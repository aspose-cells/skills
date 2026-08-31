using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Cells;

// ---------------------------------------------------------------------------
// Aspose.Cells skill verifier
//
// Scans every *.md under the skill root (SKILL.md + references/), extracts the
// C# code snippets, and reflects over the real Aspose.Cells assembly to check
// that every type/member/call shown in the docs actually exists. This is the
// anti-hallucination safety net: the docs stay truthful because the script
// re-runs against whichever Aspose.Cells version is pinned in the csproj.
//
// Exit code: 0 = all checks passed, 1 = at least one NOT FOUND.
// ---------------------------------------------------------------------------

internal static class Program
{
    private static readonly Assembly Cells = typeof(Workbook).Assembly;

    private static readonly string[] KnownSystemNamespaces =
    {
        "System", "System.IO", "System.Data", "System.Drawing", "System.Text",
        "System.Threading", "System.Collections.Generic", "System.Net",
        "System.Linq", "System.Globalization",
    };

    // Static helper types used in the docs that live outside Aspose.Cells.
    private static readonly string[] KnownStaticHelpers =
    {
        "File", "Directory", "Path", "Console", "Convert", "Math", "String",
        "Environment", "Regex", "DateTime", "TimeSpan", "Guid", "Random",
    };

    private static int _ok;
    private static int _notFound;
    private static int _skipped;
    private static int _constructors;
    private static readonly List<string> SkipSamples = new();
    private static readonly int MaxSkipSamples = 40;

    private static int Main(string[] args)
    {
        // `--compile` runs the Roslyn compile verifier (extract.py equivalent);
        // default is the reflection scan of inline references (scan-inline.py
        // equivalent). An optional trailing path overrides the skill root.
        bool compile = args.Contains("--compile");
        string? explicitRoot = args.FirstOrDefault(a => !a.StartsWith("--"));
        if (compile)
            return CompileVerifier.Run(explicitRoot);

        // Locate the skill root by walking up from the executable directory until we
        // find a folder that contains SKILL.md and references/ - works whether the
        // verifier is launched via `dotnet run` or the exe is executed directly.
        string? skillRoot = explicitRoot ?? FindSkillRoot(AppContext.BaseDirectory);
        if (skillRoot == null)
        {
            Console.Error.WriteLine("Could not locate skill root (folder with SKILL.md and references/).");
            return 1;
        }

        Console.WriteLine($"Aspose.Cells version: {Cells.GetName().Version}");
        Console.WriteLine($"Scanning: {skillRoot}");
        Console.WriteLine(new string('-', 60));

        var files = Directory.GetFiles(skillRoot, "*.md", SearchOption.AllDirectories)
            .OrderBy(f => f)
            .ToArray();

        foreach (string file in files)
        {
            CheckFile(file);
        }

        Console.WriteLine(new string('-', 60));
        Console.WriteLine($"OK: {_ok}   NOT FOUND: {_notFound}   SKIPPED: {_skipped}   CTOR-CHECKED: {_constructors}");
        if (SkipSamples.Count > 0)
        {
            Console.WriteLine(new string('-', 60));
            Console.WriteLine($"Skipped samples (first {SkipSamples.Count} of {_skipped}):");
            foreach (string s in SkipSamples) Console.WriteLine("  " + s);
        }
        return _notFound == 0 ? 0 : 1;
    }

    private static string? FindSkillRoot(string startDir)
    {
        var dir = new DirectoryInfo(startDir);
        while (dir != null)
        {
            if (File.Exists(Path.Combine(dir.FullName, "SKILL.md")) &&
                Directory.Exists(Path.Combine(dir.FullName, "references")))
                return dir.FullName;
            dir = dir.Parent;
        }
        return null;
    }

    private static void CheckFile(string file)
    {
        string rel = Path.GetRelativePath(Path.GetDirectoryName(Path.GetDirectoryName(file)) ?? ".", file);
        string[] lines = File.ReadAllLines(file);

        bool inCode = false;
        var env = new Dictionary<string, Type?>();
        for (int i = 0; i < lines.Length; i++)
        {
            string line = lines[i];
            if (line.TrimStart().StartsWith("```"))
            {
                inCode = !inCode;
                if (!inCode) env.Clear(); // fresh scope per code block
                continue;
            }
            if (!inCode) continue;

            string san = Sanitize(line);
            ScanLine(san, rel, i + 1, env, isTopLevel: true);
        }
    }

    // -----------------------------------------------------------------------
    // Sanitizer: blank out comments and string literals so the scanner only
    // sees identifiers, brackets and parens.
    // -----------------------------------------------------------------------
    private static string Sanitize(string line)
    {
        var sb = new StringBuilder(line);
        int i = 0;
        int n = line.Length;
        while (i < n)
        {
            char c = sb[i];
            if (c == '/' && i + 1 < n && sb[i + 1] == '/') { for (int k = i; k < n; k++) sb[k] = ' '; break; }
            if (c == '/' && i + 1 < n && sb[i + 1] == '*')
            {
                int end = line.IndexOf("*/", i + 2, StringComparison.Ordinal);
                int stop = end < 0 ? n : end + 2;
                for (int k = i; k < stop; k++) sb[k] = ' ';
                i = Math.Max(stop, i + 2);
                continue;
            }
            if (c == '"')
            {
                int j = i + 1;
                while (j < n)
                {
                    if (sb[j] == '"' && sb[j - 1] != '\\') break;
                    if (sb[j] == '\n') break;
                    j++;
                }
                int stop = Math.Min(j, n - 1);
                if (j < n) stop = j;
                // Mark the string literal with a control char (SOH) so the arg-type
                // matcher can tell "0" (string) from 0 (int), while keeping it an
                // opaque token that is never scanned as an identifier.
                for (int k = i; k <= stop; k++) sb[k] = '\u0001';
                i = stop + 1;
                continue;
            }
            if (c == '\'')
            {
                int j = Math.Min(i + 2, n - 1);
                if (j >= i) for (int k = i; k <= j; k++) sb[k] = '\u0002';
                i = j + 1;
                continue;
            }
            i++;
        }
        return sb.ToString();
    }

    // -----------------------------------------------------------------------
    // Line scanner: walks identifiers, builds chains, resolves them.
    // -----------------------------------------------------------------------
    private static void ScanLine(string line, string file, int lineNo, Dictionary<string, Type?> env, bool isTopLevel)
    {
        int i = 0;
        int n = line.Length;
        while (i < n)
        {
            char c = line[i];
            if (char.IsLetter(c) || c == '_')
            {
                int start = i;
                while (i < n && (char.IsLetterOrDigit(line[i]) || line[i] == '_')) i++;
                string ident = line.Substring(start, i - start);

                // Skip keywords that are not chains.
                if (ident is "return" or "foreach" or "while" or "using" or "if" or "else" or "for" or "in" or "switch" or "case" or "break" or "continue" or "var" or "this" or "params" or "public" or "private" or "class" or "override" or "new")
                {
                    if (ident == "new")
                    {
                        // Capture "new TypeName(...) { ... }" constructor usage.
                        int k = SkipWs(line, i);
                        if (k < n && char.IsLetter(line[k]))
                        {
                            int ts = k;
                            while (k < n && (char.IsLetterOrDigit(line[k]) || line[k] == '_')) k++;
                            string tname = line.Substring(ts, k - ts);
                            int p = SkipWs(line, k);
                            if (p < n && line[p] == '(')
                            {
                                (int closeP, string inner) = TakeBalanced(line, p, '(', ')');
                                if (closeP < 0) { i = p; continue; }  // multi-line: not verifiable per-line
                                int argCount = CountArgs(inner);
                                Type? t = FindTypeByName(tname);
                                if (t != null)
                                {
                                    _constructors++;
                                    if (!HasCtor(t, argCount))
                                        Report(file, lineNo, $"new {tname}({argCount} args) - no matching constructor on {t.FullName}");
                                    // recursively scan constructor args
                                    ScanLine(inner, file, lineNo, env, false);
                                }
                                else
                                {
                                    _skipped++;
                                }
                                if (closeP < n) i = closeP + 1;
                                continue;
                            }
                        }
                    }
                    i = SkipWs(line, i);
                    continue;
                }

                // It might be the start of a chain: ident (.|(|[)+
                int j = SkipWs(line, i);
                if (j < n && (line[j] == '.' || line[j] == '(' || line[j] == '['))
                {
                    ProcessChain(line, ref i, ident, file, lineNo, env);
                }
                else
                {
                    i = j;
                }
            }
            else if (c == '{' || c == '}')
            {
                i++;
            }
            else
            {
                i++;
            }
        }

        // Declaration tracking: "TypeName varName =" / "var varName = <chain>".
        TrackDeclaration(line, file, lineNo, env);
    }

    private static int SkipWs(string s, int i)
    {
        while (i < s.Length && char.IsWhiteSpace(s[i])) i++;
        return i;
    }

    private static (int close, string inner) TakeBalanced(string s, int open, char openC, char closeC)
    {
        int depth = 0;
        for (int k = open; k < s.Length; k++)
        {
            if (s[k] == openC) depth++;
            else if (s[k] == closeC)
            {
                depth--;
                if (depth == 0)
                    return (k, s.Substring(open + 1, k - open - 1));
            }
        }
        return (-1, "");
    }

    private static int CountArgs(string inner)
    {
        if (string.IsNullOrWhiteSpace(inner)) return 0;
        int depth = 0, count = 1;
        foreach (char c in inner)
        {
            if (c == '(' || c == '[' || c == '{') depth++;
            else if (c == ')' || c == ']' || c == '}') depth--;
            else if (c == ',' && depth == 0) count++;
        }
        return count;
    }

    private static void ProcessChain(string line, ref int i, string firstIdent, string file, int lineNo, Dictionary<string, Type?> env)
    {
        // Parse the chain into parts.
        var parts = new List<(string name, bool isCall, bool isIndexer, int argCount, string inner)>();
        {
            int n = line.Length;
            parts.Add((firstIdent, false, false, 0, ""));
            int k = i;
            while (true)
            {
                k = SkipWs(line, k);
                if (k >= n) break;
                char c = line[k];
                if (c == '.')
                {
                    int ks = SkipWs(line, k + 1);
                    if (ks < n && (char.IsLetter(line[ks]) || line[ks] == '_'))
                    {
                        int ke = ks;
                        while (ke < n && (char.IsLetterOrDigit(line[ke]) || line[ke] == '_')) ke++;
                        parts.Add((line.Substring(ks, ke - ks), false, false, 0, ""));
                        k = ke;
                        continue;
                    }
                    break;
                }
                if (c == '[')
                {
                    (int close, string inner) = TakeBalanced(line, k, '[', ']');
                    if (close < 0) break;   // multi-line indexer: stop this chain
                    parts.Add(("<indexer>", false, true, 0, inner));
                    ScanLine(inner, file, lineNo, env, false);
                    k = close + 1;
                    continue;
                }
                if (c == '(')
                {
                    (int close, string inner) = TakeBalanced(line, k, '(', ')');
                    if (close < 0)
                    {
                        // Multi-line call: the trailing ident is a method, arity unknown.
                        var lp = parts[^1];
                        parts[^1] = (lp.name, true, false, -1, "");
                        break;
                    }
                    var last = parts[^1];
                    parts[^1] = (last.name, true, false, CountArgs(inner), inner);
                    ScanLine(inner, file, lineNo, env, false);
                    k = close + 1;
                    continue;
                }
                break;
            }
            i = k;
        }

        // Object-initializer property check: new T { A = 1, B = 2 }
        int braceStart = SkipWs(line, i);
        if (braceStart < line.Length && line[braceStart] == '{' && parts.Count > 0)
        {
            (int close, string inner) = TakeBalanced(line, braceStart, '{', '}');
            if (close >= 0)
            {
                CheckInitializer(parts, inner, file, lineNo);
                i = close + 1;
            }
        }

        ResolveChain(parts, file, lineNo, env);
    }

    private static string ReadIdentAt(string s, int i)
    {
        int start = i;
        while (i < s.Length && (char.IsLetterOrDigit(s[i]) || s[i] == '_')) i++;
        return s.Substring(start, i - start);
    }

    // -----------------------------------------------------------------------
    // Declaration tracking: learns "TypeName var =" and "var var = <chain>".
    // -----------------------------------------------------------------------
    private static void TrackDeclaration(string line, string file, int lineNo, Dictionary<string, Type?> env)
    {
        var m = Regex.Match(line, @"^\s*(?<type>[A-Za-z_][A-Za-z0-9_<>\.]*)\s+(?<name>[A-Za-z_][A-Za-z0-9_]*)\s*=(?<rhs>.*)$");
        if (m.Success)
        {
            string typeName = m.Groups["type"].Value;
            string name = m.Groups["name"].Value;
            string rhs = m.Groups["rhs"].Value;

            Type? t = null;
            if (typeName == "var")
            {
                t = EvalExprType(rhs, env);
            }
            else
            {
                t = FindTypeByName(typeName.Trim());
            }
            if (t != null)
            {
                env[name] = t;
                // "TypeName var = new TypeName(...)" also verifies the constructor.
                var cm = Regex.Match(rhs, @"^\s*new\s+([A-Za-z_][A-Za-z0-9_\.]*)\s*\(");
                if (cm.Success)
                {
                    string tn = cm.Groups[1].Value;
                    Type? tt = FindTypeByName(tn);
                    if (tt != null)
                    {
                        int open = rhs.IndexOf('(', cm.Index);
                        (int close, string inner) = TakeBalanced(rhs, open, '(', ')');
                        if (close >= 0)
                        {
                            _constructors++;
                            if (!HasCtor(tt, CountArgs(inner)))
                                Report(file, lineNo, $"new {tn}({CountArgs(inner)} args) - no matching constructor on {tt.FullName}");
                        }
                    }
                }
            }
        }
    }

    private static Type? EvalExprType(string rhs, Dictionary<string, Type?> env)
    {
        rhs = rhs.Trim().TrimEnd(';').Trim();
        var m = Regex.Match(rhs, @"^new\s+([A-Za-z_][A-Za-z0-9_\.]*)\s*\(");
        if (m.Success) return FindTypeByName(m.Groups[1].Value);

        var mm = Regex.Match(rhs, @"^(?<chain>[A-Za-z_][A-Za-z0-9_]*(?:\.[A-Za-z_][A-Za-z0-9_]*|\[[^\]]*\]|\([^()]*\))*)\s*$");
        if (mm.Success)
        {
            var parts = new List<(string, bool, bool, int, string)>();
            foreach (Match pm in Regex.Matches(mm.Groups["chain"].Value, @"[A-Za-z_][A-Za-z0-9_]*|\[[^\]]*\]|\([^()]*\)"))
            {
                string tok = pm.Value;
                if (tok.StartsWith("[")) parts.Add(("<indexer>", false, true, 0, tok[1..^1]));
                else if (tok.StartsWith("("))
                {
                    var lp = parts[^1];
                    parts[^1] = (lp.Item1, true, false, CountArgs(tok[1..^1]), tok[1..^1]);
                }
                else parts.Add((tok, false, false, 0, ""));
            }
            return ResolveTypeOnly(parts, env);
        }
        return null;
    }

    private static Type? ResolveTypeOnly(List<(string name, bool isCall, bool isIndexer, int argCount, string inner)> parts, Dictionary<string, Type?> env)
    {
        if (parts.Count == 0) return null;
        Type? cur = null;
        int startIdx = 0;
        bool rootIsLocal = env.TryGetValue(parts[0].name, out Type? rootLocal) && rootLocal != null;
        if (!rootIsLocal)
        {
            for (int len = 1; len <= parts.Count; len++)
            {
                if (parts[len - 1].isCall || parts[len - 1].isIndexer) break;
                string joined = string.Join(".", parts.Take(len).Select(p => p.name));
                Type? candidate = FindTypeByName(joined);
                if (candidate != null) { cur = candidate; startIdx = len; }
            }
        }
        if (cur == null) cur = ResolveRoot(parts[0].name, env);
        if (startIdx == 0 && cur != null) startIdx = 1;   // root was not a fully-qualified type; skip it
        for (int i = startIdx; i < parts.Count && cur != null; i++)
        {
            var p = parts[i];
            if (p.isIndexer) cur = GetIndexerType(cur, p.inner);
            else if (p.isCall)
            {
                var (mi, _) = FindMethod(cur, p.name, p.argCount, p.argCount >= 0 ? GetArgTypes(p.inner, env) : null);
                cur = mi?.ReturnType;
            }
            else
            {
                var pi = FindPropertyOrField(cur, p.name);
                cur = pi.prop?.PropertyType ?? pi.field?.FieldType;
            }
        }
        return cur;
    }

    // -----------------------------------------------------------------------
    // Resolution + reporting for a full chain.
    // -----------------------------------------------------------------------
    private static void ResolveChain(List<(string name, bool isCall, bool isIndexer, int argCount, string inner)> parts, string file, int lineNo, Dictionary<string, Type?> env)
    {
        if (parts.Count == 0) return;

        // Try to consume fully-qualified type prefixes like "Aspose.Cells.Range"
        // or "Aspose.Cells.LowCode.PdfConverter" before the first call/indexer.
        // Skip this entirely when the root ident is a tracked local/enum value.
        int startIdx = 0;
        Type? cur = null;
        bool rootIsLocal = env.TryGetValue(parts[0].name, out Type? rootLocal) && rootLocal != null;
        if (!rootIsLocal)
        {
            for (int len = 1; len <= parts.Count; len++)
            {
                if (parts[len - 1].isCall || parts[len - 1].isIndexer) break;
                string joined = string.Join(".", parts.Take(len).Select(p => p.name));
                Type? candidate = FindTypeByName(joined);
                if (candidate != null)
                {
                    cur = candidate;
                    startIdx = len;
                }
            }
        }
        if (cur == null) cur = ResolveRoot(parts[0].name, env);
        if (cur == null)
        {
            _skipped++;
            if (SkipSamples.Count < MaxSkipSamples && parts[0].name.Length > 0)
                SkipSamples.Add($"{file}:{lineNo}  {parts[0].name}  ({string.Join(".", parts.Select(p => p.isIndexer ? "[]" : p.name))})");
            return;
        }
        if (startIdx == 0) startIdx = 1;   // the root was not a fully-qualified type; skip it

        string desc = string.Join(".", parts.Select(p => p.isIndexer ? "[]" : p.name));

        for (int i = startIdx; i < parts.Count; i++)
        {
            var p = parts[i];
            if (cur == null) break;

            if (p.isIndexer)
            {
                cur = GetIndexerType(cur, p.inner);
                continue;
            }

            if (p.isCall)
            {
                var (mi, note) = FindMethod(cur, p.name, p.argCount, p.argCount >= 0 ? GetArgTypes(p.inner, env) : null);
                if (mi == null)
                {
                    if (note == "PROPERTY")
                        Report(file, lineNo, $"{desc} - '{p.name}' is a property, not a method on {cur.Name}");
                    else if (note != null && note.StartsWith("TYPE"))
                        Report(file, lineNo, $"{desc} - argument type mismatch on {cur.Name}.{p.name}(...); expected {note.Substring(5)}");
                    else if (note != null)
                        Report(file, lineNo, $"{desc} - no '{p.name}' overload with {p.argCount} arg(s) on {cur.Name} (available arities: {note})");
                    else
                        Report(file, lineNo, $"{desc} - no method '{p.name}' on {cur.Name}");
                    return;
                }
                cur = mi.ReturnType;
                continue;
            }

            // plain property / field / enum member
            if (cur.IsEnum)
            {
                if (!Enum.IsDefined(cur, p.name))
                {
                    Report(file, lineNo, $"{desc} - enum {cur.Name} has no member '{p.name}'");
                    return;
                }
                cur = typeof(int);
                continue;
            }

            var pr = FindPropertyOrField(cur, p.name);
            if (pr.prop == null && pr.field == null)
            {
                Report(file, lineNo, $"{desc} - no property/field '{p.name}' on {cur.Name}");
                return;
            }
            cur = pr.prop?.PropertyType ?? pr.field?.FieldType;
        }

        // the final segment: report property existence when it was never a method call
        var last = parts[^1];
        if (!last.isCall && !last.isIndexer && !last.name.StartsWith("<"))
        {
            // already validated above
        }
        _ok++;
    }

    private static Type? ResolveRoot(string name, Dictionary<string, Type?> env)
    {
        if (env.TryGetValue(name, out Type? t) && t != null) return t;
        return FindTypeByName(name);
    }

    private static Type? GetIndexerType(Type? t, string inner = "")
    {
        if (t == null) return null;
        if (t.IsArray) return t.GetElementType();
        var items = t.GetProperties(BindingFlags.Public | BindingFlags.Instance)
            .Where(p => p.GetIndexParameters().Length > 0).ToArray();
        if (items.Length == 0) return null;
        // Prefer an indexer whose parameter type matches the raw argument, else the first.
        string arg = inner.Trim();
        foreach (var item in items)
        {
            var ps = item.GetIndexParameters();
            if (ps.Length == 1 && MatchesArg(ps[0].ParameterType, arg)) return item.PropertyType;
        }
        return items[0].PropertyType;
    }

    private static bool MatchesArg(Type param, string arg)
    {
        if (string.IsNullOrEmpty(arg)) return true;
        if (arg[0] == '\u0001') return param == typeof(string);   // "literal"
        if (arg[0] == '\u0002') return param == typeof(char);     // 'c'
        var n = Regex.Match(arg, @"^\d+$");
        if (n.Success) return param == typeof(int);
        return true;
    }

    private static (PropertyInfo? prop, FieldInfo? field) FindPropertyOrField(Type t, string name)
    {
        const BindingFlags flags = BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static;
        var prop = t.GetProperty(name, flags);
        if (prop != null) return (prop, null);
        var field = t.GetField(name, flags);
        return (null, field);
    }

    private static (MethodInfo? method, string? note) FindMethod(Type t, string name, int argCount = -1, List<ArgType>? argTypes = null)
    {
        const BindingFlags flags = BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static;
        var methods = t.GetMethods(flags).Where(m => m.Name == name).ToArray();
        if (methods.Length == 0)
        {
            if (t.GetProperty(name, flags) != null) return (null, "PROPERTY");
            return (null, null);
        }
        if (argCount >= 0)
        {
            var byArity = methods.Where(m => m.GetParameters().Length == argCount).ToArray();
            if (byArity.Length > 0)
            {
                // Prefer an overload whose parameter types are assignable from the
                // inferred argument types (string arg must bind a string parameter, etc.).
                var byType = byArity.Where(m => ParamsAccept(m, argTypes)).ToArray();
                if (byType.Length > 0) return (byType[0], null);
                // A params-array last parameter matches any argCount >= its position.
                foreach (var m in byArity)
                {
                    var ps = m.GetParameters();
                    if (ps.Length > 0 && ps[ps.Length - 1].GetCustomAttributes(typeof(ParamArrayAttribute), false).Length > 0)
                        return (m, null);
                }
                if (argTypes != null && argTypes.Any(a => a.Kind != ArgKind.Unknown))
                {
                    // Every arity match has a type clash: report which param expects what.
                    var sig = string.Join("  ", byArity.Select(m => $"({string.Join(", ", m.GetParameters().Select(p => p.ParameterType.Name))})"));
                    return (null, "TYPE " + sig);
                }
                return (byArity[0], null);
            }
            // A params-array last parameter matches any argCount >= its position.
            foreach (var m in methods)
            {
                var ps = m.GetParameters();
                if (ps.Length > 0 && ps.Length <= argCount &&
                    ps[ps.Length - 1].GetCustomAttributes(typeof(ParamArrayAttribute), false).Length > 0)
                    return (m, null);
            }
            // No overload with that arity: report the actual arities so the doc can be fixed.
            return (null, string.Join(",", methods.Select(m => m.GetParameters().Length).Distinct().OrderBy(x => x)));
        }
        return (methods[0], null);
    }

    private static bool ParamsAccept(MethodInfo m, List<ArgType>? argTypes)
    {
        if (argTypes == null) return true;
        var ps = m.GetParameters();
        for (int i = 0; i < argTypes.Count; i++)
        {
            var a = argTypes[i];
            if (a.Kind == ArgKind.Unknown) continue;      // unknown type - skip
            var p = ps[i];
            if (p.GetCustomAttributes(typeof(ParamArrayAttribute), false).Length > 0)
                return true;                              // params - accept
            if (!ArgAssignable(a, p.ParameterType)) return false;
        }
        return true;
    }

    private static bool ArgAssignable(ArgType arg, Type param)
    {
        // resolved real type: use the framework's assignability rules
        if (arg.Kind == ArgKind.Resolved && arg.Resolved != null)
        {
            if (param.IsAssignableFrom(arg.Resolved)) return true;
            if (param == typeof(object)) return true;
            // numeric widening for integer literals handled via Kind.Int32 below
            return false;
        }
        if (param == typeof(object)) return true;
        switch (arg.Kind)
        {
            case ArgKind.String: return param == typeof(string);
            case ArgKind.Char: return param == typeof(char);
            case ArgKind.Bool: return param == typeof(bool);
            case ArgKind.Int32:
                return param == typeof(int) || param == typeof(long) || param == typeof(double)
                    || param == typeof(decimal) || param == typeof(float) || param == typeof(short)
                    || param == typeof(byte) || param == typeof(object);
            case ArgKind.Int64:
                return param == typeof(long) || param == typeof(double) || param == typeof(decimal)
                    || param == typeof(object);
            case ArgKind.Float:
                return param == typeof(float) || param == typeof(double) || param == typeof(object);
            case ArgKind.Double:
                return param == typeof(double) || param == typeof(decimal) || param == typeof(object);
            case ArgKind.Decimal:
                return param == typeof(decimal) || param == typeof(double) || param == typeof(object);
        }
        return false;
    }

    /// <summary>Split a call's argument list on top-level commas and classify each arg.</summary>
    private static List<ArgType> GetArgTypes(string inner, Dictionary<string, Type?> env)
    {
        var parts = new List<string>();
        int depth = 0, start = 0;
        for (int i = 0; i < inner.Length; i++)
        {
            char c = inner[i];
            if (c == '(' || c == '[' || c == '{') depth++;
            else if (c == ')' || c == ']' || c == '}') depth--;
            else if (c == ',' && depth == 0) { parts.Add(inner.Substring(start, i - start)); start = i + 1; }
        }
        parts.Add(inner.Substring(start));
        return parts.Select(p => ClassifyArg(p.Trim(), env)).ToList();
    }

    private static ArgType ClassifyArg(string a, Dictionary<string, Type?> env)
    {
        if (a.Length == 0) return ArgType.UnknownType;
        if (a[0] == '\u0001') return new ArgType(ArgKind.String, null);      // "literal"
        if (a[0] == '\u0002') return new ArgType(ArgKind.Char, null);        // 'c'
        if (Regex.IsMatch(a, @"^\d+$")) return new ArgType(ArgKind.Int32, null);
        if (Regex.IsMatch(a, @"^-?\d+\.\d+([eE][+-]?\d+)?[dD]?$")) return new ArgType(ArgKind.Double, null);
        if (Regex.IsMatch(a, @"^-?\d+[fF]$")) return new ArgType(ArgKind.Float, null);
        if (Regex.IsMatch(a, @"^-?\d+[mM]$")) return new ArgType(ArgKind.Decimal, null);
        if (Regex.IsMatch(a, @"^-?\d+[lL]$")) return new ArgType(ArgKind.Int64, null);
        if (a == "true" || a == "false") return new ArgType(ArgKind.Bool, null);
        if (a == "null") return ArgType.UnknownType;
        if (a.StartsWith("new "))
        {
            // new TypeName(...) or new TypeName { ... }
            var m = Regex.Match(a, @"new\s+([A-Za-z_][A-Za-z0-9_\.<>]*)");
            if (m.Success)
            {
                var t = FindTypeByName(m.Groups[1].Value);
                if (t != null) return new ArgType(ArgKind.Resolved, t);
            }
            return ArgType.UnknownType;
        }
        // enum member like SaveFormat.Xlsx, or a static member like Color.SteelBlue
        var em = Regex.Match(a, @"^([A-Za-z_][A-Za-z0-9_]*)\.([A-Za-z_][A-Za-z0-9_]*)$");
        if (em.Success)
        {
            var et = FindTypeByName(em.Groups[1].Value);
            if (et != null) return new ArgType(ArgKind.Resolved, et);
        }
        // a tracked local's declared type (e.g. `pt` is a PivotTable)
        if (env.TryGetValue(a, out Type? loc) && loc != null) return new ArgType(ArgKind.Resolved, loc);
        // any other identifier: bind by its declared type if we can (checked elsewhere), else unknown
        return ArgType.UnknownType;
    }

    private enum ArgKind { Unknown, String, Char, Bool, Int32, Int64, Float, Double, Decimal, Resolved }

    private readonly record struct ArgType(ArgKind Kind, Type? Resolved)
    {
        public static readonly ArgType UnknownType = new(ArgKind.Unknown, null);
    }

    private static bool HasCtor(Type t, int argCount)
    {
        return t.GetConstructors(BindingFlags.Public | BindingFlags.Instance)
            .Any(c => c.GetParameters().Length == argCount);
    }

    private static void CheckInitializer(List<(string name, bool isCall, bool isIndexer, int argCount, string inner)> parts, string inner, string file, int lineNo)
    {
        // The type name is the FIRST identifier of the chain, e.g. "StyleFlag" or "opts".
        string typeName = parts[0].name;
        Type? t = FindTypeByName(typeName);
        if (t == null) { _skipped++; return; }

        foreach (Match pm in Regex.Matches(inner, @"(?<prop>[A-Za-z_][A-Za-z0-9_]*)\s*="))
        {
            string prop = pm.Groups["prop"].Value;
            const BindingFlags flags = BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static;
            if (t.GetProperty(prop, flags) == null && t.GetField(prop, flags) == null)
                Report(file, lineNo, $"{typeName} initializer - no property '{prop}' on {t.Name}");
        }
    }

    private static void Report(string file, int line, string message)
    {
        _notFound++;
        Console.WriteLine($"[NOT FOUND] {file}:{line}  {message}");
    }

    // -----------------------------------------------------------------------
    // Type lookup across Aspose.Cells + common System namespaces.
    // -----------------------------------------------------------------------
    private static Type? FindTypeByName(string name)
    {
        name = name.Trim();
        if (name.EndsWith("]")) return null;

        // Fully-qualified names like "Aspose.Cells.Range" or "Aspose.Cells.LowCode.PdfConverter"
        Type? fq = Cells.GetType(name);
        if (fq != null) return fq;

        Type? t = Cells.GetType("Aspose.Cells." + name);
        if (t != null) return t;

        // known static helper types
        if (KnownStaticHelpers.Contains(name))
        {
            Type? h = KnownSystemNamespaces
                .Select(ns => Type.GetType($"{ns}.{name}, System.Private.CoreLib"))
                .FirstOrDefault(x => x != null)
                ?? KnownSystemNamespaces.Select(ns => Type.GetType($"{ns}.{name}, System.Runtime"))
                .FirstOrDefault(x => x != null);
            if (h != null) return h;
        }

        // generic like List<T>
        var gm = Regex.Match(name, @"^(?<g>[A-Za-z_][A-Za-z0-9_]*)<");
        if (gm.Success)
        {
            string g = gm.Groups["g"].Value;
            Type? def = FindTypeByName(g);
            if (def != null && def.IsGenericTypeDefinition)
                return def;
        }

        // fallback: search the whole Aspose.Cells assembly by simple name
        try
        {
            Type? bySimple = Cells.GetTypes()
                .FirstOrDefault(x => string.Equals(x.Name, name, StringComparison.Ordinal));
            if (bySimple != null) return bySimple;
        }
        catch (ReflectionTypeLoadException ex)
        {
            Type? bySimple = ex.Types
                .Where(x => x != null)
                .FirstOrDefault(x => string.Equals(x!.Name, name, StringComparison.Ordinal));
            if (bySimple != null) return bySimple;
        }

        // plain System types
        foreach (string ns in KnownSystemNamespaces)
        {
            foreach (string asm in new[] { "System.Private.CoreLib", "System.Runtime", "System.Data.Common" })
            {
                Type? x = Type.GetType($"{ns}.{name}, {asm}");
                if (x != null) return x;
            }
        }
        return null;
    }
}
