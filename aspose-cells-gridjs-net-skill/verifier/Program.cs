using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;

// ---------------------------------------------------------------------------
// GridJs skill verifier
//
// Four modes:
//   (default)  Reflection scan of server-side inline API references
//   --compile  Roslyn-compile every ```csharp block
//   --client   Validate client-side API references against index.d.ts
//   --names    Name-level cross-check: every API name in skill must exist in assembly
//
// Exit code: 0 = all checks passed, 1 = at least one issue found.
// ---------------------------------------------------------------------------

internal static class Program
{
    private static int Main(string[] args)
    {
        bool compile = args.Contains("--compile");
        bool client = args.Contains("--client");
        bool names = args.Contains("--names");
        string? explicitRoot = args.FirstOrDefault(a => !a.StartsWith("--"));

        if (compile)
            return CompileVerifier.Run(explicitRoot);

        if (client)
            return ClientApiVerifier.Run(explicitRoot);

        if (names)
            return NameVerifier.Run(explicitRoot);

        // Default: reflection scan of server-side API
        return RunReflectionScan(explicitRoot);
    }

    // -----------------------------------------------------------------------
    // Reflection scan (server-side)
    // -----------------------------------------------------------------------
    private static readonly Assembly GridJsAssembly = typeof(Aspose.Cells.GridJs.GridJsWorkbook).Assembly;

    private static readonly string[] KnownSystemNamespaces =
    {
        "System", "System.IO", "System.Data", "System.Drawing", "System.Text",
        "System.Threading", "System.Collections.Generic", "System.Net",
        "System.Linq", "System.Globalization", "System.Threading.Tasks",
    };

    private static readonly string[] KnownStaticHelpers =
    {
        "File", "Directory", "Path", "Console", "Convert", "Math", "String",
        "Environment", "Regex", "DateTime", "TimeSpan", "Guid", "Random",
        "Task", "Encoding", "StringBuilder", "MemoryStream", "FileStream",
    };

    private static int _ok;
    private static int _notFound;
    private static int _skipped;
    private static int _constructors;
    private static readonly List<string> SkipSamples = new();
    private static readonly int MaxSkipSamples = 40;

    private static int RunReflectionScan(string? explicitRoot)
    {
        string? skillRoot = explicitRoot ?? FindSkillRoot(AppContext.BaseDirectory);
        if (skillRoot == null)
        {
            Console.Error.WriteLine("Could not locate skill root (folder with SKILL.md and references/).");
            return 1;
        }

        Console.WriteLine($"GridJs assembly version: {GridJsAssembly.GetName().Version}");
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
            // Pattern 1: SKILL.md + references/ (reference skill layout)
            if (File.Exists(Path.Combine(dir.FullName, "SKILL.md")) &&
                Directory.Exists(Path.Combine(dir.FullName, "references")))
                return dir.FullName;
            // Pattern 2: skill-index.md with .md files (template-generated layout)
            if (File.Exists(Path.Combine(dir.FullName, "skill-index.md")))
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
                if (!inCode) env.Clear();
                continue;
            }
            if (!inCode) continue;

            string san = Sanitize(line);
            ScanLine(san, rel, i + 1, env, isTopLevel: true);
        }
    }

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
                while (j < n) { if (sb[j] == '"' && sb[j - 1] != '\\') break; if (sb[j] == '\n') break; j++; }
                int stop = Math.Min(j, n - 1);
                if (j < n) stop = j;
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

                if (ident is "return" or "foreach" or "while" or "using" or "if" or "else" or "for" or "in" or "switch" or "case" or "break" or "continue" or "var" or "this" or "params" or "public" or "private" or "class" or "override" or "new")
                {
                    if (ident == "new")
                    {
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
                                if (closeP < 0) { i = p; continue; }
                                int argCount = CountArgs(inner);
                                Type? t = FindTypeByName(tname);
                                if (t != null)
                                {
                                    _constructors++;
                                    if (!HasCtor(t, argCount))
                                        Report(file, lineNo, $"new {tname}({argCount} args) - no matching constructor on {t.FullName}");
                                    ScanLine(inner, file, lineNo, env, false);
                                }
                                else { _skipped++; }
                                if (closeP < n) i = closeP + 1;
                                continue;
                            }
                        }
                    }
                    i = SkipWs(line, i);
                    continue;
                }

                int j = SkipWs(line, i);
                if (j < n && (line[j] == '.' || line[j] == '(' || line[j] == '['))
                {
                    ProcessChain(line, ref i, ident, file, lineNo, env);
                }
                else { i = j; }
            }
            else if (c == '{' || c == '}') { i++; }
            else { i++; }
        }
        TrackDeclaration(line, file, lineNo, env);
    }

    private static int SkipWs(string s, int i) { while (i < s.Length && char.IsWhiteSpace(s[i])) i++; return i; }

    private static (int close, string inner) TakeBalanced(string s, int open, char openC, char closeC)
    {
        int depth = 0;
        for (int k = open; k < s.Length; k++)
        {
            if (s[k] == openC) depth++;
            else if (s[k] == closeC) { depth--; if (depth == 0) return (k, s.Substring(open + 1, k - open - 1)); }
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
                        k = ke; continue;
                    }
                    break;
                }
                if (c == '[')
                {
                    (int close, string inner) = TakeBalanced(line, k, '[', ']');
                    if (close < 0) break;
                    parts.Add(("<indexer>", false, true, 0, inner));
                    ScanLine(inner, file, lineNo, env, false);
                    k = close + 1; continue;
                }
                if (c == '(')
                {
                    (int close, string inner) = TakeBalanced(line, k, '(', ')');
                    if (close < 0) { var lp = parts[^1]; parts[^1] = (lp.name, true, false, -1, ""); break; }
                    var last = parts[^1];
                    parts[^1] = (last.name, true, false, CountArgs(inner), inner);
                    ScanLine(inner, file, lineNo, env, false);
                    k = close + 1; continue;
                }
                break;
            }
            i = k;
        }

        int braceStart = SkipWs(line, i);
        if (braceStart < line.Length && line[braceStart] == '{' && parts.Count > 0)
        {
            (int close, string inner) = TakeBalanced(line, braceStart, '{', '}');
            if (close >= 0) { CheckInitializer(parts, inner, file, lineNo); i = close + 1; }
        }

        ResolveChain(parts, file, lineNo, env);
    }

    private static void TrackDeclaration(string line, string file, int lineNo, Dictionary<string, Type?> env)
    {
        var m = Regex.Match(line, @"^\s*(?<type>[A-Za-z_][A-Za-z0-9_<>\.]*)\s+(?<name>[A-Za-z_][A-Za-z0-9_]*)\s*=(?<rhs>.*)$");
        if (!m.Success) return;
        string typeName = m.Groups["type"].Value;
        string name = m.Groups["name"].Value;
        string rhs = m.Groups["rhs"].Value;
        Type? t = typeName == "var" ? EvalExprType(rhs, env) : FindTypeByName(typeName.Trim());
        if (t != null) env[name] = t;
    }

    private static Type? EvalExprType(string rhs, Dictionary<string, Type?> env)
    {
        rhs = rhs.Trim().TrimEnd(';').Trim();
        var m = Regex.Match(rhs, @"^new\s+([A-Za-z_][A-Za-z0-9_\.]*)\s*\(");
        if (m.Success) return FindTypeByName(m.Groups[1].Value);
        return null;
    }

    private static void ResolveChain(List<(string name, bool isCall, bool isIndexer, int argCount, string inner)> parts, string file, int lineNo, Dictionary<string, Type?> env)
    {
        if (parts.Count == 0) return;
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
                if (candidate != null) { cur = candidate; startIdx = len; }
            }
        }
        if (cur == null) cur = ResolveRoot(parts[0].name, env);
        if (cur == null) { _skipped++; return; }
        if (startIdx == 0) startIdx = 1;

        for (int i = startIdx; i < parts.Count && cur != null; i++)
        {
            var p = parts[i];
            if (p.isIndexer) { cur = GetIndexerType(cur, p.inner); continue; }
            if (p.isCall)
            {
                var (mi, note) = FindMethod(cur, p.name, p.argCount);
                if (mi == null)
                {
                    if (note == "PROPERTY") Report(file, lineNo, $"'{p.name}' is a property, not a method on {cur.Name}");
                    else if (note != null) Report(file, lineNo, $"no '{p.name}' overload with {p.argCount} arg(s) on {cur.Name} (available: {note})");
                    else Report(file, lineNo, $"no method '{p.name}' on {cur.Name}");
                    return;
                }
                cur = mi.ReturnType; continue;
            }
            if (cur.IsEnum)
            {
                if (!Enum.IsDefined(cur, p.name)) { Report(file, lineNo, $"enum {cur.Name} has no member '{p.name}'"); return; }
                cur = typeof(int); continue;
            }
            var pr = FindPropertyOrField(cur, p.name);
            if (pr.prop == null && pr.field == null) { Report(file, lineNo, $"no property/field '{p.name}' on {cur.Name}"); return; }
            cur = pr.prop?.PropertyType ?? pr.field?.FieldType;
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
        var items = t.GetProperties(BindingFlags.Public | BindingFlags.Instance).Where(p => p.GetIndexParameters().Length > 0).ToArray();
        if (items.Length == 0) return null;
        return items[0].PropertyType;
    }

    private static (PropertyInfo? prop, FieldInfo? field) FindPropertyOrField(Type t, string name)
    {
        const BindingFlags flags = BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static;
        var prop = t.GetProperty(name, flags);
        if (prop != null) return (prop, null);
        return (null, t.GetField(name, flags));
    }

    private static (MethodInfo? method, string? note) FindMethod(Type t, string name, int argCount = -1)
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
            if (byArity.Length > 0) return (byArity[0], null);
            return (null, string.Join(",", methods.Select(m => m.GetParameters().Length).Distinct().OrderBy(x => x)));
        }
        return (methods[0], null);
    }

    private static bool HasCtor(Type t, int argCount)
    {
        return t.GetConstructors(BindingFlags.Public | BindingFlags.Instance).Any(c => c.GetParameters().Length == argCount);
    }

    private static void CheckInitializer(List<(string name, bool isCall, bool isIndexer, int argCount, string inner)> parts, string inner, string file, int lineNo)
    {
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

    private static Type? FindTypeByName(string name)
    {
        name = name.Trim();
        if (name.EndsWith("]")) return null;

        // GridJs types
        Type? fq = GridJsAssembly.GetType(name);
        if (fq != null) return fq;
        Type? t = GridJsAssembly.GetType("Aspose.Cells.GridJs." + name);
        if (t != null) return t;

        // Also check Aspose.Cells namespace (GridJs re-exports some types)
        try
        {
            var cellsAsm = AppDomain.CurrentDomain.GetAssemblies().FirstOrDefault(a => a.GetName().Name == "Aspose.Cells");
            if (cellsAsm != null)
            {
                Type? ct = cellsAsm.GetType("Aspose.Cells." + name);
                if (ct != null) return ct;
            }
        }
        catch { }

        // Known static helpers
        if (KnownStaticHelpers.Contains(name))
        {
            foreach (string ns in KnownSystemNamespaces)
            {
                foreach (string asm in new[] { "System.Private.CoreLib", "System.Runtime" })
                {
                    Type? h = Type.GetType($"{ns}.{name}, {asm}");
                    if (h != null) return h;
                }
            }
        }

        // Generic
        var gm = Regex.Match(name, @"^(?<g>[A-Za-z_][A-Za-z0-9_]*)<");
        if (gm.Success) { Type? def = FindTypeByName(gm.Groups["g"].Value); if (def != null && def.IsGenericTypeDefinition) return def; }

        // Fallback: search by simple name
        try
        {
            Type? bySimple = GridJsAssembly.GetTypes().FirstOrDefault(x => string.Equals(x.Name, name, StringComparison.Ordinal));
            if (bySimple != null) return bySimple;
        }
        catch (ReflectionTypeLoadException ex)
        {
            Type? bySimple = ex.Types.Where(x => x != null).FirstOrDefault(x => string.Equals(x!.Name, name, StringComparison.Ordinal));
            if (bySimple != null) return bySimple;
        }

        // System types
        foreach (string ns in KnownSystemNamespaces)
        {
            foreach (string asm in new[] { "System.Private.CoreLib", "System.Runtime" })
            {
                Type? x = Type.GetType($"{ns}.{name}, {asm}");
                if (x != null) return x;
            }
        }
        return null;
    }
}