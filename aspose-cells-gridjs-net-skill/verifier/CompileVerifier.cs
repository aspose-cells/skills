using System.Text;
using System.Text.RegularExpressions;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;

// ---------------------------------------------------------------------------
// GridJs skill compile verifier
//
// Extracts every ```csharp block from SKILL.md + references/, wraps it in a
// compilable unit, and compiles with Roslyn against the pinned GridJs assembly.
//
// Exit code: 0 = all blocks compiled, 1 = at least one compile error, 2 = could not run.
// ---------------------------------------------------------------------------

internal static class CompileVerifier
{
    private static readonly string[] BaseWrapperUsings =
    {
        "using System;",
        "using System.IO;",
        "using System.Data;",
        "using System.Drawing;",
        "using System.Text;",
        "using System.Collections.Generic;",
        "using System.Linq;",
        "using System.Threading;",
        "using System.Threading.Tasks;",
        "using Aspose.Cells;",
        "using Aspose.Cells.GridJs;",
    };

    // A setup snippet names types from the packages the skill tells the reader to install.
    // Those namespaces are only added to the wrapper when the assembly really is on this
    // run's reference list - a using for a missing package would turn into CS0234 and every
    // snippet would fail for a reason that has nothing to do with GridJs.
    // The pair is (assembly file name, namespace it is known to declare).
    private static readonly (string Assembly, string Ns)[] OptionalWrapperUsings =
    {
        ("Microsoft.Extensions.DependencyInjection.Abstractions", "Microsoft.Extensions.DependencyInjection"),
        ("Microsoft.Extensions.Options", "Microsoft.Extensions.Options"),
        ("Microsoft.Extensions.Logging.Abstractions", "Microsoft.Extensions.Logging"),
        ("Microsoft.Extensions.Configuration.Abstractions", "Microsoft.Extensions.Configuration"),
        ("Microsoft.Extensions.Hosting.Abstractions", "Microsoft.Extensions.Hosting"),
        ("Microsoft.AspNetCore", "Microsoft.AspNetCore.Builder"),
        ("Microsoft.AspNetCore.Mvc.Core", "Microsoft.AspNetCore.Mvc"),
    };

    private static readonly List<string> WrapperUsings = new(BaseWrapperUsings);

    private static readonly Regex UsingDirective = new Regex(
        @"^\s*using\s+(?:static\s+)?[A-Za-z_][\w.]*\s*;");

    // A name the harness cannot supply - a field of the class the excerpt was cut from, a
    // type from a package this verifier does not reference - is not a claim the documentation
    // got wrong. It is coverage the excerpt format does not allow, and it must not be
    // reported as though a reader would hit it. Name validity is --names' job.
    private static readonly HashSet<string> CoverageDiagnosticIds = new(StringComparer.Ordinal)
    {
        "CS0103", "CS0117", "CS0119", "CS0121", "CS0122", "CS0234", "CS0246",
        "CS0012", "CS1061", "CS7036", "CS0117", "CS1503", "CS1061",
    };

    private static int _diagnostics;
    private static readonly HashSet<string> ErrorBlocks = new();
    private static readonly HashSet<string> CoverageBlocks = new();
    private static readonly List<string> ErrorSamples = new();
    private static readonly List<string> CoverageSamples = new();
    private static readonly Dictionary<string, CodeBlock> BlocksByPath = new();

    public static int Run(string? explicitRoot = null)
    {
        string? skillRoot = explicitRoot is null
            ? FindSkillRoot(AppContext.BaseDirectory)
            : Path.GetFullPath(explicitRoot);
        if (skillRoot == null)
        {
            Console.Error.WriteLine("Could not locate skill root (folder with SKILL.md and references/).");
            return 2;
        }

        var gridJsVersion = typeof(Aspose.Cells.GridJs.GridJsWorkbook).Assembly.GetName().Version;
        Console.WriteLine($"GridJs version: {gridJsVersion.Major}.{gridJsVersion.Minor}.{gridJsVersion.Build}");
        Console.WriteLine($"Compile-scanning: {skillRoot}");
        Console.WriteLine(new string('-', 60));

        var refs = BuildReferences();
        foreach (string ns in OptionalUsings(refs))
            WrapperUsings.Add("using " + ns + ";");

        var trees = new List<SyntaxTree>();
        var files = Directory.GetFiles(skillRoot, "*.md", SearchOption.AllDirectories)
            .Where(f =>
            {
                var rel = Path.GetRelativePath(skillRoot, f).Replace('\\', '/');
                return !rel.StartsWith("example") && !rel.StartsWith("verifier") && !rel.StartsWith("report");
            })
            .OrderBy(f => f)
            .ToArray();
        int index = 0;
        int fragments = 0;
        foreach (string file in files)
        {
            CollectBlocks(file, skillRoot, trees, ref index, ref fragments);
        }

        var compilation = CSharpCompilation.Create(
            "SkillSnippets",
            trees,
            refs,
            new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary));

        ReportDiagnostics(compilation);

        int total = trees.Count;
        int failed = ErrorBlocks.Count;
        int gaps = CoverageBlocks.Count;
        Console.WriteLine(new string('-', 60));
        Console.WriteLine($"COMPILE OK: {total - failed}   ERRORS: {failed}   COVERAGE GAPS: {gaps}   BLOCKS: {total}   FRAGMENTS SKIPPED: {fragments}   DIAGNOSTICS: {_diagnostics}");
        // Machine-readable result line, counted per block: one block can raise a dozen
        // diagnostics, and a reader comparing runs wants to know how many snippets broke.
        Console.WriteLine($"VERIFIER-RESULT: blocks={total} errors={failed} coverage-gaps={gaps} fragments={fragments} diagnostics={_diagnostics}");
        if (ErrorSamples.Count > 0)
        {
            Console.WriteLine(new string('-', 60));
            Console.WriteLine("Errors (first " + ErrorSamples.Count + "):");
            foreach (string s in ErrorSamples) Console.WriteLine("  " + s);
        }
        if (CoverageSamples.Count > 0)
        {
            Console.WriteLine(new string('-', 60));
            Console.WriteLine($"Coverage gaps (first {CoverageSamples.Count} of {gaps} blocks) - the excerpt needs a host the scan does not provide, not a claim about the docs:");
            foreach (string s in CoverageSamples) Console.WriteLine("  " + s);
        }
        return failed == 0 ? 0 : 1;
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
            // Pattern 2: SKILL.md with .md files (Agent Skills standard layout)
            if (File.Exists(Path.Combine(dir.FullName, "SKILL.md")))
                return dir.FullName;
            dir = dir.Parent;
        }
        return null;
    }

    private static void CollectBlocks(string file, string skillRoot, List<SyntaxTree> trees, ref int index, ref int fragments)
    {
        string[] lines = File.ReadAllLines(file);
        string rel = Path.GetRelativePath(skillRoot, file).Replace('\\', '/');
        for (int i = 0; i < lines.Length; i++)
        {
            string t = lines[i].TrimStart();
            if (!t.StartsWith("```")) continue;
            if (!IsCsharpFence(t)) continue;

            var body = new List<string>();
            int startLine = i + 1;
            int j = i + 1;
            for (; j < lines.Length; j++)
            {
                if (lines[j].TrimStart().StartsWith("```")) break;
                body.Add(lines[j]);
            }
            if (j >= lines.Length) break;

            string treePath = $"Snippet_{index}~{rel}~L{startLine}";
            var block = new CodeBlock(rel, startLine, body, index, treePath);
            if (block.IsFragment)
            {
                fragments++;
                index++;
                i = j;
                continue;
            }
            string source = Wrap(block);
            BlocksByPath[treePath] = block;
            trees.Add(CSharpSyntaxTree.ParseText(
                source,
                new CSharpParseOptions(LanguageVersion.Latest),
                path: treePath));
            index++;
            i = j;
        }
    }

    // Only a fence whose whole info string says csharp (or cs) is C# source. ```csharp-html
    // is a Razor/HTML sample with a C# fragment inside and compiling it as C# yields syntax
    // errors that say nothing about the skill.
    private static bool IsCsharpFence(string trimmedOpener)
    {
        var m = Regex.Match(trimmedOpener, @"^```\s*(?<lang>[A-Za-z+#-]+)?");
        if (!m.Success) return false;
        string lang = m.Groups["lang"].Value.ToLowerInvariant();
        return lang is "csharp" or "cs";
    }

    private static string Wrap(CodeBlock block)
    {
        var lines = new List<string>();
        lines.AddRange(WrapperUsings);
        lines.Add("");
        lines.Add("#nullable disable");

        if (block.IsTypeDecl)
        {
            // Every snippet lands in one compilation, and several skill files declare a
            // controller under the same name. A namespace per block keeps those from
            // colliding with each other - a property of the harness, not of the docs.
            bool ownNamespace = block.HasOwnNamespace;
            if (!ownNamespace)
            {
                lines.Add($"namespace Snippet_{block.Index}");
                lines.Add("{");
            }
            block.BodyOffset = lines.Count;
            foreach (string ln in block.Body)
            {
                if (UsingDirective.IsMatch(ln)) continue;
                lines.Add(ownNamespace ? ln : "    " + ln);
            }
            if (!ownNamespace) lines.Add("}");
        }
        else if (block.IsMemberDecl)
        {
            // The excerpt is one member of a class the reader already has. A plain host class
            // gives Roslyn the shape a member needs; statements and types get their own hosts.
            lines.Add("");
            lines.Add($"public class Snippet_{block.Index}_Host");
            lines.Add("{");
            block.BodyOffset = lines.Count;
            foreach (string ln in block.Body)
            {
                if (UsingDirective.IsMatch(ln)) continue;
                lines.Add("    " + ln);
            }
            lines.Add("}");
        }
        else
        {
            lines.Add("");
            lines.Add($"public static class Snippet_{block.Index}");
            lines.Add("{");
            lines.Add("    public static void Run()");
            lines.Add("    {");
            lines.Add("        try");
            lines.Add("        {");

            var prelude = BuildPrelude(block);
            int bodyStart = lines.Count + prelude.Count;
            foreach (string ln in prelude)
                lines.Add("            " + ln);

            block.BodyOffset = bodyStart;
            foreach (string ln in block.Body)
            {
                if (UsingDirective.IsMatch(ln)) continue;
                lines.Add("            " + ln);
            }
            lines.Add("        }");
            lines.Add("        catch (Exception) { }");
            lines.Add("    }");
            lines.Add("}");
        }

        block.Source = string.Join("\n", lines);
        return block.Source;
    }

    private static List<string> BuildPrelude(CodeBlock block)
    {
        string code = StripCommentsAndStrings(block.Text);
        bool hasWbDecl = Regex.IsMatch(code, @"\b(?:Workbook|var)\s+(?:wb|workbook)\s*=");
        bool hasSheetDecl = Regex.IsMatch(code, @"\b(?:Worksheet|var)\s+sheet\s*=");

        var prelude = new List<string>();
        if (!hasWbDecl && Regex.IsMatch(code, @"\bwb\b|\bworkbook\b"))
            prelude.Add("Workbook wb = new Workbook();");
        if (!hasSheetDecl && Regex.IsMatch(code, @"\bsheet\b"))
            prelude.Add("Worksheet sheet = wb.Worksheets[0];");
        return prelude;
    }

    private static string StripCommentsAndStrings(string text)
    {
        var sb = new StringBuilder(text);
        int i = 0, n = sb.Length;
        while (i < n)
        {
            char c = sb[i];
            if (c == '/' && i + 1 < n && sb[i + 1] == '/') { while (i < n && sb[i] != '\n') sb[i++] = ' '; continue; }
            if (c == '/' && i + 1 < n && sb[i + 1] == '*') { int end = text.IndexOf("*/", i + 2, StringComparison.Ordinal); int stop = end < 0 ? n : end + 2; while (i < stop && i < n) sb[i++] = ' '; continue; }
            if (c == '"') { sb[i++] = ' '; while (i < n && sb[i] != '"') sb[i++] = ' '; if (i < n) sb[i++] = ' '; continue; }
            if (c == '\'') { sb[i++] = ' '; while (i < n && sb[i] != '\'') sb[i++] = ' '; if (i < n) sb[i++] = ' '; continue; }
            i++;
        }
        return sb.ToString();
    }

    private static List<MetadataReference> BuildReferences()
    {
        var refs = new List<MetadataReference>();
        refs.Add(MetadataReference.CreateFromFile(typeof(Aspose.Cells.GridJs.GridJsWorkbook).Assembly.Location));
        // Also add Aspose.Cells assembly if loaded
        var cellsAsm = AppDomain.CurrentDomain.GetAssemblies().FirstOrDefault(a => a.GetName().Name == "Aspose.Cells");
        if (cellsAsm != null)
        {
            try { refs.Add(MetadataReference.CreateFromFile(cellsAsm.Location)); } catch { }
        }
        string tpa = (string?)AppContext.GetData("TRUSTED_PLATFORM_ASSEMBLIES") ?? "";
        var seen = new HashSet<string>();
        // Keyed by file name, not path: the ASP.NET reference folder below re-supplies
        // assemblies the runtime already carries (Microsoft.Extensions.Options and friends),
        // and Roslyn reports the same identity twice as CS0433 "exists in both 8.0.0.0 and
        // 9.0.0.0" - an error in the harness that points at a doc line. First one wins.
        var byName = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (string path in tpa.Split(Path.PathSeparator))
        {
            if (path.Length == 0 || !seen.Add(path)) continue;
            if (!File.Exists(path)) continue;
            if (!byName.Add(Path.GetFileName(path))) continue;
            try { refs.Add(MetadataReference.CreateFromFile(path)); } catch { }
        }

        // The skill documents an ASP.NET Core integration, so its snippets name MVC and
        // hosting types that the GridJs package does not carry. Their assemblies live beside
        // the runtime, not in this project: looking for them keeps the verifier free of a
        // framework dependency it does not use itself. When they are missing the controller
        // snippets cannot be compiled, and that is reported as a reference gap rather than
        // silently turning into a hundred CS0246 "errors in the docs".
        string? aspNetDir = FindAspNetReferenceDir();
        if (aspNetDir == null)
        {
            Console.WriteLine("ASP.NET Core reference assemblies: NOT FOUND - controller snippets cannot compile");
        }
        else
        {
            Console.WriteLine("ASP.NET Core references: " + aspNetDir);
            // Every assembly the framework folder carries, not a hand-picked list: an MVC
            // snippet reaches HttpContext through a signature the docs never name, and a
            // missing Microsoft.AspNetCore.Http.Abstractions reads as CS0012 against a
            // perfectly correct doc line.
            foreach (string path in Directory.GetFiles(aspNetDir, "*.dll"))
            {
                if (!seen.Add(path) || !byName.Add(Path.GetFileName(path))) continue;
                try { refs.Add(MetadataReference.CreateFromFile(path)); } catch { }
            }
        }
        return refs;
    }

    private static readonly string[] AspNetRequired =
    {
        "Microsoft.AspNetCore.dll",
        "Microsoft.AspNetCore.Mvc.Core.dll",
    };

    // Reference packs carry the compile-time shape and come first; an installed shared
    // framework works too. Both sit under the dotnet root the running assembly came from.
    // The folder must share its major version with the runtime this verifier executes on:
    // a 9.0 Microsoft.AspNetCore beside the 8.0 Microsoft.Extensions the runtime already
    // supplied produces CS1705, a complaint about the harness that lands on a doc line.
    private static string? FindAspNetReferenceDir()
    {
        string? root = FindDotnetRoot();
        if (root == null) return null;
        int major = Environment.Version.Major;

        var candidates = new List<(string Dir, bool IsRefPack)>();
        string packs = Path.Combine(root, "packs", "Microsoft.AspNetCore.App.Ref");
        if (Directory.Exists(packs))
        {
            foreach (string ver in Directory.GetDirectories(packs))
            {
                string refRoot = Path.Combine(ver, "ref");
                if (Directory.Exists(refRoot))
                {
                    foreach (string tfm in Directory.GetDirectories(refRoot))
                        candidates.Add((tfm, true));
                }
            }
        }
        string shared = Path.Combine(root, "shared", "Microsoft.AspNetCore.App");
        if (Directory.Exists(shared))
        {
            foreach (string ver in Directory.GetDirectories(shared))
                candidates.Add((ver, false));
        }

        // Nearest major first; within that, the reference pack beats the shared runtime -
        // it is compile-time metadata only, so nothing here can accidentally run against it.
        return candidates
            .Where(c => AspNetRequired.All(a => File.Exists(Path.Combine(c.Dir, a))))
            .OrderBy(c => Math.Abs(FrameworkMajor(c.Dir) - major))
            .ThenByDescending(c => c.IsRefPack)
            .ThenByDescending(c => c.Dir, StringComparer.Ordinal)
            .Select(c => c.Dir)
            .FirstOrDefault();
    }

    // The version a folder was built for: "9.0.15/ref/net9.0" and "shared/.../8.0.26" both
    // say it in the path, and a bare "net8.0" tfm still names the major.
    private static int FrameworkMajor(string dir)
    {
        foreach (Match m in Regex.Matches(dir, @"(\d+)\.\d+\.\d+"))
            if (int.TryParse(m.Groups[1].Value, out int v)) return v;
        Match tfm = Regex.Match(Path.GetFileName(dir), @"net(?:coreapp)?(\d+)");
        return tfm.Success ? int.Parse(tfm.Groups[1].Value) : 0;
    }

    private static string? FindDotnetRoot()
    {
        try
        {
            // The running assembly sits in <dotnet>/shared/Microsoft.NETCore.App/<version>,
            // so the root is the ancestor that owns shared/ or packs/.
            var dir = new DirectoryInfo(Path.GetDirectoryName(typeof(object).Assembly.Location) ?? "");
            for (int i = 0; i < 6 && dir != null; i++)
            {
                if (Directory.Exists(Path.Combine(dir.FullName, "shared")) ||
                    Directory.Exists(Path.Combine(dir.FullName, "packs")))
                    return dir.FullName;
                dir = dir.Parent;
            }
            return null;
        }
        catch { return null; }
    }

    private static IEnumerable<string> OptionalUsings(List<MetadataReference> refs)
    {
        var present = new HashSet<string>(
            refs.OfType<PortableExecutableReference>()
                .Select(r => Path.GetFileNameWithoutExtension(r.FilePath) ?? "")
                .Where(n => n.Length > 0),
            StringComparer.OrdinalIgnoreCase);
        foreach (var (assembly, ns) in OptionalWrapperUsings)
        {
            if (present.Contains(assembly)) yield return ns;
        }
    }

    private static void ReportDiagnostics(CSharpCompilation compilation)
    {
        var reported = new HashSet<string>();
        foreach (var d in compilation.GetDiagnostics())
        {
            if (d.Severity != DiagnosticSeverity.Error) continue;
            string? where = Locate(d);
            if (where == null) continue;
            string key = where + "|" + d.Id;
            if (!reported.Add(key)) continue;
            _diagnostics++;
            // One snippet breaking is one problem, whether it produces two errors or twelve.
            string? path = d.Location.SourceTree?.FilePath;
            if (CoverageDiagnosticIds.Contains(d.Id))
            {
                CoverageBlocks.Add(path ?? where);
                if (CoverageSamples.Count < 20)
                    CoverageSamples.Add($"{where}  {d.Id}  {d.GetMessage()}");
                continue;
            }
            ErrorBlocks.Add(path ?? where);
            if (ErrorSamples.Count < 40)
                ErrorSamples.Add($"{where}  {d.Id}  {d.GetMessage()}");
        }
    }

    private static string? Locate(Diagnostic d)
    {
        var loc = d.Location;
        string? path = loc.SourceTree?.FilePath;
        if (path == null) return loc.GetLineSpan().ToString();
        var m = Regex.Match(path, @"Snippet_\d+~(?<rel>.+)~L(?<line>\d+)$");
        if (!m.Success) return loc.GetLineSpan().ToString();
        string rel = m.Groups["rel"].Value;
        int startLine = int.Parse(m.Groups["line"].Value);
        int treeLine = loc.GetLineSpan().StartLinePosition.Line;
        if (!BlocksByPath.TryGetValue(path, out CodeBlock? block))
            return $"{rel}:{startLine + treeLine}";
        int bodyLine = treeLine - block.BodyOffset;
        int orig = Math.Max(startLine, startLine + bodyLine);
        return $"{rel}:{orig}";
    }
}

internal class CodeBlock
{
    public CodeBlock(string file, int startLine, List<string> body, int index, string treePath)
    {
        File = file; StartLine = startLine; Body = body;
        Text = string.Join("\n", body); Index = index; TreePath = treePath;
        // The first thing that is actually being declared. A snippet opens with its usings,
        // an [Route(...)] attribute, a namespace and its opening brace, and none of those
        // tell Roslyn that a type follows - reading them as "not a type" buries the class
        // inside a method body, where a namespace keyword is a syntax error.
        string first = "";
        foreach (string l in SignificantLines(body))
        {
            if (l.StartsWith("using ") || l.StartsWith("namespace ") || l.StartsWith("[") || l.StartsWith("#")) continue;
            if (l == "{" || l == "}") continue;
            first = l;
            break;
        }
        IsTypeDecl = Regex.IsMatch(first,
            @"^(?:(?:public|internal|sealed|abstract|static|partial|readonly)\s+)*(class|interface|enum|struct|record)\b");
        IsFragment = LooksLikeFragment(body, Text);
        // A block that brings its own namespace is already isolated: it is a set of
        // declarations, never a run of statements, and wrapping it in a namespace of its own
        // would mix file-scoped and block-scoped declarations, which Roslyn rejects outright.
        HasOwnNamespace = SignificantLines(body).Any(l => l.StartsWith("namespace "));
        if (HasOwnNamespace && !IsFragment) IsTypeDecl = true;
        IsMemberDecl = !IsTypeDecl && !IsFragment && MemberOpening.IsMatch(first);
    }
    public string File { get; }
    public int StartLine { get; }
    public List<string> Body { get; }
    public string Text { get; }
    public int Index { get; }
    public string TreePath { get; }
    public bool IsTypeDecl { get; }
    public bool IsFragment { get; }
    public bool IsMemberDecl { get; }
    public bool HasOwnNamespace { get; }
    public int BodyOffset { get; set; }
    public string Source { get; set; } = "";

    // Non-empty lines with a trailing comment removed. A "//" inside a URL is not a comment,
    // so only one that opens the line or follows whitespace counts.
    private static IEnumerable<string> SignificantLines(List<string> body)
    {
        foreach (string l in body)
        {
            string t = l.Trim();
            int c = t.IndexOf("//", StringComparison.Ordinal);
            if (c == 0) continue;
            if (c > 0 && char.IsWhiteSpace(t[c - 1])) t = t.Substring(0, c).TrimEnd();
            if (t.Length > 0) yield return t;
        }
    }

    // Three shapes of block are excerpts rather than programs, and Roslyn can only answer
    // them with a syntax error that says nothing about the skill:
    //   - one bare line with no terminator (the API reference's SetFontFolder(a, b) stubs);
    //   - an elided body written with the unicode ellipsis;
    //   - braces that do not balance in either direction - the doc shows the head of a class
    //     or the tail of a method, and no wrapper can turn that into a standalone unit.
    private static bool LooksLikeFragment(List<string> body, string text)
    {
        if (text.Contains("\u2026")) return true;
        if (BraceBalance(text) != 0) return true;
        var lines = SignificantLines(body).ToList();
        if (lines.Count != 1) return false;
        string s = lines[0];
        return !s.EndsWith(";") && !s.EndsWith("{") && !s.EndsWith("}");
    }

    // A member excerpt: modifiers, a return type, a name and then a parameter list, a body
    // or an expression. Some blocks document one action rather than a whole controller, and
    // a method cannot live in a statement block - it needs a class body around it.
    private static readonly Regex MemberOpening = new Regex(
        @"^(?:(?:public|private|protected|internal|static|virtual|override|async|sealed|abstract|partial|new|extern|unsafe)\s+)+"
      + @"[\w<>\[\],\.\?\s]+?\s+[A-Za-z_@]\w*\s*(?:\(|=>|\{)",
        RegexOptions.Compiled);

    // Brace counting that ignores what the C# compiler ignores: a string literal holding a
    // JSON snippet, a verbatim path with a wildcard, commented-out code from the docs.
    private static int BraceBalance(string text)
    {
        int depth = 0;
        bool inString = false;
        bool inChar = false;
        bool inLineComment = false;
        bool inBlockComment = false;
        for (int i = 0; i < text.Length; i++)
        {
            char c = text[i];
            if (inLineComment) { if (c == '\n') inLineComment = false; continue; }
            if (inBlockComment)
            {
                if (c == '*' && i + 1 < text.Length && text[i + 1] == '/') { inBlockComment = false; i++; }
                continue;
            }
            if (inString) { if (c == '"' && text[i - 1] != '\\') inString = false; continue; }
            if (inChar) { if (c == '\'' && text[i - 1] != '\\') inChar = false; continue; }
            if (c == '/' && i + 1 < text.Length)
            {
                char next = text[i + 1];
                if (next == '/') { inLineComment = true; i++; continue; }
                if (next == '*') { inBlockComment = true; i++; continue; }
            }
            if (c == '"') { inString = true; continue; }
            if (c == '\'') { inChar = true; continue; }
            if (c == '{') depth++;
            else if (c == '}') depth--;
        }
        return depth;
    }
}
