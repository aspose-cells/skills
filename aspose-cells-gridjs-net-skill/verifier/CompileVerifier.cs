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
// Exit code: 0 = all blocks compiled, 1 = at least one compile error.
// ---------------------------------------------------------------------------

internal static class CompileVerifier
{
    private static readonly string[] WrapperUsings =
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

    private static readonly Regex UsingDirective = new Regex(
        @"^\s*using\s+(?:static\s+)?[A-Za-z_][\w.]*\s*;");

    private static int _errors;
    private static readonly List<string> ErrorSamples = new();
    private static readonly Dictionary<string, CodeBlock> BlocksByPath = new();

    public static int Run(string? explicitRoot = null)
    {
        string? skillRoot = explicitRoot ?? FindSkillRoot(AppContext.BaseDirectory);
        if (skillRoot == null)
        {
            Console.Error.WriteLine("Could not locate skill root (folder with SKILL.md and references/).");
            return 1;
        }

        Console.WriteLine($"GridJs version: {typeof(Aspose.Cells.GridJs.GridJsWorkbook).Assembly.GetName().Version}");
        Console.WriteLine($"Compile-scanning: {skillRoot}");
        Console.WriteLine(new string('-', 60));

        var trees = new List<SyntaxTree>();
        var files = Directory.GetFiles(skillRoot, "*.md", SearchOption.AllDirectories)
            .OrderBy(f => f)
            .ToArray();
        int index = 0;
        foreach (string file in files)
        {
            CollectBlocks(file, trees, ref index);
        }

        var refs = BuildReferences();
        var compilation = CSharpCompilation.Create(
            "SkillSnippets",
            trees,
            refs,
            new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary));

        ReportDiagnostics(compilation);

        int total = trees.Count;
        Console.WriteLine(new string('-', 60));
        Console.WriteLine($"COMPILE OK: {total - _errors}   ERRORS: {_errors}   BLOCKS: {total}");
        if (ErrorSamples.Count > 0)
        {
            Console.WriteLine(new string('-', 60));
            Console.WriteLine("Errors (first " + ErrorSamples.Count + "):");
            foreach (string s in ErrorSamples) Console.WriteLine("  " + s);
        }
        return _errors == 0 ? 0 : 1;
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

    private static void CollectBlocks(string file, List<SyntaxTree> trees, ref int index)
    {
        string[] lines = File.ReadAllLines(file);
        string rel = Path.GetRelativePath(
            Path.GetDirectoryName(Path.GetDirectoryName(file)) ?? ".", file);
        for (int i = 0; i < lines.Length; i++)
        {
            string t = lines[i].TrimStart();
            if (!t.StartsWith("```")) continue;
            var m = Regex.Match(t, @"^```\s*(csharp|cs)\s*$");
            if (!m.Success) continue;

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

    private static string Wrap(CodeBlock block)
    {
        var lines = new List<string>();
        lines.AddRange(WrapperUsings);
        lines.Add("");
        lines.Add("#nullable disable");

        if (block.IsTypeDecl)
        {
            block.BodyOffset = lines.Count;
            foreach (string ln in block.Body)
            {
                if (UsingDirective.IsMatch(ln)) continue;
                lines.Add(ln);
            }
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
        foreach (string path in tpa.Split(Path.PathSeparator))
        {
            if (path.Length == 0 || !seen.Add(path)) continue;
            if (!File.Exists(path)) continue;
            try { refs.Add(MetadataReference.CreateFromFile(path)); } catch { }
        }
        return refs;
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
            _errors++;
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
        IsTypeDecl = Regex.IsMatch(FirstCode(body),
            @"^(?:(?:public|internal|sealed|abstract|static|partial|readonly)\s+)*(class|interface|enum|struct|record)\b");
    }
    public string File { get; }
    public int StartLine { get; }
    public List<string> Body { get; }
    public string Text { get; }
    public int Index { get; }
    public string TreePath { get; }
    public bool IsTypeDecl { get; }
    public int BodyOffset { get; set; }
    public string Source { get; set; } = "";
    private static string FirstCode(List<string> body)
    {
        foreach (string l in body) { string t = l.Trim(); if (t.Length > 0) return t; }
        return "";
    }
}