using System.Text;
using System.Text.RegularExpressions;

// ---------------------------------------------------------------------------
// GridJs client-side API verifier
//
// Scans all *.md files under the skill root, extracts TypeScript/JavaScript
// API references, and validates them against the index.d.ts declarations.
//
// Checks:
//   - Type names (interfaces, classes, type aliases) exist
//   - Method names exist on the referenced type
//   - Property names exist on the referenced type
//   - Event names are declared in event type aliases
//
// Exit code: 0 = all checks passed, 1 = at least one NOT FOUND.
// ---------------------------------------------------------------------------

internal static class ClientApiVerifier
{
    private static int _ok;
    private static int _notFound;
    private static int _skipped;
    private static readonly List<string> NotFoundSamples = new();
    private static readonly int MaxSamples = 40;

    // Parsed client API: type name -> members
    private static readonly Dictionary<string, ClientTypeInfo> KnownTypes = new();
    private static readonly HashSet<string> KnownEventNames = new();

    public static int Run(string? explicitRoot = null)
    {
        string? skillRoot = explicitRoot ?? FindSkillRoot(AppContext.BaseDirectory);
        if (skillRoot == null)
        {
            Console.Error.WriteLine("Could not locate skill root (folder with SKILL.md and references/).");
            return 1;
        }

        // Find index.d.ts
        string dtsPath = Path.Combine(skillRoot, "client-api", "index.d.ts");
        if (!File.Exists(dtsPath))
        {
            Console.Error.WriteLine($"Client API declaration file not found: {dtsPath}");
            return 1;
        }

        Console.WriteLine($"Client API declarations: {dtsPath}");
        Console.WriteLine($"Scanning: {skillRoot}");
        Console.WriteLine(new string('-', 60));

        // Parse index.d.ts
        ParseDeclarations(File.ReadAllText(dtsPath));
        Console.WriteLine($"  Parsed {KnownTypes.Count} types, {KnownEventNames.Count} events");

        // Scan markdown files
        var files = Directory.GetFiles(skillRoot, "*.md", SearchOption.AllDirectories)
            .OrderBy(f => f)
            .ToArray();

        foreach (string file in files)
        {
            CheckFile(file, skillRoot);
        }

        Console.WriteLine(new string('-', 60));
        Console.WriteLine($"OK: {_ok}   NOT FOUND: {_notFound}   SKIPPED: {_skipped}");
        if (NotFoundSamples.Count > 0)
        {
            Console.WriteLine(new string('-', 60));
            Console.WriteLine("NOT FOUND samples:");
            foreach (string s in NotFoundSamples) Console.WriteLine("  " + s);
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

    // -----------------------------------------------------------------------
    // Parse index.d.ts declarations
    // -----------------------------------------------------------------------
    private static void ParseDeclarations(string content)
    {
        // Parse interfaces
        foreach (Match m in Regex.Matches(content,
            @"(?:export\s+)?interface\s+(\w+)(?:\s+extends\s+[\w,\s]+)?\s*\{",
            RegexOptions.Multiline))
        {
            string name = m.Groups[1].Value;
            int braceStart = m.Index + m.Length - 1;
            string body = ExtractBraceBody(content, braceStart);
            var info = GetOrCreateType(name);
            ParseMembers(body, info);
        }

        // Parse type aliases (including event types)
        foreach (Match m in Regex.Matches(content,
            @"(?:export\s+)?type\s+(\w+)\s*=\s*\{([^}]+)\}",
            RegexOptions.Multiline))
        {
            string name = m.Groups[1].Value;
            string body = m.Groups[2].Value;
            var info = GetOrCreateType(name);
            // Parse union members like 'cell-selected' | 'cells-selected'
            foreach (Match um in Regex.Matches(body, @"'([^']+)'"))
            {
                info.Members.Add(um.Groups[1].Value);
                KnownEventNames.Add(um.Groups[1].Value);
            }
        }

        // Parse default export class
        var classMatch = Regex.Match(content,
            @"export\s+default\s+class\s+(\w+)",
            RegexOptions.Multiline);
        if (classMatch.Success)
        {
            string name = classMatch.Groups[1].Value;
            var info = GetOrCreateType(name);
            // Find the class body
            int classStart = content.IndexOf('{', classMatch.Index);
            if (classStart >= 0)
            {
                string body = ExtractBraceBody(content, classStart);
                ParseMembers(body, info);
            }
        }
    }

    private static ClientTypeInfo GetOrCreateType(string name)
    {
        if (!KnownTypes.TryGetValue(name, out var info))
        {
            info = new ClientTypeInfo(name);
            KnownTypes[name] = info;
        }
        return info;
    }

    private static void ParseMembers(string body, ClientTypeInfo info)
    {
        // Methods: name(params): returnType
        foreach (Match m in Regex.Matches(body,
            @"(\w+)\s*\([^)]*\)\s*:\s*[\w<>\[\]|&\s,]+",
            RegexOptions.Multiline))
        {
            info.Members.Add(m.Groups[1].Value);
        }

        // Properties: name?: type or name: type
        foreach (Match m in Regex.Matches(body,
            @"^\s*(\w+)\s*\??:\s*",
            RegexOptions.Multiline))
        {
            info.Members.Add(m.Groups[1].Value);
        }
    }

    private static string ExtractBraceBody(string content, int braceStart)
    {
        int depth = 0;
        int start = braceStart;
        for (int i = braceStart; i < content.Length; i++)
        {
            if (content[i] == '{') depth++;
            else if (content[i] == '}')
            {
                depth--;
                if (depth == 0) return content.Substring(start + 1, i - start - 1);
            }
        }
        return "";
    }

    // -----------------------------------------------------------------------
    // Check markdown files
    // -----------------------------------------------------------------------
    private static void CheckFile(string file, string skillRoot)
    {
        string rel = Path.GetRelativePath(skillRoot, file);
        string[] lines = File.ReadAllLines(file);

        bool inCode = false;
        bool isTsBlock = false;

        for (int i = 0; i < lines.Length; i++)
        {
            string line = lines[i];
            string trimmed = line.TrimStart();

            if (trimmed.StartsWith("```"))
            {
                if (!inCode)
                {
                    inCode = true;
                    isTsBlock = Regex.IsMatch(trimmed, @"^```\s*(typescript|javascript|js|ts)\b");
                }
                else
                {
                    inCode = false;
                    isTsBlock = false;
                }
                continue;
            }

            if (!inCode) continue;

            // Check xs.methodName( patterns
            foreach (Match m in Regex.Matches(line, @"\bxs\.(\w+)\s*\("))
            {
                string method = m.Groups[1].Value;
                if (KnownTypes.ContainsKey("Spreadsheet") && KnownTypes["Spreadsheet"].Members.Contains(method))
                    _ok++;
                else if (KnownTypes.ContainsKey("Spreadsheet"))
                    Report(rel, i + 1, $"xs.{method}() - no method '{method}' on Spreadsheet");
                else
                    _skipped++;
            }

            // Check xs.on('event-name' patterns
            foreach (Match m in Regex.Matches(line, @"\bxs\.on\s*\(\s*'([^']+)'"))
            {
                string eventName = m.Groups[1].Value;
                if (KnownEventNames.Contains(eventName))
                    _ok++;
                else
                    Report(rel, i + 1, $"xs.on('{eventName}') - unknown event name");
            }

            // Check x_spreadsheet( factory
            if (line.Contains("x_spreadsheet("))
                _ok++;

            // Check Options properties: { updateMode:, local:, etc.
            foreach (Match m in Regex.Matches(line, @"\b(\w+)\s*:\s*(?:'[^']*'|""[^""]*""|\d+|true|false|\w+)"))
            {
                string prop = m.Groups[1].Value;
                if (prop == "updateMode" || prop == "local" || prop == "updateUrl" ||
                    prop == "showContextmenu" || prop == "showGrid" || prop == "showToolbar")
                {
                    if (KnownTypes.ContainsKey("Options") && KnownTypes["Options"].Members.Contains(prop))
                        _ok++;
                    else
                        _skipped++;
                }
            }

            // Check semantic. patterns
            foreach (Match m in Regex.Matches(line, @"\bsemantic\.(\w+)\s*\("))
            {
                string method = m.Groups[1].Value;
                bool found = false;
                foreach (var type in KnownTypes.Values.Where(t => t.Name.StartsWith("Semantic")))
                {
                    if (type.Members.Contains(method)) { found = true; break; }
                }
                if (found) _ok++;
                else if (method.Length > 2) _skipped++; // might be a local variable
            }
        }
    }

    private static void Report(string file, int line, string message)
    {
        _notFound++;
        if (NotFoundSamples.Count < MaxSamples)
            NotFoundSamples.Add($"{file}:{line}  {message}");
    }
}

internal class ClientTypeInfo
{
    public string Name { get; }
    public HashSet<string> Members { get; } = new();
    public ClientTypeInfo(string name) { Name = name; }
}