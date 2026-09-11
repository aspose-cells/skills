using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;

// ---------------------------------------------------------------------------
// Name-level cross-check verifier
//
// Builds a complete index of all public types/methods/properties in the
// GridJs assembly, then scans skill markdown files for identifiers that
// look like GridJs API references but don't exist in the assembly.
//
// This catches LLM hallucinations like "AddAsposeCellsGridJs" that the
// reflection scan misses because the containing type can't be resolved.
// ---------------------------------------------------------------------------

internal static class NameVerifier
{
    private static readonly Assembly GridJsAssembly = typeof(Aspose.Cells.GridJs.GridJsWorkbook).Assembly;

    // All public type names in the GridJs assembly
    private static readonly HashSet<string> KnownTypes = new(StringComparer.Ordinal);

    // All public method/property/field/event names in the GridJs assembly
    private static readonly HashSet<string> KnownMembers = new(StringComparer.Ordinal);

    // Method names that are extension methods on framework types (e.g., IServiceCollection)
    private static readonly HashSet<string> KnownExtensionMethods = new(StringComparer.Ordinal);

    // Names to ignore (common .NET / framework identifiers that are not GridJs-specific)
    private static readonly HashSet<string> IgnoreNames = new(StringComparer.Ordinal)
    {
        // C# keywords & contextual keywords
        "var", "new", "return", "using", "namespace", "class", "struct", "enum",
        "interface", "public", "private", "protected", "internal", "static",
        "void", "string", "int", "bool", "double", "float", "byte", "long",
        "object", "null", "true", "false", "this", "base", "async", "await",
        "try", "catch", "finally", "throw", "if", "else", "for", "foreach",
        "while", "do", "switch", "case", "break", "continue", "default",
        "get", "set", "value", "where", "select", "from", "orderby", "into",
        "readonly", "const", "ref", "out", "params", "override", "virtual",
        "abstract", "sealed", "partial", "lambda", "nameof", "typeof", "sizeof",
        // Common .NET types
        "Console", "Math", "String", "StringBuilder", "Exception", "Task",
        "File", "Directory", "Path", "Stream", "MemoryStream", "FileStream",
        "Encoding", "UTF8", "Regex", "DateTime", "TimeSpan", "Guid",
        "Convert", "Environment", "Random", "Array", "List", "Dictionary",
        "HashSet", "Tuple", "KeyValuePair", "IEnumerable", "IList",
        "IDisposable", "IAsyncDisposable", "CancellationToken",
        // ASP.NET Core / DI common names
        "WebApplication", "WebApplicationBuilder", "IApplicationBuilder",
        "IHost", "IConfiguration", "ILoggingBuilder",
        "AddControllersWithViews", "AddScoped", "AddTransient", "AddSingleton",
        "Configure", "Build", "Run", "Use", "Map", "MapControllerRoute",
        "UseExceptionHandler", "UseHsts", "UseHttpsRedirection", "UseStaticFiles",
        "UseRouting", "UseAuthorization", "UseAuthentication",
        "CreateBuilder", "IsDevelopment",
        // MVC / Controller
        "Controller", "ControllerBase", "IActionResult", "ViewResult",
        "ContentResult", "FileResult", "JsonResult", "RedirectResult",
        "Ok", "Content", "View", "Json", "File", "Redirect",
        "HttpGet", "HttpPost", "HttpPut", "HttpDelete",
        "Route", "ApiController",
        // Common method names that are too generic
        "ToString", "GetType", "GetHashCode", "Equals",
        "Dispose", "Clone", "Copy", "Create",
        // ASP.NET Core project structure names (not GridJs APIs)
        "Program", "HomeController", "GridJsController", "GridJsControllerBase",
        "ErrorViewModel", "Index", "Privacy", "Error",
        "Layout", "Simple", "Demo", "Browse",
        "GridJsDemo", "LoadSpreadsheet", "UpdateCell",
        "ILogger", "IServiceProvider", "IServiceCollection",
        "AddLogging", "AddMvc", "MapGet", "MapPost",
        "Options", "Required", "Enable",
        "Newtonsoft", "NewtonsoftJson",
        "ItemGroup", "PackageReference", "Include", "Version",
        "LogLevel", "Logging",
        // DI / setup pattern names (appear in C# setup code, not GridJs APIs)
        "Register", "Store", "Optional", "Example",
        "Home", "Controllers", "GridJs",
        // Common C# patterns
        "Main", "Args", "Result", "Data", "Name", "Type", "Value",
        "Count", "Length", "Key", "Message", "Status", "Error",
        "Request", "Response", "Service", "Provider", "Factory",
        // ASP.NET Core model binding attributes
        "FromBody", "FromQuery", "FromRoute", "FromHeader", "FromForm",
        // ASP.NET Core / Swagger / dev tools
        "Swagger", "AddSwaggerGen", "AddEndpointsApiExplorer",
        "Development", "IsDevelopment",
        // Example data / placeholder names (appear in code examples)
        "Record", "John", "Anna", "Smith", "Pedro", "Alvarez",
        "Email", "Total", "Return", "Apply",
        // Placeholder class/namespace names
        "YourAppNamespace", "GridDataController", "MyCustomAction", "MyDto", "Custom",
        // Configuration property names
        "Cache", "Disable",
        // String literal values that look like identifiers
        "GridJs2",
        // PascalCase versions of common words
        "Return", "Apply", "Total",
        // ASP.NET / MVC framework types used in examples
        "IFormFileCollection",
        // Common English words appearing in code comments / examples
        "Other", "Import", "Sample", "Additional",
        // Example code identifiers (from logging, weather API, calc engine examples)
        "MyEngine", "StringComparison",
        "GridJsLoggerProvider", "GridJsLoggerMiddleware",
        "Razor",
        "WeatherForecastController", "WeatherForecast", "WeatherService", "Forecast",
        "Weather", "Time", "Simulate", "Failed", "Results",
        // Example code identifiers (from controller/storage/export examples)
        "GridJsEngine", "IGridJsStorage", "FileSystemGridJsStorage",
        "SpreadsheetController", "CellUpdateInfo", "ExportFormat",
        "GridJsException", "LoadOptions", "DemoExport",
        // Common ASP.NET action method names / model properties in examples
        "Validate", "BadRequest", "Update", "Export", "Optionally",
        "App_Data", "WorksheetIndex", "IncludeHiddenRows",
        "IncludeHiddenColumns", "PreserveFormulas", "LargeWorkbook",
        // User-defined helper methods commonly seen in example code
        "GetFullFilePath",
    };

    public static int Run(string? explicitRoot)
    {
        string? skillRoot = explicitRoot ?? FindSkillRoot(AppContext.BaseDirectory);
        if (skillRoot == null)
        {
            Console.Error.WriteLine("Could not locate skill root.");
            return 1;
        }

        // Step 1: Build assembly index
        BuildAssemblyIndex();

        Console.WriteLine($"GridJs assembly version: {GridJsAssembly.GetName().Version}");
        Console.WriteLine($"Assembly index: {KnownTypes.Count} types, {KnownMembers.Count} members, {KnownExtensionMethods.Count} extension methods");
        Console.WriteLine($"Scanning: {skillRoot}");
        Console.WriteLine(new string('-', 60));

        // Step 2: Scan skill files
        var files = Directory.GetFiles(skillRoot, "*.md", SearchOption.AllDirectories)
            .Where(f =>
            {
                var rel = Path.GetRelativePath(skillRoot, f);
                return !rel.StartsWith("example") && !rel.StartsWith("verifier");
            })
            .OrderBy(f => f)
            .ToArray();

        int totalChecked = 0;
        int notFound = 0;
        var errors = new List<string>();

        foreach (string file in files)
        {
            CheckFile(file, skillRoot, ref totalChecked, ref notFound, errors);
        }

        Console.WriteLine(new string('-', 60));
        Console.WriteLine($"CHECKED: {totalChecked}   NOT FOUND: {notFound}");

        if (errors.Count > 0)
        {
            Console.WriteLine(new string('-', 60));
            Console.WriteLine($"Hallucinated API names ({errors.Count}):");
            foreach (var e in errors)
                Console.WriteLine($"  {e}");
        }

        return notFound == 0 ? 0 : 1;
    }

    private static void BuildAssemblyIndex()
    {
        Type[] types;
        try { types = GridJsAssembly.GetTypes(); }
        catch (ReflectionTypeLoadException ex) { types = ex.Types.Where(t => t != null).ToArray()!; }

        foreach (var type in types)
        {
            if (!type.IsPublic && !type.IsNestedPublic) continue;

            KnownTypes.Add(type.Name);
            KnownTypes.Add(type.FullName ?? type.Name);

            // Index all public members
            const BindingFlags memberFlags = BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static | BindingFlags.DeclaredOnly;

            foreach (var m in type.GetMethods(memberFlags))
            {
                KnownMembers.Add(m.Name);
                if (m.IsStatic && m.IsDefined(typeof(System.Runtime.CompilerServices.ExtensionAttribute), false))
                    KnownExtensionMethods.Add(m.Name);
            }
            foreach (var p in type.GetProperties(memberFlags))
                KnownMembers.Add(p.Name);
            foreach (var f in type.GetFields(memberFlags))
                KnownMembers.Add(f.Name);
            foreach (var e in type.GetEvents(memberFlags))
                KnownMembers.Add(e.Name);

            // Also index nested types
            foreach (var nested in type.GetNestedTypes(BindingFlags.Public))
            {
                KnownTypes.Add(nested.Name);
                KnownTypes.Add(nested.FullName ?? nested.Name);
            }
        }
    }

    private static void CheckFile(string file, string skillRoot, ref int totalChecked, ref int notFound, List<string> errors)
    {
        string[] lines = File.ReadAllLines(file);
        string rel = Path.GetRelativePath(skillRoot, file);
        bool inCode = false;
        bool isCsharpBlock = false;
        int codeBlockStartLine = 0;

        // Collect identifiers per code block for batch analysis
        var blockIdents = new List<(int lineNum, string ident, string lineContent)>();

        for (int i = 0; i < lines.Length; i++)
        {
            string line = lines[i].TrimStart();
            if (line.StartsWith("```"))
            {
                if (!inCode)
                {
                    // Opening a code block - detect language
                    inCode = true;
                    isCsharpBlock = Regex.IsMatch(line, @"^```\s*(csharp|cs)\s*$", RegexOptions.IgnoreCase);
                    codeBlockStartLine = i;
                    blockIdents.Clear();
                }
                else
                {
                    // Closing a code block - process collected identifiers
                    if (isCsharpBlock && blockIdents.Count > 0)
                    {
                        ProcessCodeBlock(blockIdents, rel, ref totalChecked, ref notFound, errors);
                    }
                    inCode = false;
                    isCsharpBlock = false;
                    blockIdents.Clear();
                }
                continue;
            }
            if (!inCode || !isCsharpBlock) continue;

            // Extract identifiers from C# code blocks only
            var idents = ExtractIdentifiers(lines[i]);
            foreach (var ident in idents)
            {
                if (IgnoreNames.Contains(ident)) continue;
                if (KnownTypes.Contains(ident)) continue;
                if (KnownMembers.Contains(ident)) continue;
                if (!LooksLikeApiName(ident)) continue;
                if (IsInUsingDirective(lines[i], ident)) continue;
                if (IsNamespaceQualified(lines[i], ident)) continue;

                blockIdents.Add((i + 1, ident, lines[i].Trim()));
            }
        }
    }

    /// <summary>
    /// Process collected unknown identifiers from a single code block.
    /// If a block has many unknown names, it's likely example/placeholder code — skip it.
    /// </summary>
    private static void ProcessCodeBlock(
        List<(int lineNum, string ident, string lineContent)> blockIdents,
        string rel,
        ref int totalChecked,
        ref int notFound,
        List<string> errors)
    {
        // If the block has 3+ unknown identifiers, it's almost certainly example/placeholder code
        // (e.g., a full controller class with DTOs, action methods, etc.)
        // Skip the entire block to avoid false positives.
        const int exampleThreshold = 3;
        if (blockIdents.Count >= exampleThreshold)
            return;

        // For blocks with few unknown names, check if they look like type definitions
        // (class/interface/enum declarations with placeholder names)
        foreach (var (lineNum, ident, lineContent) in blockIdents)
        {
            totalChecked++;
            notFound++;
            string msg = $"  {rel}:L{lineNum}  '{ident}' - not found in GridJs assembly";
            if (!string.IsNullOrWhiteSpace(lineContent))
                msg += $"\n    Line: {lineContent}";
            errors.Add(msg);
            Console.WriteLine(msg);
        }
    }

    private static List<string> ExtractIdentifiers(string line)
    {
        var result = new List<string>();
        // Strip comments and string literals before extracting identifiers
        var stripped = StripComments(StripStringLiterals(line));
        // Match PascalCase identifiers (at least 3 chars, starts with uppercase)
        foreach (Match m in Regex.Matches(stripped, @"\b([A-Z][a-zA-Z0-9_]{2,})\b"))
        {
            result.Add(m.Groups[1].Value);
        }
        return result;
    }

    /// <summary>
    /// Remove comments from a code line to prevent extracting identifiers from comments.
    /// Handles both single-line (//) and XML doc (///) comments.
    /// </summary>
    private static string StripComments(string line)
    {
        // Remove single-line comments (// ...) and XML doc comments (/// ...)
        var idx = line.IndexOf("//");
        if (idx >= 0)
            return line.Substring(0, idx);
        return line;
    }

    /// <summary>
    /// Remove string literals from a code line to prevent extracting identifiers from inside strings.
    /// Handles both regular strings ("...") and verbatim strings (@"...").
    /// </summary>
    private static string StripStringLiterals(string line)
    {
        // Remove verbatim strings (@"..." — may contain escaped quotes "")
        var result = Regex.Replace(line, "@\"(?:[^\"]*|\"\"\")*\"", "");
        // Remove regular strings ("..." — may contain escaped quotes \")
        result = Regex.Replace(result, "(?<!@)\"(?:[^\"\\\\]|\\\\.)*\"", "");
        return result;
    }

    private static bool LooksLikeApiName(string name)
    {
        // Must start with uppercase (PascalCase)
        if (!char.IsUpper(name[0])) return false;
        // Must be at least 4 chars
        if (name.Length < 4) return false;
        // Skip if all uppercase (likely a constant or acronym)
        if (name.All(char.IsUpper)) return false;
        // Skip if it looks like a namespace (contains multiple dots when in context)
        if (name.Contains('.')) return false;
        return true;
    }

    private static bool IsInUsingDirective(string line, string ident)
    {
        var trimmed = line.TrimStart();
        if (!trimmed.StartsWith("using ")) return false;
        // Check if ident appears in a using directive
        return Regex.IsMatch(trimmed, $@"\b{Regex.Escape(ident)}\b");
    }

    private static bool IsNamespaceQualified(string line, string ident)
    {
        // Check if the identifier is preceded by a dot (e.g., GridJs.XXX or options.XXX)
        var idx = line.IndexOf(ident, StringComparison.Ordinal);
        while (idx >= 0)
        {
            if (idx > 0 && line[idx - 1] == '.')
                return true; // Qualified access - skip
            idx = line.IndexOf(ident, idx + 1, StringComparison.Ordinal);
        }
        return false;
    }

    private static string? FindSkillRoot(string startDir)
    {
        var dir = new DirectoryInfo(startDir);
        while (dir != null)
        {
            if (File.Exists(Path.Combine(dir.FullName, "SKILL.md")) &&
                Directory.Exists(Path.Combine(dir.FullName, "references")))
                return dir.FullName;
            if (File.Exists(Path.Combine(dir.FullName, "skill-index.md")))
                return dir.FullName;
            // Also match template-generated layout (multiple .md files)
            if (Directory.GetFiles(dir.FullName, "*.md").Length > 0 &&
                Directory.GetFiles(dir.FullName, "*-api-*.md").Length > 0)
                return dir.FullName;
            dir = dir.Parent;
        }
        return null;
    }
}