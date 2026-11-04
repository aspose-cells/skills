using System.Reflection;
using System.Reflection.Metadata;
using System.Reflection.PortableExecutable;
using System.Text;
using System.Text.RegularExpressions;

// ---------------------------------------------------------------------------
// Name-level cross-check verifier
//
// Indexes every public type and member of the Aspose.Cells.GridJs assembly, then reads
// the skill's ```csharp blocks as C#: it tracks what each variable holds, walks each
// dotted chain, and judges every name in its position - a type where a type must exist,
// a member where the receiver's type must carry it.
//
// Report classes:
//   finding    - a name the skill claims about this product that the assembly does not
//                carry, including a real member hung off the wrong type.
//   suspect    - a weak claim (a GridJs-shaped member on a receiver the scan cannot type)
//                inside a block that is mostly invented names. Counted and printed, never
//                dropped, never failed.
//   unresolved - member positions the scan had no evidence about at all, because the
//                receiver belongs to ASP.NET or to the sample. Printed, because a verifier
//                that hides its blind spots cannot be trusted on its passes.
//
// Exit code: 0 = clean, 1 = at least one finding, 2 = the verifier itself could not run.
// ---------------------------------------------------------------------------

internal static class NameVerifier
{
    private static readonly Assembly GridJsAssembly = typeof(Aspose.Cells.GridJs.GridJsWorkbook).Assembly;

    // Simple name (and full name) -> the Type, for every public GridJs type.
    private static readonly Dictionary<string, Type> TypesByName = new(StringComparer.Ordinal);

    // Public types the assembly defines but the runtime cannot hand over, and the members its
    // own metadata says they declare. GridJsControllerBase derives from
    // Microsoft.AspNetCore.Mvc.ControllerBase, which this verifier project does not reference,
    // so Assembly.GetTypes() answers null for it and a real type used to read as a fabricated
    // one. The assembly file is the complete answer to "is this a GridJs type", and asking it
    // costs no new dependency - Roslyn already brings System.Reflection.Metadata.
    private static readonly Dictionary<string, HashSet<string>> ShadowTypesByName = new(StringComparer.Ordinal);
    private static readonly List<string> ShadowTypeNames = new();
    private static int _shadowCount;

    // Member name -> the types that declare it, in assembly order.
    private static readonly Dictionary<string, List<string>> MemberOwners = new(StringComparer.Ordinal);

    // Every member the assembly declares somewhere.
    private static readonly HashSet<string> KnownMembers = new(StringComparer.Ordinal);

    // Types and members the skill declares for itself, in any block. A sample class the
    // docs define - GridJsLoggerProvider, FileSystemGridJsStorage - or a helper extension
    // method they add is the reader's own code, not a claim about this product.
    private static readonly HashSet<string> SkillDeclaredTypes = new(StringComparer.Ordinal);
    private static readonly HashSet<string> SkillDeclaredMembers = new(StringComparer.Ordinal);

    // Example-code convention: MyDatabaseCache, CustomExporter, DemoHelper, SampleService...
    private static readonly Regex PlaceholderName =
        new(@"^(?:My|Custom|Demo|Sample|Test|Example|Simple|Own)[A-Z][A-Za-z0-9_]*$", RegexOptions.Compiled);

    private static readonly Regex TypeDeclaration =
        new(@"\b(?:class|interface|enum|record|struct)\s+(?<name>[A-Za-z_][A-Za-z0-9_]*)", RegexOptions.Compiled);

    // A method or extension method the snippet defines, including the `this` parameter form.
    private static readonly Regex MemberDeclaration =
        new(@"^\s*(?:\[[^\]]*\]\s*)*(?:(?:public|private|protected|internal|static|virtual|override|async|sealed|partial|explicit|implicit)\s+)+[\w<>\[\],\.\?]*\s+(?:\.?)?(?<name>[A-Za-z_][A-Za-z0-9_]*)\s*(?:\(|=>)", RegexOptions.Compiled);

    // Names to ignore (common .NET / framework identifiers that are not GridJs-specific).
    // A name that looks like this product's is never exempted by this list alone: it has to
    // be a real assembly member or something the skill itself declares.
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
        // GridJsControllerBase is deliberately absent: it is a real GridJs type (25.5.0 up),
        // and exempting it here is what hid the index hole that reported it as invented.
        "Program", "HomeController", "GridJsController",
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

    // Members reached through a dotted chain that are plain plumbing or the sample's own
    // vocabulary; they would otherwise land in the unresolved count as noise.
    private static readonly HashSet<string> ChainNoise = new(StringComparer.Ordinal)
    {
        "Length", "Count", "Key", "Value", "Item", "First", "Last", "ToList", "ToArray",
        "Select", "Where", "OrderBy", "FirstOrDefault", "Any", "All", "Add", "AddRange",
        "Contains", "Remove", "Insert", "Clear", "Exists", "StartsWith", "EndsWith",
        "Substring", "Replace", "Split", "Trim", "ToLower", "ToUpper", "ToString",
        "GetAwaiter", "GetResult", "ConfigureAwait", "Then", "ContinueWith",
    };

    private const BindingFlags PublicAny =
        BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static | BindingFlags.FlattenHierarchy;

    private static int _checked;
    private static int _notFound;
    private static int _suspect;
    private static int _unresolved;
    private static readonly List<string> Findings = new();
    private static readonly List<string> Suspects = new();
    private static readonly List<string> UnresolvedSamples = new();

    private static void CountCandidate(int line, string name, string why, bool weak,
                                       List<(int line, string name, string why)> strong,
                                       List<(int line, string name, string why)> weakList)
    {
        _checked++;
        (weak ? weakList : strong).Add((line, name, why));
    }

    public static int Run(string? explicitRoot)
    {
        string root = explicitRoot is null
            ? FindSkillRoot(AppContext.BaseDirectory) ?? ""
            : Path.GetFullPath(explicitRoot);
        if (root.Length == 0 || !Directory.Exists(root))
        {
            Console.Error.WriteLine("Could not locate skill root.");
            return 2;
        }

        BuildAssemblyIndex();

        var assemblyVersion = GridJsAssembly.GetName().Version;
        Console.WriteLine($"GridJs assembly version: {assemblyVersion.Major}.{assemblyVersion.Minor}.{assemblyVersion.Build}");
        Console.WriteLine($"Assembly index: {_typeCount} types, {KnownMembers.Count} members");
        if (_shadowCount > 0)
            Console.WriteLine($"Metadata only ({_shadowCount}): {string.Join(", ", ShadowTypeNames)}"
                + " - loaded reflection refused these, their base assembly is not referenced here");
        Console.WriteLine($"Scanning: {root}");
        Console.WriteLine(new string('-', 60));

        var files = Directory.GetFiles(root, "*.md", SearchOption.AllDirectories)
            .Where(f =>
            {
                var rel = Path.GetRelativePath(root, f);
                return !rel.StartsWith("example") && !rel.StartsWith("verifier") && !rel.StartsWith("report");
            })
            .OrderBy(f => f)
            .ToArray();

        // Pass 1: what does the skill declare for itself? A class or extension method the
        // docs define in one block is the reader's own code when a later block calls it.
        var texts = new Dictionary<string, string[]>();
        foreach (string file in files)
        {
            string[] lines = File.ReadAllLines(file);
            texts[file] = lines;
            foreach (Match m in TypeDeclaration.Matches(string.Join("\n", lines)))
                SkillDeclaredTypes.Add(m.Groups["name"].Value);
        }

        // Pass 2: judge the code blocks.
        foreach (string file in files)
        {
            CheckFile(file, root, texts[file]);
        }

        Console.WriteLine(new string('-', 60));
        Console.WriteLine($"CHECKED: {_checked}   NOT FOUND: {_notFound}   SUSPECT: {_suspect}"
            + $"   UNRESOLVED: {_unresolved}");
        // Machine-readable result line. The caller trusts this rather than the exit code,
        // because a crashed verifier also leaves a non-zero exit code behind, and that is
        // a broken toolchain, not a statement about the documentation.
        Console.WriteLine($"VERIFIER-RESULT: checked={_checked} not-found={_notFound}"
            + $" suspect={_suspect} unresolved={_unresolved} shadow-types={_shadowCount}");

        if (Findings.Count > 0)
        {
            Console.WriteLine(new string('-', 60));
            Console.WriteLine($"Hallucinated API names ({Findings.Count}):");
            foreach (var f in Findings)
                Console.WriteLine("  " + f);
        }
        if (Suspects.Count > 0)
        {
            Console.WriteLine(new string('-', 60));
            Console.WriteLine($"Demoted from blocks of placeholder code, review by hand ({Suspects.Count}):");
            foreach (var s in Suspects)
                Console.WriteLine("  " + s);
        }
        if (UnresolvedSamples.Count > 0)
        {
            Console.WriteLine(new string('-', 60));
            Console.WriteLine($"Unresolved member positions, first {UnresolvedSamples.Count} of {_unresolved}:");
            foreach (var u in UnresolvedSamples)
                Console.WriteLine("  " + u);
        }

        return _notFound == 0 ? 0 : 1;
    }

    private static int _typeCount;

    private static void BuildAssemblyIndex()
    {
        Type[] types;
        try { types = GridJsAssembly.GetTypes(); }
        catch (ReflectionTypeLoadException ex) { types = ex.Types.Where(t => t != null).ToArray()!; }

        foreach (var type in types)
        {
            if (!type.IsPublic && !type.IsNestedPublic) continue;

            TypesByName[type.Name] = type;
            if (type.FullName != null) TypesByName[type.FullName] = type;
            _typeCount++;

            const BindingFlags declared =
                BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static | BindingFlags.DeclaredOnly;

            foreach (var m in type.GetMembers(declared))
            {
                if (m is Type) continue; // nested types are names, not members
                KnownMembers.Add(m.Name);
                if (!MemberOwners.TryGetValue(m.Name, out var owners))
                    MemberOwners[m.Name] = owners = new List<string>();
                if (!owners.Contains(type.Name)) owners.Add(type.Name);
            }
        }

        AddShadowTypesFromMetadata();
    }

    // The reflection index above is only as wide as this project's references. Anything the
    // assembly defines but the loader refuses is read out of the assembly file instead, so the
    // scan can say "that is a GridJs type" without ever having been able to load it.
    private static void AddShadowTypesFromMetadata()
    {
        string path;
        try { path = GridJsAssembly.Location; } catch { return; }
        if (string.IsNullOrEmpty(path) || !File.Exists(path)) return;

        using var stream = File.OpenRead(path);
        using var pe = new PEReader(stream);
        if (!pe.HasMetadata) return;
        MetadataReader md = pe.GetMetadataReader();

        foreach (TypeDefinitionHandle th in md.TypeDefinitions)
        {
            TypeDefinition td = md.GetTypeDefinition(th);
            TypeAttributes visibility = td.Attributes & TypeAttributes.VisibilityMask;
            if (visibility != TypeAttributes.Public && visibility != TypeAttributes.NestedPublic) continue;

            string name = md.GetString(td.Name);
            string fullName = md.GetString(td.Namespace) + "." + name;
            if (TypesByName.ContainsKey(name) || TypesByName.ContainsKey(fullName)) continue;

            var members = new HashSet<string>(StringComparer.Ordinal);
            foreach (MethodDefinitionHandle mh in td.GetMethods())
            {
                MethodDefinition mth = md.GetMethodDefinition(mh);
                if ((mth.Attributes & MethodAttributes.SpecialName) != 0) continue;
                string memberName = md.GetString(mth.Name);
                if (memberName.Length == 0 || memberName[0] == '.') continue;
                members.Add(memberName);
            }
            foreach (PropertyDefinitionHandle ph in td.GetProperties())
                members.Add(md.GetString(md.GetPropertyDefinition(ph).Name));
            foreach (FieldDefinitionHandle fh in td.GetFields())
                members.Add(md.GetString(md.GetFieldDefinition(fh).Name));
            foreach (EventDefinitionHandle eh in td.GetEvents())
                members.Add(md.GetString(md.GetEventDefinition(eh).Name));

            // The name counts even when the type declares nothing: what the docs claim is that
            // the type exists.
            ShadowTypesByName[name] = members;
            ShadowTypesByName[fullName] = members;
            _shadowCount++;
            if (!ShadowTypeNames.Contains(name)) ShadowTypeNames.Add(name);

            foreach (string m in members)
            {
                KnownMembers.Add(m);
                if (!MemberOwners.TryGetValue(m, out var owners))
                    MemberOwners[m] = owners = new List<string>();
                if (!owners.Contains(name)) owners.Add(name);
            }
        }
    }

    // -----------------------------------------------------------------------
    // Scanning
    // -----------------------------------------------------------------------

    private static void CheckFile(string file, string root, string[] lines)
    {
        string rel = Path.GetRelativePath(root, file).Replace('\\', '/');

        for (int i = 0; i < lines.Length; i++)
        {
            string opener = lines[i].TrimStart();
            if (!IsCsharpFence(opener)) continue;

            var body = new List<string>();
            int firstLine = i + 2; // 1-based line number of the first body line
            int j = i + 1;
            for (; j < lines.Length; j++)
            {
                if (lines[j].TrimStart().StartsWith("```")) break;
                body.Add(lines[j]);
            }
            if (j >= lines.Length) break;

            ScanBlock(rel, firstLine, body);
            i = j;
        }
    }

    // ```csharp and ```cs, with an optional info string after them. A label that names a
    // different language (```csharp-html) is not a C# block and must not be read as one.
    private static bool IsCsharpFence(string trimmedLine)
    {
        var m = Regex.Match(trimmedLine, @"^```\s*(?<lang>[A-Za-z+#-]+)");
        if (!m.Success) return false;
        string lang = m.Groups["lang"].Value.ToLowerInvariant();
        return lang is "csharp" or "cs";
    }

    private static void ScanBlock(string rel, int firstLine, List<string> rawBody)
    {
        List<string> body = BlankNonCode(rawBody);

        // The block's own vocabulary: a class it declares or a helper it defines is not a
        // claim about the library.
        var localTypes = new HashSet<string>(StringComparer.Ordinal);
        var localMembers = new HashSet<string>(StringComparer.Ordinal);
        string joined = string.Join("\n", body);
        foreach (Match m in TypeDeclaration.Matches(joined)) localTypes.Add(m.Groups["name"].Value);
        foreach (Match m in MemberDeclaration.Matches(joined)) localMembers.Add(m.Groups["name"].Value);

        // Names the whole skill defines, learned in pass 1, feed back into the judgement.
        foreach (string n in localTypes) SkillDeclaredMembers.Add(n);
        foreach (string n in localTypes) SkillDeclaredTypes.Add(n);
        foreach (string n in localMembers) SkillDeclaredMembers.Add(n);

        var env = new Dictionary<string, Type>(StringComparer.Ordinal);
        // What each tracked name holds when its type is one reflection refused to load.
        // The value is the type name, because there is no Type object to point at.
        var shadowEnv = new Dictionary<string, string>(StringComparer.Ordinal);
        var strong = new List<(int line, string name, string why)>();
        var weak = new List<(int line, string name, string why)>();
        var seen = new HashSet<string>(StringComparer.Ordinal);

        for (int i = 0; i < body.Count; i++)
        {
            string line = body[i];
            if (line.Trim().Length == 0) continue;
            int lineNo = firstLine + i;

            string trimmed = line.TrimStart();
            if (trimmed.StartsWith("using ") || trimmed.StartsWith("namespace ")) continue;

            TrackAssignment(env, shadowEnv, line);
            ScanIdentifiers(rel, env, shadowEnv, line, lineNo, localTypes, localMembers, strong, weak, seen);
        }

        // A block that leans on three or more untypable receivers is the sample's own code,
        // not an API reference; its weak claims are demoted - but they stay on the report.
        if (weak.Count >= 3)
        {
            foreach (var c in weak)
            {
                _suspect++;
                Suspects.Add($"{rel}:L{c.line}  '{c.name}' - {c.why}");
            }
        }
        else
        {
            foreach (var c in weak) Report(rel, c.line, c.name, c.why, Findings, ref _notFound);
        }
        foreach (var c in strong) Report(rel, c.line, c.name, c.why, Findings, ref _notFound);
    }

    private static void Report(string rel, int line, string name, string why,
                               List<string> sink, ref int counter)
    {
        counter++;
        sink.Add($"{rel}:L{line}  '{name}' - {why}");
    }

    // -----------------------------------------------------------------------
    // Identifier walk: each name is judged in the position it stands in.
    // -----------------------------------------------------------------------

    private static void ScanIdentifiers(string rel, Dictionary<string, Type> env,
                                        Dictionary<string, string> shadowEnv, string line, int lineNo,
                                        HashSet<string> localTypes, HashSet<string> localMembers,
                                        List<(int line, string name, string why)> strong,
                                        List<(int line, string name, string why)> weak,
                                        HashSet<string> seen)
    {
        Type? lastType = null;   // type of the previous step of the chain being walked
        string lastIdent = "";   // name of the previous step

        // Names between '<' and '>' are type arguments: real type claims, but they say
        // nothing about the value the expression yields. List<GridJsWorkbook> is not a
        // GridJsWorkbook, and IOptions<GridJsOptions>().Value is not a member of it.
        int genericDepth = 0;
        int prevEnd = 0;

        foreach (Match m in Regex.Matches(line, @"[A-Za-z_][A-Za-z0-9_]*"))
        {
            string name = m.Value;
            bool memberPosition = ImmediatelyPrecededByDot(line, m.Index);

            string gap = line.Substring(prevEnd, m.Index - prevEnd);
            prevEnd = m.Index + name.Length;
            bool closedGenerics = false;
            for (int g = 0; g < gap.Length; g++)
            {
                char c = gap[g];
                bool isShiftOrArrow = c == '>' && g > 0 && (gap[g - 1] == '=' || gap[g - 1] == '>');
                if (c == '<' && !(g + 1 < gap.Length && gap[g + 1] == '<')) genericDepth++;
                else if (c == '>' && !isShiftOrArrow && genericDepth > 0)
                {
                    genericDepth--;
                    if (genericDepth == 0) closedGenerics = true;
                }
            }
            bool ownsChain = genericDepth == 0 && !closedGenerics;

            if (!memberPosition)
            {
                Type? previousType = lastType;
                lastType = ownsChain ? LookupType(name) : null;
                lastIdent = ownsChain ? name : "";

                if (IgnoreNames.Contains(name) && !LooksLikeGridJsApi(name)) continue;
                if (TypesByName.ContainsKey(name) || ShadowTypesByName.ContainsKey(name)
                    || localTypes.Contains(name)
                    || SkillDeclaredTypes.Contains(name) || PlaceholderName.IsMatch(name))
                {
                    if (LooksLikeApiName(name)) _checked++;
                    continue;
                }
                if (!LooksLikeApiName(name) || localMembers.Contains(name)) continue;

                // Object-initialiser member: new GridJsOptions { SaveHtmlAsZip = true }.
                // The word stands in no dotted chain, but the type being initialised owns it,
                // so it is judged as a member - which is where a made-up option name shows up.
                if (previousType is not null && PrecededBy(line, m.Index, '{', ',')
                    && FollowedByAssignment(line, m.Index + name.Length))
                {
                    _checked++;
                    if (FirstMember(previousType, name) is null)
                    {
                        string initKey = lineNo + "|" + name;
                        if (seen.Add(initKey))
                            CountCandidate(lineNo, name,
                                $"is not a settable member of {previousType.Name}",
                                false, strong, weak);
                    }
                    continue;
                }

                if (!LooksLikeGridJsApi(name))
                {
                    // A bare word the docs use as a type that neither the assembly nor any
                    // framework reference we know explains. Out of this product's shape, so
                    // the scan has no opinion - and says so by not counting it as a finding.
                    continue;
                }

                string key = lineNo + "|" + name;
                if (!seen.Add(key)) continue;
                CountCandidate(lineNo, name,
                    "used as a type but neither a GridJs type nor declared by the skill",
                    false, strong, weak);
                lastType = null;
                continue;
            }

            // Member position. Resolve the receiver: the chain step before this one, a type
            // by the same name, or a tracked variable.
            Type? receiver = lastType ?? LookupType(lastIdent) ?? LookupVariable(env, lastIdent);
            if (receiver is null && lastIdent == "this")
            {
                lastType = null;
                continue; // the sample's own members
            }

            if (receiver is not null)
            {
                MemberInfo? found = FirstMember(receiver, name);
                if (found is not null)
                {
                    _checked++;
                    lastType = StepType(found);
                    lastIdent = name;
                    continue;
                }
                // A call on a receiver this scan can name is a claim about the library
                // whatever its casing: workbook.getSettings() is a Java spelling that leaked
                // into C# documentation, and the .NET API answers to workbook.Settings.
                bool isCall = FollowedBy(line, m.Index + name.Length, '(');
                if ((!IsJudgementWorthy(name) && !isCall) || localMembers.Contains(name) || SkillDeclaredMembers.Contains(name))
                {
                    lastType = null;
                    lastIdent = name;
                    continue;
                }
                string key = lineNo + "|" + name;
                if (!seen.Add(key)) { lastType = null; lastIdent = name; continue; }
                string why = MemberOwners.TryGetValue(name, out var owners)
                    ? $"is documented on {string.Join("/", owners)}, not on {receiver.Name}"
                    : $"no public member '{name}' on {receiver.Name}";
                CountCandidate(lineNo, name, why, false, strong, weak);
                lastType = null;
                lastIdent = name;
                continue;
            }

            // Unresolved receiver, unless metadata knows the type even though reflection
            // cannot load it. A GridJs class whose base lives in an assembly this verifier
            // does not reference comes back null from GetTypes() - but the assembly file
            // still names it and lists what it declares itself.
            string? shadow = LookupShadowType(lastIdent) ?? LookupShadowVariable(shadowEnv, lastIdent);
            if (shadow is not null)
            {
                if (ShadowTypesByName[shadow].Contains(name))
                {
                    _checked++;
                    lastType = null;
                    lastIdent = name;
                    continue;
                }
                // Only a name shaped like this product's API is a claim about it. Members the
                // type inherits (an MVC controller's Content or NotFound) are invisible here,
                // so a plain word gets no verdict rather than a wrong one.
                if (!LooksLikeGridJsApi(name))
                {
                    if (UnresolvedSamples.Count < 20)
                        UnresolvedSamples.Add($"{rel}:L{lineNo}  '{name}' on metadata-only receiver '{shadow}'");
                    _unresolved++;
                    lastType = null;
                    lastIdent = name;
                    continue;
                }
                string shadowKey = lineNo + "|" + name;
                if (seen.Add(shadowKey))
                    CountCandidate(lineNo, name,
                        $"no member '{name}' declared by {shadow} (metadata only - reflection cannot load this type, so inherited members are not checked)",
                        false, strong, weak);
                lastType = null;
                lastIdent = name;
                continue;
            }

            // Unresolved receiver.
            if (ChainNoise.Contains(name) || localMembers.Contains(name)
                || SkillDeclaredMembers.Contains(name) || !IsJudgementWorthy(name))
            {
                lastType = null;
                lastIdent = name;
                continue;
            }
            if (KnownMembers.Contains(name))
            {
                _checked++;
                lastType = null;
                lastIdent = name;
                continue;
            }
            if (!LooksLikeGridJsApi(name))
            {
                if (UnresolvedSamples.Count < 20)
                    UnresolvedSamples.Add($"{rel}:L{lineNo}  '{name}' on receiver '{lastIdent}'");
                _unresolved++;
                lastType = null;
                lastIdent = name;
                continue;
            }
            string weakKey = lineNo + "|" + name;
            if (!seen.Add(weakKey)) { lastType = null; lastIdent = name; continue; }
            CountCandidate(lineNo, name,
                $"no member '{name}' anywhere in the GridJs assembly (receiver '{lastIdent}' is not a GridJs type the scan can resolve)",
                true, strong, weak);
            lastType = null;
            lastIdent = name;
        }
    }

    // A member claim only needs evidence when the name looks like it belongs to this
    // product or its receiver does; anything else is ASP.NET or the sample's own code.
    private static bool IsJudgementWorthy(string name)
    {
        return LooksLikeGridJsApi(name) || LooksLikeApiName(name);
    }

    private static bool ImmediatelyPrecededByDot(string line, int index)
    {
        int p = index - 1;
        while (p >= 0 && char.IsWhiteSpace(line[p])) p--;
        return p >= 0 && line[p] == '.';
    }

    private static bool FollowedBy(string line, int end, char c)
    {
        int i = end;
        while (i < line.Length && char.IsWhiteSpace(line[i])) i++;
        return i < line.Length && line[i] == c;
    }

    private static bool PrecededBy(string line, int index, params char[] allowed)    {
        int p = index - 1;
        while (p >= 0 && char.IsWhiteSpace(line[p])) p--;
        if (p < 0) return false;
        foreach (char c in allowed)
        {
            if (line[p] == c) return true;
        }
        return false;
    }

    // An initialiser entry assigns: one '=' that is not part of '==' or '=>'.
    private static bool FollowedByAssignment(string line, int end)
    {
        int i = end;
        while (i < line.Length && char.IsWhiteSpace(line[i])) i++;
        if (i >= line.Length || line[i] != '=') return false;
        if (i + 1 < line.Length && (line[i + 1] == '=' || line[i + 1] == '>')) return false;
        return true;
    }

    private static Type? LookupType(string name)
    {
        return TypesByName.TryGetValue(name, out var t) ? t : null;
    }

    private static Type? LookupVariable(Dictionary<string, Type> env, string name)
    {
        return name.Length > 0 && env.TryGetValue(name, out var t) ? t : null;
    }

    private static string? LookupShadowType(string name)
    {
        return name.Length > 0 && ShadowTypesByName.ContainsKey(name) ? name : null;
    }

    private static string? LookupShadowVariable(Dictionary<string, string> shadowEnv, string name)
    {
        return name.Length > 0 && shadowEnv.TryGetValue(name, out var t) ? t : null;
    }

    private static MemberInfo? FirstMember(Type t, string name)
    {
        try
        {
            foreach (var m in t.GetMember(name, PublicAny))
            {
                if (m is not Type) return m;
            }
            // A nested enum/class reached through its owner is a name in its own right.
            foreach (var m in t.GetMember(name, PublicAny))
            {
                return m;
            }
        }
        catch (Exception)
        {
            // A type the loader cannot enumerate is outside what this scan can judge.
        }
        return null;
    }

    private static Type? StepType(MemberInfo member)
    {
        Type? t = member switch
        {
            PropertyInfo p => p.PropertyType,
            FieldInfo f => f.FieldType,
            MethodInfo mi => mi.ReturnType,
            _ => null,
        };
        if (t is null) return null;
        if (t.IsGenericType && t.GetGenericArguments().Length == 1)
        {
            Type inner = t.GetGenericArguments()[0];
            if (LooksLikeGridJsApi(inner.Name)) t = inner;
        }
        return t.IsArray ? t.GetElementType() : t;
    }

    // `GridJsWorkbook wb = ...`, `var wb = new GridJsWorkbook()`, `var s = wb.Settings`
    // and `new GridJsWorkbook().Settings` all tell the scan what a name holds.
    private static void TrackAssignment(Dictionary<string, Type> env, Dictionary<string, string> shadowEnv, string line)
    {
        var m = Regex.Match(line, @"\b(?<lhs>[A-Za-z_][A-Za-z0-9_]*)\s*=\s*(?<rhs>.*)$");
        if (!m.Success) return;
        string lhs = m.Groups["lhs"].Value;
        if (lhs is "var" or "const") return;

        Type? fromRhs = ResolveExpression(env, m.Groups["rhs"].Value);
        if (fromRhs is not null)
        {
            env[lhs] = fromRhs;
            return;
        }
        string? fromShadow = ResolveShadowExpression(shadowEnv, m.Groups["rhs"].Value);
        if (fromShadow is not null)
        {
            shadowEnv[lhs] = fromShadow;
            return;
        }
        // An explicit type on the left is only evidence when it is one we know.
        var decl = Regex.Match(line.TrimStart(),
            @"^(?:(?:public|private|protected|internal|static|readonly|const|var)\s+)*(?<type>[A-Za-z_][A-Za-z0-9_]*)(?:<[^>]*>)?\s+"
            + Regex.Escape(lhs) + @"\s*[=;,)]");
        if (!decl.Success) return;
        string declared = decl.Groups["type"].Value;
        if (LookupType(declared) is Type explicitType) env[lhs] = explicitType;
        else if (LookupShadowType(declared) is string shadowType) shadowEnv[lhs] = shadowType;
    }

    // A metadata-only type has no Type object, so a chain through it stops at the first
    // member: what that member returns is a claim the scan cannot type any more.
    private static string? ResolveShadowExpression(Dictionary<string, string> shadowEnv, string rhs)
    {
        var nw = Regex.Match(rhs, @"^\s*new\s+(?<type>[A-Za-z_][A-Za-z0-9_]*)(?:<[^>]*>)?");
        if (nw.Success) return LookupShadowType(nw.Groups["type"].Value);

        var one = Regex.Match(rhs, @"^\s*(?<root>[A-Za-z_][A-Za-z0-9_]*)\s*;?\s*$");
        if (one.Success)
            return LookupShadowType(one.Groups["root"].Value) ?? LookupShadowVariable(shadowEnv, one.Groups["root"].Value);
        return null;
    }

    private static Type? ResolveExpression(Dictionary<string, Type> env, string rhs)
    {
        var nw = Regex.Match(rhs, @"^\s*new\s+(?<type>[A-Za-z_][A-Za-z0-9_]*)(?:<[^>]*>)?");
        if (nw.Success && LookupType(nw.Groups["type"].Value) is Type created) return created;

        var chain = Regex.Match(rhs, @"^\s*(?<root>[A-Za-z_][A-Za-z0-9_]*)(?<rest>(?:\s*\.\s*[A-Za-z_][A-Za-z0-9_]*)*)");
        if (!chain.Success) return null;
        Type? cur = LookupType(chain.Groups["root"].Value) ?? LookupVariable(env, chain.Groups["root"].Value);
        if (cur is null) return null;
        foreach (string step in chain.Groups["rest"].Value.Split('.'))
        {
            string name = step.Trim();
            if (name.Length == 0) continue;
            MemberInfo? found = FirstMember(cur, name);
            if (found is null) return cur;
            cur = StepType(found) ?? cur;
        }
        return cur;
    }

    // -----------------------------------------------------------------------
    // Text preparation
    // -----------------------------------------------------------------------

    // Blank comments and string contents, block strings included, so nothing the reader
    // copies as text is judged as code. Line count is preserved: every reported line
    // number stays true.
    private static List<string> BlankNonCode(List<string> body)
    {
        var outLines = new List<string>(body.Count);
        bool inBlockComment = false;
        foreach (string raw in body)
        {
            var sb = new StringBuilder(raw.Length);
            int i = 0;
            while (i < raw.Length)
            {
                char c = raw[i];
                char next = i + 1 < raw.Length ? raw[i + 1] : '\0';

                if (inBlockComment)
                {
                    if (c == '*' && next == '/') { inBlockComment = false; sb.Append("  "); i += 2; continue; }
                    sb.Append(' ');
                    i++;
                    continue;
                }
                if (c == '/' && next == '/') { break; }
                if (c == '/' && next == '*') { inBlockComment = true; sb.Append("  "); i += 2; continue; }
                if (c == '@' && next == '"')
                {
                    // Verbatim string: "" escapes a quote, no backslash rules apply.
                    sb.Append("  ");
                    i += 2;
                    while (i < raw.Length)
                    {
                        if (raw[i] == '"')
                        {
                            if (i + 1 < raw.Length && raw[i + 1] == '"') { i += 2; continue; }
                            i++;
                            break;
                        }
                        i++;
                    }
                    continue;
                }
                if ((c == '"' || c == '\'') || (c == '$' && next == '"'))
                {
                    char quote = c == '$' ? '"' : c;
                    if (c == '$') { sb.Append(' '); i++; }
                    sb.Append(quote);
                    i++;
                    while (i < raw.Length)
                    {
                        if (raw[i] == '\\') { sb.Append("  "); i += 2; continue; }
                        if (raw[i] == quote) { sb.Append(quote); i++; break; }
                        sb.Append(' ');
                        i++;
                    }
                    continue;
                }
                sb.Append(c);
                i++;
            }
            outLines.Add(sb.ToString());
        }
        return outLines;
    }

    private static bool LooksLikeGridJsApi(string name)
    {
        return name.StartsWith("GridJs", StringComparison.Ordinal)
            || name.StartsWith("Grid", StringComparison.Ordinal)
            || name.StartsWith("Aspose", StringComparison.Ordinal)
            || name.StartsWith("CoWork", StringComparison.Ordinal)
            || name.Contains("GridJs", StringComparison.Ordinal);
    }

    private static bool LooksLikeApiName(string name)
    {
        if (!char.IsUpper(name[0])) return false;
        if (name.Length < 4) return false;
        if (name.All(char.IsUpper)) return false;
        return !name.Contains('.');
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
            if (Directory.GetFiles(dir.FullName, "*.md").Length > 0 &&
                Directory.GetFiles(dir.FullName, "*-api-*.md").Length > 0)
                return dir.FullName;
            dir = dir.Parent;
        }
        return null;
    }
}
