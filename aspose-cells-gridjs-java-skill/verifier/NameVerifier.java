// Name-level verifier for a GridJs Java skill.
//
// Oracle: the java API doc in api-doc/ next to this source. Candidates come from Java
// syntax (imports, declarations, tracked receivers, super calls), not from a .NET
// assembly, so documented Java names are never matched against C# members.
//
// Usage: java NameVerifier.java [skillRoot]
// Exit code: 0 = clean, 1 = at least one unknown name, 2 = the verifier could not run.
//
// A finding is a name the skill claims about this library that the API doc does not carry,
// including one it hangs off the wrong type. A suspect is a name the scan cannot attach to
// the library at all: counted and printed, never dropped, never failed.

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Comparator;
import java.util.HashMap;
import java.util.HashSet;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Stream;

public final class NameVerifier {

    private static final Set<String> TYPES = new HashSet<>();
    private static final Set<String> ALL_MEMBERS = new HashSet<>();
    private static final Map<String, Set<String>> MEMBERS_BY_TYPE = new HashMap<>();
    // Which type documents each member, in the order the oracle lists them. Once the
    // receiver's type is known, "this name exists somewhere" stops being evidence: a real
    // Config setter hung off a workbook is the most common shape an LLM hallucination takes.
    private static final Map<String, Set<String>> MEMBER_OWNERS = new HashMap<>();

    // Platform types (java.*, javax/jakarta, spring, Aspose.Cells) that the jar's own
    // signatures mention. They are evidence, not invention, so they must never be
    // reported as unknown - but they are not GridJs types either, so they stay out of TYPES.
    private static final Set<String> SIGNATURE_TYPES = new HashSet<>();

    // JDK, Spring and Servlet names that legitimately appear in the docs.
    private static final Set<String> IGNORE_TYPES = new HashSet<>(Arrays.asList(
        "String", "Integer", "Long", "Double", "Float", "Boolean", "Byte", "Short",
        "Object", "Class", "Override", "Deprecated", "Exception", "RuntimeException",
        "IllegalArgumentException", "IllegalStateException", "IOException",
        "FileNotFoundException", "InterruptedException", "StringBuilder",
        "StringBuffer", "Math", "System", "Thread", "Runnable", "Optional", "Objects",
        "List", "ArrayList", "LinkedList", "Map", "HashMap", "LinkedHashMap",
        "TreeMap", "Set", "HashSet", "LinkedHashSet", "Collection", "Collections",
        "Arrays", "Iterator", "Stream", "Collectors", "Function", "Consumer",
        "Supplier", "Predicate", "UUID", "Date", "LocalDate", "LocalDateTime",
        "Duration", "File", "Files", "Path", "Paths",
        "InputStream", "OutputStream", "BufferedReader", "FileReader", "FileWriter",
        "ByteArrayOutputStream", "StandardCharsets", "Charsets",
        "RestController", "Controller", "RequestMapping", "GetMapping", "PostMapping",
        "PutMapping", "DeleteMapping", "RequestParam", "RequestBody", "PathVariable",
        "RequestHeader", "RequestPart", "ModelAttribute", "CrossOrigin", "Autowired",
        // The exception-handling vocabulary a troubleshooting guide writes in: a global
        // advice class and its handlers are Spring types, never GridJs types.
        "ControllerAdvice", "RestControllerAdvice", "ExceptionHandler", "ResponseStatus",
        "ResponseStatusException", "Validated", "MethodArgumentNotValidException",
        "BindException", "HttpMediaTypeNotSupportedException",
        "MissingServletRequestParameterException", "AsyncRequestNotUsableException",
        "Bean", "Configuration", "Component", "Service", "Repository",
        "SpringBootApplication", "SpringApplication", "ComponentScan",
        "ConfigurationProperties", "EnableConfigurationProperties",
        "Conditional", "Profile", "Scope", "Order", "Primary", "Qualifier", "Value",
        "Resource", "InitBean", "ApplicationRunner", "ApplicationArguments",
        "HttpServletRequest", "HttpServletResponse", "HttpSession", "HttpHeaders",
        "HttpMethod", "HttpStatus", "MediaType", "ResponseEntity",
        "MultipartFile", "MultipartException", "InputStreamResource", "FileSystemResource",
        "ClassPathResource",
        "ResourceLoader", "WebMvcConfigurer", "CorsRegistry", "MappingJackson2HttpMessageConverter",
        "JsonSerializer", "JsonObject", "JSONArray", "Gson", "ObjectMapper", "JsonNode",
        "Logger", "LoggerFactory", "Slf4j",
        // Project scaffolding the docs define themselves.
        "GridJsController", "GridJsConfig", "GridJsServiceConfig", "DemoApplication",
        "SpreadsheetController", "WebConfig", "AppConfig", "Application", "Main",
        "Hello", "World", "Example", "Demo", "Sample", "Test", "Tests",
        "Excel", "JSON", "REST", "API", "UID", "URL", "URI", "ID"
    ));

    // Methods inherited from Object or so generic that matching them against the
    // GridJs doc would only produce noise.
    private static final Set<String> IGNORE_MEMBERS = new HashSet<>(Arrays.asList(
        "toString", "hashCode", "equals", "getClass", "notify", "notifyAll", "wait",
        "clone", "name", "ordinal", "values", "valueOf", "stream", "forEach",
        "get", "set", "put", "add", "remove", "clear", "size", "isEmpty", "contains",
        "containsKey", "keySet", "entrySet", "values", "iterator", "build", "of",
        "newInstance", "close", "flush", "write", "read", "length", "append",
        "startsWith", "endsWith", "substring", "replace", "split", "trim", "format",
        "getBytes", "newString", "equalsIgnoreCase", "matches", "chars",
        "exists", "create", "delete", "mkdirs", "resolve", "getParent", "getFileName",
        "toAbsolutePath", "absolutePathString", "getPath", "getName", "getCanonicalPath",
        "getInputStream", "getBody", "getStatusCode", "getStatusCodeValue", "getHeaders",
        "ok", "status", "body", "header", "contentType", "accept", "exchange",
        "getDeclaredMethod", "invoke", "getMethod", "isPresent", "orElse", "orElseGet",
        "ifPresent", "map", "filter", "collect", "sorted", "distinct", "limit", "skip",
        "log", "info", "warn", "error", "debug", "trace", "println", "print",
        // Aspose.Cells surface wrapped by GridJs but documented separately.
        "getWorksheets", "worksheets", "getCells", "cells", "getCell", "getRows",
        "getColumns", "getRange", "importFile", "save", "setValue", "setStyle",
        "getStyle", "getName", "setName", "setVisible", "getSheetName"
    ));

    // A doc that pastes an owner prefix onto a signature the reference prints without one
    // ("Config.static void setFileCacheDirectory(...)", "GridJsService.java.lang.StringBuilder
    // detailFileJsonWithUid(...)") leaves a modifier, a primitive or a package root where a
    // member name would sit. Java forbids all of these as member names, so the qualified scan
    // must not read them as an API claim.
    private static final Set<String> NOT_MEMBER_TOKENS = new HashSet<>(Arrays.asList(
        "static", "final", "public", "private", "protected", "abstract", "synchronized",
        "transient", "volatile", "native", "default", "return", "throws", "new",
        "void", "int", "long", "short", "byte", "char", "boolean", "float", "double",
        "java", "javax", "com", "org"
    ));

    private static final Pattern MEMBER_ATTR =
        Pattern.compile("<member name=\"([TFMPE]):([^\"]*)\"");
    private static final Pattern SIGNATURE_TYPE_ATTR =
        Pattern.compile("\\btype=\"([^\"]+)\"");
    private static final Pattern SIGNATURE_TYPE_ELEMENT =
        Pattern.compile("<type>([^<]+)</type>");
    // Both fences must start a line. A table cell can hold an escaped example
    // ("| ```java\nConfig.setX(1);\n``` |"), whose literal backslash-n is not a newline:
    // without the anchors the capture runs to the end of that row, takes the newline, and
    // swallows the next row's prose up to its own inline fence - every heading and property
    // name in it then reads as a code identifier.
    private static final Pattern JAVA_BLOCK =
        Pattern.compile("^[ \\t]*```java(?!\\w)[^\\n]*\\n(.*?)^[ \\t]*```",
            Pattern.DOTALL | Pattern.MULTILINE);
    private static final Pattern IMPORT =
        Pattern.compile("\\bimport\\s+(static\\s+)?com\\.aspose\\.gridjs\\.([A-Za-z0-9_.]+)\\s*;");
    private static final Pattern EXTENDS =
        Pattern.compile("\\bclass\\s+[A-Za-z0-9_]+\\s+extends\\s+([A-Z][A-Za-z0-9_]*)");
    private static final Pattern DECLARATION =
        Pattern.compile("\\b([A-Z][A-Za-z0-9_]*)\\s+([a-z][A-Za-z0-9_]*)\\s*[=;,)]");
    private static final Pattern NEW_CALL = Pattern.compile("\\bnew\\s+([A-Z][A-Za-z0-9_]*)\\s*\\(");
    private static final Pattern QUALIFIED =
        Pattern.compile("\\b([A-Za-z_][A-Za-z0-9_]*)\\.([A-Za-z_][A-Za-z0-9_]*)\\b");
    private static final Pattern ANY_PASCAL = Pattern.compile("\\b([A-Z][A-Za-z0-9_]{2,})\\b");
    private static final Pattern LOCAL_TYPE =
        Pattern.compile("\\b(?:class|interface|enum)\\s+([A-Z][A-Za-z0-9_]*)");
    private static final Pattern PLACEHOLDER_TYPE =
        Pattern.compile("(?:My|Custom|Demo|Sample|Test|Example|Simple|Own)[A-Z][A-Za-z0-9_]*");
    // Troubleshooting prose uses obvious stand-in identifiers - "yourService.someMethod()"
    // - to describe a pattern rather than a library member, so the member scan skips them.
    private static final Pattern PLACEHOLDER_MEMBER = Pattern.compile(
        "(?:some|any|our|my|your|their|dummy|fake|placeholder)_[A-Za-z0-9]+"
        + "|(?:some|any|our|my|your|their|dummy|fake)[A-Z][A-Za-z0-9]*"
        + "|foo|bar|baz|qux|x{3,}"
        + "|do[_]?[Ss]omething|something|whatever|anything");

    // The short names of packages a documentation snippet can be assumed to import. Used
    // to recognise a JDK type by resolution instead of by hand, because a list written by
    // hand always lags behind the prose (PrintWriter was reported as invented).
    private static final String[] JDK_PACKAGES = {
        "java.lang.", "java.lang.reflect.", "java.util.", "java.util.concurrent.",
        "java.util.stream.", "java.util.function.", "java.time.", "java.time.format.",
        "java.io.", "java.nio.", "java.nio.file.", "java.nio.charset.", "java.net.",
        "java.text.", "java.math.", "java.sql.", "javax.crypto.",
    };
    private static final Map<String, Boolean> JDK_TYPES = new HashMap<>();

    private NameVerifier() {
    }

    public static void main(String[] args) {
        int code;
        try {
            code = run(args);
        } catch (Throwable t) {
            // A crash says the verifier is broken, not that the skill has bad names.
            System.err.println("Verifier could not run: " + t.getClass().getName() + ": " + t.getMessage());
            code = 2;
        }
        System.exit(code);
    }

    private static int run(String[] args) throws IOException {
        Path base = Paths.get(System.getProperty("user.dir")).toAbsolutePath();
        Path root = args.length > 0 ? Paths.get(args[0]).toAbsolutePath() : findSkillRoot(base);
        if (root == null || !Files.isDirectory(root)) {
            System.err.println("Could not locate skill root.");
            return 2;
        }

        Path apiDoc = findApiDoc(base, root);
        if (apiDoc == null) {
            System.err.println("No API doc found under verifier/api-doc - cannot verify names.");
            return 2;
        }

        loadOracle(apiDoc);

        System.out.println("GridJs Java API index: " + TYPES.size() + " types, "
            + ALL_MEMBERS.size() + " members  (from " + apiDoc.getFileName() + ")");
        System.out.println("Scanning: " + root);
        System.out.println(new String(new char[60]).replace('\0', '-'));

        int[] stats = new int[3];
        List<String> findings = new ArrayList<>();
        List<String> suspects = new ArrayList<>();

        for (Path file : markdownFiles(root)) {
            String text = new String(Files.readAllBytes(file), StandardCharsets.UTF_8);
            String rel = root.relativize(file).toString().replace('\\', '/');
            Matcher m = JAVA_BLOCK.matcher(text);
            while (m.find()) {
                checkBlock(countLines(text.substring(0, m.start())) + 2, m.group(1), rel,
                    stats, findings, suspects);
            }
        }

        System.out.println(new String(new char[60]).replace('\0', '-'));
        System.out.println("CHECKED: " + stats[0] + "   NOT FOUND: " + stats[1]
            + "   SUSPECT: " + stats[2]);
        // Machine-readable result line. The caller trusts this rather than the exit code,
        // because a JVM that dies on the Windows paging file also leaves exit code 1 behind,
        // and that is a broken toolchain, not a statement about the documentation.
        System.out.println("VERIFIER-RESULT: checked=" + stats[0] + " not-found=" + stats[1]
            + " suspect=" + stats[2]);

        if (!findings.isEmpty()) {
            System.out.println(new String(new char[60]).replace('\0', '-'));
            System.out.println("Undocumented Java API names (" + findings.size() + "):");
            for (String f : findings) {
                System.out.println("  " + f);
            }
        }

        if (!suspects.isEmpty()) {
            System.out.println(new String(new char[60]).replace('\0', '-'));
            System.out.println("Demoted from blocks of placeholder code, review by hand ("
                + suspects.size() + "):");
            for (String s : suspects) {
                System.out.println("  " + s);
            }
        }

        return stats[1] == 0 ? 0 : 1;
    }

    private static int countLines(String s) {
        int n = 0;
        for (int i = 0; i < s.length(); i++) {
            if (s.charAt(i) == '\n') {
                n++;
            }
        }
        return n;
    }

    private static Path findApiDoc(Path base, Path root) {
        Path[] candidates = {
            base.resolve("api-doc"),
            root.resolve("verifier").resolve("api-doc"),
        };
        for (Path dir : candidates) {
            if (!Files.isDirectory(dir)) {
                continue;
            }
            try (Stream<Path> s = Files.list(dir)) {
                Path doc = s.filter(p -> p.getFileName().toString().endsWith(".xml"))
                    .min(Comparator.comparing(p -> p.getFileName().toString()))
                    .orElse(null);
                if (doc != null) {
                    return doc;
                }
            } catch (IOException e) {
                return null;
            }
        }
        return null;
    }

    private static Path findSkillRoot(Path start) {
        Path dir = start;
        while (dir != null) {
            if (Files.exists(dir.resolve("SKILL.md")) || Files.exists(dir.resolve("skill-index.md"))) {
                return dir;
            }
            dir = dir.getParent();
        }
        return null;
    }

    private static List<Path> markdownFiles(Path root) throws IOException {
        List<Path> files = new ArrayList<>();
        try (Stream<Path> s = Files.walk(root)) {
            for (Path p : (Iterable<Path>) s::iterator) {
                String rel = root.relativize(p).toString().replace('\\', '/');
                if (!Files.isRegularFile(p) || !rel.endsWith(".md")) {
                    continue;
                }
                String top = rel.contains("/") ? rel.substring(0, rel.indexOf('/')) : rel;
                if (top.startsWith("example") || top.equals("verifier") || top.equals("report")) {
                    continue;
                }
                files.add(p);
            }
        }
        files.sort(Comparator.comparing(Path::toString));
        return files;
    }

    private static void loadOracle(Path apiDoc) throws IOException {
        String text = new String(Files.readAllBytes(apiDoc), StandardCharsets.UTF_8);
        if (!text.isEmpty() && text.charAt(0) == '\uFEFF') {
            text = text.substring(1);
        }
        Matcher m = MEMBER_ATTR.matcher(text);
        while (m.find()) {
            String kind = m.group(1);
            String qualified = m.group(2);
            String head = qualified.contains("(") ? qualified.substring(0, qualified.indexOf('(')) : qualified;
            String[] parts = head.split("\\.");
            String last = parts[parts.length - 1];

            if (kind.equals("T")) {
                TYPES.add(last);
                continue;
            }

            String owner = parts.length > 2 ? parts[parts.length - 2] : "";
            if (!owner.isEmpty()) {
                TYPES.add(owner);
                Set<String> owned = MEMBERS_BY_TYPE.computeIfAbsent(owner, k -> new HashSet<>());
                owned.add(last);
                ALL_MEMBERS.add(last);
                MEMBER_OWNERS.computeIfAbsent(last, k -> new LinkedHashSet<>()).add(owner);

                // The doc records fields and properties .NET-style; Java code reaches
                // them through accessors, so accept the JavaBean spellings too.
                if (kind.equals("P") || kind.equals("F")) {
                    String accessor = last.substring(0, 1).toUpperCase() + last.substring(1);
                    for (String bean : Arrays.asList("set" + accessor, "get" + accessor, "is" + accessor)) {
                        owned.add(bean);
                        ALL_MEMBERS.add(bean);
                        MEMBER_OWNERS.computeIfAbsent(bean, k -> new LinkedHashSet<>()).add(owner);
                    }
                }
            }

            for (String params : matchAll(qualified)) {
                for (String decl : params.split(",")) {
                    String token = decl.trim().replaceAll("<.*>", "");
                    if (token.isEmpty()) {
                        continue;
                    }
                    token = token.contains(" ") ? token.substring(0, token.indexOf(' ')).trim() : token;
                    token = token.contains(" ") ? token.substring(token.lastIndexOf(' ') + 1) : token;
                    token = token.contains(".") ? token.substring(token.lastIndexOf('.') + 1) : token;
                    if (token.length() > 2 && Character.isUpperCase(token.charAt(0))) {
                        TYPES.add(token);
                    }
                }
            }
        }

        harvestSignatureTypes(text);
    }

    // Collect the platform types quoted by the oracle's own signatures so the scan
    // below can tell "cited from the jar" apart from "invented by the model".
    private static void harvestSignatureTypes(String text) {
        List<String> rawTypes = new ArrayList<>();
        Matcher attr = SIGNATURE_TYPE_ATTR.matcher(text);
        while (attr.find()) {
            rawTypes.add(attr.group(1));
        }
        Matcher element = SIGNATURE_TYPE_ELEMENT.matcher(text);
        while (element.find()) {
            rawTypes.add(element.group(1));
        }
        for (String raw : rawTypes) {
            String token = raw.replaceAll("<.*>", "").replace("[]", "").trim();
            if (token.contains(".")) {
                token = token.substring(token.lastIndexOf('.') + 1);
            }
            if (token.length() > 2 && Character.isUpperCase(token.charAt(0))) {
                SIGNATURE_TYPES.add(token);
            }
        }
    }
    // Parameter list of a member signature, if it has one.
    private static List<String> matchAll(String qualified) {
        List<String> out = new ArrayList<>();
        int open = qualified.indexOf('(');
        int close = qualified.lastIndexOf(')');
        if (open >= 0 && close > open) {
            out.add(qualified.substring(open + 1, close));
        }
        return out;
    }

    private static boolean looksLikeGridJsType(String name) {
        return name.startsWith("GridJs") || name.startsWith("Grid")
            || name.startsWith("Aspose") || name.startsWith("IGridJs")
            || name.contains("GridJs");
    }

    private static boolean isJdkType(String name) {
        Boolean cached = JDK_TYPES.get(name);
        if (cached != null) {
            return cached;
        }
        boolean found = false;
        for (String pkg : JDK_PACKAGES) {
            try {
                Class.forName(pkg + name, false, ClassLoader.getSystemClassLoader());
                found = true;
                break;
            } catch (Throwable ignored) {
                // Not in this package, so try the next one. Anything the JDK cannot
                // resolve stays a candidate for the unknown-name report.
            }
        }
        JDK_TYPES.put(name, found);
        return found;
    }

    private static boolean isPlaceholderMember(String name) {
        return PLACEHOLDER_MEMBER.matcher(name).matches();
    }

    // Blank out comments and string/char contents before any name is read, keeping the
    // character count so line numbers still line up. Without this pass the prose inside a
    // javadoc block and the text inside a URL are scanned as if they were code, so a word
    // like "Load" is reported as an undocumented API name.
    private static String sanitize(String block) {
        StringBuilder out = new StringBuilder(block.length());
        boolean inString = false;
        boolean inChar = false;
        boolean inLine = false;
        boolean inBlock = false;

        for (int i = 0; i < block.length(); i++) {
            char c = block.charAt(i);
            char next = i + 1 < block.length() ? block.charAt(i + 1) : '\0';

            if (inLine) {
                if (c == '\n') {
                    inLine = false;
                    out.append('\n');
                } else {
                    out.append(' ');
                }
                continue;
            }
            if (inBlock) {
                if (c == '*' && next == '/') {
                    inBlock = false;
                    out.append("  ");
                    i++;
                } else {
                    out.append(c == '\n' ? '\n' : ' ');
                }
                continue;
            }
            if (inString) {
                if (c == '\\') {
                    out.append("  ");
                    i++;
                } else if (c == '"') {
                    inString = false;
                    out.append('"');
                } else {
                    out.append(c == '\n' ? '\n' : ' ');
                }
                continue;
            }
            if (inChar) {
                if (c == '\\') {
                    out.append("  ");
                    i++;
                } else if (c == '\'') {
                    inChar = false;
                    out.append('\'');
                } else {
                    out.append(' ');
                }
                continue;
            }
            if (c == '/' && next == '/') {
                inLine = true;
                out.append("  ");
                i++;
                continue;
            }
            if (c == '/' && next == '*') {
                inBlock = true;
                out.append("  ");
                i++;
                continue;
            }
            if (c == '"') {
                inString = true;
                out.append('"');
                continue;
            }
            if (c == '\'') {
                inChar = true;
                out.append('\'');
                continue;
            }
            out.append(c);
        }
        return out.toString();
    }

    private static void checkBlock(int startLine, String rawBlock, String rel, int[] stats,
                                   List<String> findings, List<String> suspects) {
        String block = sanitize(rawBlock);

        // Names the snippet declares itself (MyCache, CustomProvider) are example code, not
        // claims about the library, so they are exempt from every check below.
        Set<String> declaredLocally = new HashSet<>();
        Matcher local = LOCAL_TYPE.matcher(block);
        while (local.find()) {
            declaredLocally.add(local.group(1));
        }

        // Structural candidates are attached to something only GridJs can be: an
        // aspose.gridjs import, a member of a tracked receiver, a super call.
        List<String[]> structural = new ArrayList<>();
        List<String[]> loose = new ArrayList<>();
        Map<String, String> varTypes = new HashMap<>();
        String superType = null;

        String[] lines = block.split("\n", -1);
        for (int i = 0; i < lines.length; i++) {
            String line = lines[i];
            if (line.trim().isEmpty()) {
                continue;
            }
            int lineNo = startLine + i;

            Matcher ext = EXTENDS.matcher(line);
            while (ext.find()) {
                if (TYPES.contains(ext.group(1))) {
                    superType = ext.group(1);
                }
            }

            Matcher imp = IMPORT.matcher(line);
            while (imp.find()) {
                String path = imp.group(2);
                if (path.endsWith("*")) {
                    continue;
                }
                String[] parts = path.split("\\.");
                String last = parts[parts.length - 1];
                boolean isStatic = imp.group(1) != null;
                if (isStatic && parts.length > 1) {
                    String owner = parts[parts.length - 2];
                    if (TYPES.contains(owner)
                        && (ALL_MEMBERS.contains(last) || MEMBERS_BY_TYPE.getOrDefault(owner, Set.of()).contains(last))) {
                        stats[0]++;
                    } else if (IGNORE_MEMBERS.contains(last)) {
                        stats[0]++;
                    } else {
                        structural.add(new String[] { String.valueOf(lineNo), last,
                            "imported from com.aspose.gridjs." + owner + " but undocumented" });
                    }
                } else if (TYPES.contains(last) || IGNORE_TYPES.contains(last)) {
                    stats[0]++;
                } else if (looksLikeGridJsType(last)) {
                    structural.add(new String[] { String.valueOf(lineNo), last,
                        "imported from com.aspose.gridjs but undocumented" });
                }
            }

            if (line.trim().startsWith("import ")) {
                continue;
            }

            Matcher decl = DECLARATION.matcher(line);
            while (decl.find()) {
                String type = decl.group(1);
                String var = decl.group(2);
                if (TYPES.contains(type)) {
                    varTypes.put(var, type);
                    stats[0]++;
                } else if (SIGNATURE_TYPES.contains(type) || IGNORE_TYPES.contains(type)) {
                    stats[0]++;
                } else if (looksLikeGridJsType(type)) {
                    structural.add(new String[] { String.valueOf(lineNo), type,
                        "not a documented GridJs type" });
                }
            }

            Matcher nw = NEW_CALL.matcher(line);
            while (nw.find()) {
                String type = nw.group(1);
                if (TYPES.contains(type) || IGNORE_TYPES.contains(type)) {
                    stats[0]++;
                } else if (looksLikeGridJsType(type)) {
                    structural.add(new String[] { String.valueOf(lineNo), type,
                        "not a documented GridJs type" });
                }
            }

            Matcher qual = QUALIFIED.matcher(line);
            while (qual.find()) {
                String receiver = qual.group(1);
                String member = qual.group(2);
                String owner = null;
                if (TYPES.contains(receiver)) {
                    owner = receiver;
                } else if (varTypes.containsKey(receiver)) {
                    owner = varTypes.get(receiver);
                } else if (receiver.equals("super") && superType != null) {
                    owner = superType;
                }
                if (owner == null) {
                    continue;
                }
                if (NOT_MEMBER_TOKENS.contains(member)) {
                    stats[0]++;
                    continue;
                }
                if (!MEMBERS_BY_TYPE.containsKey(owner)) {
                    // The name is only a type the jar's own signatures quote
                    // (HttpServletResponse, Workbook, Stream): the oracle lists no members
                    // for it, so rejecting its members would be guesswork.
                    stats[0]++;
                    continue;
                }
                Set<String> owned = MEMBERS_BY_TYPE.getOrDefault(owner, new HashSet<>());
                if (owned.contains(member) || IGNORE_MEMBERS.contains(member) || TYPES.contains(member)
                    || isPlaceholderMember(member)) {
                    stats[0]++;
                } else if (ALL_MEMBERS.contains(member)) {
                    // The oracle does document this member - just not on the type the skill
                    // puts it on. Say which type, so the fix is not a guess.
                    structural.add(new String[] { String.valueOf(lineNo), member,
                        "is documented on " + String.join("/",
                            MEMBER_OWNERS.getOrDefault(member, new LinkedHashSet<>()))
                        + ", not on " + owner });
                } else {
                    structural.add(new String[] { String.valueOf(lineNo), member,
                        "not a documented member of " + owner });
                }
            }

            Matcher pascal = ANY_PASCAL.matcher(line);
            while (pascal.find()) {
                String name = pascal.group(1);
                // A member reference (HttpStatus.INTERNAL_SERVER_ERROR) is already judged by
                // the qualified scan above, and a type the snippet declares itself
                // (class MyCache extends GridCacheForStream) or names by the placeholder
                // convention (new MyUserProvider()) is example code, not an API claim.
                boolean memberAccess = pascal.start() > 0 && line.charAt(pascal.start() - 1) == '.';
                // "@PostConstruct" and its siblings are annotations from the surrounding
                // framework, never a claim about this product; one that does look like a
                // GridJs name still has to resolve.
                boolean annotationUse = pascal.start() > 0 && line.charAt(pascal.start() - 1) == '@'
                    && !looksLikeGridJsType(name);
                if (memberAccess || annotationUse || declaredLocally.contains(name) || PLACEHOLDER_TYPE.matcher(name).matches()) {
                    stats[0]++;
                } else if (TYPES.contains(name) || ALL_MEMBERS.contains(name) || IGNORE_TYPES.contains(name)
                    || SIGNATURE_TYPES.contains(name) || isJdkType(name)) {
                    stats[0]++;
                } else {
                    loose.add(new String[] { String.valueOf(lineNo), name,
                        "not found in the GridJs Java API" });
                }
            }
        }

        Set<String> structuralKeys = new HashSet<>();
        for (String[] c : structural) {
            structuralKeys.add(c[0] + "|" + c[1]);
        }

        // Several detached unknowns mean example or placeholder code, not an API
        // reference; the loose batch is demoted there rather than dropped, because a block
        // that vanishes from the report cannot be reviewed.
        if (loose.size() >= 3) {
            Set<String> seen = new HashSet<>();
            for (String[] c : loose) {
                String key = c[0] + "|" + c[1];
                // A name one of the structural scans already reported is a finding, not a
                // suspect, and the same name reached twice by this scan is one line of review.
                if (structuralKeys.contains(key) || !seen.add(key)) {
                    continue;
                }
                stats[2]++;
                suspects.add(rel + ":L" + c[0] + "  '" + c[1] + "' - " + c[2]);
            }
            loose = new ArrayList<>();
        }

        Set<String> reported = new HashSet<>();
        for (String[] c : structural) {
            // One name can arrive through two scans (a declaration and its new call); the
            // skill has one defect to fix, so report it once.
            if (!reported.add(c[0] + "|" + c[1])) {
                continue;
            }
            report(c, rel, stats, findings);
        }
        for (String[] c : loose) {
            if (reported.contains(c[0] + "|" + c[1])) {
                continue;
            }
            reported.add(c[0] + "|" + c[1]);
            report(c, rel, stats, findings);
        }
    }

    private static void report(String[] candidate, String rel, int[] stats, List<String> findings) {
        stats[0]++;
        stats[1]++;
        findings.add(rel + ":L" + candidate[0] + "  '" + candidate[1] + "' - " + candidate[2]);
    }
}

