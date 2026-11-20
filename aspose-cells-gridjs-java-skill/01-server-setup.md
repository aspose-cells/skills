# Server-Side Setup Guide

## Prerequisites

- Java 17+ (for Java)
- Spring Boot 2.7.x or 3.x
- Maven dependencies: aspose-cells, aspose-cells (classifier gridjs), gson

## Step 1: Maven Configuration

```xml
<!-- pom.xml -->
<dependencies>
    <dependency>
        <groupId>org.springframework.boot</groupId>
        <artifactId>spring-boot-starter-web</artifactId>
    </dependency>
    <!-- Aspose.Cells core library -->
    <dependency>
        <groupId>com.aspose</groupId>
        <artifactId>aspose-cells</artifactId>
        <version>26.10</version>
    </dependency>
    <!-- Aspose.Cells GridJs extension -->
    <dependency>
        <groupId>com.aspose</groupId>
        <artifactId>aspose-cells</artifactId>
        <classifier>gridjs</classifier>
        <version>26.10</version>
    </dependency>
    <!-- Gson is a runtime requirement of the gridjs extension: without it
         every request fails with NoClassDefFoundError: com/google/gson/JsonElement -->
    <dependency>
        <groupId>com.google.code.gson</groupId>
        <artifactId>gson</artifactId>
        <version>2.10.1</version>
    </dependency>
</dependencies>

<repositories>
    <repository>
        <id>AsposeJavaAPI</id>
        <name>Aspose Java API</name>
        <url>https://releases.aspose.com/java/repo/</url>
    </repository>
</repositories>
```

## Step 2: Register the GridJs packages with Spring

`GridJsService` is already a Spring component inside the jar, so the application
class only has to add the library package to the component scan. Declaring your
own `@Bean` for it fails startup with `@Bean definition illegally overridden by
existing bean definition`.

```java
// DemoApplication.java
@SpringBootApplication
@ComponentScan(basePackages = { "com.example.gridjsdemo", "com.aspose.gridjs" })
public class DemoApplication {
    public static void main(String[] args) throws Exception {
        SpringApplication.run(DemoApplication.class, args);
    }
}
```

## Step 3: Configure GridJsOptions

The options bean is the one thing the application owns. `setFileCacheDirectory`
receives a directory that must already exist - `GridJsService` throws
`java.io.FileNotFoundException` on startup when it does not. That directory is
where GridJs writes the per-`uid` session copies; the workbook the demo loads is
application content and lives somewhere else (see Step 4).

```java
// GridJsConfig.java
@Configuration
public class GridJsConfig {
    @Bean
    public GridJsOptions gridJsOptions() {
        GridJsOptions options = new GridJsOptions();
        options.setFileCacheDirectory("./cache");  // must exist before startup
        options.setBaseRouteName("/GridJs");
        return options;
    }
}
```

## Step 4: Create Controller

```java
// GridJsController.java
@RestController
@RequestMapping("/GridJs")
public class GridJsController extends GridJsControllerBase {
    @Autowired
    public GridJsController(GridJsService gridJsService) {
        super(gridJsService);
    }

    @GetMapping("/LoadSpreadsheet")
    public ResponseEntity<String> loadSpreadsheet(
            @RequestParam String filename,
            @RequestParam String uid) throws Exception {
        String fullPath = getFullFilePath(filename);
        StringBuilder json = _gridJsService.detailFileJsonWithUid(fullPath, uid);
        return ResponseEntity.ok()
            .contentType(MediaType.parseMediaType("text/plain;charset=UTF-8"))
            .body(json.toString());
    }

    // There is no path helper in GridJsControllerBase, and the base's
    // getFileCacheDirectory() is not one either - it returns the session directory.
    // The shipped workbook is resolved against src/main/resources/files/, which is
    // what the packaged classpath sees and what `mvn spring-boot:run` (working
    // directory = project root) reaches with this relative path.
    private String getFullFilePath(String filename) {
        return "./src/main/resources/files/" + filename;
    }
}
```

## Important Notes

- For Spring Boot 2.x (javax.servlet), extend `GridJsControllerBase`
- For Spring Boot 3.x (jakarta.servlet), extend `GridJsJakartaControllerBase`
- `GridJsControllerBase` provides built-in endpoints: /UpdateCell, /ImageUrl, /AddImage, /AddImageByURL, /CopyImage, /Download, /Ole
- The `uid` (unique identifier) is used for file caching - generate a UUID per session
- `com.google.code.gson:gson` must be a declared dependency: the gridjs extension uses Gson internally, and Spring fails to build the GridJsService bean with `NoClassDefFoundError: com/google/gson/JsonElement` when it is missing
- `FileCacheDirectory` must exist and be writable before the application starts
- Source workbooks live in `src/main/resources/files/` and the controller resolves an
  incoming `filename` against that directory only. Reading it from
  `getFileCacheDirectory()` instead looks in the session dump, where
  `Sample.xlsx` never appears, and `LoadSpreadsheet` fails with
  `FileNotFoundException` on a cold server
- `com.aspose.gridjs.GridJsService` is a Spring component in the jar: register it
  with `@ComponentScan`, never with your own `@Bean`
- Constructing `GridJsService` declares `throws java.lang.Exception`, so only the
  container may build it - a `@Bean` factory method would have to catch that
