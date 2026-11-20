# Getting Started with Aspose.Cells GridJs (Java)

This guide shows the **minimum** steps required to add GridJs to an existing Spring Boot application, verify that the server starts, and display a workbook (`Sample.xlsx`) in the browser.

---

## 1️⃣ Prerequisites

| Item | Minimum version |
|------|-----------------|
| Java | 17 or newer |
| Spring Boot | 2.7.x **or** 3.x (choose the matching base controller) |
| Maven | 3.6+ |
| Internet access | To download the Aspose Maven repository and npm assets |

> **Note** – GridJs uses **Gson** internally. If the `gson` dependency is missing the server fails with `NoClassDefFoundError: com/google/gson/JsonElement`.

---

## 2️⃣ Maven dependencies

Add the following entries to your `pom.xml`. The **gridjs** classifier pulls in the GridJs extension that contains all server‑side endpoints.

```xml
<!-- pom.xml -->
<dependencies>
    <!-- Spring Web -->
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

    <!-- Gson – required at runtime -->
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

---

## 3️⃣ Create a writable cache directory

GridJs stores per‑session JSON files under a directory you configure.  
Create the folder **before** the application starts, e.g.:

```bash
mkdir -p ./cache
```

The folder must be readable and writable by the process that runs Spring Boot.

---

## 4️⃣ Minimal server configuration

### 4.1 Application entry point

```java
// DemoApplication.java
package com.example.gridjsdemo;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.annotation.ComponentScan;

@SpringBootApplication
@ComponentScan(basePackages = { "com.example.gridjsdemo", "com.aspose.gridjs" })
public class DemoApplication {
    public static void main(String[] args) {
        SpringApplication.run(DemoApplication.class, args);
    }
}
```

*Why the `@ComponentScan`?*  
`GridJsService` is already a Spring component inside the Aspose JAR. Scanning the `com.aspose.gridjs` package registers it automatically; defining your own `@Bean` would cause a bean‑definition conflict.

### 4.2 GridJs options bean

```java
// GridJsConfig.java
package com.example.gridjsdemo.config;

import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import com.aspose.gridjs.GridJsOptions;

/**
 * Application‑specific GridJs configuration.
 */
@Configuration
public class GridJsConfig {

    @Bean
    public GridJsOptions gridJsOptions() {
        GridJsOptions options = new GridJsOptions();

        // Must point to an existing, writable folder (step 3)
        options.setFileCacheDirectory("./cache");

        // Base URL that the client will call (must match @RequestMapping)
        options.setBaseRouteName("/GridJs");

        return options;
    }
}
```

### 4.3 Controller

For Spring Boot 2.x (javax.servlet) extend `GridJsControllerBase`.  
If you are on Spring Boot 3.x, replace the base class with `GridJsJakartaControllerBase`.

```java
// GridJsController.java
package com.example.gridjsdemo.controller;

import com.aspose.gridjs.GridJsControllerBase;
import com.aspose.gridjs.GridJsService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import javax.servlet.http.HttpServletRequest;

/**
 * Minimal controller exposing the built‑in GridJs endpoints
 * and a custom LoadSpreadsheet operation.
 */
@RestController
@RequestMapping("/GridJs")
public class GridJsController extends GridJsControllerBase {

    @Autowired
    public GridJsController(GridJsService gridJsService) {
        super(gridJsService);
    }

    /**
     * Returns the JSON representation of an Excel workbook together with a UID.
     *
     * @param filename workbook name located under src/main/resources/files/
     * @param uid      client‑generated unique identifier
     */
    @GetMapping("/LoadSpreadsheet")
    public ResponseEntity<String> loadSpreadsheet(@RequestParam String filename,
                                                  @RequestParam String uid) throws Exception {
        String fullPath = getFullFilePath(filename);
        StringBuilder json = _gridJsService.detailFileJsonWithUid(fullPath, uid);
        return ResponseEntity.ok()
                .header("Content-Type", "text/plain; charset=UTF-8")
                .body(json.toString());
    }

    /** Resolve the physical path of a workbook. */
    private String getFullFilePath(String filename) {
        // Workbooks are placed under src/main/resources/files/
        return "./src/main/resources/files/" + filename;
    }
}
```

### 4.4 Place the demo workbook

Add `Sample.xlsx` to the project:

```
src/main/resources/files/Sample.xlsx
```

The controller above resolves any `filename` against this folder, so the file **must** exist there.

---

## 5️⃣ Client‑side assets (static)

All static files belong under `src/main/resources/static/`.

### 5.1 `index.html`

```html
<!-- src/main/resources/static/index.html -->
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>GridJs Demo</title>

    <!-- jQuery (required by the demo script) -->
    <script src="https://code.jquery.com/jquery-2.1.1.min.js"></script>

    <!-- GridJs UI assets -->
    <link rel="stylesheet"
          href="https://unpkg.com/gridjs-spreadsheet/xspreadsheet.css"/>
    <script src="https://unpkg.com/gridjs-spreadsheet/xspreadsheet.js"></script>

    <!-- Demo script -->
    <script src="./gridjs-demo.js"></script>

    <style>
        body {font-family: Arial, sans-serif; margin: 20px;}
        #gridjs-demo-uid {width: 100%; height: 800px;}
    </style>
</head>
<body>
<h1>GridJs Spreadsheet Demo</h1>

<div id="gridjs-demo-uid"></div>

<div style="margin-top:20px;">
    <label for="filename">Workbook (under server folder): </label>
    <input type="text" id="filename" value="Sample.xlsx"/>
    <button id="loadBtn">Load Spreadsheet</button>
</div>
</body>
</html>
```

### 5.2 `gridjs-demo.js`

```javascript
// src/main/resources/static/gridjs-demo.js
$(function () {
    const queryJsonUrl = "/GridJs/LoadSpreadsheet";
    const updateUrl = "/GridJs/UpdateCell";

    // Placeholder for the GridJs instance
    let xs;

    // Simple RFC‑4122 v4 UUID generator (used as session UID)
    function generateUUID() {
        return ([1e7]+-1e3+-4e3+-8e3+-1e11).replace(/[018]/g, c =>
            (c ^ crypto.getRandomValues(new Uint8Array(1))[0] & 15 >> c / 4).toString(16)
        );
    }

    $("#loadBtn").on("click", function () {
        const filename = $("#filename").val().trim();
        if (!filename) {
            alert("Enter a file name.");
            return;
        }
        const uid = generateUUID();
        const requestUrl = `${queryJsonUrl}?filename=${encodeURIComponent(filename)}&uid=${uid}`;

        $.ajax({
            url: requestUrl,
            method: "GET",
            dataType: "text",
            success: function (response) {
                const jsondata = JSON.parse(response);
                const option = {
                    updateMode: "server",
                    updateUrl: updateUrl,
                    local: "en"
                };
                loadWithOption(jsondata, option, uid);
            },
            error: function (xhr, status, err) {
                console.error("Load failed:", err);
                alert("Error loading file: " + err);
            }
        });
    });

    function loadWithOption(jsondata, option, uid) {
        $('#gridjs-demo-uid').empty();
        const sheets = jsondata.data;
        const filename = jsondata.filename;
        const activeSheetName = jsondata.actname;

        xs = x_spreadsheet('#gridjs-demo-uid', option)
            .loadData(sheets, activeSheetName)
            .updateCellError(msg => console.error(msg));

        if (!jsondata.showtabs) xs.bottombar.hide();

        xs.setUniqueId(uid);
        xs.setFileName(filename);
        xs.setActiveCell(jsondata.actrow, jsondata.actcol);
    }
});
```

The script:

1. Generates a **UID** for the session.  
2. Calls `/GridJs/LoadSpreadsheet` with `filename` and `uid`.  
3. Receives the JSON payload and creates the GridJs UI (`x_spreadsheet`).  
4. Binds the server‑side update endpoint (`/GridJs/UpdateCell`).

---

## 6️⃣ Verify the installation step‑by‑step

| Step | Command / Action | Expected result |
|------|-------------------|-----------------|
| **6.1** | `mvn clean package` | Maven builds without errors (Gson, Aspose JARs resolved). |
| **6.2** | `mvn spring-boot:run` | Console shows `Started DemoApplication` and no `FileNotFoundException` about the cache directory. |
| **6.3** | Open <http://localhost:8080/> in a browser | The static `index.html` loads (no 404). |
| **6.4** | Click **Load Spreadsheet** (default `Sample.xlsx`) | The grid appears with the workbook’s sheets, rows, and columns. |
| **6.5** | Edit a cell and press **Enter** | The change is sent to `/GridJs/UpdateCell` and persisted on the server (no error toast). |

If any step fails, consult the **Checklist** below.

---

## 7️⃣ Checklist – Grid renders empty

When the UI shows a blank area (no sheets, rows, or columns), check the following three items in order:

1. **Cache directory**  
   * Verify `./cache` exists and is writable.  
   * Look for a startup exception like `java.io.FileNotFoundException` from `GridJsOptions.setFileCacheDirectory`.

2. **Workbook location & name**  
   * Confirm `src/main/resources/files/Sample.xlsx` is present and readable.  
   * Ensure the request URL contains the exact file name (`filename=Sample.xlsx`).  
   * Server logs will show `FileNotFoundException` if the path is wrong.

3. **UID handling**  
   * The client must send a non‑empty `uid` query parameter.  
   * If `uid` is missing or empty, `detailFileJsonWithUid` creates no session folder, leading to an empty JSON response.  
   * Check the network tab for the GET request to `/GridJs/LoadSpreadsheet` – the response body should contain a JSON object with a `data` array.

If all three items are correct and the grid is still empty, inspect the server console for `GridCellException` or any stack trace that indicates JSON conversion failure.

---

### 🎉 You’re ready!

You now have a **working GridJs integration** that can load an Excel workbook, display it in the browser, and push cell edits back to the server. From here you can explore advanced features such as image handling, collaborative editing, or custom calculation engines—all built on the same minimal foundation.
