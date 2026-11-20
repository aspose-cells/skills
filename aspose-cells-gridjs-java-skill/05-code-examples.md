# Code Examples

Complete working examples from the demo project.

## Client - JavaScript

### gridjs-demo.js

**Path**: `src\main\resources\static\gridjs-demo.js`

```javascript
// src/main/resources/static/gridjs-demo.js
/**
 * GridJs Demo – client side logic.
 * Loads a workbook from the server, displays it, and enables server‑side updates.
 */
$(function () {
    // -------------------------------------------------------------
    // 1. Define URLs used by the demo
    // -------------------------------------------------------------
    const queryJsonUrl = "/GridJs/LoadSpreadsheet";
    const updateUrl = "/GridJs/UpdateCell";
    const fileDownloadUrl = "/GridJs/Download";
    const oleDownloadUrl = "/GridJs/Ole";
    const imageUrl = "/GridJs/ImageUrl";
    const imageUploadUrl1 = "/GridJs/AddImage";
    const imageUploadUrl2 = "/GridJs/AddImageByURL";
    const imageCopyUrl = "/GridJs/CopyImage";
    // -------------------------------------------------------------
    // 2. Global GridJs instance placeholder
    // -------------------------------------------------------------
    let xs; // will hold the x_spreadsheet instance
    // -------------------------------------------------------------
    // 3. Helper to generate a UUID (used as unique session id)
    // -------------------------------------------------------------
    function generateUUID() {
        // RFC4122 version 4 compliant UUID
        return ([1e7]+-1e3+-4e3+-8e3+-1e11).replace(/[018]/g, c =>
            (c ^ crypto.getRandomValues(new Uint8Array(1))[0] & 15 >> c / 4).toString(16)
        );
    }
    // -------------------------------------------------------------
    // 4. Load Spreadsheet button click handler
    // -------------------------------------------------------------
    $("#loadBtn").on("click", function () {
        const filename = $("#filename").val().trim();
        if (!filename) {
            alert("Please enter a file name.");
            return;
        }
        const uid = generateUUID(); // unique identifier for this session
        // Build request URL with query parameters
        const requestUrl = `${queryJsonUrl}?filename=${encodeURIComponent(filename)}&uid=${uid}`;
        // ---------------------------------------------------------
        // 5. AJAX request to obtain the JSON representation of the workbook
        // ---------------------------------------------------------
        $.ajax({
            url: requestUrl,
            method: "GET",
            dataType: "text", // server returns plain text JSON
            success: function (responseJsonString) {
                const jsondata = JSON.parse(responseJsonString);
                const option = {
                    updateMode: 'server',
                    updateUrl: updateUrl,
                    local: 'en' // UI language
                };
                loadWithOption(jsondata, option, uid);
            },
            error: function (xhr, status, err) {
                console.error("Failed to load spreadsheet:", err);
                alert("Error loading file: " + err);
            }
        });
    });
    // -------------------------------------------------------------
    // 6. Function that creates the GridJs UI and binds all URLs
    // -------------------------------------------------------------
    function loadWithOption(jsondata, option, uniqueId) {
        // Clear any previous instance
        $('#gridjs-demo-uid').empty();
        const sheets = jsondata.data;
        const filename = jsondata.filename;
		const activeSheetName = jsondata.actname;
        // Initialise the GridJs UI
        xs = x_spreadsheet('#gridjs-demo-uid', option)
            .loadData(sheets,activeSheetName)
            .updateCellError((msg) => console.error(msg));
        // Hide sheet tabs if server says not to show them
        if (!jsondata.showtabs) {
            xs.bottombar.hide();
        }
        // Associate the unique identifier and file name
        xs.setUniqueId(uniqueId);
        xs.setFileName(filename);
        // ---------------------------------------------------------
        // Set the active  cell (fallback to  A1)
        // ---------------------------------------------------------
        if (xs.getActiveSheetName() == activeSheetName) {
            xs.setActiveCell(jsondata.actrow, jsondata.actcol);
        } else {
			//some times if the active sheet is invisible or chartworksheet ,just use the first worksheet
            // If the requested sheet is hidden or a chart sheet,it will set the first visilbe sheet , just set A1 as active sheet
           xs.setActiveCell(0, 0);
        }
        // ---------------------------------------------------------
        // Bind auxiliary URLs (image handling, OLE, file download, etc.)
        // ---------------------------------------------------------
        xs.setImageInfo(imageUrl, imageUploadUrl1, imageUploadUrl2, imageCopyUrl, true);
        xs.setFileDownloadInfo(fileDownloadUrl);
        xs.setOleDownloadInfo(oleDownloadUrl);
        xs.setOpenFileUrl("/"); // optional: opens home page
    }
});
```

## Other

### maven-wrapper.properties

**Path**: `.mvn\wrapper\maven-wrapper.properties`

```text
wrapperVersion=3.3.4
distributionType=only-script
distributionUrl=https://repo.maven.apache.org/maven2/org/apache/maven/apache-maven/3.9.11/apache-maven-3.9.11-bin.zip
```

### pom.xml

**Path**: `pom.xml`

```text
<?xml version="1.0" encoding="UTF-8"?>
<project xmlns="http://maven.apache.org/POM/4.0.0" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
	xsi:schemaLocation="http://maven.apache.org/POM/4.0.0 https://maven.apache.org/xsd/maven-4.0.0.xsd">
	<modelVersion>4.0.0</modelVersion>
	<parent>
		<groupId>org.springframework.boot</groupId>
		<artifactId>spring-boot-starter-parent</artifactId>
		<!-- <version>3.5.7</version> -->
		<version>2.7.18</version> 
		<relativePath/> <!-- lookup parent from repository -->
	</parent>
	<groupId>com.example</groupId>
	<artifactId>gridjs-demo</artifactId>
	<version>26.8</version>
	<name>gridjs-demo</name>
	<description>Demo project for gridjs using Spring Boot</description>
	<url/>
	<licenses>
		<license/>
	</licenses>
	<developers>
		<developer/>
	</developers>
	<scm>
		<connection/>
		<developerConnection/>
		<tag/>
		<url/>
	</scm>
	<properties>
		<java.version>17</java.version>
	</properties>
	<dependencies>
		<dependency>
			<groupId>org.springframework.boot</groupId>
			<artifactId>spring-boot-starter-web</artifactId>
		</dependency>

		<dependency>
			<groupId>org.springframework.boot</groupId>
			<artifactId>spring-boot-starter-test</artifactId>
			<scope>test</scope>
		</dependency>
		 <!-- Aspose.Cells core library -->
    <dependency>
        <groupId>com.aspose</groupId>
        <artifactId>aspose-cells</artifactId>
        <version>26.8</version>
    </dependency>
    <!-- Aspose.Cells GridJs extension -->
    <dependency>
        <groupId>com.aspose</groupId>
        <artifactId>aspose-cells</artifactId>
        <classifier>gridjs</classifier>
        <version>26.8</version>
    </dependency>
    <!-- Gson for JSON handling -->
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
	<build>
		<plugins>
			<plugin>
				<groupId>org.springframework.boot</groupId>
				<artifactId>spring-boot-maven-plugin</artifactId>
			</plugin>
		</plugins>
	</build>

</project>
```

### GridJsConfig.java

**Path**: `src\main\java\com\example\gridjsdemo\config\GridJsConfig.java`

```text
package com.example.gridjsdemo.config;

import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;

import com.aspose.gridjs.GridJsOptions;

/**
 * Configures Aspose.Cells GridJs options.
 */
@Configuration
public class GridJsConfig {

    @Bean
    public GridJsOptions gridJsOptions() {
        GridJsOptions options = new GridJsOptions();

        // Directory where temporary files (e.g., JSON cache) are stored
        options.setFileCacheDirectory("D:/storage/Aspose.Cells.GridJs/");


        // Base route used by the GridJs controller
        options.setBaseRouteName("/GridJs");

        return options;
    }
}
```

### DemoApplication.java

**Path**: `src\main\java\com\example\gridjsdemo\DemoApplication.java`

```text
package com.example.gridjsdemo;

import org.springframework.boot.SpringApplication;

import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.annotation.ComponentScan;
@ComponentScan(basePackages = { "com.example.gridjsdemo", "com.aspose.gridjs" })
@SpringBootApplication
public class DemoApplication {

	public static void main(String[] args) throws Exception {
 
		 SpringApplication.run(DemoApplication.class, args);
	}
	
}
```

### application.properties

**Path**: `src\main\resources\application.properties`

```text
spring.application.name=demo
```

### index.html

**Path**: `src\main\resources\static\index.html`

```text
<!-- src/main/resources/static/index.html -->
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>GridJs Demo – Load Spreadsheet</title>
    <!-- jQuery and jQuery‑UI -->
    <script src="https://code.jquery.com/jquery-2.1.1.min.js"></script>
    <script src="https://code.jquery.com/ui/1.12.1/jquery-ui.min.js"></script>
    <link rel="stylesheet"
          href="https://code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css"/>
    <!-- JSZip (required for export) -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.6.0/jszip.min.js"></script>
    <!-- GridJs Spreadsheet UI -->
    <link rel="stylesheet"
          href="https://unpkg.com/gridjs-spreadsheet/xspreadsheet.css"/>
    <script src="https://unpkg.com/gridjs-spreadsheet/xspreadsheet.js"></script> 
    <!-- Demo specific script -->
    <script src="./gridjs-demo.js"></script>
    
    <style>
        body {font-family: Arial, sans-serif; margin: 20px;}
        #gridjs-demo-uid {width: 100%; height: 800px;}
    </style>
</head>
<body>
<h1>GridJs Spreadsheet Demo</h1>
<!-- Container for the GridJs UI -->
<div id="gridjs-demo-uid"></div>
<!-- Simple UI to pick a file from the server -->
<div style="margin-top:20px;">
    <label for="filename">Spreadsheet file (in server folder): </label>
    <input type="text" id="filename" value="Sample.xlsx"/>
    <button id="loadBtn">Load Spreadsheet</button>
</div>
</body>
</html>
```

### DemoApplicationTests.java

**Path**: `src\test\java\com\example\demo\DemoApplicationTests.java`

```text
package com.example.demo;

import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;

@SpringBootTest
class DemoApplicationTests {

	@Test
	void contextLoads() {
	}

}
```

## Server - Controller

### GridJsController.java

**Path**: `src\main\java\com\example\gridjsdemo\controller\GridJsController.java`

```text
package com.example.gridjsdemo.controller;

import com.aspose.gridjs.GridJsControllerBase;
import com.aspose.gridjs.GridJsService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.InputStreamResource;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;
import javax.servlet.http.HttpServletRequest;
/**
 * REST controller that handles GridJs operations such as loading a workbook,
 * updating cells, managing images, and downloading files.
 */
@RestController
@RequestMapping("/GridJs")
public class GridJsController extends GridJsControllerBase {
    @Autowired
    public GridJsController(GridJsService gridJsService) {
        super(gridJsService);
    }
    /**
     * Loads a workbook and returns its JSON representation together with a unique id.
     *
     * @param filename Name of the Excel file located under the configured root folder.
     * @param uid      Unique identifier generated on the client side.
     * @throws Exception 
     */
    @GetMapping("/LoadSpreadsheet")
    public ResponseEntity<String> loadSpreadsheet(@RequestParam String filename,
                                                  @RequestParam String uid) throws Exception {
        String fullFilePath = getFullFilePath(filename);
        StringBuilder json = _gridJsService.detailFileJsonWithUid(fullFilePath, uid);
        return ResponseEntity.ok()
                .header("Content-Type", "text/plain; charset=UTF-8")
                .body(json.toString());
    }
    /**
     * Resolve the physical file path. Adjust the base folder to match your environment.
     */
    private String getFullFilePath(String filename) {
        // Example: All Excel files are placed under src/main/resources/files/
        return "./src/main/resources/files/" + filename;
    }
    // -------------------------------------------------------------------------
    // The following endpoints simply delegate to the base implementation.
    // -------------------------------------------------------------------------
    @PostMapping("/UpdateCell")
    public ResponseEntity<String> updateCell(HttpServletRequest request) {
        try {
            return super.updateCell(request);
        } catch (Exception e) {
            e.printStackTrace();
            return ResponseEntity.badRequest().body(e.getMessage());
        }
    }
    @PostMapping("/AddImage")
    public ResponseEntity<String> addImage(@RequestParam(value = "image", required = false) MultipartFile file,
                                           @RequestParam("uid") String uid,
                                           @RequestParam("p") String p,
                                           @RequestParam(value = "control", required = false) String isControl) {
        return super.addImage(file, uid, p, isControl);
    }
    @PostMapping("/CopyImage")
    public ResponseEntity<String> copyImage(HttpServletRequest request) {
        return super.copyImage(request);
    }
    @PostMapping("/AddImageByURL")
    public ResponseEntity<String> addImageByUrl(HttpServletRequest request) {
        return super.addImageByUrl(request);
    }
    @GetMapping("/Image")
    public ResponseEntity<InputStreamResource> getImage(HttpServletRequest request) {
        return super.getImage(request);
    }
    @GetMapping("/ImageUrl")
    public ResponseEntity<String> getImageUrl(@RequestParam String id,
                                               @RequestParam String uid) {
        return super.getImageUrl(id, uid);
    }
    @GetMapping("/Ole")
    public ResponseEntity<?> getOle(HttpServletRequest request) {
        return super.getOle(request);
    }
    @GetMapping("/GetFile")
    public ResponseEntity<?> getFile(@RequestParam("id") String id) {
        return super.getFile(id);
    }
    @PostMapping("/Download")
    public ResponseEntity<String> download(HttpServletRequest request) {
        return super.download(request);
    }
}
```

