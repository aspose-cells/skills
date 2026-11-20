# 10 – Configuration Guide  

This guide explains every **documented** configuration option for Aspose Cells GridJs (Java).  
Options are grouped by concern; each entry tells you:

* **What decision it represents** – why the option exists.  
* **When to change it** – scenarios where the default (when documented) or the out‑of‑the‑box value should be overridden.  
* **What breaks if it is wrong** – typical symptoms or exceptions you will see.

> **Important** – If a default value is not listed in the reference, the default is *not documented*. Do not assume any value.

---

## 1 Caching  

| Option | Where to set | Decision | When to change | What breaks if wrong |
|--------|--------------|----------|----------------|----------------------|
| `CacheImp` | `GridJsOptions` (instance) | Supplies a custom implementation of `GridCacheForStream` used for all file‑stream caching (JSON, images, PDFs, etc.). | You need a bespoke cache (e.g., Redis, database, distributed file system) or want to control cache eviction. | If the implementation does not correctly implement `saveStream`, `loadStream`, `isExisted`, or `getFileUrl`, the service will throw `java.lang.NullPointerException` or `java.io.IOException` when trying to read/write cached data. |
| `JsonCache` | `GridJsOptions` (instance) | Holds a custom JSON cache (`com.aspose.gridjs.y_`). | You want to replace the built‑in in‑memory JSON cache with a shared cache (e.g., Hazelcast). | Missing or mismatched cache leads to stale JSON being served; updates may be lost, causing “cell update failed” errors from the client. |
| `PictureCacheDirectory` | `Config` (static) – `setPictureCacheDirectory(String)` | Directory on the server file system where picture streams are cached. | Deploying to an environment where the default temp folder is not writable or you need to isolate picture data. | If the path is invalid or unwritable, image‑related endpoints (`/Image`, `/AddImage`) will throw `java.io.IOException`. |
| `FileCacheDirectory` | `Config` (static) – `setFileCacheDirectory(String)` | Directory for generic file caching (PDF, HTML, etc.). | Same reasons as picture cache; also useful for large‑scale PDF export where you want a dedicated storage area. | An invalid path causes `java.lang.Exception` from `setFileCacheDirectory` and later failures when saving PDFs (`saveToPdf`). |
| `BaseRouteName` | `Config` (static) – `setBaseRouteName(String)` | Prefix used for all GridJs controller routes (e.g., `/gridjs/...`). | You need to avoid route collisions in a multi‑module Spring Boot app or want a versioned API path. | If the route does not match the client’s `setImageInfo`, `setLazyLoadingUrl`, etc., HTTP 404 errors will appear. |
| `MessageTopic` | `Config` (static) – `setMessageTopic(String)` | Topic name used by the collaborative messaging bus (e.g., Kafka, RabbitMQ). | Switching the underlying message broker or isolating workbooks per tenant. | Collaborative operations (`GridJsCoService`) will fail with `java.lang.Exception` because messages are published to a non‑existent topic. |
| `RedactionUseClientGenerateId` | `Config` (static) **and** `GridJsOptions` (instance) | When `true`, the client generates redaction IDs; otherwise the server does. | You want deterministic IDs for client‑side replay or testing. | If the client expects to generate IDs but the server generates them (or vice‑versa), redaction updates will be rejected with `GridJsPermissionException` or “invalid redaction id”. |

### Example – Custom Cache Implementation  

```java
import com.aspose.gridjs.GridCacheForStream;
import java.io.InputStream;
import java.io.OutputStream;

public class MyCache extends GridCacheForStream {
    @Override
    public void saveStream(InputStream stream, String key) {
        // store stream in Redis, DB, etc.
    }

    @Override
    public InputStream loadStream(String key) {
        // retrieve stream from storage
        return null; // replace with real stream
    }

    @Override
    public boolean isExisted(String key) {
        // check existence
        return false;
    }

    @Override
    public String getFileUrl(String key) {
        // return public URL if applicable
        return "/mycache/" + key;
    }
}

// Wiring the cache into GridJs
GridJsOptions options = new GridJsOptions();
options.setCacheImp(new MyCache());
// pass `options` to GridJsService constructor
```

---

## 2 Streaming & Lazy Loading  

| Option | Where to set | Decision | When to change | What breaks if wrong |
|--------|--------------|----------|----------------|----------------------|
| `LazyLoading` | `Config` (static) – `setLazyLoading(boolean)` **or** `GridJsOptions` – `setLazyLoading(boolean)` | Enables the “lazy loading” mode where only the visible sheet region is sent to the client; the rest is fetched on demand. | Workbooks with many sheets or very large sheets where initial payload size hurts performance. | If set to `true` but the client never calls `/LazyLoading` endpoint, the UI will show blank cells and log “lazy loading request missing”. |
| `SameImageDetecting` | `Config` – `setSameImageDetecting(boolean)` **or** `GridJsOptions` – `setSameImageDetecting(boolean)` | Detects duplicate images and stores only one copy in the cache. | Reducing storage for workbooks that embed the same logo many times. | When disabled, each duplicate image consumes extra cache space, possibly exceeding `MaxShapeOrImageCount` limits and causing “shape limit exceeded” errors. |
| `AutoOptimizeForLargeCells` | `Config` – `setAutoOptimizeForLargeCells(boolean)` **or** `GridJsOptions` – `setAutoOptimizeForLargeCells(boolean)` | Optimizes rendering of cells that contain large amounts of text or data (e.g., long comments). | Workbooks with cells that hold megabytes of text; you want faster client rendering. | If left `false`, large cells may cause the client to freeze or the server to exceed `MaxPdfSaveSeconds` when exporting. |
| `IgnoreEmptyContent` | `Config` – `setIgnoreEmptyContent(boolean)` **or** `GridJsOptions` – `setIgnoreEmptyContent(boolean)` | When `true`, empty rows/columns are omitted from the JSON payload. | Reducing payload size for sparse worksheets. | If set to `true` but the client expects a full grid (e.g., for formulas that reference empty cells), you may see “reference out of range” errors. |
| `AutoFitRowsHeightOnLoad` | `Config` – `setAutoFitRowsHeightOnLoad(boolean)` **or** `GridJsOptions` – `setAutoFitRowsHeightOnLoad(boolean)` | Auto‑fits row heights based on content during the initial load. | You want rows to display fully without manual resizing after load. | When `true` on extremely tall rows, the server may take long to compute heights, leading to timeouts (`java.lang.Exception: operation timed out`). |
| `AutoFitColumnsOnLoad` | `Config` – `setAutoFitColumnsOnLoad(boolean)` **or** `GridJsOptions` – `setAutoFitColumnsOnLoad(boolean)` | Auto‑fits column widths on load. | Same rationale as rows; useful for data‑driven column widths. | Incorrect widths may cause horizontal scroll glitches if the client’s CSS expects a fixed width. |
| `MaxPdfSaveSeconds` | `Config` – `setMaxPdfSaveSeconds(int)` **or** `GridJsOptions` – `setMaxPdfSaveSeconds(int)` | Upper bound (seconds) for PDF export operations. | Preventing runaway PDF generation on huge workbooks. | If the limit is too low, legitimate PDF exports will abort with `java.lang.Exception: PDF generation timed out`. |
| `SaveHtmlAsZip` | `Config` – `setSaveHtmlAsZip(boolean)` **or** `GridJsOptions` – `setSaveHtmlAsZip(boolean)` | When `true`, HTML export is packaged as a ZIP file (HTML + resources). | You need a single downloadable artifact. | If the client expects a plain HTML file but receives a ZIP, the download will be corrupted. |
| `PictureCacheDirectory` / `FileCacheDirectory` (see Caching) | – | – | – | – |

### Example – Enabling Lazy Loading  

```java
GridJsOptions options = new GridJsOptions();
options.setLazyLoading(true);          // enable lazy loading
options.setSameImageDetecting(true);   // deduplicate images
options.setAutoFitRowsHeightOnLoad(true);
options.setAutoFitColumnsOnLoad(true);

// Build the service
GridJsService service = new GridJsService(options);
```

---

## 3 Images & Shapes  

| Option | Where to set | Decision | When to change | What breaks if wrong |
|--------|--------------|----------|----------------|----------------------|
| `SkipInvisibleShapes` | `Config` – `setSkipInvisibleShapes(boolean)` **or** `GridJsOptions` – `setSkipInvisibleShapes(boolean)` | Omits shapes whose `visible` flag is `false` from the JSON payload. | Reducing payload for workbooks that contain many hidden objects (e.g., templates). | If set to `true` but later a hidden shape is made visible via an API, the client will never receive it, leading to “shape not found” errors. |
| `IslimitShapeOrImage` | `Config` – `setIslimitShapeOrImage(boolean)` **or** `GridJsOptions` – `setIslimitShapeOrImage(boolean)` | Turns on the shape/image count limits (`MaxShapeOrImageCount`, `MaxTotalShapeOrImageCount`). | You want to enforce a hard ceiling to protect memory. | When `true` and the workbook exceeds the limits, the service throws `GridJsPermissionException` with `LIMITATION`. |
| `MaxShapeOrImageCount` | `Config` – `setMaxShapeOrImageCount(int)` **or** `GridJsOptions` – `setMaxShapeOrImageCount(int)` | Maximum number of **individual** shapes or images allowed per sheet. | Preventing a single sheet from exploding with thousands of objects. | Exceeding the limit results in `java.lang.Exception: shape count limit exceeded`. |
| `MaxTotalShapeOrImageCount` | `Config` – `setMaxTotalShapeOrImageCount(int)` **or** `GridJsOptions` – `setMaxTotalShapeOrImageCount(int)` | Maximum total number of shapes/images across the whole workbook. | Same as above but workbook‑wide. | Same failure mode as `MaxShapeOrImageCount`. |
| `MaxShapeOrImageWidthOrHeight` | `Config` – `setMaxShapeOrImageWidthOrHeight(int)` **or** `GridJsOptions` – `setMaxShapeOrImageWidthOrHeight(int)` | Upper bound (pixels) for any shape or image dimension. | Guarding against extremely large raster images that would blow memory. | Oversized images are rejected with `java.lang.Exception: image size exceeds limit`. |
| `EnableChartClientRendering` | `Config` – `setEnableChartClientRendering(boolean)` **or** `GridJsOptions` – `setEnableChartClientRendering(boolean)` | When `true`, charts are rendered on the client using Canvas instead of being rasterized on the server. | Reducing server CPU load for chart‑heavy workbooks. | If the client browser does not support the required Canvas features, charts appear blank. |
| `ShowChartSheet` | `Config` – `setShowChartSheet(boolean)` **or** `GridJsOptions` – `setShowChartSheet(boolean)` | Controls whether a dedicated “Chart” sheet is generated for each chart object. | You want charts to appear as separate sheets for printing. | When `true` but the client UI hides chart sheets, users may see unexpected extra tabs. |
| `RedactionUseClientGenerateId` | `Config` / `GridJsOptions` | Determines who generates redaction IDs (client vs. server). | See *Caching* section. | Mismatch leads to “invalid redaction id” errors. |

### Example – Limiting Shapes  

```java
GridJsOptions options = new GridJsOptions();
options.setIslimitShapeOrImage(true);
options.setMaxShapeOrImageCount(500);          // per sheet
options.setMaxTotalShapeOrImageCount(2000);    // workbook‑wide
options.setMaxShapeOrImageWidthOrHeight(4096); // pixels
```

---

## 4 Logging  

| Option | Where to set | Decision | When to change | What breaks if wrong |
|--------|--------------|----------|----------------|----------------------|
| `Level` | `GridJsLogger.setLevel(int)` | Controls the minimum log severity (`DEBUG`, `INFO`, `WARN`, `ERROR`, `NONE`). | During development you may set `DEBUG`; in production you usually set `WARN` or `ERROR`. | If set to `DEBUG` in production, log files can grow rapidly and exhaust disk space. |
| `EnableConsoleOutput` | `GridJsLogger.setEnableConsoleOutput(boolean)` | Sends log entries to `System.out`. | Debugging locally without a file logger. | When disabled and no file logger is configured, you lose all logs. |
| `EnableFileOutput` | `GridJsLogger.setEnableFileOutput(boolean)` | Enables writing logs to a file. | Production environments where persistent logs are required. | If disabled unintentionally, no log files are created, making troubleshooting impossible. |
| `LogFile` | `GridJsLogger.setLogFile(String)` **or** `setLogFile(String, boolean)` | Path of the log file. The boolean overload indicates whether to append (`true`) or overwrite (`false`). | Custom log location (e.g., `/var/log/aspose/gridjs.log`). | Invalid path or insufficient permissions cause `java.io.IOException` when the logger tries to open the file; subsequent logging calls are silently dropped. |
| `LogDirectory` | `GridJsLogger.setLogDirectory(String)` **or** overloads with rolling policy (`int`) and optional prefix (`String`). | Directory where log files are stored; optional rolling (`DAILY`, `MONTHLY`, `NONE`). | Centralized log directories or when you need daily log rotation. | Wrong directory leads to the same failure as `LogFile`. |
| `MaxFileSize` | `GridJsLogger.setMaxFileSize(long)` | Maximum size (bytes) before the logger rolls over (if rolling is enabled). | Preventing a single log file from growing without bound. | If set too low, frequent rollovers may cause performance degradation; if too high, disk exhaustion may occur. |
| `LogFilePath` / `MaxFileSize` (getter) | `GridJsLogger.getLogFilePath()`, `getMaxFileSize()` | Read‑only access to current log file location and size limit. | Useful for diagnostics. | N/A – getters never fail. |

### Example – Production Logging  

```java
import com.aspose.gridjs.GridJsLogger;
import com.aspose.gridjs.GridJsLogLevel;
import com.aspose.gridjs.GridJsLogRolling;

public class DemoLoggingSetup {
    public static void configure() {
        // Log only warnings and errors
        GridJsLogger.setLevel(GridJsLogLevel.WARN);
        // Write to a dedicated file and enable daily rolling
        GridJsLogger.setLogDirectory("/var/log/aspose", GridJsLogRolling.DAILY);
        GridJsLogger.setLogFile("/var/log/aspose/gridjs.log", true);
        GridJsLogger.setEnableFileOutput(true);
        GridJsLogger.setEnableConsoleOutput(false);
        // Limit each log file to 100 MiB
        GridJsLogger.setMaxFileSize(100L * 1024L * 1024L);
    }
}
```

---

## 5 Collaboration  

| Option | Where to set | Decision | When to change | What breaks if wrong |
|--------|--------------|----------|----------------|----------------------|
| `Collaborative` | `Config` – `setCollaborative(boolean)` **or** `GridJsOptions` – `setCollaborative(boolean)` | Turns on the collaborative editing engine (real‑time sync via `GridJsCoService`). | Deploying a multi‑user editing scenario (e.g., shared spreadsheets). | If enabled but the `CoWorkUserProvider` bean is missing, the service throws `java.lang.Exception: CoWorkUserProvider not configured`. |
| `MessageTopic` | `Config` – `setMessageTopic(String)` | Name of the messaging topic used for broadcasting collaborative operations. | Switching message brokers or isolating tenants. | Mismatched topic leads to no operation propagation; users see “no other users’ changes”. |
| `BaseRouteName` | `Config` – `setBaseRouteName(String)` | Affects the URL prefix for collaborative endpoints (`/UpdateCell`, `/AddImage`, etc.). | When you need a versioned API (`/v2/gridjs/...`). | If the client’s `setImageInfo`/`setLazyLoadingUrl` still point to the old prefix, HTTP 404 errors occur. |
| `RedactionUseClientGenerateId` | (see *Caching*) | Influences redaction ID generation in collaborative mode. | When you need deterministic IDs for conflict resolution. | Inconsistent ID generation causes `GridJsPermissionException` with `PERMISSION` code during collaborative updates. |

### Example – Enabling Collaboration  

```java
import com.aspose.gridjs.CoWorkUserProvider;
import com.aspose.gridjs.GridJsCoService;
import com.aspose.gridjs.GridJsOptions;
import com.aspose.gridjs.Config;

// 1. Provide a user provider (Spring bean)
public class DemoUserProvider implements CoWorkUserProvider {
    @Override
    public String getCurrentUserName() { return "alice"; }

    @Override
    public long getCurrentUserId() { return 42L; }

    @Override
    public int getPermission() {
        // EDITABLE = 1 (from CoWorkUserPermission)
        return 1;
    }
}

// 2. Configure collaborative mode
GridJsOptions options = new GridJsOptions();
options.setCollaborative(true);
Config.setMessageTopic("gridjs-collab");
Config.setBaseRouteName("/gridjs");

// 3. Build the service (the constructor will pick up the provider via DI)
GridJsService service = new GridJsService(options);
```

---

## 6 Miscellaneous Workbook‑Level Settings  

These settings are **static** on `Config` (global) but affect workbook processing. They are not grouped elsewhere but are still important.

| Option | Where to set | Decision | When to change | What breaks if wrong |
|--------|--------------|----------|----------------|----------------------|
| `UsePrintArea` | `Config` – `setUsePrintArea(boolean)` | When `true`, only the defined print area is exported (PDF/HTML). | You want to honor the user‑defined print area for printing. | If `true` and no print area is defined, the export may produce an empty file. |
| `ShowChartSheet` | (see *Images & Shapes*) | – | – | – |
| `EmptySheetMaxRow` / `EmptySheetMaxCol` | `Config` – `setEmptySheetMaxRow(int)`, `setEmptySheetMaxCol(int)` | Upper limits for rows/columns in a completely empty sheet (used to avoid generating massive empty grids). | Preventing accidental creation of sheets with millions of rows/cols. | If set too low, legitimate large empty sheets will be truncated, causing missing data. |
| `FontFolders` | `Config` – `setFontFolder(String, boolean)` / `setFontFolders(String[], boolean)` | Directories that contain custom fonts used during PDF/HTML rendering. The boolean indicates whether to load sub‑folders recursively. | Deploying custom corporate fonts. | Missing fonts cause fallback to default fonts; PDF may render with unexpected typography. |
| `CustomPdfSaveOptions` | `Config` – `setCustomPdfSaveOptions(PdfSaveOptions)` | Allows you to supply a fully configured `com.aspose.cells.PdfSaveOptions` (e.g., compliance, image quality). | Fine‑tuning PDF output (e.g., PDF/A compliance). | If the options are incompatible (e.g., invalid compliance level), PDF export throws `java.lang.Exception`. |
| `MaxPdfSaveSeconds` | (see *Streaming*) | – | – | – |
| `RedactionUseClientGenerateId` | (see *Caching*) | – | – | – |

---

## 7 Putting It All Together  

A typical Spring Boot configuration class that wires **all** relevant options:

```java
package com.example.demo;

import com.aspose.gridjs.*;
import com.aspose.gridjs.y_; // placeholder for the actual JSON cache type
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;

@Configuration
public class DemoGridJsConfig {

    @Bean
    public GridJsOptions gridJsOptions() {
        GridJsOptions opts = new GridJsOptions();

        // Caching
        opts.setCacheImp(new MyCache());               // custom stream cache
        opts.setJsonCache(new y_());                   // custom JSON cache (implementation supplied)

        // Streaming & lazy loading
        opts.setLazyLoading(true);
        opts.setSameImageDetecting(true);
        opts.setAutoOptimizeForLargeCells(true);
        opts.setIgnoreEmptyContent(true);
        opts.setAutoFitRowsHeightOnLoad(true);
        opts.setAutoFitColumnsOnLoad(true);
        opts.setMaxPdfSaveSeconds(120);                // 2 minutes

        // Images & shapes
        opts.setSkipInvisibleShapes(true);
        opts.setIslimitShapeOrImage(true);
        opts.setMaxShapeOrImageCount(800);
        opts.setMaxTotalShapeOrImageCount(3000);
        opts.setMaxShapeOrImageWidthOrHeight(4096);
        opts.setEnableChartClientRendering(true);
        opts.setShowChartSheet(false);

        // Collaboration
        opts.setCollaborative(true);
        Config.setMessageTopic("gridjs-demo");
        Config.setBaseRouteName("/gridjs");

        // Redaction ID strategy
        Config.setRedactionUseClientGenerateId(true);
        opts.setRedactionUseClientGenerateId(true);

        return opts;
    }

    @Bean
    public GridJsService gridJsService(GridJsOptions gridJsOptions) throws Exception {
        return new GridJsService(gridJsOptions);
    }

    @Bean
    public CoWorkUserProvider coWorkUserProvider() {
        return new DemoUserProvider(); // see the DemoUserProvider class above
    }

    static {
        // Global static configuration (run once at startup)
        Config.setPictureCacheDirectory("/tmp/gridjs/pictures");
        Config.setFileCacheDirectory("/tmp/gridjs/files");
        Config.setFontFolder("/opt/app/fonts", true);
        Config.setEnableConsoleOutput(false);
        Config.setEnableFileOutput(true);
        Config.setLogDirectory("/var/log/aspose", GridJsLogRolling.DAILY);
        Config.setLogFile("/var/log/aspose/gridjs.log", true);
        Config.setLevel(GridJsLogLevel.WARN);
    }
}
```

---

### Quick Checklist  

| Concern | Checklist item |
|---------|----------------|
| **Caching** | Custom `CacheImp`? Custom `JsonCache`? Valid picture/file cache directories? |
| **Streaming** | `LazyLoading` enabled for large workbooks? `SameImageDetecting` matches your storage strategy? `MaxPdfSaveSeconds` high enough for expected PDFs? |
| **Images & Shapes** | Shape limits (`IslimitShapeOrImage`, `MaxShapeOrImageCount`, `MaxTotalShapeOrImageCount`) set according to expected workbook size. `MaxShapeOrImageWidthOrHeight` prevents OOM. |
| **Logging** | Log level appropriate for environment. Log directory writable. Rolling policy set if needed. |
| **Collaboration** | `Collaborative` true only when `CoWorkUserProvider` bean present. `MessageTopic` and `BaseRouteName` match client configuration. |
| **Redaction** | `RedactionUseClientGenerateId` consistent between server and client. |
| **Misc** | `UsePrintArea` matches export expectations. `FontFolders` contain required fonts. `CustomPdfSaveOptions` validated. |

---

**Remember:** When you change any of these options, restart the Spring Boot application so the static `Config` values are re‑read and the `GridJsService` is reconstructed with the new `GridJsOptions`. Incorrect values will surface as exceptions during workbook load, export, image handling, or collaborative sync – consult the stack trace for the specific option mentioned above.
