# 10 – Configuration Guide  

This guide explains every **documented** configuration option for Aspose Cells GridJs (Python).  
For each option we describe:

* **What decision it controls** – the functional impact.  
* **When you should change it** – typical scenarios that justify a non‑default value.  
* **What can break** – symptoms you will see if the value is wrong.

The options are grouped by concern. All options exist both as **static** members of `aspose.cellsgridjs.Config` and as **instance** members of `aspose.cellsgridjs.GridJsOptions`. The same name and semantics apply to both; the examples below use the static `Config` class for brevity.

---

## 1 – Caching & Storage  

| Option | What it controls | When to change | What can break |
|--------|------------------|----------------|----------------|
| `file_cache_directory` | Directory where GridJs stores the cached workbook files. | You need a writable folder that survives application restarts (e.g., a dedicated temp folder). | If the path does not exist or the process lacks write permission, `FileNotFoundError` / `PermissionError` is raised when importing a workbook. |
| `picture_cache_directory` | Sub‑folder for cached picture streams (used when `GridJsWorkbook.cache_imp` is *null*). | Set a location with enough space for image caches, especially when many images are rendered. | Missing folder → image requests return HTTP 404; the server logs a `GridCellException` with code `IO`. |
| `cache_imp` (instance only) | Custom cache implementation (e.g., stream‑based storage). | You want to store cache in a database, cloud storage, or a custom file‑system layout. | An implementation that does not honor the `save_stream` / `load_stream` contract will cause `GridCellException` during lazy‑loading or image retrieval. |
| `islimit_shape_or_image` | Enables a hard limit on the number of shapes/images displayed per worksheet. | Large workbooks with thousands of drawings that cause memory pressure. | Setting to `false` may exhaust server memory; setting to `true` with too low limits truncates visible shapes. |
| `max_shape_or_image_count` | Maximum shapes/images displayed **per worksheet** when `islimit_shape_or_image` is `true`. | Tune when a worksheet contains many drawings but you only need a subset (e.g., 200). | If the count is lower than the actual number, excess shapes are silently omitted, leading to missing graphics in the UI. |
| `max_total_shape_or_image_count` | Maximum shapes/images displayed **in the whole workbook** when `islimit_shape_or_image` is `true`. | Same rationale as above, but workbook‑wide. | Same symptom: missing graphics; also may cause inconsistent UI if different sheets have different numbers of omitted objects. |
| `max_shape_or_image_width_or_height` | Upper bound (pixels) for a shape/image’s width or height when limiting is active. | Prevent extremely large raster images from consuming bandwidth. | Images larger than the limit are ignored → blank placeholders appear. |
| `max_pdf_save_seconds` | Timeout (seconds) for PDF export operations. | When exporting very large workbooks on a constrained server. | Export aborts with a `GridCellException` (code `INTERRUPTED`) after the timeout. |
| `load_time_out` | Timeout (milliseconds) for workbook loading. Default is `-1` (no timeout). | Protect against pathological files that take too long to parse. | Loading exceeding the limit raises a `GridCellException` (code `INTERRUPTED`). |
| `auto_optimize_for_large_cells` | When `true`, GridJs skips some style/border processing for worksheets with large cells to improve load speed. | Workbooks with many merged cells or huge column/row spans. | If you rely on precise border rendering, the visual appearance will be simplified. |

### Example – Setting a custom cache folder

```python
from aspose.cellsgridjs import Config

# Choose a folder that the process can write to
Config.set_file_cache_directory()
# The method has no parameters; the directory is taken from the environment variable
# or you can set it directly via the property (if exposed):
Config.file_cache_directory = "/var/tmp/gridjs_cache"
```

---

## 2 – Streaming & Lazy Loading  

| Option | What it controls | When to change | What can break |
|--------|------------------|----------------|----------------|
| `lazy_loading` | Loads **only the active worksheet** on the initial request. | Very large workbooks where loading all sheets would be slow or memory‑intensive. | If a client later requests a non‑loaded sheet, GridJs will fetch it on‑demand; however, if the cache is mis‑configured, the request may fail with `FileNotFoundError`. |
| `auto_fit_rows_height_on_load` | Auto‑fits **all row heights** after loading. Default is `false`. | When you need rows to display their full content immediately (e.g., wrapped text). | Setting to `true` can cause a noticeable performance hit on large sheets; the server may time out (`load_time_out`). |
| `auto_fit_columns_on_load` | Auto‑fits **all column widths** after loading. Default is `false`. | Same rationale as rows, but for column width. | Same performance impact; may also cause layout shifts that break custom UI expectations. |
| `save_html_as_zip` | When `true`, the HTML export is packaged as a ZIP archive. Default is `false`. | Required when the HTML contains many assets (images, CSS) that must stay together. | Clients expecting a plain `.html` file will receive a `.zip` and fail to render. |

### Example – Enabling lazy loading

```python
from aspose.cellsgridjs import Config

Config.set_lazy_loading()   # toggles the property; default is false
# After this call only the active sheet is loaded on the first request.
```

---

## 3 – Images, Shapes & Visual Rendering  

| Option | What it controls | When to change | What can break |
|--------|------------------|----------------|----------------|
| `skip_invisible_shapes` | Skips shapes that are invisible to the UI. Default is `true`. | When you need to preserve invisible objects for later programmatic access (e.g., redaction). | Setting to `false` may increase payload size and load time; setting to `true` hides those shapes from the UI. |
| `same_image_detecting` | Detects duplicate image sources and re‑uses the same cache entry. Default is `true`. | Workbooks that embed the same image many times – reduces memory and network traffic. | If set to `false`, each occurrence is stored separately, potentially exhausting cache space. |
| `enable_chart_client_rendering` | Enables client‑side chart rendering. When `false`, charts are rendered as static images on the server. Default is `true`. | Environments where the client has sufficient resources and you want interactive charts. | If the client lacks WebGL or the required JS libraries, charts will appear blank when the flag is `true`. |
| `ignore_empty_content` | When `true`, trailing rows/columns that contain only style (no data/formula) are omitted from the UI range. Default is `true`. | To hide large styled but empty regions that would otherwise inflate the visible grid. | Setting to `false` may cause the UI to display many empty rows/columns, degrading performance. |
| `use_print_area` | When `true`, the UI range respects the worksheet’s `PageSetup.PrintArea`. Default is `false`. | You want the UI to match the printed layout. | If the print area is mis‑configured, the UI may hide data unintentionally. |
| `show_chart_sheet` | Controls whether chart worksheets are exposed in the UI. Default is `false`. | When you need to let users view chart‑only sheets. | Enabling may expose internal chart data that you intended to keep hidden. |
| `empty_sheet_max_row` / `empty_sheet_max_col` | Default size for a completely empty worksheet (rows × columns). Defaults are 12 and 15 respectively. | When you want a larger/smaller initial grid for new worksheets. | Setting values that are too large may cause UI rendering lag; too small may truncate user‑added data before they can scroll. |
| `max_pdf_save_seconds` (see Caching) | Controls PDF export timeout. | Large PDFs that need more time. | See Caching section. |
| `auto_optimize_for_large_cells` (see Caching) | Optimizes loading for worksheets with large cells. | Workbooks with massive merged cells or very wide columns. | See Caching section. |

### Example – Disabling client‑side chart rendering

```python
from aspose.cellsgridjs import Config

# Force server‑generated chart images only
Config.set_enable_chart_client_rendering()
# The method toggles the property; after this call charts are rendered as static images.
```

---

## 4 – Logging  

| Option | What it controls | When to change | What can break |
|--------|------------------|----------------|----------------|
| `GridJsLogger.level` | Minimum log level (`DEBUG`, `INFO`, `WARN`, `ERROR`, `NONE`). Default is `None`. | During development (`DEBUG`) or production (`WARN`/`ERROR`). | If set to `NONE` you lose all diagnostic information, making troubleshooting impossible. |
| `GridJsLogger.enable_console_output` | Enables console logging. Default is `false`. | Quick debugging without configuring files. | No effect if `level` is `NONE`. |
| `GridJsLogger.enable_file_output` | Enables file logging. Must call `set_log_file` or `set_log_directory` first. Default is `false`. | Persistent logs for audit or post‑mortem analysis. | Enabling without a valid log path raises a `ValueError`. |
| `GridJsLogger.set_log_file(file_path)` | Sets a single log file (no rolling). | When you want a fixed log location. | If the directory does not exist or is not writable → `PermissionError`. |
| `GridJsLogger.set_log_directory(directory)` | Enables daily rolling logs in the given directory. | When you need log rotation. | Same permission issues as above; also, if the directory is on a read‑only volume, logging silently fails. |
| `GridJsLogger.max_file_size` | Maximum size (bytes) for a single log file before rolling (effective only with rolling enabled). Default is `0` (no size limit). | To limit disk usage on high‑traffic servers. | Setting too low may cause excessive file creation; setting to `0` may fill the disk. |

### Example – Configuring daily rolling logs

```python
from aspose.cellsgridjs import GridJsLogger, GridJsLogLevel, GridJsLogRolling

logger = GridJsLogger()
logger.set_level(GridJsLogLevel.INFO)
logger.set_log_directory("/var/log/gridjs")
logger.enable_file_output = True
logger.max_file_size = 10 * 1024 * 1024   # 10 MiB per file
```

---

## 5 – Collaboration (Co‑Work)  

| Option | What it controls | When to change | What can break |
|--------|------------------|----------------|----------------|
| `is_collaborative` | Enables collaborative editing (WebSocket‑based). Default is `false`. | When you need real‑time multi‑user editing. | If enabled but the server is not configured for WebSocket endpoints, clients will receive connection errors. |
| `base_route_name` | Base URL segment for GridJs controller routes. Default is `"/GridJs2"`. | To avoid route conflicts with other APIs or to match your existing routing scheme. | Mismatched route leads to 404 errors for all GridJs endpoints. |
| `message_topic` | WebSocket destination prefix (e.g., `"/topic/opr"`). Default is `"/topic/opr"`. | When you host multiple collaborative workspaces under different topics. | Incorrect topic prevents clients from receiving broadcast updates; changes appear only locally. |
| `redaction_use_client_generate_id` | When `true`, the client supplies its own IDs for redaction shapes instead of using the server‑generated shape ID. | When you need deterministic IDs for downstream processing (e.g., audit logs). | If the client generates duplicate IDs, redaction operations may overwrite each other, causing unexpected visual artifacts. |
| `custom_pdf_save_options` | Custom `PdfSaveOptions` object for PDF export. Default is `null`. | When you need specific PDF compliance, compression, or security settings. | Passing an incorrectly configured `PdfSaveOptions` (e.g., invalid `security_options`) raises a `GridCellException` during export. |

### Example – Enabling collaboration with a custom route

```python
from aspose.cellsgridjs import Config

Config.set_is_collaborative()          # toggles to true
Config.set_base_route_name()           # toggles; you can also assign directly:
Config.base_route_name = "/mygridjs"
Config.set_message_topic()             # toggles; assign directly if needed:
Config.message_topic = "/topic/mywork"
```

---

## 6 – Other Frequently Used Options  

| Option | What it controls | When to change | What can break |
|--------|------------------|----------------|----------------|
| `save_html_as_zip` | Determines whether HTML export is packaged as a ZIP. Default is `false`. | When you need a single downloadable archive containing HTML + assets. | Clients expecting a plain `.html` file will receive a ZIP and fail to render. |
| `empty_sheet_max_row` / `empty_sheet_max_col` | See Images & Shapes section. | When you want a larger initial grid for new sheets. | Too large values increase initial payload size; too small truncates user input before scrolling. |
| `auto_fit_rows_height_on_load` / `auto_fit_columns_on_load` | See Streaming section. | When you need immediate correct sizing at the cost of load time. | Performance degradation; possible timeout if `load_time_out` is low. |
| `max_pdf_save_seconds` | See Caching section. | When exporting very large PDFs. | Export aborts with a timeout exception. |
| `ignore_empty_content` | See Images & Shapes section. | When you want to keep styled empty rows/columns visible. | UI may become bloated with many empty rows/columns, slowing rendering. |
| `use_print_area` | See Images & Shapes section. | When you want UI to reflect the defined print area. | If the print area excludes needed data, the UI will hide it. |
| `show_chart_sheet` | See Images & Shapes section. | When you want chart‑only worksheets visible. | May expose internal chart data unintentionally. |
| `custom_pdf_save_options` | See Collaboration section. | When you need custom PDF settings (e.g., compliance, security). | Mis‑configuration leads to export failures. |
| `redaction_use_client_generate_id` | See Collaboration section. | When deterministic redaction IDs are required. | Duplicate IDs cause overwriting of redaction shapes. |

---

## 7 – Putting It All Together – Sample Startup Code  

Below is a complete, **runnable** Python snippet that configures a typical production environment:

```python
from aspose.cellsgridjs import Config, GridJsLogger, GridJsLogLevel

# ---------- Global static configuration ----------
# Cache locations
Config.file_cache_directory = "/var/tmp/gridjs_filecache"
Config.picture_cache_directory = "/var/tmp/gridjs_piccache"

# Enable lazy loading for large workbooks
Config.set_lazy_loading()

# Limit shape/image count to protect memory
Config.set_islimit_shape_or_image()
Config.max_shape_or_image_count = 200
Config.max_total_shape_or_image_count = 500

# Reduce image size for web delivery
Config.set_same_image_detecting()   # keep default true (deduplication)

# Enable collaboration
Config.set_is_collaborative()
Config.base_route_name = "/gridjs"
Config.message_topic = "/topic/workbook"

# Logging – daily rolling files, INFO level
logger = GridJsLogger()
logger.set_level(GridJsLogLevel.INFO)
logger.set_log_directory("/var/log/gridjs")
logger.enable_file_output = True

# Optional: custom PDF options (example – PDF/A‑1b compliance)
from aspose.cellsgridjs import PdfSaveOptions, PdfCompliance
pdf_opts = PdfSaveOptions()
pdf_opts.compliance = PdfCompliance.PDF_A1B
Config.custom_pdf_save_options = pdf_opts

# ---------- Instance‑specific options (if you use DI) ----------
# If you create a GridJsOptions object manually:
# from aspose.cellsgridjs import GridJsOptions
# opts = GridJsOptions()
# opts.lazy_loading = True
# opts.max_shape_or_image_count = 200
# ... (set any other per‑instance overrides)
```

**Key points in the sample**

* All paths are absolute and writable.  
* `set_lazy_loading()` is called **once**; the property toggles between `false` → `true`.  
* Limits are raised from the defaults (100/300) to accommodate a workbook with many drawings.  
* Collaboration is turned on and the route names are aligned with the host application.  
* Logging is configured to write daily‑rolled files; forgetting `set_log_directory` while `enable_file_output` is `true` would raise an exception.  

---

## 8 – Troubleshooting Checklist  

| Symptom | Likely mis‑configuration | Fix |
|---------|--------------------------|-----|
| **File import fails with `FileNotFoundError`** | `file_cache_directory` points to a non‑existent folder or lacks write permission. | Create the folder and grant write rights; verify the path. |
| **Images do not appear** | `picture_cache_directory` missing or `cache_imp` not implemented correctly. | Ensure the picture cache folder exists or provide a working `CacheImp` implementation. |
| **UI shows thousands of empty rows** | `ignore_empty_content` set to `false`. | Set back to `true` if you want empty styled rows hidden. |
| **Charts are blank** | `enable_chart_client_rendering` is `true` but the client lacks required JS libraries. | Either include the chart rendering library on the client or set `enable_chart_client_rendering` to `false`. |
| **Collaborative updates never appear** | `is_collaborative` is `true` but `base_route_name` or `message_topic` does not match the server’s WebSocket configuration. | Align the route and topic with your server’s WebSocket endpoints. |
| **Exported PDF is truncated or times out** | `max_pdf_save_seconds` too low for the workbook size. | Increase the timeout or optimize the workbook (e.g., enable `auto_optimize_for_large_cells`). |
| **Server logs are empty** | `GridJsLogger.enable_file_output` is `true` but no log file/directory was set. | Call `set_log_file` or `set_log_directory` before enabling file output. |
| **Redaction shapes disappear after a refresh** | `redaction_use_client_generate_id` is `false` and the server generates new IDs each session. | Set `redaction_use_client_generate_id` to `true` and ensure the client supplies stable IDs. |

---

## 9 – Defaults Summary  

| Option | Default (as documented) |
|--------|--------------------------|
| `save_html_as_zip` | `false` |
| `skip_invisible_shapes` | `true` |
| `lazy_loading` | `false` |
| `same_image_detecting` | `true` |
| `auto_optimize_for_large_cells` | `true` |
| `islimit_shape_or_image` | `true` |
| `max_shape_or_image_count` | `100` |
| `max_total_shape_or_image_count` | `300` |
| `max_shape_or_image_width_or_height` | `10000` |
| `max_pdf_save_seconds` | `10` |
| `ignore_empty_content` | `true` |
| `use_print_area` | `false` |
| `is_collaborative` | `false` |
| `enable_chart_client_rendering` | `true` |
| `custom_pdf_save_options` | `null` |
| `load_time_out` | `-1` (no timeout) |
| `show_chart_sheet` | `false` |
| `empty_sheet_max_row` | `12` |
| `empty_sheet_max_col` | `15` |
| `picture_cache_directory` | `"_piccache"` inside `file_cache_directory` |
| `base_route_name` | `"/GridJs2"` |
| `message_topic` | `"/topic/opr"` |
| `auto_fit_rows_height_on_load` | `false` |
| `auto_fit_columns_on_load` | `false` |
| `redaction_use_client_generate_id` | *not documented* (no default) |
| `font_folders` (instance only) | *not documented* |
| `cache_imp` (instance only) | *not documented* |

If a default is **not documented**, the guide explicitly states “default not documented”.

---

## 10 – Final Recommendations  

1. **Start with the documented defaults** – they are tuned for typical workloads.  
2. **Profile your workload** (size of workbooks, number of images, concurrency) before adjusting limits.  
3. **Always verify folder permissions** for any cache‑related path.  
4. **Enable logging early**; set a sensible `level` and a writable log directory to capture mis‑configurations.  
5. **Test collaboration locally** before deploying to production; mismatched routes are a common source of silent failures.  

By following this guide you can confidently tailor GridJs to your environment while avoiding the most common pitfalls.
