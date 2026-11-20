# Code Examples

Complete working examples from the demo project.

## Client - JavaScript

### vite.config.js

**Path**: `ClientApp\react-gridjs\vite.config.js`

```javascript
import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

const backend = 'http://127.0.0.1:2022';

export default defineConfig({
  base: '/demos/react/',
  plugins: [react()],
  server: {
    proxy: {
      '/GridJs': { target: backend, changeOrigin: true },
    },
  },
  build: {
    outDir: '../../static/demos/react',
    emptyOutDir: true,
  },
});
```

### main.js

**Path**: `ClientApp\vanilla-gridjs-npm\main.js`

```javascript
import Spreadsheet from 'gridjs-spreadsheet';
import JSZip from 'jszip';
import 'gridjs-spreadsheet/xspreadsheet.css';
import './styles.css';

const workbookName = 'Sample.xlsx';
const launcher = import.meta.env.DEV ? 'http://127.0.0.1:2022/' : '/';
const openFileUrl = import.meta.env.DEV ? 'http://127.0.0.1:2022/legacy' : '/legacy';
const host = document.getElementById('spreadsheet');
const loadingState = document.getElementById('loading-state');
const errorState = document.getElementById('error-state');
const errorMessage = document.getElementById('error-message');
let spreadsheet = null;

window.JSZip = JSZip;
document.getElementById('back-link').href = launcher;
document.getElementById('retry-button').addEventListener('click', loadWorkbook);
window.addEventListener('beforeunload', () => spreadsheet?.destroy?.());

function createUid() {
  return globalThis.crypto?.randomUUID?.() || `npm-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

async function loadWorkbook() {
  spreadsheet?.destroy?.();
  spreadsheet = null;
  host.innerHTML = '';
  loadingState.hidden = false;
  errorState.hidden = true;

  try {
    const query = new URLSearchParams({ filename: workbookName, uid: createUid() });
    const response = await fetch(`/GridJs/LoadSpreadsheet?${query}`);
    if (!response.ok) throw new Error(`${response.status} ${response.statusText}`);
    const payload = JSON.parse(await response.text());

    spreadsheet = new Spreadsheet(host, {
      updateMode: 'server',
      updateUrl: '/GridJs/UpdateCell',
      mode: 'edit',
      local: 'en',
      showToolbar: true,
      showContextmenu: true,
    });
    spreadsheet.loadData(payload.data || [], payload.actname);
    spreadsheet.setUniqueId?.(payload.uniqueid || '');
    spreadsheet.setFileName?.(payload.filename || workbookName);
    spreadsheet.setImageInfo?.('/GridJs/ImageUrl', '/GridJs/AddImage', '/GridJs/AddImageByURL', '/GridJs/CopyImage', 5678, '/content/img/loading.gif');
    spreadsheet.setFileDownloadInfo?.('/GridJs/Download');
    spreadsheet.setOleDownloadInfo?.('/GridJs/Ole');
    spreadsheet.setLazyLoadingUrl?.('/GridJs/LazyLoadingStreamJson');
    if (payload.actname) spreadsheet.setActiveSheetByName?.(payload.actname);
    spreadsheet.setActiveCell?.(payload.actrow || 0, payload.actcol || 0);
    spreadsheet.setOpenFileUrl?.(openFileUrl);
    spreadsheet.updateCellError?.((message) => console.error('[GridJS npm]', message));
    loadingState.hidden = true;
    console.log('[GridJS npm] ready', payload.filename);
  } catch (reason) {
    loadingState.hidden = true;
    errorMessage.textContent = reason.message || String(reason);
    errorState.hidden = false;
  }
}

loadWorkbook();
```

### vite.config.js

**Path**: `ClientApp\vanilla-gridjs-npm\vite.config.js`

```javascript
import { defineConfig } from 'vite';

const backend = 'http://127.0.0.1:2022';

export default defineConfig({
  base: '/demos/npm/',
  server: {
    proxy: {
      '/GridJs': { target: backend, changeOrigin: true },
    },
  },
  build: {
    outDir: '../../static/demos/npm',
    emptyOutDir: true,
  },
});
```

### main.js

**Path**: `ClientApp\vue-gridjs\src\main.js`

```javascript
import { createApp } from 'vue';
import 'gridjs-spreadsheet/xspreadsheet.css';
import './styles.css';
import App from './App.vue';

createApp(App).mount('#app');
```

### vite.config.js

**Path**: `ClientApp\vue-gridjs\vite.config.js`

```javascript
import { defineConfig } from 'vite';
import vue from '@vitejs/plugin-vue';

const backend = 'http://127.0.0.1:2022';

export default defineConfig({
  base: '/demos/vue/',
  plugins: [vue()],
  server: {
    proxy: {
      '/GridJs': { target: backend, changeOrigin: true },
    },
  },
  build: {
    outDir: '../../static/demos/vue',
    emptyOutDir: true,
  },
});
```

### gridjs-demo.js

**Path**: `static\js\gridjs-demo.js`

```javascript
/* static/js/gridjs-demo.js */
/* -------------------------------------------------------------
   Client‑side script that initializes GridJs, loads the workbook,
   and wires up all server‑side URLs.
   ------------------------------------------------------------- */

(function ($) {
    // ---------- 1. Helper: generate a UUID ----------
    function generateUUID() {
        // Simple RFC4122 version 4 UUID
        return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function (c) {
            var r = Math.random() * 16 | 0,
                v = c === 'x' ? r : (r & 0x3 | 0x8);
            return v.toString(16);
        });
    }

    // ---------- 2. Define server URLs ----------
    const queryJsonUrl      = "/GridJs/LoadSpreadsheet";
    const updateUrl         = "/GridJs/UpdateCell";
    const fileDownloadUrl   = "/GridJs/Download";
    const oleDownloadUrl    = "/GridJs/Ole";
    const imageurl          = "/GridJs/ImageUrl";
    const imageuploadurl1   = "/GridJs/AddImage";
    const imageuploadurl2   = "/GridJs/AddImageByURL";
    const imagecopyurl      = "/GridJs/CopyImage";

    const zorder = 1000;               // Canvas Zorder
    let xs;                            // GridJs instance
    const uid = generateUUID();        // Unique session id

    // ---------- 3. Load workbook JSON ----------
    $.ajax({
        url: queryJsonUrl,
        method: "GET",
        data: { uid: uid },
        dataType: "json",
        // The server returns gzip‑compressed JSON; browsers decompress automatically.
        success: function (responseData) {
            const jsondata = typeof responseData === "string"
                ? JSON.parse(responseData)
                : responseData;

            const option = {
                updateMode: "server",
                updateUrl: updateUrl,
                // UI language (change as needed)
                local: 'en',
            };

            loadWithOption(jsondata, option);
        },
        error: function (xhr) {
            console.error("Failed to load spreadsheet:", xhr);
        }
    });

    // ---------- 4. Render the UI ----------
    function loadWithOption(jsondata, option) {
        $('#gridjs-demo-uid').empty();

        const sheets = jsondata.data;
        const filename = jsondata.filename;
        const activeSheetName = jsondata.actname;
        // Initialise GridJs (x_spreadsheet) and bind to the container,load data and set active sheet
        xs = x_spreadsheet('#gridjs-demo-uid', option)
            .loadData(sheets, activeSheetName)
            .updateCellError(msg => console.error(msg));

        // Hide the bottom sheet bar when tabs are disabled
        if (!jsondata.showtabs) {
            xs.bottombar.hide();
        }

        xs.setUniqueId(jsondata.uniqueid);
        xs.setFileName(filename);
    // ----------------------------------------------------------------
    //  4. Set the active cell (fallback to A1 if needed)
    // ----------------------------------------------------------------
   
    if (xs.getActiveSheetName() == activeSheetName)  {
        xs.setActiveCell(jsondata.actrow, jsondata.actcol);
    } else {
        // Fallback – set A1 as active cell
        xs.setActiveCell(0, 0);
    }

        // ---------- 5. Register auxiliary URLs ----------
        xs.setImageInfo(
            imageurl,
            imageuploadurl1,
            imageuploadurl2,
            imagecopyurl,
            zorder,
            "/image/loading.gif"
        );
        xs.setFileDownloadInfo(fileDownloadUrl);
        xs.setOleDownloadInfo(oleDownloadUrl);
        xs.setOpenFileUrl("/GridJs/Index");

    }

})(jQuery);
```

## Other

### app.py

**Path**: `app.py`

```text
# app.py
# -------------------------------------------------------------
# Flask application that hosts Aspose.Cells GridJs services.
# -------------------------------------------------------------

from flask import (
    Flask,
    render_template,
    request,
    jsonify,
    Response,
    send_file,
    send_from_directory,
    abort,
)
import io
import gzip
import mimetypes
import os

# Aspose.Cells GridJs imports
from aspose.cellsgridjs import GridJsOptions, GridJsService

BASE_DIRECTORY = os.path.dirname(os.path.abspath(__file__))
app = Flask(__name__)


def resolve_path(value: str) -> str:
    if os.path.isabs(value):
        return value
    return os.path.join(BASE_DIRECTORY, value)


WORKBOOK_DIRECTORY = resolve_path(os.environ.get("GRIDJS_WORKBOOK_DIR", os.path.join("static", "files")))
DEFAULT_WORKBOOK = os.environ.get("GRIDJS_WORKBOOK", "Sample.xlsx")
CACHE_DIRECTORY = resolve_path(os.environ.get("GRIDJS_CACHE_DIR", "cache"))

# -------------------------------------------------------------
# GridJs service initialization
# -------------------------------------------------------------
options = GridJsOptions()
options.file_cache_directory = CACHE_DIRECTORY     # Cache folder (must exist)
options.base_route_name = "/GridJs"                # Base route used by the client


# Ensure cache folder exists
os.makedirs(options.file_cache_directory, exist_ok=True)

gridjs_service = GridJsService(options)

# -------------------------------------------------------------
# Helper: Guess MIME type from a filename
# -------------------------------------------------------------
def guess_mime_type_from_filename(filename: str) -> str:
    mime_type, _ = mimetypes.guess_type(filename)
    return mime_type or "application/octet-stream"

# -------------------------------------------------------------
# Root route – renders the UI
# -------------------------------------------------------------
@app.route("/")
def index():
    return render_template("frameworks.html")


@app.route("/legacy")
def legacy():
    return render_template("index.html")


@app.route("/health")
def health():
    return Response("Healthy", content_type="text/plain; charset=utf-8")


@app.route("/content/img/<filename>")
@app.route("/image/<filename>")
def gridjs_status_image(filename):
    if filename not in {"loading.gif", "updating.gif"}:
        abort(404)
    return send_from_directory(
        os.path.join(app.static_folder, "content", "img"),
        filename,
    )


@app.route("/demos/<framework>/")
def framework_demo(framework):
    if framework not in {"react", "vue", "angular", "npm"}:
        abort(404)
    return send_from_directory(
        os.path.join(app.static_folder, "demos", framework),
        "index.html",
    )


@app.route("/demos/<framework>/content/img/<filename>")
def framework_demo_status_image(framework, filename):
    if framework not in {"react", "vue", "angular", "npm"}:
        abort(404)
    if filename not in {"loading.gif", "updating.gif"}:
        abort(404)
    return send_from_directory(
        os.path.join(app.static_folder, "content", "img"),
        filename,
    )


@app.route("/demos/<framework>/<path:filename>")
def framework_demo_asset(framework, filename):
    if framework not in {"react", "vue", "angular", "npm"}:
        abort(404)
    return send_from_directory(
        os.path.join(app.static_folder, "demos", framework),
        filename,
    )

# -------------------------------------------------------------
# 1️⃣ Load spreadsheet JSON (gzipped)
# -------------------------------------------------------------
@app.route("/GridJs/LoadSpreadsheet", methods=["GET"])
@app.route("/GridJs2/LoadSpreadsheet", methods=["GET"])
def load_spreadsheet():
    uid = request.args.get("uid", "")
    filename = request.args.get("filename", DEFAULT_WORKBOOK)
    if not uid:
        return jsonify({"error": "uid is required"}), 400
    if filename != os.path.basename(filename):
        return jsonify({"error": "invalid filename"}), 400
    workbook_path = os.path.join(WORKBOOK_DIRECTORY, filename)
    # Prepare an in‑memory gzip stream
    gzip_buffer = io.BytesIO()
    with gzip.GzipFile(fileobj=gzip_buffer, mode="w") as gz:
        # Populate the gzip stream with GridJs JSON data
        gridjs_service.detail_stream_json_with_uid(gz, workbook_path, uid)
    gzip_buffer.seek(0)
    return Response(
        gzip_buffer.getvalue(),
        mimetype="application/json",
        headers={"Content-Encoding": "gzip"},
    )

# -------------------------------------------------------------
# 2️⃣ Update a cell (POST)
# -------------------------------------------------------------
@app.route("/GridJs/UpdateCell", methods=["POST"])
@app.route("/GridJs2/UpdateCell", methods=["POST"])
def update_cell():
    p = request.form.get("p")
    uid = request.form.get("uid")
    ret = gridjs_service.update_cell(p, uid)
    return Response(ret, content_type="text/plain; charset=utf-8")

# -------------------------------------------------------------
# 3️⃣ Add image (multipart/form‑data)
# -------------------------------------------------------------
@app.route("/GridJs/AddImage", methods=["POST"])
@app.route("/GridJs2/AddImage", methods=["POST"])
def add_image():
    uid = request.form.get("uid")
    p = request.form.get("p")
    is_control = request.form.get("control")
    file = request.files.get("image")          # May be None – handled below
    file_bytes = io.BytesIO(file.read()) if file else None
    ret = gridjs_service.add_image(p, uid, is_control, file_bytes)
    return jsonify(ret)

# -------------------------------------------------------------
# 4️⃣ Copy image
# -------------------------------------------------------------
@app.route("/GridJs/CopyImage", methods=["POST"])
@app.route("/GridJs2/CopyImage", methods=["POST"])
def copy_image():
    uid = request.form.get("uid")
    p = request.form.get("p")
    ret = gridjs_service.copy_image(p, uid)
    return jsonify(ret)

# -------------------------------------------------------------
# 5️⃣ Add image by external URL
# -------------------------------------------------------------
@app.route("/GridJs/AddImageByURL", methods=["POST"])
@app.route("/GridJs2/AddImageByURL", methods=["POST"])
def add_image_by_url():
    uid = request.form.get("uid")
    p = request.form.get("p")
    image_url = request.form.get("imageurl")
    ret = gridjs_service.add_image_by_url(p, uid, image_url)
    return jsonify(ret)

# -------------------------------------------------------------
# 6️⃣ Retrieve an image (GET)
# -------------------------------------------------------------
@app.route("/GridJs/Image", methods=["GET"])
@app.route("/GridJs2/Image", methods=["GET"])
def image():
    img_id = request.args.get("id")
    uid = request.args.get("uid")
    if not img_id or not uid:
        return "Missing required parameters", 400
    image_bytes = gridjs_service.image(uid, img_id)
    return send_file(
        image_bytes,
        as_attachment=False,
        download_name=img_id,
        mimetype="image/png",
    )

# -------------------------------------------------------------
# 7️⃣ Retrieve an embedded OLE object (GET)
# -------------------------------------------------------------
@app.route("/GridJs/Ole", methods=["GET"])
@app.route("/GridJs2/Ole", methods=["GET"])
def ole():
    obj_id = request.args.get("id")
    uid = request.args.get("uid")
    sheet = request.args.get("sheet")
    filename = None
    file_bytes = gridjs_service.ole(uid, sheet, obj_id, filename)
    if filename:
        return send_file(
            io.BytesIO(file_bytes),
            as_attachment=True,
            download_name=filename,
            mimetype=guess_mime_type_from_filename(filename),
        )
    abort(400, "File not found")

# -------------------------------------------------------------
# 8️⃣ Batch image URLs (GET)
# -------------------------------------------------------------
@app.route("/GridJs/ImageUrl", methods=["GET"])
@app.route("/GridJs2/ImageUrl", methods=["GET"])
def image_url():
    img_id = request.args.get("id")
    uid = request.args.get("uid")
    ret = gridjs_service.image_url(options.base_route_name, img_id, uid)
    return jsonify(ret)

# -------------------------------------------------------------
# 9️⃣ Download cached file (GET)
# -------------------------------------------------------------
@app.route("/GridJs/GetFile", methods=["GET"])
@app.route("/GridJs2/GetFile", methods=["GET"])
def get_file():
    file_id = request.args.get("id")
    file_bytes = gridjs_service.get_file(file_id)
    return send_file(
        file_bytes,
        as_attachment=True,
        download_name=file_id,
        mimetype=guess_mime_type_from_filename(file_id),
    )

# -------------------------------------------------------------
# 🔟 Trigger file download (POST)
# -------------------------------------------------------------
@app.route("/GridJs/Download", methods=["POST"])
@app.route("/GridJs2/Download", methods=["POST"])
def download():
    p = request.form.get("p")
    uid = request.form.get("uid")
    file_name = request.form.get("file")
    ret = gridjs_service.download(p, uid, file_name)
    return jsonify(ret)

# -------------------------------------------------------------
# Application entry point
# -------------------------------------------------------------
if __name__ == "__main__":
    app.run(
        host="0.0.0.0",
        port=int(os.environ.get("PORT", "2022")),
        debug=os.environ.get("FLASK_DEBUG", "").lower() in {"1", "true", "yes"},
    )
```

### app.component.html

**Path**: `ClientApp\angular-gridjs\src\app\app.component.html`

```text
<main class="editor-shell">
  <header class="demo-bar">
    <a class="back-link" [href]="launcherUrl" aria-label="Back to framework examples">← Examples</a>
    <div class="demo-title">
      <span class="framework-mark" aria-hidden="true">A</span>
      <span><strong>Angular</strong><small>{{ workbookName }}</small></span>
    </div>
    <span class="npm-label">npm component</span>
  </header>

  <section class="grid-stage" aria-live="polite">
    <div *ngIf="!workbook() && !error()" class="state-panel loading-state" role="status">
      <span class="loading-line wide"></span>
      <span class="loading-line"></span>
      <p>Loading workbook from Flask…</p>
    </div>
    <div *ngIf="error()" class="state-panel error-state" role="alert">
      <strong>Workbook could not be loaded</strong>
      <p>{{ error() }}</p>
      <button type="button" (click)="loadWorkbook()">Retry loading</button>
    </div>
    <gridjs-spreadsheet
      *ngIf="workbook()"
      [data]="workbook()"
      height="calc(100dvh - 64px)"
      mode="edit"
      locale="en"
      [showToolbar]="true"
      [showContextmenu]="true"
      (ready)="handleReady($event)"
      (change)="logChange()"
      (error)="handleError($event)"
    ></gridjs-spreadsheet>
  </section>
</main>
```

### index.html

**Path**: `ClientApp\angular-gridjs\src\index.html`

```text
<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8" />
    <title>GridJS Angular example</title>
    <base href="/" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <meta name="description" content="Aspose.Cells GridJS Angular example backed by Flask and Python via .NET." />
  </head>
  <body>
    <app-root></app-root>
  </body>
</html>
```

### index.html

**Path**: `ClientApp\react-gridjs\index.html`

```text
<!doctype html>
<html lang="en">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <meta name="description" content="Aspose.Cells GridJS React example backed by Flask and Python via .NET." />
    <title>GridJS React example</title>
  </head>
  <body>
    <div id="root"></div>
    <script type="module" src="/src/main.jsx"></script>
  </body>
</html>
```

### index.html

**Path**: `ClientApp\vanilla-gridjs-npm\index.html`

```text
<!doctype html>
<html lang="en">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <meta name="description" content="Aspose.Cells GridJS vanilla npm example backed by Flask and Python via .NET." />
    <title>GridJS npm example · Python via .NET</title>
  </head>
  <body>
    <main class="editor-shell">
      <header class="demo-bar">
        <a id="back-link" class="back-link" href="/" aria-label="Back to framework examples">← Examples</a>
        <div class="demo-title">
          <span class="framework-mark" aria-hidden="true">JS</span>
          <span><strong>JavaScript</strong><small>Sample.xlsx</small></span>
        </div>
        <span class="npm-label">npm import</span>
      </header>
      <section id="grid-stage" class="grid-stage" aria-live="polite">
        <div id="loading-state" class="state-panel loading-state" role="status">
          <span class="loading-line wide"></span>
          <span class="loading-line"></span>
          <p>Loading workbook from Flask…</p>
        </div>
        <div id="error-state" class="state-panel error-state" role="alert" hidden>
          <strong>Workbook could not be loaded</strong>
          <p id="error-message"></p>
          <button id="retry-button" type="button">Retry loading</button>
        </div>
        <div id="spreadsheet"></div>
      </section>
    </main>
    <script type="module" src="/main.js"></script>
  </body>
</html>
```

### index.html

**Path**: `ClientApp\vue-gridjs\index.html`

```text
<!doctype html>
<html lang="en">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <meta name="description" content="Aspose.Cells GridJS Vue example backed by Flask and Python via .NET." />
    <title>GridJS Vue example</title>
  </head>
  <body>
    <div id="app"></div>
    <script type="module" src="/src/main.js"></script>
  </body>
</html>
```

### requirements.txt

**Path**: `requirements.txt`

```text
requests  
flask
aspose-cells-gridjs-net-python
gunicorn>=23,<24
```

### frameworks.html

**Path**: `templates\frameworks.html`

```text
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <meta name="description" content="Aspose.Cells GridJS Simple examples for Python via .NET with React, Vue, Angular, npm, and direct JavaScript integrations.">
  <title>Framework examples · Aspose.Cells GridJS Simple for Python via .NET</title>
  <link rel="icon" href="{{ url_for('static', filename='frameworks/aspose-logo.svg') }}" type="image/svg+xml">
  <link rel="stylesheet" href="{{ url_for('static', filename='frameworks/site.css') }}">
</head>
<body>
  <a class="skip-link" href="#main-content">Skip to main content</a>
  <header class="site-header">
    <div class="site-header__inner">
      <a class="brand" href="/">
        <img class="brand__logo" src="{{ url_for('static', filename='frameworks/aspose-logo.svg') }}" alt="" width="34" height="34" aria-hidden="true">
        <span>Aspose.Cells GridJS Simple for Python</span>
      </a>
      <nav class="site-nav" aria-label="Primary navigation">
        <a href="/">Examples</a>
        <a href="/legacy">Original Simple demo</a>
        <a href="https://docs.aspose.com/cells/python-net/aspose-cells-gridjs/">Documentation</a>
      </nav>
    </div>
  </header>

  <main id="main-content" class="site-main">
    <section class="hero" aria-labelledby="page-title">
      <div>
        <p class="eyebrow">Flask · Python via .NET · npm</p>
        <h1 id="page-title">One Python backend.<br>Four client integrations.</h1>
        <p class="hero__copy">React, Vue, Angular, and vanilla JavaScript open the same <code>Sample.xlsx</code> workbook and reuse the existing <code>/GridJs</code> server operations.</p>
      </div>
      <aside class="run-note" aria-label="Run information">
        <span class="run-note__label">One-command start</span>
        <strong>Built static demos</strong>
        <p>Create the Python environment, build all clients, and start Flask.</p>
        <code>npm start</code>
      </aside>
    </section>

    <section class="examples-section" aria-labelledby="examples-title">
      <div class="section-heading">
        <div>
          <p class="eyebrow">Choose a client</p>
          <h2 id="examples-title">Framework examples</h2>
        </div>
        <span class="environment-chip">PYTHON</span>
      </div>

      <div class="demo-grid">
        <a class="demo-card" href="/demos/react/">
          <span class="demo-card__mark" aria-hidden="true">R</span>
          <span class="demo-card__content"><strong>React</strong><span>React lifecycle wrapper using the published npm adapter.</span></span>
          <span class="demo-card__arrow" aria-hidden="true">↗</span>
        </a>
        <a class="demo-card" href="/demos/vue/">
          <span class="demo-card__mark" aria-hidden="true">V</span>
          <span class="demo-card__content"><strong>Vue 3</strong><span>Composition API integration with cleanup and events.</span></span>
          <span class="demo-card__arrow" aria-hidden="true">↗</span>
        </a>
        <a class="demo-card" href="/demos/angular/">
          <span class="demo-card__mark" aria-hidden="true">A</span>
          <span class="demo-card__content"><strong>Angular 22</strong><span>Standalone component using signals and the Angular adapter.</span></span>
          <span class="demo-card__arrow" aria-hidden="true">↗</span>
        </a>
        <a class="demo-card" href="/demos/npm/">
          <span class="demo-card__mark" aria-hidden="true">JS</span>
          <span class="demo-card__content"><strong>JavaScript + npm</strong><span>Direct <code>Spreadsheet</code> import without a framework wrapper.</span></span>
          <span class="demo-card__arrow" aria-hidden="true">↗</span>
        </a>
      </div>
    </section>

    <section class="legacy-demo" aria-labelledby="legacy-title">
      <div>
        <p class="eyebrow">JavaScript direct import</p>
        <h2 id="legacy-title">&lt;script&gt; + CDN example</h2>
        <p>The original Python Simple page loads <code>xspreadsheet.js</code> directly from a CDN, with no npm build step.</p>
      </div>
      <a class="secondary-link" href="/legacy">Open original demo <span aria-hidden="true">→</span></a>
    </section>
  </main>

  <footer class="site-footer">
    <span>&copy; 2026 Aspose.Cells</span>
    <span>Flask backend with four npm clients and one direct JavaScript example.</span>
  </footer>
</body>
</html>
```

### index.html

**Path**: `templates\index.html`

```text
<!-- templates/index.html -->
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>GridJS JavaScript CDN Demo</title>
    <link rel="icon" href="{{ url_for('static', filename='frameworks/aspose-logo.svg') }}" type="image/svg+xml">

    <!-- jQuery & UI -->
    <script src="https://code.jquery.com/jquery-2.1.1.min.js"></script>
    <script src="https://code.jquery.com/ui/1.12.1/jquery-ui.min.js"></script>
    <link rel="stylesheet"
          href="https://code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">

    <!-- JSZip (required by GridJs) -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.6.0/jszip.min.js"></script>

    <!-- GridJs Spreadsheet UI -->
    <link rel="stylesheet"
          href="https://unpkg.com/gridjs-spreadsheet/xspreadsheet.css">
    <script src="https://unpkg.com/gridjs-spreadsheet/xspreadsheet.js"></script>

    <!-- Demo helper script -->
    <script src="{{ url_for('static', filename='js/gridjs-demo.js') }}"></script>

    <style>
        body {font-family: Arial, sans-serif; margin: 20px;}
        #gridjs-demo-uid {border: 1px solid #ccc;}
    </style>
</head>
<body>
    <h1>GridJS JavaScript &lt;script&gt; / CDN Demo</h1>
    <p>This page uses the browser build from <code>unpkg.com/gridjs-spreadsheet</code>; npm is not required.</p>

    <!-- Container for the GridJs UI -->
    <div id="gridjs-demo-uid"></div>
</body>
</html>
```

