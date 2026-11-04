# Getting Started with Aspose.Cells GridJs (Python)

This guide shows the absolute minimum you need to add GridJs to an existing Python web project, verify that it builds, runs, and renders the demo workbook **Sample.xlsx**.

---

## 1️⃣ Prerequisites

| Item | Requirement |
|------|-------------|
| Python | **3.9+** |
| Web framework | **Flask** (or FastAPI – the example uses Flask) |
| Package | `aspose-cells-gridjs-net-python` – bundles the Aspose.Cells runtime |
| OS permissions | Write permission for a folder that will hold per‑session cache files |

> **Note** – The demo workbook must be placed at `static/files/Sample.xlsx`. Do **not** move it to the cache folder; the cache is only for temporary copies created at runtime.

---

## 2️⃣ Install the dependencies

```bash
# Create a virtual environment (optional but recommended)
python -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate

# Install Flask and the GridJs package
pip install flask aspose-cells-gridjs-net-python
```

If you prefer FastAPI, replace `flask` with `fastapi uvicorn` in the command above.

---

## 3️⃣ Minimal Flask application

Create a file called **app.py** in the root of your project and paste the code below **exactly**.  
All class and method names are taken from the reference material; no extra imports are required.

```python
# app.py
# -------------------------------------------------------------
# Minimal Flask app that hosts Aspose.Cells GridJs services.
# -------------------------------------------------------------
from flask import Flask, request, Response, jsonify, send_file, abort
import os, io, gzip, mimetypes

# GridJs imports – these are the only public classes needed for a working demo
from aspose.cellsgridjs import GridJsOptions, GridJsService

# -----------------------------------------------------------------
# 1️⃣ Application and folder configuration
# -----------------------------------------------------------------
BASE_DIR = os.path.abspath(os.path.dirname(__file__))
app = Flask(__name__)

# Where the original workbooks live (do NOT change unless you move Sample.xlsx)
WORKBOOK_DIR = os.path.join(BASE_DIR, "static", "files")

# Cache folder – GridJs writes a per‑UID copy here. Must exist and be writable.
CACHE_DIR = os.path.join(BASE_DIR, "cache")
os.makedirs(CACHE_DIR, exist_ok=True)

# -----------------------------------------------------------------
# 2️⃣ Initialise GridJs service
# -----------------------------------------------------------------
options = GridJsOptions()
options.file_cache_directory = CACHE_DIR          # required
options.base_route_name = "/GridJs"               # client will call /GridJs/…
gridjs_service = GridJsService(options)

# -----------------------------------------------------------------
# 3️⃣ Helper – MIME type for file download endpoints
# -----------------------------------------------------------------
def _mime_type(filename: str) -> str:
    mime, _ = mimetypes.guess_type(filename)
    return mime or "application/octet-stream"

# -----------------------------------------------------------------
# 4️⃣ Routes required for the UI to work
# -----------------------------------------------------------------
@app.route("/")
def index():
    """Render a very small HTML page that loads the GridJs UI."""
    return """
    <!doctype html>
    <html lang="en">
    <head>
      <meta charset="UTF-8">
      <title>GridJs demo</title>
      <script type="module" src="/static/js/demo.js"></script>
    </head>
    <body>
      <div id="grid"></div>
    </body>
    </html>
    """

# -------------------------------------------------------------
# LoadSpreadsheet – returns gzipped JSON for a workbook
# -------------------------------------------------------------
@app.route("/GridJs/LoadSpreadsheet", methods=["GET"])
def load_spreadsheet():
    uid = request.args.get("uid", "")
    filename = request.args.get("filename", "Sample.xlsx")
    if not uid:
        return jsonify({"error": "uid is required"}), 400
    if filename != os.path.basename(filename):
        return jsonify({"error": "invalid filename"}), 400

    workbook_path = os.path.join(WORKBOOK_DIR, filename)

    # GridJs writes JSON directly into the provided stream.
    gzip_buffer = io.BytesIO()
    with gzip.GzipFile(fileobj=gzip_buffer, mode="w") as gz:
        gridjs_service.detail_stream_json_with_uid(gz, workbook_path, uid)
    gzip_buffer.seek(0)

    return Response(
        gzip_buffer.getvalue(),
        mimetype="application/json",
        headers={"Content-Encoding": "gzip"},
    )

# -------------------------------------------------------------
# UpdateCell – receives a cell change from the client
# -------------------------------------------------------------
@app.route("/GridJs/UpdateCell", methods=["POST"])
def update_cell():
    p = request.form.get("p")
    uid = request.form.get("uid")
    result = gridjs_service.update_cell(p, uid)
    return Response(result, content_type="text/plain; charset=utf-8")

# -----------------------------------------------------------------
# Optional: serve the tiny client script used by the HTML above
# -----------------------------------------------------------------
@app.route("/static/js/demo.js")
def demo_js():
    """A minimal client that loads the workbook and creates the spreadsheet."""
    return """
    import Spreadsheet from 'gridjs-spreadsheet';
    const host = document.getElementById('grid');

    function createUid() {
        return crypto.randomUUID?.() || `uid-${Date.now()}`;
    }

    async function load() {
        const uid = createUid();
        const resp = await fetch(`/GridJs/LoadSpreadsheet?uid=${uid}`);
        if (!resp.ok) throw new Error('Load failed');
        const payload = JSON.parse(await resp.text());

        const sheet = new Spreadsheet(host, {
            updateMode: 'server',
            updateUrl: '/GridJs/UpdateCell',
            mode: 'edit',
            local: 'en',
            showToolbar: true,
            showContextmenu: true,
        });
        sheet.loadData(payload.data || [], payload.actname);
        sheet.setUniqueId?.(payload.uniqueid || '');
        sheet.setFileName?.(payload.filename || 'Sample.xlsx');
    }

    load().catch(e => console.error(e));
    """, 200, {"Content-Type": "application/javascript"}

# -----------------------------------------------------------------
# Run the server
# -----------------------------------------------------------------
if __name__ == "__main__":
    # Port 2022 matches the demo documentation, but any free port works.
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "2022")), debug=False)
```

### What the code does

| Step | Explanation |
|------|-------------|
| **Folder setup** | `static/files` holds the source workbook (`Sample.xlsx`). `cache` holds per‑UID copies created at runtime. |
| **`GridJsOptions`** | `file_cache_directory` **must** be a writable folder. `base_route_name` tells the client where the service lives (`/GridJs`). |
| **`detail_stream_json_with_uid`** | The only documented method that streams gzipped JSON for a workbook. It writes directly to the supplied `gzip.GzipFile`. |
| **`update_cell`** | Handles cell edits sent by the client. Returns plain‑text response as required by the UI. |
| **Minimal client script** | Uses the npm‑style `gridjs-spreadsheet` entry point (`Spreadsheet`). It creates a random `uid`, fetches the JSON, and renders the sheet. No build step is needed for this demo. |

---

## 4️⃣ Build & Run the demo

```bash
# 1️⃣ Ensure the cache folder exists (the script also creates it, but run it once)
mkdir -p cache

# 2️⃣ Start the Flask server
python app.py
```

You should see something like:

```
 * Running on http://0.0.0.0:2022/ (Press CTRL+C to quit)
```

Open a browser and navigate to **http://localhost:2022/**. The page loads a tiny HTML skeleton, the client script fetches `/GridJs/LoadSpreadsheet`, and the spreadsheet UI appears with the contents of **Sample.xlsx**.

---

## 5️⃣ Verification checklist

When the grid appears **empty** (no rows/columns), run through the three most common causes:

| # | What to check | How to verify |
|---|---------------|---------------|
| 1️⃣ | **`uid` query parameter** is present and non‑empty | Look at the network request to `/GridJs/LoadSpreadsheet` in the browser dev tools. The URL must contain `uid=…`. If missing, the server returns `400`. |
| 2️⃣ | **Cache directory** exists and is writable | On the server console, confirm no `PermissionError` was logged. Manually create the folder (`mkdir -p cache`) and ensure the user running Flask can write files inside it. |
| 3️⃣ | **Workbook path** resolves to a real file | Verify that `static/files/Sample.xlsx` exists (case‑sensitive on Linux). If the file is missing, the server raises `FileNotFoundError` and the client receives a 500 error. |

If all three items are correct and the grid is still empty, open the browser console – any error printed by `Spreadsheet` (e.g., JSON parsing) will point to the next debugging step.

---

### 🎉 You’re ready

You now have the smallest possible GridJs integration:

1. Install the package.  
2. Add the minimal Flask app.  
3. Run it, open the browser, and see **Sample.xlsx** rendered.

From here you can expand the demo by adding more routes (image handling, PDF export, etc.) or swapping Flask for FastAPI – the core `GridJsOptions` and `GridJsService` usage stays the same. Happy coding!
