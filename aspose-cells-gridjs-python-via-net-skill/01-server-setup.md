# Server-Side Setup Guide

## Prerequisites

- Python 3.9+ (for Python)
- Web framework: Flask or FastAPI
- pip package: aspose-cells-gridjs-net-python (it bundles Aspose.Cells, so `aspose-cells` is not needed separately)

## Step 1: Install Dependencies

```bash
pip install flask aspose-cells-gridjs-net-python
# or for FastAPI:
pip install fastapi uvicorn aspose-cells-gridjs-net-python
```

## Step 2: Configure GridJs Options

```python
# Initialize options
from aspose.cellsgridjs import GridJsOptions

options = GridJsOptions()
options.file_cache_directory = "./cache"  # Must exist and be writable
options.base_route_name = "/GridJs"
```

## Step 3: Create Flask Application

```python
# app.py (Flask example)
from flask import Flask, request, Response, jsonify
from aspose.cellsgridjs import GridJsOptions, GridJsService
import io, gzip, os

app = Flask(__name__)

# Initialize GridJs service
options = GridJsOptions()
# Session/cache directory - GridJs writes per-uid copies here. The workbook
# the demo loads is NOT read from here; it is application content in static/files/.
options.file_cache_directory = "./cache"  # Must exist and be writable
options.base_route_name = "/GridJs"
os.makedirs(options.file_cache_directory, exist_ok=True)

# Where source workbooks are resolved from. The demo ships Sample.xlsx here.
WORKBOOK_DIR = os.path.join("static", "files")

gridjs_service = GridJsService(options)

@app.route('/GridJs/LoadSpreadsheet', methods=['GET'])
def load_spreadsheet():
    uid = request.args.get('uid', '')
    filename = request.args.get('filename', 'Sample.xlsx')
    workbook_path = os.path.join(WORKBOOK_DIR, filename)
    
    # Use detail_stream_json_with_uid to write gzipped JSON
    gzip_buffer = io.BytesIO()
    with gzip.GzipFile(fileobj=gzip_buffer, mode='w') as gz:
        gridjs_service.detail_stream_json_with_uid(gz, workbook_path, uid)
    gzip_buffer.seek(0)
    
    return Response(
        gzip_buffer.getvalue(),
        mimetype='application/json',
        headers={'Content-Encoding': 'gzip'}
    )

@app.route('/GridJs/UpdateCell', methods=['POST'])
def update_cell():
    p = request.form.get('p')
    uid = request.form.get('uid')
    ret = gridjs_service.update_cell(p, uid)
    return Response(ret, content_type='text/plain; charset=utf-8')

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)), debug=False)
```

## Important Notes

- **Key method**: Use `detail_stream_json_with_uid(stream, file_path, uid)` for LoadSpreadsheet endpoint - NOT `detail_file_json_with_uid`
- The method writes gzipped JSON data to the provided stream (use io.BytesIO + gzip.GzipFile)
- GridJsService provides methods: update_cell, add_image, copy_image, image, image_url, ole, get_file, download
- The `uid` (unique identifier) is used for file caching - generate a UUID per session
- `file_cache_directory` must exist and be writable before initializing the service
- Source workbooks are read from `static/files/` (`static/files/Sample.xlsx`).
  `file_cache_directory` holds only the per-`uid` copies GridJs wrote during
  earlier requests, so joining the requested filename onto it asks for a file that
  does not exist on a cold server and the route raises `FileNotFoundError`
- For production, use a WSGI server like gunicorn: `gunicorn app:app`
