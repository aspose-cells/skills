# 11‑troubleshooting.md  

| Symptom | Observable behaviour / message | Likely cause | Quick check | Exception raised |
|---------|--------------------------------|--------------|-------------|-------------------|
| Workbook not found | `FileNotFoundError: [Errno 2] No such file or directory: 'static/files/…'` returned from the **LoadSpreadsheet** endpoint. | The route builds the workbook path from `WORKBOOK_DIR` + `filename` but the file does not exist at that location (e.g., typo in name, missing `Sample.xlsx`). | Verify the file exists on disk: `os.path.isfile(os.path.join("static","files","Sample.xlsx"))`. | `FileNotFoundError` |
| Cache directory cannot be written | Server logs show `PermissionError: [Errno 13] Permission denied: './cache/…'` and the request fails with 500. | `options.file_cache_directory` points to a folder that the process cannot create or write to (missing folder, read‑only mount, wrong user). | Check write permission: `os.access("./cache", os.W_OK)`. | `PermissionError` |
| Gzip stream read before it is rewound | The client receives an empty response body or the server raises `ValueError: I/O operation on closed file`. | After writing gzipped JSON to a `BytesIO` the stream pointer is left at the end; reading without `seek(0)` returns no data or raises an error. | Print the current position: `print(gzip_buffer.tell())` – it should be `0` before reading. | `ValueError` (or generic `OSError` for a closed stream) |

---  

## 1. Workbook not found  

**Why it happens**  
`load_spreadsheet` builds the absolute path with  

```python
workbook_path = os.path.join(WORKBOOK_DIR, filename)
```  

If `filename` does not match a file in `static/files/`, the call  

```python
gridjs_service.detail_stream_json_with_uid(gz, workbook_path, uid)
```  

throws `FileNotFoundError`.

**How to confirm**  

```python
import os
print(os.path.isfile(os.path.join("static", "files", "Sample.xlsx")))
```  

`True` means the file is present; `False` confirms the missing‑file hypothesis.

**Fix**  

1. Ensure the demo workbook is exactly `static/files/Sample.xlsx`.  
2. If you use a custom name, pass it via the query string:  

   ```
   GET /GridJs/LoadSpreadsheet?uid=123&filename=MyBook.xlsx
   ```  

3. Restart the server after adding the file.

**Robust code example**

```python
from flask import Flask, request, Response
from aspose.cellsgridjs import GridJsService, GridJsOptions
import io, gzip, os

app = Flask(__name__)

options = GridJsOptions()
options.file_cache_directory = "./cache"
options.base_route_name = "/GridJs"
os.makedirs(options.file_cache_directory, exist_ok=True)

WORKBOOK_DIR = os.path.join("static", "files")
gridjs_service = GridJsService(options)

@app.route("/GridJs/LoadSpreadsheet", methods=["GET"])
def load_spreadsheet():
    uid = request.args.get("uid", "")
    filename = request.args.get("filename", "Sample.xlsx")
    workbook_path = os.path.join(WORKBOOK_DIR, filename)

    gzip_buffer = io.BytesIO()
    try:
        with gzip.GzipFile(fileobj=gzip_buffer, mode="w") as gz:
            gridjs_service.detail_stream_json_with_uid(gz, workbook_path, uid)
    except FileNotFoundError as e:
        # Return a clear 404 to the client
        return Response(str(e), status=404, mimetype="text/plain")

    gzip_buffer.seek(0)  # <-- important for the next step
    return Response(
        gzip_buffer.getvalue(),
        mimetype="application/json",
        headers={"Content-Encoding": "gzip"},
    )
```

---

## 2. Cache directory cannot be written  

**Why it happens**  
`GridJsOptions.file_cache_directory` is used by `GridJsService` to store per‑UID copies of the workbook. If the directory does not exist or the process lacks write permission, any attempt to create a cache file raises `PermissionError`.

**How to confirm**  

```python
import os
print(os.path.isdir("./cache"))          # should be True
print(os.access("./cache", os.W_OK))     # should be True
```

If either prints `False`, the directory is the problem.

**Fix**  

1. Create the folder before initializing the service and give the process write rights:

   ```bash
   mkdir -p ./cache
   chmod 775 ./cache   # adjust for your environment
   ```

2. In code, ensure `os.makedirs(options.file_cache_directory, exist_ok=True)` runs **before** `GridJsService(options)`.

**Robust code example**

```python
import os
from aspose.cellsgridjs import GridJsOptions, GridJsService

cache_dir = "./cache"
try:
    os.makedirs(cache_dir, exist_ok=True)
    if not os.access(cache_dir, os.W_OK):
        raise PermissionError(f"Cache directory '{cache_dir}' is not writable")
except OSError as e:
    # Fail fast with a clear message
    raise PermissionError(f"Unable to prepare cache directory: {e}") from e

options = GridJsOptions()
options.file_cache_directory = cache_dir
options.base_route_name = "/GridJs"

gridjs_service = GridJsService(options)
```

---

## 3. Gzip stream read before it is rewound  

**Why it happens**  
`detail_stream_json_with_uid` writes gzipped JSON into a `BytesIO` object. After the `with gzip.GzipFile(...):` block the stream pointer is at the end. If you call `gzip_buffer.getvalue()` (or read) without resetting the pointer, the response is empty or the runtime may raise `ValueError` because the underlying stream is considered “closed for reading”.

**How to confirm**  

```python
print(gzip_buffer.tell())   # should be 0 before reading
```

If the value is non‑zero, the stream has not been rewound.

**Fix**  

Always call `gzip_buffer.seek(0)` **after** writing and **before** reading or returning the data.

**Robust code example**

```python
import io, gzip
from aspose.cellsgridjs import GridJsService

def get_gzipped_json(service: GridJsService, path: str, uid: str) -> bytes:
    buffer = io.BytesIO()
    try:
        with gzip.GzipFile(fileobj=buffer, mode="w") as gz:
            service.detail_stream_json_with_uid(gz, path, uid)
        # Rewind before extracting bytes
        buffer.seek(0)
        return buffer.getvalue()
    except ValueError as e:
        # This catches the “stream not rewound” scenario
        raise ValueError(f"Gzip stream was not rewound: {e}") from e
```

Use the helper in the Flask endpoint:

```python
@app.route("/GridJs/LoadSpreadsheet", methods=["GET"])
def load_spreadsheet():
    uid = request.args.get("uid", "")
    filename = request.args.get("filename", "Sample.xlsx")
    workbook_path = os.path.join(WORKBOOK_DIR, filename)

    try:
        data = get_gzipped_json(gridjs_service, workbook_path, uid)
    except FileNotFoundError as e:
        return Response(str(e), status=404, mimetype="text/plain")
    except PermissionError as e:
        return Response(str(e), status=500, mimetype="text/plain")
    except ValueError as e:
        return Response(str(e), status=500, mimetype="text/plain")

    return Response(
        data,
        mimetype="application/json",
        headers={"Content-Encoding": "gzip"},
    )
```

---  

**Summary checklist**

| Issue | Verify | Fix |
|-------|--------|-----|
| Workbook missing | `os.path.isfile(workbook_path)` → `True` | Place the correct file in `static/files/` or adjust the query string. |
| Cache folder not writable | `os.path.isdir("./cache")` & `os.access("./cache", os.W_OK)` → both `True` | Create the folder and grant write permission before `GridJsService` construction. |
| Gzip stream not rewound | `gzip_buffer.tell()` returns `0` before read | Call `gzip_buffer.seek(0)` after writing, before returning data. |

Applying these steps eliminates the three most common runtime failures documented for Aspose.Cells.GridJs in Python.
