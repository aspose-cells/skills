# Cross-Platform Deployment (Linux, Docker, macOS, Python anywhere)

Read this when code works on Windows but crashes, hangs, or renders wrong or missing fonts on Linux, Docker, macOS, AWS Lambda, or any non-Windows target - including Python deployed as a wheel, a Docker image, or a server/cloud function.

Two failure modes dominate: (1) the graphics backend is missing or version-mismatched, and (2) the fonts the workbook needs are not installed. Fix both.

## Pitfalls

### Crash on first render or save on Linux

Symptom: `TypeInitializationException`, `PlatformNotSupportedException`, or a native load error - thrown the first time you save to PDF/image, render a chart, or measure text. The same code runs fine on Windows.

Cause: Aspose.Cells for Python via .NET migrated non-Windows rendering to **SkiaSharp** starting **v22.10.1**. The Python wheel bundles the SkiaSharp native library, so no extra NuGet step is needed - but the native library still must be loadable in the target environment, and fonts must be present. On very old or exotic images a missing system library (e.g. `libfontconfig1`) can still break loading.

Fix:
- Use Aspose.Cells for Python via .NET **>= 22.10.1** (25.x recommended).
- Ensure the basic system libraries are present: on Debian/Ubuntu install `libfontconfig1` (Alpine: `fontconfig`); on minimal/musl images you may also need `libstdc++` and `libgcc`.
- Install the actual `.ttf`/`.otf` fonts the workbook references (see below).

### Blank boxes, wrong metrics, or substituted fonts in output

Symptom: PDF or image output shows empty rectangles, clipped or overlapping text, wrong column widths / row heights, or a different typeface than Excel. Layout and numbers drift because glyph metrics differ from the intended font.

Cause: base OS and container images ship almost no fonts. When the workbook's font (Calibri, Segoe UI, ...) is absent, Aspose.Cells substitutes the closest available font, changing metrics and appearance.

Fix: install the actual `.ttf`/`.otf` files into the image, and - if they are not in a system font directory - register them. Every `FontConfigs` call must run **before** you construct or load any `Workbook`, and if several source-setting methods are called, only the last one takes effect.

```python
import aspose.cells as gc

# Call FontConfigs BEFORE constructing or loading any Workbook.
gc.FontConfigs.set_font_folder("/usr/share/fonts", True)   # True = search subfolders
wb = gc.Workbook("report.xlsx")
wb.save("report.pdf", gc.SaveFormat.PDF)
```

Use `set_font_folders` for several directories, and `set_font_substitutes` to redirect a font you cannot ship to one you can:

```python
import aspose.cells as gc

gc.FontConfigs.set_font_folders(["/usr/share/fonts", "/app/fonts"], True)
# Redirect absent Windows fonts to metric-compatible fonts the image ships.
gc.FontConfigs.set_font_substitutes("Calibri", ["Carlito", "DejaVu Sans"])
gc.FontConfigs.set_font_substitutes("Segoe UI", ["DejaVu Sans"])
```

`FontConfigs.set_font_sources` additionally loads a single font file or an in-memory `bytes`. The render-time last-resort default (`PdfSaveOptions.default_font`, `ImageOrPrintOptions.default_font`, `FontConfigs.default_font_name`) is a separate lever - see conversion-and-rendering.md. Installing the real fonts is always the more faithful fix.

### No printing to a physical printer on SkiaSharp builds

Symptom: printer output APIs are unavailable or fail off Windows.

Cause: the SkiaSharp backend does not implement printing to a physical printer; that capability exists only on the Windows build.

Fix: on Linux/macOS/containers, render to PDF or image (`save(..., SaveFormat.PDF)`, or `SheetRender`/`WorkbookRender` to PNG) and print that downstream. Do not design a server flow around driving a printer directly.

### EMF and TIFF not supported on Linux

Symptom: saving or rendering to EMF or TIFF fails or produces empty output on Linux.

Cause: Aspose.Cells for Python via .NET does not support EMF and TIFF on Linux.

Fix: target PNG, JPEG, BMP, or PDF for image output on Linux. Keep any EMF/TIFF paths on Windows only.

## Docker (Python wheel, SkiaSharp, Debian)

Minimal known-good multi-stage Dockerfile for a Python service that uses the wheel.

```dockerfile
FROM python:3.11-slim AS build
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

FROM python:3.11-slim AS final
WORKDIR /app
# fontconfig is needed for PDF/image text rendering
RUN apt-get update && apt-get install -y --no-install-recommends libfontconfig1 fontconfig \
    && rm -rf /var/lib/apt/lists/*
# Base images ship almost no fonts: copy your .ttf/.otf files into a system font dir.
COPY fonts/ /usr/share/fonts/
COPY --from=build /app /app
ENTRYPOINT ["python", "app.py"]
```

`requirements.txt` should pin the wheel, e.g. `aspose-cells-python==26.7.0`. Alpine delta: use `python:3.11-alpine` and replace the apt line with `RUN apk add --no-cache fontconfig`. As an apt alternative to copying `.ttf` files, install a font package (e.g. `fonts-liberation`).

## macOS

The Python wheel carries the SkiaSharp native library for macOS, so no extra native step is normally required. Install the real fonts (or use `FontConfigs.set_font_folder`) before rendering, exactly as on Linux.

## AWS Lambda / serverless Linux

Lambda runs on Amazon Linux, so every non-Windows rule applies: Aspose.Cells for Python via .NET with the bundled SkiaSharp backend, plus fonts you must ship. You cannot `apt-get` into a managed runtime, so:
- Bundle the `.ttf`/`.otf` fonts inside the deployment package (or a Lambda layer).
- Register them at cold start with `FontConfigs.set_font_folder("./fonts", True)` before the first `Workbook` - the same pattern used for read-only cloud servers where you cannot modify system directories.
- Write only under `/tmp` (the sole writable path); prefer saving to a `BytesIO` and returning the bytes.

## Related
- conversion-and-rendering.md - render-time default_font in PdfSaveOptions/ImageOrPrintOptions, PDF layout options, SheetRender/WorkbookRender image output
- setup-and-licensing.md - pip install, wheel variants, license setup
