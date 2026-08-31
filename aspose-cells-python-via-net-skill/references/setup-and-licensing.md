# Setup and Licensing

Read this when installing Aspose.Cells for Python via .NET, applying or verifying a license, or output shows an evaluation watermark ("Evaluation Only. Created with Aspose.Cells") or truncated content.

A license is a digitally signed XML file. Set it ONCE per process, at startup, BEFORE constructing any `Workbook`. The license is NOT shipped in the wheel - you get it by purchasing or by requesting a free 30-day temporary license (https://purchase.aspose.com/temporary-license).

## Pitfalls

### Evaluation watermark though you own a license

Symptom: output contains an extra worksheet titled "Evaluation Only. Created with Aspose.Cells" (or a watermark banner on PDF/image output, or truncated CSV), even though you have a valid `.lic` file. `Workbook.is_licensed` returns `False`. Four distinct causes:

a. SetLicense not called, called too late, or called on the wrong class. Apply the license before the first `Workbook` is constructed, exactly once, using `gc.License` - NOT another product's License class. Each Aspose product has its own License class and its own `.lic`.

```python
import aspose.cells as gc

# Call ONCE at process startup, BEFORE constructing any Workbook.
gc.License().set_license("Aspose.Cells.lic")  # file name only; resolved next to the wheel or as an embedded resource
workbook = gc.Workbook()
workbook.save("licensed.xlsx")
```

b. License file not found at runtime (works on the dev box, watermark on the server or in a container). See "License file not found at runtime" below.

c. Subscription-expiry trap ("worked yesterday, watermark after an upgrade"). A license unlocks any build released BEFORE your subscription end date, forever - but it does NOT unlock a build released AFTER that date. Upgrading the `aspose-cells-python` wheel to a newer version silently reverts to evaluation mode. Fix: pin the wheel to a version released within your subscription window, or renew. The `.lic` carries the expiry date; the wheel's release date must fall on or before it. (Every license includes a one-year upgrade subscription.)

d. Edited or re-saved `.lic` file. The file is digitally signed. Any change - adding a line break, re-encoding, opening and re-saving in an editor - invalidates the signature. Use it byte-for-byte as delivered; never modify or reformat it.

### License file not found at runtime

Symptom: passing just the file name works locally but produces a watermark after deployment, or `set_license` throws.

Cause: the `.lic` is not on any path Aspose.Cells searches. With the file-name overload, Aspose.Cells looks in this order:
1. The explicit path you pass.
2. The folder containing the `aspose` package / executable.
3. The current working directory.

Fix (recommended for servers/containers): resolve an absolute path to the `.lic`, or load it from bytes (blob store, secret manager):

```python
import aspose.cells as gc

gc.License().set_license("Aspose.Cells.lic")
```

Loading from a stream also works when you control the bytes:

```python
import aspose.cells as gc

gc.License().set_license("Aspose.Cells.lic")
```

## Verifying the license loaded

`Workbook.is_licensed` is `False` before a license is applied and `True` after a successful `set_license` - assert it in a smoke test. A missing license file makes `set_license` throw `FileNotFoundError` ("Cannot find license"); a corrupt or modified file throws `IOError` ("Corrupted License file.") - verified against 26.7.0. Do not swallow either.

```python
import aspose.cells as gc

try:
    gc.License().set_license("Aspose.Cells.lic")
except Exception as ex:  # FileNotFoundError if missing, OSError if corrupted
    print("License NOT applied:", ex)
    raise  # fail fast - never continue silently into evaluation mode

wb = gc.Workbook()
print(wb.is_licensed)  # prints True when the license loaded
```

## Evaluation version limitations

Recognize these as "no license loaded" symptoms, not bugs. Without a license Aspose.Cells is fully functional except:
- An extra worksheet stamped "Evaluation Only. Created with Aspose.Cells" is injected and forced to be the active sheet.
- Only 100 files may be opened per process; opening more throws an exception.
- Plain-text output (CSV, TSV) appends evaluation text and writes only the FIRST worksheet.
- PDF and image output get an evaluation watermark banner at the top.

All disappear once a valid license is applied before the first `Workbook`.

## Metered (pay-as-you-go) licensing

An alternative to a `.lic` file: bill by usage instead of a fixed license. Call `set_metered_key(public_key, private_key)` on a `Metered` instance at startup. It requires network access (consumption is reported to Aspose) and meters by processed file size.

```python
import aspose.cells as gc

metered = gc.Metered()
metered.set_metered_key("<public key>", "<private key>")
workbook = gc.Workbook()
print(workbook.is_licensed)
before = gc.Metered.get_consumption_quantity()
workbook.save("out.xlsx")
after = gc.Metered.get_consumption_quantity()
print(after - before)  # amount consumed by this run
```

## Installing Aspose.Cells for Python via .NET

The wheel is the only artifact you install; the license is separate.

```
pip install aspose-cells-python
```

Microsoft Excel is NOT required on the machine. The current package (25.x) ships a
self-contained build for Windows, Linux, and macOS (x64 and arm64). The `.lic` comes
from the purchase portal or the 30-day temporary-license portal - it is never part of
the wheel download. Pin the wheel version to stay within your subscription window
(see cause c). For first workbook operations after setup, follow the SKILL.md quickstart.

## Related
- cross-platform-deployment.md - Linux/Docker font and SkiaSharp setup for non-Windows PDF/image output.
