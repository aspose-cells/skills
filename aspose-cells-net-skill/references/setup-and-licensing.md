# Setup and Licensing

Read this when installing Aspose.Cells, applying or verifying a license, or output shows an evaluation watermark ("Evaluation Only. Created with Aspose.Cells") or truncated content.

A license is a digitally signed XML file. Set it ONCE per process, at startup, BEFORE constructing any `Workbook`. The license is NOT shipped in the NuGet package - you get it by purchasing or by requesting a free 30-day temporary license (https://purchase.aspose.com/temporary-license).

## Pitfalls

### Evaluation watermark though you own a license

Symptom: output contains an extra worksheet titled "Evaluation Only. Created with Aspose.Cells" (or a watermark banner on PDF/image output, or truncated CSV), even though you have a valid `.lic` file. `Workbook.IsLicensed` returns `false`. Four distinct causes:

a. SetLicense not called, called too late, or called on the wrong class. Apply the license before the first `Workbook` is constructed, exactly once, using `Aspose.Cells.License` - NOT another product's License class (`Aspose.Words.License`, `Aspose.Pdf.License`, ...). Each Aspose product has its own License class and its own `.lic`.

```csharp
// Call ONCE at process startup, BEFORE constructing any Workbook.
Aspose.Cells.License license = new Aspose.Cells.License();
license.SetLicense("Aspose.Cells.lic"); // file name only; resolved next to the DLL/exe or as an embedded resource
// Every Workbook created after this line is fully licensed.
Workbook workbook = new Workbook();
workbook.Save("licensed.xlsx");
```

b. License file not found at runtime (works on the dev box, watermark on the server or in a container). See "License file not found at runtime" below.

c. Subscription-expiry trap ("worked yesterday, watermark after a package upgrade"). A license unlocks any build released BEFORE your subscription end date, forever - but it does NOT unlock a build released AFTER that date. Upgrading the Aspose.Cells NuGet package to a newer DLL silently reverts to evaluation mode. Fix: pin the package to a version released within your subscription window, or renew. The `.lic` carries the expiry date; the DLL's release date must fall on or before it. (Every license includes a one-year upgrade subscription.)

d. Edited or re-saved `.lic` file. The file is digitally signed. Any change - adding a line break, re-encoding, opening and re-saving in an editor - invalidates the signature. Use it byte-for-byte as delivered; never modify or reformat it.

### License file not found at runtime

Symptom: passing just the file name works locally but produces a watermark after deployment, or `SetLicense` throws.

Cause: the `.lic` is not on any path Aspose.Cells searches. With the file-name overload, Aspose.Cells looks in this order:
1. The explicit path you pass.
2. The folder containing Aspose.Cells.dll.
3. The folder containing the assembly that called Aspose.Cells.dll.
4. The folder containing the entry assembly (your .exe).
5. An embedded resource in the assembly that called Aspose.Cells.dll.

Fix (recommended for servers/containers): embed the license so it can never be lost or mispathed. Add the `.lic` to the project, set Build Action = Embedded Resource, then pass just the file name - no reflection needed; the License class finds it in the manifest resources automatically.

```csharp
// .lic added to the project with Build Action = Embedded Resource
Aspose.Cells.License license = new Aspose.Cells.License();
license.SetLicense("Aspose.Cells.lic"); // must match the embedded resource's file name exactly
```

The string you pass must match the license file name exactly. If you renamed it to "Aspose.Cells.lic.xml", pass that.

Loading from a stream also works when you control the bytes (blob store, secret manager):

```csharp
Aspose.Cells.License license = new Aspose.Cells.License();
using (FileStream fs = new FileStream("Aspose.Cells.lic", FileMode.Open, FileAccess.Read))
{
    license.SetLicense(fs);
}
```

## Verifying the license loaded

`Workbook.IsLicensed` is `false` before a license is applied and `true` after a successful `SetLicense` - assert it in a smoke test. A missing license file makes `SetLicense` throw `FileNotFoundException` ("Cannot find license"); a corrupt or modified file throws `IOException` ("Corrupted License file.") - verified against 26.7.0. Do not swallow either.

```csharp
Aspose.Cells.License license = new Aspose.Cells.License();
try
{
    license.SetLicense("Aspose.Cells.lic");
}
catch (Exception ex) // FileNotFoundException if missing, IOException if corrupted
{
    Console.WriteLine("License NOT applied: " + ex.Message);
    throw; // fail fast - never continue silently into evaluation mode
}
Workbook wb = new Workbook();
Console.WriteLine(wb.IsLicensed); // prints True when the license loaded
```

## Evaluation version limitations

Recognize these as "no license loaded" symptoms, not bugs. Without a license Aspose.Cells is fully functional except:
- An extra worksheet stamped "Evaluation Only. Created with Aspose.Cells" is injected and forced to be the active sheet.
- Only 100 files may be opened per process; opening more throws an exception.
- Plain-text output (CSV, TSV) appends evaluation text and writes only the FIRST worksheet.
- PDF and image output get an evaluation watermark banner at the top.

All disappear once a valid license is applied before the first `Workbook`.

## Metered (pay-as-you-go) licensing

An alternative to a `.lic` file: bill by usage instead of a fixed license. Call `SetMeteredKey(publicKey, privateKey)` on a `Metered` instance at startup. It requires network access (consumption is reported to Aspose) and meters by processed file size.

```csharp
Metered metered = new Metered();
metered.SetMeteredKey("<public key>", "<private key>");
Workbook workbook = new Workbook();
Console.WriteLine(workbook.IsLicensed);
decimal before = Metered.GetConsumptionQuantity();
workbook.Save("out.xlsx");
decimal after = Metered.GetConsumptionQuantity();
Console.WriteLine(after - before); // amount consumed by this run
```

## Installing Aspose.Cells

NuGet is the only artifact you install; the license is separate.

- .NET CLI: `dotnet add package Aspose.Cells`
- Package Manager Console: `Install-Package Aspose.Cells`

Microsoft Excel is NOT required on the machine. The current package (26.x) ships builds for .NET Framework 4.0 and 4.8, .NET Standard 2.0, and .NET 6/8/9/10, each with an additional Windows-specific variant. The `.lic` comes from the purchase portal or the 30-day temporary-license portal - it is never part of the NuGet download. Pin the package version to stay within your subscription window (see cause c). For first workbook operations after setup, follow the SKILL.md quickstart.

## Related
- cross-platform-deployment.md - Linux/Docker/.NET 6+ setup, SkiaSharp, and fonts for non-Windows PDF/image output.
