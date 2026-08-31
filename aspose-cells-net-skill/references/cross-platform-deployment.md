# Cross-Platform Deployment (Linux, Docker, macOS, .NET 6+)

Read this when code works on Windows but crashes, hangs, or renders wrong or missing fonts on Linux, Docker, macOS, AWS Lambda, Blazor, or any .NET 6+ non-Windows target.

Two failure modes dominate: (1) the graphics backend is missing or version-mismatched, and (2) the fonts the workbook needs are not installed. Fix both.

## Pitfalls

### Crash on first render or save on Linux (.NET 6+)

Symptom: `TypeInitializationException`, `PlatformNotSupportedException` "System.Drawing.Common is not supported on this platform", or `DllNotFoundException` for `libgdiplus` - thrown the first time you save to PDF/image, render a chart, or measure text. The same code runs fine on Windows.

Cause: .NET 6 made `System.Drawing.Common` Windows-only. Aspose.Cells migrated non-Windows rendering to **SkiaSharp** starting **v22.10.1**. On .NET 6+, the backend is chosen from the project's **Target OS**: Target OS = Windows -> `System.Drawing.Common`; Target OS = None or any other value -> SkiaSharp. A non-Windows deployment therefore needs the SkiaSharp native library present, and its version must match the SkiaSharp version Aspose.Cells references.

Fix:
- Use Aspose.Cells **>= 22.10.1** (26.x recommended).
- Add the OS-specific SkiaSharp native-assets NuGet package: **SkiaSharp.NativeAssets.Linux**, or **SkiaSharp.NativeAssets.Linux.NoDependencies** (the NoDependencies variant bundles what it needs, so you can skip the apt/apk step below).
- **Pin its version to the SkiaSharp version your Aspose.Cells references** (table below). A version skew between the SkiaSharp that Aspose.Cells pulls in and SkiaSharp.NativeAssets.* is itself a load-time crash.
- Install the native config package: `libfontconfig1` (Debian/Ubuntu) or `fontconfig` (Alpine). Not needed with the `.NoDependencies` variant.
- Legacy path only - .NET Core 3.1 or earlier, or an old Aspose.Cells still on the System.Drawing.Common backend: install `libgdiplus` + `libc6-dev` instead of the SkiaSharp assets.

SkiaSharp.NativeAssets.* version must equal the SkiaSharp column:

| Aspose.Cells for .NET | SkiaSharp |
|:---:|:---:|
| >= 22.10.1 && <= 22.11 | 2.88.0 |
| >= 22.12 && <= 23.9 | 2.88.3 |
| >= 23.10 && <= 24.12 | 2.88.6 |
| = 25.1.1 | 3.116.1 |
| >= 25.1.2 && <= 26.4 | 2.88.9 (.NET 6.0, .NET 8.0), 3.116.1 (.NET 9.0) |
| >= 26.5.1 | 2.88.9 (.NET 6.0, .NET 8.0), 3.119.0 (.NET 9.0, .NET 10.0) |

For the 26.x-on-.NET-8 baseline this file assumes, that is **SkiaSharp 2.88.9** (except 25.1.1, which pins 3.116.1).

### Blank boxes, wrong metrics, or substituted fonts in output

Symptom: PDF or image output shows empty rectangles, clipped or overlapping text, wrong column widths / row heights, or a different typeface than Excel. Layout and numbers drift because glyph metrics differ from the intended font.

Cause: base OS and container images ship almost no fonts. When the workbook's font (Calibri, Segoe UI, ...) is absent, Aspose.Cells substitutes the closest available font, changing metrics and appearance.

Fix: install the actual `.ttf`/`.otf` files into the image, and - if they are not in a system font directory - register them. Every `FontConfigs` call must run **before** you construct or load any `Workbook`, and if several source-setting methods are called, only the last one takes effect.

```csharp
// Call FontConfigs BEFORE constructing or loading any Workbook.
FontConfigs.SetFontFolder("/usr/share/fonts", true);   // true = search subfolders
var wb = new Workbook("report.xlsx");
wb.Save("report.pdf", SaveFormat.Pdf);
```

Use `SetFontFolders` for several directories, and `SetFontSubstitutes` to redirect a font you cannot ship to one you can:

```csharp
FontConfigs.SetFontFolders(new string[] { "/usr/share/fonts", "/app/fonts" }, true);
// Redirect absent Windows fonts to metric-compatible fonts the image ships.
FontConfigs.SetFontSubstitutes("Calibri", new string[] { "Carlito", "DejaVu Sans" });
FontConfigs.SetFontSubstitutes("Segoe UI", new string[] { "DejaVu Sans" });
```

`FontConfigs.SetFontSources` additionally loads a single font file or an in-memory `byte[]`. The render-time last-resort default (`PdfSaveOptions.DefaultFont`, `ImageOrPrintOptions.DefaultFont`, `FontConfigs.DefaultFontName`) is a separate lever - see conversion-and-rendering.md. Installing the real fonts is always the more faithful fix.

### No printing to a physical printer on SkiaSharp builds

Symptom: printer output APIs are unavailable or fail off Windows.

Cause: the SkiaSharp backend does not implement printing to a physical printer; that capability exists only on the Windows `System.Drawing.Common` build.

Fix: on Linux/macOS/containers, render to PDF or image (`Save(..., SaveFormat.Pdf)`, or `SheetRender`/`WorkbookRender` to PNG) and print that downstream. Do not design a server flow around driving a printer directly.

### EMF and TIFF not supported on Linux

Symptom: saving or rendering to EMF or TIFF fails or produces empty output on Linux.

Cause: Aspose.Cells for .NET Standard does not support EMF and TIFF on Linux.

Fix: target PNG, JPEG, BMP, or PDF for image output on Linux. Keep any EMF/TIFF paths on Windows only.

## Docker (.NET 8, SkiaSharp, Debian)

Minimal known-good multi-stage Dockerfile. Add `SkiaSharp.NativeAssets.Linux` (matching the version table) to the project first.

```dockerfile
FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build
WORKDIR /src
COPY . .
RUN dotnet publish "MyApp.csproj" -c Release -o /app/publish

FROM mcr.microsoft.com/dotnet/aspnet:8.0 AS final
WORKDIR /app
# SkiaSharp.NativeAssets.Linux needs libfontconfig1 (skip if using .NoDependencies)
RUN apt-get update && apt-get install -y libfontconfig1 && rm -rf /var/lib/apt/lists/*
# Base images ship almost no fonts: copy your .ttf/.otf files into a system font dir.
COPY fonts/ /usr/share/fonts/
COPY --from=build /app/publish .
ENTRYPOINT ["dotnet", "MyApp.dll"]
```

Alpine delta: use base images `mcr.microsoft.com/dotnet/{sdk,aspnet}:8.0-alpine` and replace the apt line with `RUN apk update && apk add fontconfig`. As an apt alternative to copying `.ttf` files, install a font package (e.g. `fonts-liberation`, or `ttf-mscorefonts-installer` from the contrib repo, which prompts to accept a EULA). Fonts placed under `/usr/share/fonts` are found automatically; fonts elsewhere need `FontConfigs.SetFontFolder`.

## macOS

.NET 6+ (recommended): use the SkiaSharp backend - no `libgdiplus` needed. Adding Aspose.Cells pulls SkiaSharp, whose macOS native library ships in the base SkiaSharp package, so no extra native-assets package is normally required.

Legacy only - when you deliberately stay on the `System.Drawing.Common` backend (e.g. .NET Core 3.1): install libgdiplus via Homebrew.

```bash
/bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
brew install mono-libgdiplus
```

## AWS Lambda / serverless Linux

Lambda runs on Amazon Linux, so every non-Windows rule applies: Aspose.Cells >= 22.10.1 with `SkiaSharp.NativeAssets.Linux` (or `.NoDependencies`) matching the version table. You cannot `apt-get` into a managed runtime, so:
- Bundle the `.ttf`/`.otf` fonts inside the deployment package (or a Lambda layer).
- Register them at cold start with `FontConfigs.SetFontFolder("./fonts", true)` before the first `Workbook` - the same pattern used for read-only cloud servers where you cannot modify system directories.
- Write only under `/tmp` (the sole writable path); prefer saving to a `MemoryStream` and returning the bytes.

## Blazor

**Blazor Server**: runs server-side; treat it like any .NET server app. Add Aspose.Cells (SkiaSharp is pulled in automatically) and, on a Linux host, apply the Docker/Linux steps above.

**Blazor WebAssembly**: supported from **Aspose.Cells 25.1**; runs in-browser via WASM. Add the **SkiaSharp.Views.Blazor** package, version-matched to the SkiaSharp that Aspose.Cells references:

| Aspose.Cells for .NET | SkiaSharp |
|:---:|:---:|
| = 25.1.1 | 3.116.1 |
| >= 25.1.2 | 2.88.9 (net6.0, net8.0), 3.116.1 (net9.0) |

Known issue: publishing a **net8.0** WASM project with the **.NET 9 SDK** throws `System.PlatformNotSupportedException: PlatformNotSupported_HybridGlobalization`. Fix by pinning the 8.0 SDK with a solution-level `global.json`, retargeting the project to net9.0, or updating Visual Studio to 17.12.4 or later.

## Related
- conversion-and-rendering.md - render-time DefaultFont in PdfSaveOptions/ImageOrPrintOptions, PDF layout options, SheetRender/WorkbookRender image output
- setup-and-licensing.md - NuGet install, TFM/DLL variants (net6.0 vs net6.0-windows), license setup
