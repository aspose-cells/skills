# Server-Side Setup Guide

## Prerequisites

- .NET 6.0+ (for C#)
- NuGet packages: Aspose.Cells, Aspose.Cells.GridJs

## Step 1: Register Services

```csharp
// Program.cs
using Aspose.Cells.GridJs;

var builder = WebApplication.CreateBuilder(args);
builder.Services.AddControllersWithViews();

// Register GridJs service
builder.Services.AddScoped<IGridJsService, GridJsService>();
builder.Services.Configure<GridJsOptions>(options =>
{
    // Required: Set cache directory for workbook files
    options.FileCacheDirectory = @"D:\storage\Aspose.Cells.GridJs\";
    // Required: Set base route name for GridJs controller
    options.BaseRouteName = "/GridJs";
});
```

## Step 2: Create Controller

```csharp
using Aspose.Cells.GridJs;
using Microsoft.AspNetCore.Mvc;

[Route("[controller]/[action]")]
public class GridJsController : GridJsControllerBase
{
    private readonly IGridJsService _gridJsService;

    public GridJsController(IGridJsService gridJsService) : base(gridJsService)
    {
        _gridJsService = gridJsService;
    }

    [HttpGet]
    public IActionResult LoadSpreadsheet(string filename, string uid)
    {
        string fullFilePath = GetFullFilePath(filename);
        StringBuilder json = _gridJsService.DetailFileJsonWithUid(fullFilePath, uid);
        return Content(json.ToString(), "text/plain", Encoding.UTF8);
    }
}
```

## Step 3: Client-Side HTML

```html
<!DOCTYPE html>
<html>
<head>
    <link rel="stylesheet" href="https://unpkg.com/gridjs-spreadsheet/xspreadsheet.css" />
    <script src="https://unpkg.com/gridjs-spreadsheet/xspreadsheet.js"></script>
</head>
<body>
    <div id="spreadsheet" style="width:100%;height:95vh;"></div>
</body>
</html>
```

## Step 4: Client-Side JavaScript

```javascript
// Load workbook JSON from server
$.ajax({
    url: `/GridJs/LoadSpreadsheet?filename=Sample.xlsx&uid=${generateUUID()}`,
    method: "GET",
    success: function(responseJsonString) {
        const jsonData = JSON.parse(responseJsonString);
        const option = {
            updateMode: 'server',
            updateUrl: '/GridJs/UpdateCell',
            local: 'en'
        };
        const xs = x_spreadsheet('#spreadsheet', option)
            .loadData(jsonData.data, jsonData.actname)
            .setUniqueId(jsonData.uniqueid)
            .setFileName(jsonData.filename);
    }
});
```

## Important Notes

- `GridJsControllerBase` provides built-in actions: UpdateCell, ImageUrl, AddImage, AddImageByURL, CopyImage, Download, Ole
- The `uid` (unique identifier) is used for file caching - generate a UUID per session
- `FileCacheDirectory` must exist and be writable
- For large files, consider enabling `LazyLoading` in config
