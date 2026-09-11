# Configuration Reference

## Config (Static Settings)

Represents all the static settings for GridJs

These are static properties on the `Config` class that affect global behavior.

| Property | Description |
|----------|-------------|
| `SaveHtmlAsZip` | Gets/Sets  whether to save html file as zip archive,the default is false. |
| `SkipInvisibleShapes` | Gets/Sets  whether to skip shapes that are invisble to UI ,the default value is true. |
| `LazyLoading` | Gets/Sets  whether to load active worksheet only,the default is false. |
| `SameImageDetecting` | Gets/Sets  whether to check if images have same source,the default is true the default value is true. |
| `AutoOptimizeForLargeCells` | Gets/Sets  whether to automatically optimize the load performance for worksheet with large cells. it will ignore some style /borders to reduce the load  time. the default value is true. |
| `IslimitShapeOrImage` | Gets/Sets  whether to limit the total display shape/image count in one worksheet ,if set to true, GridJs will limit the total count of the display shapes or images in one worksheet  to MaxShapeOrImageCount the default value is true. |
| `MaxShapeOrImageCount` | Gets/Sets the total count of the display shapes or images in the active sheet,it takes effect  when IslimitShapeOrImage=true. the default value is 100. |
| `MaxTotalShapeOrImageCount` | Gets/Sets the total count of the display shapes or images  in the workbook,it takes effect  when IslimitShapeOrImage=true. the default value is 300. |
| `MaxShapeOrImageWidthOrHeight` | Gets/Sets the  max width or height for a shape or an image ,GridJs will ignore the shape or image with the width or height larger than this, it takes effect when IslimitShapeOrImage=true. the default value is 10000. |
| `MaxPdfSaveSeconds` | Gets/Sets the max timed out seconds when save to PDF. the default value is 10. |
| `IgnoreEmptyContent` | Gets/Sets whether to show  the max range which includes data ,style, merged cells and shapes. if the last row or column contains cells with  no value and formula but has custom style then we will not show this row/column when this vlaue is true。 the default value is true . |
| `UsePrintArea` | Gets/Sets whether to use PageSetup.PrintArea for the UI display range when the worksheet has PageSetup setting for PrintArea. the default value is false . |
| `IsCollaborative` | Gets/Sets  whether to support collabrative editing,the default is false. |
| `EnableChartClientRendering` | Gets/Sets whether to enable client-side chart rendering. If set to false, GridJs will fall back to server-generated chart images/shapes only. the default value is true. |
| `CustomPdfSaveOptions` | Gets/Sets the custom PdfSaveOptions for PDF export. If set, this will be used instead of the default options. the default value is null. |
| `ShowChartSheet` | Gets/Sets whether to show chart worksheet. the default value is false . |
| `EmptySheetMaxRow` | Gets/Sets default max row for an empty worksheet. the default value is 12. |
| `EmptySheetMaxCol` | Gets/Sets default max column for an empty worksheet. the default value is 15. |
| `PictureCacheDirectory` | Gets/Sets the cache directory for pictures.(this takes effect when GridJsWorkbook.CacheImp is null) the default path will be "_piccache" inside the FileCacheDirectory. |
| `FileCacheDirectory` | Gets/Sets the cache directory for storing spreadsheet file. We need to set it to a specific path before we use GridJs. |
| `BaseRouteName` | Gets/Sets the base route name for GridJs controller URL. the default is "/GridJs2". |
| `MessageTopic` | Gets/Sets the websocket destinations prefixed with "/topic". the default is "/topic/opr".used in collaborative mode only. |
| `AutoFitRowsHeightOnLoad` | Indicates whether to auto-fit row heights during file loading. The default value is false. Warning: Setting this to true will perform an auto-fit all rows operation post-load, which may have a noticeable impact on performance. |
| `AutoFitColumnsOnLoad` | Indicates whether to auto-fit column widths during file loading. The default value is false. Warning: Setting this to true will perform an auto-fit all columns operation post-load, which may have a noticeable impact on performance. |
| `RedactionUseClientGenerateId` | Indicates whether use client generate id instead of actual shape id in redaction related operations |

## GridJsOptions (Instance Settings)

Represents  all the load options for GridJs

These options are configured per-service instance via DI.

| Property | Description |
|----------|-------------|
| `SaveHtmlAsZip` | Gets/Sets  whether to save html file as zip archive,the default is false. |
| `SkipInvisibleShapes` | Gets/Sets  whether to skip shapes that are invisble to UI ,the default value is true. |
| `LazyLoading` | Gets/Sets  whether to load active worksheet only,the default is false. |
| `SameImageDetecting` | Gets/Sets  whether to check if images have same source,the default is true the default value is true. |
| `AutoOptimizeForLargeCells` | Gets/Sets  whether to automatically optimize the load performance for worksheet with large cells. it will ignore some style /borders to reduce the load  time. the default value is true. |
| `IslimitShapeOrImage` | Gets/Sets  whether to limit the total display shape/image count in one worksheet ,if set to true, GridJs will limit the total count of the display shapes or images in one worksheet  to MaxShapeOrImageCount the default value is true. |
| `MaxShapeOrImageCount` | Gets/Sets the total count of the display shapes or images in the active sheet,it takes effect  when IslimitShapeOrImage=true. the default value is 100. |
| `MaxTotalShapeOrImageCount` | Gets/Sets the total count of the display shapes or images  in the workbook,it takes effect  when IslimitShapeOrImage=true. the default value is 300. |
| `MaxShapeOrImageWidthOrHeight` | Gets/Sets the  max width or height for a shape or an image ,GridJs will ignore the shape or image with the width or height larger than this, it takes effect when IslimitShapeOrImage=true. the default value is 10000. |
| `MaxPdfSaveSeconds` | Gets/Sets the max timed out seconds when save to PDF. the default value is 10. |
| `IgnoreEmptyContent` | Gets/Sets whether to show  the max range which includes data ,style, merged cells and shapes. if the last row or column contains cells with  no value and formula but has custom style then we will not show this row/column when this vlaue is true。 the default value is true . |
| `UsePrintArea` | Gets/Sets whether to use PageSetup.PrintArea for the UI display range when the worksheet has PageSetup setting for PrintArea. the default value is false . |
| `IsCollaborative` | Gets/Sets  whether to support collabrative editing,the default is false. |
| `EnableChartClientRendering` | Gets/Sets whether to enable client-side chart rendering. If set to false, GridJs will emit chart images/shapes only and will not require client-side chart rendering. the default value is true. |
| `CustomPdfSaveOptions` | Gets/Sets the custom PdfSaveOptions for PDF export. If set, this will be used instead of the default options. the default value is null. |
| `ShowChartSheet` | Gets/Sets whether to show chart worksheet. the default value is false . |
| `EmptySheetMaxRow` | Gets/Sets default max row for an empty worksheet. the default value is 12. |
| `EmptySheetMaxCol` | Gets/Sets default max column for an empty worksheet. the default value is 15. |
| `PictureCacheDirectory` | Gets/Sets the cache directory for pictures.(this takes effect when GridJsWorkbook.CacheImp is null) the default path will be "_piccache" inside the FileCacheDirectory. |
| `FileCacheDirectory` | Gets/Sets the cache directory for storing spreadsheet file. We need to set it to a specific path before we use GridJs. |
| `FontFolders` | Gets/Sets the fonts folders for fonts in the rendered pictures/shapes |
| `BaseRouteName` | Gets/Sets the route URL base name for GridJs controller.the default is GridJs2 |
| `MessageTopic` | Gets/Sets the websocket destinations prefixed with "/topic". the default is "/topic/opr".used in collaborative mode only. |
| `AutoFitRowsHeightOnLoad` | Indicates whether to autofit rows height  when loading the file,the default value is false. |
| `AutoFitColumnsOnLoad` | Indicates whether to autofit columns width  when loading the file,the default value is false. |
| `CacheImp` | Custom  implemention for cache storage,If you want to store cache in stream way ,you  need to set and implement it. |
| `RedactionUseClientGenerateId` | Indicates whether use client generate id instead of actual shape id in redaction related operations |

## GridWorkbookSettings

Represents the settings of the workbook.

| Property | Description |
|----------|-------------|
| `MaxIteration` | Returns or sets the maximum number of iterations to resolve a circular reference, the default value is 100. |
| `Iteration` | Indicates whether use iteration to resolve circular references. |
| `ForceFullCalculate` | Indicates whether fully calculates every time when a calculation is triggered. |
| `CreateCalcChain` | Indicates whether create calculated formulas chain. Default is false. |
| `ReCalculateOnOpen` | Indicates whether re-calculate all formulas on opening file. Default is true. |
| `PrecisionAsDisplayed` | True if calculations in this workbook will be done using only the precision of the numbers as they're displayed |
| `Date1904` | Gets or sets a value which represents if the workbook uses the 1904 date system. |
| `EnableMacros` | Enable macros; Now it only works when copying a worksheet to other worksheet in a workbook. |
| `CheckCustomNumberFormat` | Indicates whether checking custom number format when setting Style.Custom, default is false. |
| `CheckExcelRestriction` | Whether check restriction of excel file when user modify cells related objects. For example, excel does not allow inputting string value longer than 32K. When you input a value longer than 32K such as by Cell.PutValue(string), if this property is true, you will get an Exception. If this property is false, we will accept your input string value as the cell's value so that later you can output the complete string value for other file formats such as CSV. However, if you have set such kind of value that is invalid for excel file format, you should not save the workbook as excel file format later. Otherwise there may be unexpected error for the generated excel file. default is false. |
| `Author` | Gets/sets the author of the file. |

