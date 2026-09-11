# Server-Side API Reference

Assembly: **Aspose.Cells.GridJs**

## Config

**Kind**: Class

**Namespace**: `Aspose.Cells.GridJs`

**Description**: Represents all the static settings for GridJs

### Properties & Fields

| Name | Type | Description |
|------|------|-------------|
| `SaveHtmlAsZip` | Property | Gets/Sets  whether to save html file as zip archive,the default is false. |
| `SkipInvisibleShapes` | Property | Gets/Sets  whether to skip shapes that are invisble to UI ,the default value is true. |
| `LazyLoading` | Property | Gets/Sets  whether to load active worksheet only,the default is false. |
| `SameImageDetecting` | Property | Gets/Sets  whether to check if images have same source,the default is true the default value is true. |
| `AutoOptimizeForLargeCells` | Property | Gets/Sets  whether to automatically optimize the load performance for worksheet with large cells. it will ignore some style /borders to reduce the load  time. the default value is true. |
| `IslimitShapeOrImage` | Property | Gets/Sets  whether to limit the total display shape/image count in one worksheet ,if set to true, GridJs will limit the total count of the display shapes or images in one worksheet  to MaxShapeOrImageCount the default value is true. |
| `MaxShapeOrImageCount` | Property | Gets/Sets the total count of the display shapes or images in the active sheet,it takes effect  when IslimitShapeOrImage=true. the default value is 100. |
| `MaxTotalShapeOrImageCount` | Property | Gets/Sets the total count of the display shapes or images  in the workbook,it takes effect  when IslimitShapeOrImage=true. the default value is 300. |
| `MaxShapeOrImageWidthOrHeight` | Property | Gets/Sets the  max width or height for a shape or an image ,GridJs will ignore the shape or image with the width or height larger than this, it takes effect when IslimitShapeOrImage=true. the default value is 10000. |
| `MaxPdfSaveSeconds` | Property | Gets/Sets the max timed out seconds when save to PDF. the default value is 10. |
| `IgnoreEmptyContent` | Property | Gets/Sets whether to show  the max range which includes data ,style, merged cells and shapes. if the last row or column contains cells with  no value and formula but has custom style then we will not show this row/column when this vlaue is true。 the default value is true . |
| `UsePrintArea` | Property | Gets/Sets whether to use PageSetup.PrintArea for the UI display range when the worksheet has PageSetup setting for PrintArea. the default value is false . |
| `IsCollaborative` | Property | Gets/Sets  whether to support collabrative editing,the default is false. |
| `EnableChartClientRendering` | Property | Gets/Sets whether to enable client-side chart rendering. If set to false, GridJs will fall back to server-generated chart images/shapes only. the default value is true. |
| `CustomPdfSaveOptions` | Property | Gets/Sets the custom PdfSaveOptions for PDF export. If set, this will be used instead of the default options. the default value is null. |
| `ShowChartSheet` | Property | Gets/Sets whether to show chart worksheet. the default value is false . |
| `EmptySheetMaxRow` | Property | Gets/Sets default max row for an empty worksheet. the default value is 12. |
| `EmptySheetMaxCol` | Property | Gets/Sets default max column for an empty worksheet. the default value is 15. |
| `PictureCacheDirectory` | Property | Gets/Sets the cache directory for pictures.(this takes effect when GridJsWorkbook.CacheImp is null) the default path will be "_piccache" inside the FileCacheDirectory. |
| `FileCacheDirectory` | Property | Gets/Sets the cache directory for storing spreadsheet file. We need to set it to a specific path before we use GridJs. |
| `BaseRouteName` | Property | Gets/Sets the base route name for GridJs controller URL. the default is "/GridJs2". |
| `MessageTopic` | Property | Gets/Sets the websocket destinations prefixed with "/topic". the default is "/topic/opr".used in collaborative mode only. |
| `AutoFitRowsHeightOnLoad` | Property | Indicates whether to auto-fit row heights during file loading. The default value is false. Warning: Setting this to true will perform an auto-fit all rows operation post-load, which may have a noticeable impact on performance. |
| `AutoFitColumnsOnLoad` | Property | Indicates whether to auto-fit column widths during file loading. The default value is false. Warning: Setting this to true will perform an auto-fit all columns operation post-load, which may have a noticeable impact on performance. |
| `RedactionUseClientGenerateId` | Field | Indicates whether use client generate id instead of actual shape id in redaction related operations |

### Methods

#### SetFontFolder

```csharp
SetFontFolder(fontFolder, recursive)
```

Sets the fonts folder

**Parameters**:

- `fontFolder`: The folder that contains TrueType fonts.
- `recursive`: Determines whether or not to scan subfolders.

#### SetFontFolders

```csharp
SetFontFolders(fontFolders, recursive)
```

Sets the fonts folders

**Parameters**:

- `fontFolders`: The folders that contains TrueType fonts.
- `recursive`: Determines whether or not to scan subfolders.

---

## CoWorkUserPermission

**Kind**: Class

**Namespace**: `Aspose.Cells.GridJs`

**Description**: represent the user permission  in collaboration mode.only available in java version now, will be available in .net/python version in future.

### Properties & Fields

| Name | Type | Description |
|------|------|-------------|
| `READ_ONLY` | Field | Read-only permission,User can only view the content but cannot make any modifications |
| `EDITABLE` | Field | Editable permission,User can view and edit the content |
| `DOWNLOAD` | Field | Download permission,User can view,edit and download the content to local storage |
| `ADMIN` | Field | Administrator permission,User has full access including management operations |

---

## CoWorkUserProvider

**Kind**: Interface

**Namespace**: `Aspose.Cells.GridJs`

**Description**: Represents the user provider inerface in collabration mode.only available in java version now, will be available in .net/python version in future. Customer application can implement this interface to provide the user information.

### Methods

#### GetCurrentUserName

```csharp
GetCurrentUserName()
```

Gets the username of the current user

**Returns**: Current username

#### GetCurrentUserId

```csharp
GetCurrentUserId()
```

Gets the unique identifier of the current user

**Returns**: Current user ID

#### GetPermission

```csharp
GetPermission()
```

Gets the permission level of the current user

**Returns**: Current user permission level

---

## GridAbstractCalculationEngine

**Kind**: Class

**Namespace**: `Aspose.Cells.GridJs`

**Description**: Represents user's custom calculation engine to extend the default calculation engine of Aspose.Cells.

### Methods

#### Calculate

```csharp
Calculate(data)
```

Calculates one function with given data.

**Parameters**:

- `data`: The required data to calculate function such as function name, parameters, ...etc.

---

## GridCacheForStream

**Kind**: Class

**Namespace**: `Aspose.Cells.GridJs`

**Description**: This class contains the cache operations for GridJs. User shall implement his own business logic for storage based on it..

### Methods

#### SaveStream

```csharp
SaveStream()
```

Implements this method to save cache,save the stream to the cache with the key uid.

#### LoadStream

```csharp
LoadStream()
```

Implements this method to load cache with the key uid,return the stream from the cache.

#### IsExisted

```csharp
IsExisted(uid)
```

Checks whether the cache with uid is existed

**Parameters**:

- `uid`: The unique id for the file cache.

**Returns**: The bool value

#### GetFileUrl

```csharp
GetFileUrl(uid)
```

Implements this method to get the file url  from the cache.

**Parameters**:

- `uid`: The unique id for the file cache.

**Returns**: The URL of the file

---

## GridCalculationData

**Kind**: Class

**Namespace**: `Aspose.Cells.GridJs`

**Description**: Represents the required data when calculating one function, such as function name, parameters, ...etc.

### Properties & Fields

| Name | Type | Description |
|------|------|-------------|
| `CalculatedValue` | Property | Gets/sets the calculated value for this function. |
| `Row` | Property | Gets the Cell Row index where the function is in. |
| `Column` | Property | Gets the Cell Column index where the function is in. |
| `StringValue` | Property | Gets the Cell DisplayStringValue where the function is in. |
| `Value` | Property | Gets the Cell value where the function is in. |
| `Formula` | Property | Gets the Cell formula where the function is in. |
| `SheetName` | Property | Gets the worksheet name where the function is in. |
| `FunctionName` | Property | Gets the function name to be calculated. |
| `ParamCount` | Property | Gets the count of parameters . |

### Methods

#### GetParamValue

```csharp
GetParamValue(index)
```

Gets the represented value object of the parameter at given index.

**Parameters**:

- `index`: The index of the parameter(0 based).

**Returns**: If the parameter is plain value, then returns the plain value. If the parameter is reference, then return ReferredArea object.

#### GetParamText

```csharp
GetParamText(index)
```

Gets the literal text of the parameter at given index.

**Parameters**:

- `index`: The index of the parameter(0 based).

**Returns**: The literal text of the parameter.

---

## GridCellException

**Kind**: Class

**Namespace**: `Aspose.Cells.GridJs`

**Description**: The exception that is thrown when GridJs specified error occurs.

### Properties & Fields

| Name | Type | Description |
|------|------|-------------|
| `Code` | Property | Represents the exception code. |
| `Message` | Property | Represents the exception message. |

### Methods

#### ToString

```csharp
ToString()
```

Creates and returns a string representation of the current exception.

---

## GridExceptionType

**Kind**: Class

**Namespace**: `Aspose.Cells.GridJs`

**Description**: Represents custom exception code for GridJs.

### Properties & Fields

| Name | Type | Description |
|------|------|-------------|
| `Chart` | Field | Invalid chart setting. |
| `DataType` | Field | Invalid data type setting. |
| `DataValidation` | Field | Invalid data validation setting. |
| `ConditionalFormatting` | Field | Invalid data validation setting. |
| `FileFormat` | Field | Invalid file format. |
| `Formula` | Field | Invalid formula. |
| `InvalidData` | Field | Invalid data. |
| `InvalidOperator` | Field | Invalid operator. |
| `IncorrectPassword` | Field | Incorrect password. |
| `License` | Field | License related errors. |
| `Limitation` | Field | Out of MS Excel limitation error. |
| `PageSetup` | Field | Invalid page setup setting. |
| `PivotTable` | Field | Invalid pivotTable setting. |
| `Shape` | Field | Invalid drawing object setting. |
| `Sparkline` | Field | Invalid sparkline object setting. |
| `SheetName` | Field | Invalid worksheet name. |
| `SheetType` | Field | Invalid worksheet type. |
| `Interrupted` | Field | The process is interrupted. |
| `IO` | Field | The file is invalid. |
| `Permission` | Field | Permission is required to open this file. |
| `UnsupportedFeature` | Field | Unsupported feature. |
| `UnsupportedStream` | Field | Unsupported stream to be opened. |
| `UndisclosedInformation` | Field | Files contains some undisclosed information. |

---

## GridInterruptMonitor

**Kind**: Class

**Namespace**: `Aspose.Cells.GridJs`

**Description**: Represents all operator about the calculation interrupt.

### Methods

#### Interrupt

```csharp
Interrupt()
```

Interrupt the current operator.

---

## GridJsLogger

**Kind**: Class

**Namespace**: `Aspose.Cells.GridJs`

**Description**: Provides a lightweight logging system for GridJs server. Supports console and file output with configurable log levels. Default level is None. In DEBUG build, the default level is Debug (outputs all messages). Supports log file rolling by day/week/month.

### Properties & Fields

| Name | Type | Description |
|------|------|-------------|
| `Level` | Property | Gets or sets the minimum log level. Messages below this level will be ignored. Default is None. |
| `EnableConsoleOutput` | Property | Gets or sets whether to output log messages to the console. Default is false. |
| `EnableFileOutput` | Property | Gets or sets whether to output log messages to a file. Default is false. Must call  or  before enabling. |
| `LogFilePath` | Property | Gets the current log file path, or null if not set. |
| `MaxFileSize` | Property | Gets or sets the maximum size (in bytes) for a single log file before rolling. Only effective when rolling is enabled. 0 means no size limit. Default is 0. |

### Methods

#### SetLogFile

```csharp
SetLogFile(filePath)
```

Sets the log file path and enables file output (single file mode, no rolling). Appends to existing file.

**Parameters**:

- `filePath`: The full path to the log file.

#### SetLogFile

```csharp
SetLogFile(filePath, append)
```

Sets the log file path and enables file output (single file mode, no rolling).

**Parameters**:

- `filePath`: The full path to the log file.
- `append`: Whether to append to existing file.

#### SetLogDirectory

```csharp
SetLogDirectory(directory)
```

Sets the log directory with daily rolling file strategy and enables file output. Log files will be named as: gridjs_{period}.log

**Parameters**:

- `directory`: The directory to store log files.

#### SetLogDirectory

```csharp
SetLogDirectory(directory, rolling)
```

Sets the log directory with rolling file strategy and enables file output. Log files will be named as: gridjs_{period}.log

**Parameters**:

- `directory`: The directory to store log files.
- `rolling`: The rolling strategy (Daily/Weekly/Monthly).

#### SetLogDirectory

```csharp
SetLogDirectory(directory, rolling, filePrefix)
```

Sets the log directory with rolling file strategy and enables file output. Log files will be named as: {prefix}_{period}.log Examples: gridjs_2026-07-27.log (Daily), gridjs_2026-W30.log (Weekly), gridjs_2026-07.log (Monthly)

**Parameters**:

- `directory`: The directory to store log files.
- `rolling`: The rolling strategy (Daily/Weekly/Monthly).
- `filePrefix`: The file name prefix.

#### CloseLogFile

```csharp
CloseLogFile()
```

Closes the log file writer and disables file output.

#### Debug

```csharp
Debug()
```

Logs a debug level message.

#### Info

```csharp
Info()
```

Logs an info level message.

#### Warn

```csharp
Warn()
```

Logs a warning level message.

#### Error

```csharp
Error()
```

Logs an error level message.

#### Log

```csharp
Log(level, message)
```

Logs a message at the specified level.

**Parameters**:

- `level`: The log level.
- `message`: The message to log.

---

## GridJsLogLevel

**Kind**: Class

**Namespace**: `Aspose.Cells.GridJs`

**Description**: Specifies the log level for GridJs logging.

### Properties & Fields

| Name | Type | Description |
|------|------|-------------|
| `Debug` | Field | Debug level, outputs all messages. |
| `Info` | Field | Info level. |
| `Warn` | Field | Warning level. |
| `Error` | Field | Error level. |
| `None` | Field | Disables all logging. |

---

## GridJsLogRolling

**Kind**: Class

**Namespace**: `Aspose.Cells.GridJs`

**Description**: Specifies the log file rolling strategy.

### Properties & Fields

| Name | Type | Description |
|------|------|-------------|
| `None` | Field | No rolling, single file. |
| `Daily` | Field | Roll by day, e.g. gridjs_2026-07-27.log |
| `Monthly` | Field | Roll by month, e.g. gridjs_2026-07.log |

---

## GridJsOptions

**Kind**: Class

**Namespace**: `Aspose.Cells.GridJs`

**Description**: Represents  all the load options for GridJs

### Properties & Fields

| Name | Type | Description |
|------|------|-------------|
| `SaveHtmlAsZip` | Property | Gets/Sets  whether to save html file as zip archive,the default is false. |
| `SkipInvisibleShapes` | Property | Gets/Sets  whether to skip shapes that are invisble to UI ,the default value is true. |
| `LazyLoading` | Property | Gets/Sets  whether to load active worksheet only,the default is false. |
| `SameImageDetecting` | Property | Gets/Sets  whether to check if images have same source,the default is true the default value is true. |
| `AutoOptimizeForLargeCells` | Property | Gets/Sets  whether to automatically optimize the load performance for worksheet with large cells. it will ignore some style /borders to reduce the load  time. the default value is true. |
| `IslimitShapeOrImage` | Property | Gets/Sets  whether to limit the total display shape/image count in one worksheet ,if set to true, GridJs will limit the total count of the display shapes or images in one worksheet  to MaxShapeOrImageCount the default value is true. |
| `MaxShapeOrImageCount` | Property | Gets/Sets the total count of the display shapes or images in the active sheet,it takes effect  when IslimitShapeOrImage=true. the default value is 100. |
| `MaxTotalShapeOrImageCount` | Property | Gets/Sets the total count of the display shapes or images  in the workbook,it takes effect  when IslimitShapeOrImage=true. the default value is 300. |
| `MaxShapeOrImageWidthOrHeight` | Property | Gets/Sets the  max width or height for a shape or an image ,GridJs will ignore the shape or image with the width or height larger than this, it takes effect when IslimitShapeOrImage=true. the default value is 10000. |
| `MaxPdfSaveSeconds` | Property | Gets/Sets the max timed out seconds when save to PDF. the default value is 10. |
| `IgnoreEmptyContent` | Property | Gets/Sets whether to show  the max range which includes data ,style, merged cells and shapes. if the last row or column contains cells with  no value and formula but has custom style then we will not show this row/column when this vlaue is true。 the default value is true . |
| `UsePrintArea` | Property | Gets/Sets whether to use PageSetup.PrintArea for the UI display range when the worksheet has PageSetup setting for PrintArea. the default value is false . |
| `IsCollaborative` | Property | Gets/Sets  whether to support collabrative editing,the default is false. |
| `EnableChartClientRendering` | Property | Gets/Sets whether to enable client-side chart rendering. If set to false, GridJs will emit chart images/shapes only and will not require client-side chart rendering. the default value is true. |
| `CustomPdfSaveOptions` | Property | Gets/Sets the custom PdfSaveOptions for PDF export. If set, this will be used instead of the default options. the default value is null. |
| `ShowChartSheet` | Property | Gets/Sets whether to show chart worksheet. the default value is false . |
| `EmptySheetMaxRow` | Property | Gets/Sets default max row for an empty worksheet. the default value is 12. |
| `EmptySheetMaxCol` | Property | Gets/Sets default max column for an empty worksheet. the default value is 15. |
| `PictureCacheDirectory` | Property | Gets/Sets the cache directory for pictures.(this takes effect when GridJsWorkbook.CacheImp is null) the default path will be "_piccache" inside the FileCacheDirectory. |
| `FileCacheDirectory` | Property | Gets/Sets the cache directory for storing spreadsheet file. We need to set it to a specific path before we use GridJs. |
| `FontFolders` | Property | Gets/Sets the fonts folders for fonts in the rendered pictures/shapes |
| `BaseRouteName` | Property | Gets/Sets the route URL base name for GridJs controller.the default is GridJs2 |
| `MessageTopic` | Property | Gets/Sets the websocket destinations prefixed with "/topic". the default is "/topic/opr".used in collaborative mode only. |
| `AutoFitRowsHeightOnLoad` | Property | Indicates whether to autofit rows height  when loading the file,the default value is false. |
| `AutoFitColumnsOnLoad` | Property | Indicates whether to autofit columns width  when loading the file,the default value is false. |
| `CacheImp` | Field | Custom  implemention for cache storage,If you want to store cache in stream way ,you  need to set and implement it. |
| `RedactionUseClientGenerateId` | Field | Indicates whether use client generate id instead of actual shape id in redaction related operations |

---

## GridJsPermissionException

**Kind**: Class

**Namespace**: `Aspose.Cells.GridJs`

**Description**: represents permission exception in collaboration mode.only available in java version now, will be available in .net/python version in future.

### Constructors

```csharp
GridJsPermissionException(operation, requiredPermission, currentPermission)
```

Constructs a permission exception

**Parameters**:

- `operation`: Operation attempted to execute
- `requiredPermission`: Permission required for the operation
- `currentPermission`: Current user permission

```csharp
GridJsPermissionException(message, operation, requiredPermission, currentPermission)
```

Constructs a permission exception

**Parameters**:

- `message`: Custom error message
- `operation`: Operation attempted to execute
- `requiredPermission`: Permission required for the operation
- `currentPermission`: Current user permission

### Methods

#### GetRequiredPermission

```csharp
GetRequiredPermission()
```

Gets the permission level required for the operation

**Returns**: Required permission level

#### GetCurrentPermission

```csharp
GetCurrentPermission()
```

Gets the current permission level of the user

**Returns**: Current user permission level

#### GetOperation

```csharp
GetOperation()
```

Gets the operation that was attempted to execute

**Returns**: Operation name

---

## GridJsService

**Kind**: Class

**Namespace**: `Aspose.Cells.GridJs`

**Description**: Provides the basic operation apis used in controller actions.

### Constructors

```csharp
GridJsService()
```

The default constructor for GridJsService

### Properties & Fields

| Name | Type | Description |
|------|------|-------------|
| `Settings` | Property | Represents the workbook settings. |

### Methods

#### GetWarningCallback

```csharp
GetWarningCallback()
```

Gets custom warning callback for import file.

**Returns**: The warning callback

#### SetWarningCallback

```csharp
SetWarningCallback(callback)
```

Sets custom warning callback for import file.

**Parameters**:

- `callback`: The warning callback

#### CheckInCacheForCollaborative

```csharp
CheckInCacheForCollaborative(uid)
```

Check wether workbook instance is in memory cache .this method is apply for Collaborative mode only.

**Parameters**:

- `uid`: 

#### UpdateCell

```csharp
UpdateCell(p, uid)
```

Applies the update operation.

**Parameters**:

- `p`: The JSON format string of update operation.
- `uid`: The unique id for the file cache.

**Returns**: The JSON format string of the update result.

#### DetailFileJsonWithUid

```csharp
DetailFileJsonWithUid(filePath, uid)
```

Gets JSON string for the file by the specified unique id. .

**Parameters**:

- `filePath`: The file path
- `uid`: The unique id for the file cache.

**Returns**: The JSON StringBuilder.

#### DetailFileJsonWithUid

```csharp
DetailFileJsonWithUid(wb, filename, uid)
```

Gets JSON string for the Workbook  by the specified unique id.

**Parameters**:

- `wb`: the Workbook instance
- `filename`: The file name
- `uid`: The unique id for the file cache.

**Returns**: The JSON StringBuilder.

#### DetailStreamJsonWithUid

```csharp
DetailStreamJsonWithUid(stream, filePath, uid)
```

Write the JSON string  for the file to the stream  by the specified unique id.

**Parameters**:

- `stream`: The stream that will be written
- `filePath`: The file path
- `uid`: The unique id for the file cache.

#### DetailStreamJsonWithUid

```csharp
DetailStreamJsonWithUid(stream, wb, filename, uid)
```

Write the JSON string  for the Workbook to the stream  by the specified unique id.

**Parameters**:

- `stream`: The stream that will be written
- `wb`: The Workbook instance
- `filename`: The file name
- `uid`: The unique id for the file cache.

#### DetailStreamJson

```csharp
DetailStreamJson(stream, filePath)
```

Write the JSON string  for the file to the stream .

**Parameters**:

- `stream`: The stream that will be written
- `filePath`: The file path

#### DetailStreamJson

```csharp
DetailStreamJson(stream, wb, filename)
```

Write the JSON string  for the Workbook to the stream

**Parameters**:

- `stream`: The stream that will be written
- `wb`: The Workbook instance
- `filename`: The file name

#### LazyLoadingJson

```csharp
LazyLoadingJson(sheetName, uid)
```

Gets the JSON string of the specified sheet in the file from the cache using the specified unique id.

**Parameters**:

- `sheetName`: the sheet name.
- `uid`: The unique id for the file cache.

**Returns**: The JSON string StringBuilder

#### LazyLoadingStreamJson

```csharp
LazyLoadingStreamJson(stream, sheetName, uid)
```

Writes the JSON string of the specified sheet in the file from the cache using the specified unique id  to the stream..

**Parameters**:

- `stream`: The stream that will be written
- `sheetName`: The sheet name.
- `uid`: The unique id for the file cache.

#### AddImage

```csharp
AddImage(p, uid, iscontrol, files)
```

Applies the add image from local file operation.

**Parameters**:

- `p`: The JSON string parameter
- `uid`: The unique id for the file cache.
- `iscontrol`: Specify whether it is a control.
- `files`: The form file of the image

**Returns**: The JSON string result

#### AddImageByURL

```csharp
AddImageByURL(p, uid, imageurl)
```

Applies the add image from remote URL operation.

**Parameters**:

- `p`: The JSON string parameter
- `uid`: The unique id for the file cache.
- `imageurl`: Specify the image URL.

**Returns**: The JSON string result

#### CopyImage

```csharp
CopyImage(p, uid)
```

Applies the copy image operation.

**Parameters**:

- `p`: The JSON string parameter
- `uid`: The unique id for the file cache.

**Returns**: The JSON string result

#### Load

```csharp
Load(uid, filename)
```

Gets the JSON  string  of the file from the cache using the specified unique id,set the output filename in the JSON.

**Parameters**:

- `uid`: The unique id for the file cache.
- `filename`: Specifies the file name in the JSON. If set to null,the default filename is: book1.

**Returns**: The JSON string

#### Image

```csharp
Image(uid, picid)
```

Get Stream of image.

**Parameters**:

- `uid`: The unique id for the file cache.
- `picid`: The image id.

**Returns**: The image stream

#### Ole

```csharp
Ole(uid, sheetname, oleid, label)
```

Gets the byte array data of the  embedded ole object .

**Parameters**:

- `uid`: The unique id for the file cache.
- `sheetname`: The worksheet name.
- `oleid`: The  id for the embedded ole object.
- `label`: The display label of the embedded ole object.

**Returns**: The byte array data of the  embedded ole object .

#### ImageUrl

```csharp
ImageUrl(baseURL, picid, uid)
```

Gets the image URL.

**Parameters**:

- `baseURL`: The base action URL.
- `picid`: The image id.
- `uid`: The unique id for the file cache.

**Returns**: The image URL

#### GetFile

```csharp
GetFile(fileid)
```

Get file stream

**Parameters**:

- `fileid`: the file id

**Returns**: The stream of the file

#### TranslateSheetAsync

```csharp
TranslateSheetAsync(uid, sheetName, translator, targetLanguage )
```

Translate all the string value to the target language in the worksheet

**Parameters**:

- `uid`: The unique id for the file cache.
- `sheetName`: The sheet name
- `translator`: The translator which implement translate function
- `targetLanguage `: The target language

#### Download

```csharp
Download(p, uid, filename)
```

Applies the download file operation

**Parameters**:

- `p`: The JSON parameter
- `uid`: The unique id for the file cache.
- `filename`: The file name

**Returns**: The file URL

#### Dispose

```csharp
Dispose()
```

Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.

#### Destroy

```csharp
Destroy()
```

this method is for java only ,it implement the destroy method in DisposableBean ,which will be called automatically by Spring when finish request, actually it will call Dispose method

---

## GridJsWorkbook

**Kind**: Class

**Namespace**: `Aspose.Cells.GridJs`

**Description**: Represents the main entry class for GridJs

### Constructors

```csharp
GridJsWorkbook()
```

Creates a new instance of GridJsWorkbook.

### Properties & Fields

| Name | Type | Description |
|------|------|-------------|
| `Settings` | Property | Represents the workbook settings. |
| `WarningCallback` | Property | Gets or sets warning callback for import file. |
| `CacheDirectory` | Property | Gets/Sets the cache directory for storing spreadsheet file. if this is not set ,we will use the static propertie from  or  . |
| `CacheImp` | Field | Custom  implemention for cache storage,If you want to store cache in stream way ,you  need to set and implement it. |
| `CalculateEngine` | Field | Custom  implemention for calculation engine ,If you want to do custom calculation, you  need to set and implement it. |
| `UpdateMonitor` | Field | Gets/Sets the update monitor to track update operation |
| `pictureType` | Field | const value for the type of the image |

### Methods

#### SetInterruptMonitorForLoad

```csharp
SetInterruptMonitorForLoad(monitor, calculateTimeoutMilliseconds)
```

Sets InterruptMonitor for load operation.

**Parameters**:

- `monitor`: The InterruptMonitor instance.
- `calculateTimeoutMilliseconds`: The time out in millisecond for load file.

#### SetInterruptMonitorForSave

```csharp
SetInterruptMonitorForSave(monitor)
```

Sets InterruptMonitor for save operation.

**Parameters**:

- `monitor`: The InterruptMonitor instance.

#### GetJsonByUid

```csharp
GetJsonByUid(uid, filename)
```

Gets the JSON  string  of the file from the cache using the specified unique id,set the output filename in the JSON.

**Parameters**:

- `uid`: The unique id for the file cache.
- `filename`: Specifies the file name in the JSON. If set to null,the default filename is: book1.

**Returns**: The JSON string StringBuilder

#### JsonToStreamByUid

```csharp
JsonToStreamByUid(stream, uid, filename)
```

Retrieve the JSON string of the file from the cache using the specified unique id,set the output filename in the JSON,and write it to the stream.

**Parameters**:

- `stream`: The stream that will be written
- `uid`: The unique id for the file cache.
- `filename`: Specifies the file name in the JSON. If set to null,the default filename is: book1.

#### JsonToStream

```csharp
JsonToStream(stream, filename)
```

Retrieve the JSON string from memory data,set the output filename in the JSON, and write it to the stream.

**Parameters**:

- `stream`: The stream that will be written
- `filename`: Specifies the file name in the JSON. If set to null,the default filename is: book1.

#### LazyLoadingStream

```csharp
LazyLoadingStream(stream, uid, sheetName)
```

Retrieve the JSON string of the specified sheet in the file from the cache using the specified unique id, and write it to the stream.

**Parameters**:

- `stream`: The stream that will be written
- `uid`: The unique id for the file cache.
- `sheetName`: the sheet name.

#### LazyLoadingJson

```csharp
LazyLoadingJson(uid, sheetName)
```

Gets the JSON string of the specified sheet in the file from the cache using the specified unique id.

**Parameters**:

- `uid`: The unique id for the file cache.
- `sheetName`: the sheet name.

**Returns**: The JSON string StringBuilder

#### ImportExcelFile

```csharp
ImportExcelFile(uid, fileName, password)
```

Imports the excel file from file path and open password.

**Parameters**:

- `uid`: The unique id for the file cache, if set to null,it will be generated automatically.
- `fileName`: The full path of the file.
- `password`: The open password  of the excel file.The value can be null If no passowrd is set.

#### ImportExcelFile

```csharp
ImportExcelFile(uid, fileName)
```

Imports the excel file from the file path.

**Parameters**:

- `uid`: The unique id for the file cache, if set to null,it will be generated automatically.
- `fileName`: The full path of the file.

#### ImportExcelFile

```csharp
ImportExcelFile(fileName)
```

Imports the excel file from the file path.

**Parameters**:

- `fileName`: The full path of the file.

#### ImportExcelFile

```csharp
ImportExcelFile(wb)
```

Imports the excel file from the Workbook object.

**Parameters**:

- `wb`: The Workbook object.

#### ImportExcelFile

```csharp
ImportExcelFile(uid, wb)
```

Imports the excel file from the Workbook object.

**Parameters**:

- `uid`: The unique id for the file cache, if set to null,it will be generated automatically.
- `wb`: The Workbook object .

#### GetUidForFile

```csharp
GetUidForFile(fileName)
```

Generates a new unique id for the file cache using the given file name.

**Parameters**:

- `fileName`: The file name.

#### ImportExcelFile

```csharp
ImportExcelFile(uid, filestream, format, password)
```

Imports the excel file from  file stream with load format and open password.

**Parameters**:

- `uid`: The unique id for the file cache, if set to null,it will be generated automatically.
- `filestream`: The stream of the excel file .
- `format`: The LoadFormat of the excel file.
- `password`: The open password  of the excel file.The value can be null If no passowrd is set

#### ImportExcelFile

```csharp
ImportExcelFile(uid, filestream, format)
```

Imports the excel file from file stream.

**Parameters**:

- `uid`: The unique id for the file cache, if set to null,it will be generated automatically.
- `filestream`: The stream of the excel file .
- `format`: The LoadFormat of the excel file.

#### ImportExcelFile

```csharp
ImportExcelFile(filestream, format, password)
```

Imports the excel file from file stream with load format and open password.

**Parameters**:

- `filestream`: The stream of the excel file .
- `format`: The LoadFormat of the excel file.
- `password`: The open password  of the excel file.The value can be null If no passowrd is set.

#### ImportExcelFile

```csharp
ImportExcelFile(filestream, format)
```

Imports the excel file from file stream with load format.

**Parameters**:

- `filestream`: The stream of the excel file .
- `format`: The LoadFormat of the excel file.

#### ImportExcelFileFromJson

```csharp
ImportExcelFileFromJson(json)
```

Imports the excel file from JSON format string.

**Parameters**:

- `json`: The JSON format string.

#### MergeExcelFileFromJson

```csharp
MergeExcelFileFromJson(uid, json)
```

Applies a batch update to the memory data.

**Parameters**:

- `uid`: The unique id for the file cache.
- `json`: The update JSON format string.

#### ExportToJson

```csharp
ExportToJson(filename)
```

Gets JSON  string from memory data,set the output filename in the JSON.

**Parameters**:

- `filename`: Specifies the file name in the JSON. If set to null,the default filename is: book1..

**Returns**: The JSON string.

#### ExportToJson

```csharp
ExportToJson()
```

Gets JSON string from memory data, the default filename in the JSON is: book1.

**Returns**: The JSON string.

#### ExportToJsonStringBuilder

```csharp
ExportToJsonStringBuilder(filename)
```

Gets JSON string from memory data,set the output filename in the JSON.

**Parameters**:

- `filename`: Specifies the file name in the JSON. If set to null,the default filename is: book1..

**Returns**: The JSON StringBuilder.

#### SaveToExcelFile

```csharp
SaveToExcelFile(stream)
```

Saves the memory data to the sream, baseed on the origin file format.

**Parameters**:

- `stream`: The stream to save.

#### SaveToExcelFile

```csharp
SaveToExcelFile(path)
```

Saves the memory data to the file path,if the file has extension ,save format is baseed on the file extension .

**Parameters**:

- `path`: The file path to save.

#### SaveToCacheWithFileName

```csharp
SaveToCacheWithFileName(uid, filename, password)
```

Saves the memory data to the cache file with the specified filename and also set the open password, the save format is baseed on the file extension of the filename  .

**Parameters**:

- `uid`: The unique id for the file cache.
- `filename`: The filename to save.
- `password`: The excel file's open password. The value can be null If no passowrd is set.

#### SaveToPdf

```csharp
SaveToPdf(path)
```

Saves the memory data to the file path,the save format is pdf.

**Parameters**:

- `path`: The file path to save.

#### SaveToXlsx

```csharp
SaveToXlsx(path)
```

Saves the memory data to the file path,the save format is xlsx.

**Parameters**:

- `path`: The file path to save.

#### SaveToHtml

```csharp
SaveToHtml(path)
```

Saves the memory data to the file path,the save format is html.

**Parameters**:

- `path`: The file path to save.

#### SaveToPdf

```csharp
SaveToPdf(stream)
```

Saves the memory data to the sream,the save format is pdf.

**Parameters**:

- `stream`: The stream to save.

#### SaveToXlsx

```csharp
SaveToXlsx(stream)
```

Saves the memory data to the sream,the save format is xlsx.

**Parameters**:

- `stream`: The stream to save.

#### SaveToHtml

```csharp
SaveToHtml(stream)
```

Saves the memory data to the sream,the save format is html

**Parameters**:

- `stream`: The stream to save.

#### GetImageStream

```csharp
GetImageStream(uid, picid)
```

Get Stream of image.

**Parameters**:

- `uid`: The unique id for the file cache.
- `picid`: The image id.

**Returns**: The image stream

#### GetOle

```csharp
GetOle(uid, sheetname, oleid, label)
```

Gets the byte array data of the  embedded ole object .

**Parameters**:

- `uid`: The unique id for the file cache.
- `sheetname`: The worksheet name.
- `oleid`: The  id for the embedded ole object.
- `label`: The display label of the embedded ole object.

**Returns**: The byte array data of the  embedded ole object .

#### CheckInCacheForCollaborative

```csharp
CheckInCacheForCollaborative(uid)
```

Check wether workbook instance is in memory cache .this method is apply for Collaborative mode only.

**Parameters**:

- `uid`: 

#### UpdateCell

```csharp
UpdateCell(p, uid)
```

Applies the update operation.

**Parameters**:

- `p`: The JSON format string of update operation.
- `uid`: The unique id for the file cache.

**Returns**: The JSON format string of the update result.

#### RedactFile

```csharp
RedactFile(excelFilePath, uid, arrayOfRedactionOpr)
```

Performs redaction on an Excel file based on an array of JSON operations.

**Parameters**:

- `excelFilePath`: The file path of the Excel file to be redacted.
- `uid`: The unique identifier for the workbook. If null or empty, a new uid will be generated based on the file path.
- `arrayOfRedactionOpr`: An array of JSON strings representing the redaction operations to be applied.

#### SetTransParentView

```csharp
SetTransParentView(excelFilePath, uid, isTransparent)
```

Sets the transparency of redaction shapes in the workbook.

**Parameters**:

- `excelFilePath`: The file path of the Excel file.
- `uid`: The unique identifier for the workbook. If null or empty, a new uid will be generated based on the file path.
- `isTransparent`: If true, sets transparency to 0.89 (semi-transparent); if false, sets transparency to 1 (fully opaque/invisible).

#### BurnRedactionFile

```csharp
BurnRedactionFile(excelFilePath, uid)
```

Burns (applies) all redaction operations in the workbook by removing redaction shapes and their target shapes or clearing target cell range contents.

**Parameters**:

- `excelFilePath`: The file path of the Excel file.
- `uid`: The unique identifier for the workbook. If null or empty, a new uid will be generated based on the file path.

#### Dispose

```csharp
Dispose()
```

Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.

#### InsertImage

```csharp
InsertImage(uid, p, s, imageUrl)
```

Inserts image in the worksheet from file stream or the URL,(either the file stream or the URL shall be provided) or Inserts shape ,when the p.type is one of AutoShapeType

**Parameters**:

- `uid`: The unique id for the file cache
- `p`: The JSON format string for the operation which specify the cell location  ,the worksheet name,upper left row,upper left column for the image，etc  {name:'sheet1',ri:1,ci:1}
- `s`: The file stream of the image file
- `imageUrl`: The URL of the image file

**Returns**: The JSON format string of the inserted image

#### CopyImageOrShape

```csharp
CopyImageOrShape(uid, p)
```

Copys image or shape.

**Parameters**:

- `uid`: The unique id for the file cache.
- `p`: The JSON string for the operation which specify the cell location ,it contains the worksheet name,upper left row,upper left column for the image or shape，etc  {name:'sheet1',ri:1,ci:1,srcid:2,srcname:'sheet2',isshape:true}

**Returns**: The JSON string of the new copied image

#### ErrorJson

```csharp
ErrorJson(msg)
```

Gets the error message string in JSON format.

**Parameters**:

- `msg`: The error message.

**Returns**: The JSON string.

#### GetGridLoadFormat

```csharp
GetGridLoadFormat(extension)
```

Gets the load format by file extension

**Parameters**:

- `extension`: The file extention ,usually start with '.' .

#### GetImageUrl

```csharp
GetImageUrl(uid, picid, delimiter)
```

Gets the image URL.

**Parameters**:

- `uid`: The unique id for the file cache.
- `picid`: The image id.
- `delimiter`: The string delimiter.

#### SetImageUrlBase

```csharp
SetImageUrlBase(baseImageURL)
```

Set the base image get action URL from controller .

**Parameters**:

- `baseImageURL`: the base image get action URL.

#### TranslateSheetAsync

```csharp
TranslateSheetAsync(uid, sheetName, translator, targetLanguage )
```

Translate all the string value to the target language in the worksheet

**Parameters**:

- `uid`: The unique id for the file cache.
- `sheetName`: The sheet name
- `translator`: The translator which implement translate function
- `targetLanguage `: The target language

---

## GridLoadFormat

**Kind**: Class

**Namespace**: `Aspose.Cells.GridJs`

**Description**: Represents the load file format.

### Properties & Fields

| Name | Type | Description |
|------|------|-------------|
| `Auto` | Field | Represents recognizing the format automatically. |
| `Csv` | Field | Comma-Separated Values(CSV) text file. |
| `Xlsx` | Field | Represents Office Open XML spreadsheetML workbook or template, with or without macros. |
| `Tsv` | Field | Tab-Separated Values(TSV) text file. |
| `TabDelimited` | Field | Represents a tab delimited text file, same with . |
| `Html` | Field | Represents a html file. |
| `MHtml` | Field | Represents a mhtml file. |
| `Ods` | Field | Open Document Sheet(ODS) file. |
| `Excel97To2003` | Field | Represents an Excel97-2003 xls file. |
| `SpreadsheetML` | Field | Represents an Excel 2003 xml file. |
| `Xlsb` | Field | Represents an xlsb file. |
| `Ots` | Field | Open Document Template Sheet(OTS) file. |
| `Numbers` | Field | Represents a numbers file. |
| `Fods` | Field | Represents OpenDocument Flat XML Spreadsheet (.fods) file format. |
| `Sxc` | Field | Represents StarOffice Calc Spreadsheet (.sxc) file format. |
| `Xml` | Field | Represents a simple xml file. |
| `Epub` | Field | Reprents an EPUB file. |
| `Azw3` | Field | Represents an AZW3 file. |
| `Chm` | Field | Represents a CHM file. |
| `Markdown` | Field | Represents a Markdown file. |
| `Unknown` | Field | Represents unrecognized format, cannot be loaded. |
| `Image` | Field | Image |
| `Json` | Field | Json |
| `Dif` | Field | Data Interchange Format. |
| `Dbf` | Field | Xbase Data file |

---

## GridReferredArea

**Kind**: Class

**Namespace**: `Aspose.Cells.GridJs`

**Description**: Represents a referred area by the formula.

### Properties & Fields

| Name | Type | Description |
|------|------|-------------|
| `IsExternalLink` | Property | Indicates whether this is an external link. |
| `ExternalFileName` | Property | Get the external file name if this is an external reference. |
| `SheetName` | Property | Indicates which sheet this is in |
| `IsArea` | Property | Indicates whether this is an area. |
| `EndColumn` | Property | The end column of the area. |
| `StartColumn` | Property | The start column of the area. |
| `EndRow` | Property | The end row of the area. |
| `StartRow` | Property | The start row of the area. |

### Methods

#### GetValues

```csharp
GetValues()
```

Gets cell values in this area.

**Returns**: If this area is invalid, "#REF!" will be returned; If this area is one single cell, then return the cell value object; Otherwise return one array for all values in this area.

#### GetValue

```csharp
GetValue(rowOffset, colOffset)
```

Gets cell value with given offset from the top-left of this area.

**Parameters**:

- `rowOffset`: row offset from the start row of this area
- `colOffset`: column offset from the start row of this area

**Returns**: "#REF!" if this area is invalid; "#N/A" if given offset out of this area; Otherwise return the cell value at given position.

---

## GridUpdateMonitor

**Kind**: Class

**Namespace**: `Aspose.Cells.GridJs`

**Description**: Monitor for user to track the change of update operation.

### Methods

#### BeforeUpdate

```csharp
BeforeUpdate(op, uid, wb)
```

before update operation

**Parameters**:

- `op`: The JSON string for update operation
- `uid`: The unique id for the file cache
- `wb`: The Workbook object

#### AfterUpdate

```csharp
AfterUpdate(op, uid, cells)
```

after update operation

**Parameters**:

- `op`: The JSON string for update operation
- `uid`: The unique id for the file cache.
- `cells`: The Updated Cells list,include cells which has style change,value change or formula change

---

## GridWorkbookSettings

**Kind**: Class

**Namespace**: `Aspose.Cells.GridJs`

**Description**: Represents the settings of the workbook.

### Properties & Fields

| Name | Type | Description |
|------|------|-------------|
| `MaxIteration` | Property | Returns or sets the maximum number of iterations to resolve a circular reference, the default value is 100. |
| `Iteration` | Property | Indicates whether use iteration to resolve circular references. |
| `ForceFullCalculate` | Property | Indicates whether fully calculates every time when a calculation is triggered. |
| `CreateCalcChain` | Property | Indicates whether create calculated formulas chain. Default is false. |
| `ReCalculateOnOpen` | Property | Indicates whether re-calculate all formulas on opening file. Default is true. |
| `PrecisionAsDisplayed` | Property | True if calculations in this workbook will be done using only the precision of the numbers as they're displayed |
| `Date1904` | Property | Gets or sets a value which represents if the workbook uses the 1904 date system. |
| `EnableMacros` | Property | Enable macros; Now it only works when copying a worksheet to other worksheet in a workbook. |
| `CheckCustomNumberFormat` | Property | Indicates whether checking custom number format when setting Style.Custom, default is false. |
| `CheckExcelRestriction` | Property | Whether check restriction of excel file when user modify cells related objects. For example, excel does not allow inputting string value longer than 32K. When you input a value longer than 32K such as by Cell.PutValue(string), if this property is true, you will get an Exception. If this property is false, we will accept your input string value as the cell's value so that later you can output the complete string value for other file formats such as CSV. However, if you have set such kind of value that is invalid for excel file format, you should not save the workbook as excel file format later. Otherwise there may be unexpected error for the generated excel file. default is false. |
| `Author` | Property | Gets/sets the author of the file. |

---

## IGridJsService

**Kind**: Interface

**Namespace**: `Aspose.Cells.GridJs`

**Description**: Reprensents the basic operation apis interface used in controller actions.

### Methods

#### UpdateCell

```csharp
UpdateCell(p, uid)
```

Applies the update operation.

**Parameters**:

- `p`: The JSON format string of update operation.
- `uid`: The unique id for the file cache.

**Returns**: The JSON format string of the update result.

#### CheckInCacheForCollaborative

```csharp
CheckInCacheForCollaborative(uid)
```

Check wether workbook instance is in memory cache .this method is apply for Collaborative mode only.

**Parameters**:

- `uid`: 

#### DetailFileJsonWithUid

```csharp
DetailFileJsonWithUid(filePath, uid)
```

Gets JSON string for the file by the specified unique id. .

**Parameters**:

- `filePath`: The file path
- `uid`: The unique id for the file cache.

**Returns**: The JSON StringBuilder.

#### DetailFileJsonWithUid

```csharp
DetailFileJsonWithUid(wb, filename, uid)
```

Gets JSON string for the Workbook  by the specified unique id.

**Parameters**:

- `wb`: the Workbook instance
- `filename`: The file name
- `uid`: The unique id for the file cache.

**Returns**: The JSON StringBuilder.

#### DetailStreamJsonWithUid

```csharp
DetailStreamJsonWithUid(stream, filePath, uid)
```

Write the JSON string  for the file to the stream  by the specified unique id.

**Parameters**:

- `stream`: The stream that will be written
- `filePath`: The file path
- `uid`: The unique id for the file cache.

#### DetailStreamJsonWithUid

```csharp
DetailStreamJsonWithUid(stream, wb, filename, uid)
```

Write the JSON string  for the Workbook to the stream  by the specified unique id.

**Parameters**:

- `stream`: The stream that will be written
- `wb`: The Workbook instance
- `filename`: The file name
- `uid`: The unique id for the file cache.

#### DetailStreamJson

```csharp
DetailStreamJson(stream, filePath)
```

Write the JSON string  for the Workbook to the stream

**Parameters**:

- `stream`: The stream that will be written
- `filePath`: The file path

#### DetailStreamJson

```csharp
DetailStreamJson(stream, wb, filename)
```

Write the JSON string  for the Workbook to the stream

**Parameters**:

- `stream`: The stream that will be written
- `wb`: The Workbook instance
- `filename`: The file name

#### LazyLoadingJson

```csharp
LazyLoadingJson(sheetName, uid)
```

Gets the JSON string of the specified sheet in the file from the cache using the specified unique id.

**Parameters**:

- `sheetName`: the sheet name.
- `uid`: The unique id for the file cache.

**Returns**: The JSON string StringBuilder

#### LazyLoadingStreamJson

```csharp
LazyLoadingStreamJson(stream, sheetName, uid)
```

Writes the JSON string of the specified sheet in the file from the cache using the specified unique id  to the stream..

**Parameters**:

- `stream`: The stream that will be written
- `sheetName`: The sheet name.
- `uid`: The unique id for the file cache.

#### AddImage

```csharp
AddImage(p, uid, iscontrol, files)
```

Applies the add image from local file operation.

**Parameters**:

- `p`: The JSON string parameter
- `uid`: The unique id for the file cache.
- `iscontrol`: Specify whether it is a control.
- `files`: The form file of the image

**Returns**: The JSON string result

#### AddImageByURL

```csharp
AddImageByURL(p, uid, imageurl)
```

Applies the add image from remote URL operation.

**Parameters**:

- `p`: The JSON string parameter
- `uid`: The unique id for the file cache.
- `imageurl`: Specify the image URL.

**Returns**: The JSON string result

#### CopyImage

```csharp
CopyImage(p, uid)
```

Applies the copy image operation.

**Parameters**:

- `p`: The JSON string parameter
- `uid`: The unique id for the file cache.

**Returns**: The JSON string result

#### Load

```csharp
Load(uid, filename)
```

Gets the JSON  string  of the file from the cache using the specified unique id,set the output filename in the JSON.

**Parameters**:

- `uid`: The unique id for the file cache.
- `filename`: Specifies the file name in the JSON. If set to null,the default filename is: book1.

**Returns**: The JSON string

#### Image

```csharp
Image(uid, picid)
```

Get Stream of image.

**Parameters**:

- `uid`: The unique id for the file cache.
- `picid`: The image id.

**Returns**: The image stream

#### Ole

```csharp
Ole(uid, sheetname, oleid, label)
```

Gets the byte array data of the  embedded ole object .

**Parameters**:

- `uid`: The unique id for the file cache.
- `sheetname`: The worksheet name.
- `oleid`: The  id for the embedded ole object.
- `label`: The display label of the embedded ole object.

**Returns**: The byte array data of the  embedded ole object .

#### ImageUrl

```csharp
ImageUrl(baseURL, picid, uid)
```

Gets the image URL.

**Parameters**:

- `baseURL`: The base action URL.
- `picid`: The image id.
- `uid`: The unique id for the file cache.

**Returns**: The image URL

#### GetFile

```csharp
GetFile(fileid)
```

Get file stream

**Parameters**:

- `fileid`: the file id

**Returns**: The stream of the file

#### Download

```csharp
Download(p, uid, filename)
```

Applies the download file operation

**Parameters**:

- `p`: The JSON parameter
- `uid`: The unique id for the file cache.
- `filename`: The file name

**Returns**: The file URL

#### GetWarningCallback

```csharp
GetWarningCallback()
```

Gets custom warning callback for import file.

**Returns**: The warning callback

#### SetWarningCallback

```csharp
SetWarningCallback(callback)
```

Sets custom warning callback for import file.

**Parameters**:

- `callback`: The warning callback

---

## ITextTranslator

**Kind**: Interface

**Namespace**: `Aspose.Cells.GridJs`

**Description**: Represents the interface for translate

### Methods

#### TranslateAsync

```csharp
TranslateAsync(texts, targetLanguage)
```

translate texts to targetLanguage

**Parameters**:

- `texts`: string list
- `targetLanguage`: target language

---

## NamespaceDoc

**Kind**: Struct

**Namespace**: `Aspose.Cells.GridJs.Chart`

**Description**: The Chart namespace encapsulates all classes of GridJs , providing  basic data structure for Charts JSON generation.

---

## NamespaceDoc

**Kind**: Class

**Namespace**: `Aspose.Cells.GridJs`

**Description**: The Aspose.Cells.GridJs namespace encapsulates all classes of GridJs, providing simple APIs for viewing or editing spreadsheet files through JSON operations.

---

## OprMessageService

**Kind**: Class

**Namespace**: `Aspose.Cells.GridJs`

**Description**: This class provide all the operations for messages sync in Collaborative mode .

---

# Enumerations

## CoWorkOperationType

Represents the action operation type in collabration mode.only available in java version now, will be available in .net/python version in future.

| Member | Description |
|--------|-------------|
| `VIEW` | View operation (read-only) |
| `LOAD` | Load operation (read-only) |
| `LOAD_SHEET` | Load sheet operation (read-only) |
| `GET_IMAGE` | Get image operation (read-only) |
| `GET_OLE` | Get OLE object operation (read-only) |
| `EDIT` | Edit operation (editable) |
| `UPDATE_CELL` | Update cell operation (editable) |
| `ADD_IMAGE` | Add image operation (editable) |
| `COPY_IMAGE` | Copy image operation (editable) |
| `INSERT_ROW` | Insert row operation (editable) |
| `DELETE_ROW` | Delete row operation (editable) |
| `INSERT_COLUMN` | Insert column operation (editable) |
| `DELETE_COLUMN` | Delete column operation (editable) |
| `DOWNLOAD` | Download operation |
| `ADMIN_CONFIG` | Admin configuration operation (administrative) |
| `USER_MANAGEMENT` | User management operation (administrative) |

