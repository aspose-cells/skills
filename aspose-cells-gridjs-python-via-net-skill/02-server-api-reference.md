# Server-Side API Reference

Assembly: **aspose-cells-gridjs**

## Config

**Kind**: Class


**Description**: Represents all the static settings for GridJs

### Properties & Fields

| Name | Type | Description |
|------|------|-------------|
| `save_html_as_zip` | Property | Gets/Sets whether to save html file as zip archive,the default is false. |
| `skip_invisible_shapes` | Property | Gets/Sets whether to skip shapes that are invisble to UI ,the default value is true. |
| `lazy_loading` | Property | Gets/Sets whether to load active worksheet only,the default is false. |
| `same_image_detecting` | Property | Gets/Sets whether to check if images have same source,the default is true the default value is true. |
| `auto_optimize_for_large_cells` | Property | Gets/Sets whether to automatically optimize the load performance for worksheet with large cells. it will ignore some style /borders to reduce the load time. the default value is true. |
| `islimit_shape_or_image` | Property | Gets/Sets whether to limit the total display shape/image count in one worksheet ,if set to true, GridJs will limit the total count of the display shapes or images in one worksheet to MaxShapeOrImageCount the default value is true. |
| `max_shape_or_image_count` | Property | Gets/Sets the total count of the display shapes or images in the active sheet,it takes effect when IslimitShapeOrImage=true. the default value is 100. |
| `max_total_shape_or_image_count` | Property | Gets/Sets the total count of the display shapes or images in the workbook,it takes effect when IslimitShapeOrImage=true. the default value is 300. |
| `max_shape_or_image_width_or_height` | Property | Gets/Sets the max width or height for a shape or an image ,GridJs will ignore the shape or image with the width or height larger than this, it takes effect when IslimitShapeOrImage=true. the default value is 10000. |
| `max_pdf_save_seconds` | Property | Gets/Sets the max timed out seconds when save to PDF. the default value is 10. |
| `ignore_empty_content` | Property | Gets/Sets whether to show the max range which includes data ,style, merged cells and shapes. if the last row or column contains cells with no value and formula but has custom style then we will not show this row/column when this vlaue is true。 the default value is true . |
| `use_print_area` | Property | Gets/Sets whether to use PageSetup.PrintArea for the UI display range when the worksheet has PageSetup setting for PrintArea. the default value is false . |
| `is_collaborative` | Property | Gets/Sets whether to support collabrative editing,the default is false. |
| `enable_chart_client_rendering` | Property | Gets/Sets whether to enable client-side chart rendering. If set to false, GridJs will fall back to server-generated chart images/shapes only. the default value is true. |
| `custom_pdf_save_options` | Property | Gets/Sets the custom PdfSaveOptions for PDF export. If set, this will be used instead of the default options. the default value is null. |
| `load_time_out` | Property | Gets/Sets a timeout interrupt in milliseconds in loading file, when the cost time period of loading file is longer than the expectation ，it will raise exception. the default value is -1,which means no timeout interrupt is set . |
| `show_chart_sheet` | Property | Gets/Sets whether to show chart worksheet. the default value is false . |
| `empty_sheet_max_row` | Property | Gets/Sets default max row for an empty worksheet. the default value is 12. |
| `empty_sheet_max_col` | Property | Gets/Sets default max column for an empty worksheet. the default value is 15. |
| `picture_cache_directory` | Property | Gets/Sets the cache directory for pictures.(this takes effect when GridJsWorkbook.CacheImp is null) the default path will be "_piccache" inside the FileCacheDirectory. |
| `file_cache_directory` | Property | Gets/Sets the cache directory for storing spreadsheet file. We need to set it to a specific path before we use GridJs. |
| `base_route_name` | Property | Gets/Sets the base route name for GridJs controller URL. the default is "/GridJs2". |
| `message_topic` | Property | Gets/Sets the websocket destinations prefixed with "/topic". the default is "/topic/opr".used in collaborative mode only. |
| `auto_fit_rows_height_on_load` | Property | Indicates whether to auto-fit row heights during file loading. The default value is ``false``. Warning: Setting this to ``true`` will perform an auto-fit all rows operation post-load, which may have a noticeable impact on performance. |
| `auto_fit_columns_on_load` | Property | Indicates whether to auto-fit column widths during file loading. The default value is ``false``. Warning: Setting this to ``true`` will perform an auto-fit all columns operation post-load, which may have a noticeable impact on performance. |
| `redaction_use_client_generate_id` | Property | Indicates whether use client generate id instead of actual shape id in redaction related operations |

### Methods

#### __init__

```python
__init__()
```

Config __init__.

#### set_license

```python
set_license(license_name)
```

Licenses the component. Use an empty string to switch to evaluation mode.

**Parameters**:

- `license_name`: Can be a full or short file name or name of an embedded resource.

#### set_font_folder

```python
set_font_folder(font_folder, recursive)
```

Sets the fonts folder

**Parameters**:

- `font_folder`: The folder that contains TrueType fonts.
- `recursive`: Determines whether or not to scan subfolders.

#### set_font_folders

```python
set_font_folders(font_folders, recursive)
```

Sets the fonts folders

**Parameters**:

- `font_folders`: The folders that contains TrueType fonts.
- `recursive`: Determines whether or not to scan subfolders.

#### set_save_html_as_zip

```python
set_save_html_as_zip()
```

Gets/Sets whether to save html file as zip archive,the default is false.

#### set_skip_invisible_shapes

```python
set_skip_invisible_shapes()
```

Gets/Sets whether to skip shapes that are invisble to UI ,the default value is true.

#### set_lazy_loading

```python
set_lazy_loading()
```

Gets/Sets whether to load active worksheet only,the default is false.

#### set_same_image_detecting

```python
set_same_image_detecting()
```

Gets/Sets whether to check if images have same source,the default is true the default value is true.

#### set_auto_optimize_for_large_cells

```python
set_auto_optimize_for_large_cells()
```

Gets/Sets whether to automatically optimize the load performance for worksheet with large cells. it will ignore some style /borders to reduce the load time. the default value is true.

#### set_islimit_shape_or_image

```python
set_islimit_shape_or_image()
```

Gets/Sets whether to limit the total display shape/image count in one worksheet ,if set to true, GridJs will limit the total count of the display shapes or images in one worksheet to MaxShapeOrImageCount the default value is true.

#### set_max_shape_or_image_count

```python
set_max_shape_or_image_count()
```

Gets/Sets the total count of the display shapes or images in the active sheet,it takes effect when IslimitShapeOrImage=true. the default value is 100.

#### set_max_total_shape_or_image_count

```python
set_max_total_shape_or_image_count()
```

Gets/Sets the total count of the display shapes or images in the workbook,it takes effect when IslimitShapeOrImage=true. the default value is 300.

#### set_max_shape_or_image_width_or_height

```python
set_max_shape_or_image_width_or_height()
```

Gets/Sets the max width or height for a shape or an image ,GridJs will ignore the shape or image with the width or height larger than this, it takes effect when IslimitShapeOrImage=true. the default value is 10000.

#### set_max_pdf_save_seconds

```python
set_max_pdf_save_seconds()
```

Gets/Sets the max timed out seconds when save to PDF. the default value is 10.

#### set_ignore_empty_content

```python
set_ignore_empty_content()
```

Gets/Sets whether to show the max range which includes data ,style, merged cells and shapes. if the last row or column contains cells with no value and formula but has custom style then we will not show this row/column when this vlaue is true。 the default value is true .

#### set_use_print_area

```python
set_use_print_area()
```

Gets/Sets whether to use PageSetup.PrintArea for the UI display range when the worksheet has PageSetup setting for PrintArea. the default value is false .

#### set_is_collaborative

```python
set_is_collaborative()
```

Gets/Sets whether to support collabrative editing,the default is false.

#### set_enable_chart_client_rendering

```python
set_enable_chart_client_rendering()
```

Gets/Sets whether to enable client-side chart rendering. If set to false, GridJs will fall back to server-generated chart images/shapes only. the default value is true.

#### set_custom_pdf_save_options

```python
set_custom_pdf_save_options()
```

Gets/Sets the custom PdfSaveOptions for PDF export. If set, this will be used instead of the default options. the default value is null.

#### set_load_time_out

```python
set_load_time_out()
```

Gets/Sets a timeout interrupt in milliseconds in loading file, when the cost time period of loading file is longer than the expectation ，it will raise exception. the default value is -1,which means no timeout interrupt is set .

#### set_show_chart_sheet

```python
set_show_chart_sheet()
```

Gets/Sets whether to show chart worksheet. the default value is false .

#### set_empty_sheet_max_row

```python
set_empty_sheet_max_row()
```

Gets/Sets default max row for an empty worksheet. the default value is 12.

#### set_empty_sheet_max_col

```python
set_empty_sheet_max_col()
```

Gets/Sets default max column for an empty worksheet. the default value is 15.

#### set_picture_cache_directory

```python
set_picture_cache_directory()
```

Gets/Sets the cache directory for pictures.(this takes effect when GridJsWorkbook.CacheImp is null) the default path will be "_piccache" inside the FileCacheDirectory.

#### set_file_cache_directory

```python
set_file_cache_directory()
```

Gets/Sets the cache directory for storing spreadsheet file. We need to set it to a specific path before we use GridJs.

#### set_base_route_name

```python
set_base_route_name()
```

Gets/Sets the base route name for GridJs controller URL. the default is "/GridJs2".

#### set_message_topic

```python
set_message_topic()
```

Gets/Sets the websocket destinations prefixed with "/topic". the default is "/topic/opr".used in collaborative mode only.

#### set_auto_fit_rows_height_on_load

```python
set_auto_fit_rows_height_on_load()
```

Indicates whether to auto-fit row heights during file loading. The default value is ``false``. Warning: Setting this to ``true`` will perform an auto-fit all rows operation post-load, which may have a noticeable impact on performance.

#### set_auto_fit_columns_on_load

```python
set_auto_fit_columns_on_load()
```

Indicates whether to auto-fit column widths during file loading. The default value is ``false``. Warning: Setting this to ``true`` will perform an auto-fit all columns operation post-load, which may have a noticeable impact on performance.

#### set_redaction_use_client_generate_id

```python
set_redaction_use_client_generate_id()
```

Indicates whether use client generate id instead of actual shape id in redaction related operations

---

## CoWorkUserProvider

**Kind**: Interface


**Description**: Represents the user provider inerface in collabration mode.only available in java version now, will be available in .net/python version in future. Customer application can implement this interface to provide the user information.

### Methods

#### get_current_user_name

```python
get_current_user_name()
```

Gets the username of the current user

**Returns**: Current username

#### get_current_user_id

```python
get_current_user_id()
```

Gets the unique identifier of the current user

**Returns**: Current user ID

#### get_permission

```python
get_permission()
```

Gets the permission level of the current user

**Returns**: Current user permission level

---

## GridAbstractCalculationEngine

**Kind**: Class


**Description**: Represents user\'s custom calculation engine to extend the default calculation engine of Aspose.Cells.

### Methods

#### calculate

```python
calculate(data)
```

Calculates one function with given data.

**Parameters**:

- `data`: The required data to calculate function such as function name, parameters, ...etc.

---

## GridCacheForStream

**Kind**: Class


**Description**: This class contains the cache operations for GridJs. User shall implement his own business logic for storage based on it..

### Methods

#### save_stream

```python
save_stream()
```

Implements this method to save cache,save the stream to the cache with the key uid.

#### load_stream

```python
load_stream()
```

Implements this method to load cache with the key uid,return the stream from the cache.

**Returns**: Returns io._IOBase.

#### is_existed

```python
is_existed(uid)
```

Checks whether the cache with uid is existed

**Parameters**:

- `uid`: The unique id for the file cache.

**Returns**: The bool value

#### get_file_url

```python
get_file_url(uid)
```

Implements this method to get the file url from the cache.

**Parameters**:

- `uid`: The unique id for the file cache.

**Returns**: The URL of the file

---

## GridCalculationData

**Kind**: Class


**Description**: Represents the required data when calculating one function, such as function name, parameters, ...etc.

### Properties & Fields

| Name | Type | Description |
|------|------|-------------|
| `calculated_value` | Property | Gets/sets the calculated value for this function. |
| `row` | Property | Gets the Cell Row index where the function is in. |
| `column` | Property | Gets the Cell Column index where the function is in. |
| `string_value` | Property | Gets the Cell DisplayStringValue where the function is in. |
| `value` | Property | Gets the Cell value where the function is in. |
| `formula` | Property | Gets the Cell formula where the function is in. |
| `sheet_name` | Property | Gets the worksheet name where the function is in. |
| `function_name` | Property | Gets the function name to be calculated. |
| `param_count` | Property | Gets the count of parameters . |

### Methods

#### get_param_value

```python
get_param_value(index)
```

Gets the represented value object of the parameter at given index. If the parameter is reference, then return ReferredArea object.

**Parameters**:

- `index`: The index of the parameter(0 based).

**Returns**: If the parameter is plain value, then returns the plain value.

#### get_param_text

```python
get_param_text(index)
```

Gets the literal text of the parameter at given index.

**Parameters**:

- `index`: The index of the parameter(0 based).

**Returns**: The literal text of the parameter.

---

## GridCellException

**Kind**: Class


**Description**: The exception that is thrown when GridJs specified error occurs.

### Properties & Fields

| Name | Type | Description |
|------|------|-------------|
| `code` | Property | Represents the exception code. |

### Methods

#### to_string

```python
to_string()
```

Creates and returns a string representation of the current exception.

**Returns**: Returns str.

---

## GridJsControllerBase

**Kind**: Class


**Description**: Represents the GridJsControllerBase type.

### Methods

#### __init__

```python
__init__()
```

GridJsControllerBase __init__.

---

## GridJsLogger

**Kind**: Class


**Description**: Provides a lightweight logging system for GridJs server. Supports console and file output with configurable log levels. Default level is None. In DEBUG build, the default level is Debug (outputs all messages). Supports log file rolling by day/week/month.

### Properties & Fields

| Name | Type | Description |
|------|------|-------------|
| `level` | Property | Sets the minimum log level. Messages below this level will be ignored. Default is None. |
| `enable_console_output` | Property | Sets whether to output log messages to the console. Default is false. |
| `enable_file_output` | Property | Sets whether to output log messages to a file. Default is false. Must call :py:func:`aspose.cellsgridjs.GridJsLogger.set_log_file` or :py:func:`aspose.cellsgridjs.GridJsLogger.set_log_directory` before enabling. |
| `log_file_path` | Property | Gets the current log file path, or null if not set. |
| `max_file_size` | Property | Sets the maximum size (in bytes) for a single log file before rolling. Only effective when rolling is enabled. 0 means no size limit. Default is 0. |

### Methods

#### set_log_file

```python
set_log_file(file_path)
```

Sets the log file path and enables file output (single file mode, no rolling). Appends to existing file.

**Parameters**:

- `file_path`: The full path to the log file.

#### set_log_directory

```python
set_log_directory(directory)
```

Sets the log directory with daily rolling file strategy and enables file output. Log files will be named as: gridjs_{period}.log

**Parameters**:

- `directory`: The directory to store log files.

#### close_log_file

```python
close_log_file()
```

Closes the log file writer and disables file output.

#### debug

```python
debug()
```

Logs a debug level message.

#### info

```python
info()
```

Logs an info level message.

#### warn

```python
warn()
```

Logs a warning level message.

#### error

```python
error()
```

Logs an error level message.

#### log

```python
log(level, message)
```

Logs a message at the specified level.

**Parameters**:

- `level`: The log level.
- `message`: The message to log.

#### set_level

```python
set_level()
```

Gets the minimum log level. Messages below this level will be ignored. Default is None.

#### set_enable_console_output

```python
set_enable_console_output()
```

Gets whether to output log messages to the console. Default is false.

#### set_enable_file_output

```python
set_enable_file_output()
```

Gets whether to output log messages to a file. Default is false. Must call :py:func:`aspose.cellsgridjs.GridJsLogger.set_log_file` or :py:func:`aspose.cellsgridjs.GridJsLogger.set_log_directory` before enabling.

#### set_max_file_size

```python
set_max_file_size()
```

Gets the maximum size (in bytes) for a single log file before rolling. Only effective when rolling is enabled. 0 means no size limit. Default is 0.

---

## GridJsOptions

**Kind**: Class


**Description**: Represents all the load options for GridJs

### Properties & Fields

| Name | Type | Description |
|------|------|-------------|
| `save_html_as_zip` | Property | Gets/Sets whether to save html file as zip archive,the default is false. |
| `skip_invisible_shapes` | Property | Gets/Sets whether to skip shapes that are invisble to UI ,the default value is true. |
| `lazy_loading` | Property | Gets/Sets whether to load active worksheet only,the default is false. |
| `same_image_detecting` | Property | Gets/Sets whether to check if images have same source,the default is true the default value is true. |
| `auto_optimize_for_large_cells` | Property | Gets/Sets whether to automatically optimize the load performance for worksheet with large cells. it will ignore some style /borders to reduce the load time. the default value is true. |
| `islimit_shape_or_image` | Property | Gets/Sets whether to limit the total display shape/image count in one worksheet ,if set to true, GridJs will limit the total count of the display shapes or images in one worksheet to MaxShapeOrImageCount the default value is true. |
| `max_shape_or_image_count` | Property | Gets/Sets the total count of the display shapes or images in the active sheet,it takes effect when IslimitShapeOrImage=true. the default value is 100. |
| `max_total_shape_or_image_count` | Property | Gets/Sets the total count of the display shapes or images in the workbook,it takes effect when IslimitShapeOrImage=true. the default value is 300. |
| `max_shape_or_image_width_or_height` | Property | Gets/Sets the max width or height for a shape or an image ,GridJs will ignore the shape or image with the width or height larger than this, it takes effect when IslimitShapeOrImage=true. the default value is 10000. |
| `max_pdf_save_seconds` | Property | Gets/Sets the max timed out seconds when save to PDF. the default value is 10. |
| `ignore_empty_content` | Property | Gets/Sets whether to show the max range which includes data ,style, merged cells and shapes. if the last row or column contains cells with no value and formula but has custom style then we will not show this row/column when this vlaue is true。 the default value is true . |
| `use_print_area` | Property | Gets/Sets whether to use PageSetup.PrintArea for the UI display range when the worksheet has PageSetup setting for PrintArea. the default value is false . |
| `is_collaborative` | Property | Gets/Sets whether to support collabrative editing,the default is false. |
| `enable_chart_client_rendering` | Property | Gets/Sets whether to enable client-side chart rendering. If set to false, GridJs will emit chart images/shapes only and will not require client-side chart rendering. the default value is true. |
| `custom_pdf_save_options` | Property | Gets/Sets the custom PdfSaveOptions for PDF export. If set, this will be used instead of the default options. the default value is null. |
| `load_time_out` | Property | Gets/Sets a timeout interrupt in milliseconds in loading file, when the cost time period of loading file is longer than the expectation ，it will raise exception. the default value is -1,which means no timeout interrupt is set . |
| `show_chart_sheet` | Property | Gets/Sets whether to show chart worksheet. the default value is false . |
| `empty_sheet_max_row` | Property | Gets/Sets default max row for an empty worksheet. the default value is 12. |
| `empty_sheet_max_col` | Property | Gets/Sets default max column for an empty worksheet. the default value is 15. |
| `picture_cache_directory` | Property | Gets/Sets the cache directory for pictures.(this takes effect when GridJsWorkbook.CacheImp is null) the default path will be "_piccache" inside the FileCacheDirectory. |
| `file_cache_directory` | Property | Gets/Sets the cache directory for storing spreadsheet file. We need to set it to a specific path before we use GridJs. |
| `font_folders` | Property | Gets/Sets the fonts folders for fonts in the rendered pictures/shapes |
| `base_route_name` | Property | Gets/Sets the route URL base name for GridJs controller.the default is GridJs2 |
| `message_topic` | Property | Gets/Sets the websocket destinations prefixed with "/topic". the default is "/topic/opr".used in collaborative mode only. |
| `auto_fit_rows_height_on_load` | Property | Indicates whether to autofit rows height when loading the file,the default value is false. |
| `auto_fit_columns_on_load` | Property | Indicates whether to autofit columns width when loading the file,the default value is false. |
| `cache_imp` | Property | Custom implemention for cache storage,If you want to store cache in stream way ,you need to set and implement it. |
| `redaction_use_client_generate_id` | Property | Indicates whether use client generate id instead of actual shape id in redaction related operations |

### Methods

#### __init__

```python
__init__()
```

GridJsOptions __init__.

---

## GridJsPermissionException

**Kind**: Class


**Description**: represents permission exception in collaboration mode.only available in java version now, will be available in .net/python version in future.

### Methods

#### __init__

```python
__init__(operation, required_permission, current_permission)
```

Constructs a permission exception

**Parameters**:

- `operation`: Operation attempted to execute
- `required_permission`: Permission required for the operation
- `current_permission`: Current user permission

#### get_required_permission

```python
get_required_permission()
```

Gets the permission level required for the operation

**Returns**: Required permission level

#### get_current_permission

```python
get_current_permission()
```

Gets the current permission level of the user

**Returns**: Current user permission level

#### get_operation

```python
get_operation()
```

Gets the operation that was attempted to execute

**Returns**: Operation name

---

## GridJsService

**Kind**: Class


**Description**: Provides the basic operation apis used in controller actions.

### Properties & Fields

| Name | Type | Description |
|------|------|-------------|
| `settings` | Property | Represents the workbook settings. |

### Methods

#### __init__

```python
__init__()
```

The default constructor for GridJsService

#### get_warning_callback

```python
get_warning_callback()
```

Gets custom warning callback for import file.

**Returns**: The warning callback

#### set_warning_callback

```python
set_warning_callback(callback)
```

Sets custom warning callback for import file.

**Parameters**:

- `callback`: The warning callback

#### check_in_cache_for_collaborative

```python
check_in_cache_for_collaborative()
```

Check wether workbook instance is in memory cache .this method is apply for Collaborative mode only.

**Returns**: Returns bool.

#### update_cell

```python
update_cell(p, uid)
```

Applies the update operation.

**Parameters**:

- `p`: The JSON format string of update operation.
- `uid`: The unique id for the file cache.

**Returns**: The JSON format string of the update result.

#### detail_stream_json_with_uid

```python
detail_stream_json_with_uid(stream, file_path, uid)
```

Write the JSON string for the file to the stream by the specified unique id.

**Parameters**:

- `stream`: The stream that will be written
- `file_path`: The file path
- `uid`: The unique id for the file cache.

#### detail_stream_json

```python
detail_stream_json(stream, file_path)
```

Write the JSON string for the file to the stream .

**Parameters**:

- `stream`: The stream that will be written
- `file_path`: The file path

#### lazy_loading_stream_json

```python
lazy_loading_stream_json(stream, sheet_name, uid)
```

Writes the JSON string of the specified sheet in the file from the cache using the specified unique id to the stream..

**Parameters**:

- `stream`: The stream that will be written
- `sheet_name`: The sheet name.
- `uid`: The unique id for the file cache.

#### add_image

```python
add_image(p, uid, iscontrol, file_input_stream)
```

Applies the add image from local file operation.

**Parameters**:

- `p`: The JSON string parameter
- `uid`: The unique id for the file cache.
- `iscontrol`: Specify whether it is a control.
- `file_input_stream`: The input stream form file of the image

**Returns**: The JSON string result

#### add_image_by_url

```python
add_image_by_url(p, uid, imageurl)
```

Applies the add image from remote URL operation.

**Parameters**:

- `p`: The JSON string parameter
- `uid`: The unique id for the file cache.
- `imageurl`: Specify the image URL.

**Returns**: The JSON string result

#### copy_image

```python
copy_image(p, uid)
```

Applies the copy image operation.

**Parameters**:

- `p`: The JSON string parameter
- `uid`: The unique id for the file cache.

**Returns**: The JSON string result

#### load

```python
load(uid, filename)
```

Gets the JSON string of the file from the cache using the specified unique id,set the output filename in the JSON.

**Parameters**:

- `uid`: The unique id for the file cache.
- `filename`: Specifies the file name in the JSON. If set to null,the default filename is: book1.

**Returns**: The JSON string

#### image

```python
image(uid, picid)
```

Get Stream of image.

**Parameters**:

- `uid`: The unique id for the file cache.
- `picid`: The image id.

**Returns**: The image stream

#### ole

```python
ole(uid, sheetname, oleid, label)
```

Gets the byte array data of the embedded ole object .

**Parameters**:

- `uid`: The unique id for the file cache.
- `sheetname`: The worksheet name.
- `oleid`: The  id for the embedded ole object.
- `label`: The display label of the embedded ole object.

**Returns**: The byte array data of the  embedded ole object .

#### image_url

```python
image_url(base_url, picid, uid)
```

Gets the image URL.

**Parameters**:

- `base_url`: The base action URL.
- `picid`: The image id.
- `uid`: The unique id for the file cache.

**Returns**: The image URL

#### get_file

```python
get_file(fileid)
```

Get file stream

**Parameters**:

- `fileid`: the file id

**Returns**: The stream of the file

#### download

```python
download(p, uid, filename)
```

Applies the download file operation

**Parameters**:

- `p`: The JSON parameter
- `uid`: The unique id for the file cache.
- `filename`: The file name

**Returns**: The file URL

#### destroy

```python
destroy()
```

this method is for java only ,it implement the destroy method in DisposableBean ,which will be called automatically by Spring when finish request, actually it will call Dispose method

---

## GridJsWorkbook

**Kind**: Class


**Description**: Represents the main entry class for GridJs

### Properties & Fields

| Name | Type | Description |
|------|------|-------------|
| `settings` | Property | Represents the workbook settings. |
| `warning_callback` | Property | Gets warning callback for import file. |
| `cache_directory` | Property | Gets/Sets the cache directory for storing spreadsheet file. if this is not set ,we will use the static propertie from :py:attr:`aspose.cellsgridjs.Config.file_cache_directory` or :py:attr:`aspose.cellsgridjs.GridJsWorkbook.cache_imp` . |
| `cache_imp` | Property | Custom implemention for cache storage,If you want to store cache in stream way ,you need to set and implement it. |
| `calculate_engine` | Property | Custom implemention for calculation engine ,If you want to do custom calculation, you need to set and implement it. |
| `update_monitor` | Property | Gets/Sets the update monitor to track update operation |
| `PICTURE_TYPE` | Property | const value for the type of the image |

### Methods

#### __init__

```python
__init__()
```

Creates a new instance of GridJsWorkbook.

#### import_excel_file

```python
import_excel_file(uid, file_name, password)
```

Imports the excel file from file path and open password.

**Parameters**:

- `uid`: The unique id for the file cache, if set to null,it will be generated automatically.
- `file_name`: The full path of the file.
- `password`: The open password  of the excel file.The value can be null If no passowrd is set.

#### export_to_json

```python
export_to_json(filename)
```

Gets JSON string from memory data,set the output filename in the JSON.

**Parameters**:

- `filename`: Specifies the file name in the JSON. If set to null,the default filename is: book1..

**Returns**: The JSON string.

#### save_to_excel_file

```python
save_to_excel_file(stream)
```

Saves the memory data to the sream, baseed on the origin file format.

**Parameters**:

- `stream`: The stream to save.

#### save_to_pdf

```python
save_to_pdf(path)
```

Saves the memory data to the file path,the save format is pdf.

**Parameters**:

- `path`: The file path to save.

#### save_to_xlsx

```python
save_to_xlsx(path)
```

Saves the memory data to the file path,the save format is xlsx.

**Parameters**:

- `path`: The file path to save.

#### save_to_html

```python
save_to_html(path)
```

Saves the memory data to the file path,the save format is html.

**Parameters**:

- `path`: The file path to save.

#### json_to_stream_by_uid

```python
json_to_stream_by_uid(stream, uid, filename)
```

Retrieve the JSON string of the file from the cache using the specified unique id,set the output filename in the JSON,and write it to the stream.

**Parameters**:

- `stream`: The stream that will be written
- `uid`: The unique id for the file cache.
- `filename`: Specifies the file name in the JSON. If set to null,the default filename is: book1.

**Returns**: Returns bool.

#### json_to_stream

```python
json_to_stream(stream, filename)
```

Retrieve the JSON string from memory data,set the output filename in the JSON, and write it to the stream.

**Parameters**:

- `stream`: The stream that will be written
- `filename`: Specifies the file name in the JSON. If set to null,the default filename is: book1.

#### lazy_loading_stream

```python
lazy_loading_stream(stream, uid, sheet_name)
```

Retrieve the JSON string of the specified sheet in the file from the cache using the specified unique id, and write it to the stream.

**Parameters**:

- `stream`: The stream that will be written
- `uid`: The unique id for the file cache.
- `sheet_name`: the sheet name.

#### get_json_str_by_uid

```python
get_json_str_by_uid(uid, filename)
```

Gets the JSON string of the file from the cache using the specified unique id,set the output filename in the JSON.

**Parameters**:

- `uid`: The unique id for the file cache.
- `filename`: Specifies the file name in the JSON. If set to null,the default filename is: book1.

**Returns**: The JSON  string

#### lazy_loading_json_str

```python
lazy_loading_json_str(uid, sheet_name)
```

Gets the JSON string of the specified sheet in the file from the cache using the specified unique id.

**Parameters**:

- `uid`: The unique id for the file cache.
- `sheet_name`: the sheet name .

**Returns**: The JSON string

#### get_uid_for_file

```python
get_uid_for_file(file_name)
```

Generates a new unique id for the file cache using the given file name.

**Parameters**:

- `file_name`: The file name.

**Returns**: Returns str.

#### import_excel_file_from_json

```python
import_excel_file_from_json(json)
```

Imports the excel file from JSON format string.

**Parameters**:

- `json`: The JSON format string.

#### merge_excel_file_from_json

```python
merge_excel_file_from_json(uid, json)
```

Applies a batch update to the memory data.

**Parameters**:

- `uid`: The unique id for the file cache.
- `json`: The update JSON format string.

#### save_to_cache_with_file_name

```python
save_to_cache_with_file_name(uid, filename, password)
```

Saves the memory data to the cache file with the specified filename and also set the open password, the save format is baseed on the file extension of the filename .

**Parameters**:

- `uid`: The unique id for the file cache.
- `filename`: The filename to save.
- `password`: The excel file\'s open password. The value can be null If no passowrd is set.

#### get_image_stream

```python
get_image_stream(uid, picid)
```

Get Stream of image.

**Parameters**:

- `uid`: The unique id for the file cache.
- `picid`: The image id.

**Returns**: The image stream

#### get_ole

```python
get_ole(uid, sheetname, oleid, label)
```

Gets the byte array data of the embedded ole object .

**Parameters**:

- `uid`: The unique id for the file cache.
- `sheetname`: The worksheet name.
- `oleid`: The  id for the embedded ole object.
- `label`: The display label of the embedded ole object.

**Returns**: The byte array data of the  embedded ole object .

#### check_in_cache_for_collaborative

```python
check_in_cache_for_collaborative()
```

Check wether workbook instance is in memory cache .this method is apply for Collaborative mode only.

**Returns**: Returns bool.

#### update_cell

```python
update_cell(p, uid)
```

Applies the update operation.

**Parameters**:

- `p`: The JSON format string of update operation.
- `uid`: The unique id for the file cache.

**Returns**: The JSON format string of the update result.

#### redact_file

```python
redact_file(excel_file_path, uid, array_of_redaction_opr)
```

Performs redaction on an Excel file based on an array of JSON operations.

**Parameters**:

- `excel_file_path`: The file path of the Excel file to be redacted.
- `uid`: The unique identifier for the workbook. If null or empty, a new uid will be generated based on the file path.
- `array_of_redaction_opr`: An array of JSON strings representing the redaction operations to be applied.

#### set_trans_parent_view

```python
set_trans_parent_view(excel_file_path, uid, is_transparent)
```

Sets the transparency of redaction shapes in the workbook.

**Parameters**:

- `excel_file_path`: The file path of the Excel file.
- `uid`: The unique identifier for the workbook. If null or empty, a new uid will be generated based on the file path.
- `is_transparent`: If true, sets transparency to 0.89 (semi-transparent); if false, sets transparency to 1 (fully opaque/invisible).

#### burn_redaction_file

```python
burn_redaction_file(excel_file_path, uid)
```

Burns (applies) all redaction operations in the workbook by removing redaction shapes and their target shapes or clearing target cell range contents.

**Parameters**:

- `excel_file_path`: The file path of the Excel file.
- `uid`: The unique identifier for the workbook. If null or empty, a new uid will be generated based on the file path.

#### dispose

```python
dispose()
```

Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.

#### insert_image

```python
insert_image(uid, p, s, image_url)
```

Inserts image in the worksheet from file stream or the URL,(either the file stream or the URL shall be provided) or Inserts shape ,when the p.type is one of AutoShapeType

**Parameters**:

- `uid`: The unique id for the file cache
- `p`: The JSON format string for the operation which specify the cell location  ,the worksheet name,upper left row,upper left column for the image，etc  {name:\'sheet1\',ri:1,ci:1}
- `s`: The file stream of the image file
- `image_url`: The URL of the image file

**Returns**: The JSON format string of the inserted image

#### copy_image_or_shape

```python
copy_image_or_shape(uid, p)
```

Copys image or shape.

**Parameters**:

- `uid`: The unique id for the file cache.
- `p`: The JSON string for the operation which specify the cell location ,it contains the worksheet name,upper left row,upper left column for the image or shape，etc  {name:\'sheet1\',ri:1,ci:1,srcid:2,srcname:\'sheet2\',isshape:true}

**Returns**: The JSON string of the new copied image

#### error_json

```python
error_json(msg)
```

Gets the error message string in JSON format.

**Parameters**:

- `msg`: The error message.

**Returns**: The JSON string.

#### get_grid_load_format

```python
get_grid_load_format(extension)
```

Gets the load format by file extension

**Parameters**:

- `extension`: The file extention ,usually start with \'.\' .

**Returns**: Returns GridLoadFormat.

#### get_image_url

```python
get_image_url(uid, picid, delimiter)
```

Gets the image URL.

**Parameters**:

- `uid`: The unique id for the file cache.
- `picid`: The image id.
- `delimiter`: The string delimiter.

**Returns**: Returns str.

#### set_image_url_base

```python
set_image_url_base(base_image_url)
```

Set the base image get action URL from controller .

**Parameters**:

- `base_image_url`: the base image get action URL.

#### set_cache_imp

```python
set_cache_imp()
```

Custom implemention for cache storage,If you want to store cache in stream way ,you need to set and implement it.

#### set_calculate_engine

```python
set_calculate_engine()
```

Custom implemention for calculation engine ,If you want to do custom calculation, you need to set and implement it.

#### set_update_monitor

```python
set_update_monitor()
```

Gets/Sets the update monitor to track update operation

---

## GridReferredArea

**Kind**: Class


**Description**: Represents a referred area by the formula.

### Properties & Fields

| Name | Type | Description |
|------|------|-------------|
| `is_external_link` | Property | Indicates whether this is an external link. |
| `external_file_name` | Property | Get the external file name if this is an external reference. |
| `sheet_name` | Property | Indicates which sheet this is in |
| `is_area` | Property | Indicates whether this is an area. |
| `end_column` | Property | The end column of the area. |
| `start_column` | Property | The start column of the area. |
| `end_row` | Property | The end row of the area. |
| `start_row` | Property | The start row of the area. |

### Methods

#### get_values

```python
get_values()
```

Gets cell values in this area. If this area is one single cell, then return the cell value object; Otherwise return one array for all values in this area.

**Returns**: If this area is invalid, "#REF!" will be returned;

#### get_value

```python
get_value(row_offset, col_offset)
```

Gets cell value with given offset from the top-left of this area. "#N/A" if given offset out of this area; Otherwise return the cell value at given position.

**Parameters**:

- `row_offset`: row offset from the start row of this area
- `col_offset`: column offset from the start row of this area

**Returns**: "#REF!" if this area is invalid;

---

## GridUpdateMonitor

**Kind**: Class


**Description**: Monitor for user to track the change of update operation.

### Methods

#### after_update

```python
after_update(op, uid, cells)
```

after update operation

**Parameters**:

- `op`: The JSON string for update operation
- `uid`: The unique id for the file cache.
- `cells`: The Updated Cells list,include cells which has style change,value change or formula change

---

## GridWarningInfo

**Kind**: Class


**Description**: Represents warning information for GridJs.

### Properties & Fields

| Name | Type | Description |
|------|------|-------------|
| `type` | Property | Gets the warning type. |
| `description` | Property | Gets the warning description. |
| `corrected_object` | Property | Gets the corrected object. |
| `error_object` | Property | Gets the error object. |

### Methods

#### __init__

```python
__init__()
```

GridWarningInfo __init__.

---

## GridWorkbookSettings

**Kind**: Class


**Description**: Represents the settings of the workbook.

### Properties & Fields

| Name | Type | Description |
|------|------|-------------|
| `max_iteration` | Property | Returns the maximum number of iterations to resolve a circular reference, the default value is 100. |
| `iteration` | Property | Indicates whether use iteration to resolve circular references. |
| `force_full_calculate` | Property | Indicates whether fully calculates every time when a calculation is triggered. |
| `create_calc_chain` | Property | Indicates whether create calculated formulas chain. Default is false. |
| `re_calculate_on_open` | Property | Indicates whether re-calculate all formulas on opening file. Default is true. |
| `precision_as_displayed` | Property | True if calculations in this workbook will be done using only the precision of the numbers as they\'re displayed |
| `date1904` | Property | Gets a value which represents if the workbook uses the 1904 date system. |
| `enable_macros` | Property | Enable macros; Now it only works when copying a worksheet to other worksheet in a workbook. |
| `check_custom_number_format` | Property | Indicates whether checking custom number format when setting Style.Custom, default is false. |
| `check_excel_restriction` | Property | Whether check restriction of excel file when user modify cells related objects. For example, excel does not allow inputting string value longer than 32K. When you input a value longer than 32K such as by Cell.PutValue(string), if this property is true, you will get an Exception. If this property is false, we will accept your input string value as the cell\'s value so that later you can output the complete string value for other file formats such as CSV. However, if you have set such kind of value that is invalid for excel file format, you should not save the workbook as excel file format later. Otherwise there may be unexpected error for the generated excel file. default is false. |
| `author` | Property | Gets/sets the author of the file. |

### Methods

#### __init__

```python
__init__()
```

GridWorkbookSettings __init__.

---

## IGridJsService

**Kind**: Interface


**Description**: Reprensents the basic operation apis interface used in controller actions.

### Methods

#### update_cell

```python
update_cell(p, uid)
```

Applies the update operation.

**Parameters**:

- `p`: The JSON format string of update operation.
- `uid`: The unique id for the file cache.

**Returns**: The JSON format string of the update result.

#### check_in_cache_for_collaborative

```python
check_in_cache_for_collaborative()
```

Check wether workbook instance is in memory cache .this method is apply for Collaborative mode only.

**Returns**: Returns bool.

#### detail_stream_json_with_uid

```python
detail_stream_json_with_uid(stream, file_path, uid)
```

Write the JSON string for the file to the stream by the specified unique id.

**Parameters**:

- `stream`: The stream that will be written
- `file_path`: The file path
- `uid`: The unique id for the file cache.

#### detail_stream_json

```python
detail_stream_json(stream, file_path)
```

Write the JSON string for the Workbook to the stream

**Parameters**:

- `stream`: The stream that will be written
- `file_path`: The file path

#### lazy_loading_stream_json

```python
lazy_loading_stream_json(stream, sheet_name, uid)
```

Writes the JSON string of the specified sheet in the file from the cache using the specified unique id to the stream..

**Parameters**:

- `stream`: The stream that will be written
- `sheet_name`: The sheet name.
- `uid`: The unique id for the file cache.

#### add_image

```python
add_image(p, uid, iscontrol, file_input_stream)
```

Applies the add image from local file operation.

**Parameters**:

- `p`: The JSON string parameter
- `uid`: The unique id for the file cache.
- `iscontrol`: Specify whether it is a control.
- `file_input_stream`: The input stream form file of the image

**Returns**: The JSON string result

#### add_image_by_url

```python
add_image_by_url(p, uid, imageurl)
```

Applies the add image from remote URL operation.

**Parameters**:

- `p`: The JSON string parameter
- `uid`: The unique id for the file cache.
- `imageurl`: Specify the image URL.

**Returns**: The JSON string result

#### copy_image

```python
copy_image(p, uid)
```

Applies the copy image operation.

**Parameters**:

- `p`: The JSON string parameter
- `uid`: The unique id for the file cache.

**Returns**: The JSON string result

#### load

```python
load(uid, filename)
```

Gets the JSON string of the file from the cache using the specified unique id,set the output filename in the JSON.

**Parameters**:

- `uid`: The unique id for the file cache.
- `filename`: Specifies the file name in the JSON. If set to null,the default filename is: book1.

**Returns**: The JSON string

#### image

```python
image(uid, picid)
```

Get Stream of image.

**Parameters**:

- `uid`: The unique id for the file cache.
- `picid`: The image id.

**Returns**: The image stream

#### ole

```python
ole(uid, sheetname, oleid, label)
```

Gets the byte array data of the embedded ole object .

**Parameters**:

- `uid`: The unique id for the file cache.
- `sheetname`: The worksheet name.
- `oleid`: The  id for the embedded ole object.
- `label`: The display label of the embedded ole object.

**Returns**: The byte array data of the  embedded ole object .

#### image_url

```python
image_url(base_url, picid, uid)
```

Gets the image URL.

**Parameters**:

- `base_url`: The base action URL.
- `picid`: The image id.
- `uid`: The unique id for the file cache.

**Returns**: The image URL

#### get_file

```python
get_file(fileid)
```

Get file stream

**Parameters**:

- `fileid`: the file id

**Returns**: The stream of the file

#### download

```python
download(p, uid, filename)
```

Applies the download file operation

**Parameters**:

- `p`: The JSON parameter
- `uid`: The unique id for the file cache.
- `filename`: The file name

**Returns**: The file URL

#### get_warning_callback

```python
get_warning_callback()
```

Gets custom warning callback for import file.

**Returns**: The warning callback

#### set_warning_callback

```python
set_warning_callback(callback)
```

Sets custom warning callback for import file.

**Parameters**:

- `callback`: The warning callback

---

## ITextTranslator

**Kind**: Interface


**Description**: Represents the interface for translate

---

## IWarningCallback

**Kind**: Interface


**Description**: Public warning callback interface for GridJs. Use this interface instead of Aspose.Cells.IWarningCallback which is internal under PYTHONVIANET.

### Methods

#### warning

```python
warning(warning_info)
```

Called when a warning occurs.

**Parameters**:

- `warning_info`: The warning information.

---

## OprMessageService

**Kind**: Class


**Description**: This class provide all the operations for messages sync in Collaborative mode .

### Methods

#### __init__

```python
__init__()
```

OprMessageService __init__.

---

## PdfSaveOptions

**Kind**: Class


**Description**: Represents the options for saving pdf file.

### Properties & Fields

| Name | Type | Description |
|------|------|-------------|
| `font_encoding` | Property | Gets embedded font encoding in pdf. |
| `display_doc_title` | Property | Indicates whether the window\'s title bar should display the document title. |
| `export_document_structure` | Property | Indicates whether to export document structure. |
| `custom_properties_export` | Property | Gets a value determining the way of custom document properties are exported to PDF file. Default value is None. |
| `optimization_type` | Property | Gets and sets pdf optimization type. |
| `producer` | Property | Gets and sets producer of generated pdf document. |
| `created_time` | Property | Gets and sets the time of generating the pdf document. |
| `pdf_compression` | Property | Indicate the compression algorithm |
| `calculate_formula` | Property | Indicates whether to calculate formulas before saving pdf file. |
| `security_options` | Property | Set this options, when security is need in xls2pdf result. |
| `compliance` | Property | Gets the PDF standards compliance level for output documents. |
| `embed_standard_windows_fonts` | Property | True to embed true type fonts. Affects only ASCII characters 32-127. Fonts for character codes greater than 127 are always embedded. Fonts are always embedded for PDF/A-1a, PDF/A-1b standard. Default is true. |
| `embed_attachments` | Property | Indicates whether to embed attachment for Ole objects in Excel. |
| `default_edit_language` | Property | Gets default edit language. |
| `text_cross_type` | Property | Gets displaying text type when the text width is larger than cell width. |
| `gridline_color` | Property | Gets gridline colr. |
| `gridline_type` | Property | Gets gridline type. |
| `printing_page_type` | Property | Indicates which pages will not be printed. |
| `emf_render_setting` | Property | Setting for rendering Emf metafile. |
| `page_count` | Property | Gets the number of pages to save. |
| `page_index` | Property | Gets the 0-based index of the first page to save. |
| `ignore_error` | Property | Indicates if you need to hide the error while rendering. The error can be error in shape, image, chart rendering, etc. |
| `output_blank_page_when_nothing_to_print` | Property | Indicates whether to output a blank page when there is nothing to print. |
| `default_font` | Property | When characters in the Excel are Unicode and not be set with correct font in cell style, They may appear as block in pdf,image. Set the DefaultFont such as MingLiu or MS Gothic to show these characters. If this property is not set, Aspose.Cells will use system default font to show these unicode characters. |
| `check_workbook_default_font` | Property | When characters in the Excel are Unicode and not be set with correct font in cell style, They may appear as block in pdf,image. Set this to true to try to use workbook\'s default font to show these characters first. |
| `check_font_compatibility` | Property | Indicates whether to check font compatibility for every character in text. |
| `is_font_substitution_char_granularity` | Property | Indicates whether to only substitute the font of character when the cell font is not compatibility for it. |
| `one_page_per_sheet` | Property | If OnePagePerSheet is true , all content of one sheet will output to only one page in result. The paper size of pagesetup will be invalid, and the other settings of pagesetup will still take effect. |
| `all_columns_in_one_page_per_sheet` | Property | If AllColumnsInOnePagePerSheet is true , all column content of one sheet will output to only one page in result. The width of paper size of pagesetup will be ignored, and the other settings of pagesetup will still take effect. |

### Methods

#### __init__

```python
__init__()
```

Creates the options for saving pdf file.

#### set_sheet_set

```python
set_sheet_set(sheet_indexes)
```

Creates a sheet set based on exact sheet indexes.

**Parameters**:

- `sheet_indexes`: zero based sheet indexes.

#### set_image_resample

```python
set_image_resample(desired_ppi, jpeg_quality)
```

Sets desired PPI(pixels per inch) of resample images and jpeg quality. All images will be converted to JPEG with the specified quality setting, and images that are greater than the specified PPI (pixels per inch) will be resampled.

**Parameters**:

- `desired_ppi`: Desired pixels per inch. 220 high quality. 150 screen quality. 96 email quality.
- `jpeg_quality`: 0 - 100% JPEG quality.

---

## PdfSecurityOptions

**Kind**: Class


**Description**: Options for encrypting and access permissions for a PDF document. PDF/A does not allow security setting.

### Properties & Fields

| Name | Type | Description |
|------|------|-------------|
| `user_password` | Property | Gets the user password required for opening the encrypted PDF document. |
| `owner_password` | Property | Gets the owner password for the encrypted PDF document. |
| `print_permission` | Property | Indicates whether to allow to print the document. |
| `modify_document_permission` | Property | Indicates whether to allow to modify the contents of the document by operations other than those controlled by :py:attr:`aspose.cellsgridjs.PdfSecurityOptions.annotations_permission`, :py:attr:`aspose.cellsgridjs.PdfSecurityOptions.fill_forms_permission` and :py:attr:`aspose.cellsgridjs.PdfSecurityOptions.assemble_document_permission`. |
| `annotations_permission` | Property | Indicates whether to allow to add or modify text annotations, fill in interactive form fields. |
| `fill_forms_permission` | Property | Indicates whether to allow to fill in existing interactive form fields (including signature fields), even if :py:attr:`aspose.cellsgridjs.PdfSecurityOptions.modify_document_permission` is clear. |
| `extract_content_permission` | Property | Indicates whether to allow to copy or otherwise extract text and graphics from the document by operations other than that controlled by :py:attr:`aspose.cellsgridjs.PdfSecurityOptions.accessibility_extract_content`. |
| `accessibility_extract_content` | Property | Indicates whether to allow to extract text and graphics (in support of accessibility to users with disabilities or for other purposes). |
| `assemble_document_permission` | Property | Indicates whether to allow to assemble the document (insert, rotate, or delete pages and create bookmarks or thumbnail images), even if :py:attr:`aspose.cellsgridjs.PdfSecurityOptions.modify_document_permission` is clear. |
| `full_quality_print_permission` | Property | Indicates whether to allow to print the document to a representation from which a faithful digital copy of the PDF content could be generated. |

### Methods

#### __init__

```python
__init__()
```

The constructor of PdfSecurityOptions

---

## SheetSet

**Kind**: Class


**Description**: Describes a set of sheets.

### Methods

#### __init__

```python
__init__(sheet_indexes)
```

Creates a sheet set based on exact sheet indexes.

**Parameters**:

- `sheet_indexes`: zero based sheet indexes.

---

# Enumerations

## CoWorkOperationType

Represents the CoWorkOperationType enumeration.

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

## CoWorkUserPermission

Represents the CoWorkUserPermission enumeration.

| Member | Description |
|--------|-------------|
| `READ_ONLY` | Read-only permission,User can only view the content but cannot make any modifications |
| `EDITABLE` | Editable permission,User can view and edit the content |
| `DOWNLOAD` | Download permission,User can view,edit and download the content to local storage |
| `ADMIN` | Administrator permission,User has full access including management operations |

## DefaultEditLanguage

Represents the DefaultEditLanguage enumeration.

| Member | Description |
|--------|-------------|
| `AUTO` | Represents auto detecting edit language according to the text itself. |
| `ENGLISH` | Represents English language. |
| `CJK` | Represents Chinese, Japanese, Korean language. |

## EmfRenderSetting

Represents the EmfRenderSetting enumeration.

| Member | Description |
|--------|-------------|
| `EMF_ONLY` | Only rendering Emf records. |
| `EMF_PLUS_PREFER` | Prefer rendering EmfPlus records. |

## GridExceptionType

Represents the GridExceptionType enumeration.

| Member | Description |
|--------|-------------|
| `CHART` | Invalid chart setting. |
| `DATA_TYPE` | Invalid data type setting. |
| `DATA_VALIDATION` | Invalid data validation setting. |
| `CONDITIONAL_FORMATTING` | Invalid data validation setting. |
| `FILE_FORMAT` | Invalid file format. |
| `FORMULA` | Invalid formula. |
| `INVALID_DATA` | Invalid data. |
| `INVALID_OPERATOR` | Invalid operator. |
| `INCORRECT_PASSWORD` | Incorrect password. |
| `LICENSE` | License related errors. |
| `LIMITATION` | Out of MS Excel limitation error. |
| `PAGE_SETUP` | Invalid page setup setting. |
| `PIVOT_TABLE` | Invalid pivotTable setting. |
| `SHAPE` | Invalid drawing object setting. |
| `SPARKLINE` | Invalid sparkline object setting. |
| `SHEET_NAME` | Invalid worksheet name. |
| `SHEET_TYPE` | Invalid worksheet type. |
| `INTERRUPTED` | The process is interrupted. |
| `IO` | The file is invalid. |
| `PERMISSION` | Permission is required to open this file. |
| `UNSUPPORTED_FEATURE` | Unsupported feature. |
| `UNSUPPORTED_STREAM` | Unsupported stream to be opened. |
| `UNDISCLOSED_INFORMATION` | Files contains some undisclosed information. |

## GridJsLogLevel

Represents the GridJsLogLevel enumeration.

| Member | Description |
|--------|-------------|
| `DEBUG` | Debug level, outputs all messages. |
| `INFO` | Info level. |
| `WARN` | Warning level. |
| `ERROR` | Error level. |
| `NONE` | Disables all logging. |

## GridJsLogRolling

Represents the GridJsLogRolling enumeration.

| Member | Description |
|--------|-------------|
| `NONE` | No rolling, single file. |
| `DAILY` | Roll by day, e.g. gridjs_2026-07-27.log |
| `MONTHLY` | Roll by month, e.g. gridjs_2026-07.log |

## GridlineType

Enumerates grid line Type.

| Member | Description |
|--------|-------------|
| `DOTTED` | Represents dotted line. |
| `HAIR` | Represents hair line. |

## GridLoadFormat

Represents the GridLoadFormat enumeration.

| Member | Description |
|--------|-------------|
| `AUTO` | Represents recognizing the format automatically. |
| `CSV` | Comma-Separated Values(CSV) text file. |
| `XLSX` | Represents Office Open XML spreadsheetML workbook or template, with or without macros. |
| `TSV` | Tab-Separated Values(TSV) text file. |
| `TAB_DELIMITED` | Represents a tab delimited text file, same with :py:attr:`aspose.cellsgridjs.GridLoadFormat.TSV`. |
| `HTML` | Represents a html file. |
| `M_HTML` | Represents a mhtml file. |
| `ODS` | Open Document Sheet(ODS) file. |
| `EXCEL_97_TO_2003` | Represents an Excel97-2003 xls file. |
| `SPREADSHEET_ML` | Represents an Excel 2003 xml file. |
| `XLSB` | Represents an xlsb file. |
| `OTS` | Open Document Template Sheet(OTS) file. |
| `NUMBERS` | Represents a numbers file. |
| `FODS` | Represents OpenDocument Flat XML Spreadsheet (.fods) file format. |
| `SXC` | Represents StarOffice Calc Spreadsheet (.sxc) file format. |
| `XML` | Represents a simple xml file. |
| `EPUB` | Reprents an EPUB file. |
| `AZW3` | Represents an AZW3 file. |
| `CHM` | Represents a CHM file. |
| `MARKDOWN` | Represents a Markdown file. |
| `UNKNOWN` | Represents unrecognized format, cannot be loaded. |
| `IMAGE` | Image |
| `JSON` | Json |
| `DIF` | Data Interchange Format. |
| `DBF` | Xbase Data file |

## GridWarningType

Represents the GridWarningType enumeration.

| Member | Description |
|--------|-------------|
| `CHART` | Invalid chart setting. |
| `DATA_TYPE` | Invalid data type setting. |
| `DATA_VALIDATION` | Invalid data validation setting. |
| `CONDITIONAL_FORMATTING` | Invalid conditional formatting setting. |
| `FILE_FORMAT` | Invalid file format. |
| `FORMULA` | Invalid formula. |
| `INVALID_DATA` | Invalid data. |
| `INVALID_OPERATOR` | Invalid operator. |
| `INCORRECT_PASSWORD` | Incorrect password. |
| `LICENSE` | License related errors. |
| `LIMITATION` | Out of MS Excel limitation error. |
| `PAGE_SETUP` | Invalid page setup setting. |
| `PIVOT_TABLE` | Invalid pivotTable setting. |
| `SHAPE` | Invalid drawing object setting. |
| `SPARKLINE` | Invalid sparkline object setting. |
| `SHEET_NAME` | Invalid worksheet name. |
| `SHEET_TYPE` | Invalid worksheet type. |
| `INTERRUPTED` | The process is interrupted. |
| `IO` | The file is invalid. |
| `PERMISSION` | Permission is required to open this file. |
| `UNSUPPORTED_FEATURE` | Unsupported feature. |
| `UNSUPPORTED_STREAM` | Unsupported stream to be opened. |
| `UNDISCLOSED_INFORMATION` | Files contains some undisclosed information. |
| `FILE_CORRUPTED` | File content is corrupted. |
| `DEFINED_NAME` | Invalid defined name. |
| `FONT` | Invalid font. |
| `AUTO_FILTER` | Invalid auto filter setting. |
| `FONT_SUBSTITUTION` | Font substitution warning type. |

## PdfCompliance

Represents the PdfCompliance enumeration.

| Member | Description |
|--------|-------------|
| `PDF14` | Pdf format compatible with PDF 1.4 |
| `PDF15` | Pdf format compatible with PDF 1.5 |
| `PDF16` | Pdf format compatible with PDF 1.6 |
| `PDF17` | Pdf format compatible with PDF 1.7 |
| `PDF_A1B` | Pdf format compatible with PDF/A-1b(ISO 19005-1) |
| `PDF_A1A` | Pdf format compatible with PDF/A-1a(ISO 19005-1) |
| `PDF_A2B` | Pdf format compatible with PDF/A-2b(ISO 19005-2) |
| `PDF_A2U` | Pdf format compatible with PDF/A-2u(ISO 19005-2) |
| `PDF_A2A` | Pdf format compatible with PDF/A-2a(ISO 19005-2) |
| `PDF_A3B` | Pdf format compatible with PDF/A-3b(ISO 19005-3) |
| `PDF_A3U` | Pdf format compatible with PDF/A-3u(ISO 19005-3) |
| `PDF_A3A` | Pdf format compatible with PDF/A-3a(ISO 19005-3) |

## PdfCompressionCore

Represents the PdfCompressionCore enumeration.

| Member | Description |
|--------|-------------|
| `NONE` | None |
| `RLE` | Rle |
| `LZW` | Lzw |
| `FLATE` | Flate |

## PdfCustomPropertiesExport

Represents the PdfCustomPropertiesExport enumeration.

| Member | Description |
|--------|-------------|
| `NONE` | No custom properties are exported. |
| `STANDARD` | Custom properties are exported as entries in Info dictionary. |

## PdfFontEncoding

Represents the PdfFontEncoding enumeration.

| Member | Description |
|--------|-------------|
| `IDENTITY` | Represents use Identity-H encoding for all embedded fonts in pdf. |
| `ANSI_PREFER` | Represents prefer to use WinAnsiEncoding for TrueType fonts with characters 32-127, otherwise, Identity-H encoding will be used for embedded fonts in pdf. |

## PdfOptimizationType

Represents the PdfOptimizationType enumeration.

| Member | Description |
|--------|-------------|
| `STANDARD` | High print quality |
| `MINIMUM_SIZE` | File size is more important than print quality |

## PrintingPageType

Represents the PrintingPageType enumeration.

| Member | Description |
|--------|-------------|
| `DEFAULT` | Prints all pages. |
| `IGNORE_BLANK` | Don\'t print the pages which the cells are blank. |
| `IGNORE_STYLE` | Don\'t print the pages which cells only contain styles. |

## TextCrossType

Enumerates displaying text type when the text width is larger than cell width.

| Member | Description |
|--------|-------------|
| `DEFAULT` | Display text like in Microsoft Excel. |
| `CROSS_KEEP` | Display all the text by crossing other cells and keep text of crossed cells. |
| `CROSS_OVERRIDE` | Display all the text by crossing other cells and override text of crossed cells. |
| `STRICT_IN_CELL` | Only display the text within the width of cell. |

