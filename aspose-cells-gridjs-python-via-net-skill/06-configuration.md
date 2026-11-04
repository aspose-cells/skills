# Configuration Reference

## Config (Static Settings)

Represents all the static settings for GridJs

These are static properties on the `Config` class that affect global behavior.

| Property | Description |
|----------|-------------|
| `save_html_as_zip` | Gets/Sets whether to save html file as zip archive,the default is false. |
| `skip_invisible_shapes` | Gets/Sets whether to skip shapes that are invisble to UI ,the default value is true. |
| `lazy_loading` | Gets/Sets whether to load active worksheet only,the default is false. |
| `same_image_detecting` | Gets/Sets whether to check if images have same source,the default is true the default value is true. |
| `auto_optimize_for_large_cells` | Gets/Sets whether to automatically optimize the load performance for worksheet with large cells. it will ignore some style /borders to reduce the load time. the default value is true. |
| `islimit_shape_or_image` | Gets/Sets whether to limit the total display shape/image count in one worksheet ,if set to true, GridJs will limit the total count of the display shapes or images in one worksheet to MaxShapeOrImageCount the default value is true. |
| `max_shape_or_image_count` | Gets/Sets the total count of the display shapes or images in the active sheet,it takes effect when IslimitShapeOrImage=true. the default value is 100. |
| `max_total_shape_or_image_count` | Gets/Sets the total count of the display shapes or images in the workbook,it takes effect when IslimitShapeOrImage=true. the default value is 300. |
| `max_shape_or_image_width_or_height` | Gets/Sets the max width or height for a shape or an image ,GridJs will ignore the shape or image with the width or height larger than this, it takes effect when IslimitShapeOrImage=true. the default value is 10000. |
| `max_pdf_save_seconds` | Gets/Sets the max timed out seconds when save to PDF. the default value is 10. |
| `ignore_empty_content` | Gets/Sets whether to show the max range which includes data ,style, merged cells and shapes. if the last row or column contains cells with no value and formula but has custom style then we will not show this row/column when this vlaue is true。 the default value is true . |
| `use_print_area` | Gets/Sets whether to use PageSetup.PrintArea for the UI display range when the worksheet has PageSetup setting for PrintArea. the default value is false . |
| `is_collaborative` | Gets/Sets whether to support collabrative editing,the default is false. |
| `enable_chart_client_rendering` | Gets/Sets whether to enable client-side chart rendering. If set to false, GridJs will fall back to server-generated chart images/shapes only. the default value is true. |
| `custom_pdf_save_options` | Gets/Sets the custom PdfSaveOptions for PDF export. If set, this will be used instead of the default options. the default value is null. |
| `load_time_out` | Gets/Sets a timeout interrupt in milliseconds in loading file, when the cost time period of loading file is longer than the expectation ，it will raise exception. the default value is -1,which means no timeout interrupt is set . |
| `show_chart_sheet` | Gets/Sets whether to show chart worksheet. the default value is false . |
| `empty_sheet_max_row` | Gets/Sets default max row for an empty worksheet. the default value is 12. |
| `empty_sheet_max_col` | Gets/Sets default max column for an empty worksheet. the default value is 15. |
| `picture_cache_directory` | Gets/Sets the cache directory for pictures.(this takes effect when GridJsWorkbook.CacheImp is null) the default path will be "_piccache" inside the FileCacheDirectory. |
| `file_cache_directory` | Gets/Sets the cache directory for storing spreadsheet file. We need to set it to a specific path before we use GridJs. |
| `base_route_name` | Gets/Sets the base route name for GridJs controller URL. the default is "/GridJs2". |
| `message_topic` | Gets/Sets the websocket destinations prefixed with "/topic". the default is "/topic/opr".used in collaborative mode only. |
| `auto_fit_rows_height_on_load` | Indicates whether to auto-fit row heights during file loading. The default value is ``false``. Warning: Setting this to ``true`` will perform an auto-fit all rows operation post-load, which may have a noticeable impact on performance. |
| `auto_fit_columns_on_load` | Indicates whether to auto-fit column widths during file loading. The default value is ``false``. Warning: Setting this to ``true`` will perform an auto-fit all columns operation post-load, which may have a noticeable impact on performance. |
| `redaction_use_client_generate_id` | Indicates whether use client generate id instead of actual shape id in redaction related operations |

## GridJsOptions (Instance Settings)

Represents all the load options for GridJs

These options are configured per-service instance via DI.

| Property | Description |
|----------|-------------|
| `save_html_as_zip` | Gets/Sets whether to save html file as zip archive,the default is false. |
| `skip_invisible_shapes` | Gets/Sets whether to skip shapes that are invisble to UI ,the default value is true. |
| `lazy_loading` | Gets/Sets whether to load active worksheet only,the default is false. |
| `same_image_detecting` | Gets/Sets whether to check if images have same source,the default is true the default value is true. |
| `auto_optimize_for_large_cells` | Gets/Sets whether to automatically optimize the load performance for worksheet with large cells. it will ignore some style /borders to reduce the load time. the default value is true. |
| `islimit_shape_or_image` | Gets/Sets whether to limit the total display shape/image count in one worksheet ,if set to true, GridJs will limit the total count of the display shapes or images in one worksheet to MaxShapeOrImageCount the default value is true. |
| `max_shape_or_image_count` | Gets/Sets the total count of the display shapes or images in the active sheet,it takes effect when IslimitShapeOrImage=true. the default value is 100. |
| `max_total_shape_or_image_count` | Gets/Sets the total count of the display shapes or images in the workbook,it takes effect when IslimitShapeOrImage=true. the default value is 300. |
| `max_shape_or_image_width_or_height` | Gets/Sets the max width or height for a shape or an image ,GridJs will ignore the shape or image with the width or height larger than this, it takes effect when IslimitShapeOrImage=true. the default value is 10000. |
| `max_pdf_save_seconds` | Gets/Sets the max timed out seconds when save to PDF. the default value is 10. |
| `ignore_empty_content` | Gets/Sets whether to show the max range which includes data ,style, merged cells and shapes. if the last row or column contains cells with no value and formula but has custom style then we will not show this row/column when this vlaue is true。 the default value is true . |
| `use_print_area` | Gets/Sets whether to use PageSetup.PrintArea for the UI display range when the worksheet has PageSetup setting for PrintArea. the default value is false . |
| `is_collaborative` | Gets/Sets whether to support collabrative editing,the default is false. |
| `enable_chart_client_rendering` | Gets/Sets whether to enable client-side chart rendering. If set to false, GridJs will emit chart images/shapes only and will not require client-side chart rendering. the default value is true. |
| `custom_pdf_save_options` | Gets/Sets the custom PdfSaveOptions for PDF export. If set, this will be used instead of the default options. the default value is null. |
| `load_time_out` | Gets/Sets a timeout interrupt in milliseconds in loading file, when the cost time period of loading file is longer than the expectation ，it will raise exception. the default value is -1,which means no timeout interrupt is set . |
| `show_chart_sheet` | Gets/Sets whether to show chart worksheet. the default value is false . |
| `empty_sheet_max_row` | Gets/Sets default max row for an empty worksheet. the default value is 12. |
| `empty_sheet_max_col` | Gets/Sets default max column for an empty worksheet. the default value is 15. |
| `picture_cache_directory` | Gets/Sets the cache directory for pictures.(this takes effect when GridJsWorkbook.CacheImp is null) the default path will be "_piccache" inside the FileCacheDirectory. |
| `file_cache_directory` | Gets/Sets the cache directory for storing spreadsheet file. We need to set it to a specific path before we use GridJs. |
| `font_folders` | Gets/Sets the fonts folders for fonts in the rendered pictures/shapes |
| `base_route_name` | Gets/Sets the route URL base name for GridJs controller.the default is GridJs2 |
| `message_topic` | Gets/Sets the websocket destinations prefixed with "/topic". the default is "/topic/opr".used in collaborative mode only. |
| `auto_fit_rows_height_on_load` | Indicates whether to autofit rows height when loading the file,the default value is false. |
| `auto_fit_columns_on_load` | Indicates whether to autofit columns width when loading the file,the default value is false. |
| `cache_imp` | Custom implemention for cache storage,If you want to store cache in stream way ,you need to set and implement it. |
| `redaction_use_client_generate_id` | Indicates whether use client generate id instead of actual shape id in redaction related operations |

## GridWorkbookSettings

Represents the settings of the workbook.

| Property | Description |
|----------|-------------|
| `max_iteration` | Returns the maximum number of iterations to resolve a circular reference, the default value is 100. |
| `iteration` | Indicates whether use iteration to resolve circular references. |
| `force_full_calculate` | Indicates whether fully calculates every time when a calculation is triggered. |
| `create_calc_chain` | Indicates whether create calculated formulas chain. Default is false. |
| `re_calculate_on_open` | Indicates whether re-calculate all formulas on opening file. Default is true. |
| `precision_as_displayed` | True if calculations in this workbook will be done using only the precision of the numbers as they\'re displayed |
| `date1904` | Gets a value which represents if the workbook uses the 1904 date system. |
| `enable_macros` | Enable macros; Now it only works when copying a worksheet to other worksheet in a workbook. |
| `check_custom_number_format` | Indicates whether checking custom number format when setting Style.Custom, default is false. |
| `check_excel_restriction` | Whether check restriction of excel file when user modify cells related objects. For example, excel does not allow inputting string value longer than 32K. When you input a value longer than 32K such as by Cell.PutValue(string), if this property is true, you will get an Exception. If this property is false, we will accept your input string value as the cell\'s value so that later you can output the complete string value for other file formats such as CSV. However, if you have set such kind of value that is invalid for excel file format, you should not save the workbook as excel file format later. Otherwise there may be unexpected error for the generated excel file. default is false. |
| `author` | Gets/sets the author of the file. |

