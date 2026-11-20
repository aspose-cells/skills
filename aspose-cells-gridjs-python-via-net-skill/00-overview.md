# GridJs Overview

## What is GridJs?

GridJs (Aspose.Cells.GridJs) is a web-based spreadsheet component that provides:

- **Server-side**: Excel file parsing, JSON conversion, cell update, image handling, PDF/XLSX/HTML export
- **Client-side**: Spreadsheet UI rendering, cell editing, event handling, toolbar, context menu

## Architecture

```
┌──────────────────────────────────────────────────────────┐
│                    Browser (Client)                       │
│  ┌────────────────────────────────────────────────────┐  │
│  │  gridjs-spreadsheet (x_spreadsheet)                │  │
│  │  - Spreadsheet UI rendering                        │  │
│  │  - Cell editing & event handling                   │  │
│  │  - Toolbar, context menu, sheet tabs               │  │
│  │  - Semantic Automation API (for AI agents)         │  │
│  └────────────────────────────────────────────────────┘  │
│                      ↕ HTTP/JSON                         │
├──────────────────────────────────────────────────────────┤
│        Server (Python)                                   │
│  ┌────────────────────────────────────────────────────┐  │
│  │  Flask routes -> GridJsService                     │  │
│  │  - LoadSpreadsheet: Excel → JSON                   │  │
│  │  - UpdateCell: Apply cell changes                  │  │
│  │  - Image/OLE handling                              │  │
│  │  - File download                                   │  │
│  └────────────────────────────────────────────────────┘  │
│  ┌────────────────────────────────────────────────────┐  │
│  │  IGridJsService / GridJsService                    │  │
│  │  - Core business logic                             │  │
│  │  - File caching & workbook management              │  │
│  └────────────────────────────────────────────────────┘  │
│  ┌────────────────────────────────────────────────────┐  │
│  │  GridJsWorkbook (low-level API)                    │  │
│  │  - Import/Export Excel files                       │  │
│  │  - Custom calculation engine                       │  │
│  │  - Redaction support                               │  │
│  └────────────────────────────────────────────────────┘  │
└──────────────────────────────────────────────────────────┘
```

## Key Server-Side Classes

- **Config** (Class): Represents all the static settings for GridJs
- **CoWorkUserProvider** (Interface): Represents the user provider inerface in collabration mode.only available in java version now, will be available in .net...
- **GridAbstractCalculationEngine** (Class): Represents user\'s custom calculation engine to extend the default calculation engine of Aspose.Cells.
- **GridCacheForStream** (Class): This class contains the cache operations for GridJs. User shall implement his own business logic for storage based on it...
- **GridCalculationData** (Class): Represents the required data when calculating one function, such as function name, parameters, ...etc.
- **GridCellException** (Class): The exception that is thrown when GridJs specified error occurs.
- **GridJsControllerBase** (Class): Represents the GridJsControllerBase type.
- **GridJsLogger** (Class): Provides a lightweight logging system for GridJs server. Supports console and file output with configurable log levels. ...
- **GridJsOptions** (Class): Represents all the load options for GridJs
- **GridJsPermissionException** (Class): represents permission exception in collaboration mode.only available in java version now, will be available in .net/pyth...
- **GridJsService** (Class): Provides the basic operation apis used in controller actions.
- **GridJsWorkbook** (Class): Represents the main entry class for GridJs
- **GridReferredArea** (Class): Represents a referred area by the formula.
- **GridUpdateMonitor** (Class): Monitor for user to track the change of update operation.
- **GridWarningInfo** (Class): Represents warning information for GridJs.
- **GridWorkbookSettings** (Class): Represents the settings of the workbook.
- **IGridJsService** (Interface): Reprensents the basic operation apis interface used in controller actions.
- **ITextTranslator** (Interface): Represents the interface for translate
- **IWarningCallback** (Interface): Public warning callback interface for GridJs. Use this interface instead of Aspose.Cells.IWarningCallback which is inter...
- **OprMessageService** (Class): This class provide all the operations for messages sync in Collaborative mode .
- **PdfSaveOptions** (Class): Represents the options for saving pdf file.
- **PdfSecurityOptions** (Class): Options for encrypting and access permissions for a PDF document. PDF/A does not allow security setting.
- **SheetSet** (Class): Describes a set of sheets.

## Key Server-Side Enums

- **CoWorkOperationType**: Represents the CoWorkOperationType enumeration.
- **CoWorkUserPermission**: Represents the CoWorkUserPermission enumeration.
- **DefaultEditLanguage**: Represents the DefaultEditLanguage enumeration.
- **EmfRenderSetting**: Represents the EmfRenderSetting enumeration.
- **GridExceptionType**: Represents the GridExceptionType enumeration.
- **GridJsLogLevel**: Represents the GridJsLogLevel enumeration.
- **GridJsLogRolling**: Represents the GridJsLogRolling enumeration.
- **GridlineType**: Enumerates grid line Type.
- **GridLoadFormat**: Represents the GridLoadFormat enumeration.
- **GridWarningType**: Represents the GridWarningType enumeration.
- **PdfCompliance**: Represents the PdfCompliance enumeration.
- **PdfCompressionCore**: Represents the PdfCompressionCore enumeration.
- **PdfCustomPropertiesExport**: Represents the PdfCustomPropertiesExport enumeration.
- **PdfFontEncoding**: Represents the PdfFontEncoding enumeration.
- **PdfOptimizationType**: Represents the PdfOptimizationType enumeration.
- **PrintingPageType**: Represents the PrintingPageType enumeration.
- **TextCrossType**: Enumerates displaying text type when the text width is larger than cell width.

## Client-Side Entry Point

The client-side API is provided by the `gridjs-spreadsheet` npm package.
Main entry point: `x_spreadsheet(container, options)` which returns a `Spreadsheet` instance.

