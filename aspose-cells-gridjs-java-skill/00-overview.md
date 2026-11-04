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
│          Server (Java)                                   │
│  ┌────────────────────────────────────────────────────┐  │
│  │  GridJsController extends GridJsControllerBase     │  │
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

- **Config** (Class): Represents the Config class shipped in aspose-cells-26.10-gridjs.jar.
- **CoWorkOperationType** (Class): Represents the CoWorkOperationType class shipped in aspose-cells-26.10-gridjs.jar.
- **CoWorkUserPermission** (Class): Represents the CoWorkUserPermission class shipped in aspose-cells-26.10-gridjs.jar.
- **CoWorkUserProvider** (Interface): Represents the CoWorkUserProvider interface shipped in aspose-cells-26.10-gridjs.jar.
- **GridAbstractCalculationEngine** (Class): Represents the GridAbstractCalculationEngine class shipped in aspose-cells-26.10-gridjs.jar.
- **GridCacheForStream** (Class): Abstract class for custom cache implementations. Extend this to provide custom caching logic for GridJs files.
- **GridCalculationData** (Class): Represents the GridCalculationData class shipped in aspose-cells-26.10-gridjs.jar.
- **GridCellException** (Class): Represents the GridCellException class shipped in aspose-cells-26.10-gridjs.jar.
- **GridExceptionType** (Class): Represents the GridExceptionType class shipped in aspose-cells-26.10-gridjs.jar.
- **GridInterruptMonitor** (Class): Represents the GridInterruptMonitor class shipped in aspose-cells-26.10-gridjs.jar.
- **GridJsControllerBase** (Class): Base controller class for GridJs operations. Provides built-in endpoints for cell updates, image handling, OLE objects, ...
- **GridJsCoService** (Class): Service for collaborative editing operations.
- **GridJsJakartaControllerBase** (Class): Base controller class for GridJs operations with Jakarta EE support. For Spring Boot 3.x with jakarta.servlet. For Sprin...
- **GridJsLogger** (Class): Logger for GridJs operations.
- **GridJsLogLevel** (Class): Log levels for GridJs logging.
- **GridJsLogRolling** (Class): Log rolling configuration.
- **GridJsOptions** (Class): Configuration options for GridJs. Contains settings for caching, lazy loading, image handling, and other GridJs behavior...
- **GridJsPermissionException** (Class): Represents the GridJsPermissionException class shipped in aspose-cells-26.10-gridjs.jar.
- **GridJsService** (Class): Default implementation of IGridJsService. Handles all GridJs operations including workbook loading, cell updates, image ...
- **GridJsWorkbook** (Class): Represents a workbook for GridJs operations. Provides methods for importing/exporting Excel files, managing images, and ...
- **GridLoadFormat** (Class): Represents the GridLoadFormat class shipped in aspose-cells-26.10-gridjs.jar.
- **GridReferredArea** (Class): Represents the GridReferredArea class shipped in aspose-cells-26.10-gridjs.jar.
- **GridUpdateMonitor** (Class): Represents the GridUpdateMonitor class shipped in aspose-cells-26.10-gridjs.jar.
- **GridWorkbookSettings** (Class): Represents the GridWorkbookSettings class shipped in aspose-cells-26.10-gridjs.jar.
- **IGridJsService** (Interface): Interface for GridJs service operations. Provides methods for loading workbooks, updating cells, managing images, and ha...
- **ITextTranslator** (Interface): Represents the ITextTranslator interface shipped in aspose-cells-26.10-gridjs.jar.
- **OprMessage** (Class): Represents the OprMessage class shipped in aspose-cells-26.10-gridjs.jar.
- **OprMessageDto** (Class): Represents the OprMessageDto class shipped in aspose-cells-26.10-gridjs.jar.

## Key Server-Side Enums


## Client-Side Entry Point

The client-side API is provided by the `gridjs-spreadsheet` npm package.
Main entry point: `x_spreadsheet(container, options)` which returns a `Spreadsheet` instance.

