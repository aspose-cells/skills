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
│                    Server (C#        )                          │
│  ┌────────────────────────────────────────────────────┐  │
│  │  GridJsController (extends GridJsControllerBase)   │  │
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
- **CoWorkUserPermission** (Class): represent the user permission  in collaboration mode.only available in java version now, will be available in .net/pytho...
- **CoWorkUserProvider** (Interface): Represents the user provider inerface in collabration mode.only available in java version now, will be available in .net...
- **GridAbstractCalculationEngine** (Class): Represents user's custom calculation engine to extend the default calculation engine of Aspose.Cells.
- **GridCacheForStream** (Class): This class contains the cache operations for GridJs. User shall implement his own business logic for storage based on it...
- **GridCalculationData** (Class): Represents the required data when calculating one function, such as function name, parameters, ...etc.
- **GridCellException** (Class): The exception that is thrown when GridJs specified error occurs.
- **GridExceptionType** (Class): Represents custom exception code for GridJs.
- **GridInterruptMonitor** (Class): Represents all operator about the calculation interrupt.
- **GridJsLogger** (Class): Provides a lightweight logging system for GridJs server. Supports console and file output with configurable log levels. ...
- **GridJsLogLevel** (Class): Specifies the log level for GridJs logging.
- **GridJsLogRolling** (Class): Specifies the log file rolling strategy.
- **GridJsOptions** (Class): Represents  all the load options for GridJs
- **GridJsPermissionException** (Class): represents permission exception in collaboration mode.only available in java version now, will be available in .net/pyth...
- **GridJsService** (Class): Provides the basic operation apis used in controller actions.
- **GridJsWorkbook** (Class): Represents the main entry class for GridJs
- **GridLoadFormat** (Class): Represents the load file format.
- **GridReferredArea** (Class): Represents a referred area by the formula.
- **GridUpdateMonitor** (Class): Monitor for user to track the change of update operation.
- **GridWorkbookSettings** (Class): Represents the settings of the workbook.
- **IGridJsService** (Interface): Reprensents the basic operation apis interface used in controller actions.
- **ITextTranslator** (Interface): Represents the interface for translate
- **NamespaceDoc** (Struct): The Chart namespace encapsulates all classes of GridJs , providing  basic data structure for Charts JSON generation.
- **NamespaceDoc** (Class): The Aspose.Cells.GridJs namespace encapsulates all classes of GridJs, providing simple APIs for viewing or editing sprea...
- **OprMessageService** (Class): This class provide all the operations for messages sync in Collaborative mode .

## Key Server-Side Enums

- **CoWorkOperationType**: Represents the action operation type in collabration mode.only available in java version now, will b...

## Client-Side Entry Point

The client-side API is provided by the `gridjs-spreadsheet` npm package.
Main entry point: `x_spreadsheet(container, options)` which returns a `Spreadsheet` instance.

## Quick Start: Controller Setup

Create a controller that inherits from `GridJsControllerBase`. The base class requires an `IGridJsService` instance via constructor injection:

```csharp
[Route("[controller]/[action]")]
public class GridJsController : GridJsControllerBase
{
    public GridJsController(IGridJsService gridJsService) : base(gridJsService)
    {
    }
}
```
