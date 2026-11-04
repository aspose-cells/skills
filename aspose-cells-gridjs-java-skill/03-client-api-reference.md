# Client-Side API Reference

Package: `gridjs-spreadsheet`

## Spreadsheet (Main Class)

The main entry point for GridJs client-side API.

### Constructor

```typescript
new Spreadsheet(container: string | HTMLElement, opts?: Options)
```

### Factory Function

```typescript
x_spreadsheet(container: string | HTMLElement, opts?: Options): Spreadsheet
```

### Methods

#### cellStyle()

retrieve cell style

```typescript
cellStyle(rowIndex: number,
      colIndex: number,
      sheetIndex: number): CellStyle
```

**Parameters**:

- `rowIndex`: `number`
- `colIndex`: `number`
- `sheetIndex`: `number`

#### cellText()

get/set cell text

```typescript
cellText(rowIndex: number,
      colIndex: number,
      text: string,
      sheetIndex?: number): string
```

**Parameters**:

- `rowIndex`: `number`
- `colIndex`: `number`
- `text`: `string`
- `sheetIndex`: `number` (optional)

#### deleteSheet()

remove current sheet

```typescript
deleteSheet(): void
```

#### moveSheet()

move sheet to a new tab index

```typescript
moveSheet(fromIndex: number, toIndex: number): this
```

**Parameters**:

- `fromIndex`: `number`
  - current sheet index
- `toIndex`: `number`
  - target sheet index

#### insertRedactionForShape()

insert a redaction shape on a target shape/image

```typescript
insertRedactionForShape(reason: string, color: string, targetId: string, sheetName?: string): Promise<void>
```

**Parameters**:

- `reason`: `string`
  - redaction reason text
- `color`: `string`
  - background color
- `targetId`: `string`
  - target shape/image id
- `sheetName`: `string` (optional)
  - sheet name, defaults to active sheet

#### insertRedactionForRange()

insert a redaction on a cell range

```typescript
insertRedactionForRange(reason: string, color: string, range: CellRange, sheetName?: string): Promise<void>
```

**Parameters**:

- `reason`: `string`
  - redaction reason text
- `color`: `string`
  - background color
- `range`: `CellRange`
  - cell range {sri, sci, eri, eci}
- `sheetName`: `string` (optional)
  - sheet name, defaults to active sheet

#### removeRedaction()

remove a redaction by id

```typescript
removeRedaction(id: string, sheetName?: string): Promise<void>
```

**Parameters**:

- `id`: `string`
  - redaction id
- `sheetName`: `string` (optional)
  - sheet name, defaults to active sheet

#### syncRedactionOprClient()

sync redaction operations from history records

```typescript
syncRedactionOprClient(historyOprArray: any[], isSyncToServer?: boolean): Promise<void>
```

**Parameters**:

- `historyOprArray`: `any[]`
  - array of operation history records
- `isSyncToServer`: `boolean` (optional)
  - whether to sync to server

#### burnAllRedactions()

burn all redactions permanently

```typescript
burnAllRedactions(): void
```

#### clearRedactionClient()

clear all redaction shapes on a sheet or sheets array

```typescript
clearRedactionClient(sheetNameOrArray?: string | string[] | null, isSyncToServer?: boolean): Promise<void>
```

**Parameters**:

- `sheetNameOrArray`: `string | string[] | null` (optional)
  - sheet name, array of sheet names, or null for active sheet
- `isSyncToServer`: `boolean` (optional)
  - whether to sync to server

#### loadData()

load data

```typescript
loadData(json: Record<string, any>, activeSheetName?: string): this
```

**Parameters**:

- `json`: `Record<string, any>`
- `activeSheetName`: `string` (optional)
  - optional sheet name to activate after loading

#### setActiveSheet()

set active sheet by index

```typescript
setActiveSheet(id: number, isReActive?: boolean): this
```

**Parameters**:

- `id`: `number`
  - sheet index
- `isReActive`: `boolean` (optional)
  - force re-activate even if already active

#### setActiveSheetByName()

set active sheet by name

```typescript
setActiveSheetByName(sheetname: string, isReActive?: boolean): this
```

**Parameters**:

- `sheetname`: `string`
  - sheet name
- `isReActive`: `boolean` (optional)
  - force re-activate even if already active

#### setActiveCell()

set active cell

```typescript
setActiveCell(rowIndex: number, colIndex: number, scrollToCell?: boolean): this
```

**Parameters**:

- `rowIndex`: `number`
  - row index
- `colIndex`: `number`
  - column index
- `scrollToCell`: `boolean` (optional)
  - whether to scroll the active cell into view, defaults to true

#### getData()

get data

```typescript
getData(): Record<string, any>
```

#### getUpdateDatas()

get pending server update payloads

```typescript
getUpdateDatas(): Record<string, any>[]
```

#### refreshToken()

update authorization token used by server requests

```typescript
refreshToken(token: string): void
```

**Parameters**:

- `token`: `string`

#### destroy()

release GridJS DOM and runtime resources

```typescript
destroy(): void
```

#### setUniqueId()

configure server-side workbook id

```typescript
setUniqueId(uid: string): void
```

**Parameters**:

- `uid`: `string`

#### setFileName()

configure display/download file name

```typescript
setFileName(fname: string): void
```

**Parameters**:

- `fname`: `string`

#### setImageInfo()

configure image endpoints

```typescript
setImageInfo(imageUrl: string,
      uploadByLocalUrl: string,
      uploadByUrlUrl: string,
      copyUrl: string,
      zorder?: number,
      loadingGifUrl?: string): void
```

**Parameters**:

- `imageUrl`: `string`
- `uploadByLocalUrl`: `string`
- `uploadByUrlUrl`: `string`
- `copyUrl`: `string`
- `zorder`: `number` (optional)
- `loadingGifUrl`: `string` (optional)

#### setFileDownloadInfo()

configure file download endpoint

```typescript
setFileDownloadInfo(url: string): void
```

**Parameters**:

- `url`: `string`

#### setOleDownloadInfo()

configure OLE download endpoint

```typescript
setOleDownloadInfo(url: string): void
```

**Parameters**:

- `url`: `string`

#### setLazyLoadingUrl()

configure lazy loading endpoint

```typescript
setLazyLoadingUrl(url: string): void
```

**Parameters**:

- `url`: `string`

#### updateServerError()

bind server update error handler

```typescript
updateServerError(callback: (...args: any[]) => void): this
```

**Parameters**:

- `callback`: `(...args: any[]) => void`

#### updateCellError()

bind update-cell error handler

```typescript
updateCellError(callback: (...args: any[]) => void): this
```

**Parameters**:

- `callback`: `(...args: any[]) => void`

#### change()

bind handler to change event, including data change and user actions

```typescript
change(callback: (json: Record<string, any>) => void): this
```

**Parameters**:

- `callback`: `(json: Record<string, any>) => void`

#### locale()

set locale

```typescript
locale(lang: string, message: string): void
```

**Parameters**:

- `lang`: `string`
- `message`: `string`

### Properties

- **opts**: `Options)` (optional)
- **version**: `string`
  - Build version string (injected at compile time)
- **semantic**: `SemanticFacade`
  - Public Semantic Automation facade.
- **on**: `SpreadsheetEventHandler`

## Options Interface

Configuration options for initializing the Spreadsheet.

```typescript
interface Options {
    mode?: 'edit' | 'read';
    showToolbar?: boolean;
    showGrid?: boolean;
    showContextmenu?: boolean;
    showFileName?: boolean;
    local?: string;
    locale?: string;
    updateMode?: 'server' | 'client' | string;
    updateUrl?: string;
    token?: string;
    loadingGif?: string;
    showPartToolbar?: boolean;
}
```

### Option Details

- **mode**: `'edit' | 'read'` (optional)
- **showToolbar**: `boolean` (optional)
- **showGrid**: `boolean` (optional)
- **showContextmenu**: `boolean` (optional)
- **showFileName**: `boolean` (optional)
- **local**: `string` (optional)
- **locale**: `string` (optional)
- **updateMode**: `'server' | 'client' | string` (optional)
- **updateUrl**: `string` (optional)
- **token**: `string` (optional)
- **loadingGif**: `string` (optional)
- **showPartToolbar**: `boolean` (optional)

## Other Interfaces

### CellData

Data for representing a cell

- **text**: `string`
- **style**: `number` (optional)
- **merge**: `CellMerge` (optional)

### CellRange

- **sri**: `number`
- **sci**: `number`
- **eri**: `number`
- **eci**: `number`

### CellStyle

- **align**: `'left' | 'center' | 'right'` (optional)
- **valign**: `'top' | 'middle' | 'bottom'` (optional)
- **bold**: `boolean` (optional)

### ColProperties

- **width**: `number` (optional)

### CustomToolbarButton

Runtime instance of a custom toolbar button (passed to onClick and `custom-button` event)

- **tag**: `string`
  - Unique tag identifier
- **tip**: `string`
  - Tooltip text
- **spreadsheet**: `Spreadsheet`
  - Reference to the owning Spreadsheet instance (injected by Toolbar)
- **config**: `CustomToolbarButtonConfig`
  - Original config passed in via Options.customToolbarButtons

### CustomToolbarButtonConfig

Configuration for a single custom toolbar button

- **tag**: `string`
  - Unique identifier; forwarded as the second argument of the `custom-button` change event
- **tooltip**: `string` (optional)
  - Tooltip shown on hover
- **icon**: `CustomToolbarButtonIcon` (optional)
  - Icon descriptor (url / className / html / text)
- **active**: `boolean` (optional)
  - Click handler; `button` is the CustomToolbarButton instance, `spreadsheet` is the Spreadsheet instance */ onClick?: (button: CustomToolbarButton, spreadsheet: Spreadsheet) => void; Initial active state
- **disabled**: `boolean` (optional)
  - Initial disabled state
- **width**: `number` (optional)
  - Optional fixed button width in px

### CustomToolbarButtonIcon

Icon descriptor for a custom toolbar button (choose one form)

- **url**: `string` (optional)
  - Image URL used as background-image
- **className**: `string` (optional)
  - CSS class applied to the inner icon element (user supplies CSS)
- **html**: `string` (optional)
  - Raw HTML fragment, e.g. inline <svg>...</svg>
- **text**: `string` (optional)
  - Plain text / emoji, e.g. '★'
- **width**: `number` (optional)
  - Width in px when using `url` (default: 16)
- **height**: `number` (optional)
  - Height in px when using `url` (default: 16)

### Editor

- **opts**: `Options)` (optional)
- **version**: `string`
  - Build version string (injected at compile time)
- **semantic**: `SemanticFacade`
  - Public Semantic Automation facade.
- **on**: `SpreadsheetEventHandler`

### HeaderFont

Shared canvas render font for both column headers (A, B, C, ...) and row headers (1, 2, 3, ...). Fields in colHeaderFont / rowHeaderFont override it field by field. headerFont?: HeaderFont; Canvas render font for column headers (A, B, C, ...), overrides headerFont */ colHeaderFont?: HeaderFont; Canvas render font for row headers (1, 2, 3, ...), overrides headerFont */ rowHeaderFont?: HeaderFont; Enable redaction feature */ enableRedactionShape?: boolean; Predefined redaction reasons */ redactionReasons?: string[]; Default redaction color */ redactionDefaultColor?: string; Custom toolbar buttons injected before the "more" dropdown. Each item defines a tag, tooltip, icon, and onClick handler. customToolbarButtons?: CustomToolbarButtonConfig[]; Custom callback for adding new redaction reason. If provided, replaces the default modal dialog. onRedactionAddReason?: (existingReasons: string[]) => Promise<string | null>; Versioned Object API for AI Agent and automated tests. Disabled by default. */ semanticAutomation?: SemanticAutomationOptions; } export interface SemanticAutomationOptions { enabled?: boolean; instanceId?: string; } export type SemanticRuntimePhase = | 'disabled' | 'initializing' | 'loading' | 'rendering' | 'ready' | 'updating' | 'failed' | 'destroyed'; export interface SemanticRuntimeState { contractVersion: string; instanceId: string; instanceGeneration: string; workbookSessionId: string | null; phase: SemanticRuntimePhase; revision: number; failure: { code: string; message: string; retryable: boolean } | null; activeOperationIds: string[]; lastOperationId: string | null; activeSheetId: string | null; } export interface SemanticOperation { operationId: string; kind: string; action: string; instanceGeneration: string; workbookSessionId: string | null; parentOperationId: string | null; phase: 'pending' | 'running' | 'completed' | 'failed' | 'cancelled'; progress: number; revision: number | null; result: Record<string, any> | null; error: { code: string; message: string; retryable: boolean } | null; } export interface SemanticStateRequest { kind: 'workbook' | 'sheets' | 'sheet' | 'cell' | 'range' | 'selection'; sheetId?: string; address?: string; range?: string; offset?: number; limit?: number; } export interface SemanticStateSnapshot { contractVersion: string; instanceId: string; instanceGeneration: string; workbookSessionId: string; revision: number; target: Record<string, any>; state: Record<string, any>; } export type SemanticCellValue = | { kind: 'blank' } | { kind: 'text'; value: string } | { kind: 'number'; value: number } | { kind: 'boolean'; value: boolean }; export type SemanticEditableCellValue = | { kind: 'blank' } | { kind: 'text'; value: string }; export type SemanticActionRequest = | { action: 'sheet.activate'; parameters: { sheetId: string }; invocationId?: string } | { action: 'cell.activate'; parameters: { sheetId: string; address: string }; invocationId?: string } | { action: 'range.select'; parameters: { sheetId: string; address?: string; range?: string }; invocationId?: string } | { action: 'cell.setValue'; parameters: { sheetId: string; address: string; value: SemanticEditableCellValue }; invocationId?: string }; export interface SemanticActionReceipt { operationId: string; action: string; instanceGeneration: string; workbookSessionId: string; phase: 'running'; } export interface SemanticFacade { describeContract(): Record<string, any>; getCapabilities(): Record<string, any>; getContract(): Record<string, any>; getRuntimeState(): SemanticRuntimeState; waitForRuntime(options?: { phase?: SemanticRuntimePhase; timeoutMs?: number; signal?: AbortSignal }): Promise<SemanticRuntimeState>; subscribe(options: { event?: 'runtimechange' | 'operationchange' | 'sheetchange' | 'selectionchange' | 'cellchange' | '*'; fromRevision?: number; onEvent: (event: Record<string, any>) => void }): { unsubscribe(): void }; getOperation(operationId: string): SemanticOperation; listOperations(): SemanticOperation[]; waitForOperation(options: { operationId: string; timeoutMs?: number; signal?: AbortSignal }): Promise<SemanticOperation>; getState(request: SemanticStateRequest): SemanticStateSnapshot; query(request: SemanticStateRequest): SemanticStateSnapshot; getAvailableActions(request?: Pick<SemanticStateRequest, 'kind' | 'sheetId'>): ReadonlyArray<{ action: string; capability: string }>; perform(request: SemanticActionRequest): SemanticActionReceipt; } export interface SemanticInstanceDescriptor { instanceId: string; instanceGeneration: string; contractVersion: string; runtimePhase: SemanticRuntimePhase; workbookSessionId: string | null; } export interface GridJSSemanticAutomationRegistry { readonly version: '1.0.0'; listInstances(): SemanticInstanceDescriptor[]; getInstance(instanceId: string): SemanticFacade | null; } export type CELL_SELECTED = 'cell-selected'; export type CELLS_SELECTED = 'cells-selected'; export type CELL_EDITED = 'cell-edited'; export type OBJECT_SELECTED = 'object-selected'; export type SHEET_SELECTED = 'sheet-selected'; export type SHEET_LOADED = 'sheet-loaded'; export type VERTICAL_SCROLLED = 'vertical-scrolled'; export type HORIZONTAL_SCROLLED = 'horizontal-scrolled'; export type CELLS_DELETED = 'cells-deleted'; export type CELLS_UPDATED = 'cells-updated'; export type ROWS_INSERTED = 'rows-inserted'; export type COLUMNS_INSERTED = 'columns-inserted'; export type ROWS_DELETED = 'rows-deleted'; export type COLUMNS_DELETED = 'columns-deleted'; export type CHANGE = 'change'; export type REDACTION_BURNED = 'redaction-burned'; export type REDACTION_INSERTED = 'redaction-inserted'; export type REDACTION_DELETED = 'redaction-deleted'; export type REDACTION_UPDATED = 'redaction-updated'; export type REDACTION_REASON_INSERTED = 'redactionReason-inserted'; export type REDACTION_TRANSPARENCY_TOGGLED = 'redactionTransparency-toggled'; export type SEARCH_REDACTED = 'search-redacted'; export type CUSTOM_BUTTON = 'custom-button'; export type UPDATE_SERVER_ERROR = 'updateServerError'; export type UPDATE_CELL_ERROR = 'updateCellError'; export type CellMerge = [number, number]; Font setting for fixed headers rendered on canvas

- **name**: `string` (optional)
  - Font family name, e.g. 'Arial'
- **size**: `number` (optional)
  - Font size in px (canvas font size before dpr scaling)
- **weight**: `number | 'normal' | 'bold' | 'lighter' | 'bolder'` (optional)
  - Font weight, e.g. 400 | 500 | 'bold'
- **bold**: `boolean` (optional)
  - Shorthand for weight: 'bold'
- **italic**: `boolean` (optional)
- **color**: `string` (optional)
  - Text color

### RedactionData

- **id**: `string`
- **originId**: `string` (optional)
- **redactionReason**: `string` (optional)
- **redactionLabel**: `string` (optional)
- **redactionColor**: `string` (optional)
- **bgColor**: `string` (optional)
- **targetType**: `string` (optional)
- **targetId**: `string` (optional)
- **isRedaction**: `boolean` (optional)
- **fabobj**: `any` (optional)

### RowData

Data for representing a row

### SheetData

Data for representing a sheet

- **name**: `string` (optional)
- **freeze**: `string` (optional)
- **styles**: `CellStyle[]` (optional)
- **merges**: `string[]` (optional)
- **len**: `number` (optional)

### SpreadsheetData

Data for representing a spreadsheet

### SpreadsheetEventHandler

- **id**: `string`
- **type**: `string`
- **left**: `number`
- **top**: `number`
- **width**: `number`

## Type Aliases

### SemanticActionRequest

Values:
- `{ action: 'sheet.activate`

### SemanticCellValue

Values:
- `{ kind: 'blank' }`
- `{ kind: 'text`

### SemanticEditableCellValue

Values:
- `{ kind: 'blank' }`
- `{ kind: 'text`

### SemanticRuntimePhase

Values:
- `disabled`
- `initializing`
- `loading`
- `rendering`
- `ready`
- `updating`
- `failed`
- `destroyed`

