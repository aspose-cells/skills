declare module 'gridjs-spreadsheet' {
  export interface Options {
    mode?: 'edit' | 'read';
    showToolbar?: boolean;
    showGrid?: boolean;
    showContextmenu?: boolean;
    showFileName?: boolean;
    local?:string;
    locale?: string;
    updateMode?: 'server' | 'client' | string;
    updateUrl?: string;
    token?: string;
    loadingGif?: string;
    showPartToolbar?: boolean;
    view?: {
      height: () => number;
      width: () => number;
    };
    row?: {
      len: number;
      height: number;
    };
    col?: {
      len: number;
      width: number;
      indexWidth: number;
      minWidth: number;
    };
    style?: {
      bgcolor: string;
      align: 'left' | 'center' | 'right';
      valign: 'top' | 'middle' | 'bottom';
      textwrap: boolean;
      strike: boolean;
      underline: boolean;
      color: string;
      font: {
        name: 'Helvetica';
        size: number;
        bold: boolean;
        italic: false;
      };
    };
    /**
     * Shared canvas render font for both column headers (A, B, C, ...)
     * and row headers (1, 2, 3, ...).
     * Fields in colHeaderFont / rowHeaderFont override it field by field.
     */
    headerFont?: HeaderFont;
    /** Canvas render font for column headers (A, B, C, ...), overrides headerFont */
    colHeaderFont?: HeaderFont;
    /** Canvas render font for row headers (1, 2, 3, ...), overrides headerFont */
    rowHeaderFont?: HeaderFont;
    /** Enable redaction feature */
    enableRedactionShape?: boolean;
    /** Predefined redaction reasons */
    redactionReasons?: string[];
    /** Default redaction color */
    redactionDefaultColor?: string;
    /**
     * Custom toolbar buttons injected before the "more" dropdown.
     * Each item defines a tag, tooltip, icon, and onClick handler.
     */
    customToolbarButtons?: CustomToolbarButtonConfig[];
    /**
     * Custom callback for adding new redaction reason.
     * If provided, replaces the default modal dialog.
     * @param existingReasons - Current list of existing reasons
     * @returns Promise<string> for the new reason, or null to cancel
     */
    onRedactionAddReason?: (existingReasons: string[]) => Promise<string | null>;
    /** Versioned Object API for AI Agent and automated tests. Disabled by default. */
    semanticAutomation?: SemanticAutomationOptions;
  }

  export interface SemanticAutomationOptions {
    enabled?: boolean;
    instanceId?: string;
  }

  export type SemanticRuntimePhase =
    | 'disabled'
    | 'initializing'
    | 'loading'
    | 'rendering'
    | 'ready'
    | 'updating'
    | 'failed'
    | 'destroyed';

  export interface SemanticRuntimeState {
    contractVersion: string;
    instanceId: string;
    instanceGeneration: string;
    workbookSessionId: string | null;
    phase: SemanticRuntimePhase;
    revision: number;
    failure: { code: string; message: string; retryable: boolean } | null;
    activeOperationIds: string[];
    lastOperationId: string | null;
    activeSheetId: string | null;
  }

  export interface SemanticOperation {
    operationId: string;
    kind: string;
    action: string;
    instanceGeneration: string;
    workbookSessionId: string | null;
    parentOperationId: string | null;
    phase: 'pending' | 'running' | 'completed' | 'failed' | 'cancelled';
    progress: number;
    revision: number | null;
    result: Record<string, any> | null;
    error: { code: string; message: string; retryable: boolean } | null;
  }

  export interface SemanticStateRequest {
    kind: 'workbook' | 'sheets' | 'sheet' | 'cell' | 'range' | 'selection';
    sheetId?: string;
    address?: string;
    range?: string;
    offset?: number;
    limit?: number;
  }

  export interface SemanticStateSnapshot {
    contractVersion: string;
    instanceId: string;
    instanceGeneration: string;
    workbookSessionId: string;
    revision: number;
    target: Record<string, any>;
    state: Record<string, any>;
  }

  export type SemanticCellValue =
    | { kind: 'blank' }
    | { kind: 'text'; value: string }
    | { kind: 'number'; value: number }
    | { kind: 'boolean'; value: boolean };

  export type SemanticEditableCellValue =
    | { kind: 'blank' }
    | { kind: 'text'; value: string };

  export type SemanticActionRequest =
    | { action: 'sheet.activate'; parameters: { sheetId: string }; invocationId?: string }
    | { action: 'cell.activate'; parameters: { sheetId: string; address: string }; invocationId?: string }
    | { action: 'range.select'; parameters: { sheetId: string; address?: string; range?: string }; invocationId?: string }
    | { action: 'cell.setValue'; parameters: { sheetId: string; address: string; value: SemanticEditableCellValue }; invocationId?: string };

  export interface SemanticActionReceipt {
    operationId: string;
    action: string;
    instanceGeneration: string;
    workbookSessionId: string;
    phase: 'running';
  }

  export interface SemanticFacade {
    describeContract(): Record<string, any>;
    getCapabilities(): Record<string, any>;
    getContract(): Record<string, any>;
    getRuntimeState(): SemanticRuntimeState;
    waitForRuntime(options?: { phase?: SemanticRuntimePhase; timeoutMs?: number; signal?: AbortSignal }): Promise<SemanticRuntimeState>;
    subscribe(options: { event?: 'runtimechange' | 'operationchange' | 'sheetchange' | 'selectionchange' | 'cellchange' | '*'; fromRevision?: number; onEvent: (event: Record<string, any>) => void }): { unsubscribe(): void };
    getOperation(operationId: string): SemanticOperation;
    listOperations(): SemanticOperation[];
    waitForOperation(options: { operationId: string; timeoutMs?: number; signal?: AbortSignal }): Promise<SemanticOperation>;
    getState(request: SemanticStateRequest): SemanticStateSnapshot;
    query(request: SemanticStateRequest): SemanticStateSnapshot;
    getAvailableActions(request?: Pick<SemanticStateRequest, 'kind' | 'sheetId'>): ReadonlyArray<{ action: string; capability: string }>;
    perform(request: SemanticActionRequest): SemanticActionReceipt;
  }

  export interface SemanticInstanceDescriptor {
    instanceId: string;
    instanceGeneration: string;
    contractVersion: string;
    runtimePhase: SemanticRuntimePhase;
    workbookSessionId: string | null;
  }

  export interface GridJSSemanticAutomationRegistry {
    readonly version: '1.0.0';
    listInstances(): SemanticInstanceDescriptor[];
    getInstance(instanceId: string): SemanticFacade | null;
  }

  export type CELL_SELECTED = 'cell-selected';
  export type CELLS_SELECTED = 'cells-selected';
  export type CELL_EDITED = 'cell-edited';
  export type OBJECT_SELECTED = 'object-selected';
  export type SHEET_SELECTED = 'sheet-selected';
  export type SHEET_LOADED = 'sheet-loaded';
  export type VERTICAL_SCROLLED = 'vertical-scrolled';
  export type HORIZONTAL_SCROLLED = 'horizontal-scrolled';
  export type CELLS_DELETED = 'cells-deleted';
  export type CELLS_UPDATED = 'cells-updated';
  export type ROWS_INSERTED = 'rows-inserted';
  export type COLUMNS_INSERTED = 'columns-inserted';
  export type ROWS_DELETED = 'rows-deleted';
  export type COLUMNS_DELETED = 'columns-deleted';
  export type CHANGE = 'change';
  export type REDACTION_BURNED = 'redaction-burned';
  export type REDACTION_INSERTED = 'redaction-inserted';
  export type REDACTION_DELETED = 'redaction-deleted';
  export type REDACTION_UPDATED = 'redaction-updated';
  export type REDACTION_REASON_INSERTED = 'redactionReason-inserted';
  export type REDACTION_TRANSPARENCY_TOGGLED = 'redactionTransparency-toggled';
  export type SEARCH_REDACTED = 'search-redacted';
  export type CUSTOM_BUTTON = 'custom-button';
  export type UPDATE_SERVER_ERROR = 'updateServerError';
  export type UPDATE_CELL_ERROR = 'updateCellError';

  export type CellMerge = [number, number];

  /** Font setting for fixed headers rendered on canvas */
  export interface HeaderFont {
    /** Font family name, e.g. 'Arial' */
    name?: string;
    /** Font size in px (canvas font size before dpr scaling) */
    size?: number;
    /** Font weight, e.g. 400 | 500 | 'bold' */
    weight?: number | 'normal' | 'bold' | 'lighter' | 'bolder';
    /** Shorthand for weight: 'bold' */
    bold?: boolean;
    italic?: boolean;
    /** Text color */
    color?: string;
  }

  export interface CellRange {
    sri: number;
    sci: number;
    eri: number;
    eci: number;
  }

  /** Icon descriptor for a custom toolbar button (choose one form) */
  export interface CustomToolbarButtonIcon {
    /** Image URL used as background-image */
    url?: string;
    /** CSS class applied to the inner icon element (user supplies CSS) */
    className?: string;
    /** Raw HTML fragment, e.g. inline <svg>...</svg> */
    html?: string;
    /** Plain text / emoji, e.g. '★' */
    text?: string;
    /** Width in px when using `url` (default: 16) */
    width?: number;
    /** Height in px when using `url` (default: 16) */
    height?: number;
  }

  /** Runtime instance of a custom toolbar button (passed to onClick and `custom-button` event) */
  export interface CustomToolbarButton {
    /** Unique tag identifier */
    tag: string;
    /** Tooltip text */
    tip: string;
    /** Reference to the owning Spreadsheet instance (injected by Toolbar) */
    spreadsheet: Spreadsheet;
    /** Original config passed in via Options.customToolbarButtons */
    config: CustomToolbarButtonConfig;
    /** Update the icon at runtime */
    setIcon(iconConfig: CustomToolbarButtonIcon): void;
    /** Update the tooltip text at runtime */
    setTooltip(text: string): void;
    /** Toggle active (highlighted) state */
    setActive(active: boolean): void;
    /** Toggle disabled state (click events suppressed) */
    setDisabled(disabled: boolean): void;
    /** Show the button */
    show(): void;
    /** Hide the button */
    hide(): void;
  }

  /** Configuration for a single custom toolbar button */
  export interface CustomToolbarButtonConfig {
    /** Unique identifier; forwarded as the second argument of the `custom-button` change event */
    tag: string;
    /** Tooltip shown on hover */
    tooltip?: string;
    /** Icon descriptor (url / className / html / text) */
    icon?: CustomToolbarButtonIcon;
    /** Click handler; `button` is the CustomToolbarButton instance, `spreadsheet` is the Spreadsheet instance */
    onClick?: (button: CustomToolbarButton, spreadsheet: Spreadsheet) => void;
    /** Initial active state */
    active?: boolean;
    /** Initial disabled state */
    disabled?: boolean;
    /** Optional fixed button width in px */
    width?: number;
  }

  export interface RedactionData {
    id: string;
    originId?: string;
    redactionReason?: string;
    redactionLabel?: string;
    redactionColor?: string;
    bgColor?: string;
    targetType?: string;
    targetId?: string;
    isRedaction?: boolean;
    fabobj?: any;
  }

  export interface SpreadsheetEventHandler {
    (
      envt: CELL_SELECTED,
      callback: (cell: Cell, rowIndex: number, colIndex: number) => void
    ): void;
    (
      envt: CELLS_SELECTED,
      callback: (
        cell: Cell,
        range: CellRange
      ) => void
    ): void;
    (
      envt: OBJECT_SELECTED,
      callback: (obj: { id: string; type: string; left: number; top: number; width: number; height: number }) => void
    ): void;
    (
      envt: SHEET_SELECTED,
      callback: (index: number, name: string) => void
    ): void;
    (
      envt: CELL_EDITED,
      callback: (text: string, rowIndex: number, colIndex: number) => void
    ): void;
    (
      envt: SHEET_LOADED,
      callback: (index: number, name: string) => void
    ): void;
    (
      envt: VERTICAL_SCROLLED,
      callback: (rowIndex: number) => void
    ): void;
    (
      envt: HORIZONTAL_SCROLLED,
      callback: (colIndex: number) => void
    ): void;
    (
      envt: CELLS_DELETED,
      callback: (range: CellRange) => void
    ): void;
    (
      envt: CELLS_UPDATED,
      callback: (sheetName: string, cells: any) => void
    ): void;
    (
      envt: ROWS_INSERTED,
      callback: (rowIndex: number, count: number) => void
    ): void;
    (
      envt: COLUMNS_INSERTED,
      callback: (colIndex: number, count: number) => void
    ): void;
    (
      envt: ROWS_DELETED,
      callback: (rowIndex: number, count: number) => void
    ): void;
    (
      envt: COLUMNS_DELETED,
      callback: (colIndex: number, count: number) => void
    ): void;
    (
      envt: CHANGE,
      callback: (...args: any[]) => void
    ): void;
    (
      envt: REDACTION_BURNED,
      callback: () => void
    ): void;
    (
      envt: REDACTION_INSERTED,
      callback: (sheetName: string, redactionData: RedactionData) => void
    ): void;
    (
      envt: REDACTION_DELETED,
      callback: (sheetName: string, redactionData: RedactionData) => void
    ): void;
    (
      envt: REDACTION_UPDATED,
      callback: (sheetName: string, shape: any) => void
    ): void;
    (
      envt: REDACTION_REASON_INSERTED,
      callback: (reason: string) => void
    ): void;
    (
      envt: REDACTION_TRANSPARENCY_TOGGLED,
      callback: (isTransparent: boolean) => void
    ): void;
    (
      envt: SEARCH_REDACTED,
      callback: (sheetNames: string[], detail: { keywords: string, isCaseSensitive: boolean, matchWholeCell: boolean, reason: string, color: string, baseId: string }) => void
    ): void;
    (
      envt: CUSTOM_BUTTON,
      callback: (tag: string, button: CustomToolbarButton) => void
    ): void;
    (
      envt: UPDATE_SERVER_ERROR,
      callback: (...args: any[]) => void
    ): void;
    (
      envt: UPDATE_CELL_ERROR,
      callback: (...args: any[]) => void
    ): void;
  }

  export interface ColProperties {
    width?: number;
  }

  /**
   * Data for representing a cell
   */
  export interface CellData {
    text: string;
    style?: number;
    merge?: CellMerge;
  }
  /**
   * Data for representing a row
   */
  export interface RowData {
    cells: {
      [key: number]: CellData;
    }
  }

  /**
   * Data for representing a sheet
   */
  export interface SheetData {
    name?: string;
    freeze?: string;
    styles?: CellStyle[];
    merges?: string[];
    cols?: {
      len?: number;
      [key: number]: ColProperties;
    };
    rows?: {
      [key: number]: RowData
    };
  }

  /**
   * Data for representing a spreadsheet
   */
  export interface SpreadsheetData {
    [index: number]: SheetData;
  }

  export interface CellStyle {
    align?: 'left' | 'center' | 'right';
    valign?: 'top' | 'middle' | 'bottom';
    font?: {
      bold?: boolean;
    }
    bgcolor?: string;
    textwrap?: boolean;
    color?: string;
    border?: {
      top?: string[];
      right?: string[];
      bottom?: string[];
      left?: string[];
    };
  }
  export interface Editor {}
  export interface Element {}

  export interface Row {}
  export interface Table {}
  export interface Cell {}
  export interface Sheet {}

  export default class Spreadsheet {
    constructor(container: string | HTMLElement, opts?: Options);
    /** Build version string (injected at compile time) */
    readonly version: string;
    /** Public Semantic Automation facade. */
    semantic: SemanticFacade;
    on: SpreadsheetEventHandler;
    /**
     * retrieve cell
     * @param rowIndex {number} row index
     * @param colIndex {number} column index
     * @param sheetIndex {number} sheet iindex
     */
    cell(rowIndex: number, colIndex: number, sheetIndex: number): Cell;
    /**
     * retrieve cell style
     * @param rowIndex
     * @param colIndex
     * @param sheetIndex
     */
    cellStyle(
      rowIndex: number,
      colIndex: number,
      sheetIndex: number
    ): CellStyle;
    /**
     * get/set cell text
     * @param rowIndex
     * @param colIndex
     * @param text
     * @param sheetIndex
     */
    cellText(
      rowIndex: number,
      colIndex: number,
      text: string,
      sheetIndex?: number
    ): string;
    /**
     * remove current sheet
     */
    deleteSheet(): void;
    /**
     * move sheet to a new tab index
     * @param fromIndex current sheet index
     * @param toIndex target sheet index
     */
    moveSheet(fromIndex: number, toIndex: number): this;

    /**
     * insert a redaction shape on a target shape/image
     * @param reason redaction reason text
     * @param color background color
     * @param targetId target shape/image id
     * @param sheetName sheet name, defaults to active sheet
     */
    insertRedactionForShape(reason: string, color: string, targetId: string, sheetName?: string): Promise<void>;

    /**
     * insert a redaction on a cell range
     * @param reason redaction reason text
     * @param color background color
     * @param range cell range {sri, sci, eri, eci}
     * @param sheetName sheet name, defaults to active sheet
     */
    insertRedactionForRange(reason: string, color: string, range: CellRange, sheetName?: string): Promise<void>;

    /**
     * remove a redaction by id
     * @param id redaction id
     * @param sheetName sheet name, defaults to active sheet
     */
    removeRedaction(id: string, sheetName?: string): Promise<void>;

    /**
     * sync redaction operations from history records
     * @param historyOprArray array of operation history records
     * @param isSyncToServer whether to sync to server
     */
    syncRedactionOprClient(historyOprArray: any[], isSyncToServer?: boolean): Promise<void>;

    /**
     * burn all redactions permanently
     */
    burnAllRedactions(): void;

    /**
     * clear all redaction shapes on a sheet or sheets array
     * @param sheetNameOrArray sheet name, array of sheet names, or null for active sheet
     * @param isSyncToServer whether to sync to server
     */
    clearRedactionClient(sheetNameOrArray?: string | string[] | null, isSyncToServer?: boolean): Promise<void>;

    /**
     * load data
     * @param json
     * @param activeSheetName optional sheet name to activate after loading
     */
    loadData(json: Record<string, any>, activeSheetName?: string): this;
    /**
     * set active sheet by index
     * @param id sheet index
     * @param isReActive force re-activate even if already active
     */
    setActiveSheet(id: number, isReActive?: boolean): this;
    /**
     * set active sheet by name
     * @param sheetname sheet name
     * @param isReActive force re-activate even if already active
     */
    setActiveSheetByName(sheetname: string, isReActive?: boolean): this;
    /**
     * set active cell
     * @param rowIndex row index
     * @param colIndex column index
     * @param scrollToCell whether to scroll the active cell into view, defaults to true
     */
    setActiveCell(rowIndex: number, colIndex: number, scrollToCell?: boolean): this;
    /**
     * get data
     */
    getData(): Record<string, any>;
    /**
     * get pending server update payloads
     */
    getUpdateDatas(): Record<string, any>[];
    /**
     * update authorization token used by server requests
     */
    refreshToken(token: string): void;
    /**
     * release GridJS DOM and runtime resources
     */
    destroy(): void;
    /**
     * configure server-side workbook id
     */
    setUniqueId(uid: string): void;
    /**
     * configure display/download file name
     */
    setFileName(fname: string): void;
    /**
     * configure image endpoints
     */
    setImageInfo(
      imageUrl: string,
      uploadByLocalUrl: string,
      uploadByUrlUrl: string,
      copyUrl: string,
      zorder?: number,
      loadingGifUrl?: string
    ): void;
    /**
     * configure file download endpoint
     */
    setFileDownloadInfo(url: string): void;
    /**
     * configure OLE download endpoint
     */
    setOleDownloadInfo(url: string): void;
    /**
     * configure lazy loading endpoint
     */
    setLazyLoadingUrl(url: string): void;
    /**
     * bind server update error handler
     */
    updateServerError(callback: (...args: any[]) => void): this;
    /**
     * bind update-cell error handler
     */
    updateCellError(callback: (...args: any[]) => void): this;
    /**
     * bind handler to change event, including data change and user actions
     * @param callback
     */
    change(callback: (json: Record<string, any>) => void): this;
    /**
     * set locale
     * @param lang
     * @param message
     */
    locale(lang: string, message: string): void;
  }
  global {
    interface Window {
      x_spreadsheet(container: string | HTMLElement, opts?: Options): Spreadsheet; 
      GridJSSemanticAutomation?: GridJSSemanticAutomationRegistry;
    }
  }
}
