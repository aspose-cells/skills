# Client-Side Events

GridJs Spreadsheet provides an event system for handling user interactions.

## Event Types

### SemanticActionRequest

- `{ action: 'sheet.activate`

### SemanticCellValue

- `{ kind: 'blank' }`
- `{ kind: 'text`

### SemanticEditableCellValue

- `{ kind: 'blank' }`
- `{ kind: 'text`

### SemanticRuntimePhase

- `disabled`
- `initializing`
- `loading`
- `rendering`
- `ready`
- `updating`
- `failed`
- `destroyed`

## Event Handler Usage

```javascript
// Using the on() method
xs.on('cell-selected', (cell, rowIndex, colIndex) => {
    console.log(`Cell selected: ${rowIndex}, ${colIndex}`);
});

xs.on('cells-selected', (cell, range) => {
    console.log(`Range selected: ${range.sri},${range.sci} to ${range.eri},${range.eci}`);
});

xs.on('cell-edited', (text, rowIndex, colIndex) => {
    console.log(`Cell edited: ${text} at ${rowIndex},${colIndex}`);
});

xs.on('sheet-selected', (index, name) => {
    console.log(`Sheet selected: ${name}`);
});

xs.on('sheet-loaded', (index, name) => {
    console.log(`Sheet loaded: ${name}`);
});

xs.on('change', (json) => {
    console.log('Data changed:', json);
});

// Using shortcut methods
xs.change((json) => {
    console.log('Change event:', json);
});

xs.updateCellError((msg) => {
    console.error('Cell update error:', msg);
});
```

## Event Callback Signatures

| Event | Callback Parameters |
|-------|-------------------|
| `cell-selected` | `(cell: Cell, rowIndex: number, colIndex: number)` |
| `cells-selected` | `(cell: Cell, range: CellRange)` |
| `cell-edited` | `(text: string, rowIndex: number, colIndex: number)` |
| `object-selected` | `(obj: { id, type, left, top, width, height })` |
| `sheet-selected` | `(index: number, name: string)` |
| `sheet-loaded` | `(index: number, name: string)` |
| `vertical-scrolled` | `(rowIndex: number)` |
| `horizontal-scrolled` | `(colIndex: number)` |
| `cells-deleted` | `(range: CellRange)` |
| `cells-updated` | `(sheetName: string, cells: any)` |
| `rows-inserted` | `(rowIndex: number, count: number)` |
| `columns-inserted` | `(colIndex: number, count: number)` |
| `rows-deleted` | `(rowIndex: number, count: number)` |
| `columns-deleted` | `(colIndex: number, count: number)` |
| `change` | `(...args: any[])` |
| `custom-button` | `(tag: string, button: CustomToolbarButton)` |
| `updateServerError` | `(...args: any[])` |
| `updateCellError` | `(...args: any[])` |
