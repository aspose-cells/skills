# Semantic Automation API

The Semantic Automation API provides a versioned, structured interface for AI agents and automated tests to interact with GridJs.

> **Note**: This feature is disabled by default. Enable it via `Options.semanticAutomation`.

## SemanticActionRequest

**Kind**: type

## SemanticCellValue

**Kind**: type

## SemanticEditableCellValue

**Kind**: type

## SemanticRuntimePhase

**Kind**: type

## Usage Example

```javascript
// Enable semantic automation
const xs = x_spreadsheet('#container', {
    semanticAutomation: { enabled: true, instanceId: 'my-instance' }
});

// Access the semantic facade
const semantic = xs.semantic;

// Get runtime state
const state = semantic.getRuntimeState();
console.log(state.phase); // 'ready'

// Query cell value
const cellState = semantic.getState({ kind: 'cell', address: 'A1' });

// Perform action
const receipt = semantic.perform({
    action: 'cell.setValue',
    parameters: { sheetId: 'sheet1', address: 'A1', value: { kind: 'text', value: 'Hello' } }
});

// Wait for operation to complete
const result = await semantic.waitForOperation({ operationId: receipt.operationId });
```
