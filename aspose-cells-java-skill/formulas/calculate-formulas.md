---
name: calculate-formulas
description: Recalculate all or selective formulas in Aspose.Cells for Java — workbook.calculateFormula, worksheet.calculateFormula, formula calculation options, custom calculation engine, and ignoring errors.
applies_to: Java
keywords: [calculate, recalc, formula engine, custom engine, calculation options]
---

# Calculate formulas

`Workbook.calculateFormula()` is the recalculate-everything method.
Use it after editing any formula, before saving, and before reading
`cell.getValue()`.

## Imports

```java
import com.aspose.cells.*;
```

## Quick example

```java
Workbook wb = new Workbook("with-formulas.xlsx");

// recompute every formula in the workbook
wb.calculateFormula();
// or just one sheet:
CalculationOptions opts = new CalculationOptions();
opts.setIgnoreError(true);
wb.getWorksheets().get("Q1").calculateFormula(opts, true);
// compute a single expression against the workbook's data
Object v = wb.getWorksheets().get(0).calculateFormula("MAX(A1:A100)");
```

`Worksheet.calculateFormula(expression)` returns the calculated value
of a one-off expression; no cell is modified.

## Calculation options

```java
Workbook wb = new Workbook();
CalculationOptions opts = new CalculationOptions();
opts.setIgnoreError(true);                      // skip #REF! / #DIV/0! cells
opts.setCalcStackSize(32);                      // bigger evaluation stack
opts.setPrecisionStrategy(CalculationPrecisionStrategy.DECIMAL);

wb.calculateFormula(opts);
```

`setIgnoreError(true)` lets the workbook compute formulas even when
the workbook contains `#REF!` / `#DIV/0!` errors.

## Custom calculation engine (UDF / unsupported functions)

```java
public class MyEngine extends AbstractCalculationEngine {
    @Override public void calculate(CalculationData data) {
        if (data.getFunctionName().equals("MY_UDF")) {
            // replace the cell value with the first argument
            data.setCalculatedValue(data.getParamValue(0));
        }
    }
}
```

Register:

```java
Workbook wb = new Workbook();
CalculationOptions opts = new CalculationOptions();
opts.setCustomEngine(new AbstractCalculationEngine() {
    @Override public void calculate(CalculationData data) {
        if (data.getFunctionName().equals("MY_UDF")) {
            data.setCalculatedValue(data.getParamValue(0));
        }
    }
});
wb.calculateFormula(opts);
```

## Pitfalls

- **Recalculate BEFORE save** — failing to do so writes a workbook
  containing only the formula text, not the cached values.
- **Recalculation mode is a workbook setting** — switch it with
  `wb.getSettings().getFormulaSettings().setForceFullCalculation(true)`
  or the equivalent
  per-CalculationOptions flag when you need manual recalc behaviour.
- **Custom engines** must thread-safely return values; returning
  `null` makes the cell value `null`.
- **`setIgnoreError(true)`** is not a substitute for repairing
  circular references — calculation may produce nonsense numbers.

### Circular references compute to 0

Symptom: an intentionally circular formula (iterative model — e.g.
interest-on-balance, Euler integration) reads 0 in every cell.

Cause: iterative calculation is off by default, so the cycle cannot
converge and the engine bails out before any value stabilizes.

Fix: enable and bound iterative calculation, then `calculateFormula`.
`FormulaSettings` lives on the workbook's settings — `wb.getSettings()
.getFormulaSettings()`. Set `setEnableIterativeCalculation(true)`,
bound the iteration with `setMaxIteration` and `setMaxChange`.

```java
Workbook wb = new Workbook("circular.xlsx");
FormulaSettings fs = wb.getSettings().getFormulaSettings();
fs.setEnableIterativeCalculation(true);   // off by default → cycles resolve to 0
fs.setMaxIteration(100);
fs.setMaxChange(0.001);
wb.calculateFormula();
```

A `MaxChange` close to 0 forces tighter convergence at the cost of
more iterations; `MaxIteration` is the hard stop. Reach it without
converging and the cells hold the last value, not zero.

### Recalculation is slow — repeated full passes

Symptom: `calculateFormula()` is slow; a small edit forces a full
recompute.

Cause: no calculation chain is retained, so every pass re-evaluates
every formula. With a chain, only the cells whose inputs changed
(and their dependents) are recomputed.

Fix: enable the chain once.

```java
Workbook wb = new Workbook("big.xlsx");
wb.getSettings().getFormulaSettings().setEnableCalculationChain(true);
wb.calculateFormula();                       // first pass builds the chain
// ... later edits ...
wb.calculateFormula();                       // subsequent passes recalc only what changed
```

For many single-cell recalcs in a loop, pass `CalculationOptions`
with `setRecursive(false)` so dependents are not re-walked on every
call.

## Related

- [set-get-formula.md](set-get-formula.md)
- [array-r1c1-formula.md](array-r1c1-formula.md)
- [getting-started/workbook-settings.md](../getting-started/workbook-settings.md)
