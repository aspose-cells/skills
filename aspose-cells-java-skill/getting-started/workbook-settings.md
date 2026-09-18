---
name: workbook-settings
description: Workbook-level settings for Aspose.Cells for Java — calculation chain, date base (1904 system), shared/track-change mode, default font, compliance mode, write protection, built-in/custom properties.
applies_to: Java
keywords: [workbook settings, calculation, properties, 1904, shared, metadata]
---

# Workbook settings

`Workbook.getSettings()` exposes `WorkbookSettings` — globals that
apply across every sheet in the workbook.

## Imports

```java
import com.aspose.cells.*;
```

## Quick example — read & write at once

```java
Workbook wb = new Workbook();
WorkbookSettings s = wb.getSettings();

// Re-calculate on open (Excel does this after editing formulas)
s.getFormulaSettings().setForceFullCalculation(true);
// Recalculate formulas now
wb.calculateFormula();
```

## Common settings

| Setting                                  | When to use                                |
| ---------------------------------------- | ------------------------------------------ |
| `getFormulaSettings().setForceFullCalculation(true)` | Workbook opened in Excel triggers recalc  |
| `getFormulaSettings().setPrecisionAsDisplayed(true)` | Round stored values to the displayed precision |
| `setDate1904(true)`                      | Match Excel-for-Mac 1904 date base        |
| `getFormulaSettings().setEnableIterativeCalculation(true)` | Allow iterative (circular) calculations |
| `getFormulaSettings().getForceFullCalculation()`     | Read back whether full recalc is enabled   |
| `getWriteProtection().setRecommendReadOnly(true)` (+ `.setPassword(p)`) | Recommend open-as-read-only (and optionally password-protect) |
| `setCheckExcelRestriction()`             | Reject unknown XLSX parts (safer round-trip)|

## Calculation options

```java
CalculationOptions co = new CalculationOptions();
co.setIgnoreError(true);                       // do not propagate #REF etc.
co.setCalcStackSize(64);                       // parallel calc thread stack
co.setCustomEngine(new AbstractCalculationEngine() {
    @Override public void calculate(CalculationData data) {
        if ("MY_UDF".equals(data.getFunctionName())) {
            data.setCalculatedValue(data.getParamValue(0)); // first arg
        }
    }
});
wb.calculateFormula(co);
```

## Workbook properties

```java
BuiltInDocumentPropertyCollection builtIn = wb.getBuiltInDocumentProperties();
builtIn.setTitle("Q2 Sales");
builtIn.setAuthor("reports-bot");
builtIn.setCreatedTime(DateTime.getNow());

CustomDocumentPropertyCollection custom = wb.getCustomDocumentProperties();
custom.add("Reviewed", true);
custom.add("Source", "DataWarehouse");
```

## Pitfalls

- **Don't enable iteration unless you really need circular refs** —
  Excel shows a warning, and recalc may never converge.
- **`setDate1904(true)` is permanent** for the workbook — switching
  afterwards shifts every date by 4 years.
- **Calculation options apply per workbook** — they don't follow
  between `Workbook` instances.

## Related

- [_shared/license.md](../_shared/license.md)
- [formulas/calculate-formulas.md](../formulas/calculate-formulas.md)
- [workbook-protection is in security/](../security/encryption-decryption.md)
