---
name: large-files-memory
description: Avoid java.lang.OutOfMemoryError and slow random-access when loading or writing large workbooks (>100k rows) with Aspose.Cells for Java. Use MemorySetting.MEMORY_PREFERENCE for read or LightCells API for write/stream.
applies_to: Java
keywords: [oom, outofmemory, memory, lightcells, interrupt, large, performance]
---

# Large files & memory

Default `new Workbook("big.xlsx")` loads the whole sheet graph into
RAM and fails with `OutOfMemoryError` for files with hundreds of
thousands of rows, many styles, or many images.

## Choose your strategy

| Workload                                    | Strategy                          |
| ------------------------------------------- | --------------------------------- |
| Read large file, iterate row-by-row         | `MemorySetting.MEMORY_PREFERENCE` |
| Read large file, random cell access         | `MemorySetting.NORMAL` + `-Xmx8g`  |
| Write/stream large file                     | `LightCellsDataProvider`          |
| Long-running conversion, want to cancel     | `InterruptMonitor`                |

## Read: `MemorySetting.MEMORY_PREFERENCE`

```java
import com.aspose.cells.*;

LoadOptions opt = new LoadOptions();
opt.setMemorySetting(MemorySetting.MEMORY_PREFERENCE);
Workbook wb = new Workbook("big.xlsx", opt);
```

Traverse row-by-row, do not call `cell.getValue()` randomly.

## Write: LightCells API

```java
LightCellsDataProvider provider = new LightCellsDataProvider() {
    @Override public boolean startSheet(int sheetIndex) { return true; /* keep going */ }
    @Override public int    nextRow()                   { return 0; /* rowIndex, -1 = end */ }
    @Override public void   startRow(Row row)           { /* configure */ }
    @Override public int    nextCell()                  { return 0; /* column index, -1 = end */ }
    @Override public void   startCell(Cell cell)        { /* setValue / setFormula */ }
    @Override public boolean isGatherString()           { return false; }
};

OoxmlSaveOptions so = new OoxmlSaveOptions();
so.setLightCellsDataProvider(provider);
Workbook wb = new Workbook();
wb.save("out.xlsx", so);
```

For reading huge files use `LightCellsDataHandler` instead of the
default model.

## Cancel a long-running operation

```java
InterruptMonitor monitor = new InterruptMonitor();
Thread t = new Thread(() -> {
    try {
        Workbook wb = new Workbook("huge.xlsx");
        wb.save("out.pdf");
    } catch (Exception e) { /* swallow */ }
});
t.start();

new Thread(() -> {
    if (!monitor.isInterruptionRequested()) {
        try {
            Thread.sleep(30_000);            // hard timeout
            monitor.interrupt();
        } catch (InterruptedException e) { /* swallow */ }
    }
}).start();
t.join();
```

If `monitor` was passed via `LoadOptions.setInterruptMonitor(...)`,
both load and save honour the interrupt.

## Pitfalls

- **`MEMORY_PREFERENCE` + random access** → extremely slow. Don't
  combine the two.
- **Styles/images** keep memory pressure: call
  `Workbook.removeUnusedStyles()` / purge unused
  `ConditionalFormattingCollection` before saving huge files.
- **Threaded read** — `Workbook` is not thread-safe; use one
  `Workbook` per thread or per-thread copy.
- **Don't dispose `LoadOptions`** before passing to `Workbook`.
