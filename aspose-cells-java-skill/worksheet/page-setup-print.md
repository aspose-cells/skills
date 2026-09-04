---
name: page-setup-print
description: Configure page setup, margins, header / footer, print area, page breaks, fit-to-page, paper size, and PDF orientation in Aspose.Cells for Java. Use whenever the user asks about printing, paper size, header text, or print scaling.
applies_to: Java
keywords: [page setup, print, margins, header, footer, fit to page, paper size, orientation]
---

# Page setup & print

## Imports

```java
import com.aspose.cells.*;
```

## Quick example — set up one sheet for printing

```java
Worksheet ws = wb.getWorksheets().get(0);
PageSetup ps = ws.getPageSetup();

ps.setOrientation(PageOrientationType.LANDSCAPE);
ps.setPaperSize(PaperSizeType.PAPER_A_4);
ps.setFitToPages(1, 0);                          // (wide, tall); 0 = any

ps.setTopMargin(1.0);                           // margins are in inches
ps.setBottomMargin(0.5);
ps.setLeftMargin(0.75);
ps.setRightMargin(0.75);

ps.setHeader(0, "&\"Calibri,Bold\"&14 Quarterly Report");  // page header centre
ps.setFooter(0, "&P of &N");                                 // page footer centre
```

The `&` codes mirror Excel's header/footer markup: `&P` page number,
`&N` total, `&D` date, `&T` time, `&A` sheet name, `&F` file name.

## Print area & page breaks

```java
Worksheet ws = wb.getWorksheets().get(0);
PageSetup ps = ws.getPageSetup();
ps.setPrintArea("A1:N50");
ps.setPrintTitleRows("$1:$2");                  // repeat first 2 rows on every page

ws.getHorizontalPageBreaks().add("A21");       // break before row 21
```

## Fit to page

```java
Worksheet ws = wb.getWorksheets().get(0);
PageSetup ps = ws.getPageSetup();
ps.setFitToPages(2, 3);                         // wide=2, tall=3  (0 = any)
```

When both fields are `0`, "Fit to page" is off; setting either makes
Excel scale to that ratio.

## Pitfalls

- **Margins units** — Excel internally stores margins in inches, so
  `setTopMargin(double)` already takes inches; for centimetres divide
  by 2.54 yourself.
- **`setFitToPages(wide, tall)`** — both arguments are required; pass `0`
  for the dimension that should scale freely.
- **Print titles** support only rows or only columns per call; pass
  them separately.
- **`PageOrientationType` and `PaperSizeType`** are enums but the
  underlying value is persisted in XLSX — verify round-trips if the
  user is mixing old and new workbooks.

## Related

- [manage-worksheets.md](manage-worksheets.md)
- [panes-zoom-views.md](panes-zoom-views.md)
- [rendering/convert-to-pdf.md](../rendering/convert-to-pdf.md)
