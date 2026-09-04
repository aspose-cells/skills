---
name: document-properties
description: Read and write built-in (Title, Author, Subject, ...) and custom (typed) document properties on Aspose.Cells for Java workbooks. Use when the user wants to set metadata visible under File → Info → Properties, or to read metadata for a search/audit pipeline.
applies_to: Java
keywords: [document properties, BuiltInDocumentProperties, CustomDocumentProperties, metadata, Title, Author, custom property, DateTime, linked property]
---

# Document properties

Two collections hang off `Workbook`:

- `Workbook.getBuiltInDocumentProperties()` — a
  `BuiltInDocumentPropertyCollection` exposing the standard Excel
  properties (`Title`, `Author`, `Subject`, `Keywords`, `Comments`,
  `Category`, `Company`, `Manager`, `CreatedTime`, `LastSavedTime`,
  …). Access them via typed getters/setters on the collection.
- `Workbook.getCustomDocumentProperties()` — a
  `CustomDocumentPropertyCollection` for arbitrary name/value pairs.
  Values are typed (string, int, double, bool, `DateTime`); reading
  the wrong type throws `InvalidCastException`.

## Imports

```java
import com.aspose.cells.*;
```

## Built-in document properties

The built-in collection extends `DocumentPropertyCollection`. The
typed members (`getTitle()`, `getAuthor()`, `getSubject()`, …) are
read/write; mutations are persisted on save.

```java
Workbook wb = new Workbook();
BuiltInDocumentPropertyCollection bp = wb.getBuiltInDocumentProperties();

bp.setTitle("Q3 Report");
bp.setAuthor("Finance Team");
bp.setSubject("Quarterly Forecast");
bp.setKeywords("quarterly,forecast,2026");
bp.setComments("Prepared by Finance");
bp.setCategory("Reports");
bp.setCompany("Aspose Demo Co.");
bp.setManager("Jane Doe");

// Times use com.aspose.cells.DateTime (not java.util.Date).
bp.setCreatedTime(DateTime.getNow());
bp.setLastSavedTime(DateTime.getNow());

wb.save("out.xlsx");
```

For properties that don't have a typed setter, fall back to the
indexed `DocumentProperty` access:

```java
Workbook wb = new Workbook();
DocumentPropertyCollection bp = wb.getBuiltInDocumentProperties();

// Any name — even non-built-in — can be reached through the indexer.
bp.get("Application").setValue("Aspose.Cells");
wb.save("out.xlsx");
```

## Custom document properties

Use the typed `CustomDocumentPropertyCollection.add(name, value)`
overloads (`String`, `int`, `double`, `boolean`, `DateTime`).
Read back via `get(name)` and `.getValue()` (returns `Object`) — call
`toInt()`, `toDouble()`, `toBool()`, `toDateTime()` for typed access.

```java
Workbook wb = new Workbook();
CustomDocumentPropertyCollection props = wb.getCustomDocumentProperties();

props.add("Department", "Finance");          // String
props.add("Approved", true);                 // boolean
props.add("ApprovalCount", 3);               // int
props.add("Ratio", 0.85);                    // double
props.add("ApprovedOn", DateTime.getNow());  // DateTime

for (int i = 0; i < props.getCount(); i++) {
    DocumentProperty p = props.get(i);
    System.out.println(p.getName() + " = " + p.getValue());
}

wb.save("out.xlsx");
```

For typed reads:

```java
Workbook wb = new Workbook();
CustomDocumentPropertyCollection props = wb.getCustomDocumentProperties();
props.add("Approved", true);
props.add("ApprovalCount", 3);

DocumentProperty approved = props.get("Approved");
if (approved.toBool()) {
    System.out.println("Approved");
}
int n = props.get("ApprovalCount").toInt();
```

## Pitfalls

### Custom property "does not stick" or reads back as the wrong type

Symptom: a custom property added with a string is missing or reads
back as `null` / the wrong type.
Cause: `add(name, value)` has typed overloads. Adding a `String` then
reading as `int` (via `toInt()`) throws `InvalidCastException`.
Fix: pick the typed `add` overload that matches the value
(`String`, `int`, `double`, `boolean`, `DateTime`) and read it back
with the matching typed accessor (`toString()`, `toInt()`, `toDouble()`,
`toBool()`, `toDateTime()`).

### Built-in `DateTime` is `com.aspose.cells.DateTime`, not `java.util.Date`

Symptom: code using `new java.util.Date()` for `setCreatedTime(...)`
does not compile.
Cause: the Aspose type is `com.aspose.cells.DateTime`.
Fix: use `DateTime.getNow()` (or the static helpers on
`com.aspose.cells.DateTime`) and the Aspose overloads.

### Document protection vs signatures vs document properties — distinct concepts

Do not confuse document properties with the protection/encryption or
signature features:

| Feature                       | API                                              | Effect                                       |
| ----------------------------- | ------------------------------------------------ | -------------------------------------------- |
| Workbook digital signature    | `Workbook.setDigitalSignature(ds)`               | Proves authorship/integrity; needs a cert.   |
| Open-password **encryption**  | `WorkbookSettings.setPassword(pwd)`              | Encrypts the file; opening requires the password. |
| **Write protection** (read-only) | `Workbook.getSettings().getWriteProtection().setPassword(pwd) + setRecommendReadOnly(true)` | Excel opens the file read-only; password lifts the recommendation. File bytes are **not** encrypted. |
| **Workbook structure protection** | `Workbook.protect(ProtectionType.STRUCTURE, pwd)` | Inside an open file, blocks sheet add/rename/delete/reorder. |
| **Worksheet protection**      | `Worksheet.protect(ProtectionType, pwd, null)`    | Inside an open file, blocks cell edits on that sheet. |
| **Document properties**       | `Workbook.getBuiltInDocumentProperties()`,        | Metadata visible under File → Info.          |
|                               | `Workbook.getCustomDocumentProperties()`         |                                              |

These are independent — a workbook can carry any combination.
For the full write-up on structure protection, write protection,
and encryption (including the difference between
`WorkbookSettings.setPassword` and `WriteProtection.setPassword`)
see [encryption-decryption.md](encryption-decryption.md).

### Custom linked properties vs static values

`addLinkToContent(name, source)` registers a custom property whose
value is bound to a cell range. After the source cells change, call
`updateLinkedPropertyValue()` (or `updateLinkedRange()`) to refresh
the property. Plain `add` values are static and unaffected.

```java
CustomDocumentPropertyCollection props = wb.getCustomDocumentProperties();
props.addLinkToContent("SourceCell", "Sheet1!$A$1");

// Later, after editing A1:
props.updateLinkedPropertyValue();
```

## Related

- [digital-signature.md](digital-signature.md) — workbook and VBA
  signatures.
- [encryption-decryption.md](encryption-decryption.md) — open-password
  encryption and structure protection.
