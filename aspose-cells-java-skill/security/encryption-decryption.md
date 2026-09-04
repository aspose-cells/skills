---
name: encryption-decryption
description: Three independent workbook-security mechanisms in Aspose.Cells for Java — workbook structure protection (Workbook.protect — no add/rename/delete sheets after open), write protection (WriteProtection — file opens read-only), and byte-level file encryption (WorkbookSettings.setPassword — password required to open). Use when the user wants to lock / unlock a workbook or to recommend a file open as read-only.
applies_to: Java
keywords: [encryption, decryption, password, protection, AES, JCE, write-protection, workbook-protection, worksheet-protection, STRUCTURE, read-only, recommended]
---

# Workbook security — three independent mechanisms

These three are often confused. They address different threats and can
be combined freely:

| Need                                            | API                                                                                  | What it does                                                                                                       |
| ----------------------------------------------- | ------------------------------------------------------------------------------------ | ------------------------------------------------------------------------------------------------------------------ |
| **Workbook structure protection** (no add / rename / delete / reorder sheets) | `Workbook.protect(ProtectionType.STRUCTURE, pwd)`                                     | File opens normally; structural edits are blocked until the password is supplied.                                  |
| **Worksheet protection** (no cell edits on a sheet) | `Worksheet.protect(ProtectionType, pwd, null)`                                       | File opens normally; edits on that sheet need the password.                                                        |
| **Write protection** (file opens read-only)     | `Workbook.getSettings().getWriteProtection().setPassword(pwd) + setRecommendReadOnly(true)` | File opens read-only in Excel; password lifts the read-only. The file bytes are **not** encrypted.                 |
| **File encryption** (password required to OPEN) | `Workbook.getSettings().setPassword(pwd)` (+ optional `setEncryptionOptions(...)` for `.xls`) | The saved file bytes are encrypted; without the password Excel cannot read the workbook at all.                     |

Each applies independently — a file can carry any combination of
the four (structure + worksheet + write-protect + encrypt).

## Imports

```java
import com.aspose.cells.*;
```

## Workbook structure protection — `Workbook.protect`

Locks the workbook's structure: `add` / `rename` / `delete` /
`reorder` / `move` worksheets. The file opens normally; the
password is required only to make structural changes.

```java
Workbook wb = new Workbook("report.xlsx");
wb.protect(ProtectionType.STRUCTURE, "structure-pwd");
wb.save("report.xlsx");

// Read state, lift protection
// wb.isWorkbookProtectedWithPassword();
// wb.unprotect("structure-pwd");
```

`ProtectionType.STRUCTURE` is the only meaningful type at workbook
level. `ProtectionType.ALL`, `CONTENTS`, `OBJECTS`, `SCENARIOS` are
worksheet-level types and apply only when calling
`Worksheet.protect(...)`.

## Worksheet protection — `Worksheet.protect`

Locks a specific sheet. With `ProtectionType.ALL` the user can't
edit cells, formulas, formatting, rows, or columns. The file opens
normally; cell edits on this sheet require the password.

```java
Worksheet ws = wb.getWorksheets().get(0);
ws.protect(ProtectionType.ALL, "edit-pwd", null);
// ws.unprotect("edit-pwd");
```

Granular types: `CONTENTS` blocks cell edits, `OBJECTS` blocks
drawing objects, `SCENARIOS` blocks scenarios, `STRUCTURE` (here)
locks row/column insert/delete/format, `WINDOWS` keeps the user
from changing window layout, `NONE` clears protection.

## Write protection (open read-only) — `WriteProtection`

Tells Excel to open the file in read-only mode. With a password
set, the user must enter that password before they can edit and
save. The file is **not** encrypted — the bytes are readable to
anyone with filesystem access; only the write/edit path is
gated.

```java
Workbook wb = new Workbook("report.xlsx");
wb.getSettings().getWriteProtection().setPassword("write-pwd");
wb.getSettings().getWriteProtection().setRecommendReadOnly(true);
wb.save("read-only.xlsx");
```

Without a password the file still opens read-only, but the user can
flip to read-write in one click. With a password, removing the
read-only requires `setRecommendReadOnly(false)` + `validatePassword(pwd)`
on the read side, or supplying the password through Excel's
"Stop read-only" dialog.

## File encryption (open password) — `WorkbookSettings.setPassword`

Encrypts the file's bytes. Without the password, Excel cannot read
the workbook — that's a stronger guarantee than write-protection,
which only affects the editing UI after open.

```java
Workbook wb = new Workbook("report.xlsx");
wb.getSettings().setPassword("open-pwd");   // opens require this password
wb.save("encrypted.xlsx");
```

For legacy `.xls` also pick the algorithm and key length via
`Workbook.setEncryptionOptions(...)`:

```java
Workbook wb = new Workbook("legacy.xls");
wb.getSettings().setPassword("open-pwd");
wb.setEncryptionOptions(EncryptionType.STRONG_CRYPTOGRAPHIC_PROVIDER, 128);
wb.save("encrypted.xls");
```

`.xlsx` and `.xlsm` always use AES regardless; `setEncryptionOptions(...)`
is ignored on those formats. `EncryptionType.XOR` and
`EncryptionType.COMPATIBLE` are the legacy Office 97/2000 defaults
(crackable); `EncryptionType.STRONG_CRYPTOGRAPHIC_PROVIDER` with
128 bits is the strongest option on `.xls`. The other `ENHANCED_*`
and `STRONG_*` variants are tied to legacy Windows CSPs; prefer
them only when matching an old `xlFileMode`.

## Open an encrypted file

```java
LoadOptions loadOpts = new LoadOptions();
loadOpts.setPassword("open-pwd");
Workbook wb = new Workbook("encrypted.xlsx", loadOpts);
```

`loadOpts.setPassword(null)` works for files without a password;
when there is one, Excel throws `InvalidPasswordException` on a
wrong password.

## Detect encryption without loading

```java
FileFormatInfo info = FileFormatUtil.detectFileFormat("locked.xlsx");
if (info.isEncrypted()) {
    System.out.println("Type: " + info.getFileFormatType());    // OOXML vs XLS97
    System.out.println("Encrypted: yes");
}
```

`FileFormatType` distinguishes modern OOXML encryption
(AES inside the OOXML package) from legacy binary `.xls`
encryption — the password path is the same, but the key
derivation differs.

## Pitfalls

- **The three mechanisms are independent.** A file can carry structure
  protection, write-protection, and encryption all at once. They
  protect against different threats; pick by intent.
- **`WorkbookSettings.setPassword` ≠ `WriteProtection.setPassword`.**
  The former encrypts the file (open password); the latter sets
  the recommendation to open read-only. They use different
  mechanisms and persist in different XML parts of the workbook.
- **BouncyCastle** is suggested for any project, and **required**
  for ODS encrypt / decrypt (different code path from
  XLSX / XLS).
- **AES-256 XLSX + `InvalidKeyException: Illegal key size`** →
  install **JCE Unlimited Strength Jurisdiction Policy** for the
  JRE (`${java.home}/lib/security/`), or upgrade to JDK 8u161+ /
  JDK 17+ where it's bundled.
- **`setEncryptionOptions(...)` is `.xls`-only.** On
  `.xlsx` / `.xlsm` Excel always uses AES regardless.
- **Streaming detection** — pass the file path *or* stream to
  `FileFormatUtil.detectFileFormat`; once the stream is consumed
  the bytes are gone — open the workbook with new bytes if needed.

## Related

- [digital-signature.md](digital-signature.md) — workbook and VBA
  signatures (an unrelated integrity guarantee, often combined
  with the above).
- [../getting-started/detect-format.md](../getting-started/detect-format.md)
  — `FileFormatUtil.detectFileFormat` for `FileFormatInfo`.
