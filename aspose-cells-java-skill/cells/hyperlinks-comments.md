---
name: hyperlinks-comments
description: Add hyperlinks (web, file, cell, range) and cell comments (text, threaded, picture) in Aspose.Cells for Java. Use when the user wants clickable links or comment annotations on cells.
applies_to: Java
keywords: [hyperlink, link, comment, threaded comment, annotation]
---

# Hyperlinks & comments

## Imports

```java
import com.aspose.cells.*;
```

## Quick example — hyperlink

```java
Worksheet ws = wb.getWorksheets().get(0);
ws.getHyperlinks().add("A1", 1, 1, "https://aspose.com");   // cellName, rows, cols, url
```

Signatures:

- `add(cellName, rows, cols, url)` — place at the named cell, jump to
  a (rows, cols) offset, link to `url`.
- `add(cellName, url)` — link to URL, display = cell text.

## Cell-internal hyperlink (jump to another cell)

```java
ws.getHyperlinks().add(1, 1, 9, 2, "Sheet1!C10");   // jump from B1 to Sheet1!C10
```

## Comment

```java
CommentCollection comments = ws.getComments();
int i = comments.add("B5");
Comment c = comments.get(i);
c.setNote("Reviewed on 2024-04-12 by Alice");
c.setAuthor("Alice");
c.setWidth(220);
c.setHeight(80);
c.setVisible(true);
```

## Threaded comment

```java
Worksheet ws = wb.getWorksheets().get(0);
CommentCollection comments = ws.getComments();
int i = comments.add("B5");
Comment c = comments.get(i);

ThreadedCommentAuthorCollection authors = wb.getWorksheets().getThreadedCommentAuthors();
int bobIdx = authors.add("Bob", "user-id-bob", "provider-id");
ThreadedCommentAuthor bob = authors.get(bobIdx);

ThreadedCommentCollection tc = c.getThreadedComments();
tc.add("Looks good", bob);
```

## Pitfalls

- **Hyperlink target uses 1-based row/column in the second overload.**
- **External file links** are sandboxed by Excel; warn the user
  before adding `file://` paths.
- **`Comment.setVisible(true)`** is required for Excel to show it on
  open; default is `false`.
- **Threaded comments require Office 365 / 2019+** — older Excel
  ignores them.

## Related

- [read-write-values.md](read-write-values.md)
- [insert-image.md](../images-shapes/insert-image.md)
