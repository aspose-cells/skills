---
name: fonts-linux-docker
description: Setup fonts for Aspose.Cells for Java on Linux, Docker, Alpine, Ubuntu, headless containers. Required whenever the workbook is converted to PDF, HTML, or image. Skipping this produces tofu (□) or Fontconfig RuntimeException.
applies_to: Java
keywords: [fonts, linux, docker, headless, fontconfig, tofu, dejavu]
---

# Fonts on Linux / Docker

`Fontconfig head is null` and "tofu" squares in output PDF/image are
**always** a font problem, never an Aspose bug. They occur because
the XLSX references fonts (Calibri, Arial, …) that the OS image does
not ship with.

## Two-line fix for every Linux container

JVM flag:

```
-Djava.awt.headless=true
```

OS packages:

```bash
# Debian / Ubuntu
apt-get install -y fontconfig ttf-dejavu

# Alpine
apk add --no-cache fontconfig ttf-dejavu

# RHEL / CentOS
yum install -y fontconfig dejavu-sans-fonts
```

After install: `fc-cache -f -v`.

## Minimal Dockerfile

```dockerfile
FROM eclipse-temurin:17-jre

RUN apt-get update && \
    apt-get install -y fontconfig ttf-dejavu font-noto-color-emoji && \
    rm -rf /var/lib/apt/lists/* && \
    fc-cache -f -v

ENV JAVA_TOOL_OPTIONS="-Djava.awt.headless=true"

COPY app.jar /app.jar
ENTRYPOINT ["java","-jar","/app.jar"]
```

## Using custom fonts (e.g. corporate brand)

Place `.ttf` files in a directory inside the container and register
them at startup:

```java
import com.aspose.cells.*;

FontConfigs.setFontFolder("/opt/fonts", true);      // recursive
// or several folders:
FontConfigs.setFontFolders(new String[]{"/opt/fonts", "/usr/share/fonts"}, true);
// or one folder at a time (non-recursive):
FontConfigs.setFontFolder("/opt/fonts", false);
```

## Pitfalls

- **Skip the headless flag** — even with fonts installed, GUI calls
  inside the JVM silently fail and may corrupt the PDF.
- **Just-in-case fonts**: also install `fonts-liberation` so the
  default Excel substitute fonts (Arial / Times) exist.
- **CJK fonts**: for Chinese/Japanese/Korean spreadsheets install
  `fonts-noto-cjk` or specific `.ttf` files. Otherwise CJK text in
  PDFs will be squares.
- **Alpine is the worst offender** — its base JDK image has zero
  fonts; almost every Azure / AWS Lambda / Cloud Run user hits the
  tofu bug on Alpine.
- **WebP inside XLSX**: `twelvemonkeys` ImageIO plugin is required if
  the workbook embeds WebP — see
  [images-shapes/insert-image.md](../images-shapes/insert-image.md).

## Related

- [rendering/convert-to-pdf.md](../rendering/convert-to-pdf.md)
- [rendering/worksheet-image.md](../rendering/worksheet-image.md)
- [rendering/convert-html.md](../rendering/convert-html.md)
