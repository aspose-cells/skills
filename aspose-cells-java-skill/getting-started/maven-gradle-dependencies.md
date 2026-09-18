---
name: maven-gradle-dependencies
description: Aspose.Cells for Java Maven and Gradle coordinates, including the official Aspose repository. Use when the user needs to add Aspose.Cells to a build file or asks about the dependency declaration.
applies_to: Java
keywords: [maven, gradle, dependency, repository, pom, build, jitpack]
---

# Dependencies — Maven & Gradle

Aspose.Cells for Java is published to the official Aspose repository
(`https://releases.aspose.com/java/repo`), **not** Maven Central.

## Maven (`pom.xml`)

```xml
<repositories>
    <repository>
        <id>aspose</id>
        <url>https://releases.aspose.com/java/repo</url>
    </repository>
</repositories>

<dependencies>
    <dependency>
        <groupId>com.aspose</groupId>
        <artifactId>aspose-cells</artifactId>
        <version>26.8</version>
    </dependency>
</dependencies>
```

For Java 9+ module projects add `--add-modules=java.scripting` to the
Surefire/compiler args if you hit `Nashorn` errors in formulas.

## Gradle (`build.gradle`)

```gradle
repositories {
    maven { url 'https://releases.aspose.com/java/repo' }
}

dependencies {
    implementation 'com.aspose:aspose-cells:26.8'
}
```

## Suggested: BouncyCastle

Worth adding to any project. **Required** for VBA project signatures
and for ODS encrypt / decrypt.

```xml
<dependency>
    <groupId>org.bouncycastle</groupId>
    <artifactId>bcprov-jdk18on</artifactId>
    <version>1.78</version>
</dependency>
<dependency>
    <groupId>org.bouncycastle</groupId>
    <artifactId>bcpkix-jdk18on</artifactId>
    <version>1.78</version>
</dependency>
```

Aspose.Cells picks the jars up and registers the provider itself —
no setup code needed.

## Optional: WebP image support inside XLSX

```xml
<dependency>
    <groupId>com.twelvemonkeys.imageio</groupId>
    <artifactId>imageio-webp</artifactId>
    <version>3.10</version>
</dependency>
```

## Pitfalls

- **Do not substitute `org.apache.poi`** — both define a `Workbook`
  class; classpath collisions crash the app.
- **Maven Central alone** — gives `Could not resolve com.aspose:aspose-cells`.
  Always add the Aspose repo too.
- **JitPack does not host Aspose artifacts** as primary source; do not
  point there.
- **JDK 1.8 is the minimum.** Newer `aspose-cells` builds may target
  11+; verify with `mvn dependency:tree | findstr jdk`.
