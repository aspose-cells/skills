# Troubleshooting Guide for Aspose.Cells GridJs (Java)

| # | Symptom (short description) |
|---|------------------------------|
| 1 | **`NoClassDefFoundError: com/google/gson/JsonElement`** on application start |
| 2 | **`java.io.FileNotFoundException`** thrown by `GridJsService` because the file‑cache directory does not exist |
| 3 | **`java.io.FileNotFoundException`** when calling `loadSpreadsheet` – the workbook cannot be found |
| 4 | **`@Bean definition illegally overridden by existing bean definition`** when a custom `GridJsService` bean is declared |
| 5 | **Compilation error**: “unhandled exception type `Exception`” when invoking any `GridJsService` method that declares `throws java.lang.Exception` |
| 6 | **Compilation error**: “unhandled exception type `IOException`” when overriding `lazyLoadingStreamJson` in a controller |
| 7 | **`IllegalStateException` / compile‑time “method does not override or implement a method from a supertype”** when a controller method is redeclared with the wrong return type |

Below each symptom you will find:

* **Observable** – what you see in the console, IDE, or HTTP response.  
* **Likely cause** – the configuration or code mistake that triggers it.  
* **Check** – a quick way to verify the cause.  
* **Fix** – the exact change required (including code snippets) to resolve the issue.

---

## 1. `NoClassDefFoundError: com/google/gson/JsonElement`

**Observable**  
```
Exception in thread "main" java.lang.NoClassDefFoundError: com/google/gson/JsonElement
    at com.aspose.gridjs.GridJsService.<init>(GridJsService.java:...)
    ...
```

**Likely cause**  
The Gson library (`com.google.code.gson:gson`) is missing from the Maven classpath. The GridJs extension uses Gson internally; without it the Spring container cannot create `GridJsService`.

**Check**  
Open `pom.xml` and verify that the following dependency exists:

```xml
<dependency>
    <groupId>com.google.code.gson</groupId>
    <artifactId>gson</artifactId>
    <version>2.10.1</version>
</dependency>
```

Run `mvn dependency:tree | grep gson` – you should see the Gson artifact listed.

**Fix**  
Add the dependency (or correct the version) in `pom.xml`, then rebuild:

```bash
mvn clean install
```

---

## 2. `java.io.FileNotFoundException` – missing file‑cache directory

**Observable**  
Application fails to start with a stack trace similar to:

```
java.io.FileNotFoundException: ./cache (No such file or directory)
    at com.aspose.gridjs.GridJsService.<init>(GridJsService.java:...)
```

**Likely cause**  
`GridJsOptions.setFileCacheDirectory("./cache")` points to a directory that does not exist at startup. `GridJsService` validates the path and throws `FileNotFoundException` if it cannot find it.

**Check**  
From the project root, run:

```bash
ls -ld ./cache
```

If the directory is missing, the command will report “No such file or directory”.

**Fix**  
Create the directory **before** the Spring application starts:

```bash
mkdir -p ./cache
```

Ensure the process has write permission:

```bash
chmod u+w ./cache
```

No code change is required; the existing `GridJsConfig` bean can stay unchanged.

---

## 3. `java.io.FileNotFoundException` when loading a workbook

**Observable**  
Calling the `/GridJs/LoadSpreadsheet` endpoint returns HTTP 500 with a message:

```
java.io.FileNotFoundException: ./src/main/resources/files/Unknown.xlsx (No such file or directory)
```

**Likely cause**  
`GridJsController.getFullFilePath` builds a relative path that does not resolve to the packaged workbook. The path works only when the working directory is the project root (e.g., `mvn spring-boot:run`). If the application is launched from another directory (e.g., a packaged JAR), the file cannot be found.

**Check**  
Print the resolved absolute path inside the controller:

```java
System.out.println("Resolved path: " + new java.io.File(fullPath).getAbsolutePath());
```

Run the endpoint and verify that the printed path points to an existing file.

**Fix**  
Load the workbook from the classpath instead of a file system path:

```java
private String getFullFilePath(String filename) throws java.io.IOException {
    java.net.URL url = getClass().getResource("/files/" + filename);
    if (url == null) {
        throw new java.io.FileNotFoundException("Resource not found: " + filename);
    }
    return url.getPath();
}
```

Now the controller works regardless of the current working directory.

---

## 4. `@Bean definition illegally overridden by existing bean definition`

**Observable**  
Application fails to start with:

```
org.springframework.beans.factory.support.BeanDefinitionOverrideException:
   Invalid bean definition with name 'gridJsService' defined in class path resource
   [com/example/.../DemoConfig.class] and overriding bean definition defined in
   class path resource [com/aspose/gridjs/GridJsService.class]
```

**Likely cause**  
A custom `@Bean` method creates a `GridJsService` instance. The GridJs JAR already registers `GridJsService` as a Spring component, so the container detects a duplicate bean name.

**Check**  
Search your source for a method annotated with `@Bean` that returns `GridJsService`:

```bash
grep -R "GridJsService" src/main/java | grep "@Bean"
```

**Fix**  
Remove the custom bean definition. The component scan already registers the service:

```java
// DemoConfig.java – delete the following method if present
/*
@Bean
public GridJsService gridJsService() throws Exception {
    return new GridJsService(); // <-- remove this
}
*/
```

Keep only the `GridJsOptions` bean (which you own) and the `@ComponentScan` entry for `com.aspose.gridjs`.

---

## 5. Unhandled `java.lang.Exception` when calling `GridJsService` methods

**Observable**  
Compilation fails with:

```
error: unhandled exception type Exception
    StringBuilder json = gridJsService.detailFileJsonWithUid(fullPath, uid);
                                 ^
```

**Likely cause**  
All public methods of `GridJsService` (e.g., `detailFileJsonWithUid`, `updateCell`, `clearJsonCacheForUid`) declare `throws java.lang.Exception`. The Java compiler requires these checked exceptions to be either caught or declared.

**Check**  
Open the method call in your controller or service and look for a missing `try/catch` or `throws` clause.

**Fix**  
Wrap the call in a try‑catch block and translate the exception to an appropriate HTTP response, or propagate it by adding `throws Exception` to the enclosing method signature.

**Example – catching the exception:**

```java
@GetMapping("/LoadSpreadsheet")
public ResponseEntity<String> loadSpreadsheet(
        @RequestParam String filename,
        @RequestParam String uid) {
    try {
        String fullPath = getFullFilePath(filename);
        StringBuilder json = _gridJsService.detailFileJsonWithUid(fullPath, uid);
        return ResponseEntity.ok()
                .contentType(MediaType.parseMediaType("text/plain;charset=UTF-8"))
                .body(json.toString());
    } catch (Exception e) {
        // Log and return a 500 with a friendly message
        return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR)
                .body("Failed to load spreadsheet: " + e.getMessage());
    }
}
```

If you prefer to propagate, change the method signature:

```java
public ResponseEntity<String> loadSpreadsheet(...) throws Exception { ... }
```

---

## 6. Unhandled `java.io.IOException` in `lazyLoadingStreamJson` overrides

**Observable**  
When compiling a custom controller that overrides `lazyLoadingStreamJson`, the compiler reports:

```
error: unhandled exception type IOException
    @Override
    public void lazyLoadingStreamJson(String uid, String sheetName, HttpServletResponse response) {
        // implementation
    }
```

**Likely cause**  
Both `GridJsControllerBase` (Servlet) and `GridJsJakartaControllerBase` (Jakarta) declare `lazyLoadingStreamJson` with `throws java.io.IOException`. An overriding method must either declare the same exception or a subclass thereof.

**Check**  
Inspect the method signature in your subclass; ensure it matches the base method’s `throws` clause.

**Fix**  
Add `throws IOException` to the overriding method, or catch the exception inside the method.

**Example – declaring the exception:**

```java
@Override
public void lazyLoadingStreamJson(String uid, String sheetName,
        HttpServletResponse response) throws IOException {
    // Your streaming logic here
}
```

**Or – handling it internally:**

```java
@Override
public void lazyLoadingStreamJson(String uid, String sheetName,
        HttpServletResponse response) {
    try {
        // streaming logic
    } catch (IOException ex) {
        // translate to HTTP 500
        response.sendError(HttpServletResponse.SC_INTERNAL_SERVER_ERROR,
                "Streaming failed: " + ex.getMessage());
    }
}
```

---

## 7. Mismatched‑override failure – wrong return type in a controller method

**Observable**  
During compilation you see:

```
error: method does not override or implement a method from a supertype
    @Override
    public String loadSpreadsheet(String filename, String uid) { ... }
```

Or at runtime a Spring `IllegalStateException` indicating that a bean method signature does not match the expected mapping.

**Likely cause**  
A subclass (your controller) redeclares a method that exists in the base class (`GridJsControllerBase` or `GridJsJakartaControllerBase`) but changes the return type (e.g., from `ResponseEntity<String>` to plain `String`). Java does not allow covariant return types for overridden methods unless the new type is a subclass of the original, which `String` is not.

**Check**  
Compare the method signature in the base class (look at the source or Javadoc) with the one in your controller. The base method `loadSpreadsheet` is **not** defined there, but any method you intend to expose as a Spring endpoint must not clash with a method of the same name in the base class.

**Fix**  
Either:

* **Rename** the method in your controller to avoid overriding, or
* **Match** the exact signature, including return type, of the method you intend to override.

**Correct example – matching the base signature (if you truly need to override):**

```java
@Override
public ResponseEntity<String> loadSpreadsheet(
        @RequestParam String filename,
        @RequestParam String uid) throws Exception {
    // implementation identical to the original example
}
```

**If you only wanted a new endpoint, rename it:**

```java
@GetMapping("/LoadSpreadsheet")
public ResponseEntity<String> loadSpreadsheetEndpoint(
        @RequestParam String filename,
        @RequestParam String uid) throws Exception {
    // implementation
}
```

Avoid using the same method name as any protected or public method in `GridJsControllerBase` unless you deliberately intend to replace its behavior.

---

### Quick Reference Checklist

| ✅ | Action |
|---|--------|
| 1 | Verify `gson` dependency is present in `pom.xml`. |
| 2 | Ensure the directory set via `options.setFileCacheDirectory("./cache")` exists and is writable. |
| 3 | Load workbooks from the classpath (`/files/…`) rather than a hard‑coded file system path. |
| 4 | Do **not** declare a `@Bean` for `GridJsService`; rely on the component scan. |
| 5 | Wrap every `GridJsService` call that throws `Exception` in `try/catch` **or** add `throws Exception` to the method signature. |
| 6 | When overriding `lazyLoadingStreamJson`, either declare `throws IOException` or handle it inside the method. |
| 7 | Keep method signatures (name, parameters, return type) identical to the base class when using `@Override`; otherwise rename the method. |

Following this guide should eliminate the most common startup and compile‑time problems when integrating Aspose.Cells GridJs into a Spring Boot application. If a symptom is not listed here, it is not documented in the current reference material.
