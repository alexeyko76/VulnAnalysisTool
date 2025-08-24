# Implementation Details

This repository provides a single-file Java 8 utility (`app.ExcelTool`) that implements the behavior defined in **README.md**.

## Design Notes
- **Single Java file**: All logic is in `src/main/java/app/ExcelTool.java`.
- **Java 1.8** compliance.
- **Excel I/O**: Apache POI via `WorkbookFactory` (supports `.xlsx` and `.xls`).
- **Required columns**: `PlatformName`, `FilePath`, `HostName` must already exist; otherwise the tool exits **without saving**.
- **Auto columns**: Ensures `FileExists`, `FileModificationDate`, `JarVersion` headers exist.
- **Hostname filter**: Only rows matching the running host are processed.
- **Cross-platform paths**: Normalizes `\` to `/` and uses `Paths.get(...).normalize()`.
- **Jar manifest parsing**: Reads `META-INF/MANIFEST.MF` and extracts `Implementation-Version` for `.jar` paths.
- **Timestamps**: ISO-8601 with offset, local timezone.

## Configuration

Create `config.properties` (UTF‑8):

```
excel.path=./data/sample.xlsx
column.PlatformName=PlatformName
column.FilePath=FilePath
column.HostName=HostName
```

Pass a custom path as the first CLI arg or rely on default `./config.properties`.

## Build Options

### 1) Maven (fat/uber JAR)
```
mvn -q -DskipTests package
```
Outputs:
- `target/java-excel-tool.jar` (thin)
- `target/java-excel-tool-jar-with-dependencies.jar` (single runnable JAR)

Run:
```
java -jar target/java-excel-tool-jar-with-dependencies.jar config.properties
```

### 2) No Maven — Thin JAR + deps folder
Place Apache POI jars into `deps/` (see `deps/README.txt`), then:
- Linux/macOS: `./build_no_maven.sh`
- Windows: `build_no_maven.bat`

Run:
```
java -cp "target/java-excel-tool.jar:deps/*" app.ExcelTool config.properties
```
(Windows uses `;` instead of `:`)

### 3) No Maven — Single runnable JAR (uber)
This embeds all deps into one JAR.
- Linux/macOS: `./build_uber.sh`
- Windows: `build_uber.bat`

Run:
```
java -jar target/java-excel-tool-uber.jar config.properties
```

**Notes:**
- Scripts remove `META-INF/*.SF|*.DSA|*.RSA` from unpacked libraries to avoid signature errors.
- If distributing, keep Apache POI LICENSE/NOTICE in the repo.

## Exit Codes
- `0` success
- `2` required columns missing
- `3` invalid Excel format
- `4` configuration error
- `5` unexpected error

## Library Compatibility Notes

- **Java baseline:** Apache POI **4.0+** requires **Java 8 or newer** at runtime.  
- **POI 5.4.1:** Still runs on **Java 8**. If you *build POI itself with Java modules*, JDK **11+** is required, but that does not affect using POI as a library in Java 8 apps.  
- **Recommendation:** Prefer **POI 5.4.1** for security and bug fixes. 