# Implementation Details

This repository provides a single-file Java 8 utility (`app.ExcelTool`) that implements the behavior defined in **README.md**.

## Design Notes
- **Single Java file**: All logic is in `src/main/java/app/ExcelTool.java`.
- **Java 1.8** compliance.
- **Excel I/O**: Apache POI via `WorkbookFactory` (supports `.xlsx` and `.xls`).
- **Required columns**: `PlatformName`, `FilePath`, `HostName` must already exist; otherwise the tool exits **without saving**.
- **Auto columns**: Ensures `FileExists`, `FileModificationDate`, `JarVersion`, `ScanError` headers exist.
- **Hostname filter**: Only rows matching the running host are processed.
- **Cross-platform paths**: Normalizes `\` to `/` and uses `Paths.get(...).normalize()`.
- **JAR manifest parsing**: Reads `META-INF/MANIFEST.MF` and extracts `Implementation-Version` for `.jar` paths.
- **Timestamps**: Human-readable format `yyyy-MM-dd HH:mm:ss`, local timezone.
- **Error tracking**: `ScanError` column captures file access issues, JAR processing errors, and other scanning problems.

## Configuration

Create `config.properties` (UTF‑8):

```
excel.path=./sample-data/sample.xlsx
sheet.name=Sheet1
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

### 2) No Maven — Single runnable JAR (uber) - **CURRENT METHOD**
This embeds all dependencies into one JAR for easy deployment.
- **Windows**: `build.bat` (fixed and working)
- **Linux/macOS**: `build.sh`

**Dependencies Required** (see `deps/README.md` for complete list):
- Apache POI 5.4.1 (poi, poi-ooxml, poi-ooxml-lite, xmlbeans)
- Apache Commons (collections4, compress, io, lang3)
- Apache Log4j 2.17.2 (api, core)

**Build Process:**
1. Compiles Java source from `src/main/java/app/`
2. Creates thin JAR with compiled classes
3. Extracts all dependency JARs into staging area
4. Removes signature files to avoid conflicts
5. Creates final uber JAR with all dependencies

Run:
```
java -jar java-excel-tool-uber.jar config.properties
```

**Notes:**
- build.bat now properly extracts all dependencies (path issues resolved)
- Scripts remove `META-INF/*.SF|*.DSA|*.RSA` from unpacked libraries to avoid signature errors
- All required dependencies are documented in `deps/README.md`

## Exit Codes
- `0` success
- `2` required columns missing
- `3` invalid Excel format
- `4` configuration error
- `5` unexpected error
- `6` specified sheet does not exist

## Library Compatibility Notes

- **Java baseline:** Apache POI **4.0+** requires **Java 8 or newer** at runtime.  
- **POI 5.4.1:** Still runs on **Java 8**. If you *build POI itself with Java modules*, JDK **11+** is required, but that does not affect using POI as a library in Java 8 apps.  
- **Recommendation:** Prefer **POI 5.4.1** for security and bug fixes. 