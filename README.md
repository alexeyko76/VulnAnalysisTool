# Java Vulnerability Analysis Tool

A defensive security tool that processes Excel files to analyze file existence, modification dates, and JAR versions across different platforms. Supports both local file analysis and remote Windows file access via UNC paths.

**Implementation**: Single Java file, fully compatible with **Java 1.8**.  

## Configuration (config.properties)
- `excel.path` - Excel file path  
- `sheet.name` - Sheet name to process
- Column names: `PlatformName`, `FilePath`, `HostName`
- `platform.windows` - Windows platform identifier (e.g., "Windows_2019")
- `remote.unc.enabled` - Enable/disable remote Windows UNC access (true/false)  

## Processing Steps
1. Open the Excel file.  
   - Locate the specified sheet by name from the configuration.
     - If the **sheet does not exist**, exit with an error message listing available sheets and **do not process** any data.
   - Verify that columns `PlatformName`, `FilePath`, and `HostName` exist in the specified sheet.  
     - If **any of these columns are missing**, exit with a clear error message and **do not save** the Excel file.  
   - Ensure the following additional columns exist (create them if missing):  
     - `FileExists`  
     - `FileModificationDate` (readable format: `yyyy-MM-dd HH:mm:ss`)
     - `JarVersion` (optional, filled only for `.jar` files)
     - `ScanError` (captures any scanning issues or errors)  

2. Read the Excel file row by row.  

3. For each row:  
   - Process files for the local host name, or optionally for remote Windows hosts (if UNC access is enabled).
   - **UNC Access**: For remote Windows hosts, converts paths like `C:\path\file.jar` to `\\hostname\C$\path\file.jar`.
   - **Smart Exclusion**: Hosts that fail UNC access are added to exclusion list for the current run.
   - Resolve the `FilePath` value in a **platform-independent way** (handle both Windows `\` and Linux `/` path formats).  
   - Check if the file in the `FilePath` column exists.  
     - If the file exists:  
       - Write `"Y"` into the `FileExists` column.  
       - Write the file's last modified timestamp into the `FileModificationDate` column (format: `yyyy-MM-dd HH:mm:ss`).
       - If the file has a `.jar` extension:  
         - Open it as a ZIP archive.  
         - If it contains `META-INF/MANIFEST.MF`:  
           - **Robust Parsing**: Uses both Manifest API and manual text parsing as fallback.
           - Extract the value from the line starting with `Implementation-Version:`.  
           - Write this value into the `JarVersion` column.
         - If any JAR processing errors occur, record them in the `ScanError` column.
     - If the file does **not** exist:  
       - Write `"N"` into the `FileExists` column.  
       - Write `"File does not exist"` into the `ScanError` column.
       - Leave `FileModificationDate` and `JarVersion` blank.
   - If any other scanning errors occur (e.g., permission issues, corrupted files, UNC access failures), record them in the `ScanError` column.  

4. Save the updated Excel file after all rows are processed.  

## Recommended Libraries
- **Apache POI** – for reading and writing Excel files.  
- **java.nio.file.Paths** – for platform-independent path handling.  
- **java.nio.file.Files / java.io.File** – for checking file existence and last modified date.  
- **java.util.zip.ZipFile** – for reading `.jar` files as ZIP archives.  
- **java.util.Properties or BufferedReader** – for parsing `MANIFEST.MF`.  

## Compatibility
- Must run on **Java 1.8** (no newer language features or APIs beyond Java 8).  
- Must be implemented in **a single Java file**.  

## Build & Dependencies

### Build Options
1. **Maven Build** (Recommended):
   - Dependencies managed automatically via `pom.xml`
   - Windows: `maven-build.bat`
   - Creates: `java-excel-tool-uber.jar`

2. **Manual Build**:
   - Dependencies stored in `deps/` folder
   - Windows: `build.bat`
   - Linux/macOS: `./build.sh`

### Running the Tool
```bash
java -jar java-excel-tool-uber.jar config.properties
```

Windows batch scripts:
- `maven-run.bat` (Maven-based)  
- `run.bat` (Manual build)

**Output**: Single executable uber JAR with all dependencies included.  