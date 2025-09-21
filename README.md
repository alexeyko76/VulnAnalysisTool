# Java Vulnerability Analysis Tool

**Version 2.0.0**

A comprehensive defensive security tool that processes Excel files to analyze file existence, modification dates, and file versions (JAR and EXE) across different platforms. Features include automatic path corruption detection and fixing, duplicate detection system, CVE information sheet creation with NIST NVD integration, Oracle WebLogic vulnerability detection, and support for both local file analysis and remote Windows file access via UNC paths.

**Implementation**: Single Java file, fully compatible with **Java 1.8**.  

## Configuration (config.properties)
- `excel.path` - Excel file path
- `sheet.name` - Sheet name to process
- Column names: `PlatformName`, `FilePath`, `HostName`, `CVE`
- `platform.windows` - Windows platform identifiers (comma-separated, supports spaces, e.g., "Windows Server 2019, Windows Server 2022")
- `remote.unc.enabled` - Enable/disable remote Windows UNC access (true/false)
- `remote.unc.timeout` - UNC access timeout in seconds (default: 7, prevents infinite hangs)
- `log.filename` - Log file name (optional) - if specified, console output will also be saved to this file with hostname prefix (e.g., `excel-tool.log` becomes `HOSTNAME-excel-tool.log`)
- `invalid.path.detection` - Enable/disable invalid path pattern validation (true/false, default: true)
- `duplicate.search.enabled` - Enable/disable duplicate detection system (true/false, default: false)
- `cve.sheet.creation.enabled` - Enable/disable CVE information sheet creation with NIST NVD data (true/false, default: false)  

## Processing Steps
1. Open the Excel file.  
   - Locate the specified sheet by name from the configuration.
     - If the **sheet does not exist**, exit with an error message listing available sheets and **do not process** any data.
   - Verify that columns `PlatformName`, `FilePath`, and `HostName` exist in the specified sheet.  
     - If **any of these columns are missing**, exit with a clear error message and **do not save** the Excel file.  
   - Ensure the following additional columns exist (create them if missing):  
     - `FileExists`  
     - `FileModificationDate` (readable format: `yyyy-MM-dd HH:mm:ss`)
     - `FileVersion` (optional, filled for `.jar` files on all platforms, `.exe` files on Windows platforms)
     - `ScanError` (captures local file scanning issues or errors)
     - `RemoteScanError` (captures remote UNC access issues, cleared for successful remote scans)
     - `ScanDate` (timestamp when scanning session started, format: `yyyy-MM-dd HH:mm:ss`)
     - `FixedFilename` (stores corrected file path when path fixing is applied)
     - `FixedFileExists` (shows existence status of fixed path: Y/N/Error)  

2. Read the Excel file row by row.  

3. For each row:  
   - **Hostname Filtering**: Skip rows that don't match the local hostname (and aren't remote Windows hosts if UNC is enabled)
   - **Scan Timestamp**: Record the session start timestamp in `ScanDate` column for all processed hosts (same timestamp for entire scanning session)
   - Process files for the local host name, or optionally for remote Windows hosts (if UNC access is enabled).
   - **UNC Access**: For remote Windows hosts, converts paths like `C:\path\file.jar` to `\\hostname\C$\path\file.jar`.
   - **Timeout Protection**: UNC access operations have configurable timeout to prevent infinite hangs on unreachable hosts.
   - **Smart Exclusion**: Hosts that fail UNC access (network errors, timeouts, or permission issues) are added to exclusion list for the current run.
   - **UNC Access Preservation**: When UNC access fails, only the `RemoteScanError` column is updated - existing values in `FileExists`, `FileModificationDate`, and `JarVersion` are preserved.
   - **File Type Validation**: Only processes regular files, excludes directories and special files.
   - **Path Processing Workflow**:
     1. **Try Original Path**: First attempt to find the file using the original path from `FilePath` column
     2. **Path Corruption Detection**: If original file doesn't exist, check if path contains corrupted patterns
     3. **Automatic Path Fixing**: If corruption detected, attempt to fix by removing trailing garbage data after first space in filename
     4. **Fixed Path Validation**: Test the fixed path and record results in `FixedFilename` and `FixedFileExists` columns
     5. **Metadata Extraction**: Extract file information from whichever file exists (original or fixed)
   - **Path Fixing Examples**:
     - `C:\app\tool.exe garbage_data` → `C:\app\tool.exe`
     - `C:\path\file.jar result.filename=xyz` → `C:\path\file.jar`
     - `setup.exe version 1.2.3 build` → `setup.exe`
   - Resolve file paths in a **platform-independent way** (handle both Windows `\` and Linux `/` path formats).
   - **File Existence Results**:
     - `FileExists` column: Shows original file status (`Y`=exists, `N`=not found, `X`=invalid/corrupted path)
     - `FixedFileExists` column: Shows fixed file status when path fixing was applied (`Y`/`N`/`Error`)
   - **File Analysis**: If a file exists (original or fixed):
       - Write the file's last modified timestamp into the `FileModificationDate` column (format: `yyyy-MM-dd HH:mm:ss`).
       - **File Version Extraction**:
         - If the file has a `.jar` extension (all platforms):
           - Open it as a ZIP archive.
           - If it contains `META-INF/MANIFEST.MF`:
             - **Robust Parsing**: Uses both Manifest API and manual text parsing as fallback.
             - Extract the value from the line starting with `Implementation-Version:`.
             - Write this value into the `FileVersion` column.
           - If any JAR processing errors occur, record them in the `ScanError` column.
         - If the file has a `.exe` extension (Windows platforms only):
           - **PE Header Parsing**: Uses direct binary parsing of Windows PE headers (no external dependencies)
           - Extracts version information from PE resource section (VS_FIXEDFILEINFO structure)
           - **UNC Compatible**: Works seamlessly with remote Windows files via UNC paths
           - Formats version as `major.minor.build.revision` (e.g., `1.2.3.4`)
           - Write this value into the `FileVersion` column.
           - If any EXE processing errors occur, record them in the `ScanError` column.
     - If the file does **not** exist:  
       - Write `"N"` into the `FileExists` column.  
       - Clear the `ScanError` column (successful scan - file simply doesn't exist).
       - Leave `FileModificationDate` and `FileVersion` blank.
   - **Enhanced File Validation**: 
     - Uses `Files.exists()` and `Files.notExists()` to differentiate access issues from non-existence
     - Uses `Files.isRegularFile()` to ensure paths point to actual files (not directories)
     - **Invalid Path Detection**: Configurable validation (`invalid.path.detection=true/false`) that identifies:
       - Empty or blank file paths
       - Paths marked as "N/A" or similar placeholder values  
       - JAR files with spaces in the filename (not directory path)
       - Paths with trailing spaces after "Program Files" (exception: "C:\Program Files (x86)" is valid)
       - Directories mistakenly listed as files
       - Malformed path patterns containing "result.filename=" strings
     - **Error Classifications**:
       - Invalid/corrupted paths: `FileExists = "X"`, `ScanError = "[specific reason]"` (e.g., "JAR filename contains spaces: file name.jar")
       - Files that genuinely don't exist: `FileExists = "N"`, `ScanError = ""` (cleared for successful scan)
       - Local access permission issues: `FileExists = "N"`, `ScanError = "Access denied - cannot determine file existence"`
       - JAR processing errors: `ScanError = "JAR processing error: [details]"`
       - UNC access permission issues: `RemoteScanError = "UNC access denied - cannot determine file existence"`
       - UNC timeout issues: `RemoteScanError = "UNC access timeout - host may be unreachable or slow"`
       - UNC connection failures: `RemoteScanError = "Cannot access remote host via UNC: [details]"`
       - Invalid UNC paths: `RemoteScanError = "Invalid path format for UNC conversion"`
   - **Error Column Usage**:
     - `ScanError`: Records local file scanning issues (file access, JAR processing, etc.) - cleared for successful local scans
     - `RemoteScanError`: Records UNC access issues for remote hosts - cleared for successful remote scans
   - If any other local scanning errors occur (e.g., permission issues, corrupted files), record them in the `ScanError` column.  

4. Save the updated Excel file after all rows are processed.

5. **Progress Display & Console Reporting**: 
   - **Timestamped Logging**: Displays timestamped row-by-row logging with detailed messages
   - **Final Summary**: Comprehensive execution summary including:
     - Total rows processed
     - Rows skipped due to hostname mismatch
     - Rows skipped due to inaccessible remote hosts
     - Complete list of inaccessible hosts identified during the run

## CVE Information Sheet Creation

When `cve.sheet.creation.enabled=true`, the tool creates a comprehensive "CVEs" sheet with detailed vulnerability information:

### Features
- **NIST NVD Integration**: Fetches detailed CVE data from the National Vulnerability Database API v2.0
- **Weblogic Detection**: Automatically identifies Oracle WebLogic Server vulnerabilities by analyzing CPE configurations
- **Oracle Advisory Extraction**: Extracts Oracle security advisory URLs from CVE references (supports both HTTP and HTTPS)
- **Rate Limiting**: Built-in 2-second delays between API requests to respect NIST API limits
- **Error Handling**: Graceful handling of API errors (404 for missing CVEs, 429 for rate limiting)

### CVE Sheet Columns
1. **CVE ID**: The CVE identifier
2. **Description**: Detailed vulnerability description from NIST
3. **References**: All reference URLs associated with the CVE
4. **Affected Software**: CPE (Common Platform Enumeration) configurations
5. **Weblogic**: Y/N flag indicating if this is an Oracle WebLogic Server vulnerability
6. **Weblogic Sec Advisories**: Oracle security advisory URLs extracted from references

### Usage
- Set `cve.sheet.creation.enabled=true` in `config.properties`
- When enabled, only CVE sheet creation occurs (normal file processing is bypassed)
- Successfully tested with 15+ real WebLogic CVEs from 2020-2024
- Handles real-world data including escaped JSON URLs and Oracle advisory extraction

## Recommended Libraries
- **Apache POI** – for reading and writing Excel files.  
- **java.nio.file.Paths** – for platform-independent path handling.  
- **java.nio.file.Files / java.io.File** – for checking file existence and last modified date.  
- **java.util.zip.ZipFile** – for reading `.jar` files as ZIP archives.  
- **java.util.Properties or BufferedReader** – for parsing `MANIFEST.MF`.  

## Compatibility
- Must run on **Java 1.8** (no newer language features or APIs beyond Java 8).  
- Must be implemented in **a single Java file**.  
- **Code Quality**: Implements standardized error handling patterns and helper methods to reduce duplication while maintaining single-file architecture.
- **Flexible Configuration**: Supports multiple Windows platform values separated by commas, with proper handling of spaces in platform names.  

## Build & Dependencies

### Build Options
1. **Maven Build** (Recommended):
   - Dependencies managed automatically via `pom.xml`
   - Windows: `maven-build.bat`
   - Creates: `java-excel-tool-uber.jar`

2. **Manual Build**:
   - Dependencies stored in `deps/` folder
   - Windows: `build.bat` (with JAVA_HOME support, progress messages, and automatic cleanup)
   - Linux/macOS: `./build.sh`
   
**Enhanced build.bat Features**: The `build.bat` script provides:
- **JAVA_HOME Support**: Uses `JAVA_HOME` environment variable if set, falls back to default Java 8 installation path: `C:\Program Files\Java\jdk1.8.0_401`
- **Build Validation**: Validates that Java tools (javac, jar) exist before building
- **Progress Messages**: Clear status messages for each build stage (compiling, extracting dependencies, creating JAR, cleanup)
- **Automatic Cleanup**: Removes temporary build artifacts (`target` folder) after successful build
- **Error Handling**: Provides clear error messages if Java installation is not found or build fails

### Running the Tool
```bash
java -jar java-excel-tool-uber.jar config.properties
```

Windows batch scripts:
- `maven-run.bat` (Maven-based)  
- `run.bat` (Manual build)

**Output**: Single executable uber JAR with all dependencies included.  