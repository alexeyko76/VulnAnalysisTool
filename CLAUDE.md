# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Version Information

**Current Version: 3.0.0**

## Recent Enhancements

**Key improvements implemented in version 3.0.0:**
- **Streamlined 4-Pass Processing**: Redesigned file processing logic with distinct phases: Path fixing, File operations, Duplicate detection, Invalid path detection
- **Enhanced Path Replacement System**: Multi-line configuration support with platform-aware case sensitivity for Windows vs Linux
- **Improved Column Structure**: Replaced FixedFilename with FixedFilePath, added IsFixed and IsInvalidPath columns for better audit trails
- **Platform-Specific Processing**: Path replacements use row-level platform detection instead of OS detection for accurate processing
- **CVE Information Sheet Creation**: Comprehensive NIST NVD API integration with detailed CVE data extraction, Weblogic detection, and Oracle security advisory parsing
- **Duplicate Detection System**: Configurable duplicate search functionality with UniqueID generation based on hostname, CVE, and file path
- **Enhanced File Version Extraction**: Comprehensive FileVersion column supporting both JAR and EXE files with Bundle-Version fallback
- **Pure Java PE Parsing**: EXE version extraction using efficient chunk-based binary parsing (no PowerShell dependencies)
- **Automatic Path Fixing**: Detects and corrects corrupted file paths with trailing garbage data and space corruption
- **Split Scan Date Tracking**: Separate LocalScanDate and RemoteScanDate columns for distinct audit trails
- **Improved UNC Compatibility**: Pure Java implementations work seamlessly with remote Windows file access
- **Performance Optimization**: Removed external process dependencies and streamlined operations for faster execution

## Project Overview

This is a single-file Java 8 vulnerability analysis tool that processes Excel files to check file existence and extract metadata across different platforms. The tool is implemented in `ExcelTool.java` and uses Apache POI for Excel processing.

**Purpose**: Comprehensive vulnerability analysis tool that checks file existence, modification dates, and extracts file versions (JAR and EXE) across different platforms based on hostname matching in Excel spreadsheets. Features automatic path corruption detection and fixing, separate audit tracking for local vs remote operations, and seamless remote Windows file access via UNC paths.

## Architecture

- **Single Java File**: All logic in `src/main/java/app/ExcelTool.java` (Java 1.8 compatible)
- **Configuration-Driven**: Uses `config.properties` for Excel file path, column mappings, feature settings, and progress display mode
- **Cross-Platform**: Handles both Windows and Linux path formats
- **Flexible Host Processing**: Processes local host files and optionally remote Windows files via UNC
- **Smart Exclusion**: Maintains exclusion list for inaccessible remote hosts during execution
- **Excel I/O**: Uses Apache POI via WorkbookFactory (supports both .xlsx and .xls)
- **Path Fixing**: Automatic detection and correction of corrupted file paths with trailing suffixes

## Key Components

### Core Logic (src/main/java/app/ExcelTool.java:220-300)
- **4-Pass Processing Architecture**: Streamlined processing with distinct phases for better maintainability
  - **Pass 1**: Path fixing and special replacements for ALL rows
  - **Pass 2**: File operations using FixedFilePath for hostname scope
  - **Pass 3**: Duplicate detection using normalized paths
  - **Pass 4**: Invalid path detection with pattern matching
- Required columns validation (exits without saving if missing)
- Auto-creation of output columns: FileExists, FileModificationDate, FileVersion, ScanError, RemoteScanError, LocalScanDate, RemoteScanDate, FixedFilePath, IsFixed, IsInvalidPath, UniqueID, Duplicate
- **Platform-Aware Processing**: Row-level platform detection for accurate path replacement logic
- **Enhanced Path Fixing**: Handles space corruption, "key=" patterns, and special replacement mappings
- **Unified Cell Writing**: Streamlined operations with writeCells() function for multiple column updates
- **Standardized error handling**: Consistent error recording through dedicated helper methods

### Path Resolution (src/main/java/app/ExcelTool.java:267-270)
- Cross-platform path normalization (converts `\` to `/`)
- Uses `Paths.get().normalize()` for consistent path handling

### File Version Analysis
- **JAR Analysis**: Extracts Implementation-Version from META-INF/MANIFEST.MF using robust parsing with Manifest API and manual text parsing fallback
- **EXE Analysis**: Extracts file version from Windows PE headers using pure Java binary parsing (no PowerShell/external dependencies)
- **PE Header Parsing**: Efficient chunk-based file scanning for VS_FIXEDFILEINFO signatures, handles files of any size
- **Cross-Platform**: JAR analysis works on all platforms, EXE analysis works on Windows platforms only
- **UNC Compatible**: Both JAR and EXE analysis work seamlessly with remote Windows files via UNC paths
- **Performance**: Pure Java implementation avoids external process overhead and UNC compatibility issues
- **Error Handling**: Comprehensive error reporting captures file version processing issues in ScanError column

### Path Corruption Detection and Fixing (Pass 1)
- **3-Step Processing Order**: Optimized sequence for accurate path resolution
  1. **"key=" Pattern Removal**: Handles `filename.ext key=data` corruption first
  2. **Special Path Replacements**: Platform-aware mapping before generic fixing
  3. **Space Corruption Fixing**: Removes trailing garbage only if no replacement applied
- **Platform-Aware Replacements**: Uses case-insensitive comparison on Windows platforms, case-sensitive on Linux
- **Properties File Format**: Standard Java properties format with proper backslash escaping
- **Result Tracking**:
  - `FixedFilePath` column shows the corrected path (always populated)
  - `IsFixed` column indicates whether path fixing was applied (Y/N)
- **Intelligent Fallback**: Generic space corruption fixing only applies when specific replacements don't match
- **Invalid Path Detection**: Separate Pass 4 marks truly invalid paths with `IsInvalidPath` column
- **Configurable**: Can be enabled/disabled via `invalid.path.detection` setting

### Remote Windows Access (UNC Support)
- **UNC Path Conversion**: Converts `C:\path\file.jar` to `\\hostname\C$\path\file.jar`
- **Smart Exclusion**: Hosts that fail UNC access are added to exclusion list (both exception-based and permission-based failures)
- **Data Preservation**: UNC access failures only update `RemoteScanError` column, preserving existing data in other columns
- **Configurable**: Can be enabled/disabled via `remote.unc.enabled` setting
- **Error Handling**: Captures UNC access failures in RemoteScanError column
- **Console Reporting**: Reports inaccessible hosts in real-time and final summary

### Progress Display System
- **Timestamped Logging**: Row-by-row logging with detailed messages and UNC access notifications  
- **Real-time Updates**: Shows current processing status and file being analyzed

### Hostname-Prefixed Logging System
- **Automatic Prefixing**: Log filenames automatically include hostname for multi-machine identification
- **Path Preservation**: Maintains directory structure and file extensions from original configuration
- **Character Normalization**: Invalid filename characters in hostnames replaced with underscores
- **Format**: `hostname-originalname.ext` (e.g., `LPRIME-excel-tool.log` from `excel-tool.log`)

### Scan Date Tracking System
- **Separate Audit Trails**: Maintains distinct timestamps for local vs remote file operations
- **LocalScanDate**: Records session timestamp for files processed on the local host
- **RemoteScanDate**: Records session timestamp for files processed on remote Windows hosts via UNC
- **Session Consistency**: All files processed in the same session receive identical timestamps (session start time)
- **Mutual Exclusivity**: Each row populates either LocalScanDate OR RemoteScanDate, never both
- **Audit Capability**: Enables tracking of when and how files were last scanned (local vs remote access method)

### Duplicate Detection System
- **Configurable Detection**: Enabled/disabled via `duplicate.search.enabled` setting (default: false)
- **UniqueID Generation**: Creates unique identifiers by concatenating normalized hostname, CVE, and file path
- **Smart Path Selection**: Uses FixedFilename if available, otherwise uses original FilePath
- **Normalization**: Hostname (lowercase), CVE (uppercase), file path (lowercase) for consistent comparison
- **HashMap Tracking**: Maintains in-memory map of seen UniqueIDs to detect duplicates
- **Duplicate Marking**: Sets "Y" for duplicates, "N" for first occurrences in Duplicate column
- **Two-Stage Processing**: Initial check with original path, final update with corrected path if applicable

### Error Handling (src/main/java/app/ExcelTool.java:Error columns)
- **ScanError column**: Automatically created to track local file scanning issues
- **RemoteScanError column**: Automatically created to track remote UNC scanning issues
- **Invalid Path Detection**: Extensible pattern system to identify corrupted or invalid file paths
- **FileExists column values**:
  - `Y`: File exists and is accessible
  - `N`: File does not exist (but path is valid)
  - `X`: Invalid path (empty, corrupted, or directory instead of file)
- **Error types captured in ScanError**:
  - Invalid path patterns (empty paths, "N/A", "N\A", directories)
  - Local file access permissions errors (differentiated from file non-existence)
  - JAR processing failures (missing MANIFEST.MF, corrupted files)
  - File modification date read errors
  - Access denied scenarios for local files (using Files.exists() + Files.notExists() logic)
- **Error types captured in RemoteScanError**:
  - UNC access failures for remote Windows hosts
  - UNC access timeouts (hosts unreachable or slow)
  - UNC access denied scenarios
  - Invalid path format for UNC conversion
  - Host exclusion list messages
- **Error behavior**: 
  - ScanError is cleared (set to blank) for successful local scans where file status can be determined
  - RemoteScanError is cleared (set to blank) for successful remote scans
- **Error aggregation**: Multiple errors for same file are combined with `;` separator

### CVE Information Sheet Creation (NIST NVD Integration)
- **CVE Data Fetching**: Retrieves detailed vulnerability information from NIST National Vulnerability Database (NVD) API v2.0
- **Weblogic Detection**: Automatically identifies Oracle WebLogic Server vulnerabilities by analyzing CPE configurations for `weblogic_server` patterns
- **Oracle Advisory Extraction**: Extracts Oracle security advisory URLs from CVE references (supports both HTTP and HTTPS, handles escaped JSON URLs)
- **Sheet Creation**: Creates comprehensive "CVEs" sheet with columns:
  - CVE ID, Description, References, Affected Software (CPE data)
  - Weblogic (Y/N flag for WebLogic vulnerabilities)
  - Weblogic Sec Advisories (Oracle security advisory URLs)
- **Rate Limiting**: Built-in 2-second delays between API requests to respect NIST API limits
- **Error Handling**: Graceful handling of 404 (CVE not found) and 429 (rate limiting) responses
- **Configuration**: Enabled via `cve.sheet.creation.enabled=true` setting (bypasses normal file processing when enabled)
- **Real-World Testing**: Successfully tested with 15+ real WebLogic CVEs from 2020-2024 timeframe

## Build and Development Commands

### Build Options

1. **Build uber JAR (recommended)**:
   - Windows: `build.bat` ✅ (manual dependency extraction with JAVA_HOME support, progress messages, automatic cleanup)
   - Maven: `maven-build.bat` ✅ (Maven-based with automatic dependency resolution)
   - Linux/macOS: `./build.sh`
   - Output: `java-excel-tool-uber.jar`

**Enhanced build.bat Features**:
- **JAVA_HOME Support**: Automatically detects and uses `JAVA_HOME` environment variable, with fallback to default Java 8 installation path if not set
- **Progress Messages**: Shows clear progress indicators for each build stage (compiling, extracting dependencies, creating JAR, etc.)
- **Automatic Cleanup**: Removes temporary build artifacts (`target` folder) after successful build to keep workspace clean

2. **Run the tool**:
   ```bash
   java -jar java-excel-tool-uber.jar config.properties
   ```
   Or on Windows:
   ```cmd
   run.bat
   maven-run.bat
   ```

### Configuration

Edit `config.properties` to set:
- `excel.path`: Path to Excel file
- `sheet.name`: Name of Excel sheet to process
- Column names for PlatformName, FilePath, HostName, CVE
- `platform.windows`: Windows platform values (comma-separated, e.g., "Windows Server 2019, Windows Server 2022")
- `remote.unc.enabled`: Enable/disable remote Windows UNC access (true/false)
- `remote.unc.timeout`: UNC access timeout in seconds (default: 6)
- `log.filename`: Log file name (optional) - saves console output to hostname-prefixed file (e.g., `HOSTNAME-excel-tool.log`)
- `invalid.path.detection`: Enable/disable invalid path pattern validation (true/false, default: true)
- `duplicate.search.enabled`: Enable/disable duplicate detection system (true/false, default: true)
- `cve.sheet.creation.enabled`: Enable/disable CVE information sheet creation with NIST NVD data (true/false, default: false)
- `path.replacements`: Special path replacement mappings (standard Java properties format):
  ```properties
  # Single replacement
  path.replacements=D:\\Apps\\Notepad++ otepad++.exe=D:\\Apps\\Notepad++\\Notepad++.exe

  # Multiple replacements (comma-separated)
  path.replacements=D:\\Apps\\Notepad++ otepad++.exe=D:\\Apps\\Notepad++\\Notepad++.exe,C:\\old\\path\\app.exe=C:\\new\\path\\app.exe

  # With line continuation
  path.replacements=D:\\Apps\\Notepad++ otepad++.exe=D:\\Apps\\Notepad++\\Notepad++.exe,\
                    C:\\old\\path\\app.exe=C:\\new\\path\\app.exe
  ```

### Dependencies

Dependencies are managed through:
- **Manual**: Dependencies stored in `deps/` folder
- **Maven**: Configured in `pom.xml` with automatic resolution

Key dependencies:
- Apache POI 5.4.1 (poi-*.jar)
- Commons libraries for compression and utilities  
- XMLBeans for XML processing

## Exit Codes

- `0`: Success
- `2`: Required columns missing
- `3`: Invalid Excel format
- `4`: Configuration error  
- `5`: Unexpected error
- `6`: Specified sheet does not exist

## Development Notes

- **Java 1.8 Compatibility**: Code must remain compatible with Java 8
- **Security**: Tool is designed for defensive file analysis, not exploitation
- **Error Handling**: Tool exits without saving if required columns are missing
- **Enhanced Error Reporting**: ScanError column captures local scanning issues (cleared for successful scans), RemoteScanError column captures remote UNC issues (cleared for successful remote scans)
- **Standardized Error Handling**: Consistent error recording patterns through dedicated helper methods (`recordScanError`, `recordRemoteScanError`, `addHostToExclusionList`)
- **Code Quality**: Reduced code duplication through helper methods for common operations (`setFileNotFound`, `normalizeHostname`, `clearScanErrors`)
- **Logging**: Uses System.out/System.err for status messages and warnings
- **Cross-Platform**: Handles different path separators and hostname detection methods
- **Date Format**: Human-readable timestamps (`yyyy-MM-dd HH:mm:ss`) instead of ISO format
- **Resource Management**: Proper try-with-resources to prevent memory leaks
- **Remote Access**: Smart UNC path handling with exclusion lists for performance
- **Flexible Configuration**: Boolean settings with sensible defaults, support for multiple platform values
- **Enhanced File Analysis**: Uses Files.exists() + Files.notExists() to differentiate access issues
- **File Type Validation**: Uses Files.isRegularFile() to exclude directories and special files
- **Build System Robustness**: JAVA_HOME detection with intelligent fallbacks, progress messaging, and automatic cleanup
- **Comprehensive Reporting**: Real-time and summary reporting of inaccessible hosts
- **Progress Display**: Timestamped logging for detailed execution tracking
- **Data Preservation**: UNC access failures preserve existing column data while recording errors in RemoteScanError column
- **Counter Consistency**: Accurate tracking of processed vs skipped rows across all failure scenarios
- **Scan Auditing**: LocalScanDate and RemoteScanDate columns record session start timestamp for local and remote hosts respectively (same timestamp for entire scanning session)
- **Invalid Path Handling**: Extensible system for detecting and marking invalid file paths with "X" in FileExists column
- **Optimized Path Fixing**: 3-step processing order ensures special replacements take precedence over generic space corruption fixing
- **PE Header Parsing**: Efficient chunk-based binary parsing for EXE version extraction without external dependencies
- **UNC Compatibility**: Pure Java implementations work seamlessly with remote Windows file access
- **Streamlined Architecture**: 4-pass processing design separates concerns for better maintainability and performance
- **Enhanced Column System**: FixedFilePath, IsFixed, and IsInvalidPath columns provide comprehensive audit trails
- **Platform-Aware Logic**: Row-level platform detection enables accurate Windows vs Linux processing
- **Properties-Based Configuration**: Path replacements use standard Java properties format with proper escaping and platform-specific case sensitivity
- **Unified Operations**: writeCells() function streamlines multiple column updates and reduces code duplication
- **Bundle-Version Support**: JAR version extraction includes Implementation-Version and Bundle-Version fallback for OSGi compatibility
- **Scan Date Separation**: LocalScanDate and RemoteScanDate columns provide separate audit trails for different access methods
- **Hostname-Prefixed Logging**: Log files automatically include hostname prefix for multi-machine deployments (e.g., `HOSTNAME-excel-tool.log`)
- **Maintainability**: Consistent patterns and helper methods improve code readability and reduce maintenance overhead
- **Performance Optimization**: Removed PowerShell dependencies and streamlined processing for faster execution