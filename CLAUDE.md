# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a single-file Java 8 vulnerability analysis tool that processes Excel files to check file existence and extract metadata across different platforms. The tool is implemented in `ExcelTool.java` and uses Apache POI for Excel processing.

**Purpose**: Check file existence, modification dates, and JAR versions across different platforms based on hostname matching in Excel spreadsheets. Supports both local file access and remote Windows file access via UNC paths with configurable progress display modes.

## Architecture

- **Single Java File**: All logic in `src/main/java/app/ExcelTool.java` (Java 1.8 compatible)
- **Configuration-Driven**: Uses `config.properties` for Excel file path, column mappings, feature settings, and progress display mode
- **Cross-Platform**: Handles both Windows and Linux path formats
- **Flexible Host Processing**: Processes local host files and optionally remote Windows files via UNC
- **Smart Exclusion**: Maintains exclusion list for inaccessible remote hosts during execution
- **Excel I/O**: Uses Apache POI via WorkbookFactory (supports both .xlsx and .xls)

## Key Components

### Core Logic (src/main/java/app/ExcelTool.java:45-177)
- Main processing loop that reads Excel rows and updates file status
- Required columns validation (exits without saving if missing)
- Auto-creation of output columns: FileExists, FileModificationDate, JarVersion, ScanError, RemoteScanError, ScanDate
- **Standardized error handling**: Consistent error recording through dedicated helper methods
- **Code quality improvements**: Reduced duplication via helper methods for common operations

### Path Resolution (src/main/java/app/ExcelTool.java:267-270)
- Cross-platform path normalization (converts `\` to `/`)
- Uses `Paths.get().normalize()` for consistent path handling

### JAR Analysis (src/main/java/app/ExcelTool.java:316-346)
- Extracts Implementation-Version from META-INF/MANIFEST.MF
- Only processes files with `.jar` extension
- **Robust Parsing**: Uses both Manifest API and manual text parsing as fallback
- Enhanced error reporting captures JAR processing issues in ScanError column
- RemoteScanError column captures UNC access issues for remote hosts

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
- Column names for PlatformName, FilePath, HostName
- `platform.windows`: Windows platform values (comma-separated, e.g., "Windows Server 2019, Windows Server 2022")
- `remote.unc.enabled`: Enable/disable remote Windows UNC access (true/false)
- `remote.unc.timeout`: UNC access timeout in seconds (default: 7)
- `log.filename`: Log file name (optional) - saves console output to specified file

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
- **Scan Auditing**: ScanDate column records session start timestamp for all processed hosts (same timestamp for entire scanning session)
- **Invalid Path Handling**: Extensible system for detecting and marking invalid file paths with "X" in FileExists column
- **Maintainability**: Consistent patterns and helper methods improve code readability and reduce maintenance overhead