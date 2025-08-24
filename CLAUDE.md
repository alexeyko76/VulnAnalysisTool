# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a single-file Java 8 vulnerability analysis tool that processes Excel files to check file existence and extract metadata across different platforms. The tool is implemented in `ExcelTool.java` and uses Apache POI for Excel processing.

**Purpose**: Check file existence, modification dates, and JAR versions across different platforms based on hostname matching in Excel spreadsheets.

## Architecture

- **Single Java File**: All logic in `src/main/java/app/ExcelTool.java` (Java 1.8 compatible)
- **Configuration-Driven**: Uses `config.properties` for Excel file path and column mappings
- **Cross-Platform**: Handles both Windows and Linux path formats
- **Hostname Filtering**: Only processes rows matching the current system hostname
- **Excel I/O**: Uses Apache POI via WorkbookFactory (supports both .xlsx and .xls)

## Key Components

### Core Logic (src/main/java/app/ExcelTool.java:45-177)
- Main processing loop that reads Excel rows and updates file status
- Required columns validation (exits without saving if missing)
- Auto-creation of output columns: FileExists, FileModificationDate, JarVersion, ScanError

### Path Resolution (src/main/java/app/ExcelTool.java:267-270)
- Cross-platform path normalization (converts `\` to `/`)
- Uses `Paths.get().normalize()` for consistent path handling

### JAR Analysis (src/main/java/app/ExcelTool.java:316-346)
- Extracts Implementation-Version from META-INF/MANIFEST.MF
- Only processes files with `.jar` extension
- Enhanced error reporting captures JAR processing issues in ScanError column

### Error Handling (src/main/java/app/ExcelTool.java:ScanError column)
- **ScanError column**: Automatically created to track scanning issues
- **Error types captured**:
  - Empty file paths
  - File access permissions errors
  - JAR processing failures (missing MANIFEST.MF, corrupted files)
  - File modification date read errors
- **Error aggregation**: Multiple errors for same file are combined with `;` separator

## Build and Development Commands

### Build Options

1. **Build uber JAR (recommended)**:
   - Windows: `build.bat` ✅ (fixed - now properly includes all dependencies)
   - Linux/macOS: `./build.sh`
   - Output: `java-excel-tool-uber.jar`

2. **Run the tool**:
   ```bash
   java -jar java-excel-tool-uber.jar config.properties
   ```
   Or on Windows:
   ```cmd
   run.bat
   ```

### Configuration

Edit `config.properties` to set:
- `excel.path`: Path to Excel file
- `sheet.name`: Name of Excel sheet to process
- Column names for PlatformName, FilePath, HostName

### Dependencies

Dependencies are stored in `deps/` folder:
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
- **Enhanced Error Reporting**: ScanError column captures and reports all scanning issues
- **Logging**: Uses System.out/System.err for status messages and warnings
- **Cross-Platform**: Handles different path separators and hostname detection methods
- **Date Format**: Human-readable timestamps (`yyyy-MM-dd HH:mm:ss`) instead of ISO format
- **Resource Management**: Proper try-with-resources to prevent memory leaks