# Comprehensive Test Cases for Java Vulnerability Analysis Tool

## Excel File Structure: sample.xlsx

**Sheet Name:** Export

**Columns:** AFFECTED_PLATFORMS | XTRACT_PATH | HOSTNAME

Replace `LPRIME` with your actual local hostname in the test cases below.

## Test Case Data (Copy to Excel)

| AFFECTED_PLATFORMS | XTRACT_PATH | HOSTNAME | TEST_CASE_DESCRIPTION | EXPECTED_RESULT |
|-------------------|-------------|----------|----------------------|-----------------|
| Windows Server 2019 | C:\Windows\System32\notepad.exe | LPRIME | Valid EXE - Windows System | FileExists=Y, extract EXE version |
| Windows Server 2019 | C:\Windows\System32\calc.exe | LPRIME | Valid EXE - Calculator | FileExists=Y, extract EXE version |
| Windows Server 2019 | C:\Windows\System32\cmd.exe | LPRIME | Valid EXE - Command Prompt | FileExists=Y, extract EXE version |
| Windows Server 2019 | C:\Program Files\Java\jre\lib\rt.jar | LPRIME | Valid JAR - Java Runtime | FileExists=Y or N, extract JAR version if exists |
| Windows Server 2019 | C:\apps\springframework\spring-core-5.3.21.jar | LPRIME | Valid JAR with version | FileExists=N, no version extraction |
| Windows Server 2019 | C:\Windows\System32\calc.exe some_garbage_data_here | LPRIME | Corrupted EXE - Trailing garbage | FileExists=N, FixedFilename=C:\Windows\System32\calc.exe, FixedFileExists=Y |
| Windows Server 2019 | C:\apps\myapp.jar result.filename=corrupted_data | LPRIME | Corrupted JAR - result.filename | FileExists=N, FixedFilename=C:\apps\myapp.jar, FixedFileExists=N |
| Windows Server 2019 | C:\tools\setup.exe version 1.2.3 build 456 | LPRIME | Corrupted EXE - Multiple words | FileExists=N, FixedFilename=C:\tools\setup.exe, FixedFileExists=N |
| Windows Server 2019 | C:\Windows\explorer.exe additional data [info] | LPRIME | Corrupted EXE - Brackets pattern | FileExists=N, FixedFilename=C:\Windows\explorer.exe, FixedFileExists=Y |
| Windows Server 2019 |  | LPRIME | Empty path | FileExists=X, ScanError=Empty file path |
| Windows Server 2019 | N/A | LPRIME | N/A placeholder | FileExists=X, ScanError=Path marked as N/A |
| Windows Server 2019 | n\a | LPRIME | N\A placeholder | FileExists=X, ScanError=Path marked as N/A |
| Windows Server 2019 | C:\Program Files  | LPRIME | Program Files trailing space | FileExists=X, ScanError=Invalid path format |
| Windows Server 2019 | C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe | LPRIME | Valid Program Files (x86) | FileExists=Y or N (valid processing) |
| Windows Server 2019 | C:\Windows\System32 | LPRIME | Directory not file | FileExists=X, ScanError=Path is a directory |
| Windows Server 2019 | C:\apps\my application.jar | LPRIME | JAR with spaces in filename | FileExists=X, ScanError=JAR filename contains spaces |
| Windows Server 2019 | C:\Windows\System32\kernel32.dll | REMOTE-SERVER-01 | Remote Windows valid | UNC: \\\\REMOTE-SERVER-01\\C$\\Windows\\System32\\kernel32.dll |
| Windows Server 2019 | C:\apps\myapp.exe garbage_data | REMOTE-SERVER-02 | Remote Windows corrupted | Fixed path via UNC |
| Linux RHEL 8 | /usr/bin/java | LINUX-SERVER-01 | Linux platform | Row skipped (hostname mismatch) |
| Windows Server 2022 | C:\Windows\explorer.exe | LPRIME | Windows Server 2022 | Processed (matches platform config) |
| Windows Server 2019 | \\\\shared-server\\apps\\tool.exe | LPRIME | UNC path original | Process as-is (already UNC) |
| Windows Server 2019 | C:\data\README | LPRIME | File without extension | FileExists=Y or N, no version extraction |
| Windows Server 2019 | C:\very\long\path\that\goes\deep\into\subdirectories\app.exe | LPRIME | Long valid path | FileExists=N, normal processing |
| Windows Server 2019 | C:\invalid\path\does\not\exist.exe | LPRIME | Non-existent valid path | FileExists=N, no fixing needed |
| Windows Server 2019 | C:\temp\test.jar extra_content | LPRIME | JAR with trailing content | FixedFilename=C:\temp\test.jar |

## Creating the Excel File

1. **Open Excel**
2. **Create new workbook**
3. **Rename Sheet1 to "Export"**
4. **Add column headers in row 1:**
   - A1: AFFECTED_PLATFORMS
   - B1: XTRACT_PATH
   - C1: HOSTNAME

5. **Copy test case data from the table above (excluding the description columns)**
6. **Replace "LPRIME" with your actual computer hostname**
7. **Save as sample.xlsx in the sample-data directory**

## Test Categories Included

### ✅ **JAR Version Extraction Tests**
- Valid JAR files that exist on most Windows systems
- JAR files that may not exist (tests file-not-found handling)
- Corrupted JAR paths that can be fixed

### ✅ **EXE Version Extraction Tests**
- Common Windows system executables (notepad, calc, cmd)
- Program Files executables
- Corrupted EXE paths with trailing garbage

### ✅ **Path Corruption & Fixing Tests**
- Trailing garbage data after filenames
- result.filename= corruption patterns
- Multiple words/spaces after valid filenames
- Bracket patterns [additional_data]

### ✅ **Invalid Path Detection Tests**
- Empty/blank paths
- N/A placeholder values
- Program Files with trailing spaces (invalid)
- Program Files (x86) paths (valid exception)
- Directories mistakenly listed as files
- JAR filenames with spaces

### ✅ **Remote Windows UNC Tests**
- Valid remote Windows hostnames
- Remote paths with corruption (tests fixing + UNC)
- UNC conversion validation

### ✅ **Platform Filtering Tests**
- Linux platforms (should be skipped)
- Different Windows Server versions
- Hostname matching logic

### ✅ **Edge Cases**
- Files without extensions
- Long paths
- UNC paths as original input
- Non-existent but valid paths

## Expected Tool Behavior

When you run the tool on this test data, you should see:

1. **FileExists column** populated with Y/N/X values
2. **FileModificationDate column** populated for existing files with modification timestamps
3. **FileVersion column** populated for existing JAR/EXE files using pure Java parsing
4. **FixedFilename column** populated when path fixing occurs
5. **FixedFileExists column** showing results of fixed path testing (always populated)
6. **ScanError column** with specific error messages for invalid paths
7. **RemoteScanError column** with UNC access issues for remote hosts
8. **LocalScanDate column** with session timestamp for local host files
9. **RemoteScanDate column** with session timestamp for remote host files

## Running the Test

```bash
java -jar java-excel-tool-uber.jar config.properties
```

Ensure your config.properties has:
```properties
excel.path=./sample-data/sample.xlsx
sheet.name=Export
```