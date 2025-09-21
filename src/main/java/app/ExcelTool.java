package app;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.CellType;

import java.io.*;
import java.net.InetAddress;
import java.net.UnknownHostException;
import java.net.URL;
import java.net.HttpURLConnection;
import java.nio.charset.StandardCharsets;
import java.nio.file.*;
import java.time.Instant;
import java.time.ZoneId;
import java.time.ZonedDateTime;
import java.time.format.DateTimeFormatter;
import java.util.concurrent.*;
import java.util.*;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import java.util.jar.Manifest;
import java.util.regex.Pattern;
import java.util.regex.Matcher;

/**
 * Excel processing utility implemented in a single Java file.
 * Java version: 1.8
 *
 * Behavior is defined by README.md. Configuration is read from a properties file.
 *
 * Usage:
 *   java -jar target/java-excel-tool-jar-with-dependencies.jar [path/to/config.properties]
 * or (no Maven):
 *   java -cp target/java-excel-tool.jar;deps/* app.ExcelTool [path/to/config.properties]
 */
public class ExcelTool {

    // Version information
    private static final String VERSION = "2.0.0";

    // Config keys
    private static final String KEY_EXCEL_PATH = "excel.path";
    private static final String KEY_SHEET_NAME = "sheet.name";
    private static final String KEY_COL_PLATFORM = "column.PlatformName";
    private static final String KEY_COL_FILEPATH = "column.FilePath";
    private static final String KEY_COL_HOSTNAME = "column.HostName";
    private static final String KEY_COL_CVE = "column.CVE";
    private static final String KEY_PLATFORM_WINDOWS = "platform.windows";
    private static final String KEY_REMOTE_UNC_ENABLED = "remote.unc.enabled";
    private static final String KEY_REMOTE_UNC_TIMEOUT = "remote.unc.timeout";
    private static final String KEY_LOG_FILENAME = "log.filename";
    private static final String KEY_INVALID_PATH_DETECTION = "invalid.path.detection";
    private static final String KEY_DUPLICATE_SEARCH_ENABLED = "duplicate.search.enabled";
    private static final String KEY_CVE_SHEET_CREATION_ENABLED = "cve.sheet.creation.enabled";

    // Additional columns to ensure exist
    private static final String COL_FILE_EXISTS = "FileExists";
    private static final String COL_FILE_MOD_DATE = "FileModificationDate";
    private static final String COL_FILE_VERSION = "FileVersion";
    private static final String COL_SCAN_ERROR = "ScanError";
    private static final String COL_REMOTE_SCAN_ERROR = "RemoteScanError";
    private static final String COL_LOCAL_SCAN_DATE = "LocalScanDate";
    private static final String COL_REMOTE_SCAN_DATE = "RemoteScanDate";
    private static final String COL_FIXED_FILENAME = "FixedFilename";
    private static final String COL_FIXED_FILE_EXISTS = "FixedFileExists";
    private static final String COL_UNIQUE_ID = "UniqueID";
    private static final String COL_DUPLICATE = "Duplicate";

    private static final DateTimeFormatter TS_FMT = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");
    
    // Dual logging support
    private static PrintWriter logWriter = null;

    public static void main(String[] args) {
        int exit = 0;
        try {
            String configPath = args != null && args.length > 0 ? args[0] : "config.properties";
            Properties cfg = loadConfig(configPath);

            String excelPath = require(cfg, KEY_EXCEL_PATH);
            String sheetName = require(cfg, KEY_SHEET_NAME);
            String colPlatform = require(cfg, KEY_COL_PLATFORM);
            String colFilePath = require(cfg, KEY_COL_FILEPATH);
            String colHostName = require(cfg, KEY_COL_HOSTNAME);
            String colCVE = require(cfg, KEY_COL_CVE);
            String windowsPlatformValues = require(cfg, KEY_PLATFORM_WINDOWS);
            Set<String> windowsPlatforms = parseWindowsPlatforms(windowsPlatformValues);
            boolean remoteUncEnabled = getBoolean(cfg, KEY_REMOTE_UNC_ENABLED, true);
            int remoteUncTimeout = getInteger(cfg, KEY_REMOTE_UNC_TIMEOUT, 30);
            String logFilename = getString(cfg, KEY_LOG_FILENAME, "");
            boolean invalidPathDetection = getBoolean(cfg, KEY_INVALID_PATH_DETECTION, true);
            boolean duplicateSearchEnabled = getBoolean(cfg, KEY_DUPLICATE_SEARCH_ENABLED, false);
            boolean cveSheetCreationEnabled = getBoolean(cfg, KEY_CVE_SHEET_CREATION_ENABLED, false);
            
            // Get hostname first for log file prefixing
            String localHost = getLocalHostName();
            
            // Initialize log file with hostname prefix if specified
            String prefixedLogFilename = createHostnamePrefixedLogFilename(logFilename, localHost);
            initializeLogFile(prefixedLogFilename);

            logMessage("Excel Vulnerability Analysis Tool v" + VERSION);
            logMessage("Local hostname: " + localHost);
            logMessage("Windows platform values: " + windowsPlatforms);
            logMessage("Remote UNC access enabled: " + remoteUncEnabled);
            if (remoteUncEnabled) {
                logMessage("UNC access timeout: " + remoteUncTimeout + " seconds");
            }
            logMessage("Invalid path detection enabled: " + invalidPathDetection);
            logMessage("Duplicate search enabled: " + duplicateSearchEnabled);
            logMessage("CVE sheet creation enabled: " + cveSheetCreationEnabled);
            if (!isBlank(prefixedLogFilename)) {
                logMessage("Log file: " + prefixedLogFilename);
            }

            File excelFile = new File(excelPath);
            if (!excelFile.exists()) {
                throw new IllegalArgumentException("Excel file does not exist: " + excelFile.getAbsolutePath());
            }

            try (FileInputStream fis = new FileInputStream(excelFile);
                 Workbook wb = WorkbookFactory.create(fis)) {

                // Find the specified sheet by name
                Sheet sheet = wb.getSheet(sheetName);
                if (sheet == null) {
                    logError("Sheet '" + sheetName + "' does not exist in the Excel file.");
                    logError("Available sheets:");
                    for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                        logError("  - " + wb.getSheetName(i));
                    }
                    logError("The tool will exit without processing.");
                    System.exit(6);
                }

                // CVE Sheet Creation Mode - skip file processing and create CVE information sheet
                if (cveSheetCreationEnabled) {
                    logMessage("CVE Sheet Creation Mode: Creating CVEs sheet with NIST NVD data...");
                    createCVESheet(wb, sheet, colCVE);

                    // Save the file with CVE sheet
                    try (FileOutputStream fos = new FileOutputStream(excelFile)) {
                        wb.write(fos);
                    }
                    logMessage("CVE sheet creation completed successfully.");
                    return; // Exit early without file processing
                }
                
                if (sheet.getPhysicalNumberOfRows() == 0) {
                    sheet.createRow(0);
                }
                Row header = sheet.getRow(0);
                if (header == null) header = sheet.createRow(0);

                Map<String, Integer> colIndex = mapHeaderIndices(header);

                // Validate required columns exist; do not save if missing
                List<String> missingRequired = new ArrayList<String>();
                if (!colIndex.containsKey(colPlatform)) missingRequired.add(colPlatform);
                if (!colIndex.containsKey(colFilePath)) missingRequired.add(colFilePath);
                if (!colIndex.containsKey(colHostName)) missingRequired.add(colHostName);
                if (!colIndex.containsKey(colCVE)) missingRequired.add(colCVE);

                if (!missingRequired.isEmpty()) {
                    logError("Required columns missing: " + missingRequired);
                    logError("The tool will exit without saving changes.");
                    System.exit(2);
                }

                // Ensure additional columns exist (create if missing)
                ensureColumn(header, colIndex, COL_FILE_EXISTS);
                ensureColumn(header, colIndex, COL_FILE_MOD_DATE);
                ensureColumn(header, colIndex, COL_FILE_VERSION);
                ensureColumn(header, colIndex, COL_SCAN_ERROR);
                ensureColumn(header, colIndex, COL_REMOTE_SCAN_ERROR);
                ensureColumn(header, colIndex, COL_LOCAL_SCAN_DATE);
                ensureColumn(header, colIndex, COL_REMOTE_SCAN_DATE);
                ensureColumn(header, colIndex, COL_FIXED_FILENAME);
                ensureColumn(header, colIndex, COL_FIXED_FILE_EXISTS);
                ensureColumn(header, colIndex, COL_UNIQUE_ID);
                ensureColumn(header, colIndex, COL_DUPLICATE);

                int idxPlatform = colIndex.get(colPlatform);
                int idxFilePath = colIndex.get(colFilePath);
                int idxHostName = colIndex.get(colHostName);
                int idxCVE = colIndex.get(colCVE);
                int idxFileExists = colIndex.get(COL_FILE_EXISTS);
                int idxFileMod = colIndex.get(COL_FILE_MOD_DATE);
                int idxFileVersion = colIndex.get(COL_FILE_VERSION);
                int idxScanError = colIndex.get(COL_SCAN_ERROR);
                int idxRemoteScanError = colIndex.get(COL_REMOTE_SCAN_ERROR);
                int idxLocalScanDate = colIndex.get(COL_LOCAL_SCAN_DATE);
                int idxRemoteScanDate = colIndex.get(COL_REMOTE_SCAN_DATE);
                int idxFixedFilename = colIndex.get(COL_FIXED_FILENAME);
                int idxFixedFileExists = colIndex.get(COL_FIXED_FILE_EXISTS);
                int idxUniqueID = colIndex.get(COL_UNIQUE_ID);
                int idxDuplicate = colIndex.get(COL_DUPLICATE);

                int processed = 0;
                int skippedHost = 0;
                int skippedRemote = 0;
                
                // Exclusion list for hosts that cannot be accessed via UNC from current machine
                Set<String> inaccessibleHosts = new HashSet<String>();

                // Duplicate detection HashMap (only used if duplicate search is enabled)
                Map<String, Boolean> uniqueIDMap = new HashMap<String, Boolean>();

                // Single timestamp for entire scanning session
                String sessionScanTimestamp = getCurrentTimestamp();
                
                // Progress tracking
                int totalRows = sheet.getLastRowNum();
                
                // Phase 1: Duplicate detection for ALL rows (if enabled)
                if (duplicateSearchEnabled) {
                    logMessage("Phase 1: Processing duplicate detection for " + totalRows + " rows:");
                    for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                        Row row = sheet.getRow(r);
                        if (row == null) continue;

                        String targetHost = getStringCell(row, idxHostName);
                        String cveValue = getStringCell(row, idxCVE);
                        String rawPath = getStringCell(row, idxFilePath);

                        // Generate initial UniqueID (may be updated later if path gets fixed)
                        String uniqueID = generateUniqueID(targetHost, cveValue, rawPath, "");
                        writeCell(row, idxUniqueID, uniqueID);

                        // Check for duplicates and set Duplicate column
                        boolean isDuplicate = uniqueIDMap.containsKey(uniqueID);
                        writeCell(row, idxDuplicate, isDuplicate ? "Y" : "N");
                        uniqueIDMap.put(uniqueID, true);
                    }
                } else {
                    // Clear duplicate detection columns when disabled
                    for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                        Row row = sheet.getRow(r);
                        if (row == null) continue;
                        writeCell(row, idxUniqueID, "");
                        writeCell(row, idxDuplicate, "");
                    }
                }

                logMessage("Phase 2: Processing file operations for " + totalRows + " rows:");

                for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                    Row row = sheet.getRow(r);
                    if (row == null) continue;

                    String targetHost = getStringCell(row, idxHostName);
                    String targetPlatform = getStringCell(row, idxPlatform);
                    boolean isLocalHost = !isBlank(targetHost) && targetHost.trim().equalsIgnoreCase(localHost);
                    boolean isWindowsPlatform = !isBlank(targetPlatform) && isWindowsPlatformMatch(targetPlatform, windowsPlatforms);
                    boolean isRemoteWindows = remoteUncEnabled && isWindowsPlatform && !isLocalHost && !isBlank(targetHost);

                    // Skip if not local host and (UNC disabled or not a Windows platform for remote access)
                    if (!isLocalHost && !isRemoteWindows) {
                        skippedHost++;
                        continue;
                    }
                    
                    // Record scan timestamp for hosts that will be processed
                    if (isLocalHost) {
                        writeCell(row, idxLocalScanDate, sessionScanTimestamp);
                    } else if (isRemoteWindows) {
                        writeCell(row, idxRemoteScanDate, sessionScanTimestamp);
                    }
                    
                    // Skip if remote host is in exclusion list (but update RemoteScanError for this row)
                    if (isRemoteWindows && inaccessibleHosts.contains(normalizeHostname(targetHost))) {
                        recordRemoteScanError(row, idxRemoteScanError, "Host previously identified as inaccessible via UNC");
                        skippedRemote++;
                        continue;
                    }

                    String rawPath = getStringCell(row, idxFilePath);

                    // Initialize fixed path columns
                    writeCell(row, idxFixedFilename, "");
                    writeCell(row, idxFixedFileExists, "");

                    // Step 1: First try the original path
                    FileProcessingResult originalResult = processFilePath(rawPath, targetHost, isRemoteWindows,
                                                                        isWindowsPlatform, r, totalRows,
                                                                        inaccessibleHosts, remoteUncTimeout);

                    if (originalResult.shouldSkip) {
                        if (originalResult.isRemoteSkip) {
                            skippedRemote++;
                        } else {
                            processed++;
                        }
                        if (originalResult.errorMessage != null) {
                            if (originalResult.isRemoteError) {
                                recordRemoteScanError(row, idxRemoteScanError, originalResult.errorMessage);
                            } else {
                                recordScanError(row, idxScanError, originalResult.errorMessage);
                            }
                        }
                        continue;
                    }

                    boolean originalExists = originalResult.exists;
                    Path originalResolved = originalResult.resolvedPath;

                    // Step 2: Always attempt path fixing to provide FixedFileExists information
                    boolean useFixedPath = false;
                    FileProcessingResult fixedResult = null;

                    if (invalidPathDetection) {
                        PathFixResult pathFixResult = attemptPathFix(rawPath, isWindowsPlatform);

                        if (pathFixResult.errorReason != null && !pathFixResult.wasFixed) {
                            // Path is corrupted and cannot be fixed
                            if (!originalExists) {
                                // Only mark as invalid if original also doesn't exist
                                writeCell(row, idxFileExists, "X");
                                writeCell(row, idxFileMod, "");
                                writeCell(row, idxFileVersion, "");
                                recordScanError(row, idxScanError, pathFixResult.errorReason);
                                processed++;
                                continue;
                            }
                            // If original exists but path has corruption patterns, keep original result
                        } else if (pathFixResult.wasFixed) {
                            // Path was fixed - always test the fixed path for FixedFileExists
                            writeCell(row, idxFixedFilename, pathFixResult.fixedPath);

                            fixedResult = processFilePath(pathFixResult.fixedPath, targetHost, isRemoteWindows,
                                                        isWindowsPlatform, r, totalRows,
                                                        inaccessibleHosts, remoteUncTimeout);

                            if (fixedResult.shouldSkip) {
                                writeCell(row, idxFixedFileExists, "Error");
                                if (fixedResult.errorMessage != null) {
                                    if (fixedResult.isRemoteError) {
                                        recordRemoteScanError(row, idxRemoteScanError, "Fixed path: " + fixedResult.errorMessage);
                                    } else {
                                        recordScanError(row, idxScanError, "Fixed path: " + fixedResult.errorMessage);
                                    }
                                }
                            } else {
                                writeCell(row, idxFixedFileExists, fixedResult.exists ? "Y" : "N");
                                // Use fixed path for metadata extraction only if original doesn't exist
                                if (!originalExists && fixedResult.exists) {
                                    useFixedPath = true;
                                }
                            }
                        } else {
                            // Path is already valid, but always check if there's a potential fix
                            String potentialFixedPath = tryRemoveTrailingSuffix(rawPath.trim());
                            if (potentialFixedPath != null && !potentialFixedPath.equals(rawPath.trim())) {
                                // There was a potential fix, test it for FixedFileExists
                                writeCell(row, idxFixedFilename, potentialFixedPath);

                                FileProcessingResult potentialResult = processFilePath(potentialFixedPath, targetHost, isRemoteWindows,
                                                                    isWindowsPlatform, r, totalRows,
                                                                    inaccessibleHosts, remoteUncTimeout);

                                if (potentialResult.shouldSkip) {
                                    writeCell(row, idxFixedFileExists, "Error");
                                } else {
                                    writeCell(row, idxFixedFileExists, potentialResult.exists ? "Y" : "N");
                                }
                            } else {
                                // No potential fix available - FixedFileExists should be same as original
                                writeCell(row, idxFixedFileExists, originalExists ? "Y" : "N");
                            }
                        }
                    } else {
                        // Invalid path detection is disabled - always try potential path fix for FixedFileExists
                        String potentialFixedPath = tryRemoveTrailingSuffix(rawPath.trim());
                        if (potentialFixedPath != null && !potentialFixedPath.equals(rawPath.trim())) {
                            // There was a potential fix, test it for FixedFileExists
                            writeCell(row, idxFixedFilename, potentialFixedPath);

                            FileProcessingResult potentialResult = processFilePath(potentialFixedPath, targetHost, isRemoteWindows,
                                                                isWindowsPlatform, r, totalRows,
                                                                inaccessibleHosts, remoteUncTimeout);

                            if (potentialResult.shouldSkip) {
                                writeCell(row, idxFixedFileExists, "Error");
                            } else {
                                writeCell(row, idxFixedFileExists, potentialResult.exists ? "Y" : "N");
                            }
                        } else {
                            // No potential fix available - FixedFileExists should be same as original
                            writeCell(row, idxFixedFileExists, originalExists ? "Y" : "N");
                        }
                    }

                    // Step 3: Determine which file to process for metadata extraction
                    Path resolvedForMetadata;
                    boolean existsForMetadata;

                    if (useFixedPath && fixedResult != null) {
                        resolvedForMetadata = fixedResult.resolvedPath;
                        existsForMetadata = fixedResult.exists;
                        writeCell(row, idxFileExists, originalExists ? "Y" : "N"); // Original file status
                    } else {
                        resolvedForMetadata = originalResolved;
                        existsForMetadata = originalExists;
                        writeCell(row, idxFileExists, originalExists ? "Y" : "N");
                    }

                    // Update UniqueID with the correct file path (if duplicate search is enabled)
                    if (duplicateSearchEnabled) {
                        String cveValue = getStringCell(row, idxCVE);
                        String fixedFilename = getStringCell(row, idxFixedFilename);
                        String updatedUniqueID = generateUniqueID(targetHost, cveValue, rawPath, fixedFilename);

                        // Always update the UniqueID with the corrected path
                        writeCell(row, idxUniqueID, updatedUniqueID);
                    }

                    // Step 4: Extract metadata (modification date and version) if file exists
                    if (existsForMetadata) {
                        StringBuilder scanErrors = new StringBuilder();

                        try {
                            Instant lm = Files.getLastModifiedTime(resolvedForMetadata).toInstant();
                            ZonedDateTime zdt = ZonedDateTime.ofInstant(lm, ZoneId.systemDefault());
                            writeCell(row, idxFileMod, TS_FMT.format(zdt.toLocalDateTime()));
                        } catch (IOException e) {
                            writeCell(row, idxFileMod, "");
                            scanErrors.append("Cannot read modification date: ").append(e.getMessage());
                            logError("WARN: Could not read last modified for: " + resolvedForMetadata + " -> " + e.getMessage());
                        }

                        // File version handling (JAR and EXE files)
                        if (resolvedForMetadata.getFileName() != null) {
                            FileVersionResult result = extractFileVersion(resolvedForMetadata.toFile(), isWindowsPlatform);
                            writeCell(row, idxFileVersion, result.version != null ? result.version : "");
                            if (result.error != null) {
                                if (scanErrors.length() > 0) scanErrors.append("; ");
                                scanErrors.append("File version processing error: ").append(result.error);
                            }
                        } else {
                            writeCell(row, idxFileVersion, "");
                        }

                        // Write scan errors or clear if none
                        writeCell(row, idxScanError, scanErrors.length() > 0 ? scanErrors.toString() : "");
                    } else {
                        writeCell(row, idxFileMod, "");
                        writeCell(row, idxFileVersion, "");
                        // Clear ScanError for successful determination that file doesn't exist
                        if (!isRemoteWindows) {
                            recordScanError(row, idxScanError, ""); // Clear error for successful scan
                        }
                    }

                    // Clear error columns for successful scans
                    clearScanErrors(row, idxScanError, idxRemoteScanError, isRemoteWindows);

                    processed++;
                }

                // Phase 3: Rebuild duplicate detection with corrected UniqueIDs (if enabled)
                if (duplicateSearchEnabled) {
                    logMessage("Phase 3: Rebuilding duplicate detection with corrected paths:");
                    uniqueIDMap.clear(); // Clear the map to rebuild it

                    for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                        Row row = sheet.getRow(r);
                        if (row == null) continue;

                        String currentUniqueID = getStringCell(row, idxUniqueID);
                        if (!isBlank(currentUniqueID)) {
                            // Check for duplicates with the final UniqueID
                            boolean isDuplicate = uniqueIDMap.containsKey(currentUniqueID);
                            writeCell(row, idxDuplicate, isDuplicate ? "Y" : "N");
                            uniqueIDMap.put(currentUniqueID, true);
                        }
                    }
                }

                // Save back to the same file
                try (FileOutputStream fos = new FileOutputStream(excelFile)) {
                    wb.write(fos);
                }
                
                // Print final results
                logMessage("Done. Rows processed: " + processed + ", skipped (hostname mismatch): " + skippedHost + ", skipped (remote inaccessible): " + skippedRemote);
                if (!inaccessibleHosts.isEmpty()) {
                    logMessage("Inaccessible hosts identified during this run: " + inaccessibleHosts);
                }
            }
        } catch (java.io.IOException e) {
            logError("IO failure: " + e.getMessage());
            exit = 3;
        } catch (IllegalArgumentException e) {
            logError(e.getMessage());
            exit = 4;
        } catch (Exception e) {
            logError("Unexpected failure: " + e.getMessage());
            e.printStackTrace(System.err);
            if (logWriter != null) {
                e.printStackTrace(logWriter);
            }
            exit = 5;
        }
        if (exit != 0) {
            System.exit(exit);
        }
        
        // Close log file
        closeLogFile();
    }

    private static Properties loadConfig(String path) throws IOException {
        Properties p = new Properties();
        File f = new File(path);
        if (!f.exists()) {
            throw new IllegalArgumentException("Config file not found: " + f.getAbsolutePath());
        }
        try (InputStream is = new FileInputStream(f);
             Reader reader = new InputStreamReader(is, StandardCharsets.UTF_8)) {
            p.load(reader);
        }
        return p;
    }

    private static String require(Properties p, String key) {
        String v = p.getProperty(key);
        if (v == null || v.trim().isEmpty()) {
            throw new IllegalArgumentException("Missing config property: " + key);
        }
        return v.trim();
    }

    private static boolean getBoolean(Properties p, String key, boolean defaultValue) {
        String v = p.getProperty(key);
        if (v == null || v.trim().isEmpty()) {
            return defaultValue;
        }
        return Boolean.parseBoolean(v.trim());
    }

    private static String getString(Properties p, String key, String defaultValue) {
        String v = p.getProperty(key);
        if (v == null || v.trim().isEmpty()) {
            return defaultValue;
        }
        return v.trim();
    }

    private static int getInteger(Properties p, String key, int defaultValue) {
        String v = p.getProperty(key);
        if (v == null || v.trim().isEmpty()) {
            return defaultValue;
        }
        try {
            return Integer.parseInt(v.trim());
        } catch (NumberFormatException e) {
            return defaultValue;
        }
    }

    private static Map<String, Integer> mapHeaderIndices(Row header) {
        Map<String, Integer> map = new HashMap<String, Integer>();
        short lastCell = header.getLastCellNum();
        if (lastCell < 0) lastCell = 0;
        for (int i = 0; i < lastCell; i++) {
            Cell c = header.getCell(i, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
            if (c == null) continue;
            String name = c.getStringCellValue();
            if (name != null && !name.trim().isEmpty()) {
                map.put(name.trim(), i);
            }
        }
        return map;
    }

    private static void ensureColumn(Row header, Map<String, Integer> map, String name) {
        if (!map.containsKey(name)) {
            int idx = header.getLastCellNum();
            if (idx < 0) idx = 0;
            Cell cell = header.createCell(idx, CellType.STRING);
            cell.setCellValue(name);
            map.put(name, idx);
        }
    }

    private static String getLocalHostName() {
        try {
            String host = InetAddress.getLocalHost().getHostName();
            if (host != null && !host.trim().isEmpty()) return host.trim();
        } catch (UnknownHostException ignored) {}
        String envHost = System.getenv("COMPUTERNAME");
        if (envHost == null || envHost.trim().isEmpty()) {
            envHost = System.getenv("HOSTNAME");
        }
        if (envHost != null && !envHost.trim().isEmpty()) return envHost.trim();
        return "UNKNOWN_HOST";
    }

    private static String getStringCell(Row row, int idx) {
        Cell c = row.getCell(idx, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
        if (c == null) return "";
        if (c.getCellType() == CellType.STRING) {
            return c.getStringCellValue() != null ? c.getStringCellValue().trim() : "";
        } else if (c.getCellType() == CellType.NUMERIC) {
            double d = c.getNumericCellValue();
            if (Math.floor(d) == d) {
                return String.valueOf((long) d);
            } else {
                return String.valueOf(d);
            }
        } else if (c.getCellType() == CellType.BOOLEAN) {
            return String.valueOf(c.getBooleanCellValue());
        } else {
            return "";
        }
    }

    private static void writeCell(Row row, int idx, String value) {
        Cell c = row.getCell(idx, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
        c.setCellValue(value == null ? "" : value);
    }

    private static boolean isBlank(String s) {
        return s == null || s.trim().isEmpty();
    }

    private static Path resolvePathCrossPlatform(String raw) {
        String normalized = raw.replace('\\', '/');
        return Paths.get(normalized).normalize();
    }

    private static String convertToUncPath(String hostname, String localPath) {
        if (isBlank(hostname) || isBlank(localPath)) {
            return null;
        }
        
        String trimmedPath = localPath.trim();
        
        // Handle Windows drive paths like C:\path or C:/path
        if (trimmedPath.length() >= 2 && trimmedPath.charAt(1) == ':') {
            char driveLetter = trimmedPath.charAt(0);
            String restOfPath = trimmedPath.length() > 2 ? trimmedPath.substring(2) : "";
            
            // Convert to UNC format: \\hostname\drive$\rest_of_path
            String uncPath = "\\\\" + hostname + "\\" + driveLetter + "$" + restOfPath.replace('/', '\\');
            return uncPath;
        }
        
        // Handle UNC paths that are already in UNC format - just return as is
        if (trimmedPath.startsWith("\\\\")) {
            return trimmedPath;
        }
        
        // Cannot convert other path formats
        return null;
    }

    
    private static class FileVersionResult {
        String version;
        String error;

        FileVersionResult(String version, String error) {
            this.version = version;
            this.error = error;
        }
    }
    
    private static FileVersionResult extractFileVersion(File file, boolean isWindowsPlatform) {
        if (file == null || !file.exists()) {
            return new FileVersionResult(null, "File does not exist");
        }

        String fileName = file.getName().toLowerCase(Locale.ENGLISH);

        if (fileName.endsWith(".jar")) {
            return extractJarVersion(file);
        } else if (fileName.endsWith(".exe") && isWindowsPlatform) {
            return extractExeVersion(file);
        } else {
            return new FileVersionResult(null, null); // No version extraction needed for other file types
        }
    }

    private static FileVersionResult extractJarVersion(File jarFile) {
        if (jarFile == null || !jarFile.exists()) {
            return new FileVersionResult(null, "JAR file does not exist");
        }

        ZipFile zip = null;
        try {
            zip = new ZipFile(jarFile);
            ZipEntry entry = zip.getEntry("META-INF/MANIFEST.MF");
            if (entry == null) {
                return new FileVersionResult(null, "No MANIFEST.MF found in JAR");
            }

            try (InputStream is = zip.getInputStream(entry)) {
                Manifest mf = new Manifest(is);
                String v = mf.getMainAttributes().getValue("Implementation-Version");
                if (v != null) {
                    return new FileVersionResult(v.trim(), null);
                } else {
                    // Fallback: manually parse MANIFEST.MF for Implementation-Version
                    String manualVersion = parseManifestManually(zip, entry);
                    if (manualVersion != null) {
                        return new FileVersionResult(manualVersion, null);
                    } else {
                        return new FileVersionResult(null, "No Implementation-Version in MANIFEST.MF");
                    }
                }
            }
        } catch (IOException e) {
            logError("WARN: Failed to read manifest from jar: " + jarFile + " -> " + e.getMessage());
            return new FileVersionResult(null, e.getMessage());
        } finally {
            if (zip != null) {
                try { zip.close(); } catch (IOException ignored) {}
            }
        }
    }

    private static FileVersionResult extractExeVersion(File exeFile) {
        if (exeFile == null || !exeFile.exists()) {
            return new FileVersionResult(null, "EXE file does not exist");
        }

        // Use pure Java PE parsing - works with UNC paths and is much faster
        return extractExeVersionUsingPEHeaders(exeFile);
    }

    private static FileVersionResult extractExeVersionUsingSystem(File exeFile) {
        try {
            // Use PowerShell Get-ItemProperty which is more reliable
            String filePath = exeFile.getAbsolutePath();
            ProcessBuilder pb = new ProcessBuilder(
                "powershell.exe", "-Command",
                "(Get-ItemProperty '" + filePath + "').VersionInfo.FileVersion"
            );

            Process process = pb.start();
            StringBuilder output = new StringBuilder();
            StringBuilder errorOutput = new StringBuilder();

            try (BufferedReader reader = new BufferedReader(new InputStreamReader(process.getInputStream()))) {
                String line;
                while ((line = reader.readLine()) != null) {
                    output.append(line.trim());
                }
            }

            try (BufferedReader reader = new BufferedReader(new InputStreamReader(process.getErrorStream()))) {
                String line;
                while ((line = reader.readLine()) != null) {
                    errorOutput.append(line.trim()).append(" ");
                }
            }

            int exitCode = process.waitFor();

            if (exitCode == 0) {
                String version = output.toString().trim();
                if (!version.isEmpty() && !version.equals("null") && !version.equals("")) {
                    return new FileVersionResult(version, null);
                } else {
                    logError("DEBUG: PowerShell returned empty version for: " + filePath + " (output: '" + version + "')");
                }
            } else {
                logError("DEBUG: PowerShell failed for: " + filePath + " (exit code: " + exitCode + ", error: " + errorOutput.toString() + ")");
            }

            // If PowerShell fails, try WMIC as fallback
            return extractExeVersionUsingWMIC(exeFile);

        } catch (Exception e) {
            // If PowerShell fails, try WMIC as fallback
            return extractExeVersionUsingWMIC(exeFile);
        }
    }

    private static FileVersionResult extractExeVersionUsingWMIC(File exeFile) {
        try {
            // Use wmic to get file version information
            ProcessBuilder pb = new ProcessBuilder(
                "wmic", "datafile",
                "where", "name='" + exeFile.getAbsolutePath().replace("\\", "\\\\") + "'",
                "get", "Version", "/format:list"
            );

            Process process = pb.start();
            StringBuilder output = new StringBuilder();

            try (BufferedReader reader = new BufferedReader(new InputStreamReader(process.getInputStream()))) {
                String line;
                while ((line = reader.readLine()) != null) {
                    line = line.trim();
                    if (line.startsWith("Version=") && !line.equals("Version=")) {
                        String version = line.substring("Version=".length()).trim();
                        if (!version.isEmpty() && !version.equals("null")) {
                            return new FileVersionResult(version, null);
                        }
                    }
                }
            }

            int exitCode = process.waitFor();
            if (exitCode != 0) {
                return new FileVersionResult(null, "Both PowerShell and WMIC failed to extract version");
            }

            return new FileVersionResult(null, "No version found using system methods");

        } catch (Exception e) {
            return new FileVersionResult(null, "System version extraction error: " + e.getMessage());
        }
    }

    private static FileVersionResult extractExeVersionUsingPEHeaders(File exeFile) {
        try (RandomAccessFile raf = new RandomAccessFile(exeFile, "r")) {

            // Simple brute-force approach: scan the entire file for VS_FIXEDFILEINFO signature
            // This is much more reliable than trying to parse the complex PE resource structure

            long fileLength = raf.length();

            // Read file in chunks and search for version signature
            int chunkSize = 8192;
            byte[] buffer = new byte[chunkSize + 64]; // Extra bytes for signature overlap
            long position = 0;

            while (position < fileLength) {
                raf.seek(position);
                int bytesRead = raf.read(buffer, 0, Math.min(buffer.length, (int)(fileLength - position)));

                if (bytesRead < 52) break; // Need at least 52 bytes for VS_FIXEDFILEINFO

                // Look for VS_FIXEDFILEINFO signature (0xFEEF04BD) in this chunk
                for (int i = 0; i <= bytesRead - 52; i++) {
                    int signature = readLittleEndianInt(buffer, i);
                    if (signature == 0xFEEF04BD) {
                        // Found VS_FIXEDFILEINFO structure!

                        // Extract file version (offsets +8 to +15 from signature)
                        int fileVersionMS = readLittleEndianInt(buffer, i + 8);
                        int fileVersionLS = readLittleEndianInt(buffer, i + 12);

                        int major = (fileVersionMS >> 16) & 0xFFFF;
                        int minor = fileVersionMS & 0xFFFF;
                        int build = (fileVersionLS >> 16) & 0xFFFF;
                        int revision = fileVersionLS & 0xFFFF;

                        // Only return version if it's not all zeros
                        if (major != 0 || minor != 0 || build != 0 || revision != 0) {
                            String version = major + "." + minor + "." + build + "." + revision;
                            return new FileVersionResult(version, null);
                        }
                    }
                }

                // Move to next chunk with overlap to avoid missing signatures at chunk boundaries
                position += chunkSize;
            }

            return new FileVersionResult(null, "No version info found in file");

        } catch (Exception e) {
            return new FileVersionResult(null, "Error reading file: " + e.getMessage());
        }
    }


    private static int readLittleEndianInt(byte[] buffer, int offset) {
        return (buffer[offset] & 0xFF) |
               ((buffer[offset + 1] & 0xFF) << 8) |
               ((buffer[offset + 2] & 0xFF) << 16) |
               ((buffer[offset + 3] & 0xFF) << 24);
    }

    private static int readLittleEndianShort(byte[] buffer, int offset) {
        return (buffer[offset] & 0xFF) |
               ((buffer[offset + 1] & 0xFF) << 8);
    }
    
    private static String parseManifestManually(ZipFile zip, ZipEntry entry) {
        try (InputStream is = zip.getInputStream(entry);
             BufferedReader reader = new BufferedReader(new InputStreamReader(is, StandardCharsets.UTF_8))) {
            
            String line;
            StringBuilder continuation = new StringBuilder();
            
            while ((line = reader.readLine()) != null) {
                // Handle line continuation (lines starting with space)
                if (line.startsWith(" ") || line.startsWith("\t")) {
                    continuation.append(line.substring(1));
                    continue;
                }
                
                // Process the complete line (previous + current)
                String completeLine = continuation.toString();
                continuation.setLength(0);
                continuation.append(line);
                
                if (completeLine.startsWith("Implementation-Version:")) {
                    String version = completeLine.substring("Implementation-Version:".length()).trim();
                    if (!version.isEmpty()) {
                        return version;
                    }
                }
            }
            
            // Check the last line
            String lastLine = continuation.toString();
            if (lastLine.startsWith("Implementation-Version:")) {
                String version = lastLine.substring("Implementation-Version:".length()).trim();
                if (!version.isEmpty()) {
                    return version;
                }
            }
            
        } catch (IOException e) {
            logError("WARN: Failed to manually parse manifest: " + e.getMessage());
        }
        return null;
    }
    
    private static void updateProgress(int currentRow, int totalRows, String currentFile) {
        // Verbose mode - simple row-by-row logging with timestamp
        logMessage(getCurrentTimestamp() + " Row " + currentRow + "/" + totalRows + ": " + currentFile);
    }
    
    private static void printVerboseMessage(String message) {
        // Indent additional messages for clarity with timestamp
        logMessage(getCurrentTimestamp() + "   -> " + message);
    }
    
    private static String getCurrentTimestamp() {
        return ZonedDateTime.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss"));
    }
    
    // Helper class for UNC access results
    private static class UncAccessResult {
        boolean exists = false;
        boolean timedOut = false;
        boolean accessDenied = false;
        Exception exception = null;
        
        UncAccessResult(boolean exists) {
            this.exists = exists;
        }
        
        UncAccessResult(Exception exception) {
            this.exception = exception;
        }
        
        static UncAccessResult timeout() {
            UncAccessResult result = new UncAccessResult(false);
            result.timedOut = true;
            return result;
        }
        
        static UncAccessResult accessDenied() {
            UncAccessResult result = new UncAccessResult(false);
            result.accessDenied = true;
            return result;
        }
    }
    
    // Timeout-protected UNC path access
    private static UncAccessResult checkUncPathWithTimeout(Path uncPath, int timeoutSeconds) {
        ExecutorService executor = Executors.newSingleThreadExecutor();
        Future<UncAccessResult> future = executor.submit(new Callable<UncAccessResult>() {
            @Override
            public UncAccessResult call() {
                try {
                    boolean exists = Files.exists(uncPath);
                    return new UncAccessResult(exists);
                } catch (Exception e) {
                    return new UncAccessResult(e);
                }
            }
        });
        
        try {
            // Wait for result with timeout
            return future.get(timeoutSeconds, TimeUnit.SECONDS);
        } catch (TimeoutException e) {
            // Cancel the task and return timeout result
            future.cancel(true);
            return UncAccessResult.timeout();
        } catch (Exception e) {
            return new UncAccessResult(e);
        } finally {
            executor.shutdown();
            try {
                if (!executor.awaitTermination(1, TimeUnit.SECONDS)) {
                    executor.shutdownNow();
                }
            } catch (InterruptedException e) {
                executor.shutdownNow();
            }
        }
    }
    
    // Dual logging utility methods
    private static void initializeLogFile(String logFilename) {
        if (isBlank(logFilename)) {
            return; // No log file specified
        }
        
        try {
            File logFile = new File(logFilename);
            logWriter = new PrintWriter(new FileWriter(logFile, false), true); // Overwrite existing log
            logWriter.println("=== Excel Tool Log Started at " + getCurrentTimestamp() + " ===");
        } catch (IOException e) {
            System.err.println("WARNING: Could not create log file " + logFilename + ": " + e.getMessage());
            logWriter = null;
        }
    }
    
    private static void closeLogFile() {
        if (logWriter != null) {
            logWriter.println("=== Excel Tool Log Ended at " + getCurrentTimestamp() + " ===");
            logWriter.close();
            logWriter = null;
        }
    }
    
    private static void logMessage(String message) {
        // Always output to console
        System.out.println(message);
        
        // Also write to log file if available
        if (logWriter != null) {
            logWriter.println(message);
        }
    }
    
    private static void logError(String message) {
        // Always output to console error
        System.err.println(message);
        
        // Also write to log file if available (don't prefix ERROR if message already has a prefix)
        if (logWriter != null) {
            if (message.startsWith("WARN:") || message.startsWith("ERROR:")) {
                logWriter.println(message);
            } else {
                logWriter.println("ERROR: " + message);
            }
        }
    }
    
    // Create hostname-prefixed log filename
    private static String createHostnamePrefixedLogFilename(String logFilename, String hostname) {
        if (isBlank(logFilename)) {
            return ""; // No log file specified
        }
        
        // Normalize hostname for filename use (remove invalid characters)
        String normalizedHostname = hostname.replaceAll("[^a-zA-Z0-9.-]", "_");
        
        // Extract directory, filename, and extension
        File logFile = new File(logFilename);
        String dir = logFile.getParent();
        String name = logFile.getName();
        
        // Split filename and extension
        int lastDot = name.lastIndexOf('.');
        String baseName, extension;
        if (lastDot > 0 && lastDot < name.length() - 1) {
            baseName = name.substring(0, lastDot);
            extension = name.substring(lastDot); // includes the dot
        } else {
            baseName = name;
            extension = "";
        }
        
        // Create prefixed filename: hostname-originalname.ext
        String prefixedName = normalizedHostname + "-" + baseName + extension;
        
        // Combine with directory if present
        if (dir != null) {
            return new File(dir, prefixedName).getPath();
        } else {
            return prefixedName;
        }
    }
    
    // Standardized error handling methods
    private static void recordScanError(Row row, int idxScanError, String errorMessage) {
        writeCell(row, idxScanError, errorMessage);
    }
    
    private static void recordRemoteScanError(Row row, int idxRemoteScanError, String errorMessage) {
        writeCell(row, idxRemoteScanError, errorMessage);
    }
    
    private static void clearScanErrors(Row row, int idxScanError, int idxRemoteScanError, boolean isRemoteWindows) {
        if (isRemoteWindows) {
            writeCell(row, idxRemoteScanError, "");
        } else {
            writeCell(row, idxScanError, "");
        }
    }
    
    private static void addHostToExclusionList(Set<String> inaccessibleHosts, String hostname, 
                                              Row row, int idxRemoteScanError, String reason) {
        String normalizedHostname = normalizeHostname(hostname);
        inaccessibleHosts.add(normalizedHostname);
        recordRemoteScanError(row, idxRemoteScanError, reason);
        printVerboseMessage("Added " + hostname.trim() + " to exclusion list - " + getReasonForLogging(reason));
    }
    
    private static String getReasonForLogging(String reason) {
        if (reason.contains("timeout")) return "UNC access timeout";
        if (reason.contains("access denied")) return "access denied detected";
        if (reason.contains("Cannot access")) return "UNC access failed";
        return reason.toLowerCase();
    }
    
    private static void setFileNotFound(Row row, int idxFileExists, int idxFileMod, int idxFileVersion) {
        writeCell(row, idxFileExists, "N");
        writeCell(row, idxFileMod, "");
        writeCell(row, idxFileVersion, "");
        // Note: FixedFilename is handled separately in the main processing logic
    }
    
    // Helper method for hostname normalization to reduce repeated toLowerCase() calls
    private static String normalizeHostname(String hostname) {
        return isBlank(hostname) ? "" : hostname.trim().toLowerCase();
    }
    
    // Parse comma-separated Windows platform values, handling spaces in values
    private static Set<String> parseWindowsPlatforms(String platformValues) {
        Set<String> platforms = new HashSet<String>();
        if (!isBlank(platformValues)) {
            String[] values = platformValues.split(",");
            for (String value : values) {
                String trimmed = value.trim();
                if (!trimmed.isEmpty()) {
                    platforms.add(trimmed);
                }
            }
        }
        return platforms;
    }
    
    // Check if target platform matches any of the configured Windows platforms
    private static boolean isWindowsPlatformMatch(String targetPlatform, Set<String> windowsPlatforms) {
        if (isBlank(targetPlatform) || windowsPlatforms.isEmpty()) {
            return false;
        }
        String normalizedTarget = targetPlatform.trim();
        for (String platform : windowsPlatforms) {
            if (normalizedTarget.equalsIgnoreCase(platform)) {
                return true;
            }
        }
        return false;
    }
    
    private static class FileProcessingResult {
        boolean exists;
        Path resolvedPath;
        boolean shouldSkip;
        boolean isRemoteSkip;
        boolean isRemoteError;
        String errorMessage;

        FileProcessingResult(boolean exists, Path resolvedPath) {
            this.exists = exists;
            this.resolvedPath = resolvedPath;
            this.shouldSkip = false;
        }

        static FileProcessingResult skip(boolean isRemoteSkip, boolean isRemoteError, String errorMessage) {
            FileProcessingResult result = new FileProcessingResult(false, null);
            result.shouldSkip = true;
            result.isRemoteSkip = isRemoteSkip;
            result.isRemoteError = isRemoteError;
            result.errorMessage = errorMessage;
            return result;
        }
    }

    private static class PathFixResult {
        String fixedPath;
        String errorReason;
        boolean wasFixed;

        PathFixResult(String fixedPath, String errorReason, boolean wasFixed) {
            this.fixedPath = fixedPath;
            this.errorReason = errorReason;
            this.wasFixed = wasFixed;
        }

        static PathFixResult noFixNeeded(String originalPath) {
            return new PathFixResult(originalPath, null, false);
        }

        static PathFixResult fixed(String fixedPath) {
            return new PathFixResult(fixedPath, null, true);
        }

        static PathFixResult invalid(String errorReason) {
            return new PathFixResult(null, errorReason, false);
        }
    }

    private static PathFixResult attemptPathFix(String rawPath, boolean isWindows) {
        // Case 1: Blank/empty path - cannot fix
        if (isBlank(rawPath)) {
            return PathFixResult.invalid("Empty file path");
        }

        String trimmedPath = rawPath.trim();
        String lowerPath = trimmedPath.toLowerCase();

        // Case 2: Specific invalid patterns that cannot be fixed
        if (lowerPath.equals("n/a") || lowerPath.equals("n\\a") || lowerPath.equals("na")) {
            return PathFixResult.invalid("Path marked as N/A");
        }

        // Case 3: Try to fix paths with trailing suffixes
        // Look for patterns like "C:\path\file.exe some_suffix_here"
        String fixedPath = tryRemoveTrailingSuffix(trimmedPath);
        if (fixedPath != null && !fixedPath.equals(trimmedPath)) {
            // Check if the fixed path is now valid
            String fixedInvalidReason = checkForInvalidPath(fixedPath, null, isWindows);
            if (fixedInvalidReason == null) {
                return PathFixResult.fixed(fixedPath);
            }
            // Fixed path is still invalid, continue with original validation
        }

        // No fix was possible or fix didn't help - check if original path is valid
        String originalInvalidReason = checkForInvalidPath(rawPath, null, isWindows);
        if (originalInvalidReason == null) {
            return PathFixResult.noFixNeeded(rawPath);
        } else {
            return PathFixResult.invalid(originalInvalidReason);
        }
    }

    private static String tryRemoveTrailingSuffix(String path) {
        // Separate file path and file name
        int lastSlash = Math.max(path.lastIndexOf('/'), path.lastIndexOf('\\'));

        if (lastSlash >= 0 && lastSlash < path.length() - 1) {
            // Path has directory component
            String dirPath = path.substring(0, lastSlash + 1); // Include the slash
            String fileName = path.substring(lastSlash + 1);

            // Only attempt to fix if the filename has an extension (contains a dot)
            // This prevents fixing valid directory names like "Program Files"
            if (fileName.contains(".")) {
                // Check if filename contains space
                int spaceIndex = fileName.indexOf(' '); // Find FIRST space in filename
                if (spaceIndex > 0) {
                    // Truncate filename at the FIRST space (including the space)
                    String cleanFileName = fileName.substring(0, spaceIndex);
                    return dirPath + cleanFileName;
                }
            }
        } else {
            // Path is just a filename (no directory)
            // Only attempt to fix if it looks like a filename with extension
            if (path.contains(".")) {
                int spaceIndex = path.indexOf(' '); // Find FIRST space in filename
                if (spaceIndex > 0) {
                    // Truncate at the FIRST space (including the space)
                    return path.substring(0, spaceIndex);
                }
            }
        }

        return null; // No fix possible (no space found in filename or no extension)
    }

    private static FileProcessingResult processFilePath(String filePath, String targetHost, boolean isRemoteWindows,
                                                       boolean isWindowsPlatform, int currentRow, int totalRows,
                                                       Set<String> inaccessibleHosts, int remoteUncTimeout) {
        try {
            Path resolved;
            boolean exists = false;
            String displayPath = filePath;

            if (isRemoteWindows) {
                // Try UNC path for remote Windows host
                String uncPath = convertToUncPath(targetHost.trim(), filePath);
                if (uncPath == null) {
                    return FileProcessingResult.skip(false, true, "Invalid path format for UNC conversion");
                }

                resolved = Paths.get(uncPath);
                displayPath = uncPath;

                // Update progress display for UNC path (this may take time)
                updateProgress(currentRow, totalRows, displayPath);

                try {
                    // Use timeout for UNC access to prevent infinite hangs
                    UncAccessResult result = checkUncPathWithTimeout(resolved, remoteUncTimeout);
                    if (result.timedOut) {
                        String normalizedHostname = normalizeHostname(targetHost);
                        inaccessibleHosts.add(normalizedHostname);
                        printVerboseMessage("Added " + targetHost.trim() + " to exclusion list - UNC access timeout");
                        return FileProcessingResult.skip(true, true, "UNC access timeout");
                    } else if (result.exception != null) {
                        throw result.exception;
                    }

                    exists = result.exists;
                    // For UNC paths, also check if we can determine file existence
                    if (!exists && !result.accessDenied) {
                        boolean notExists = Files.notExists(resolved);
                        if (!notExists) {
                            String normalizedHostname = normalizeHostname(targetHost);
                            inaccessibleHosts.add(normalizedHostname);
                            printVerboseMessage("Added " + targetHost.trim() + " to exclusion list - UNC access denied");
                            return FileProcessingResult.skip(true, true, "UNC access denied");
                        }
                    }
                } catch (Exception e) {
                    String normalizedHostname = normalizeHostname(targetHost);
                    inaccessibleHosts.add(normalizedHostname);
                    printVerboseMessage("Added " + targetHost.trim() + " to exclusion list - UNC access failed");
                    return FileProcessingResult.skip(true, true, "UNC access failed: " + e.getMessage());
                }
            } else {
                // Local host - use normal path resolution
                resolved = resolvePathCrossPlatform(filePath);
                exists = Files.exists(resolved);

                // Update progress display
                updateProgress(currentRow, totalRows, displayPath);
            }

            // Check for access permission issues using both exists() and notExists()
            if (!exists) {
                boolean notExists = Files.notExists(resolved);
                if (!notExists) {
                    // Neither exists() nor notExists() returned true - likely access denied
                    if (isRemoteWindows) {
                        String normalizedHostname = normalizeHostname(targetHost);
                        inaccessibleHosts.add(normalizedHostname);
                        printVerboseMessage("Added " + targetHost.trim() + " to exclusion list - UNC access denied");
                        return FileProcessingResult.skip(true, true, "UNC access denied");
                    } else {
                        return FileProcessingResult.skip(false, false, "Access denied - cannot determine file existence");
                    }
                }
            }

            // Check if it's a regular file (not a directory)
            if (exists && !Files.isRegularFile(resolved)) {
                exists = false; // Treat directories as not found
            }

            return new FileProcessingResult(exists, resolved);

        } catch (Exception e) {
            if (isRemoteWindows) {
                String normalizedHostname = normalizeHostname(targetHost);
                inaccessibleHosts.add(normalizedHostname);
                printVerboseMessage("Added " + targetHost.trim() + " to exclusion list - Exception during UNC access");
                return FileProcessingResult.skip(true, true, "Exception during UNC access: " + e.getMessage());
            } else {
                return FileProcessingResult.skip(false, false, "Exception during file access: " + e.getMessage());
            }
        }
    }


    private static String checkForInvalidPath(String rawPath, Path resolvedPath, boolean isWindows) {
        // Case 1: Blank/empty path
        if (isBlank(rawPath)) {
            return "Empty file path";
        }
        
        String trimmedPath = rawPath.trim();
        String lowerPath = trimmedPath.toLowerCase();
        
        // Case 2: Specific invalid path patterns
        if (lowerPath.equals("n/a") || 
            lowerPath.equals("n\\a") || 
            lowerPath.equals("na")) {
            return "Path marked as N/A";
        }
        
        // Case 3: Paths containing result.filename patterns
        if (lowerPath.contains(" result.filename=") || lowerPath.contains("result.filename=")) {
            return "Invalid path format containing result.filename pattern";
        }
        
        // Case 4: Paths starting with "C:\Program Files " (with trailing space)
        // Exception: "C:\Program Files (x86)" is valid
        if (lowerPath.startsWith("c:\\program files ") && !lowerPath.startsWith("c:\\program files (x86)")) {
            return "Invalid path format with trailing space after Program Files";
        }
        
        // Case 5: Windows .jar files with spaces in filename (invalid pattern)
        if (lowerPath.endsWith(".jar") && isWindows) {
            // Extract just the filename after the last slash
            String filename = trimmedPath;
            int lastSlash = Math.max(trimmedPath.lastIndexOf('/'), trimmedPath.lastIndexOf('\\'));
            if (lastSlash >= 0 && lastSlash < trimmedPath.length() - 1) {
                filename = trimmedPath.substring(lastSlash + 1);
            }
            // Only mark invalid if the filename itself (not the path) contains spaces
            if (filename.contains(" ")) {
                return "JAR filename contains spaces: " + filename;
            }
        }
        
        // Case 6: Directory instead of file (if path exists and is accessible)
        if (resolvedPath != null) {
            try {
                if (Files.exists(resolvedPath) && Files.isDirectory(resolvedPath)) {
                    return "Path is a directory, not a file";
                }
            } catch (Exception e) {
                // If we can't determine, don't mark as invalid due to directory check
                // Let normal file processing handle access issues
            }
        }
        
        // Extensible: Add more invalid path patterns here as needed
        // Future patterns could include:
        // - Invalid characters or formats
        // - Specific known bad path patterns
        // - Length restrictions
        // - etc.
        
        return null; // Path is valid
    }
    
    private static void setFileInvalid(Row row, int idxFileExists, int idxFileMod, int idxFileVersion, int idxScanError, String reason) {
        writeCell(row, idxFileExists, "X");
        writeCell(row, idxFileMod, "");
        writeCell(row, idxFileVersion, "");
        recordScanError(row, idxScanError, reason); // Record specific reason for invalid path
        // Note: FixedFilename is handled separately in the main processing logic
    }

    private static String generateUniqueID(String hostname, String cve, String originalFilePath, String fixedFilePath) {
        // Use fixed file path if it exists and is not empty, otherwise use original
        String actualFilePath = (!isBlank(fixedFilePath)) ? fixedFilePath.trim() : originalFilePath.trim();

        // Normalize hostname and CVE for consistent comparison
        String normalizedHostname = isBlank(hostname) ? "" : hostname.trim().toLowerCase();
        String normalizedCVE = isBlank(cve) ? "" : cve.trim().toUpperCase();
        String normalizedFilePath = isBlank(actualFilePath) ? "" : actualFilePath.trim().toLowerCase();

        // Concatenate with separator to create unique ID
        return normalizedHostname + "|" + normalizedCVE + "|" + normalizedFilePath;
    }

    // CVE Data Structure
    private static class CVEData {
        String cveId;
        String description;
        List<String> references;
        List<String> affectedSoftware;
        boolean hasWeblogic;
        List<String> oracleAdvisories;

        CVEData(String cveId) {
            this.cveId = cveId;
            this.references = new ArrayList<String>();
            this.affectedSoftware = new ArrayList<String>();
            this.oracleAdvisories = new ArrayList<String>();
            this.hasWeblogic = false;
        }
    }

    // Create CVE information sheet with data from NIST NVD
    private static void createCVESheet(Workbook wb, Sheet sourceSheet, String cveColumnName) {
        try {
            // Get unique CVE IDs from the source sheet
            Set<String> uniqueCVEs = extractUniqueCVEs(sourceSheet, cveColumnName);
            logMessage("Found " + uniqueCVEs.size() + " unique CVE IDs to process");

            // Create or get CVEs sheet
            Sheet cveSheet = wb.getSheet("CVEs");
            if (cveSheet != null) {
                wb.removeSheetAt(wb.getSheetIndex(cveSheet));
            }
            cveSheet = wb.createSheet("CVEs");

            // Create header row
            Row headerRow = cveSheet.createRow(0);
            headerRow.createCell(0).setCellValue("CVE ID");
            headerRow.createCell(1).setCellValue("Description");
            headerRow.createCell(2).setCellValue("References");
            headerRow.createCell(3).setCellValue("Affected Software");
            headerRow.createCell(4).setCellValue("Weblogic");
            headerRow.createCell(5).setCellValue("Weblogic Sec Advisories");

            // Fetch CVE data and populate sheet
            int rowNum = 1;
            for (String cveId : uniqueCVEs) {
                if (isBlank(cveId)) continue;

                logMessage("Fetching data for " + cveId + "...");
                CVEData cveData = fetchCVEData(cveId);

                if (cveData != null) {
                    Row dataRow = cveSheet.createRow(rowNum++);
                    dataRow.createCell(0).setCellValue(cveData.cveId);
                    dataRow.createCell(1).setCellValue(cveData.description != null ? cveData.description : "");
                    dataRow.createCell(2).setCellValue(String.join("; ", cveData.references));
                    dataRow.createCell(3).setCellValue(String.join("; ", cveData.affectedSoftware));
                    dataRow.createCell(4).setCellValue(cveData.hasWeblogic ? "Y" : "N");
                    dataRow.createCell(5).setCellValue(String.join("; ", cveData.oracleAdvisories));
                } else {
                    logError("Failed to fetch data for " + cveId);
                    Row dataRow = cveSheet.createRow(rowNum++);
                    dataRow.createCell(0).setCellValue(cveId);
                    dataRow.createCell(1).setCellValue("ERROR: Could not fetch CVE data");
                }

                // Add delay to avoid overwhelming the NIST API
                try {
                    Thread.sleep(2000); // 2 second delay between requests to avoid rate limiting
                } catch (InterruptedException e) {
                    Thread.currentThread().interrupt();
                    break;
                }
            }

            // Auto-size columns
            for (int i = 0; i < 6; i++) {
                cveSheet.autoSizeColumn(i);
            }

            logMessage("CVE sheet created with " + (rowNum - 1) + " CVE entries");

        } catch (Exception e) {
            logError("Error creating CVE sheet: " + e.getMessage());
        }
    }

    // Extract unique CVE IDs from the source sheet
    private static Set<String> extractUniqueCVEs(Sheet sheet, String cveColumnName) {
        Set<String> uniqueCVEs = new HashSet<String>();

        // Find CVE column index
        Row headerRow = sheet.getRow(0);
        if (headerRow == null) return uniqueCVEs;

        int cveColumnIndex = -1;
        for (int i = 0; i < headerRow.getLastCellNum(); i++) {
            Cell cell = headerRow.getCell(i);
            if (cell != null && cveColumnName.equals(cell.getStringCellValue())) {
                cveColumnIndex = i;
                break;
            }
        }

        if (cveColumnIndex == -1) {
            logError("CVE column '" + cveColumnName + "' not found");
            return uniqueCVEs;
        }

        // Extract CVE IDs from all rows
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                Cell cveCell = row.getCell(cveColumnIndex);
                if (cveCell != null) {
                    String cveId = cveCell.getStringCellValue();
                    if (!isBlank(cveId) && cveId.trim().startsWith("CVE-")) {
                        uniqueCVEs.add(cveId.trim().toUpperCase());
                    }
                }
            }
        }

        return uniqueCVEs;
    }

    // Fetch CVE data from NIST NVD
    private static CVEData fetchCVEData(String cveId) {
        try {
            // NIST NVD REST API endpoint
            String apiUrl = "https://services.nvd.nist.gov/rest/json/cves/2.0?cveId=" + cveId;

            URL url = new URL(apiUrl);
            HttpURLConnection connection = (HttpURLConnection) url.openConnection();
            connection.setRequestMethod("GET");
            connection.setRequestProperty("Accept", "application/json");
            connection.setConnectTimeout(10000); // 10 seconds
            connection.setReadTimeout(15000); // 15 seconds

            int responseCode = connection.getResponseCode();
            if (responseCode != 200) {
                logError("Failed to fetch CVE data for " + cveId + ". HTTP response code: " + responseCode);
                return null;
            }

            // Read response
            StringBuilder response = new StringBuilder();
            try (BufferedReader reader = new BufferedReader(new InputStreamReader(connection.getInputStream(), StandardCharsets.UTF_8))) {
                String line;
                while ((line = reader.readLine()) != null) {
                    response.append(line);
                }
            }

            // Parse JSON response (basic parsing without external JSON library)
            return parseCVEResponse(cveId, response.toString());

        } catch (Exception e) {
            logError("Error fetching CVE data for " + cveId + ": " + e.getMessage());
            return null;
        }
    }

    // Parse CVE JSON response (basic JSON parsing for Java 8 compatibility)
    private static CVEData parseCVEResponse(String cveId, String jsonResponse) {
        CVEData cveData = new CVEData(cveId);

        try {
            // Extract description
            String descPattern = "\"value\"\\s*:\\s*\"([^\"]+)\"";
            Pattern pattern = Pattern.compile(descPattern);
            Matcher matcher = pattern.matcher(jsonResponse);
            if (matcher.find()) {
                cveData.description = matcher.group(1).replaceAll("\\\\\"", "\"").replaceAll("\\\\n", " ");
            }

            // Extract references
            String refPattern = "\"url\"\\s*:\\s*\"([^\"]+)\"";
            pattern = Pattern.compile(refPattern);
            matcher = pattern.matcher(jsonResponse);
            while (matcher.find()) {
                String url = matcher.group(1);
                cveData.references.add(url);

                // Check for Oracle advisories (handle escaped URLs from JSON)
                String unescapedUrl = url.replace("\\/", "/");
                if (unescapedUrl.startsWith("https://www.oracle.com/") ||
                    unescapedUrl.startsWith("http://www.oracle.com/")) {
                    cveData.oracleAdvisories.add(unescapedUrl);
                }
            }

            // Extract affected software configurations (CPE)
            // NIST NVD API v2.0 uses "criteria" instead of "cpe23Uri"
            String cpePattern = "\"criteria\"\\s*:\\s*\"([^\"]+)\"";
            pattern = Pattern.compile(cpePattern);
            matcher = pattern.matcher(jsonResponse);
            while (matcher.find()) {
                String cpe = matcher.group(1);
                cveData.affectedSoftware.add(cpe);

                // Check for Weblogic
                if (cpe.contains("weblogic_server")) {
                    cveData.hasWeblogic = true;
                }
            }

        } catch (Exception e) {
            logError("Error parsing CVE response for " + cveId + ": " + e.getMessage());
        }

        return cveData;
    }
}
