package app;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.CellType;

import java.io.*;
import java.net.InetAddress;
import java.net.UnknownHostException;
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

    // Config keys
    private static final String KEY_EXCEL_PATH = "excel.path";
    private static final String KEY_SHEET_NAME = "sheet.name";
    private static final String KEY_COL_PLATFORM = "column.PlatformName";
    private static final String KEY_COL_FILEPATH = "column.FilePath";
    private static final String KEY_COL_HOSTNAME = "column.HostName";
    private static final String KEY_PLATFORM_WINDOWS = "platform.windows";
    private static final String KEY_REMOTE_UNC_ENABLED = "remote.unc.enabled";
    private static final String KEY_REMOTE_UNC_TIMEOUT = "remote.unc.timeout";
    private static final String KEY_LOG_FILENAME = "log.filename";

    // Additional columns to ensure exist
    private static final String COL_FILE_EXISTS = "FileExists";
    private static final String COL_FILE_MOD_DATE = "FileModificationDate";
    private static final String COL_JAR_VERSION = "JarVersion";
    private static final String COL_SCAN_ERROR = "ScanError";
    private static final String COL_REMOTE_SCAN_ERROR = "RemoteScanError";
    private static final String COL_SCAN_DATE = "ScanDate";

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
            String windowsPlatformValues = require(cfg, KEY_PLATFORM_WINDOWS);
            Set<String> windowsPlatforms = parseWindowsPlatforms(windowsPlatformValues);
            boolean remoteUncEnabled = getBoolean(cfg, KEY_REMOTE_UNC_ENABLED, true);
            int remoteUncTimeout = getInteger(cfg, KEY_REMOTE_UNC_TIMEOUT, 30);
            String logFilename = getString(cfg, KEY_LOG_FILENAME, "");
            
            // Get hostname first for log file prefixing
            String localHost = getLocalHostName();
            
            // Initialize log file with hostname prefix if specified
            String prefixedLogFilename = createHostnamePrefixedLogFilename(logFilename, localHost);
            initializeLogFile(prefixedLogFilename);
            
            logMessage("Local hostname: " + localHost);
            logMessage("Windows platform values: " + windowsPlatforms);
            logMessage("Remote UNC access enabled: " + remoteUncEnabled);
            if (remoteUncEnabled) {
                logMessage("UNC access timeout: " + remoteUncTimeout + " seconds");
            }
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

                if (!missingRequired.isEmpty()) {
                    logError("Required columns missing: " + missingRequired);
                    logError("The tool will exit without saving changes.");
                    System.exit(2);
                }

                // Ensure additional columns exist (create if missing)
                ensureColumn(header, colIndex, COL_FILE_EXISTS);
                ensureColumn(header, colIndex, COL_FILE_MOD_DATE);
                ensureColumn(header, colIndex, COL_JAR_VERSION);
                ensureColumn(header, colIndex, COL_SCAN_ERROR);
                ensureColumn(header, colIndex, COL_REMOTE_SCAN_ERROR);
                ensureColumn(header, colIndex, COL_SCAN_DATE);

                int idxPlatform = colIndex.get(colPlatform);
                int idxFilePath = colIndex.get(colFilePath);
                int idxHostName = colIndex.get(colHostName);
                int idxFileExists = colIndex.get(COL_FILE_EXISTS);
                int idxFileMod = colIndex.get(COL_FILE_MOD_DATE);
                int idxJarVersion = colIndex.get(COL_JAR_VERSION);
                int idxScanError = colIndex.get(COL_SCAN_ERROR);
                int idxRemoteScanError = colIndex.get(COL_REMOTE_SCAN_ERROR);
                int idxScanDate = colIndex.get(COL_SCAN_DATE);

                int processed = 0;
                int skippedHost = 0;
                int skippedRemote = 0;
                
                // Exclusion list for hosts that cannot be accessed via UNC from current machine
                Set<String> inaccessibleHosts = new HashSet<String>();
                
                // Single timestamp for entire scanning session
                String sessionScanTimestamp = getCurrentTimestamp();
                
                // Progress tracking
                int totalRows = sheet.getLastRowNum();
                
                logMessage("Processing " + totalRows + " rows:");

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
                    writeCell(row, idxScanDate, sessionScanTimestamp);
                    
                    // Skip if remote host is in exclusion list (but update RemoteScanError for this row)
                    if (isRemoteWindows && inaccessibleHosts.contains(normalizeHostname(targetHost))) {
                        recordRemoteScanError(row, idxRemoteScanError, "Host previously identified as inaccessible via UNC");
                        skippedRemote++;
                        continue;
                    }

                    String rawPath = getStringCell(row, idxFilePath);
                    
                    // Check for invalid path patterns first (before path resolution)
                    if (checkForInvalidPath(rawPath, null, isWindowsPlatform)) {
                        setFileInvalid(row, idxFileExists, idxFileMod, idxJarVersion, idxScanError);
                        processed++;
                        continue;
                    }

                    Path resolved;
                    boolean exists = false;
                    String displayPath = rawPath; // Default display path
                    
                    if (isRemoteWindows) {
                        // Try UNC path for remote Windows host
                        String uncPath = convertToUncPath(targetHost.trim(), rawPath);
                        if (uncPath != null) {
                            resolved = Paths.get(uncPath);
                            displayPath = uncPath; // Show UNC path in progress
                            
                            // Update progress display for UNC path (this may take time)
                            updateProgress(r, totalRows, displayPath);
                            
                            try {
                                // Use timeout for UNC access to prevent infinite hangs
                                UncAccessResult result = checkUncPathWithTimeout(resolved, remoteUncTimeout);
                                if (result.timedOut) {
                                    // Timeout occurred - add host to exclusion list
                                    addHostToExclusionList(inaccessibleHosts, targetHost, row, idxRemoteScanError, 
                                                         "UNC access timeout - host may be unreachable or slow");
                                    skippedRemote++;
                                    continue;
                                } else if (result.exception != null) {
                                    // Exception during access
                                    throw result.exception;
                                }
                                
                                exists = result.exists;
                                // For UNC paths, also check if we can determine file existence
                                if (!exists && !result.accessDenied) {
                                    boolean notExists = Files.notExists(resolved);
                                    if (!notExists) {
                                        // Neither exists() nor notExists() returned true for UNC path - access denied
                                        addHostToExclusionList(inaccessibleHosts, targetHost, row, idxRemoteScanError, 
                                                             "UNC access denied - cannot determine file existence");
                                        skippedRemote++;
                                        continue;
                                    }
                                }
                            } catch (Exception e) {
                                // UNC access failed - add host to exclusion list
                                addHostToExclusionList(inaccessibleHosts, targetHost, row, idxRemoteScanError, 
                                                     "Cannot access remote host via UNC: " + e.getMessage());
                                skippedRemote++;
                                continue;
                            }
                        } else {
                            // Invalid path format for UNC conversion
                            resolved = resolvePathCrossPlatform(rawPath);
                            setFileNotFound(row, idxFileExists, idxFileMod, idxJarVersion);
                            recordRemoteScanError(row, idxRemoteScanError, "Invalid path format for UNC conversion");
                            processed++;
                            continue;
                        }
                    } else {
                        // Local host - use normal path resolution
                        resolved = resolvePathCrossPlatform(rawPath);
                        exists = Files.exists(resolved);
                    }
                    
                    // Update progress display
                    updateProgress(r, totalRows, displayPath);
                    
                    // Check for access permission issues using both exists() and notExists()
                    if (!exists) {
                        boolean notExists = Files.notExists(resolved);
                        if (!notExists) {
                            // Neither exists() nor notExists() returned true - likely access denied
                            // Add remote host to exclusion list to avoid repeated attempts
                            if (isRemoteWindows) {
                                addHostToExclusionList(inaccessibleHosts, targetHost, row, idxRemoteScanError, 
                                                     "UNC access denied - cannot determine file existence");
                            } else {
                                recordScanError(row, idxScanError, "Access denied - cannot determine file existence");
                                setFileNotFound(row, idxFileExists, idxFileMod, idxJarVersion);
                            }
                            if (isRemoteWindows) {
                                skippedRemote++;
                            } else {
                                processed++;
                            }
                            continue;
                        }
                    }
                    
                    // Check if it's a regular file (not a directory) - if not, mark as invalid
                    if (exists && !Files.isRegularFile(resolved)) {
                        // Re-check for invalid path now that we have resolved path (for directory detection)
                        if (checkForInvalidPath(rawPath, resolved, isWindowsPlatform)) {
                            setFileInvalid(row, idxFileExists, idxFileMod, idxJarVersion, idxScanError);
                            processed++;
                            continue;
                        } else {
                            // Fallback: treat as not found with error message
                            exists = false;
                        }
                    }
                    
                    writeCell(row, idxFileExists, exists ? "Y" : "N");

                    if (exists) {
                        StringBuilder scanErrors = new StringBuilder();
                        
                        try {
                            Instant lm = Files.getLastModifiedTime(resolved).toInstant();
                            ZonedDateTime zdt = ZonedDateTime.ofInstant(lm, ZoneId.systemDefault());
                            writeCell(row, idxFileMod, TS_FMT.format(zdt.toLocalDateTime()));
                        } catch (IOException e) {
                            writeCell(row, idxFileMod, "");
                            scanErrors.append("Cannot read modification date: ").append(e.getMessage());
                            logError("WARN: Could not read last modified for: " + resolved + " -> " + e.getMessage());
                        }
                        // Jar handling
                        if (resolved.getFileName() != null && resolved.getFileName().toString().toLowerCase(Locale.ENGLISH).endsWith(".jar")) {
                            JarResult result = extractImplementationVersionWithError(resolved.toFile());
                            writeCell(row, idxJarVersion, result.version != null ? result.version : "");
                            if (result.error != null) {
                                if (scanErrors.length() > 0) scanErrors.append("; ");
                                scanErrors.append("JAR processing error: ").append(result.error);
                            }
                        } else {
                            writeCell(row, idxJarVersion, "");
                        }
                        
                        // Write scan errors or clear if none
                        writeCell(row, idxScanError, scanErrors.length() > 0 ? scanErrors.toString() : "");
                    } else {
                        writeCell(row, idxFileMod, "");
                        writeCell(row, idxJarVersion, "");
                        // For local scans, clear ScanError when file simply doesn't exist (successful scan)
                        // For remote scans, preserve any existing ScanError and use RemoteScanError for UNC issues
                        if (isRemoteWindows) {
                            // Don't clear ScanError for remote scans - it may contain local file processing errors
                        } else {
                            // For local scans, clear ScanError for successful determination that file doesn't exist
                            recordScanError(row, idxScanError, ""); // Clear error for successful scan
                        }
                    }
                    
                    // Clear error columns for successful scans
                    clearScanErrors(row, idxScanError, idxRemoteScanError, isRemoteWindows);

                    processed++;
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

    
    private static class JarResult {
        String version;
        String error;
        
        JarResult(String version, String error) {
            this.version = version;
            this.error = error;
        }
    }
    
    private static JarResult extractImplementationVersionWithError(File jarFile) {
        if (jarFile == null || !jarFile.exists()) {
            return new JarResult(null, "JAR file does not exist");
        }
        
        ZipFile zip = null;
        try {
            zip = new ZipFile(jarFile);
            ZipEntry entry = zip.getEntry("META-INF/MANIFEST.MF");
            if (entry == null) {
                return new JarResult(null, "No MANIFEST.MF found in JAR");
            }
            
            try (InputStream is = zip.getInputStream(entry)) {
                Manifest mf = new Manifest(is);
                String v = mf.getMainAttributes().getValue("Implementation-Version");
                if (v != null) {
                    return new JarResult(v.trim(), null);
                } else {
                    // Fallback: manually parse MANIFEST.MF for Implementation-Version
                    String manualVersion = parseManifestManually(zip, entry);
                    if (manualVersion != null) {
                        return new JarResult(manualVersion, null);
                    } else {
                        return new JarResult(null, "No Implementation-Version in MANIFEST.MF");
                    }
                }
            }
        } catch (IOException e) {
            logError("WARN: Failed to read manifest from jar: " + jarFile + " -> " + e.getMessage());
            return new JarResult(null, e.getMessage());
        } finally {
            if (zip != null) {
                try { zip.close(); } catch (IOException ignored) {}
            }
        }
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
    
    private static void setFileNotFound(Row row, int idxFileExists, int idxFileMod, int idxJarVersion) {
        writeCell(row, idxFileExists, "N");
        writeCell(row, idxFileMod, "");
        writeCell(row, idxJarVersion, "");
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
    
    private static boolean checkForInvalidPath(String rawPath, Path resolvedPath, boolean isWindows) {
        // Case 1: Blank/empty path
        if (isBlank(rawPath)) {
            return true;
        }
        
        String trimmedPath = rawPath.trim();
        String lowerPath = trimmedPath.toLowerCase();
        
        // Case 2: Specific invalid path patterns
        if (lowerPath.equals("n/a") || 
            lowerPath.equals("n\\a") || 
            lowerPath.equals("na")) {
            return true;
        }
        
        // Case 3: Paths containing result.filename patterns
        if (lowerPath.contains(" result.filename=") || lowerPath.contains("result.filename=")) {
            return true;
        }
        
        // Case 4: Paths starting with "C:\Program Files " (with trailing space)
        if (lowerPath.startsWith("c:\\program files ")) {
            return true;
        }
        
        // Case 5: Windows .jar files with spaces in filename (invalid pattern)
        if (lowerPath.endsWith(".jar") && trimmedPath.contains(" ") && isWindows) {
            return true;
        }
        
        // Case 6: Directory instead of file (if path exists and is accessible)
        if (resolvedPath != null) {
            try {
                if (Files.exists(resolvedPath) && Files.isDirectory(resolvedPath)) {
                    return true;
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
        
        return false;
    }
    
    private static void setFileInvalid(Row row, int idxFileExists, int idxFileMod, int idxJarVersion, int idxScanError) {
        writeCell(row, idxFileExists, "X");
        writeCell(row, idxFileMod, "");
        writeCell(row, idxJarVersion, "");
        recordScanError(row, idxScanError, ""); // Clear ScanError - invalid path indicated by FileExists="X"
    }
}
