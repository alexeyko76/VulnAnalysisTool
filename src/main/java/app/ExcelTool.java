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
import java.util.concurrent.atomic.AtomicBoolean;
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
    private static final String KEY_PROGRESS_DISPLAY = "progress.display";

    // Additional columns to ensure exist
    private static final String COL_FILE_EXISTS = "FileExists";
    private static final String COL_FILE_MOD_DATE = "FileModificationDate";
    private static final String COL_JAR_VERSION = "JarVersion";
    private static final String COL_SCAN_ERROR = "ScanError";

    private static final DateTimeFormatter TS_FMT = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");

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
            String windowsPlatformValue = require(cfg, KEY_PLATFORM_WINDOWS);
            boolean remoteUncEnabled = getBoolean(cfg, KEY_REMOTE_UNC_ENABLED, true);
            int remoteUncTimeout = getInteger(cfg, KEY_REMOTE_UNC_TIMEOUT, 30);
            String progressDisplay = getString(cfg, KEY_PROGRESS_DISPLAY, "bar");

            String localHost = getLocalHostName();
            System.out.println("Local hostname: " + localHost);
            System.out.println("Windows platform value: " + windowsPlatformValue);
            System.out.println("Remote UNC access enabled: " + remoteUncEnabled);
            if (remoteUncEnabled) {
                System.out.println("UNC access timeout: " + remoteUncTimeout + " seconds");
            }
            System.out.println("Progress display mode: " + progressDisplay);

            File excelFile = new File(excelPath);
            if (!excelFile.exists()) {
                throw new IllegalArgumentException("Excel file does not exist: " + excelFile.getAbsolutePath());
            }

            try (FileInputStream fis = new FileInputStream(excelFile);
                 Workbook wb = WorkbookFactory.create(fis)) {

                // Find the specified sheet by name
                Sheet sheet = wb.getSheet(sheetName);
                if (sheet == null) {
                    System.err.println("ERROR: Sheet '" + sheetName + "' does not exist in the Excel file.");
                    System.err.println("Available sheets:");
                    for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                        System.err.println("  - " + wb.getSheetName(i));
                    }
                    System.err.println("The tool will exit without processing.");
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
                    System.err.println("ERROR: Required columns missing: " + missingRequired);
                    System.err.println("The tool will exit without saving changes.");
                    System.exit(2);
                }

                // Ensure additional columns exist (create if missing)
                ensureColumn(header, colIndex, COL_FILE_EXISTS);
                ensureColumn(header, colIndex, COL_FILE_MOD_DATE);
                ensureColumn(header, colIndex, COL_JAR_VERSION);
                ensureColumn(header, colIndex, COL_SCAN_ERROR);

                int idxPlatform = colIndex.get(colPlatform);
                int idxFilePath = colIndex.get(colFilePath);
                int idxHostName = colIndex.get(colHostName);
                int idxFileExists = colIndex.get(COL_FILE_EXISTS);
                int idxFileMod = colIndex.get(COL_FILE_MOD_DATE);
                int idxJarVersion = colIndex.get(COL_JAR_VERSION);
                int idxScanError = colIndex.get(COL_SCAN_ERROR);

                int processed = 0;
                int skippedHost = 0;
                int skippedRemote = 0;
                
                // Exclusion list for hosts that cannot be accessed via UNC from current machine
                Set<String> inaccessibleHosts = new HashSet<String>();
                
                // Progress tracking
                int totalRows = sheet.getLastRowNum();
                boolean useProgressBar = "bar".equalsIgnoreCase(progressDisplay);
                
                if (useProgressBar) {
                    System.out.println("Processing " + totalRows + " rows...");
                    System.out.println(); // Empty line for progress display
                } else {
                    System.out.println("Processing " + totalRows + " rows (verbose mode):");
                }

                for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                    Row row = sheet.getRow(r);
                    if (row == null) continue;

                    String targetHost = getStringCell(row, idxHostName);
                    String targetPlatform = getStringCell(row, idxPlatform);
                    boolean isLocalHost = !isBlank(targetHost) && targetHost.trim().equalsIgnoreCase(localHost);
                    boolean isWindowsPlatform = !isBlank(targetPlatform) && targetPlatform.trim().equalsIgnoreCase(windowsPlatformValue);
                    boolean isRemoteWindows = remoteUncEnabled && isWindowsPlatform && !isLocalHost && !isBlank(targetHost);
                    
                    // Skip if not local host and (UNC disabled or not a Windows platform for remote access)
                    if (!isLocalHost && !isRemoteWindows) {
                        skippedHost++;
                        continue;
                    }
                    
                    // Skip if remote host is in exclusion list
                    if (isRemoteWindows && inaccessibleHosts.contains(targetHost.trim().toLowerCase())) {
                        skippedRemote++;
                        continue;
                    }

                    String rawPath = getStringCell(row, idxFilePath);
                    if (isBlank(rawPath)) {
                        writeCell(row, idxFileExists, "N");
                        writeCell(row, idxFileMod, "");
                        writeCell(row, idxJarVersion, "");
                        writeCell(row, idxScanError, "Empty file path");
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
                            updateProgress(r, totalRows, displayPath, useProgressBar);
                            
                            try {
                                // Use timeout for UNC access to prevent infinite hangs
                                UncAccessResult result = checkUncPathWithTimeout(resolved, remoteUncTimeout);
                                if (result.timedOut) {
                                    // Timeout occurred - add host to exclusion list
                                    inaccessibleHosts.add(targetHost.trim().toLowerCase());
                                    writeCell(row, idxScanError, "UNC access timeout - host may be unreachable or slow");
                                    skippedRemote++;
                                    printVerboseMessage("Added " + targetHost.trim() + " to exclusion list - UNC access timeout", useProgressBar);
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
                                        inaccessibleHosts.add(targetHost.trim().toLowerCase());
                                        writeCell(row, idxScanError, "UNC access denied - cannot determine file existence");
                                        skippedRemote++;
                                        printVerboseMessage("Added " + targetHost.trim() + " to exclusion list - access denied detected", useProgressBar);
                                        continue;
                                    }
                                }
                            } catch (Exception e) {
                                // UNC access failed - add host to exclusion list
                                inaccessibleHosts.add(targetHost.trim().toLowerCase());
                                writeCell(row, idxScanError, "Cannot access remote host via UNC: " + e.getMessage());
                                skippedRemote++;
                                printVerboseMessage("Added " + targetHost.trim() + " to exclusion list - UNC access failed", useProgressBar);
                                continue;
                            }
                        } else {
                            // Invalid path format for UNC conversion
                            resolved = resolvePathCrossPlatform(rawPath);
                            writeCell(row, idxFileExists, "N");
                            writeCell(row, idxFileMod, "");
                            writeCell(row, idxJarVersion, "");
                            writeCell(row, idxScanError, "Invalid path format for UNC conversion");
                            processed++;
                            continue;
                        }
                    } else {
                        // Local host - use normal path resolution
                        resolved = resolvePathCrossPlatform(rawPath);
                        exists = Files.exists(resolved);
                    }
                    
                    // Update progress display
                    updateProgress(r, totalRows, displayPath, useProgressBar);
                    
                    // Check for access permission issues using both exists() and notExists()
                    if (!exists) {
                        boolean notExists = Files.notExists(resolved);
                        if (!notExists) {
                            // Neither exists() nor notExists() returned true - likely access denied
                            // Add remote host to exclusion list to avoid repeated attempts
                            if (isRemoteWindows) {
                                inaccessibleHosts.add(targetHost.trim().toLowerCase());
                                printVerboseMessage("Added " + targetHost.trim() + " to exclusion list - access denied detected", useProgressBar);
                                writeCell(row, idxScanError, "UNC access denied - cannot determine file existence");
                            } else {
                                writeCell(row, idxScanError, "Access denied - cannot determine file existence");
                                writeCell(row, idxFileExists, "N");
                                writeCell(row, idxFileMod, "");
                                writeCell(row, idxJarVersion, "");
                            }
                            if (isRemoteWindows) {
                                skippedRemote++;
                            } else {
                                processed++;
                            }
                            continue;
                        }
                    }
                    
                    // Check if it's a regular file (not a directory) - if not, treat as not found
                    String notRegularFileError = null;
                    if (exists && !Files.isRegularFile(resolved)) {
                        exists = false; // Treat non-regular files as not found
                        notRegularFileError = "Path exists but is not a regular file (directory or special file)";
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
                            System.err.println("WARN: Could not read last modified for: " + resolved + " -> " + e.getMessage());
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
                        // Use specific error message for non-regular files, otherwise default message
                        writeCell(row, idxScanError, notRegularFileError != null ? notRegularFileError : "File does not exist");
                    }

                    processed++;
                }

                // Save back to the same file
                try (FileOutputStream fos = new FileOutputStream(excelFile)) {
                    wb.write(fos);
                }
                
                // Clear progress bar and print final results
                if (useProgressBar) {
                    // Complete the progress bar with final newline
                    System.out.println();
                }
                System.out.println("Done. Rows processed: " + processed + ", skipped (hostname mismatch): " + skippedHost + ", skipped (remote inaccessible): " + skippedRemote);
                if (!inaccessibleHosts.isEmpty()) {
                    System.out.println("Inaccessible hosts identified during this run: " + inaccessibleHosts);
                }
            }
        } catch (java.io.IOException e) {
            System.err.println("ERROR: IO failure: " + e.getMessage());
            exit = 3;
        } catch (IllegalArgumentException e) {
            System.err.println("ERROR: " + e.getMessage());
            exit = 4;
        } catch (Exception e) {
            System.err.println("ERROR: Unexpected failure: " + e.getMessage());
            e.printStackTrace(System.err);
            exit = 5;
        }
        if (exit != 0) {
            System.exit(exit);
        }
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
            System.err.println("WARN: Failed to read manifest from jar: " + jarFile + " -> " + e.getMessage());
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
            System.err.println("WARN: Failed to manually parse manifest: " + e.getMessage());
        }
        return null;
    }
    
    private static void updateProgress(int currentRow, int totalRows, String currentFile, boolean useProgressBar) {
        if (!useProgressBar) {
            // Verbose mode - simple row-by-row logging with timestamp
            System.out.println(getCurrentTimestamp() + " Row " + currentRow + "/" + totalRows + ": " + currentFile);
            return;
        }
        
        // Progress bar mode (original logic)
        updateProgressBar(currentRow, totalRows, currentFile);
    }
    
    private static void printVerboseMessage(String message, boolean useProgressBar) {
        if (useProgressBar) {
            // In progress bar mode, suppress verbose messages
            return;
        } else {
            // In verbose mode, indent additional messages for clarity with timestamp
            System.out.println(getCurrentTimestamp() + "   -> " + message);
        }
    }
    
    private static String getCurrentTimestamp() {
        return ZonedDateTime.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss"));
    }
    
    private static void updateProgressBar(int currentRow, int totalRows, String currentFile) {
        // Calculate progress percentage
        int percentage = (int) ((currentRow * 100.0) / totalRows);
        
        // Create progress bar (30 characters wide for better fit)
        int progressChars = (percentage * 30) / 100; // 30 chars = 100%
        StringBuilder progressBar = new StringBuilder();
        progressBar.append("[");
        for (int i = 0; i < 30; i++) {
            if (i < progressChars) {
                progressBar.append("=");
            } else if (i == progressChars && percentage < 100) {
                progressBar.append(">");
            } else {
                progressBar.append(" ");
            }
        }
        progressBar.append("] ");
        progressBar.append(String.format("%3d%%", percentage));
        progressBar.append(String.format(" (%d/%d)", currentRow, totalRows));
        
        // Truncate file path if too long for display
        String displayFile = currentFile;
        if (displayFile.length() > 60) {
            displayFile = "..." + displayFile.substring(displayFile.length() - 57);
        }
        
        // Simple approach: print progress line with carriage return for overwrite
        String progressLine = "Progress: " + progressBar.toString() + " | File: " + displayFile;
        
        // Add padding spaces at the end to clear previous longer lines
        progressLine = progressLine + "                    "; // Fixed padding
        
        System.out.print("\r" + progressLine);
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
}
