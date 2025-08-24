package app;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.net.InetAddress;
import java.net.UnknownHostException;
import java.nio.charset.StandardCharsets;
import java.nio.file.*;
import java.time.Instant;
import java.time.ZoneId;
import java.time.ZonedDateTime;
import java.time.format.DateTimeFormatter;
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

            String localHost = getLocalHostName();
            System.out.println("Local hostname: " + localHost);

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

                for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                    Row row = sheet.getRow(r);
                    if (row == null) continue;

                    String targetHost = getStringCell(row, idxHostName);
                    if (isBlank(targetHost) || !targetHost.trim().equalsIgnoreCase(localHost)) {
                        skippedHost++;
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

                    Path resolved = resolvePathCrossPlatform(rawPath);
                    boolean exists = Files.exists(resolved);
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
                        writeCell(row, idxScanError, "File does not exist");
                    }

                    processed++;
                }

                // Save back to the same file
                try (FileOutputStream fos = new FileOutputStream(excelFile)) {
                    wb.write(fos);
                }
                System.out.println("Done. Rows processed: " + processed + ", skipped (hostname mismatch): " + skippedHost);
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
                    return new JarResult(null, "No Implementation-Version in MANIFEST.MF");
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
}
