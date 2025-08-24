# Dependencies

This directory contains all required JAR dependencies for the Excel vulnerability analysis tool.

## Required Dependencies for Apache POI 5.4.1

The following JARs are required for the tool to function properly:

### Apache POI (Excel processing)
- `poi-5.4.1.jar` - Core POI library
- `poi-ooxml-5.4.1.jar` - OOXML format support (.xlsx)
- `poi-ooxml-lite-5.4.1.jar` - Lightweight OOXML support
- `xmlbeans-5.3.0.jar` - XML processing for OOXML

### Apache Commons (Utilities)
- `commons-collections4-4.5.0.jar` - Collection utilities
- `commons-compress-1.28.0.jar` - Compression/archive support
- `commons-io-2.20.0.jar` - I/O utilities
- `commons-lang3-3.12.0.jar` - Language utilities (required for POI 5.4.1)

### Apache Log4j (Logging)
- `log4j-api-2.17.2.jar` - Logging API (required by POI)
- `log4j-core-2.17.2.jar` - Logging implementation

## Version Compatibility

- **Java Runtime**: Java 8+ compatible
- **Apache POI**: Version 5.4.1 (security patched)
- **Log4j**: Version 2.17.2 (security patched, compatible with POI 5.4.1)

## Installation Methods

### Method 1: Download Individual JARs
Download each JAR from Maven Central and place in this directory.

### Method 2: Official POI Distribution
1. Download the official POI binary distribution ZIP for 5.4.1
2. Copy all JARs from the `lib/` directory to `deps/`
3. Add the additional Log4j and Commons Lang3 dependencies listed above

## Notes

- All dependencies are required for the uber JAR build process
- Missing dependencies will cause ClassNotFoundException at runtime
- Log4j dependencies are mandatory for POI 5.4.1 logging functionality
- Commons Lang3 is required by Commons Compress used internally by POI