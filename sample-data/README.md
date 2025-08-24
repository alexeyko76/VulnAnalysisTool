# Sample Data

This directory contains sample Excel files for testing the vulnerability analysis tool.

## Files

- `sample.xlsx` - Sample Excel workbook with test data for vulnerability analysis
  - Contains sheets with file path information
  - Used by default configuration for testing and demonstration

## Usage

The sample files are referenced in the main `config.properties` file:

```properties
excel.path=./sample-data/sample.xlsx
sheet.name=Export
```

## Adding New Sample Files

You can add additional Excel files to this directory for testing different scenarios:
- Different column layouts
- Multiple sheets
- Various file path formats
- Different hostname configurations

Make sure to update the `config.properties` file to point to your desired sample file when testing.