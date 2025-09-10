@echo off
rem Minimal memory configuration for systems with limited RAM
rem Suitable for smaller files (<10,000 rows) or systems with limited memory
rem -Xms512m: Initial heap size 512MB
rem -Xmx1g: Maximum heap size 1GB
java -Xms512m -Xmx1g -jar java-excel-tool-uber.jar config.properties