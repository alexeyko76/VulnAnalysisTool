@echo off
rem Optimized for very large Excel files (100,000+ rows)
rem Requires 8GB+ system RAM
rem -Xms4g: Initial heap size 4GB
rem -Xmx8g: Maximum heap size 8GB
rem -XX:+UseG1GC: Use G1 garbage collector
rem -XX:MaxGCPauseMillis=200: Target GC pause time
rem -XX:G1HeapRegionSize=16m: Optimize region size for large objects
java -Xms4g -Xmx8g -XX:+UseG1GC -XX:MaxGCPauseMillis=200 -XX:G1HeapRegionSize=16m -jar java-excel-tool-uber.jar config.properties