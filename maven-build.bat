@echo off
rem Maven build script for Java Excel Tool
rem This builds the uber JAR using Maven instead of manual compilation

echo Building Java Excel Tool using Maven...
echo.

rem Clean previous builds
echo Cleaning previous builds...
mvn clean

if %ERRORLEVEL% neq 0 (
    echo ERROR: Maven clean failed
    exit /b 1
)

rem Compile and package
echo Compiling and creating uber JAR...
mvn package

if %ERRORLEVEL% neq 0 (
    echo ERROR: Maven build failed
    exit /b 1
)

rem Copy the uber JAR to root directory for consistency with existing scripts
if exist "target\java-excel-tool-uber.jar" (
    copy "target\java-excel-tool-uber.jar" "java-excel-tool-uber.jar" >nul
    echo.
    echo SUCCESS: Build completed successfully!
    echo Output: java-excel-tool-uber.jar
    echo Run with: java -jar java-excel-tool-uber.jar config.properties
    echo Or use: maven-run.bat
) else (
    echo ERROR: Uber JAR was not created in target directory
    exit /b 1
)

exit /b 0