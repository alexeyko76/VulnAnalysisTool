@echo off
rem Maven run script for Java Excel Tool
rem Runs the tool using Maven exec plugin or the built uber JAR

echo Running Java Excel Tool...
echo.

rem Check if uber JAR exists
if exist "java-excel-tool-uber.jar" (
    echo Using existing uber JAR: java-excel-tool-uber.jar
    java -jar java-excel-tool-uber.jar config.properties
) else if exist "target\java-excel-tool-uber.jar" (
    echo Using uber JAR from target directory: target\java-excel-tool-uber.jar
    java -jar target\java-excel-tool-uber.jar config.properties
) else (
    echo Uber JAR not found. Building first...
    call maven-build.bat
    if %ERRORLEVEL% neq 0 (
        echo ERROR: Build failed, cannot run
        exit /b 1
    )
    echo.
    echo Running newly built JAR...
    java -jar java-excel-tool-uber.jar config.properties
)

echo.
if %ERRORLEVEL% equ 0 (
    echo Tool completed successfully
) else (
    echo Tool exited with error code: %ERRORLEVEL%
)

exit /b %ERRORLEVEL%