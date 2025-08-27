@echo off
setlocal ENABLEDELAYEDEXPANSION

rem Check if JAVA_HOME is set, if not set a default value
if "%JAVA_HOME%"=="" (
    echo JAVA_HOME is not set, using default Java installation path
    set "JAVA_HOME=C:\Program Files\Java\jdk1.8.0_401"
    echo Using JAVA_HOME: !JAVA_HOME!
) else (
    echo Using JAVA_HOME: %JAVA_HOME%
)

rem Set Java executables path
set "JAVA_BIN=%JAVA_HOME%\bin"
set "JAVAC_CMD=%JAVA_BIN%\javac.exe"
set "JAR_CMD=%JAVA_BIN%\jar.exe"

rem Verify Java tools exist
if not exist "%JAVAC_CMD%" (
    echo ERROR: javac not found at %JAVAC_CMD%
    echo Please check your JAVA_HOME setting or install JDK 1.8
    exit /b 1
)

if not exist "%JAR_CMD%" (
    echo ERROR: jar not found at %JAR_CMD%
    echo Please check your JAVA_HOME setting or install JDK 1.8
    exit /b 1
)

rmdir /S /Q target\uber 2>NUL
mkdir target\uber\classes
mkdir target\uber\stage

rem 1) Compile
"%JAVAC_CMD%" -source 1.8 -target 1.8 -cp "deps\*" -d target\uber\classes src\main\java\app\ExcelTool.java || goto :err

rem 2) Thin jar
"%JAR_CMD%" cfe target\uber\app-thin.jar app.ExcelTool -C target\uber\classes . || goto :err

rem 3) Unpack
pushd target\uber\stage
"%JAR_CMD%" xf ..\app-thin.jar

rem Extract each dependency JAR individually  
"%JAR_CMD%" xf ..\..\..\deps\poi-5.4.1.jar
"%JAR_CMD%" xf ..\..\..\deps\poi-ooxml-5.4.1.jar
"%JAR_CMD%" xf ..\..\..\deps\poi-ooxml-lite-5.4.1.jar
"%JAR_CMD%" xf ..\..\..\deps\xmlbeans-5.3.0.jar
"%JAR_CMD%" xf ..\..\..\deps\commons-collections4-4.5.0.jar
"%JAR_CMD%" xf ..\..\..\deps\commons-compress-1.28.0.jar
"%JAR_CMD%" xf ..\..\..\deps\commons-io-2.20.0.jar
"%JAR_CMD%" xf ..\..\..\deps\commons-lang3-3.12.0.jar
"%JAR_CMD%" xf ..\..\..\deps\log4j-api-2.17.2.jar
"%JAR_CMD%" xf ..\..\..\deps\log4j-core-2.17.2.jar

rem 4) Remove signatures
del /Q META-INF\*.SF 2>NUL
del /Q META-INF\*.DSA 2>NUL
del /Q META-INF\*.RSA 2>NUL

rem 5) Manifest
> MANIFEST.MF echo Main-Class: app.ExcelTool

rem 6) Uber jar
"%JAR_CMD%" cfm ..\..\..\java-excel-tool-uber.jar MANIFEST.MF . || goto :err
popd

echo Built uber jar: java-excel-tool-uber.jar
echo Run with: java -jar java-excel-tool-uber.jar config.properties
exit /b 0
:err
echo Build failed. Check errors above.
exit /b 1
