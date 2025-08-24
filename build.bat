@echo off
setlocal ENABLEDELAYEDEXPANSION
rmdir /S /Q target\uber 2>NUL
mkdir target\uber\classes
mkdir target\uber\stage

rem 1) Compile
javac -source 1.8 -target 1.8 -cp "deps\*" -d target\uber\classes ExcelTool.java || goto :err

rem 2) Thin jar
jar cfe target\uber\app-thin.jar app.ExcelTool -C target\uber\classes . || goto :err

rem 3) Unpack
pushd target\uber\stage
jar xf ..\app-thin.jar

for %%J in (..\..\deps\*.jar) do (
  powershell -NoLogo -NoProfile -Command "Add-Type -A 'System.IO.Compression.FileSystem'; [IO.Compression.ZipFile]::ExtractToDirectory('%%~fJ', '.')" 2>NUL
)

rem 4) Remove signatures
del /Q META-INF\*.SF 2>NUL
del /Q META-INF\*.DSA 2>NUL
del /Q META-INF\*.RSA 2>NUL

rem 5) Manifest
> MANIFEST.MF echo Main-Class: app.ExcelTool

rem 6) Uber jar
jar cfm ..\..\..\java-excel-tool-uber.jar MANIFEST.MF . || goto :err
popd

echo Built uber jar: java-excel-tool-uber.jar
echo Run with: java -jar java-excel-tool-uber.jar config.properties
exit /b 0
:err
echo Build failed. Check errors above.
exit /b 1
