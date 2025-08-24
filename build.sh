#!/usr/bin/env bash
set -euo pipefail
rm -rf target/uber
mkdir -p target/uber/classes target/uber/stage

# 1) Compile thin classes against deps
javac -source 1.8 -target 1.8 -cp "deps/*" -d target/uber/classes src/main/java/app/ExcelTool.java

# 2) Thin jar
jar cfe target/uber/app-thin.jar app.ExcelTool -C target/uber/classes .

# 3) Unpack thin jar and all deps
pushd target/uber/stage >/dev/null
jar xf ../app-thin.jar
for j in ../../../deps/*.jar; do
  [ -f "$j" ] || continue
  unzip -q "$j" || true
done

# 4) Remove signatures (avoid SecurityException)
rm -f META-INF/*.SF META-INF/*.DSA META-INF/*.RSA || true

# 5) Ensure manifest
printf "Main-Class: app.ExcelTool
" > MANIFEST.MF

# 6) Repack uber jar
jar cfm ../../../java-excel-tool-uber.jar MANIFEST.MF .
popd >/dev/null

echo "Built uber jar: java-excel-tool-uber.jar"
echo "Run with: java -jar java-excel-tool-uber.jar config.properties"
