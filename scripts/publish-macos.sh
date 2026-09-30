#!/bin/bash
set -euo pipefail
repo_dir="$(cd "$(dirname "$0")/.." && pwd)"
rid="${1:-osx-arm64}"
case "$rid" in osx-arm64|osx-x64) ;; *) echo '支持 osx-arm64 / osx-x64' >&2; exit 2;; esac
app_dir="$repo_dir/artifacts/$rid/LubanExcelMerge.app"
mkdir -p "$app_dir/Contents/MacOS" "$app_dir/Contents/Resources"
dotnet publish "$repo_dir/LubanExcelMerge.Desktop/LubanExcelMerge.Desktop.csproj" \
    -c Release -r "$rid" --self-contained true -p:PublishTrimmed=false \
    -o "$app_dir/Contents/MacOS"
cat > "$app_dir/Contents/Info.plist" <<'PLIST'
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0"><dict>
<key>CFBundleIdentifier</key><string>com.lubanexcelmerge.desktop</string>
<key>CFBundleName</key><string>LubanExcelMerge</string>
<key>CFBundleDisplayName</key><string>Luban Excel Merge</string>
<key>CFBundleExecutable</key><string>LubanExcelMerge.Desktop</string>
<key>CFBundlePackageType</key><string>APPL</string>
<key>CFBundleShortVersionString</key><string>1.2.2</string>
<key>CFBundleVersion</key><string>1</string>
<key>LSMinimumSystemVersion</key><string>12.0</string>
<key>NSHighResolutionCapable</key><true/>
</dict></plist>
PLIST
cp "$repo_dir/scripts/luban-excel-merge" "$app_dir/Contents/MacOS/luban-excel-merge"
chmod +x "$app_dir/Contents/MacOS/luban-excel-merge"
codesign --force --deep --sign - "$app_dir"
echo "已生成本机测试包：$app_dir"
