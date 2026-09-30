#!/bin/bash
set -euo pipefail
repo_dir="$(cd "$(dirname "$0")/.." && pwd)"
case "$(uname -m)" in arm64) rid=osx-arm64;; x86_64) rid=osx-x64;; *) exit 2;; esac
source_app="$repo_dir/artifacts/$rid/LubanExcelMerge.app"
destination_dir="${1:-$HOME/Applications}"
if [ ! -d "$source_app" ]; then "$repo_dir/scripts/publish-macos.sh" "$rid"; fi
mkdir -p "$destination_dir"
if [ -e "$destination_dir/LubanExcelMerge.app" ]; then
    mv "$destination_dir/LubanExcelMerge.app" "$destination_dir/LubanExcelMerge.backup-$(date +%Y%m%d-%H%M%S).app"
fi
ditto "$source_app" "$destination_dir/LubanExcelMerge.app"
codesign --verify --deep --strict "$destination_dir/LubanExcelMerge.app"
echo "Fork 自定义工具路径：$destination_dir/LubanExcelMerge.app/Contents/MacOS/luban-excel-merge"
echo 'Fork 参数："$BASE" "$LOCAL" "$REMOTE" "$MERGED"'
