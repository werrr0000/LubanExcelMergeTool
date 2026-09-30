#!/bin/bash
set -euo pipefail
repo_dir="$(cd "$(dirname "$0")/.." && pwd)"
for component in Core Luban OpenXml Cli Desktop; do
    dotnet run --project "$repo_dir/LubanExcelMerge.$component.Tests" -c Release
done
