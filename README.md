# Luban Excel Merge Tool

面向 Git 和 Fork 的 Luban `.xlsx` 三方冲突合并工具。程序读取 BASE、LOCAL、REMOTE，在桌面界面中展示差异和冲突，并将用户确认后的结果原子保存到 MERGED。

当前版本：`1.1.0`

## 功能

- BASE、LOCAL、REMOTE、MERGED 四表同步对比
- 红色冲突、黄色删除、绿色新增、橙色元数据变更高亮
- 上一处/下一处冲突定位及自动编辑循环定位
- 单元格、整行和整列批量选择 BASE/LOCAL/REMOTE
- 多 Sheet 切换与整工作簿保存门禁
- 新增行、删除行、修改和删除/修改冲突处理
- LOCAL/REMOTE 任意位置新增 Luban 字段列及字段并集合并
- 既有字段列删除、重命名、类型修改和位置移动的三方合并与冲突选择
- 危险列结构变化保存前二次确认
- Git/Fork resolved 与 staged 集成
- Git LFS 输入解析
- WPS 优先、Microsoft Excel 回退的公式重算
- 原子保存、重新打开验证、项目校验与诊断包

## 项目结构

- `LubanExcelMerge.Core`：三方合并算法
- `LubanExcelMerge.Luban`：Luban 元数据、主键和逻辑表解析
- `LubanExcelMerge.OpenXml`：工作簿读取、最小编辑和原子保存
- `LubanExcelMerge.Git`：Git LFS、mergetool 和 staging 集成
- `LubanExcelMerge.Cli`：命令行协调器
- `LubanExcelMerge.Gui`：Windows WPF 用户界面
- `*.Tests`：无外部测试框架依赖的回归测试程序

## 构建

需要 Windows 和 .NET 8 SDK。

```powershell
dotnet restore LubanExcelMerge.sln
dotnet build LubanExcelMerge.sln -c Release --no-restore
```

## 测试

```powershell
dotnet run --project LubanExcelMerge.Core.Tests -c Release
dotnet run --project LubanExcelMerge.Luban.Tests -c Release
dotnet run --project LubanExcelMerge.OpenXml.Tests -c Release
dotnet run --project LubanExcelMerge.Cli.Tests -c Release
```

当前基线为 `105` 项测试。

## 发布

```powershell
dotnet publish LubanExcelMerge.Gui/LubanExcelMerge.Gui.csproj `
  -c Release -r win-x64 --self-contained true `
  -p:PublishSingleFile=true `
  -p:IncludeNativeLibrariesForSelfExtract=true `
  -p:PublishTrimmed=false
```

## 文档

- [用户手册](docs/LubanExcelMerge-User-Guide.zh-CN.md)
- [Fork / Git 集成](docs/Fork-Git-Integration.zh-CN.md)
- [配置示例](docs/luban-excel-merge.example.json)

## 当前结构限制

- 输入必须为 `.xlsx`。
- 三方工作簿的 Sheet 名称、数量和顺序必须一致。
- 不支持插入或删除 Luban 元数据行，也不支持改变数据起始行。
- 删除唯一可用的主键字段后，如果无法建立可靠的行身份，主键校验会阻止合并。
- 列移动会保留各分支已有的公式文本并触发重算，但不会猜测或重写公式引用；保存前应重点复核公式字段。
