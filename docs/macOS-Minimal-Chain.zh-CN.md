# macOS / Fork 最小链路

此版本增加 Avalonia 桌面入口，复用现有 Core、Luban、OpenXml、CLI 合并会话。
Windows WPF 和新桌面入口共用 `MergeCompletion.Save`，按保存、验证、暂存的顺序完成合并。

当前支持 Apple Silicon 自包含应用、四文件参数、四表只读预览、多 Sheet、逐项选择
BASE/LOCAL/REMOTE、清除选择、保存复核、取消和 Git 暂存。打开应用但没有参数时显示接入说明。
尚未实现原 WPF 的同步滚动、批量区域选择、撤销重做、完整高亮和大表性能优化。

## 构建与安装

需要 macOS、能构建 net8.0 的 .NET SDK，以及 Git；LFS 仓库还需要 Git LFS。

```bash
./scripts/publish-macos.sh osx-arm64
./scripts/install-macos.sh
```

默认安装到 `~/Applications/LubanExcelMerge.app`，已有同名应用会保留为带时间戳的备份。
Intel 构建参数为 `osx-x64`，本次没有在 Intel 硬件上验证。
发布包包含 .NET 运行时，使用者无需安装 SDK。当前使用本机 ad-hoc 签名，尚未做开发者签名、公证。

## Mac Fork 自定义合并工具

在 Fork 的设置/偏好设置中找到合并工具，选择自定义工具。配置：

- 可执行文件：`/Users/你的用户名/Applications/LubanExcelMerge.app/Contents/MacOS/luban-excel-merge`
- 参数：`"$LOCAL" "$REMOTE" "$BASE" "$MERGED"`

实际路径以安装脚本输出为准。启动脚本根据 MERGED 所在目录确定 Git 根目录，支持空格、中文和 worktree。
不要把 `.app` 目录本身作为命令，也不要使用 `open -a` 作为 Git 的等待入口。
启动脚本以 `exec` 传递进程退出码；每次调用都启动独立会话。

遇到 `.xlsx` 冲突时，从 Fork 的外部合并工具入口打开。逐项选择版本，切换并复核全部 Sheet，
勾选复核确认后点击“保存并标记已解决”。成功后写入并验证 MERGED，只暂存该文件，然后返回成功。
取消不会写入或暂存。保存失败保持非零退出码；若保存已成功但暂存失败，保留输出并显示错误。

本机 Fork 2.70.1 的程序资源包含上述四个占位符。当前环境不允许通过 System Events 自动操作 Fork，
因此没有修改用户的全局 Fork 设置，也未完成真实 Fork 菜单点击验收；真实 `git mergetool`、启动脚本、
桌面生命周期、成功/取消退出码与索引状态已由集成测试验证。

## 直接运行和 Git 配置

```bash
~/Applications/LubanExcelMerge.app/Contents/MacOS/luban-excel-merge \
  /path/base.xlsx /path/local.xlsx /path/remote.xlsx /path/repo/ConfigLuban/Datas/Test.xlsx
```

也支持原有命名参数，包括 `--repo-root`、`--tables`、`--data-root`、`--config`、`--log`。
显式 `--headless` 使用原 CLI 行为，不显示窗口，不自动解决冲突，也不自动暂存。

生成仓库级 Git 配置（仅输出，不直接修改）：

```bash
~/Applications/LubanExcelMerge.app/Contents/MacOS/luban-excel-merge git-config \
  --gui "$HOME/Applications/LubanExcelMerge.app/Contents/MacOS/luban-excel-merge" \
  --repo-root /path/to/repository
```

确认并执行输出的命令后，可使用 `git mergetool --tool=LubanExcelMerge`。
Git 的工具配置和 Fork 自定义工具配置是两个入口，不假设 Fork 自动读取 `merge.tool`。

## 公式和校验脚本

Mac 首版没有 Office 自动重算引擎。沿用公共保存逻辑：

- `never`：保留来源公式缓存；公式可能受影响时标记下次打开重算。
- `auto`：没有重算引擎时安全退回延迟重算，并明确提示缓存未验证。
- `always`：没有可用引擎时阻止保存，返回错误。

公式缓存未验证时，保存并暂存后窗口仍显示提示，需要点击“完成并返回 Fork”。
依赖公式结果的表，应在 Excel/WPS 中重算并保存，重新暂存后再导出。最小版本不能保证公式缓存可直接用于 Luban 导出。

校验配置新增 `validation.unixCommand` 和 `validation.unixFullExportCommand`。
Windows 仍使用 `windowsCommand` 和 `fullExportCommand`。Unix 完整导出命令未指定时允许使用原 `fullExportCommand`，
但 `.bat/.cmd` 会给出明确错误，不能在 Mac 上执行。默认脚本分别为 `ConfigLuban/check.sh`、`ConfigLuban/gen-pipeline.sh`。
`.sh` 通过 `/bin/bash` 运行，因此不依赖复制后是否保留可执行位。

## 验证

```bash
./scripts/test-cross-platform.sh
# 在 Mac 上检查原 WPF 代码仍可编译，不代表 Windows 运行验证：
dotnet build LubanExcelMerge.Gui -c Release -p:EnableWindowsTargeting=true
```

桌面测试使用 Avalonia Headless，实际执行界面按钮、公共引擎和临时 Git 仓库。
测试驱动只存在于测试项目，不向正式应用添加自动选择冲突的隐藏参数。

本次验证：原有非 GUI 测试 144 项，新增桌面/集成测试 12 项；Windows WPF 交叉编译成功。
新增覆盖：取消不修改文件、三个版本选择、保存复核门禁、公式提示、保存期间禁止关闭、多 Sheet 门禁、
暂存失败、输入错误、Unix 引号、符号链接边界，以及真实 `git mergetool` 经启动脚本的成功/取消流程。
另已启动 Apple Silicon 原生应用并确认窗口创建；截图和 Fork 菜单操作未获得系统权限。

## 临时文件生命周期修复

2026-09-30 实测：合并窗口仍打开时，Fork 传入的 BASE、LOCAL、REMOTE 临时文件已不存在。
原逻辑在保存时重新打开 LOCAL，导致 FileNotFoundException（退出码 4）。
桌面入口现在在加载完成前保留独立的 LOCAL 工作簿模板；BASE/REMOTE 已解析为会话数据，
保存不再依赖调用方的三个临时文件。关闭窗口释放该模板，保存失败时保持模板以供重试。
异常日志新增 missingFiles 字段，仅记录缺失文件路径，不记录单元格内容。
新增回归测试会在窗口加载后删除全部三个输入，再选择 REMOTE 并保存，检查结果与 Git 索引。
本次新增测试通过，桌面测试累计 13 项。
