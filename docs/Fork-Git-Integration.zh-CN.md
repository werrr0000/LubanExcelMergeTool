# Fork / Git 合并工具集成

正式发布版本：`Luban Excel Merge 1.2.2`。面向使用者的操作说明参见 [用户手册](LubanExcelMerge-User-Guide.zh-CN.md)。`auto` 重算发生包部件丢失或结构校验失败时，会回退到重算前的完整 MERGED，并以警告状态返回，不会丢失冲突合并结果。

LubanExcelMerge 提供 Git 四文件协议：`BASE`、`LOCAL`、`REMOTE` 和 `MERGED`。Fork 应启动 GUI 可执行文件并等待其退出。

## 生成仓库级配置

先构建 Release 版本，然后运行：

```powershell
dotnet run --project LubanExcelMerge.Cli/LubanExcelMerge.Cli.csproj -c Release -- git-config `
  --gui "<安装目录>\LubanExcelMerge.Gui.exe" `
  --repo-root "<Git仓库根目录>"
```

该命令只输出待确认的 `.gitconfig` 片段和 `git config --local` 命令，不会修改全局或仓库 Git 配置。确认路径后执行输出的三条仓库级命令。

生成的核心配置等价于：

```ini
[merge]
    tool = LubanExcelMerge

[mergetool "LubanExcelMerge"]
    cmd = "'<安装目录>/LubanExcelMerge.Gui.exe' merge --base \"$BASE\" --local \"$LOCAL\" --remote \"$REMOTE\" --output \"$MERGED\" --repo-root '<Git仓库根目录>' --recalculate-with-excel auto"
    trustExitCode = true
```

## Fork 中使用

在 Fork 的自定义合并工具界面中，程序路径填写 `LubanExcelMerge.Gui.exe`，Arguments 填写：

```text
--base "$BASE" --local "$LOCAL" --remote "$REMOTE" --output "$MERGED" --repo-root "<Git仓库根目录>" --recalculate-with-excel auto
```

GUI 的 Arguments 可以省略 `merge`；为兼容已有配置，带 `merge` 的形式也继续支持。若直接编辑 Git 的 `mergetool.<tool>.cmd`，则按上方生成的配置保留 `merge`。

1. 在 Fork 中执行产生 Excel 冲突的 merge 或 rebase。
2. 对冲突的 `.xlsx` 选择使用配置的 external merge tool。
3. 在 LubanExcelMerge 中解决全部冲突并保存 MERGED。
4. GUI 仅对当前 MERGED 执行 `git add -- <仓库相对路径>`，确认 unmerged 条目消失且 Git index 存在 stage-0 条目。
5. 保存和暂存均成功后 GUI 自动退出并返回 `0`；Fork 中该文件应显示为 resolved 并位于 staged。
6. 未保存直接关闭或仍有冲突时返回 `1`；输入、工作簿安全、写入校验、项目校验失败分别返回 `2`、`3`、`4`、`5`，Fork 不应将文件标记为 resolved。

GUI 的“取消合并”与窗口关闭行为相同。外部合并调用中，保存成功前的取消返回 `1`；如果界面已经显示加载或保存错误，关闭时会保留该错误对应的准确退出码。

自动暂存要求 `$MERGED` 位于 `--repo-root` 目录内并且已经被 Git 跟踪。工具会从 `$MERGED` 所在目录向上发现最近的 Git 根目录，因此支持父项目忽略 `ConfigLuban`、而 `ConfigLuban` 自身是嵌套 Git 仓库的布局。工具不会执行 `git add .`，也不会暂存同目录的其他修改。Git 或 Git LFS 暂存失败时，GUI 会保留已保存的 MERGED 供人工恢复、显示 Git 的具体错误，但不会返回成功或自动关闭。

## 公式重算

Fork 配置默认使用兼容参数 `--recalculate-with-excel auto`。参数名暂未更改，但重算后端会优先使用 WPS 表格，再回退到 Microsoft Excel：

- 合并未影响公式时不启动 WPS 或 Excel。
- 公式可能受影响时，依次检测 WPS 的 `ket.Application` / `et.Application` 和 Excel 的 `Excel.Application` 自动化接口。
- 自动重算只打开工具生成的临时工作簿；如果自动化接口复用了已有办公软件进程，工具会停止重算，不接管或关闭用户当前打开的文档。
- `auto` 模式下，WPS/Excel 自动化接口不可用、失败或超过 30 秒时会安全回退：保存已经通过结构验证的 MERGED，保留来源缓存并标记下次打开时完整重算。
- `always` 模式下重算失败仍不覆盖 MERGED；也可改用 `never` 直接保留来源缓存并标记下次打开时完整重算。
- 发生安全回退或使用 `never` 后，应使用 WPS/Excel 打开并保存一次，以确认最新公式结果。

## 多工作表合并

当 BASE、LOCAL、REMOTE 包含多个工作表时，工具会按工作表名称和顺序逐一建立独立合并计划：

- 三个版本的工作表名称、数量和顺序必须一致。
- 每个工作表独立解析 Luban 结构、选择主键、统计自动合并结果并生成冲突。
- 界面顶部的工作表标签显示该 sheet 的未解决冲突数；切换标签会同步切换 BASE、LOCAL、REMOTE 和 MERGED 四张表。
- “上一处/下一处”可以跨工作表跳转；行列批量选择、撤销和重做只作用于冲突所属工作表。
- 保存按钮以整本工作簿为门禁：任一 sheet 仍有未解决冲突时都不能保存。
- 所有 sheet 解决后，全部工作表编辑会在同一次原子保存中写入 MERGED。

当前多工作表工作簿要求每个 sheet 都能安全解析为平面 Luban 表，并暂不与 `validateLogicalTableUniqueness` 同时启用。包含说明页、图表页或受限层级结构时会指出具体不支持的工作表并阻止保存。

`__tables__.csv` 中声明为 `mode=one` 的单例逻辑表也支持合并：工具要求 BASE、LOCAL、REMOTE 每个工作表恰好包含一条数据记录，不使用主键，而是对这条记录的每个字段执行 BASE/LOCAL/REMOTE 三方规则。单边修改或双方相同修改自动合并，双方不同修改进入字段冲突复核；缺少记录或出现多条记录时会安全拒绝保存。

### Luban 元数据复核

BASE、LOCAL、REMOTE 的 Luban 元数据内容不一致时，不再仅因内容差异拒绝打开：

- 元数据内容变化以及元数据行的插入、删除、数据起始行变化都会作为“元数据复核”项进入合并界面。
- 元数据单元格使用橙色高亮，启动时优先选中并定位第一处未确认的元数据变化。
- 即使只有 LOCAL 或 REMOTE 单边修改，或者两边进行了相同修改，也需要明确接受 BASE、LOCAL 或 REMOTE 后才能保存。
- “上一处/下一处”、行列批量选择、撤销和重做同样适用于元数据复核项。
- 无界面模式遇到元数据变化时返回未解决冲突，不会自动写入 MERGED。

LOCAL 或 REMOTE 可以在任意位置新增字段列。单侧新增字段会自动写入 MERGED；双方新增不同字段时生成字段并集。MERGED 以 LOCAL 布局为基准，发生列位碰撞的 REMOTE 独有字段移动到首个安全空列。新字段数据按主键对齐，同名新字段的数据差异进入内容冲突，类型或元数据差异进入元数据复核。自动新增的元数据和数据可通过底部中央的“查看结果”按钮循环定位。

删除、重命名、重排或修改既有字段类型仍作为结构变化进入确认流程。元数据行插入或删除会以橙色高亮，选择 BASE、LOCAL 或 REMOTE 后自动移动后续数据行。REMOTE 独有的新字段如果包含公式且必须移动列位，工具会阻止合并，避免未经解析地改写公式引用。数据区中首个非空单元格以 `##` 开头的注释/停用行不会参与主键校验或数据合并。

## 项目配置与快速校验

工具会自动查找仓库根目录或 `ConfigLuban` 目录中的 `luban-excel-merge.json`。可复制 [配置示例](luban-excel-merge.example.json) 后按仓库路径调整；命令行显式传入的 `--data-root`、`--tables` 和公式策略优先于配置文件。

配置中的 `validation.enabled` 为 `true`，或命令行带 `--validate` 时，保存后会运行 `validation.windowsCommand`。未指定命令时默认使用 `ConfigLuban/check.bat`。

- 校验成功后 GUI/Fork 才会收到成功结果。
- 校验返回非零退出码或超时，工具会恢复保存前的 MERGED；原本没有 MERGED 时会删除本次候选输出。
- 校验开始前会创建恢复标记。若进程在写入候选 MERGED 后被终止，下次分析或保存同一路径时会先恢复旧 MERGED；原本不存在输出时会删除未提交的候选文件。
- 自动化测试不会运行业务仓库的真实 `check.bat`，避免改写项目生成目录；接入时应在项目副本上手工验证一次命令行为。

`keyOverrides` 以逻辑表全名为键，字段数组按顺序组成复合主键，并且优先于 `__tables__.csv` 的 `index`。`ignoredFields` 同样按逻辑表配置；已有记录的这些字段不参与冲突判断，MERGED 始终保留 LOCAL 值。工具会在界面和 CLI 中列出实际应用的忽略字段。

启用 `validateLogicalTableUniqueness` 后，工具会将当前文件的 BASE、LOCAL、REMOTE 分别与 `input` 中的兄弟工作簿组合校验。重复键错误包含版本、文件、工作表、行号和键值。当前跨文件扫描要求兄弟输入均为单工作表 `.xlsx`；遇到其他格式会明确阻止合并。

`inactivePaths` 支持仓库相对路径以及 `*`、`**`、`?` 通配符。匹配的兄弟文件不参与全逻辑表扫描；如果当前 MERGED 本身位于停用路径，工具会拒绝自动合并。绝对路径、包含 `..` 的模式和重复模式均会被拒绝。

### 隔离完整导出

完整导出默认关闭。通过 `--validate-full` 或配置 `validation.fullExportEnabled=true` 启用后，工具会：

1. 将 `ConfigLuban` 和不含 `.git` 的 `ConfigOutput` 复制到唯一临时目录。
2. 在临时仓库中运行 `validation.fullExportCommand`，默认是 `ConfigLuban/gen-pipeline.bat`。
3. 只依据隔离命令退出码判断结果，真实 `ConfigOutput` 不会被命令修改。
4. 校验失败时恢复原 MERGED；完成后尝试清理隔离副本，清理失败也不会触碰真实仓库。

隔离副本限制为最多 200000 个文件、4096 MB，并拒绝目录链接和文件链接。真实项目约需复制 `ConfigLuban` 与 `ConfigOutput`，因此完整校验明显慢于快速校验，适合手工确认或 CI，不建议每次冲突默认开启。

## 诊断日志

CLI 和 GUI 默认把 JSON Lines 日志写入 `%LOCALAPPDATA%\LubanExcelMerge\logs\日期`；CLI 可用 `--log <路径>` 指定位置。日志包含：

- 工具版本、系统、运行模式。
- BASE、LOCAL、REMOTE、MERGED 的绝对路径、SHA-256、存在状态和 LFS 指针状态。
- 逻辑表、主键、自动编辑/冲突/新增/删除数量。
- 公式重算、快速校验、隔离完整导出和全逻辑表唯一性结果。
- 异常类型、退出码和调用栈。

日志不记录单元格正文，也不会写入可能包含记录键值的异常消息；日志不会上传到网络。

GUI 工具栏可直接导出诊断包。命令行等价操作为：

```powershell
dotnet LubanExcelMerge.dll diagnostic-package `
  --log "<日志.jsonl>" `
  --output "<诊断包.zip>"
```

诊断包仅包含 `diagnostics.jsonl` 和 `manifest.json`，后者记录版本、系统、日志哈希及 `includesWorkbooks=false`；默认不会扫描或打包日志目录旁的任何 `.xlsx`。

## Git LFS

四个输入若是 Git LFS 指针，工具会从仓库本地 LFS 对象库解析真实工作簿。对象缺失时先执行：

```powershell
git lfs fetch
```

工具不会修改 `.git/lfs/objects`。

自动化端到端测试会建立临时 Git LFS 仓库，让两个分支修改同一 `.xlsx`，再通过真实 `git mergetool` 完成四文件协议、解除冲突、提交及再次 checkout 校验。在 Git 2.36.1 / Git LFS 3.1.4 的当前验收环境中，mergetool 传入的是 LFS 已解包的实体工作簿；指针输入路径由独立测试覆盖。Fork 图形界面的菜单启动和视觉状态仍需按发布版本执行一次手工验收。
