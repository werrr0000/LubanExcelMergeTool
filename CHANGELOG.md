# Changelog

## Unreleased

### Added

- 既有字段列删除、重命名、类型修改和位置移动按 BASE/LOCAL/REMOTE 三方规则自动合并或生成可解决冲突。
- 保存时根据最终列结构重新映射行级自动编辑和冲突选择，避免列移动后数据写入旧列。
- GUI 对既有字段列结构变化显示默认拒绝的二次确认，确认后继续原子保存、校验和 Git staging。

### Safety

- 结构性缺列不再被误判为逐单元格清空；选择保留另一侧列时会继续合并该侧数据。
- 保存开始后显示不可关闭的进度窗口，禁用主窗口并置灰系统关闭按钮；成功或失败完成前拒绝所有窗口关闭请求。
- 列移动保留公式文本并触发既有重算流程，确认窗口明确提示复核公式引用。

### Validation

- 自动化回归套件：105 个测试全部通过。

## Luban Excel Merge 1.1.0 - 2026-07-29

### Added

- 支持 LOCAL 或 REMOTE 在任意位置新增 Luban 字段列。
- LOCAL 与 REMOTE 分别追加不同字段时，MERGED 自动生成字段并集；发生列位碰撞的远端字段会写入下一个安全空列。
- 新字段数据按记录主键对齐，覆盖已有记录、远端新增记录和多工作表保存。
- 两侧新增同名字段但数据不同会进入单元格冲突；新字段类型或元数据不同会进入元数据复核。
- BASE、LOCAL、REMOTE 和 MERGED 四表按目标字段布局显示新增列，自动编辑提示可遍历新增元数据与数据。

### Safety

- 删除、重命名、重排或修改既有字段类型仍会阻止保存。
- MERGED 保留 LOCAL 的字段布局；REMOTE 独有字段发生列位碰撞时移动到首个安全空列，包含公式且需要移动的远端新字段会明确阻止以避免公式引用失真。

### Validation

- 自动化回归套件：98 个测试全部通过。

## Luban Excel Merge 1.0.1 - 2026-07-27

### Fixed

- Fork 自定义合并工具可直接传入 `--base`、`--local`、`--remote`、`--output` 等 Arguments，不再因省略 `merge` 子命令而拒绝启动。
- 未知子命令仍会被明确拒绝，现有带 `merge` 的 Git mergetool 配置保持兼容。

### Validation

- 自动化回归套件：90 个测试全部通过。

## Luban Excel Merge 1.0.0 - 2026-07-24

### Added

- BASE/LOCAL/REMOTE three-way merge for Luban `.xlsx` workbooks.
- Pixel-art application and window icon.
- Conflict review with synchronized BASE, LOCAL, REMOTE and MERGED grids.
- Cell, row and column conflict resolution with undo and redo.
- Previous/next conflict navigation across multiple worksheets.
- Automatic-edit navigation that cycles across every affected cell or row.
- Multi-sheet merge with a workbook-wide save gate.
- Luban metadata review with dedicated highlighting and save blocking.
- WPS-first formula recalculation with Microsoft Excel fallback.
- Atomic save, reopen validation and rollback recovery.
- Fork/Git four-file protocol integration, Git LFS input handling and automatic staging.
- Project validation, isolated full-export validation and sanitized diagnostic packages.

### Performance

- Indexed worksheet rows, cells, Luban fields, records and conflicts.
- Cached automatic-edit locations and static comparison tables.
- Lazy comparison-row materialization for large worksheets.
- Incremental row and column selection updates during navigation.
- Phase-level preparation timings in diagnostic logs.

### Validation

- Accepted in the target Fork, Git LFS and WPS workflow.
- Automated regression suite: 88 tests passed.
- Synthetic 10,000-row benchmark: preparation below one second in the acceptance environment.
