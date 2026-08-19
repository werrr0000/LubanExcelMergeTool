# Changelog

- 支持 Luban 元数据行插入和删除：差异以橙色元数据复核展示，确认 BASE、LOCAL 或 REMOTE 后自动移动数据区并保留公式。

## Unreleased

### Added

- 支持 Luban `mode=one` 单例逻辑表：BASE、LOCAL、REMOTE 各含一条记录时按字段执行三方合并，兼容修改自动合并，字段冲突进入现有复核流程。
- 既有字段列删除、重命名、类型修改和位置移动按 BASE/LOCAL/REMOTE 三方规则自动合并或生成可解决冲突。
- 保存时根据最终列结构重新映射行级自动编辑和冲突选择，避免列移动后数据写入旧列。
- GUI 对既有字段列结构变化显示默认拒绝的二次确认，确认后继续原子保存、校验和 Git staging。
- 底部中央新增“查看结果”按钮，用于逐个浏览完整的自动合并结果（包括无需额外写入的 LOCAL 变化）；左侧状态文字恢复为纯状态显示并提高对比度。
- 蓝色统一表示已处理的修改结果和用户已解决冲突；绿色新增、黄色删除保持原有语义。“查看结果”按工作表、行、列逐格循环定位全部蓝色格子，左下角同步显示总数和当前序号。

### Safety

- 仅包含样式或格式的空行、空列不再被误判为有效数据边界；保存 MERGED 时会移除这些空白格式占位并收缩工作表使用区域，新增记录和字段紧接最后一个有内容的行、列写入。
- 空白格式清理不移动任何有效单元格坐标，并保留有内容行列上的样式、公式文本及公式缓存，避免物理移位导致公式引用变化。
- 结构性缺列不再被误判为逐单元格清空；选择保留另一侧列时会继续合并该侧数据。
- 保存开始后显示不可关闭的进度窗口，禁用主窗口并置灰系统关闭按钮；成功或失败完成前拒绝所有窗口关闭请求。
- 列移动保留公式文本并触发既有重算流程，确认窗口明确提示复核公式引用。
- `auto` 公式重算超过 30 秒、自动化接口不可用或重算失败时，安全保留已验证的 MERGED、来源公式缓存及下次打开完整重算标记；`always` 继续保持严格失败语义。
- 诊断日志记录脱敏的完整异常类型链和公式重算降级原因，不记录单元格正文。
- GUI 保存失败提示会显示最具体的底层验证原因，不再只显示统一的 MERGED 验证失败消息。
- 用户解决冲突后，四表对应格子可靠地由红色刷新为自动合并蓝色；撤销解决后立即恢复红色。
- `__tables__.csv` 兼容 Luban 元数据行中未加引号字段包含裸引号的写法，不再在忽略该行前提前报错。

### Validation

- 自动化回归套件：115 个测试全部通过。

## Luban Excel Merge 1.2.2 - 2026-08-13

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
