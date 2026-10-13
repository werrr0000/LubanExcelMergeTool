# 项目表格兼容性改进

## 识别与合并规则

- 横向表允许顶部和元数据区内的空白行，兼容标记两侧空格及 BOM。识别时不改写原始单元格，保留实际行号和未序列化的空行。
- `read_schema_from_file=FALSE` 时，在 `Datas` 同级的 `Defines` 下递归读取 XML。通过逻辑表模块和 `value_type` 精确匹配 bean，按 Excel `##var` 字段名映射类型，不按 XML 顺序猜测 Excel 列位置。
- 支持嵌套 bean 继承、一致类型的多态字段以及 `$type`、`__type__`。没有显式索引时使用 XML 首字段，避免 Excel 调整列顺序后误选主键。
- 数值主键识别 `int!#ref=...` 等修饰类型；可空主键保留空值与零的区别。
- `#`、`#Name` 等注释列仍参与内容合并，允许重复标签，不作为推断主键，也不要求出现在 XML 中。
- `##column##var` 纵向表按 A 列字段名对齐物理行，支持值与说明的三方合并、字段新增/删除和冲突选择。不转置实际文件，继续使用原单元格地址和样式来源。
- `singleton` 与 `one` 同样支持。`list` 仅在存在三方唯一稳定键、已有记录相对顺序不变时支持，不通过枚举行组合猜测主键。
- 无 Luban 标记的普通字段行在三方布局一致、具有唯一 Id 字段时按主键合并。说明页、辅助页和空页在三方一致时原样保留；内容变化按固定单元格地址复核，布局、样式、批注和图形保留 LOCAL 并在 GUI 保存前确认。含有 Luban 标记但缺少 `##var` 的页仍报错，不会被悄悄略过。详见 [普通字段行与辅助页规则](Remaining-Table-Compatibility.zh-CN.md)。
- 横向表解析错误指出 BASE / LOCAL / REMOTE、Sheet 与具体字段。

## 安全边界

- 三方 Sheet 名称、数量、顺序仍必须一致。
- 无 `##var` 的横向数据不按列顺序套用 XML；普通字段行从 Excel 字段名读取列位置并验证 Id，未知 XML 字段、缺失 bean、同名异类型字段仍拒绝映射。
- XML 来自当前工作目录。四文件调用没有提供历史 XML 的版本信息；合并涉及 XML 定义变化时，应先解决定义冲突，再使用项目导出校验确认结果。
- 显式 `parent` 属性继承、`*field` 跨行展开尚不支持；嵌套 bean 继承已支持。
- 纵向表类型变化、列头布局变化、顶部说明区行数变化仍拦截；暂不支持 `keyOverrides`、`ignoredFields` 和跨文件唯一性扫描。字段名为空但含其他内容的行不会静默丢弃。
- 辅助页不自动推断记录身份和行列移动；GUI 必须复核内容变化并确认 LOCAL 布局保留规则，命令行不能绕过确认直接保存。
- 本次未改变既有公式引用处理策略，也未替代项目自身的 Luban 导出校验。

## 项目检查（2026-10-13）

只读扫描 `GameSSSC/ConfigLuban/Datas/__tables__.csv` 登记且实际存在的 491 个工作簿。把每个文件同时作为 BASE、LOCAL、REMOTE，在工具工作区复制表目录及 XML 进行预检查，未保存或修改源项目 Excel。

483 个完成预检查（98.4%）。这验证当前文件读取、字段映射和键校验，不等于覆盖所有历史分支冲突。原始扫描还发现了 47 个纵向 Sheet，以及多处空页、说明页和辅助页。

解除占用后，单独复查 `DialogConfig_Xjd_Branch.xlsx` 已通过：识别 1 个数据 Sheet、0 个冲突。结合此前扫描，已确认通过的文件累计为 484/491（98.6%）；其余文件未在此次单文件复查中重新扫描。

| 未通过文件 | 原因 |
| --- | --- |
| `DialogConfig_Xjd_Branch.xlsx` | 原为文件占用；解除占用后复查通过 |
| `game_function/GameFunction.xlsx` | 工作簿格式不可读 |
| `disc/DiscSlotGroup.xlsx` | `*slot_info` 跨行展开字段 |
| `disc/ConstDisc.xlsx` | 第 31、53 行缺少字段名但有其他内容 |
| `equip/EquipmentBreakThroughConfigV2.xlsx` | 默认键重复，需确认稳定复合键 |
| `equip/EquipmentLevelUpTableConfigV2.xlsx` | 默认键重复，需确认稳定复合键 |
| `garden/GrowthStage.xlsx` | 默认键重复，需确认稳定复合键 |
| `NewHandWindow.xlsx` | 默认键 0 重复 |

Release 构建通过，146 项测试全部通过。新增回归覆盖 XML 继承、XML 首字段与 Excel 列序不同、带修饰数值键、重复注释列内容保留、纵向字段重排与冲突选择、纵向新增字段、空白表头行保存、辅助页保留和变化拦截。
