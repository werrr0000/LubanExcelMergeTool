# 1.2.2.2 保存性能修复

## 原因与变更

现场 1.2.2.1 日志显示 GameFunction、ItemConfig 保存都等待 Excel 30 秒超时。
工具先写入 forceFullCalc=1 和 fullCalcOnLoad=1，Excel 打开后即使 CalculateFullRebuild 已返回，CalculationState 仍保持 xlPending。

- 在实际调用 Excel/WPS 重算前，将候选文件这两个标记设为 0，然后照常显式执行 CalculateFullRebuild 并确认计算完成。
- 原子保存器保留独立的重算前副本；auto 重算失败仍恢复带完整延迟重算标记的候选文件，never 模式不变。
- Office 省略 ZIP 空目录项不再误报为 OpenXML 部件丢失。实际文件部件（含空文件）、批注和协作者信息的丢失仍会被拒绝。

## 验证

- OpenXml 41 项测试全部通过，新增延迟标记切换及目录项省略测试；覆盖样式安全、重算失败回退、persons 部件丢失等场景。
- 真实三方输入在独立临时目录测量，包含 CLI 启动、读取、合并、保存、完整 Excel 重算及结果校验，不包含 Fork 的 Git staging：GameFunction 2745ms，ItemConfig 3754ms。结果均为 Completed，无重算警告。
- EPPlus 8.0.3 逐格读取 Text：GameFunction 1472 次、ItemConfig 24625 次全部通过。
- 未更改配置源表；旧安装在替换前备份。版本 1.2.2.2 为本地修复版，未推送上游。
