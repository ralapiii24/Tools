# catX 设备分类迁移到 `CiscoBase` 可行性分析

## 目标
将 `cat1` ~ `cat6` 的设备分类规则统一抽象到 `CiscoBase.py`，为 `DeviceBackupTask`、`ACLDupCheckTask`、`ACLCrossCheckTask` 等任务提供一致的判定接口，避免多个模块维护近似的正则和列映射逻辑。

## 当前使用情况
- `DeviceBackupTask` 使用 `cat1` ~ `cat6` 作为分类标签，并在 `_DEVICE_CATEGORY_RULES` 等结构中维护文件路径、命名前缀和设备类型描述；后续的路径归类、日志分组、命令拆分也以该分类为依据。
- `ACLDupCheckTask` 居中实现 `_get_device_classification_rules()` + `_is_cat1_device/_is_cat2_device`，通过第一行判断列类型，将 cat1/cat2 设备映射到 Excel 列。
- `ACLCrossCheckTask` 除了 cat1/cat2，还扩展了 cat3/cat4/cat5/cat6 的识别；`analyze_first_row_for_cat1_cat2` 使用设备名称匹配和预定义编号进行筛选，并返回多个分类目标列。

## CiscoBase 可以提供的能力
1. **分类规则集中管理**：把 `catX` 标签与所需关键字（如 `cs`, `n9k`, `link-as` 等）写成常量规则表，可在同一位置维护。
2. **通用判定函数**：
   - `classify_device_name(name: str) -> str | None`：返回 `cat1`~`cat6` 或 `None`。
   - `gather_device_cols(worksheet) -> dict[str, list[int]]`：根据分类返回列索引。
3. **配置驱动**：支持通过 `Config.yaml` 或 `CiscoBase` 常量控制 cat1/cat2 等的编号范围，便于 `cat3` 以后扩展。

## 迁移流程建议
1. 在 `CiscoBase` 里定义 `DEVICE_CATEGORY_RULES`（关键字、编号、描述）和 `classify_device_name()`。
2. 将 `DeviceBackupTask` 中的分类数据结构调整为引用该规则，保持 `catX` 名称和路径结构同步。
3. 让 `ACLDupCheckTask`/`ACLCrossCheckTask` 调用 `CiscoBase.classify_device_name()`，只负责 `cat` 结果对应的列映射与后续处理；必要时提供 `classify_device_row()` 等辅助函数。
4. 现有的 `_is_catX_device`/`analyze_first_row...` 可逐步移除或转为简单封装调用共享函数，确保逻辑一致。

## 风险与注意点
- 新分类接口需要提供足够的灵活性，如同时支持名称关键字匹配和设备编号控制。
- `ACLCrossCheckTask` 的 `cat6` 还以 `OOB-DS` 描述存在特定处理逻辑，迁移前需确认 shared 版本与原判断完全一致。
- `DeviceBackupTask` 可能依赖 `catX` 作为字典 key，在迁移过程中尽量保持接口不变（或提供兼容适配）。

## 结论
- 目前所有涉及设备分类的逻辑都集中在三份任务中，且规则高度可重用。
- 建议在 `CiscoBase` 中新增分类模块，任务层通过统一接口访问，可显著降低冗余并方便未来扩展 `cat3/4/5/6`。

需要我把这个分类能力也写进现有的 `ACLDelimRefactorFeasibility.md` 里统一管理，还是保持两个分析文件并填写更详细的迁移步骤？*** End Patch**##

