# ACL 定界剥离到 `CiscoBase` 可行性分析

## 目标
将 `ACLArpCheckTask.py`、`ACLDupCheckTask.py` 和（当前部分已使用）`ACLCrossCheckTask.py` 中重复的 ACL 定界逻辑集中到 `CiscoBase.py`，以便统一维护正则、网络转换、解析和块提取等基础能力，减少重复实现和潜在偏差。

## 当前重复逻辑
- `ACLArpCheckTask.py` 自行维护了 `NXOS/IOSXE/ASA` 正则、`parse_acl_line`、`_is_acl_rule` 和 `find_acl_blocks_in_column` 等函数；几乎与 `CiscoBase` 中已有的 `parse_acl`、`is_acl_rule`、`find_acl_blocks_in_column` 完整重叠。
- `ACLDupCheckTask.py` 也包含与 `CiscoBase` 中 `service_to_port`、`ip_and_wildcard_to_network`、`cidr_to_network`、`parse_acl`、`rule_covers` 等高度相似的实现，导致未来 Regex/解析规则更新需要三处同步。
- `ACLCrossCheckTask.py` 已在顶部 `from .CiscoBase import ...`，说明核心解析与定界能力已经在向共享模块迁移，但它仍保留一些额外的辅助匹配函数（如图论分组）与任务特有逻辑，未完全依赖 `CiscoBase` 中的块提取函数。

## `CiscoBase` 中可复用的能力
1. **地址/端口转换**：`ip_and_wildcard_to_network`、`host_to_network`、`cidr_to_network`、`any_to_network`，确保不同任务使用同一套网络解析结果。
2. **端口解析**：`service_to_port` 与 `_extract_all_ports`，供覆盖判断/多端口展开。
3. **规则实体**：`ACLRule` 数据类封装了 `style`、`ports` 等字段，可直接用于 `ACLDup`/`ACLCross` 的覆盖判定。
4. **解析函数**：
   - `parse_acl_full`：支持各种 Cisco ACL 格式的底层解析。
   - `parse_acl`：在 `any` 规则、日志关键字等方面做了筛选，可在 `ACLArpCheckTask` 中复用。
   - `parse_acl_network_only`：分析仅网络，不需要端口，可供 `ACLDup`/`ACLArp` 在 `platform_network` 判断时使用。
5. **定界辅助**：
   - `is_acl_rule`：统一的 ACL 判定逻辑，便于 `ACLArpCheckTask` 拦截非规则行。
   - `find_acl_blocks_in_column`/`extract_acl_rules_from_column`：直接返回 `(start_row, end_row)` 和 解析后的 `ACLRule` 列表，适配正在重复实现的“按列找块”的需求。

## 迁移方案建议
1. **抽象接口**：在 `CiscoBase` 中保持 `find_acl_blocks_in_column`/`extract_acl_rules_from_column` 接口不变，确保任务端调用参数与返回值一致。
2. **任务层改造**：
   - `ACLArpCheckTask`：移除内部解析/定界函数，改为 `from .CiscoBase import (is_acl_rule, parse_acl, find_acl_blocks_in_column)`。保留自定义的 `platform_network_map` 读取与 Unicode 标记逻辑，仅在处理单元格时调用共享函数。
   - `ACLDupCheckTask`：将 `parse_acl`、`rule_covers` 等使用 `CiscoBase.ACLRule` 的部分改成直接调用共享函数；若 `rule_covers` 需要扩展（如更严格的端口比对），可在 `CiscoBase` 中提供可配置入口或保持任务内实现但以共享 `ACLRule` 为输入。
3. **逐步剥离**：先从每个任务中移除重复的 regex/parser 定义，确保 import `CiscoBase` 后逻辑保持一致，再清理多余代码。

## 潜在影响与注意点
- `CiscoBase.parse_acl` 默认忽略 `any` 规则；若某任务需要保留 `any`（例如 `ACLArpCheck` 需要识别 `any` 以决定是否跳过），需在调用后基于 `raw` 自行判断，而不是重新实现解析器。
- `ACLArpCheckTask` 目前依赖 `CONFIG` 中 `ignore_third_octet` 并在 `process_acl_block_with_unicode_marking` 内拼接 `platform_network`，这些上下文不在 `CiscoBase`，迁移仅限解析/定界，不影响业务逻辑。
- `ACLDupCheckTask` 的覆盖逻辑大量使用了 `ACLRule` 的 `ports`/`style` 字段，迁移时应确保 `parse_acl` 填充这些字段（现 `CiscoBase` 已实现）。
- 新增依赖 `openpyxl.Worksheet` 的类型提示（当前 `CiscoBase` 已通过 `try/except` 处理）。

## 结论与下一步
1. **可行性**：> 90% 的定界与解析逻辑在 `CiscoBase` 中就绪，其他任务只需替换 import，便能共享统一的正则与网络/端口转换。
2. **下一步**：
   - 逐个任务进行导入改造并去除重复定义。
   - 保留 `CiscoBase` 中的 `rule_covers`/`rule_matches` 等通用工具，供 `ACLDup`/`ACLCross` 使用。
   - 在改造过程中运行现有测试/任务，确保行为一致。


