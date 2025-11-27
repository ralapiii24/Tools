概览
入口流程：Main.py → 预检RuntimeEnvCheck.py → 自更新Updater.py → 执行核心任务APP/Core.py
主要职责：
Main.py：入口。先跑环境预检（依赖库 & Playwright 浏览器），通过后才进入后续逻辑。
################################################################################################################################################
RuntimeEnvCheck.py：环境预检。核查 pyyaml/tqdm/requests/lxml/paramiko/playwright/colorama 等包，并确认已安装 Playwright Chromium。失败会写入报告到 REPORT/<YYYYMMDD>_DependencyCheck.log，并退出（阻止更新 & 巡检）。
################################################################################################################################################

Updater.py：Updater.py：自更新器。读取 YAML/Version.yaml 的 FTP 参数与本地版本，从FTP读取最新版本号与版本目录，必要时下载解压、校验（大小/MD5），备份旧版本并替换，再启动 APP/Core.py 执行巡检。环境变量可覆盖 FTP 连接配置若需更新：递归下载对应远端版本目录 → 校验（大小/MD5，若服务器支持）→ 备份旧文件 → 替换（保留白名单：Main.py、Updater.py、YAML/Version.yaml）。不论是否更新，最终执行 APP/Core.py。
################################################################################################################################################
Core.py：多任务巡检（FXOS/FortiGate镜像/Oxidized配置抓取/各类Linux服务器SSH获取CPU内存硬盘指标/Kibana ESN9K日志扫描/Flow服务检查），写入日报与任务明细。
创建当日日志目录：LOG/<YYYYMMDD>/；日报目录：REPORT/。
依次运行任务并记录结果；按任务生成明细日志，并在 REPORT/<YYYYMMDD>巡检日报.log 写入汇总与异常摘要。
从 YAML/Config.yaml 读取配置；settings.show_progress 控制 tqdm 进度条；所有任务输出统一分级：OK/WARN/CRIT/ERROR。
每个任务为 BaseTask 子类：实现 items()（要巡检的对象列表）与 run_single(item)（单对象巡检逻辑）。
结果对象 Result(level, message, meta) 可带 meta 附加信息（如样例、原始行等）。
################################################################################################################################################
YAML/：
Config.yaml：所有任务的目标清单、账号口令、阈值、索引规则等运行时配置。
Version.yaml：自更新器所需 FTP 参数与本地版本号。
Ignore_alarm.yaml：ES/N9K 日志扫描的忽略规则（包含子串 or 正则）。
PipLibrary.bat：Windows 下一键安装依赖并安装 Playwright 的 Chromium。

################################################################################################################################################
FXOSWebTask（FXOS WEB 巡检）
技术栈：Playwright（Chromium，无头模式），自动忽略自签名/不可信证书。
目标：对 fxos.devices 中的每个 FXOS 管理地址进行自动登录验证。
交互细节：支持自动模拟 Enter/Tab/Continue/Proceed/OK/确认/确定 点击，处理浏览器安全/提示对话；登录后等待指定 XPath 成功标志。
输出：每台设备 “登录成功/失败”。
################################################################################################################################################
MirrorFortiGateTask（镜像飞塔防火墙巡检）
技术栈：Paramiko SSH。
指标：
日志盘占用：diagnose sys logdisk usage → 解析 used/total MB → 与 fortigate.thresholds.disk_percent 比较（默认 WARN≥60，CRIT≥80）。
CPU/内存/Uptime：get system performance status；如遇 --More-- 分页或缺失 Uptime，则切换交互分页（空格翻页）或再跑 get system status 兜底；
CPU：由 idle% 推算 used%
MEM：解析 used( xx % )
Uptime：支持多种格式解析（天/小时/分钟混排、中文/英文/冒号/空格等）。
输出：针对每台主机分别给出“日志盘 / CPU / 内存 / Uptime”的分级结果；解析失败会给 ERROR。
################################################################################################################################################
OxidizedTask（网络设备配置备份到本地）
目标：访问 oxidized.base_urls 下的 /nodes 页面，解析设备与分组，然后调用 /node/fetch/<group>/<device> 拉取配置。
落盘：保存为 LOG/<YYYYMMDD>/<YYYYMMDD>-<设备名>.log。
输出：统计“备份成功/失败”数量；失败附错误信息。
################################################################################################################################################
BaseLinuxServerTask（三类 Linux 主机健康）
适用：LogstashServerTask / ElasticsearchServerTask / FlowServerTask
通用检查：
根分区占用：df -h 解析 / 的 used%，与 servers.thresholds.disk_percent 比较（默认 WARN≥50，CRIT≥80）。
内存占用：free -m 解析 used/total 推算占用率，各类型有不同阈值：
LOGSTASH：WARN≥50，CRIT≥80
ES/FLOW：WARN≥80，CRIT≥90
################################################################################################################################################
FlowServerTask 额外检查：
关键端口：netstat -tulnp 寻找 5601/9200/9300/4739/2055/6343 等（来自 flow_checks.require_ports，并有兜底端口集合）。
容器状态：docker ps --format "{{.Names}} {{.Status}}"，要求 opt-kibana-1、elastiflow-logstash、elastiflow-elasticsearch 等容器处于 Up；若 docker 失败，会回退到端口命中结果作为提示。
ES 索引大小：/_cat/indices?v 过滤 index_prefix（默认 elastiflow-4.0.1-），解析末列大小（支持 K/M/G/T），超过 index_size_limit_bytes（默认 1GiB）则记录。
Segments 行数：/_cat/segments?v 针对非今昨的索引，若行数 > segment_max_non_recent（默认 3）则 WARN。
################################################################################################################################################
ESN9KLogInspectTask（Kibana 上 Cisco N9K 日志异常扫描）
目标：在一组 kibana_bases 中挑选可用的 Kibana（调用 /api/status），然后经 Kibana 反代查询 ES：
索引匹配：index_pattern: "*-n9k-*-*"
时间范围：time_gte: now-3d 到 time_lt: now（可改）
取数方式：_search?scroll 滚动查询，字段 @timestamp 与 message。
告警提取：从 message 中解析 Cisco 标准样式的严重级别（形如 %...-<sev>-... 里的 <sev> 数字）；
映射：sev<=2 → CRITICAL，sev==3 → ERROR，sev==4 → WARN，sev>=5 或无 sev → OK。
忽略列表：按 Ignore_alarm.yaml 的 message_contains 与 message_regex 过滤噪声。
输出：统计 scanned/matched，给出最严重等级与最多 10 条样例（时间戳、sev、sev 文本、等级、截断消息）。
################################################################################################################################################