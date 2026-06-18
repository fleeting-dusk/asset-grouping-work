# 资产分组脚本使用说明

## 文件

- `资产分组导入模板.xlsx`: 人工整理分组、科室映射和终端清单。
- `资产分组导入模板.example.xlsx`: 可提交到 GitHub 的脱敏模板示例。
- `asset_grouping_runner.py`: Python 主执行脚本。
- `asset_grouping_config.json`: 本地配置文件，可放 API 地址、Cookie、模板路径等。
- `asset_grouping_config.example.json`: 配置示例。
- `api_contract.md`: 已确认的办公智盾接口契约。
- `asset-grouping-run-latest.json`: 最近一次运行报告，记录分组计划、终端查询、逐台移动任务、复查结果、错误和警告。
- `asset-grouping-summary-latest.txt`: 最近一次运行的人类可读摘要，适合快速核对内外网分布、分组动作、终端匹配数、移动复查数。

## JSON 文件是什么

- `asset_grouping_config.json` 是配置文件，给脚本读取。可以填写 Cookie、接口地址、模板路径、输出目录等。
- `asset-grouping-run-latest.json` 是运行报告，给人核对结果。它记录本次解析出的分组计划、目标分组使用统计、终端查询结果、逐台移动任务、复查结果、错误和警告。
- `asset-grouping-summary-latest.txt` 是摘要报告，适合快速核对内外网分布、分组动作、终端匹配数、移动复查数。
- 默认只保留最新报告，不会不断堆历史文件；需要历史报告时加 `--keep-history`。

报告里的 `路径保障` 不是你手工启用的分组数量。默认情况下，脚本只会保障“本次终端实际目标路径”以及这些路径的父级，例如本次终端实际指向：

```text
全部终端>门急诊>门诊>门诊部01护理单元
全部终端>门急诊>急诊>医疗单元
```

它会同时保障 `全部终端`、`门急诊`、`门诊`、`急诊` 这些父路径存在，所以会显示 `启用分组行 2 行，路径保障 6 项`。平台 dry-run/正式执行时，已存在的父级只会标记为 `exists`，不会重复创建。

`分组树` 全部启用并不代表全部创建。它主要提供分组描述、是否允许新建等属性。只有当 `执行配置` 里的 `create_declared_groups_without_terminals` 改成 `TRUE` 时，脚本才会预建分组树或映射表中启用但暂无终端指向的路径；默认 `FALSE` 更适合正式执行，避免内网/外网误建不需要的分组。

如果某个“不该建”的分组仍出现在计划里，通常说明有终端被 `科室映射` 或 `手动目标分组路径` 指到了它。可以在 `asset-grouping-run-latest.json` 的 `localPlan.targetGroupUsage` 里查看每个目标分组的内外网终端数量和样例源行。

## 安全原则

- 正式执行移动终端时，脚本会逐台提交 `/terminals/move`，不会把一批 GUID 一次性发出去。
- 每台终端移动后会按 MAC 回查平台，确认同一个终端 GUID 的 `groupGuid` 已经变成目标分组；没生效会写入错误报告。
- 控制台会显示分组保障、终端查询、移动复查三个阶段的进度条，进度条完成后才会输出最终报告路径。
- 真正执行必须双确认：
  - 模板 `执行配置` 中 `dry_run` 改为 `FALSE`
  - 命令行加 `--execute`，或在交互菜单选择“正式执行”并输入确认

## 内外网分流

脚本会读取 `终端清单` 里的 `网络区域`：

- `内网` 连接 `https://IP:Prot`
- `外网` 连接 `https://IP:Prot`
- 空白或无法识别时，按配置里的 `default_network` 处理，默认是 `内网`

配置文件只保留内外网专用 Cookie：

```json
{
  "internal_api_base_url": "https://IP:Prot",
  "external_api_base_url": "https://IP:Prot",
  "internal_cookie": "",
  "external_cookie": ""
}
```

旧版配置里的通用 `cookie` 字段只会在脚本保存配置时自动迁移到 `internal_cookie`，之后不再写回配置文件。外网 Cookie 必须明确填写 `external_cookie`，避免误拿内网登录态去访问外网平台。命令行 `--cookie` 和环境变量 `UES_COOKIE` 仍保留为临时覆盖方式。

## 本地解析，不访问平台

```powershell
'python .\asset_grouping_runner.py --local-only' 
```

## 终端交互界面

直接运行脚本会进入交互界面：

```powershell
'python .\asset_grouping_runner.py' 
```

菜单里可以：

- 本地解析模板
- 平台 dry-run
- 正式执行
- 查看/修改配置
- 查看最近报告
- 比对导出的分组 Excel
- 导出不合法 MAC
- 导出未查询到终端

## 初始化/补全配置文件

```powershell
'python .\asset_grouping_runner.py --local-only --init-config'
```

配置文件路径：

```text
asset-grouping-work\asset_grouping_config.json
```

## 平台 dry-run，查平台但不新建/不移动

```powershell
$env:UES_COOKIE='JSESSIONID=...; token=...; language=zh-CN'

'python .\asset_grouping_runner.py --run' 
```

这个模式会：

- 查询 `/groups` 建立分组路径索引。
- 判断哪些分组已存在、哪些会新建。
- 按 `MAC地址` 调 `/terminals/query` 查询终端 GUID。
- 生成逐台移动任务，但不调用 `/groups` 新建，也不调用 `/terminals/move`。

## 比对导出的分组 Excel

```powershell
& 'C:\Users\Heoflare\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' `
  'D:\tmp\asset-grouping-work\asset_grouping_runner.py' `
  --compare-export 'C:\Users\Heoflare\Downloads\分组数据20260617124022.xlsx'
```

这个功能会实时查询平台 `/groups`，再和导出的 Excel 分组路径比对，输出：

- 平台有、导出没有
- 导出有、平台没有
- 名称不一致

比对报告会写到：

```text
D:\tmp\asset-grouping-work\group-export-compare-latest.json
```

## 正式执行

确认平台 dry-run 报告无误后：

1. 打开 `资产分组导入模板.xlsx`
2. 把 `执行配置` 里的 `dry_run` 改成 `FALSE`
3. 运行：

```powershell
$env:UES_COOKIE='JSESSIONID=...; token=...; language=zh-CN'

'python .\asset_grouping_runner.py --execute'
```

正式执行时：

- 已经在目标分组的终端会标记为 `already_in_target` 并跳过。
- 需要移动的终端会按 1 台 1 次提交，再立刻复查。
- 报告摘要里的 `移动复查` 会显示 `verified`、`not_effective`、`not_found_after_move`、`verify_failed` 等结果。
- 如果平台接口返回成功但复查没到位，本次运行状态会变成 `completed_with_errors`，你可以用同一份模板补跑未生效项。

## 多资产 / 多 GUID 规则

同一个 MAC 如果在平台查询到多个终端 GUID，不算异常。脚本会把这些 GUID 全部移动到模板指定目标分组。

如果同一个终端 GUID 在同一次计划中被不同模板行指向了不同目标分组，脚本会标记冲突并跳过该 GUID。

## MAC 地址格式

模板里 `MAC地址` 支持常见写法，例如：

- `AA-BB-CC-DD-EE-FF`
- `AA:BB:CC:DD:EE:FF`
- `AABBCCDDEEFF`

脚本查询平台前会统一转换为 `AA:BB:CC:DD:EE:FF`。如果不能规范化为 12 位十六进制，会在模板校验阶段报错并停止。

可以单独导出不合法 MAC 清单：

```powershell
'python .\asset_grouping_runner.py --export-invalid-mac'
```

默认输出：

```text
.\asset-grouping-work\invalid-mac-latest.xlsx
```

## 导出未查询到终端

这个功能读取最近一次平台 dry-run 或正式执行报告，把 `/terminals/query` 返回 `not_found` 的终端导出成 Excel。它不能只靠本地解析判断，必须先访问过平台。

```powershell
'python .\asset_grouping_runner.py --export-not-found-terminals'
```

默认输出：

```text
.\asset-grouping-work\not-found-terminals-latest.xlsx
```

导出字段包含源行、网络、资产编号、原 MAC、查询 MAC、主机名、责任科室、目标分组路径、目标来源、位置和说明。

## 报告保留

默认只更新 `asset-grouping-run-latest.json`。如果需要保留每次运行的历史报告，命令里加：

```powershell
--keep-history
```
