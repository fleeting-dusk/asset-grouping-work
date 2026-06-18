# 资产分组脚本使用说明

这个工具用于按 Excel 模板自动规划办公智盾终端分组，支持本地解析、平台 dry-run、正式建组和逐台移动终端。

## 文件

- `asset_grouping_runner.py`: Python 主执行脚本。
- `资产分组导入模板.xlsx`: 本地真实模板，填写分组树、科室映射和终端清单。此文件不会提交到 Git。
- `资产分组导入模板.example.xlsx`: 脱敏模板示例，可提交到 Git。
- `asset_grouping_config.json`: 本地真实配置，填写 API 地址、Cookie、模板路径等。此文件不会提交到 Git。
- `asset_grouping_config.example.json`: 脱敏配置示例。
- `api_contract.md`: 已确认的接口契约。
- `asset-grouping-run-latest.json`: 最近一次运行的 JSON 明细报告。
- `asset-grouping-summary-latest.txt`: 最近一次运行的人类可读摘要。

## 快速开始

在仓库根目录执行：

```powershell
python -m pip install openpyxl
Copy-Item .\asset_grouping_config.example.json .\asset_grouping_config.json
```

然后把真实模板放到仓库根目录，文件名默认使用：

```text
资产分组导入模板.xlsx
```

编辑 `asset_grouping_config.json`：

```json
{
  "template_path": "资产分组导入模板.xlsx",
  "output_dir": ".",
  "api_base_url": "",
  "internal_api_base_url": "https://<internal-host>:<port>",
  "external_api_base_url": "https://<external-host>:<port>",
  "query_scope_path": "全部终端",
  "internal_cookie": "",
  "external_cookie": "",
  "default_network": "内网",
  "strict_tls": false,
  "keep_history": false
}
```

`template_path` 和 `output_dir` 支持相对路径。相对路径会以 `asset_grouping_config.json` 所在目录为基准解析，方便同事克隆后直接使用。

`api_base_url` 是旧配置兼容字段，建议保持为空，实际使用 `internal_api_base_url` 和 `external_api_base_url`。

## 交互界面

直接运行脚本会进入交互菜单：

```powershell
python .\asset_grouping_runner.py
```

菜单支持：

- 本地解析模板
- 平台 dry-run
- 正式执行
- 查看/修改配置
- 查看最近报告
- 比对导出的分组 Excel
- 导出不合法 MAC
- 导出未映射责任科室
- 导出未查询到终端

## 常用命令

本地解析，不访问平台：

```powershell
python .\asset_grouping_runner.py --local-only
```

补全配置文件：

```powershell
python .\asset_grouping_runner.py --init-config
```

平台 dry-run，查询平台但不新建、不移动：

```powershell
python .\asset_grouping_runner.py --run
```

正式执行：

```powershell
python .\asset_grouping_runner.py --execute
```

比对导出的分组 Excel：

```powershell
python .\asset_grouping_runner.py --compare-export .\分组数据.xlsx
```

导出不合法 MAC：

```powershell
python .\asset_grouping_runner.py --export-invalid-mac
```

导出未映射责任科室：

```powershell
python .\asset_grouping_runner.py --export-unmapped-depts
```

这个检查只解析模板，不访问平台。它会找出“手动目标分组路径为空、设备责任科室有值、但科室映射未命中”的终端，并按责任科室汇总，方便补 `科室映射`。

导出平台未查询到的终端：

```powershell
python .\asset_grouping_runner.py --export-not-found-terminals
```

## Cookie 和地址

配置文件只保留内外网专用 Cookie：

```json
{
  "internal_api_base_url": "https://<internal-host>:<port>",
  "external_api_base_url": "https://<external-host>:<port>",
  "internal_cookie": "",
  "external_cookie": ""
}
```

也可以用环境变量临时覆盖：

```powershell
$env:UES_INTERNAL_API_BASE_URL='https://<internal-host>:<port>'
$env:UES_EXTERNAL_API_BASE_URL='https://<external-host>:<port>'
$env:UES_INTERNAL_COOKIE='JSESSIONID=...; token=...; language=zh-CN'
$env:UES_EXTERNAL_COOKIE='JSESSIONID=...; token=...; language=zh-CN'
```

兼容环境变量：

- `UES_API_BASE_URL` 或 `OFFICE_SHIELD_API_BASE_URL`
- `UES_INTERNAL_API_BASE_URL` 或 `OFFICE_SHIELD_INTERNAL_API_BASE_URL`
- `UES_EXTERNAL_API_BASE_URL` 或 `OFFICE_SHIELD_EXTERNAL_API_BASE_URL`
- `UES_COOKIE` 或 `OFFICE_SHIELD_COOKIE`
- `UES_INTERNAL_COOKIE` 或 `OFFICE_SHIELD_INTERNAL_COOKIE`
- `UES_EXTERNAL_COOKIE` 或 `OFFICE_SHIELD_EXTERNAL_COOKIE`

外网 Cookie 必须明确填写 `external_cookie` 或外网环境变量，避免误拿内网登录态访问外网平台。

## 内外网分流

脚本会读取 `终端清单` 的 `网络区域` 字段：

- `内网`: 使用 `internal_api_base_url` 和 `internal_cookie`
- `外网`: 使用 `external_api_base_url` 和 `external_cookie`
- 空白或无法识别: 使用 `default_network`，默认 `内网`

## 安全原则

- Cookie、Token 不写入模板、不写入脚本、不写入报告。
- 真实 `asset_grouping_config.json`、真实 `资产分组导入模板.xlsx`、运行报告和导出结果已加入 `.gitignore`。
- 默认不执行平台变更。
- 正式执行必须同时满足：
  - 模板 `执行配置` 中 `dry_run` 为 `FALSE`
  - 命令行加 `--execute`，或在交互菜单选择正式执行并输入确认
- 正式移动终端时逐台调用 `/terminals/move`，每台移动后按 MAC 回查复核。

## 分组计划说明

报告里的 `路径保障` 不是手工启用分组行数。脚本会保障本次终端目标路径及其父路径都存在，例如目标路径为：

```text
全部终端>门急诊>门诊>门诊部01护理单元
全部终端>门急诊>急诊>医疗单元
```

则会同时保障 `全部终端`、`门急诊`、`门诊`、`急诊` 等父路径。已存在的分组会标记为 `exists`，不会重复创建。

默认情况下，只有实际被终端指向的目标分组及其父路径会参与保障。只有当模板 `执行配置` 中 `create_declared_groups_without_terminals` 设置为 `TRUE` 时，才会预建启用但暂无终端指向的分组。

分组创建或补齐时会维护平台 `description`：新建分组会写入 `分组树.分组描述` 和 `备注`；已有分组如果平台描述为空，会补入模板描述和备注；如果平台已有描述，只会在末尾追加缺少的模板备注，不覆盖原描述，重复运行不会重复追加同一条备注。

## 终端规则

- 同一个 MAC 查询到多个终端 GUID 时，脚本会全部移动到模板指定目标分组。
- 如果某个终端 GUID 在同一次计划中被指向多个目标分组，脚本会标记冲突并跳过该 GUID。
- 已经在目标分组的终端会标记为 `already_in_target` 并跳过移动。
- 平台查不到 MAC 时不会移动，会写入报告；可用 `--export-not-found-terminals` 导出。

## MAC 格式

模板支持以下格式：

```text
AA-BB-CC-DD-EE-FF
AA:BB:CC:DD:EE:FF
AABBCCDDEEFF
```

脚本查询平台前会统一转换为 `AA:BB:CC:DD:EE:FF`。不能规范化为 12 位十六进制的 MAC 会在模板校验阶段报错，也可以用 `--export-invalid-mac` 单独导出。

## 报告

默认只更新最新报告：

```text
asset-grouping-run-latest.json
asset-grouping-summary-latest.txt
```

需要保留每次运行的历史报告时添加：

```powershell
--keep-history
```
