# 办公智盾资产分组 API 契约

本文件只记录接口形态和程序执行规则，不保存 Cookie、Token 等认证信息。

## 通用

- Base URL: `https://IP:Prot`
- 请求头:
  - `Accept: application/json, text/plain, */*`
  - `Content-Type: application/json`，仅 POST 必需
  - `Cookie: <运行时提供>`，不要写入模板或代码
  - `Referer: https://IP:Prot/ues/base/terminals`
- HTTPS 证书如为内网自签，执行器需要支持忽略证书校验的开关。
- 成功判断以 `code === 0` 为主，部分接口还会带 `success: true`。

## 查询完整分组树

```http
GET /groups
```

无请求体。

返回结构:

```json
{
  "code": 0,
  "msg": "",
  "data": [
    {
      "guid": "0f2736c478d44ecd99284c7393dfb18e",
      "name": "全部终端",
      "parentGuid": "-1",
      "children": []
    }
  ]
}
```

程序用途:

- 建立 `完整路径 -> groupGuid` 索引。
- 校验模板中的父路径是否存在。
- 避免重复创建已有分组。

## 查询指定分组及子组

```http
GET /groups/{groupGuid}
```

无请求体。

返回结构:

```json
{
  "code": 0,
  "msg": "",
  "data": {
    "guid": "0f81104556f34db5bc012a460afdb21e",
    "name": "门急诊",
    "parentGuid": "0f2736c478d44ecd99284c7393dfb18e",
    "children": []
  }
}
```

程序用途:

- 可作为增量校验父分组、子分组的备用接口。
- 如果 `GET /groups` 已经返回完整树，优先使用完整树，减少请求次数。

## 新建分组

```http
POST /groups
```

请求体:

```json
{
  "description": "",
  "lockChildGroupPolicy": false,
  "name": "测试",
  "parentGuid": "28b8e9390c4e4038b93503d41de2b6f1",
  "policyType": "",
  "autoGroupRule": {
    "ruleItemsBatch": [],
    "reGroup": 0,
    "syncBind": false
  }
}
```

返回结构:

```json
{
  "code": 0,
  "msg": "保存成功",
  "data": {
    "guid": "cb3c1089dd0d4a87b845bf721fed7310",
    "name": "测试",
    "parentGuid": "28b8e9390c4e4038b93503d41de2b6f1",
    "children": [],
    "terminals": []
  }
}
```

程序用途:

- 按模板 `分组树.完整分组路径` 从浅到深创建缺失分组。
- 创建前必须先通过路径索引确认父分组 GUID。
- 创建成功后要把返回的 `data.guid` 写回内存索引，后续子组和移动终端会用到。

## 查询终端

```http
POST /terminals/query?curPage=1&pageSize=200
```

请求体，按分组查:

```json
{
  "sort": "",
  "groupGuidList": ["0f2736c478d44ecd99284c7393dfb18e"],
  "subProductCode": "DAS-UES-SMP",
  "includeChild": true,
  "dumbTerminal": 0
}
```

请求体，按 MAC 查:

```json
{
  "sort": "",
  "groupGuidList": ["2ab8a04c63b94cc0934c50c87f2584f5"],
  "subProductCode": "DAS-UES-SMP",
  "includeChild": true,
  "dumbTerminal": 0,
  "mac": "70:CF:49:AF:09:BF"
}
```

返回结构:

```json
{
  "code": 0,
  "msg": "",
  "data": {
    "page": 1,
    "pageSize": 200,
    "totalRow": 474,
    "list": [
      {
        "guid": "369bf6a6-838e-201c-a679-cde2c14d9dec",
        "terminalName": "3B3FYSB088",
        "hostName": "3B3FYSB088",
        "groupGuid": "2ab8a04c63b94cc0934c50c87f2584f5"
      }
    ]
  },
  "success": true
}
```

程序用途:

- 优先按模板 `终端清单.MAC地址` 查询。
- 一个 MAC 返回多个终端 GUID 时，不视为异常，全部纳入移动计划。
- 查询结果需要分页取完，直到 `page * pageSize >= totalRow` 或返回列表为空。

## 移动终端

```http
POST /terminals/move
```

请求体:

```json
{
  "terminalGuidList": ["369bf6a6-838e-201c-a679-cde2c14d9dec"],
  "groupGuidList": [],
  "targetGroupGuid": "2ab8a04c63b94cc0934c50c87f2584f5"
}
```

返回结构:

```json
[
  {
    "code": 0,
    "msg": "编辑成功",
    "data": null,
    "success": true
  }
]
```

程序用途:

- 对每个目标分组聚合终端 GUID 后批量移动。
- 同一个 MAC 命中多个终端 GUID 时，把所有 GUID 一起移动到模板指定的目标分组。
- 同一个 `terminalGuid + targetGroupGuid` 重复出现时去重。
- 如果同一个终端 GUID 在同一次计划中指向多个目标分组，标记为冲突，不自动移动。

