---
phase: 2
plan: 1
title: "数据库扩展 — 新增 SewingDetail、RotocastItem 表和 format_type 字段"
wave: 1
depends_on: []
requirements: [DB-01, DB-02, DB-03]
files_modified:
  - server/services/db.js
autonomous: true
---

# Plan 01: 数据库扩展

## Objective

新增 SewingDetail 和 RotocastItem 表，QuoteVersion 添加 format_type 字段，为毛绒公仔数据存储做准备。

## Tasks

<task id="1">
<title>新增 SewingDetail 表</title>
<read_first>
- server/services/db.js
</read_first>
<action>
在 db.js 的 CREATE TABLE 区域（BodyAccessory 表之后）添加：

```sql
CREATE TABLE IF NOT EXISTS SewingDetail (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  version_id INTEGER NOT NULL REFERENCES QuoteVersion(id) ON DELETE CASCADE,
  product_name TEXT,
  fabric_name TEXT,
  position TEXT,
  cut_pieces INTEGER,
  usage_amount REAL,
  material_price_rmb REAL,
  price_rmb REAL,
  markup_point REAL DEFAULT 1.15,
  total_price_rmb REAL,
  sort_order INTEGER DEFAULT 0
);
```

字段说明：
- product_name: 产品名（如"11.25寸绿色鳄鱼女孩"）— 来自 R4 B列
- fabric_name: 布料名称（如"粉色亮光丝绒"）— C列
- position: 部位（如"头鬃毛"）— D列
- cut_pieces: 裁片数 — E列
- usage_amount: 用量 — F列
- material_price_rmb: 物料价RMB — G列
- price_rmb: 价钱RMB — H列
- markup_point: 码点 — I列（默认1.15）
- total_price_rmb: 总价钱RMB — J列
</action>
<acceptance_criteria>
- server/services/db.js 包含 `CREATE TABLE IF NOT EXISTS SewingDetail`
- SewingDetail 表包含 version_id, fabric_name, position, cut_pieces, usage_amount, material_price_rmb, price_rmb, markup_point, total_price_rmb 字段
</acceptance_criteria>
</task>

<task id="2">
<title>新增 RotocastItem 表</title>
<read_first>
- server/services/db.js
</read_first>
<action>
在 SewingDetail 之后添加：

```sql
CREATE TABLE IF NOT EXISTS RotocastItem (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  version_id INTEGER NOT NULL REFERENCES QuoteVersion(id) ON DELETE CASCADE,
  mold_no TEXT,
  name TEXT,
  output_qty INTEGER,
  usage_pcs INTEGER DEFAULT 1,
  unit_price_hkd REAL,
  total_hkd REAL,
  remark TEXT,
  sort_order INTEGER DEFAULT 0
);
```

字段说明（来自 3K报价 R20-R23）：
- mold_no: 模号（如"S01"）— A列
- name: 名称（如"毛绒公仔搪胶脸"）— B列
- output_qty: 出数 — C列
- usage_pcs: 用量pcs — D列
- unit_price_hkd: 单价HK$ — E列
- total_hkd: 合计HK$ — F列
- remark: 备注 — G列
</action>
<acceptance_criteria>
- server/services/db.js 包含 `CREATE TABLE IF NOT EXISTS RotocastItem`
- RotocastItem 表包含 mold_no, name, output_qty, usage_pcs, unit_price_hkd, total_hkd, remark 字段
</acceptance_criteria>
</task>

<task id="3">
<title>QuoteVersion 添加 format_type 字段</title>
<read_first>
- server/services/db.js
</read_first>
<action>
在 db.js 的 migration 区域（现有 PRAGMA table_info 检查之后）添加 format_type 字段迁移：

```javascript
// Migrate: add format_type to QuoteVersion
if (!existingCols.includes('format_type')) {
  db.prepare("ALTER TABLE QuoteVersion ADD COLUMN format_type TEXT DEFAULT 'injection'").run();
}
```

format_type 值：'injection'（注塑）或 'plush'（毛绒公仔）
</action>
<acceptance_criteria>
- server/services/db.js 包含 `ALTER TABLE QuoteVersion ADD COLUMN format_type`
- 默认值为 'injection'
</acceptance_criteria>
</task>

<task id="4">
<title>在 versions.js 路由中注册新表</title>
<read_first>
- server/routes/versions.js
</read_first>
<action>
在 SECTION_TABLES 映射中添加：

```javascript
'sewing-detail': 'SewingDetail',
'rotocast': 'RotocastItem',
```

在 ALL_SECTION_TABLES 映射中添加：

```javascript
sewing_details: 'SewingDetail',
rotocast_items: 'RotocastItem',
```
</action>
<acceptance_criteria>
- server/routes/versions.js 包含 `'sewing-detail': 'SewingDetail'`
- server/routes/versions.js 包含 `sewing_details: 'SewingDetail'`
- server/routes/versions.js 包含 `'rotocast': 'RotocastItem'`
- server/routes/versions.js 包含 `rotocast_items: 'RotocastItem'`
</acceptance_criteria>
</task>

## Verification

- 启动服务器，确认无 SQL 错误
- 查询 `PRAGMA table_info(SewingDetail)` 返回所有字段
- 查询 `PRAGMA table_info(RotocastItem)` 返回所有字段
- 查询 `PRAGMA table_info(QuoteVersion)` 包含 format_type 列

## must_haves

- DB-01: SewingDetail 表创建成功
- DB-02: RotocastItem 表创建成功
- DB-03: format_type 字段存在
