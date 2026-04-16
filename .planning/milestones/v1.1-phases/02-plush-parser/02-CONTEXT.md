# Phase 2: 毛绒公仔解析引擎 - Context

**Gathered:** 2026-03-28
**Status:** Ready for planning

<domain>
## Phase Boundary

解析毛绒公仔报价 Excel（如 L21014）的所有数据并存入数据库。同时确保注塑产品格式（如 47712）不受影响。

</domain>

<decisions>
## Implementation Decisions

### 格式检测
- **D-01:** 通过内容探测区分格式：检查 workbook 是否包含"车缝明细"或"搪胶" sheet → 毛绒公仔格式；否则 → 注塑格式
- **D-02:** QuoteVersion 新增 format_type 字段（'injection' | 'plush'）标记格式类型

### 数据表映射
- **D-03:** 新建 SewingDetail 表存车缝明细数据（布料名称、部位、裁片数、用量、物料价RMB、码点、总价RMB）
- **D-04:** 新建 RotocastItem 表存搪胶件数据（模号、名称、出数、用量pcs、单价HK$、合计HK$、备注）
- **D-05:** 毛绒公仔的 MoldPart 同样使用现有 MoldPart 表（R16 header, R17 数据行，结构相同）

### 解析范围
- **D-06:** 只解析 3K报价主 sheet + 车缝明细 sheet，其他子 sheet（五金、吊咭、贴纸、PE袋等）跳过
- **D-07:** 3K报价主 sheet 解析内容：
  - R1: 产品编号（B1）
  - R2-R10: 材料价格表 + 机型价格表（与注塑格式相同）
  - R11-R15: 汇率参数（与注塑格式相同）
  - R16-R18: MoldPart（header 在 R16 不是 R17，数据从 R17 开始）
  - R20-R23: 搪胶件（模号、名称、出数、用量、单价、合计）
  - R25-R53: 成本明细（料价、人工、五金、搪胶、车缝件、包装等各项成本）
  - R57-R69: 运费/码点/总价计算

### Claude's Discretion
- 3K报价主 sheet 的成本明细行（R25-R53）如何映射到现有数据表（HardwareItem、PackagingItem、laborItems 等）
- parseMoldParts 如何适配两种格式的不同起始行（注塑 R18，毛绒 R17）

</decisions>

<canonical_refs>
## Canonical References

**Downstream agents MUST read these before planning or implementing.**

### 现有解析器
- `server/services/excel-parser.js` — 当前注塑格式解析逻辑，新代码需在此扩展
- `server/routes/import.js` — 导入事务逻辑，需添加新表插入

### 数据库
- `server/services/db.js` — 现有 13 表 schema，需新增 SewingDetail 和 RotocastItem

### 参考 Excel
- `L21014毛绒公仔报价2026.03.25.xlsx` — 毛绒公仔报价样本文件

</canonical_refs>

<code_context>
## Existing Code Insights

### Reusable Assets
- `cellVal()`, `numVal()`, `strVal()` — Excel 单元格值提取 helpers，可直接复用
- `parseHeader()` — 材料价格表和汇率参数解析（R1-R15），两种格式相同可直接复用
- `parseCostItems()` — 成本项解析，可能需适配毛绒公仔的行范围

### Established Patterns
- 每个解析函数独立（parseHeader, parseMoldParts, parseCostItems 等）
- import.js 在一个事务中插入所有数据
- 所有子表通过 version_id 外键关联

### Integration Points
- `parseWorkbook()` 是唯一入口，需在此添加格式检测分支
- `insertAll` 事务块需添加 SewingDetail 和 RotocastItem 插入

</code_context>

<specifics>
## Specific Ideas

- 毛绒公仔 3K报价 sheet 的 R16 header 与注塑 R17 header 格式相同（模号/名称/料型/料重...），只是行号差 1
- 车缝明细 sheet 的 R1 是合并单元格标题，R3 是 header（图片/名称/布料名称/部位/裁片数/用量/物料价/价钱/码点/总价），R4 开始是数据（R4 是产品名称行，R5+ 是明细行）
- 成本汇总行（R25-R53）是名称+金额的简单结构，不需要独立表，可映射到现有 laborItems/hardwareItems/packagingItems

</specifics>

<deferred>
## Deferred Ideas

- 其他子 sheet（五金、吊咭、贴纸、PE袋等）的详细解析 — 目前主 sheet 已有汇总数据
- 多种毛绒公仔格式变体支持（如不含搪胶的纯布偶）

</deferred>

---

*Phase: 02-plush-parser*
*Context gathered: 2026-03-28*
