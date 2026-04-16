---
phase: 2
plan: 2
title: "毛绒公仔解析器 — 格式检测、3K报价主 sheet 和车缝明细解析"
wave: 2
depends_on: [1]
requirements: [PLUSH-01, PLUSH-02, PLUSH-03, PLUSH-04, PLUSH-06, PLUSH-07]
files_modified:
  - server/services/excel-parser.js
  - server/routes/import.js
autonomous: true
---

# Plan 02: 毛绒公仔解析器

## Objective

在 excel-parser.js 中实现格式检测和毛绒公仔专用解析函数，在 import.js 中添加新表插入逻辑。

## Tasks

<task id="1">
<title>实现格式检测函数 detectFormat()</title>
<read_first>
- server/services/excel-parser.js
</read_first>
<action>
在 excel-parser.js 中 detectLatestSheet 函数之后添加：

```javascript
function detectFormat(workbook) {
  const sheetNames = workbook.worksheets.map(ws => ws.name);
  const hasPlushIndicator = sheetNames.some(n =>
    n.includes('车缝明细') || n.includes('搪胶')
  );
  return hasPlushIndicator ? 'plush' : 'injection';
}
```

在 module.exports 中导出 detectFormat。
</action>
<acceptance_criteria>
- server/services/excel-parser.js 包含 `function detectFormat(workbook)`
- 函数检查 sheet 名称是否包含 '车缝明细' 或 '搪胶'
- module.exports 包含 detectFormat
</acceptance_criteria>
</task>

<task id="2">
<title>修改 parseMoldParts 支持不同起始行</title>
<read_first>
- server/services/excel-parser.js
</read_first>
<action>
修改 parseMoldParts 函数签名，添加 startRow 参数：

```javascript
function parseMoldParts(ws, startRow = 18) {
  const moldParts = [];
  let row = startRow;
  let sortOrder = 0;
  // ... rest of function unchanged
```

parseWorkbook 中调用时根据格式传入不同起始行：
- injection: parseMoldParts(ws, 18) — 默认值，不变
- plush: parseMoldParts(ws, 17)
</action>
<acceptance_criteria>
- parseMoldParts 函数签名包含 `startRow = 18` 参数
- 函数内 `let row = startRow;` 而不是硬编码 18
</acceptance_criteria>
</task>

<task id="3">
<title>实现搪胶件解析函数 parseRotocastItems()</title>
<read_first>
- server/services/excel-parser.js
- L21014毛绒公仔报价2026.03.25.xlsx（参考 3K报价 sheet R20-R23 结构）
</read_first>
<action>
在 excel-parser.js 中添加搪胶件解析函数：

```javascript
function parseRotocastItems(ws) {
  const items = [];
  // R20 is header: 模号 | 名称 | 出数 | 用量（pcs） | 单价（HK） | 合计 | 备注
  // R21+ is data until "搪胶件合计" row
  let row = 21;
  let sortOrder = 0;

  while (row <= 40) {
    const colA = strVal(ws.getCell(row, 1));
    const colE = strVal(ws.getCell(row, 5));

    // Stop on 合计 row
    if ((colA && colA.includes('合计')) || (colE && colE.includes('合计'))) break;

    const mold_no = strVal(ws.getCell(row, 1));
    const name = strVal(ws.getCell(row, 2));
    if (!mold_no && !name) { row++; continue; }

    const output_qty = numVal(ws.getCell(row, 3));
    const usage_pcs = numVal(ws.getCell(row, 4));
    const unit_price_hkd = numVal(ws.getCell(row, 5));
    const total_hkd = numVal(ws.getCell(row, 6));
    const remark = strVal(ws.getCell(row, 7));

    items.push({
      mold_no, name,
      output_qty: output_qty ? Math.round(output_qty) : null,
      usage_pcs: usage_pcs ? Math.round(usage_pcs) : null,
      unit_price_hkd, total_hkd, remark,
      sort_order: sortOrder++,
    });
    row++;
  }
  return items;
}
```
</action>
<acceptance_criteria>
- server/services/excel-parser.js 包含 `function parseRotocastItems(ws)`
- 函数从 R21 开始读取直到合计行
- 返回包含 mold_no, name, output_qty, usage_pcs, unit_price_hkd, total_hkd, remark 的数组
</acceptance_criteria>
</task>

<task id="4">
<title>实现车缝明细解析函数 parseSewingDetails()</title>
<read_first>
- server/services/excel-parser.js
- L21014毛绒公仔报价2026.03.25.xlsx（参考车缝明细 sheet 结构）
</read_first>
<action>
在 excel-parser.js 中添加车缝明细解析函数：

```javascript
function parseSewingDetails(workbook) {
  const wsNames = workbook.worksheets.map(ws => ws.name);
  const sewingSheet = wsNames.find(n => n.includes('车缝明细'));
  if (!sewingSheet) return [];

  const ws = workbook.getWorksheet(sewingSheet);
  const items = [];
  // R3: header (图片 | 名称 | 布料名称 | 部位 | 裁片数 | 用量 | 物料价RMB | 价钱RMB | 码点 | 总价钱RMB)
  // R4: product_name row (B4)
  // R5+: data rows until 合计 row

  let currentProductName = null;
  let sortOrder = 0;

  for (let row = 4; row <= 100; row++) {
    const colI = strVal(ws.getCell(row, 9));
    // Stop at 合计 row
    if (colI && colI.includes('合计')) break;

    const colB = strVal(ws.getCell(row, 2));
    const colC = strVal(ws.getCell(row, 3));
    const colD = strVal(ws.getCell(row, 4));

    // Product name row: B has value but C and D are empty
    if (colB && !colC && !colD) {
      currentProductName = colB;
      continue;
    }

    // Data row: must have fabric_name (C) or position (D)
    if (!colC && !colD) continue;

    const fabric_name = colC;
    const position = colD;
    const cut_pieces = numVal(ws.getCell(row, 5));
    const usage_amount = numVal(ws.getCell(row, 6));
    const material_price_rmb = numVal(ws.getCell(row, 7));
    const price_rmb = numVal(ws.getCell(row, 8));
    const markup_point = numVal(ws.getCell(row, 9));
    const total_price_rmb = numVal(ws.getCell(row, 10));

    items.push({
      product_name: currentProductName,
      fabric_name, position,
      cut_pieces: cut_pieces ? Math.round(cut_pieces) : null,
      usage_amount, material_price_rmb, price_rmb,
      markup_point: markup_point || 1.15,
      total_price_rmb,
      sort_order: sortOrder++,
    });
  }
  return items;
}
```
</action>
<acceptance_criteria>
- server/services/excel-parser.js 包含 `function parseSewingDetails(workbook)`
- 函数查找"车缝明细" sheet
- R4 检测为 product_name 行（B有值但C/D无值）
- R5+ 解析布料明细行直到合计
- 返回包含 product_name, fabric_name, position, cut_pieces, usage_amount, material_price_rmb, price_rmb, markup_point, total_price_rmb 的数组
</acceptance_criteria>
</task>

<task id="5">
<title>修改 parseWorkbook 添加格式分支</title>
<read_first>
- server/services/excel-parser.js
</read_first>
<action>
修改 parseWorkbook 函数：

1. 调用 detectFormat(workbook) 获取格式类型
2. 根据格式调用 parseMoldParts 传入不同 startRow
3. 毛绒公仔格式时额外调用 parseRotocastItems 和 parseSewingDetails
4. 在返回对象中添加 format_type, rotocastItems, sewingDetails 字段

```javascript
async function parseWorkbook(filePath) {
  // ... existing code ...
  const format = detectFormat(workbook);

  const header = parseHeader(ws);
  // ... existing product_no fallback logic ...

  const moldStartRow = format === 'plush' ? 17 : 18;
  const moldParts = parseMoldParts(ws, moldStartRow);

  // Plush-specific parsing
  const rotocastItems = format === 'plush' ? parseRotocastItems(ws) : [];
  const sewingDetails = format === 'plush' ? parseSewingDetails(workbook) : [];

  // ... existing costItems, summary, transport, electronics, painting parsing ...

  return {
    format_type: format,
    sheetName,
    // ... existing fields ...
    rotocastItems,
    sewingDetails,
  };
}
```
</action>
<acceptance_criteria>
- parseWorkbook 调用 detectFormat
- moldParts 使用格式相关的 startRow（plush=17, injection=18）
- 返回对象包含 format_type, rotocastItems, sewingDetails 字段
- 注塑格式时 rotocastItems=[], sewingDetails=[]
</acceptance_criteria>
</task>

<task id="6">
<title>修改 import.js 添加新表插入和 format_type</title>
<read_first>
- server/routes/import.js
</read_first>
<action>
1. QuoteVersion INSERT 添加 format_type 字段：
```javascript
const vr = db.prepare(
  `INSERT INTO QuoteVersion (product_id, version_name, source_sheet, date_code, quote_date, status, format_type, created_at, updated_at)
   VALUES (?, ?, ?, ?, ?, 'draft', ?, ?, ?)`
).run(product.id, versionName, data.sheetName, data.product.date_code, data.product.date_code, data.format_type, now, now);
```

2. 在 insertAll 事务中（ProductDimension 之前）添加搪胶件插入：
```javascript
// RotocastItem
if (data.rotocastItems && data.rotocastItems.length > 0) {
  const insertRoto = db.prepare(
    `INSERT INTO RotocastItem (version_id, mold_no, name, output_qty, usage_pcs, unit_price_hkd, total_hkd, remark, sort_order)
     VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`
  );
  for (const r of data.rotocastItems) {
    insertRoto.run(versionId, r.mold_no, r.name, r.output_qty, r.usage_pcs, r.unit_price_hkd, r.total_hkd, r.remark, r.sort_order);
  }
}
```

3. 添加车缝明细插入：
```javascript
// SewingDetail
if (data.sewingDetails && data.sewingDetails.length > 0) {
  const insertSew = db.prepare(
    `INSERT INTO SewingDetail (version_id, product_name, fabric_name, position, cut_pieces, usage_amount, material_price_rmb, price_rmb, markup_point, total_price_rmb, sort_order)
     VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
  );
  for (const s of data.sewingDetails) {
    insertSew.run(versionId, s.product_name, s.fabric_name, s.position, s.cut_pieces, s.usage_amount, s.material_price_rmb, s.price_rmb, s.markup_point, s.total_price_rmb, s.sort_order);
  }
}
```
</action>
<acceptance_criteria>
- import.js 的 QuoteVersion INSERT 包含 format_type 字段
- import.js 包含 `INSERT INTO RotocastItem` 语句
- import.js 包含 `INSERT INTO SewingDetail` 语句
- 插入逻辑在 insertAll 事务内
</acceptance_criteria>
</task>

<task id="7">
<title>验证：导入 L21014 并检查所有数据</title>
<read_first>
- server/services/excel-parser.js
- server/routes/import.js
</read_first>
<action>
运行验证脚本：

```javascript
const {parseWorkbook} = require('./server/services/excel-parser');
(async () => {
  const data = await parseWorkbook('L21014毛绒公仔报价2026.03.25.xlsx');
  console.log('format_type:', data.format_type);
  console.log('product_no:', data.product.product_no);
  console.log('sheetName:', data.sheetName);
  console.log('moldParts:', data.moldParts.length);
  if (data.moldParts.length > 0) {
    console.log('First moldPart:', data.moldParts[0]);
  }
  console.log('rotocastItems:', data.rotocastItems.length);
  for (const r of data.rotocastItems) {
    console.log('  Rotocast:', r.mold_no, r.name, r.unit_price_hkd, r.total_hkd);
  }
  console.log('sewingDetails:', data.sewingDetails.length);
  for (const s of data.sewingDetails.slice(0, 5)) {
    console.log('  Sewing:', s.fabric_name, s.position, s.total_price_rmb);
  }
})();
```

同时验证 47712 注塑格式不受影响：

```javascript
const data2 = await parseWorkbook('47712 本厂报价明细20260310 （电子加价改内部码点）.xlsx');
console.log('47712 format:', data2.format_type);
console.log('47712 moldParts:', data2.moldParts.length);
console.log('47712 rotocast:', data2.rotocastItems.length);
console.log('47712 sewing:', data2.sewingDetails.length);
```
</action>
<acceptance_criteria>
- L21014 format_type 输出 'plush'
- L21014 moldParts >= 1（包胶左右手 PVC 22g）
- L21014 rotocastItems >= 2（搪胶脸 + 搪胶脚）
- L21014 sewingDetails >= 20（布料明细行）
- 47712 format_type 输出 'injection'
- 47712 moldParts 输出 20（不变）
- 47712 rotocastItems 输出 0
- 47712 sewingDetails 输出 0
</acceptance_criteria>
</task>

## Verification

- 两种格式均正确检测
- L21014 的 MoldPart R17 行正确解析
- 搪胶件数据（S01 搪胶脸 3.77、S02 搪胶脚 3.43）正确提取
- 车缝明细数据（粉色亮光丝绒、白色软戟绒等布料）正确提取
- 47712 注塑格式无回归

## must_haves

- PLUSH-01: 毛绒公仔格式正确识别
- PLUSH-02: MoldPart R17 数据解析（PVC 22g 包胶件）
- PLUSH-03: 搪胶件数据正确提取
- PLUSH-04: 车缝明细数据正确提取
- PLUSH-06: 3K报价主 sheet 成本区域解析
- PLUSH-07: 所有数据正确映射到数据库表
