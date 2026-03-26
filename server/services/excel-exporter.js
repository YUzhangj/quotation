/**
 * Excel Export Service — Template-Driven
 * Loads VQ-template.xlsx and fills data from DB, preserving all TOMY formatting and formulas.
 */
const ExcelJS = require('exceljs');
const path = require('path');
const { getDb } = require('./db');

const TEMPLATE_PATH = path.join(__dirname, '../templates/VQ-template.xlsx');

// ─── Load all version data from DB ───────────────────────────────────────────

function loadData(versionId) {
  const db = getDb();
  const version = db.prepare('SELECT * FROM QuoteVersion WHERE id = ?').get(versionId);
  if (!version) throw new Error(`Version ${versionId} not found`);
  const product  = db.prepare('SELECT * FROM Product WHERE id = ?').get(version.product_id);
  const params   = db.prepare('SELECT * FROM QuoteParams WHERE version_id = ?').get(versionId) || {};
  return {
    version, product, params,
    moldParts:      db.prepare('SELECT * FROM MoldPart     WHERE version_id = ? ORDER BY sort_order').all(versionId),
    hardwareItems:  db.prepare('SELECT * FROM HardwareItem WHERE version_id = ? ORDER BY sort_order').all(versionId),
    electronicItems:db.prepare('SELECT * FROM ElectronicItem WHERE version_id = ? ORDER BY sort_order').all(versionId),
    packagingItems: db.prepare('SELECT * FROM PackagingItem WHERE version_id = ? ORDER BY sort_order').all(versionId),
    paintingDetail: db.prepare('SELECT * FROM PaintingDetail WHERE version_id = ?').get(versionId) || {},
    transportConfig:db.prepare('SELECT * FROM TransportConfig WHERE version_id = ?').get(versionId) || {},
    moldCost:       db.prepare('SELECT * FROM MoldCost WHERE version_id = ?').get(versionId) || {},
    productDim:     db.prepare('SELECT * FROM ProductDimension WHERE version_id = ?').get(versionId) || {},
    materialPrices: db.prepare('SELECT * FROM MaterialPrice WHERE version_id = ? ORDER BY id').all(versionId),
    machinePrices:  db.prepare('SELECT * FROM MachinePrice  WHERE version_id = ? ORDER BY id').all(versionId),
  };
}

// ─── Cell helper — only write to data cells, skip formula cells ───────────────

function setVal(ws, row, col, value) {
  const cell = ws.getCell(row, col);
  // Never overwrite formula cells — they calculate automatically
  if (cell.value && typeof cell.value === 'object' && cell.value.formula) return;
  // Guard against NaN (invalid XML)
  if (typeof value === 'number' && isNaN(value)) value = null;
  cell.value = (value === undefined) ? null : value;
}

// Clear a range of data columns (skip formula columns)
function clearRows(ws, startRow, endRow, dataCols) {
  for (let r = startRow; r <= endRow; r++) {
    for (const c of dataCols) {
      const cell = ws.getCell(r, c);
      if (!(cell.value && typeof cell.value === 'object' && cell.value.formula)) {
        cell.value = null;
      }
    }
  }
}

// ─── Fill Vendor Quotation sheet ─────────────────────────────────────────────

function fillVQ(ws, d) {
  const { version, product, params, packagingItems, productDim, transportConfig } = d;

  // ── Header (rows 2–5) ──────────────────────────────────────────────────────
  setVal(ws, 2, 3, product?.vendor || '');
  setVal(ws, 2, 8, params.prepared_by || '');
  setVal(ws, 3, 3, product?.item_no || '');
  setVal(ws, 3, 8, version.quote_date ? new Date(version.quote_date) : '');
  setVal(ws, 4, 3, product?.item_desc || '');
  setVal(ws, 4, 8, version.version_name || '');

  // ── Section A (rows 11–16): Body Cost — row 11 formula stays (='BCD'!F23) ─
  // Just update part no and description for main body row
  if (product?.item_no) {
    setVal(ws, 11, 1, product.item_no + '-00');
    setVal(ws, 11, 2, product.item_desc || '');
    setVal(ws, 11, 5, 2500);  // default MOQ
    setVal(ws, 11, 6, 1);     // usage 1 per toy
    // G11 = ='Body Cost Breakdown'!F23 — do NOT overwrite (formula)
  }
  // Clear accessory rows 12–16 data
  clearRows(ws, 12, 16, [1, 2, 5, 6, 7]);

  // ── Section B (rows 23–35): Packaging ──────────────────────────────────────
  const PKG_START = 23, PKG_END = 35;
  // Clear first
  clearRows(ws, PKG_START, PKG_END, [1, 2, 3, 5, 6, 7]);
  // Fill packaging items
  const moq = params.moq_default || 2500;
  packagingItems.slice(0, PKG_END - PKG_START + 1).forEach((item, i) => {
    const r = PKG_START + i;
    setVal(ws, r, 2, item.name || '');
    setVal(ws, r, 3, item.remark || '');    // specifications / remark
    setVal(ws, r, 5, moq);
    setVal(ws, r, 6, item.quantity || 1);
    setVal(ws, r, 7, parseFloat(item.new_price) || 0);
    // H col = formula =ROUND(F*G,2) — not touched
  });
  // Mark Up row 36: G36 = markup%
  const pkgMarkup = parseFloat(params.markup_packaging) || 0.12;
  setVal(ws, 36, 7, pkgMarkup);

  // ── Section D (row 52): Master Carton ──────────────────────────────────────
  if (productDim) {
    setVal(ws, 52, 2, parseFloat(productDim.carton_l_inch) || null);
    setVal(ws, 52, 3, parseFloat(productDim.carton_w_inch) || null);
    setVal(ws, 52, 4, parseFloat(productDim.carton_h_inch) || null);
    setVal(ws, 52, 6, parseInt(productDim.pcs_per_carton) || null);
    setVal(ws, 52, 7, parseFloat(productDim.carton_price) || null);
  }

  // ── Section E (row 58): Transport cost parameters ──────────────────────────
  // Template: F58=Ex-Factory cost/CuFt, G58=FOB FCL cost/CuFt, H58=FOB LCL cost/CuFt
  // C58 = formula =B52*C52*D52/1728*F52 (CuFt per toy, don't touch)
  if (transportConfig) {
    setVal(ws, 58, 6, parseFloat(transportConfig.hk_10t_cost) || 0.5);   // Ex-Factory
    setVal(ws, 58, 7, parseFloat(transportConfig.yt_40_cost)  || 4.3);   // FOB FCL
    setVal(ws, 58, 8, parseFloat(transportConfig.hk_40_cost)  || 15.85); // FOB LCL
  }
}

// ─── Fill Body Cost Breakdown sheet ──────────────────────────────────────────

function fillBCD(ws, d) {
  const { version, product, params, moldParts, hardwareItems, electronicItems,
          paintingDetail, materialPrices } = d;

  // ── Header (row 7) ─────────────────────────────────────────────────────────
  setVal(ws, 7, 2, product?.item_desc || '');
  setVal(ws, 7, 3, '0');  // body cost revision
  setVal(ws, 7, 4, product?.vendor || '');
  setVal(ws, 7, 6, params.prepared_by || '');
  setVal(ws, 7, 8, version.quote_date ? new Date(version.quote_date) : '');

  // ── Summary section markup % (rows 14–22, col E) ──────────────────────────
  const bodyMkup = parseFloat(params.markup_body) || 0.18;
  for (const r of [14, 15, 16, 19, 20, 21, 22]) {
    setVal(ws, r, 5, bodyMkup);
  }
  setVal(ws, 17, 5, 0); // D (Expensive Components) — markup 0%

  // ── Section A: Raw Material (rows 31–34 = ABS, PP, PCTG, PVC) ─────────────
  // Group mold parts by material type → sum weight_g
  const matWeight = {};
  moldParts.forEach(p => {
    const m = (p.material || '').toUpperCase().trim();
    if (m) matWeight[m] = (matWeight[m] || 0) + (parseFloat(p.weight_g) || 0);
  });

  // Helper: get price per KG for a material type
  function matPricePerKg(matType) {
    const mt = matType.toUpperCase();
    const found = materialPrices.find(mp => (mp.material_type || '').toUpperCase() === mt);
    if (!found) return null;
    if (found.price_hkd_per_g) return found.price_hkd_per_g * 1000;  // g→kg
    if (found.price_hkd_per_lb) return found.price_hkd_per_lb * 2.20462; // lb→kg
    return null;
  }

  // Fixed template rows for plastic types (add more if needed)
  const plasticRows = [
    { row: 31, type: 'ABS' },
    { row: 32, type: 'PP'  },
    { row: 33, type: 'PCTG' },
    { row: 34, type: 'PVC' },
  ];

  // Fill known plastic rows
  const usedTypes = new Set();
  plasticRows.forEach(({ row, type }) => {
    const weight = matWeight[type] || null;
    const priceKg = matPricePerKg(type);
    setVal(ws, row, 2, type);
    setVal(ws, row, 4, weight);
    setVal(ws, row, 5, priceKg);
    if (weight) usedTypes.add(type);
    // F col = formula =ROUND(D*E/1000,3) — not touched
  });

  // If there are other material types not in the fixed rows, we skip (template limitation)

  // ── Section B: Molding Labour (rows 67–90 = injection molding data) ────────
  const MOLD_START = 67, MOLD_END = 90;
  clearRows(ws, MOLD_START, MOLD_END, [1, 2, 3, 4, 5]);

  moldParts.slice(0, MOLD_END - MOLD_START + 1).forEach((part, i) => {
    const r = MOLD_START + i;
    const shots = parseFloat(part.sets_per_toy) || 1;
    const laborPerToy = parseFloat(part.molding_labor) || 0;
    const costPerShot = shots > 0 ? laborPerToy / shots : 0;

    setVal(ws, r, 1, part.part_no || '');
    setVal(ws, r, 2, part.description || '');
    setVal(ws, r, 3, part.machine_type || '');
    setVal(ws, r, 4, shots);
    setVal(ws, r, 5, parseFloat(costPerShot.toFixed(6)));
    // F col = formula =D*E — not touched (already in template)
  });

  // ── Section C: Electronics (rows 101–103) ──────────────────────────────────
  const ELEC_START = 101, ELEC_END = 103;
  clearRows(ws, ELEC_START, ELEC_END, [2, 3, 4, 5]);

  electronicItems.slice(0, ELEC_END - ELEC_START + 1).forEach((item, i) => {
    const r = ELEC_START + i;
    setVal(ws, r, 2, item.part_name || '');
    setVal(ws, r, 3, 'pc');
    setVal(ws, r, 4, parseFloat(item.quantity) || 1);
    setVal(ws, r, 5, parseFloat(item.unit_price_usd) || 0);
    // F col = formula =E*D — not touched
  });

  // ── Section C: Other Hardware (rows 113–134) ───────────────────────────────
  const HW_START = 113, HW_END = 134;
  clearRows(ws, HW_START, HW_END, [2, 3, 4, 5]);

  hardwareItems.slice(0, HW_END - HW_START + 1).forEach((item, i) => {
    const r = HW_START + i;
    setVal(ws, r, 2, item.name || '');
    setVal(ws, r, 3, 'pc');
    setVal(ws, r, 4, parseFloat(item.quantity) || 1);
    setVal(ws, r, 5, parseFloat(item.new_price) || 0);
    // F col = formula =D*E — not touched
  });

  // ── Section E: Decoration (row 153) ────────────────────────────────────────
  if (paintingDetail) {
    const sprayOps = parseInt(paintingDetail.spray_count) || 0;
    const laborPerOp = sprayOps > 0
      ? (parseFloat(paintingDetail.labor_cost_hkd) || 0) / sprayOps
      : 0;
    setVal(ws, 153, 4, sprayOps || null);
    setVal(ws, 153, 5, parseFloat(laborPerOp.toFixed(4)) || null);
    // F153 = formula =E153*D153 — not touched
  }

  // ── Section E: Assembly (row 165) ──────────────────────────────────────────
  // Assembly hours from labor_hkd param (hourly rate) — use a default assembly op
  const assemblyHours = parseFloat(params.assembly_hours) || null;
  const laborRate = parseFloat(params.labor_hkd) || null;
  setVal(ws, 165, 4, assemblyHours);
  setVal(ws, 165, 5, laborRate);
}

// ─── Main Export Function ─────────────────────────────────────────────────────

async function exportVersion(versionId) {
  const d = loadData(versionId);

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(TEMPLATE_PATH);

  // ExcelJS may write NaN formula results as invalid XML — clear them
  wb.eachSheet(ws => {
    ws.eachRow({ includeEmpty: false }, row => {
      row.eachCell({ includeEmpty: false }, cell => {
        if (cell.value && typeof cell.value === 'object' && cell.value.formula !== undefined) {
          const r = cell.value.result;
          if (r === null || r === undefined || (typeof r === 'number' && isNaN(r))) {
            cell.value = { formula: cell.value.formula };  // keep formula, drop bad result
          }
        }
      });
    });
  });

  const vqWs  = wb.getWorksheet('Vendor Quotation');
  const bcdWs = wb.getWorksheet('Body Cost Breakdown');

  if (!vqWs)  throw new Error('Template missing "Vendor Quotation" sheet');
  if (!bcdWs) throw new Error('Template missing "Body Cost Breakdown" sheet');

  fillVQ(vqWs, d);
  fillBCD(bcdWs, d);

  return wb.xlsx.writeBuffer();
}

module.exports = { exportVersion };
