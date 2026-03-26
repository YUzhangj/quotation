const ExcelJS = require('exceljs');

// ─── Cell Value Helper ────────────────────────────────────────────────────────

function cellVal(cell) {
  if (!cell) return null;
  const v = cell.value;
  if (v === null || v === undefined) return null;
  if (typeof v === 'object' && v.result !== undefined) return v.result; // formula
  if (typeof v === 'object' && v.text !== undefined) return v.text;     // rich text
  return v;
}

function numVal(cell) {
  const v = cellVal(cell);
  if (v === null || v === undefined || v === '') return null;
  const n = parseFloat(String(v).replace(/,/g, ''));
  return isNaN(n) ? null : n;
}

function strVal(cell) {
  const v = cellVal(cell);
  if (v === null || v === undefined) return null;
  return String(v).trim();
}

// ─── Sheet Detection ──────────────────────────────────────────────────────────

function detectLatestSheet(workbook) {
  const sheets = workbook.worksheets.map(ws => ws.name);

  // Try to find sheets with "报价明细" prefix
  const candidates = sheets.filter(n => n.includes('报价明细'));

  if (candidates.length === 0) {
    // Fallback: return last sheet
    return sheets[sheets.length - 1];
  }

  if (candidates.length === 1) return candidates[0];

  // Parse date/version suffix and pick the latest
  function parseSheetDate(name) {
    const suffix = name.replace(/.*报价明细[-\s]*/, '').trim();
    // YYMMDD format e.g. "260310"
    if (/^\d{6}$/.test(suffix)) {
      return parseInt(suffix, 10);
    }
    // MMDD format e.g. "0725"
    if (/^\d{4}$/.test(suffix)) {
      return parseInt(suffix, 10);
    }
    // V2, V3 style
    const vm = suffix.match(/^V(\d+)$/i);
    if (vm) return parseInt(vm[1], 10);
    return 0;
  }

  candidates.sort((a, b) => parseSheetDate(b) - parseSheetDate(a));
  return candidates[0];
}

// ─── Header Parser (R1–R16) ───────────────────────────────────────────────────

function parseHeader(ws) {
  // R1: product_no in B1
  const product_no = strVal(ws.getCell('B1'));

  // R2-R6: Material price table
  // Row 2: 料型 labels (B2:V2)
  // Row 4: 单价HKD/磅 (B4:V4)
  // Row 5: 料单价HKD/g (B5:V5)
  // Row 6: 料单价RMB/g (B6:V6)
  const materialPrices = [];
  // Columns B through V (2..22)
  for (let col = 2; col <= 22; col++) {
    const material_type = strVal(ws.getCell(2, col));
    if (!material_type) continue;
    const price_hkd_per_lb = numVal(ws.getCell(4, col));
    const price_hkd_per_g = numVal(ws.getCell(5, col));
    const price_rmb_per_g = numVal(ws.getCell(6, col));
    if (material_type) {
      materialPrices.push({ material_type, price_hkd_per_lb, price_hkd_per_g, price_rmb_per_g });
    }
  }

  // R8-R10: Machine price table
  // Row 8: 机型 labels (B8:N8)
  // Row 9: 啤工价HKD (B9:N9)
  // Row 10: 啤工价RMB (B10:N10)
  const machinePrices = [];
  for (let col = 2; col <= 14; col++) {
    const machine_type = strVal(ws.getCell(8, col));
    if (!machine_type) continue;
    const price_hkd = numVal(ws.getCell(9, col));
    const price_rmb = numVal(ws.getCell(10, col));
    machinePrices.push({ machine_type, price_hkd, price_rmb });
  }

  // R11-R14: Exchange rates and params
  // C11: hkd_rmb_quote, C12: hkd_rmb_check, C13: rmb_hkd, C14: hkd_usd
  // F13: labor_hkd, F14: box_price_hkd
  const hkd_rmb_quote = numVal(ws.getCell('C11'));
  const hkd_rmb_check = numVal(ws.getCell('C12'));
  const rmb_hkd = numVal(ws.getCell('C13'));
  const hkd_usd = numVal(ws.getCell('C14'));
  const labor_hkd = numVal(ws.getCell('F13'));
  const box_price_hkd = numVal(ws.getCell('F14'));

  // R15: date_code, R16: reference number
  const date_code = strVal(ws.getCell('B15')) || strVal(ws.getCell('C15'));
  const ref_no = strVal(ws.getCell('B16')) || strVal(ws.getCell('C16'));

  return {
    product_no,
    materialPrices,
    machinePrices,
    params: { hkd_rmb_quote, hkd_rmb_check, rmb_hkd, hkd_usd, labor_hkd, box_price_hkd },
    date_code,
    ref_no,
  };
}

// ─── Mold Parts Parser (R17+) ────────────────────────────────────────────────

function parseMoldParts(ws) {
  const moldParts = [];

  // R17: header — verify A17 = "模号"
  // R18 onward: data rows
  let row = 18;
  let sortOrder = 0;

  while (row <= 200) {
    const colA = strVal(ws.getCell(row, 1));
    const colC = strVal(ws.getCell(row, 3));
    const colI = strVal(ws.getCell(row, 9));

    // Stop on 合计 row
    if (
      (colA && colA.includes('合计')) ||
      (colC && colC.includes('合计')) ||
      (colI && colI.includes('合计'))
    ) {
      break;
    }

    // Skip empty rows (no part_no and no description)
    const part_no = strVal(ws.getCell(row, 1));
    const description = strVal(ws.getCell(row, 2));

    if (!part_no && !description) {
      // Allow up to 3 consecutive empty rows before stopping
      row++;
      continue;
    }

    const material = strVal(ws.getCell(row, 3));
    const weight_g = numVal(ws.getCell(row, 4));
    const unit_price_hkd_g = numVal(ws.getCell(row, 5));
    const machine_type = strVal(ws.getCell(row, 6));
    const cavity_count = numVal(ws.getCell(row, 7));
    const sets_per_toy = numVal(ws.getCell(row, 8));
    const target_qty = numVal(ws.getCell(row, 9));
    const molding_labor = numVal(ws.getCell(row, 10));
    const material_cost_hkd = numVal(ws.getCell(row, 11));
    const mold_cost_rmb = numVal(ws.getCell(row, 12));
    const remark = strVal(ws.getCell(row, 13));

    const is_old_mold = (remark && remark.includes('旧模')) || mold_cost_rmb === null ? 1 : 0;

    moldParts.push({
      part_no,
      description,
      material,
      weight_g,
      unit_price_hkd_g,
      machine_type,
      cavity_count: cavity_count ? Math.round(cavity_count) : null,
      sets_per_toy: sets_per_toy ? Math.round(sets_per_toy) : null,
      target_qty: target_qty ? Math.round(target_qty) : null,
      molding_labor,
      material_cost_hkd,
      mold_cost_rmb,
      remark,
      is_old_mold,
      sort_order: sortOrder++,
    });

    row++;
  }

  return moldParts;
}

// ─── Cost Items Parser ───────────────────────────────────────────────────────

function parseCostItems(ws) {
  // Parse a range of rows into items [{name, quantity, old_price, new_price, difference, tax_type}]
  // R40 is the header row; R41-R43 are summary computed rows (料价进口料, 料价国内采购, 啤工)
  // Actual labor items start at R44
  function parseItemRange(startRow, endRow) {
    const items = [];
    for (let r = startRow; r <= endRow; r++) {
      const name = strVal(ws.getCell(r, 1));
      if (!name) continue;
      const quantity = numVal(ws.getCell(r, 2));
      const old_price = numVal(ws.getCell(r, 3));
      const new_price = numVal(ws.getCell(r, 4));
      const difference = numVal(ws.getCell(r, 5));
      const tax_type = strVal(ws.getCell(r, 9));
      items.push({ name, quantity, old_price, new_price, difference, tax_type });
    }
    return items;
  }

  // R44-R47: Labor items (装配人工, 包装人工, 喷油人工, 油漆)
  // R40 = header, R41-R43 = computed summary rows, R44+ = actual items
  const laborItems = parseItemRange(44, 47);

  // R48-R76: Hardware items (五金件, 电镀件, 贴纸, IC, PCBA, 电池)
  const hardwareItems = parseItemRange(48, 76);

  // R77-R93: Packaging items (Window Box, Insert card, etc.)
  const packagingItems = parseItemRange(77, 93);

  return { laborItems, hardwareItems, packagingItems };
}

// ─── Summary Parser ──────────────────────────────────────────────────────────

function parseSummary(ws) {
  // R94: 包装合计 (C94), R95: 附加税 (C95)
  const packaging_total = numVal(ws.getCell('C94'));
  const surcharge = numVal(ws.getCell('C95'));

  // R97-R106: Cost progression (column C = 盐田40柜 scenario)
  const factory_price = numVal(ws.getCell('C97'));
  const transport_cost = numVal(ws.getCell('C99'));    // 运费
  const mark_point = numVal(ws.getCell('C102'));        // 码点
  const payment_adj = numVal(ws.getCell('C104'));       // 找数 ÷
  const total_hkd = numVal(ws.getCell('C105'));         // TOTAL HK$
  const total_usd = numVal(ws.getCell('C106'));         // USD

  // Dimensions at R108-R110, columns H(8)/J(10)/L(12)
  // R107: headers (L, W, H)
  // R108: product dimensions in inches
  // R109: carton dimensions in inches; I109 = paper type (e.g. "A=B")
  // R110: CU.FT, carton price, pcs per carton
  const product_l = numVal(ws.getCell(108, 8));    // H108
  const product_w = numVal(ws.getCell(108, 10));   // J108
  const product_h = numVal(ws.getCell(108, 12));   // L108
  const carton_l = numVal(ws.getCell(109, 8));     // H109
  const carton_w = numVal(ws.getCell(109, 10));    // J109
  const carton_h = numVal(ws.getCell(109, 12));    // L109
  const carton_paper = strVal(ws.getCell(109, 9)); // I109
  const carton_cuft = numVal(ws.getCell(110, 8));  // H110
  const carton_price = numVal(ws.getCell(110, 10)); // J110
  const pcs_per_carton = numVal(ws.getCell(110, 12)); // L110

  // R129-R136: Mold costs
  // R129-R130: section headers
  // R131=模具费用, R132=五金模/夹具, R133=喷油模具, R134=模具总计
  // R135=客补贴模费美金, R136=模费分摊
  const mold_cost_rmb = numVal(ws.getCell('C131'));
  const hardware_mold_cost_rmb = numVal(ws.getCell('C132'));
  const paint_mold_cost_rmb = numVal(ws.getCell('C133'));
  const total_mold_rmb = numVal(ws.getCell('C134'));
  const customer_subsidy_usd = numVal(ws.getCell('C135'));
  const amortization_rmb = numVal(ws.getCell('C136'));
  const amortization_usd = numVal(ws.getCell('D136'));
  const hkd_usd = numVal(ws.getCell('C14'));
  const total_mold_usd = total_mold_rmb && hkd_usd ? total_mold_rmb * hkd_usd : null;

  return {
    pricing: { packaging_total, surcharge, factory_price, transport_cost, mark_point, payment_adj, total_hkd, total_usd },
    dimensions: {
      product_l_inch: product_l, product_w_inch: product_w, product_h_inch: product_h,
      carton_l_inch: carton_l, carton_w_inch: carton_w, carton_h_inch: carton_h,
      carton_cuft, carton_price, pcs_per_carton: pcs_per_carton ? Math.round(pcs_per_carton) : null,
      carton_paper,
    },
    moldCost: {
      mold_cost_rmb, hardware_mold_cost_rmb, paint_mold_cost_rmb,
      total_mold_rmb, total_mold_usd, customer_subsidy_usd,
      amortization_qty: null,
      amortization_rmb, amortization_usd, customer_quote_usd: null,
    },
  };
}

// ─── Transport Parser (R141–R155) ────────────────────────────────────────────

function parseTransport(ws) {
  // Actual layout (verified against real file):
  // R141: section header
  // R142: 1箱的CUFT: [B]=value [C]=CUFT
  // R143: 1箱装的个数: [B]=value [C]=PCS
  // R144: 10吨车: [B]=cuft
  // R145: 5吨车: [B]=cuft
  // R146: 40": [B]=cuft
  // R147: 20": [B]=cuft
  // R148-R155: shipping costs in B column (HK40, HK20, YT40, YT20, HK10T, YT10T, HK5T, YT5T)
  const cuft_per_box = numVal(ws.getCell(142, 2));  // B142
  const pcs_per_box = numVal(ws.getCell(143, 2));   // B143

  const truck_10t_cuft = numVal(ws.getCell(144, 2));
  const truck_5t_cuft = numVal(ws.getCell(145, 2));
  const container_40_cuft = numVal(ws.getCell(146, 2));
  const container_20_cuft = numVal(ws.getCell(147, 2));

  const hk_40_cost = numVal(ws.getCell(148, 2));
  const hk_20_cost = numVal(ws.getCell(149, 2));
  const yt_40_cost = numVal(ws.getCell(150, 2));
  const yt_20_cost = numVal(ws.getCell(151, 2));
  const hk_10t_cost = numVal(ws.getCell(152, 2));
  const yt_10t_cost = numVal(ws.getCell(153, 2));
  const hk_5t_cost = numVal(ws.getCell(154, 2));
  const yt_5t_cost = numVal(ws.getCell(155, 2));
  // transport_pct and handling_pct are calculated from the totals, not stored directly
  const transport_pct = null;
  const handling_pct = null;

  return {
    cuft_per_box, pcs_per_box: pcs_per_box ? Math.round(pcs_per_box) : null,
    truck_10t_cuft, truck_5t_cuft, container_40_cuft, container_20_cuft,
    hk_40_cost, hk_20_cost, yt_40_cost, yt_20_cost,
    hk_10t_cost, yt_10t_cost, hk_5t_cost, yt_5t_cost,
    transport_pct, handling_pct,
  };
}

// ─── Electronics Parser ──────────────────────────────────────────────────────

function parseElectronics(workbook) {
  const wsNames = workbook.worksheets.map(ws => ws.name);
  const elecSheet = wsNames.find(n => n === '电子' || n.includes('电子'));
  if (!elecSheet) return { electronicItems: [], electronicSummary: null };

  const ws = workbook.getWorksheet(elecSheet);
  const electronicItems = [];

  // R6-R35: component list (columns A=part_name, B=spec, C=quantity, D=unit_price_usd, E=total_usd, F=remark)
  for (let r = 6; r <= 35; r++) {
    const part_name = strVal(ws.getCell(r, 1));
    if (!part_name) continue;
    const spec = strVal(ws.getCell(r, 2));
    const quantity = numVal(ws.getCell(r, 3));
    const unit_price_usd = numVal(ws.getCell(r, 4));
    const total_usd = numVal(ws.getCell(r, 5));
    const remark = strVal(ws.getCell(r, 6));
    electronicItems.push({ part_name, spec, quantity, unit_price_usd, total_usd, remark, sort_order: r - 6 });
  }

  // Summary section (after component list)
  const parts_cost = numVal(ws.getCell('D37')) || numVal(ws.getCell('E37'));
  const bonding_cost = numVal(ws.getCell('D38')) || numVal(ws.getCell('E38'));
  const smt_cost = numVal(ws.getCell('D39')) || numVal(ws.getCell('E39'));
  const labor_cost = numVal(ws.getCell('D40')) || numVal(ws.getCell('E40'));
  const test_cost = numVal(ws.getCell('D41')) || numVal(ws.getCell('E41'));
  const packaging_transport = numVal(ws.getCell('D42')) || numVal(ws.getCell('E42'));
  const total_cost = numVal(ws.getCell('D43')) || numVal(ws.getCell('E43'));
  const profit_margin = numVal(ws.getCell('D44')) || numVal(ws.getCell('E44'));
  const final_price_usd = numVal(ws.getCell('D45')) || numVal(ws.getCell('E45'));
  const pcb_mold_cost_usd = numVal(ws.getCell('D46')) || numVal(ws.getCell('E46'));

  const electronicSummary = {
    parts_cost, bonding_cost, smt_cost, labor_cost, test_cost,
    packaging_transport, total_cost, profit_margin, final_price_usd, pcb_mold_cost_usd,
  };

  return { electronicItems, electronicSummary };
}

// ─── Painting Parser ─────────────────────────────────────────────────────────

function parsePainting(ws) {
  // R46-R47: labor and paint costs
  const labor_cost_hkd = numVal(ws.getCell('C46')) || numVal(ws.getCell('D46'));
  const paint_cost_hkd = numVal(ws.getCell('C47')) || numVal(ws.getCell('D47'));

  // R129-R132: painting detail counts
  const clamp_count = numVal(ws.getCell('C129')) || numVal(ws.getCell('D129'));
  const print_count = numVal(ws.getCell('C130')) || numVal(ws.getCell('D130'));
  const wipe_count = numVal(ws.getCell('C131')) || numVal(ws.getCell('D131'));
  const edge_count = numVal(ws.getCell('C132')) || numVal(ws.getCell('D132'));
  const spray_count = numVal(ws.getCell('C133')) || numVal(ws.getCell('D133'));
  const total_operations = numVal(ws.getCell('C134')) || numVal(ws.getCell('D134'));
  const quoted_price_hkd = numVal(ws.getCell('C135')) || numVal(ws.getCell('D135'));

  return {
    labor_cost_hkd, paint_cost_hkd,
    clamp_count: clamp_count ? Math.round(clamp_count) : null,
    print_count: print_count ? Math.round(print_count) : null,
    wipe_count: wipe_count ? Math.round(wipe_count) : null,
    edge_count: edge_count ? Math.round(edge_count) : null,
    spray_count: spray_count ? Math.round(spray_count) : null,
    total_operations: total_operations ? Math.round(total_operations) : null,
    quoted_price_hkd,
  };
}

// ─── Main Parse Function ─────────────────────────────────────────────────────

async function parseWorkbook(filePath) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);

  const sheetName = detectLatestSheet(workbook);
  const ws = workbook.getWorksheet(sheetName);

  if (!ws) {
    throw new Error(`Sheet "${sheetName}" not found in workbook`);
  }

  const header = parseHeader(ws);
  const moldParts = parseMoldParts(ws);
  const costItems = parseCostItems(ws);
  const summary = parseSummary(ws);
  const transport = parseTransport(ws);
  const { electronicItems, electronicSummary } = parseElectronics(workbook);
  const paintingDetail = parsePainting(ws);

  return {
    sheetName,
    product: {
      product_no: header.product_no,
      date_code: header.date_code,
      ref_no: header.ref_no,
    },
    params: header.params,
    materialPrices: header.materialPrices,
    machinePrices: header.machinePrices,
    moldParts,
    hardwareItems: costItems.hardwareItems,
    laborItems: costItems.laborItems,
    packagingItems: costItems.packagingItems,
    electronicItems,
    electronicSummary,
    paintingDetail,
    transportConfig: transport,
    productDimension: summary.dimensions,
    moldCost: summary.moldCost,
    pricing: summary.pricing,
  };
}

module.exports = { parseWorkbook, detectLatestSheet };
