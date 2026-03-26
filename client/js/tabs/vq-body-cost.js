/* Tab: vq-body-cost — A. Body Cost (VQ Summary of BD breakdown) */
const tab_vq_body_cost = {
  render(versionData) {
    const parts = versionData.mold_parts || [];
    const hw = versionData.hardware_items || [];
    const pd = versionData.painting_detail || {};
    const params = versionData.params || {};
    const markup = parseFloat(params.markup_body) || 0;

    // Replicate calcBodyBreakdown logic for display
    const rawSub = parts.reduce((s, p) => s + (parseFloat(p.material_cost_hkd) || 0), 0);
    const moldSub = parts.reduce((s, p) => s + (parseFloat(p.molding_labor) || 0), 0);
    const purSub  = hw.reduce((s, h) => s + (parseFloat(h.new_price) || 0), 0);
    const decSub  = (parseFloat(pd.labor_cost_hkd) || 0) + (parseFloat(pd.paint_cost_hkd) || 0);
    const othSub  = 0;

    const rawAmt  = rawSub * (1 + markup);
    const moldAmt = moldSub * (1 + markup);
    const purAmt  = purSub * (1 + markup);
    const decAmt  = decSub * (1 + markup);
    const othAmt  = othSub * (1 + markup);
    const total   = rawAmt + moldAmt + purAmt + decAmt + othAmt;

    function pct(amt) {
      return total > 0 ? (amt / total * 100).toFixed(1) + '%' : '—';
    }

    const sections = [
      { label: 'A. Raw Material',      sub: rawSub,  amt: rawAmt,  pct: pct(rawAmt) },
      { label: 'B. Molding Labour',    sub: moldSub, amt: moldAmt, pct: pct(moldAmt) },
      { label: 'C. Purchase Parts',    sub: purSub,  amt: purAmt,  pct: pct(purAmt) },
      { label: 'D. Decoration (喷油)', sub: decSub,  amt: decAmt,  pct: pct(decAmt) },
      { label: 'E. Others',            sub: othSub,  amt: othAmt,  pct: pct(othAmt) },
    ];

    const rows = sections.map(s => `
      <tr>
        <td>${s.label}</td>
        <td class="num">${formatNumber(s.sub, 4)}</td>
        <td class="num">${(markup * 100).toFixed(1)}%</td>
        <td class="num">${formatNumber(s.amt, 4)}</td>
        <td class="num">${s.pct}</td>
      </tr>
    `).join('');

    return `
      <div class="toolbar">
        <span class="toolbar-title">A. Body Cost</span>
        <span class="toolbar-spacer"></span>
        <span class="toolbar-stats">
          Total Body Cost: <b>${formatNumber(total, 4)}</b> HKD
        </span>
      </div>
      <div class="data-table-wrap">
        <table class="data-table">
          <thead>
            <tr>
              <th>Section</th>
              <th>Sub Total HKD</th>
              <th>Mark Up</th>
              <th>Amount HKD</th>
              <th>% of Body</th>
            </tr>
          </thead>
          <tbody>
            ${rows}
            <tr style="font-weight:bold;background:#f0f4fa">
              <td>合计 Total</td>
              <td class="num">—</td>
              <td class="num">—</td>
              <td class="num">${formatNumber(total, 4)}</td>
              <td class="num">100%</td>
            </tr>
          </tbody>
        </table>
      </div>
      <p style="margin:12px 0 0;font-size:12px;color:#888">
        * 明细请在 Body Cost Breakdown (BD) 各标签编辑
      </p>
    `;
  },

  init(container, versionData, versionId) {
    // Read-only view; editing done in BD tabs
  },
};
