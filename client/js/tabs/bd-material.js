/* Tab: bd-material — A. Raw Material Cost (Breakdown) */
const tab_bd_material = {
  render(versionData) {
    const parts = versionData.mold_parts || [];
    const matPrices = versionData.material_prices || [];
    const params = versionData.params || {};
    const markup = parseFloat(params.markup_body) || 0;

    const subTotal = parts.reduce((s, p) => s + (parseFloat(p.material_cost_hkd) || 0), 0);
    const amount = subTotal * (1 + markup);

    const rows = parts.map((p, i) => `
      <tr data-idx="${i}">
        <td class="center"><input type="checkbox" class="row-check" data-id="${p.id}"></td>
        <td>${escapeHtml(p.part_no || '')}</td>
        <td style="max-width:180px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${escapeHtml(p.description || '')}">${escapeHtml(p.description || '')}</td>
        <td class="editable" data-id="${p.id}" data-field="material" data-type="select">${escapeHtml(p.material || '')}</td>
        <td class="editable num" data-id="${p.id}" data-field="weight_g" data-type="number">${p.weight_g != null ? p.weight_g : ''}</td>
        <td class="num">${formatNumber(p.unit_price_hkd_g, 6)}</td>
        <td class="num">${formatNumber(p.material_cost_hkd, 4)}</td>
      </tr>
    `).join('');

    return `
      <div class="toolbar">
        <span class="toolbar-title">A. Raw Material Cost</span>
        <button class="btn btn-danger" id="bdMatDelete">删除选中</button>
        <span class="toolbar-spacer"></span>
        <span class="toolbar-stats">
          Sub Total: <b>${formatNumber(subTotal, 4)}</b> &nbsp;|&nbsp;
          Mark Up: <b>${(markup * 100).toFixed(1)}%</b> &nbsp;|&nbsp;
          Amount: <b>${formatNumber(amount, 4)}</b>
        </span>
      </div>
      <div class="data-table-wrap">
        <table class="data-table">
          <thead>
            <tr>
              <th><input type="checkbox" id="bdMatAll"></th>
              <th>模号</th><th>名称</th><th>料型</th>
              <th>料重(G)</th><th>料价HKD/g</th><th>料金额HKD</th>
            </tr>
          </thead>
          <tbody>${rows}</tbody>
        </table>
      </div>
    `;
  },

  init(container, versionData, versionId) {
    const matOptions = (versionData.material_prices || []).map(m => m.material_type).filter(Boolean);

    container.querySelector('#bdMatAll')?.addEventListener('change', e => {
      container.querySelectorAll('.row-check').forEach(cb => cb.checked = e.target.checked);
    });

    container.querySelector('#bdMatDelete')?.addEventListener('click', async () => {
      const ids = [...container.querySelectorAll('.row-check:checked')].map(cb => cb.dataset.id);
      if (!ids.length) return showToast('请先选择要删除的行', 'info');
      if (!confirm(`确定删除 ${ids.length} 行？`)) return;
      try {
        await Promise.all(ids.map(id => api.deleteSectionItem(versionId, 'mold-parts', id)));
        await api.calculate(versionId);
        app.selectVersion(null, versionId);
      } catch (e) { showToast('删除失败: ' + e.message, 'error'); }
    });

    container.querySelectorAll('td.editable').forEach(td => {
      const id = td.dataset.id;
      const field = td.dataset.field;
      const type = td.dataset.type;
      const part = versionData.mold_parts.find(p => String(p.id) === id) || {};
      makeEditable(td, {
        type,
        choices: type === 'select' ? matOptions : [],
        value: part[field],
        onSave: async (val) => {
          try {
            await api.updateSectionItem(versionId, 'mold-parts', id, { [field]: val });
            await api.calculate(versionId);
            app.selectVersion(null, versionId);
          } catch (e) { showToast('保存失败: ' + e.message, 'error'); }
        },
      });
    });
  },
};
