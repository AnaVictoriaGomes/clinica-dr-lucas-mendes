import * as XLSX from 'xlsx';

const esc = (s: any) => String(s ?? '').replace(/[&<>"']/g, m => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[m] || m));

export interface ExportColumn {
  key: string;
  label: string;
  selected: boolean;
  getter?: (item: any) => any;
}

export function exportDataFile(
  filenameBase: string,
  format: 'csv' | 'xlsx',
  columns: ExportColumn[],
  items: any[]
) {
  const activeCols = columns.filter(c => c.selected);
  if (!activeCols.length) {
    (window as any).KommoApp?.toast?.('Selecione pelo menos um campo para exportar.');
    return;
  }

  const headers = activeCols.map(c => c.label);
  const rows = items.map(item => {
    return activeCols.map(c => {
      if (typeof c.getter === 'function') {
        return c.getter(item);
      }
      return item[c.key] ?? '';
    });
  });

  const timestamp = new Date().toISOString().slice(0, 10);
  const fileName = `${filenameBase}_${timestamp}.${format}`;

  if (format === 'xlsx') {
    const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Dados');
    XLSX.writeFile(wb, fileName);
  } else {
    // CSV with UTF-8 BOM
    const q = (v: any) => '"' + String(v ?? '').replace(/"/g, '""') + '"';
    const csvContent = '\ufeff' + [headers, ...rows].map(r => r.map(q).join(';')).join('\r\n');
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 1500);
  }
}

export function renderExportModal(
  title: string,
  columns: ExportColumn[],
  totalFiltered: number,
  onConfirm: (format: 'csv' | 'xlsx', selectedCols: ExportColumn[]) => void,
  onClose: () => void
) {
  const modalId = 'export_dialog_modal';
  const existing = document.getElementById(modalId);
  if (existing) existing.remove();

  const overlay = document.createElement('div');
  overlay.id = modalId;
  overlay.style.cssText = `
    position: fixed; inset: 0; z-index: 9999;
    background: rgba(0,0,0,0.5); backdrop-filter: blur(3px);
    display: flex; align-items: center; justify-content: center; padding: 16px;
  `;

  let currentCols = columns.map(c => ({ ...c }));

  function renderContent() {
    overlay.innerHTML = `
      <div style="background:var(--card);border:1px solid var(--line);border-radius:14px;padding:22px;max-width:540px;width:100%;box-shadow:0 12px 38px rgba(0,0,0,0.35);animation:fadeIn .15s ease">
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:12px">
          <div>
            <h2 style="margin:0;font-size:17px;font-weight:700">📥 Escolher campos para exportar</h2>
            <div class="note" style="font-size:12px;margin-top:2px">${esc(title)} · <b>${totalFiltered}</b> registro(s) correspondente(s) aos filtros</div>
          </div>
          <button type="button" class="x" id="exp_close_btn" style="font-size:18px">×</button>
        </div>

        <div style="display:flex;gap:8px;margin-bottom:12px;align-items:center">
          <button type="button" class="b s" id="exp_sel_all" style="font-size:11.5px;padding:4px 8px">Selecionar todos</button>
          <button type="button" class="b s" id="exp_desel_all" style="font-size:11.5px;padding:4px 8px">Desmarcar todos</button>
        </div>

        <div style="max-height:260px;overflow-y:auto;border:1px solid var(--line);border-radius:9px;padding:10px;background:var(--bg);display:grid;grid-template-columns:1fr 1fr;gap:6px">
          ${currentCols.map((col, idx) => `
            <label style="display:flex;align-items:center;gap:6px;font-size:12.5px;cursor:pointer;user-select:none;padding:2px 4px;border-radius:5px">
              <input type="checkbox" data-col-idx="${idx}" ${col.selected ? 'checked' : ''} style="width:auto;margin:0">
              <span style="font-weight:${col.selected ? '600' : '400'}">${esc(col.label)}</span>
            </label>
          `).join('')}
        </div>

        <div style="margin-top:14px;padding-top:12px;border-top:1px solid var(--line);display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:8px">
          <div style="display:flex;align-items:center;gap:6px">
            <span style="font-size:12px;font-weight:600;color:var(--mut)">Formato:</span>
            <select id="exp_fmt_select" style="min-height:30px;padding:3px 8px;font-size:12px">
              <option value="xlsx">Excel (.xlsx)</option>
              <option value="csv">CSV (.csv)</option>
            </select>
          </div>
          <div style="display:flex;gap:6px">
            <button type="button" class="b s" id="exp_cancel_btn">Cancelar</button>
            <button type="button" class="b" id="exp_confirm_btn">Baixar arquivo</button>
          </div>
        </div>
      </div>
    `;

    overlay.querySelector('#exp_close_btn')?.addEventListener('click', () => { overlay.remove(); onClose(); });
    overlay.querySelector('#exp_cancel_btn')?.addEventListener('click', () => { overlay.remove(); onClose(); });

    overlay.querySelector('#exp_sel_all')?.addEventListener('click', () => {
      currentCols.forEach(c => c.selected = true);
      renderContent();
    });

    overlay.querySelector('#exp_desel_all')?.addEventListener('click', () => {
      currentCols.forEach(c => c.selected = false);
      renderContent();
    });

    overlay.querySelectorAll('[data-col-idx]').forEach(inp => {
      inp.addEventListener('change', (e: any) => {
        const idx = parseInt(e.target.dataset.colIdx, 10);
        if (currentCols[idx]) {
          currentCols[idx].selected = e.target.checked;
        }
      });
    });

    overlay.querySelector('#exp_confirm_btn')?.addEventListener('click', () => {
      const selected = currentCols.filter(c => c.selected);
      if (!selected.length) {
        (window as any).KommoApp?.toast?.('Selecione ao menos um campo para exportar.');
        return;
      }
      const fmt = (overlay.querySelector('#exp_fmt_select') as HTMLSelectElement)?.value as 'csv' | 'xlsx';
      overlay.remove();
      onConfirm(fmt, currentCols);
    });
  }

  renderContent();
  document.body.appendChild(overlay);
}
