import * as XLSX from 'xlsx';

export interface ChartSeriesConfig {
  id: string;
  name: string;
  source: 'semanas' | 'consultas' | 'metas' | 'planilha';
  field: string;
  agg: 'SUM' | 'AVG' | 'COUNT' | 'MIN' | 'MAX' | 'PCT';
  color: string;
  chartSubType?: 'bar' | 'line';
  yAxis?: 'left' | 'right';
}

export interface CustomChartConfig {
  id: string;
  title: string;
  type: 'group' | 'stack' | 'hbar' | 'line' | 'area' | 'pie' | 'donut' | 'combo' | 'scatter';
  dataSource: 'semanas' | 'consultas' | 'cruzado' | 'metas' | 'planilha';
  xAxisField: string;
  xAxisGrouping: 'individual' | 'week' | 'month' | 'year' | 'category';
  xAxisSort: 'asc' | 'desc' | 'val_desc' | 'val_asc';
  spreadsheetId?: string;
  filterOrigin?: string;
  filterStatus?: string;
  filterFrom?: string;
  filterTo?: string;
  series: ChartSeriesConfig[];
  showInDashboard: boolean;
  width?: 'half' | 'full';
  isDefaultOverride?: boolean;
  defaultKey?: string;
}

export interface SpreadsheetSource {
  id: string;
  name: string;
  uploadedAt: string;
  headers: string[];
  rows: Record<string, any>[];
}

// Site palette synced with CSS variables
export function getSitePalette(): string[] {
  const style = getComputedStyle(document.documentElement);
  const g = style.getPropertyValue('--g').trim() || '#1f7a6d';
  const b = style.getPropertyValue('--b').trim() || '#3d6fa8';
  const a = style.getPropertyValue('--a').trim() || '#d99a1c';
  const r = style.getPropertyValue('--r').trim() || '#c2455a';
  const p = style.getPropertyValue('--p').trim() || '#7a5aa6';
  const k = style.getPropertyValue('--k').trim() || '#9aa8a4';
  return [g, b, a, r, p, k, '#2d9d86', '#e07a5f'];
}

export const STANDARD_CHARTS_META: Record<string, { title: string; type: CustomChartConfig['type']; desc: string }> = {
  std_fu: { title: 'Mensagens de follow-up por tentativa', type: 'stack', desc: 'FU 2 a FU 5 empilhados por semana' },
  std_resp: { title: 'Resposta e conversão do follow-up', type: 'line', desc: 'Taxa de resposta e conversão (%)' },
  std_status: { title: 'Status dos pacientes (total acumulado em cada semana)', type: 'stack', desc: 'Sem resposta, pensando, aguardando, não vão seguir' },
  std_agenda: { title: 'Agenda da semana: realizadas, canceladas e reagendadas', type: 'group', desc: 'Previstas, realizadas, canceladas e reagendadas' },
  std_funnel: { title: 'Funil: leads, respostas e agendamentos', type: 'group', desc: 'Novos leads, respostas e agendamentos' },
  std_goals: { title: 'Acompanhamento de Metas por Semana', type: 'line', desc: 'Meta estipulada vs Resultado alcançado' }
};

function esc(s: any): string {
  return String(s ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c] || c));
}

function NICE(m: number): number {
  if (m <= 0) return 10;
  const p = Math.pow(10, Math.floor(Math.log10(m)));
  const r = m / p;
  return (r <= 1 ? 1 : r <= 2 ? 2 : r <= 5 ? 5 : 10) * p;
}

export function getAvailableXFields(dataSource: string, ctx: any): { key: string; label: string }[] {
  const { FIELDS, cfg } = ctx;
  if (dataSource === 'semanas') {
    return [
      { key: 'id', label: 'Data da Semana' },
      { key: 'month', label: 'Mês' },
      { key: 'year', label: 'Ano' },
      ...FIELDS.map(([k]: [string, string]) => ({ key: k, label: ctx.lab(k) }))
    ];
  }
  if (dataSource === 'consultas') {
    return [
      { key: 'consult', label: 'Data da Consulta' },
      { key: 'schedOn', label: 'Data de Agendamento' },
      { key: 'newDate', label: 'Nova Data (Reagendamento)' },
      { key: 'origin', label: 'Origem (Veio de)' },
      { key: 'status', label: 'Status da Consulta' },
      { key: 'statusNewDate', label: 'Status Nova Data' },
      { key: 'regained', label: 'Conseguiu Reagendar?' },
      { key: 'weekday', label: 'Dia da Semana' }
    ];
  }
  if (dataSource === 'cruzado') {
    return [
      { key: 'id', label: 'Semana (Cronológica)' },
      { key: 'month', label: 'Mês' },
      { key: 'year', label: 'Ano' },
      { key: 'origin', label: 'Origem do Paciente (Consultas)' },
      { key: 'status', label: 'Status da Consulta' }
    ];
  }
  if (dataSource === 'metas') {
    return [
      { key: 'id', label: 'Semana (Cronológica)' },
      { key: 'metric', label: 'Indicador da Meta' }
    ];
  }
  if (dataSource === 'planilha') {
    const sp = (cfg.spreadsheets || [])[0];
    if (sp && sp.headers) {
      return sp.headers.map((h: string) => ({ key: h, label: `Coluna: ${h}` }));
    }
    return [{ key: 'Data', label: 'Nenhuma planilha carregada' }];
  }
  return [{ key: 'id', label: 'Data' }];
}

export function getAvailableYFields(source: string, ctx: any): { key: string; label: string }[] {
  const { FIELDS, ORIG, cfg } = ctx;
  if (source === 'semanas') {
    return [
      ...FIELDS.map(([k]: [string, string]) => ({ key: k, label: ctx.lab(k) })),
      { key: 'msgs', label: 'Mensagens de follow-up (FU2 a FU5)' },
      { key: 'cprev', label: 'Consultas previstas da semana' },
      { key: 'tot', label: 'Total de consultas finalizadas' },
      { key: 'rate_resp', label: 'Taxa de resposta do FU (%)' },
      { key: 'rate_conv', label: 'Conversão FU em Agendamento (%)' }
    ];
  }
  if (source === 'consultas') {
    const base = [
      { key: 'count_all', label: 'Quantidade de Consultas (Total)' },
      { key: 'count_agendada', label: 'Consultas Agendadas' },
      { key: 'count_realizada', label: 'Consultas Realizadas' },
      { key: 'count_cancelada', label: 'Consultas Canceladas' },
      { key: 'count_reagendada', label: 'Consultas Reagendadas' },
      { key: 'count_nao_compareceu', label: 'Não Compareceu' },
      { key: 'rate_realizada', label: 'Taxa de Comparecimento (%)' },
      { key: 'rate_cancelada', label: 'Taxa de Cancelamento (%)' }
    ];
    Object.entries(ORIG).forEach(([k, n]) => {
      base.push({ key: 'orig_' + k, label: `Consultas de "${n}"` });
    });
    return base;
  }
  if (source === 'metas') {
    return [
      { key: 'meta_target', label: 'Valor da Meta' },
      { key: 'meta_result', label: 'Resultado Alcançado' }
    ];
  }
  if (source === 'planilha') {
    const sps: SpreadsheetSource[] = cfg.spreadsheets || [];
    const fields: { key: string; label: string }[] = [];
    sps.forEach(sp => {
      sp.headers.forEach(h => {
        fields.push({ key: `${sp.id}::${h}`, label: `[${sp.name}] ${h}` });
      });
    });
    if (!fields.length) fields.push({ key: 'none', label: 'Importe uma planilha abaixo' });
    return fields;
  }
  return [];
}

export function computeChartData(config: CustomChartConfig, ctx: any) {
  const { W, appts, cfg, ORIG, STAT, fd, fdf, normIso, addDays, MONTH_NAMES, pct } = ctx;
  const WA = W();
  let allAppts: any[] = Object.values(appts || {});

  // Apply chart-level filters
  if (config.filterOrigin) {
    allAppts = allAppts.filter(a => a.origin === config.filterOrigin);
  }
  if (config.filterStatus) {
    allAppts = allAppts.filter(a => a.status === config.filterStatus);
  }
  if (config.filterFrom) {
    allAppts = allAppts.filter(a => (normIso(a.consult) || normIso(a.schedOn) || '') >= config.filterFrom!);
  }
  if (config.filterTo) {
    allAppts = allAppts.filter(a => (normIso(a.consult) || normIso(a.schedOn) || '') <= config.filterTo!);
  }

  let filteredWeeks = WA;
  if (config.filterFrom) filteredWeeks = filteredWeeks.filter((w: any) => w.id >= config.filterFrom!);
  if (config.filterTo) filteredWeeks = filteredWeeks.filter((w: any) => w.id <= config.filterTo!);

  const buckets: Record<string, { label: string; cat: string; weeks: any[]; appts: any[]; rows: any[] }> = {};
  const ds = config.dataSource || 'semanas';
  const xField = config.xAxisField || 'id';
  const grp = config.xAxisGrouping || 'individual';

  const getGroupKey = (isoDate: string) => {
    if (!isoDate) return { key: 'sem_data', label: 'Sem data', cat: 'Sem data' };
    if (grp === 'month') {
      const p = isoDate.split('-');
      const mIdx = parseInt(p[1], 10) - 1;
      const mName = (MONTH_NAMES[mIdx] || p[1]).slice(0, 3);
      return { key: `${p[0]}-${p[1]}`, label: `${mName}/${p[0].slice(2)}`, cat: `${MONTH_NAMES[mIdx]} ${p[0]}` };
    }
    if (grp === 'year') {
      const y = isoDate.split('-')[0];
      return { key: y, label: y, cat: `Ano ${y}` };
    }
    return { key: isoDate, label: fd(isoDate), cat: fdf(isoDate) };
  };

  if (ds === 'semanas' || ds === 'cruzado' || ds === 'metas') {
    if (ds === 'cruzado' && (xField === 'origin' || xField === 'status')) {
      allAppts.forEach(a => {
        const val = xField === 'origin' ? (ORIG[a.origin] || a.origin) : (STAT[a.status] || a.status);
        const key = a[xField] || 'outro';
        if (!buckets[key]) buckets[key] = { label: val, cat: val, weeks: [], appts: [], rows: [] };
        buckets[key].appts.push(a);
      });
    } else {
      filteredWeeks.forEach((w: any) => {
        const g = getGroupKey(w.id);
        if (!buckets[g.key]) buckets[g.key] = { label: g.label, cat: `Semana ${g.cat}`, weeks: [], appts: [], rows: [] };
        buckets[g.key].weeks.push(w);
        const wEnd = addDays(w.id, 6);
        const matchedAppts = allAppts.filter(a => {
          const d = normIso(a.consult) || normIso(a.schedOn);
          return d >= w.id && d <= wEnd;
        });
        buckets[g.key].appts.push(...matchedAppts);

        // Match spreadsheet rows by date if spreadsheets exist
        (cfg.spreadsheets || []).forEach((sp: SpreadsheetSource) => {
          sp.rows.forEach(r => {
            const dCol = Object.keys(r).find(k => /data|date|semana/i.test(k));
            if (dCol) {
              const rIso = normIso(String(r[dCol]));
              if (rIso >= w.id && rIso <= wEnd) {
                buckets[g.key].rows.push({ ...r, _spId: sp.id });
              }
            }
          });
        });
      });
    }
  } else if (ds === 'consultas') {
    allAppts.forEach(a => {
      let key = '';
      let label = '';
      let cat = '';
      if (xField === 'origin') {
        key = a.origin || 'novo';
        label = ORIG[key] || key;
        cat = `Origem: ${label}`;
      } else if (xField === 'status') {
        key = a.status || 'agendada';
        label = STAT[key] || key;
        cat = `Status: ${label}`;
      } else if (xField === 'statusNewDate') {
        key = a.statusNewDate || 'nenhum';
        label = ctx.STAT_NEW[a.statusNewDate || ''] || '—';
        cat = label;
      } else if (xField === 'weekday') {
        const iso = normIso(a.consult) || normIso(a.schedOn);
        const days = ['Dom', 'Seg', 'Ter', 'Qua', 'Qui', 'Sex', 'Sáb'];
        const dt = iso ? new Date(iso + 'T12:00:00') : null;
        const dIdx = dt ? dt.getDay() : 0;
        key = String(dIdx);
        label = days[dIdx] || '—';
        cat = label;
      } else {
        const iso = normIso(a[xField]) || normIso(a.consult) || normIso(a.schedOn);
        const g = getGroupKey(iso);
        key = g.key;
        label = g.label;
        cat = g.cat;
      }
      if (!buckets[key]) buckets[key] = { label, cat, weeks: [], appts: [], rows: [] };
      buckets[key].appts.push(a);
    });
  } else if (ds === 'planilha') {
    const sp: SpreadsheetSource | undefined = (cfg.spreadsheets || []).find((s: SpreadsheetSource) => s.id === config.spreadsheetId) || (cfg.spreadsheets || [])[0];
    if (sp && sp.rows) {
      sp.rows.forEach((r, idx) => {
        const rawVal = r[xField] ?? `Reg ${idx + 1}`;
        const key = String(rawVal);
        if (!buckets[key]) buckets[key] = { label: key, cat: key, weeks: [], appts: [], rows: [] };
        buckets[key].rows.push({ ...r, _spId: sp.id });
      });
    }
  }

  let keys = Object.keys(buckets);
  if (config.xAxisSort === 'desc') keys.sort((a, b) => b.localeCompare(a));
  else keys.sort((a, b) => a.localeCompare(b));

  const labels = keys.map(k => buckets[k].label);
  const categories = keys.map(k => buckets[k].cat);

  const palette = getSitePalette();
  const seriesData = (config.series || []).map((s, idx) => {
    const isPct = s.agg === 'PCT' || s.field.startsWith('rate_');
    const values = keys.map(k => {
      const b = buckets[k];
      if (s.source === 'semanas') {
        if (!b.weeks.length) return 0;
        const vals = b.weeks.map(w => {
          if (s.field === 'msgs') return (w.fu2 || 0) + (w.fu3 || 0) + (w.fu4 || 0) + (w.fu5 || 0);
          if (s.field === 'tot') return (w.creal || 0) + (w.ccanc || 0) + (w.creag || 0);
          if (s.field === 'rate_resp') {
            const m = (w.fu2 || 0) + (w.fu3 || 0) + (w.fu4 || 0) + (w.fu5 || 0);
            return m > 0 ? pct(w.resp || 0, m) : 0;
          }
          if (s.field === 'rate_conv') {
            return (w.resp || 0) > 0 ? pct(w.agend || 0, w.resp || 0) : 0;
          }
          return Number(w[s.field]) || 0;
        });
        if (s.agg === 'AVG' || isPct) return vals.reduce((a, c) => a + c, 0) / vals.length;
        if (s.agg === 'MAX') return Math.max(...vals);
        if (s.agg === 'MIN') return Math.min(...vals);
        if (s.agg === 'COUNT') return vals.length;
        return vals.reduce((a, c) => a + c, 0);
      }

      if (s.source === 'consultas') {
        const ap = b.appts;
        if (!ap.length) return 0;
        if (s.field === 'count_all') return ap.length;
        if (s.field === 'count_agendada') return ap.filter(a => a.status === 'agendada').length;
        if (s.field === 'count_realizada') return ap.filter(a => a.status === 'realizada').length;
        if (s.field === 'count_cancelada') return ap.filter(a => a.status === 'cancelada').length;
        if (s.field === 'count_reagendada') return ap.filter(a => a.status === 'reagendada').length;
        if (s.field === 'count_nao_compareceu') return ap.filter(a => a.status === 'nao_compareceu').length;
        if (s.field === 'rate_realizada') return pct(ap.filter(a => a.status === 'realizada').length, ap.length);
        if (s.field === 'rate_cancelada') return pct(ap.filter(a => a.status === 'cancelada' || a.status === 'nao_compareceu').length, ap.length);
        if (s.field.startsWith('orig_')) {
          const origKey = s.field.replace('orig_', '');
          return ap.filter(a => a.origin === origKey).length;
        }
        return ap.length;
      }

      if (s.source === 'planilha') {
        const [spId, colName] = s.field.includes('::') ? s.field.split('::') : ['', s.field];
        const matchingRows = spId ? b.rows.filter(r => r._spId === spId) : b.rows;
        if (!matchingRows.length) return 0;
        if (s.agg === 'COUNT') return matchingRows.length;
        const nums = matchingRows.map(r => parseFloat(String(r[colName] ?? '0').replace(',', '.'))).filter(n => !isNaN(n));
        if (!nums.length) return 0;
        if (s.agg === 'AVG') return nums.reduce((a, c) => a + c, 0) / nums.length;
        if (s.agg === 'MAX') return Math.max(...nums);
        if (s.agg === 'MIN') return Math.min(...nums);
        return nums.reduce((a, c) => a + c, 0);
      }

      return 0;
    });

    return {
      n: s.name,
      c: s.color || palette[idx % palette.length],
      v: values.map(v => Math.round(v * 100) / 100),
      subType: s.chartSubType || 'bar',
      isPct
    };
  });

  if (config.xAxisSort === 'val_desc' || config.xAxisSort === 'val_asc') {
    const indices = labels.map((_, i) => i);
    const firstVals = seriesData[0]?.v || [];
    indices.sort((i, j) => config.xAxisSort === 'val_desc' ? (firstVals[j] || 0) - (firstVals[i] || 0) : (firstVals[i] || 0) - (firstVals[j] || 0));
    return {
      labels: indices.map(i => labels[i]),
      categories: indices.map(i => categories[i]),
      series: seriesData.map(s => ({ ...s, v: indices.map(i => s.v[i]) }))
    };
  }

  return { labels, categories, series: seriesData };
}

export function renderExtendedChartSvg(
  labels: string[],
  series: { n: string; c: string; v: number[]; subType?: string; isPct?: boolean }[],
  mode: CustomChartConfig['type'],
  categories?: string[]
): string {
  if (!labels.length || !series.length) {
    return '<p class="note" style="padding:18px;text-align:center">Sem dados compatíveis para esta combinação de filtros e eixos.</p>';
  }

  const isPct = series.some(s => s.isPct);

  // Pie & Donut
  if (mode === 'pie' || mode === 'donut') {
    const useMultiLabels = series.length === 1 && labels.length > 1;
    const palette = getSitePalette();
    const slices = useMultiLabels
      ? labels.map((l, i) => ({ n: l, c: palette[i % palette.length], val: series[0].v[i] || 0 }))
      : series.map((s, i) => ({ n: s.n, c: s.c || palette[i % palette.length], val: s.v.reduce((a, b) => a + b, 0) }));

    const total = slices.reduce((a, s) => a + Math.max(0, s.val), 0);
    if (total <= 0) return '<p class="note" style="padding:18px;text-align:center">Valores zerados para o gráfico de pizza/rosca.</p>';

    const cx = 130, cy = 110, r = 86, ir = mode === 'donut' ? 46 : 0;
    let angle = -Math.PI / 2;
    let paths = '';

    slices.forEach(sl => {
      if (sl.val <= 0) return;
      const sliceAng = (sl.val / total) * Math.PI * 2;
      const endAng = angle + sliceAng;
      const large = sliceAng > Math.PI ? 1 : 0;
      const x1 = cx + r * Math.cos(angle), y1 = cy + r * Math.sin(angle);
      const x2 = cx + r * Math.cos(endAng), y2 = cy + r * Math.sin(endAng);
      let d = '';
      if (sliceAng >= Math.PI * 1.999) {
        d = `M ${cx} ${cy - r} A ${r} ${r} 0 1 1 ${cx - 0.01} ${cy - r} Z`;
      } else if (ir > 0) {
        const x3 = cx + ir * Math.cos(endAng), y3 = cy + ir * Math.sin(endAng);
        const x4 = cx + ir * Math.cos(angle), y4 = cy + ir * Math.sin(angle);
        d = `M ${x1} ${y1} A ${r} ${r} 0 ${large} 1 ${x2} ${y2} L ${x3} ${y3} A ${ir} ${ir} 0 ${large} 0 ${x4} ${y4} Z`;
      } else {
        d = `M ${cx} ${cy} L ${x1} ${y1} A ${r} ${r} 0 ${large} 1 ${x2} ${y2} Z`;
      }
      const pTxt = ((sl.val / total) * 100).toFixed(1).replace('.', ',') + '%';
      paths += `<path d="${d}" fill="${sl.c}" stroke="var(--card)" stroke-width="1.5" data-tip-series="${esc(sl.n)}" data-tip-color="${sl.c}" data-tip-cat="${pTxt} do total" data-tip-val="${sl.val.toLocaleString('pt-BR')} (${pTxt})"/>`;
      angle = endAng;
    });

    const legendItems = slices.filter(s => s.val > 0).map(sl => {
      const pTxt = ((sl.val / total) * 100).toFixed(1).replace('.', ',') + '%';
      return `<div style="display:flex;align-items:center;gap:6px;font-size:12px"><i style="display:inline-block;width:10px;height:10px;border-radius:2px;background:${sl.c};flex-shrink:0"></i><span>${esc(sl.n)}: <b>${sl.val.toLocaleString('pt-BR')}</b> (${pTxt})</span></div>`;
    }).join('');

    return `
      <div style="display:flex;flex-wrap:wrap;align-items:center;gap:16px;justify-content:center">
        <svg viewBox="0 0 260 220" style="width:240px;max-width:100%;height:auto">${paths}</svg>
        <div style="display:flex;flex-direction:column;gap:5px;max-height:200px;overflow-y:auto">${legendItems}</div>
      </div>
    `;
  }

  // Horizontal bar
  if (mode === 'hbar') {
    const barH = 22, gap = 10;
    const mx = NICE(Math.max(1, ...series.flatMap(s => s.v)));
    const H = Math.max(180, labels.length * (barH + gap) + 30);
    let rows = '';
    labels.forEach((l, i) => {
      const y = 16 + i * (barH + gap);
      const val = series[0].v[i] || 0;
      const w = (val / mx) * 280;
      const c = series[0].c;
      const valStr = isPct ? val.toFixed(1).replace('.', ',') + '%' : val.toLocaleString('pt-BR');
      rows += `
        <text x="125" y="${y + 15}" text-anchor="end" font-weight="600" fill="var(--ink)">${esc(l.slice(0, 18))}</text>
        <rect x="135" y="${y}" width="${Math.max(2, w)}" height="${barH}" fill="${c}" rx="3" data-tip-series="${esc(series[0].n)}" data-tip-color="${c}" data-tip-cat="${esc(l)}" data-tip-val="${valStr}"/>
        <text x="${142 + w}" y="${y + 15}" font-weight="700" fill="var(--ink)" font-size="10.5px">${valStr}</text>
      `;
    });
    return `<div class="chart-scroll-wrap"><svg viewBox="0 0 490 ${H}" style="min-width:460px;height:auto">${rows}</svg></div>`;
  }

  // Standard Cartesian SVG (group, stack, line, area, combo, scatter)
  const n = labels.length || 1;
  const minColWidth = (mode === 'group' || mode === 'combo') ? Math.max(82, series.length * 20 + 24) : 72;
  const L = 40, R = 18, T = 16, B = 28;
  const Wd = Math.max(520, L + R + n * minColWidth);
  const H = 220;
  const iw = Wd - L - R, ih = H - T - B;
  const bw = iw / n;

  let mx = mode === 'stack'
    ? Math.max(1, ...labels.map((_, i) => series.reduce((a, s) => a + (s.v[i] || 0), 0)))
    : Math.max(1, ...series.flatMap(s => s.v));
  mx = NICE(mx);
  const y = (v: number) => T + ih - (v / mx) * ih;
  let o = '';

  for (let i = 0; i <= 4; i++) {
    const v = (mx * i) / 4, yy = y(v);
    o += `<line x1="${L}" x2="${Wd - R}" y1="${yy}" y2="${yy}" stroke="var(--line)" stroke-dasharray="${i === 0 ? '' : '3 3'}" opacity="${i === 0 ? 1 : 0.8}"/><text x="${L - 6}" y="${yy + 3.5}" text-anchor="end">${isPct ? Math.round(v) + '%' : Math.round(v)}</text>`;
  }

  labels.forEach((l, i) => {
    o += `<text x="${L + bw * i + bw / 2}" y="${H - 8}" text-anchor="middle" font-weight="600" fill="var(--ink)">${esc(l)}</text>`;
  });

  if (mode === 'stack') {
    labels.forEach((l, i) => {
      let acc = 0;
      const tot = series.reduce((a, s) => a + (s.v[i] || 0), 0);
      const cat = categories?.[i] || l;
      series.forEach(s => {
        const v = s.v[i] || 0;
        if (!v) return;
        const h = (v / mx) * ih;
        const valStr = s.isPct ? v.toFixed(1).replace('.', ',') + '%' : v.toLocaleString('pt-BR');
        o += `<rect x="${L + bw * i + bw * 0.18}" y="${y(acc + v)}" width="${bw * 0.64}" height="${h}" fill="${s.c}" rx="3" data-tip-series="${esc(s.n)}" data-tip-color="${s.c}" data-tip-cat="${esc(cat)}" data-tip-val="${valStr}" data-tip-total="${tot.toLocaleString('pt-BR')}"/>`;
        if (h >= 14) o += `<text x="${L + bw * i + bw * 0.5}" y="${y(acc + v) + h / 2 + 3.5}" text-anchor="middle" fill="#fff" font-size="9.5px" font-weight="700">${valStr}</text>`;
        acc += v;
      });
      if (tot > 0) o += `<text x="${L + bw * i + bw * 0.5}" y="${Math.max(12, y(tot) - 5)}" text-anchor="middle" font-weight="700" fill="var(--ink)" font-size="10px">${tot.toLocaleString('pt-BR')}</text>`;
    });
  } else if (mode === 'group' || mode === 'combo') {
    const barSeries = mode === 'combo' ? series.filter(s => s.subType !== 'line') : series;
    const lineSeries = mode === 'combo' ? series.filter(s => s.subType === 'line') : [];
    const k = Math.max(1, barSeries.length);
    const w = (bw * 0.78) / k;

    labels.forEach((l, i) => {
      const cat = categories?.[i] || l;
      barSeries.forEach((s, j) => {
        const v = s.v[i] || 0;
        if (v > 0) {
          const h = (v / mx) * ih;
          const valStr = s.isPct ? v.toFixed(1).replace('.', ',') + '%' : v.toLocaleString('pt-BR');
          o += `<rect x="${L + bw * i + bw * 0.11 + w * j}" y="${y(v)}" width="${Math.max(2, w - 3)}" height="${h}" fill="${s.c}" rx="3" data-tip-series="${esc(s.n)}" data-tip-color="${s.c}" data-tip-cat="${esc(cat)}" data-tip-val="${valStr}"/>`;
          o += `<text x="${L + bw * i + bw * 0.11 + w * j + (w - 3) / 2}" y="${Math.max(12, y(v) - 5)}" text-anchor="middle" font-weight="700" font-size="9px">${valStr}</text>`;
        }
      });
    });

    lineSeries.forEach(s => {
      const pts = s.v.map((v, i) => [L + bw * i + bw / 2, y(v || 0)]);
      o += `<polyline fill="none" stroke="${s.c}" stroke-width="2.5" stroke-linecap="round" points="${pts.map(p => p.join(',')).join(' ')}"/>`;
      pts.forEach((p, i) => {
        const v = s.v[i] || 0;
        const valStr = s.isPct ? v.toFixed(1).replace('.', ',') + '%' : v.toLocaleString('pt-BR');
        o += `<circle cx="${p[0]}" cy="${p[1]}" r="4.5" fill="${s.c}" stroke="var(--card)" stroke-width="2" data-tip-series="${esc(s.n)}" data-tip-color="${s.c}" data-tip-cat="${esc(categories?.[i] || labels[i])}" data-tip-val="${valStr}"/>`;
      });
    });
  } else {
    // line, area, scatter
    series.forEach(s => {
      const pts = s.v.map((v, i) => [L + bw * i + bw / 2, y(v || 0)]);
      if (mode === 'area' && pts.length) {
        const areaD = `M ${pts[0][0]} ${T + ih} L ${pts.map(p => `${p[0]} ${p[1]}`).join(' L ')} L ${pts[pts.length - 1][0]} ${T + ih} Z`;
        o += `<path d="${areaD}" fill="${s.c}" opacity="0.18"/>`;
      }
      if (mode !== 'scatter') {
        o += `<polyline fill="none" stroke="${s.c}" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round" points="${pts.map(p => p.join(',')).join(' ')}"/>`;
      }
      pts.forEach((p, i) => {
        const v = s.v[i] || 0;
        const valStr = s.isPct ? v.toFixed(1).replace('.', ',') + '%' : v.toLocaleString('pt-BR');
        o += `<circle cx="${p[0]}" cy="${p[1]}" r="${mode === 'scatter' ? 6 : 4.5}" fill="${s.c}" stroke="var(--card)" stroke-width="2" data-tip-series="${esc(s.n)}" data-tip-color="${s.c}" data-tip-cat="${esc(categories?.[i] || labels[i])}" data-tip-val="${valStr}"/>`;
        o += `<text x="${p[0]}" y="${p[1] - 8}" text-anchor="middle" font-weight="600" font-size="9.5px">${valStr}</text>`;
      });
    });
  }

  return `
    <div class="chart-scroll-wrap" title="Arraste para os lados para visualizar todas as semanas">
      <div style="min-width:${Wd}px;width:100%">
        <svg viewBox="0 0 ${Wd} ${H}" role="img" style="width:100%;height:auto;min-height:220px;display:block">${o}</svg>
      </div>
    </div>
  `;
}

// Builder State
let activeEditId = '';
let draftChart: CustomChartConfig | null = null;

function createDefaultDraft(): CustomChartConfig {
  const pal = getSitePalette();
  return {
    id: '',
    title: 'Novo Gráfico Personalizado',
    type: 'group',
    dataSource: 'semanas',
    xAxisField: 'id',
    xAxisGrouping: 'individual',
    xAxisSort: 'asc',
    series: [
      { id: 's1', name: 'Novos leads', source: 'semanas', field: 'novos', agg: 'SUM', color: pal[0], chartSubType: 'bar' },
      { id: 's2', name: 'Consultas agendadas', source: 'semanas', field: 'cagend', agg: 'SUM', color: pal[1], chartSubType: 'bar' }
    ],
    showInDashboard: true,
    width: 'half'
  };
}

export function openChartInEditor(chartId: string, ctx: any) {
  const { cfg } = ctx;
  const pal = getSitePalette();
  const customList: CustomChartConfig[] = cfg.customCharts || [];
  const existing = customList.find(c => c.id === chartId);

  if (existing) {
    activeEditId = existing.id;
    draftChart = JSON.parse(JSON.stringify(existing));
  } else if (STANDARD_CHARTS_META[chartId]) {
    const std = STANDARD_CHARTS_META[chartId];
    activeEditId = chartId;
    const override = customList.find(c => c.defaultKey === chartId);
    if (override) {
      draftChart = JSON.parse(JSON.stringify(override));
      activeEditId = override.id;
    } else {
      // Create editable config pre-populated for this standard chart
      let series: ChartSeriesConfig[] = [];
      let ds: CustomChartConfig['dataSource'] = 'semanas';
      if (chartId === 'std_fu') {
        series = [2, 3, 4, 5].map((k, i) => ({ id: 'sfu' + k, name: 'FU ' + k, source: 'semanas', field: 'fu' + k, agg: 'SUM', color: pal[[1, 0, 2, 4][i] % pal.length] }));
      } else if (chartId === 'std_resp') {
        series = [
          { id: 'sr1', name: 'Taxa de resposta', source: 'semanas', field: 'rate_resp', agg: 'PCT', color: pal[1] },
          { id: 'sr2', name: 'Conversão FU (%)', source: 'semanas', field: 'rate_conv', agg: 'PCT', color: pal[0] }
        ];
      } else if (chartId === 'std_status') {
        series = [
          { id: 'ss1', name: 'Sem resposta', source: 'semanas', field: 'semresp', agg: 'SUM', color: pal[5] },
          { id: 'ss2', name: 'Pensando', source: 'semanas', field: 'pensando', agg: 'SUM', color: pal[2] },
          { id: 'ss3', name: 'Aguardando cadastro', source: 'semanas', field: 'aguard', agg: 'SUM', color: pal[0] },
          { id: 'ss4', name: 'Não vão seguir', source: 'semanas', field: 'nao', agg: 'SUM', color: pal[3] }
        ];
      } else if (chartId === 'std_agenda') {
        series = [
          { id: 'sa1', name: 'Previstas', source: 'semanas', field: 'cprev', agg: 'SUM', color: pal[1] },
          { id: 'sa2', name: 'Realizadas', source: 'semanas', field: 'creal', agg: 'SUM', color: pal[0] },
          { id: 'sa3', name: 'Canceladas', source: 'semanas', field: 'ccanc', agg: 'SUM', color: pal[3] },
          { id: 'sa4', name: 'Reagendadas', source: 'semanas', field: 'creag', agg: 'SUM', color: pal[2] }
        ];
      } else {
        series = [
          { id: 'sfl1', name: 'Novos leads', source: 'semanas', field: 'novos', agg: 'SUM', color: pal[0] },
          { id: 'sfl2', name: 'Respostas FU', source: 'semanas', field: 'resp', agg: 'SUM', color: pal[1] },
          { id: 'sfl3', name: 'Agendamentos FU', source: 'semanas', field: 'agend', agg: 'SUM', color: pal[2] }
        ];
      }
      draftChart = {
        id: 'ovr_' + chartId,
        title: std.title,
        type: std.type,
        dataSource: ds,
        xAxisField: 'id',
        xAxisGrouping: 'individual',
        xAxisSort: 'asc',
        series,
        showInDashboard: true,
        width: 'half',
        isDefaultOverride: true,
        defaultKey: chartId
      };
    }
  }
}

export function renderGraficosPage(container: HTMLElement, ctx: any) {
  const { cfg, saveCfg, toast, ORIG, STAT, render } = ctx;
  if (!draftChart) draftChart = createDefaultDraft();

  const pal = getSitePalette();
  const xFields = getAvailableXFields(draftChart.dataSource, ctx);
  if (!xFields.some(f => f.key === draftChart!.xAxisField)) {
    draftChart.xAxisField = xFields[0]?.key || 'id';
  }

  const customList: CustomChartConfig[] = cfg.customCharts || [];
  const sps: SpreadsheetSource[] = cfg.spreadsheets || [];

  const computed = computeChartData(draftChart, ctx);
  const previewSvg = renderExtendedChartSvg(computed.labels, computed.series, draftChart.type, computed.categories);
  const previewLeg = `<div class="lg" role="list"><b style="color:var(--ink)">Legenda:</b>${computed.series.map(x => `<span><i style="background:${x.c}"></i>${esc(x.n)}</span>`).join('')}</div>`;

  const isEditing = Boolean(activeEditId);

  container.innerHTML = `
    <div class="card">
      <div style="display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:8px">
        <div>
          <h2 style="margin:0">📊 Construtor de Gráficos (Estilo Data Studio)</h2>
          <p class="note" style="margin:2px 0 0">Cruze dados de semanas, consultas, metas e planilhas. Escolha livremente Eixo X, Eixo Y, séries, filtros e cores.</p>
        </div>
        <div style="display:flex;gap:6px">
          <button type="button" class="b s" id="gb_new_btn">➕ Novo Gráfico</button>
          <button type="button" class="b s" id="gb_palette_btn" title="Aplicar cores do tema do site às séries">🎨 Sincronizar Cores do Site</button>
        </div>
      </div>

      ${isEditing ? `<div style="margin-top:10px;padding:8px 12px;border-radius:8px;background:color-mix(in srgb,var(--b) 12%,var(--card));border:1px solid color-mix(in srgb,var(--b) 35%,var(--line));font-size:13px;display:flex;justify-content:space-between;align-items:center"><span>✏️ Editando gráfico: <b>${esc(draftChart.title)}</b></span><button type="button" class="b s" id="gb_cancel_top" style="padding:3px 9px;font-size:12px">Cancelar edição</button></div>` : ''}

      <div class="f" style="margin-top:14px;padding:14px;background:var(--bg);border:1px solid var(--line);border-radius:12px">
        <label style="grid-column:span 2">Título do Gráfico
          <input type="text" id="gb_title" value="${esc(draftChart.title)}">
        </label>

        <label>Tipo de Gráfico
          <select id="gb_type">
            <option value="group" ${draftChart.type === 'group' ? 'selected' : ''}>Colunas Agrupadas</option>
            <option value="stack" ${draftChart.type === 'stack' ? 'selected' : ''}>Colunas Empilhadas</option>
            <option value="hbar" ${draftChart.type === 'hbar' ? 'selected' : ''}>Barras Horizontais</option>
            <option value="line" ${draftChart.type === 'line' ? 'selected' : ''}>Linhas</option>
            <option value="area" ${draftChart.type === 'area' ? 'selected' : ''}>Área</option>
            <option value="combo" ${draftChart.type === 'combo' ? 'selected' : ''}>Combinado (Barra + Linha)</option>
            <option value="pie" ${draftChart.type === 'pie' ? 'selected' : ''}>Pizza</option>
            <option value="donut" ${draftChart.type === 'donut' ? 'selected' : ''}>Rosca</option>
            <option value="scatter" ${draftChart.type === 'scatter' ? 'selected' : ''}>Dispersão</option>
          </select>
        </label>

        <label>Fonte Principal (Eixo X)
          <select id="gb_ds">
            <option value="semanas" ${draftChart.dataSource === 'semanas' ? 'selected' : ''}>Semanas Lançadas</option>
            <option value="consultas" ${draftChart.dataSource === 'consultas' ? 'selected' : ''}>Consultas Registradas</option>
            <option value="cruzado" ${draftChart.dataSource === 'cruzado' ? 'selected' : ''}>Cruzamento (Semanas × Consultas × Planilhas)</option>
            <option value="planilha" ${draftChart.dataSource === 'planilha' ? 'selected' : ''}>Planilha Importada</option>
          </select>
        </label>

        <label>Campo do Eixo X (Dimensão)
          <select id="gb_x">
            ${xFields.map(f => `<option value="${esc(f.key)}" ${draftChart!.xAxisField === f.key ? 'selected' : ''}>${esc(f.label)}</option>`).join('')}
          </select>
        </label>

        <label>Agrupamento (Eixo X)
          <select id="gb_grp">
            <option value="individual" ${draftChart.xAxisGrouping === 'individual' ? 'selected' : ''}>Por semana / registro</option>
            <option value="month" ${draftChart.xAxisGrouping === 'month' ? 'selected' : ''}>Por Mês</option>
            <option value="year" ${draftChart.xAxisGrouping === 'year' ? 'selected' : ''}>Por Ano</option>
          </select>
        </label>

        <label>Ordenação
          <select id="gb_sort">
            <option value="asc" ${draftChart.xAxisSort === 'asc' ? 'selected' : ''}>Crescente (Data / A-Z)</option>
            <option value="desc" ${draftChart.xAxisSort === 'desc' ? 'selected' : ''}>Decrescente</option>
            <option value="val_desc" ${draftChart.xAxisSort === 'val_desc' ? 'selected' : ''}>Maior Valor Primeiro</option>
            <option value="val_asc" ${draftChart.xAxisSort === 'val_asc' ? 'selected' : ''}>Menor Valor Primeiro</option>
          </select>
        </label>

        <label>Largura no Painel
          <select id="gb_width">
            <option value="half" ${draftChart.width !== 'full' ? 'selected' : ''}>Padrão (1 coluna)</option>
            <option value="full" ${draftChart.width === 'full' ? 'selected' : ''}>Largo (Largura total)</option>
          </select>
        </label>
      </div>

      <!-- Filtros do Gráfico -->
      <div class="f" style="margin-top:10px;padding:10px 14px;background:var(--bg);border:1px solid var(--line);border-radius:10px">
        <label>Filtrar Origem (Consultas)
          <select id="gb_f_orig">
            <option value="">Todas as origens</option>
            ${Object.entries(ORIG).map(([k, v]) => `<option value="${k}" ${draftChart!.filterOrigin === k ? 'selected' : ''}>${esc(v)}</option>`).join('')}
          </select>
        </label>
        <label>Filtrar Status (Consultas)
          <select id="gb_f_stat">
            <option value="">Todos os status</option>
            ${Object.entries(STAT).map(([k, v]) => `<option value="${k}" ${draftChart!.filterStatus === k ? 'selected' : ''}>${esc(v)}</option>`).join('')}
          </select>
        </label>
        <label>Período De
          <input type="date" id="gb_f_from" value="${draftChart.filterFrom || ''}">
        </label>
        <label>Período Até
          <input type="date" id="gb_f_to" value="${draftChart.filterTo || ''}">
        </label>
      </div>

      <!-- Eixo Y: Séries independentes -->
      <div style="margin-top:14px">
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px">
          <h3 style="margin:0;font-size:14px">Eixo Y — Séries de Dados (Múltiplas fontes podem ser cruzadas aqui)</h3>
          <button type="button" class="b s" id="gb_add_series" style="font-size:12px;padding:4px 10px">➕ Adicionar Série</button>
        </div>
        <div id="gb_series_list" style="display:flex;flex-direction:column;gap:8px">
          ${draftChart.series.map((s, idx) => {
            const yFields = getAvailableYFields(s.source, ctx);
            if (!yFields.some(yf => yf.key === s.field)) s.field = yFields[0]?.key || '';
            return `
              <div class="row" style="margin:0;padding:8px 10px;background:var(--bg);border:1px solid var(--line);border-radius:9px;align-items:center;gap:8px">
                <span class="tag">#${idx + 1}</span>
                <label style="font-size:11px;color:var(--mut);flex:1;min-width:120px">Nome na legenda
                  <input type="text" data-gbs="${idx}" data-gbk="name" value="${esc(s.name)}">
                </label>
                <label style="font-size:11px;color:var(--mut);min-width:120px">Fonte da Série
                  <select data-gbs="${idx}" data-gbk="source">
                    <option value="semanas" ${s.source === 'semanas' ? 'selected' : ''}>Semanas</option>
                    <option value="consultas" ${s.source === 'consultas' ? 'selected' : ''}>Consultas</option>
                    <option value="planilha" ${s.source === 'planilha' ? 'selected' : ''}>Planilha</option>
                  </select>
                </label>
                <label style="font-size:11px;color:var(--mut);flex:1.3;min-width:160px">Campo / Métrica (Eixo Y)
                  <select data-gbs="${idx}" data-gbk="field">
                    ${yFields.map(yf => `<option value="${esc(yf.key)}" ${s.field === yf.key ? 'selected' : ''}>${esc(yf.label)}</option>`).join('')}
                  </select>
                </label>
                <label style="font-size:11px;color:var(--mut);min-width:95px">Operação
                  <select data-gbs="${idx}" data-gbk="agg">
                    <option value="SUM" ${s.agg === 'SUM' ? 'selected' : ''}>Soma</option>
                    <option value="AVG" ${s.agg === 'AVG' ? 'selected' : ''}>Média</option>
                    <option value="COUNT" ${s.agg === 'COUNT' ? 'selected' : ''}>Contagem</option>
                    <option value="MAX" ${s.agg === 'MAX' ? 'selected' : ''}>Máximo</option>
                    <option value="MIN" ${s.agg === 'MIN' ? 'selected' : ''}>Mínimo</option>
                    <option value="PCT" ${s.agg === 'PCT' ? 'selected' : ''}>Taxa (%)</option>
                  </select>
                </label>
                ${draftChart!.type === 'combo' ? `
                  <label style="font-size:11px;color:var(--mut);min-width:85px">Exibição
                    <select data-gbs="${idx}" data-gbk="chartSubType">
                      <option value="bar" ${s.chartSubType !== 'line' ? 'selected' : ''}>Coluna</option>
                      <option value="line" ${s.chartSubType === 'line' ? 'selected' : ''}>Linha</option>
                    </select>
                  </label>
                ` : ''}
                <label style="font-size:11px;color:var(--mut)">Cor
                  <input type="color" data-gbs="${idx}" data-gbk="color" value="${s.color || pal[idx % pal.length]}" style="width:38px;height:34px;padding:2px;cursor:pointer">
                </label>
                ${draftChart!.series.length > 1 ? `<button type="button" class="x" data-gb-del-s="${idx}" title="Remover série">×</button>` : ''}
              </div>
            `;
          }).join('')}
        </div>
      </div>

      <!-- Prévia em Tempo Real -->
      <div style="margin-top:16px;padding:16px;background:var(--card);border:1px solid var(--line);border-radius:12px">
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px">
          <span style="font-size:11.5px;font-weight:700;color:var(--mut);text-transform:uppercase">Pré-visualização em Tempo Real</span>
          <span class="note" style="font-size:11.5px">Cores e layout integrados ao Painel KOMMO</span>
        </div>
        <h3 style="margin:0 0 6px;font-size:15.5px">${esc(draftChart.title)}</h3>
        ${previewLeg}
        ${previewSvg}
      </div>

      <div class="row" style="margin-top:14px;justify-content:space-between;align-items:center">
        <label style="display:flex;align-items:center;gap:6px;font-size:13px;font-weight:600;cursor:pointer">
          <input type="checkbox" id="gb_show_dash" ${draftChart.showInDashboard ? 'checked' : ''} style="width:auto;margin:0">
          <span>Exibir no Painel Principal</span>
        </label>
        <div style="display:flex;gap:8px">
          <button type="button" class="b s" id="gb_cancel_btn">Cancelar</button>
          <button type="button" class="b" id="gb_save_btn">${isEditing ? '💾 Salvar alterações' : '💾 Salvar gráfico no painel'}</button>
        </div>
      </div>
    </div>

    <!-- EDITAR GRÁFICOS DO PAINEL -->
    <div class="card" style="margin-top:14px">
      <h2>EDITAR GRÁFICOS DO PAINEL</h2>
      <p class="note">Selecione qualquer gráfico padrão ou personalizado para editar, ocultar, duplicar ou excluir.</p>

      <div style="display:flex;flex-direction:column;gap:8px;margin-top:10px">
        ${Object.entries(STANDARD_CHARTS_META).map(([key, meta]) => {
          const override = customList.find(c => c.defaultKey === key);
          const isHidden = Boolean(cfg.hide && cfg.hide['c:' + meta.title]);
          return `
            <div class="row" style="margin:0;padding:10px 14px;background:var(--bg);border:1px solid var(--line);border-radius:10px;justify-content:space-between;align-items:center">
              <div>
                <b>${esc(override ? override.title : meta.title)}</b>
                <span class="tag" style="margin-left:6px;font-size:11px">Padrão${override ? ' (Editado)' : ''}</span>
                ${isHidden ? '<span class="tag" style="margin-left:4px;font-size:11px;color:var(--r)">Oculto</span>' : ''}
                <div class="note" style="font-size:12px;margin-top:2px">${esc(meta.desc)}</div>
              </div>
              <div style="display:flex;gap:6px;align-items:center">
                <button type="button" class="b s" data-edit-std="${key}" style="font-size:12px;padding:4px 10px">✏️ Editar</button>
                <button type="button" class="b s" data-toggle-std="${esc(meta.title)}" style="font-size:12px;padding:4px 10px">${isHidden ? '👁️ Reexibir' : '🙈 Ocultar'}</button>
                ${override ? `<button type="button" class="b s" data-reset-std="${override.id}" style="font-size:12px;padding:4px 10px;color:var(--r)">Restaurar original</button>` : ''}
              </div>
            </div>
          `;
        }).join('')}

        ${customList.filter(c => !c.isDefaultOverride).map(c => `
          <div class="row" style="margin:0;padding:10px 14px;background:var(--card);border:1px solid var(--line);border-radius:10px;justify-content:space-between;align-items:center">
            <div>
              <b>${esc(c.title)}</b>
              <span class="tag" style="margin-left:6px;font-size:11px">${c.type}</span>
              <span class="tag" style="margin-left:4px;font-size:11px">${c.showInDashboard ? 'Visível no Painel' : 'Oculto'}</span>
            </div>
            <div style="display:flex;gap:6px;align-items:center">
              <button type="button" class="b s" data-edit-cust="${c.id}" style="font-size:12px;padding:4px 10px">✏️ Editar</button>
              <button type="button" class="b s" data-dup-cust="${c.id}" style="font-size:12px;padding:4px 10px">Duplicar</button>
              <button type="button" class="b s" data-toggle-cust="${c.id}" style="font-size:12px;padding:4px 10px">${c.showInDashboard ? 'Ocultar' : 'Exibir'}</button>
              <button type="button" class="x" data-del-cust="${c.id}" title="Excluir gráfico" style="font-size:18px;padding:2px 8px">×</button>
            </div>
          </div>
        `).join('')}
      </div>
    </div>

    <!-- IMPORTAÇÃO DE PLANILHAS (.xlsx / .csv) -->
    <div class="card" style="margin-top:14px">
      <h2>Importação de Planilhas Externas (.xlsx e .csv)</h2>
      <p class="note">Importe planilhas externas (ex.: investimento em anúncios, tráfego, financeiro) para cruzar com os dados de semanas e consultas no construtor acima.</p>
      <div class="row" style="margin:10px 0">
        <input type="file" id="sp_file_input" accept=".xlsx,.xls,.csv" style="display:none">
        <button type="button" class="b" id="sp_upload_btn">📁 Importar Excel / CSV</button>
        <button type="button" class="b s" id="sp_template_btn">📄 Baixar modelo padrão (.xlsx)</button>
      </div>
      ${sps.length ? `
        <div style="display:flex;flex-direction:column;gap:6px;margin-top:10px">
          ${sps.map(sp => `
            <div class="row" style="margin:0;padding:8px 12px;background:var(--bg);border:1px solid var(--line);border-radius:8px;justify-content:space-between;align-items:center">
              <div>
                <b>${esc(sp.name)}</b> · <span class="note">${sp.rows.length} linhas · Colunas: ${esc(sp.headers.join(', '))}</span>
              </div>
              <button type="button" class="x" data-del-sp="${sp.id}" title="Remover planilha">×</button>
            </div>
          `).join('')}
        </div>
      ` : '<p class="note">Nenhuma planilha importada ainda.</p>'}
    </div>

    <!-- MÉTRICAS E CARDS PERSONALIZADOS (Movido das Configurações para cá) -->
    <div id="graf_metrics_slot">
      ${ctx.metricHtml()}
    </div>
  `;

  // Attach Metric Builder events from original code
  ctx.attachMetricBuilderEvents();

  // Attach Chart Builder events
  const bindVal = (id: string, prop: keyof CustomChartConfig, transform?: (v: any) => any) => {
    const el = container.querySelector('#' + id) as any;
    if (!el) return;
    el.oninput = el.onchange = () => {
      if (!draftChart) return;
      (draftChart as any)[prop] = transform ? transform(el.value) : el.value;
      renderGraficosPage(container, ctx);
    };
  };

  const titleInp = container.querySelector('#gb_title') as HTMLInputElement;
  if (titleInp) {
    titleInp.oninput = () => {
      if (draftChart) draftChart.title = titleInp.value;
      const h3 = container.querySelector('#gb_preview_title');
      if (h3) h3.textContent = titleInp.value;
    };
    titleInp.onchange = () => renderGraficosPage(container, ctx);
  }

  bindVal('gb_type', 'type');
  bindVal('gb_ds', 'dataSource');
  bindVal('gb_x', 'xAxisField');
  bindVal('gb_grp', 'xAxisGrouping');
  bindVal('gb_sort', 'xAxisSort');
  bindVal('gb_width', 'width');
  bindVal('gb_f_orig', 'filterOrigin');
  bindVal('gb_f_stat', 'filterStatus');
  bindVal('gb_f_from', 'filterFrom');
  bindVal('gb_f_to', 'filterTo');

  const showDashEl = container.querySelector('#gb_show_dash') as HTMLInputElement;
  if (showDashEl) {
    showDashEl.onchange = () => {
      if (draftChart) draftChart.showInDashboard = showDashEl.checked;
    };
  }

  container.querySelector('#gb_new_btn')?.addEventListener('click', () => {
    activeEditId = '';
    draftChart = createDefaultDraft();
    renderGraficosPage(container, ctx);
  });

  container.querySelector('#gb_palette_btn')?.addEventListener('click', () => {
    if (!draftChart) return;
    const p = getSitePalette();
    draftChart.series.forEach((s, i) => { s.color = p[i % p.length]; });
    renderGraficosPage(container, ctx);
    toast('Cores sincronizadas com a paleta do painel!');
  });

  const cancelAction = () => {
    activeEditId = '';
    draftChart = createDefaultDraft();
    renderGraficosPage(container, ctx);
    toast('Edição cancelada.');
  };
  container.querySelector('#gb_cancel_btn')?.addEventListener('click', cancelAction);
  container.querySelector('#gb_cancel_top')?.addEventListener('click', cancelAction);

  container.querySelector('#gb_add_series')?.addEventListener('click', () => {
    if (!draftChart) return;
    const idx = draftChart.series.length;
    draftChart.series.push({
      id: 's' + Date.now().toString(36),
      name: `Série ${idx + 1}`,
      source: draftChart.dataSource === 'consultas' ? 'consultas' : 'semanas',
      field: draftChart.dataSource === 'consultas' ? 'count_realizada' : 'creal',
      agg: 'SUM',
      color: pal[idx % pal.length],
      chartSubType: 'bar'
    });
    renderGraficosPage(container, ctx);
  });

  container.querySelectorAll('[data-gbs]').forEach((el: any) => {
    el.onchange = () => {
      if (!draftChart) return;
      const idx = parseInt(el.dataset.gbs, 10);
      const key = el.dataset.gbk;
      if (draftChart.series[idx]) {
        (draftChart.series[idx] as any)[key] = el.value;
        renderGraficosPage(container, ctx);
      }
    };
  });

  container.querySelectorAll('[data-gb-del-s]').forEach((btn: any) => {
    btn.onclick = () => {
      if (!draftChart) return;
      const idx = parseInt(btn.dataset.gbDelS, 10);
      draftChart.series.splice(idx, 1);
      renderGraficosPage(container, ctx);
    };
  });

  // Save chart
  container.querySelector('#gb_save_btn')?.addEventListener('click', async () => {
    if (!draftChart) return;
    const titleInput = container.querySelector('#gb_title') as HTMLInputElement;
    if (titleInput) draftChart.title = titleInput.value.trim() || 'Gráfico Personalizado';

    const list: CustomChartConfig[] = [...(cfg.customCharts || [])];
    const saveId = activeEditId || ('c_' + Date.now().toString(36));
    const toSave: CustomChartConfig = { ...draftChart, id: saveId };

    const existingIdx = list.findIndex(c => c.id === saveId);
    if (existingIdx >= 0) {
      list[existingIdx] = toSave;
    } else {
      list.push(toSave);
    }

    await saveCfg({ customCharts: list });
    activeEditId = '';
    draftChart = createDefaultDraft();
    renderGraficosPage(container, ctx);
    render();
    toast('Gráfico salvo no painel com sucesso!');
  });

  // Manager actions
  container.querySelectorAll('[data-edit-std]').forEach((btn: any) => {
    btn.onclick = () => {
      openChartInEditor(btn.dataset.editStd, ctx);
      renderGraficosPage(container, ctx);
      window.scrollTo({ top: 0, behavior: 'smooth' });
    };
  });

  container.querySelectorAll('[data-toggle-std]').forEach((btn: any) => {
    btn.onclick = async () => {
      const title = btn.dataset.toggleStd;
      const h = { ...(cfg.hide || {}) };
      if (h['c:' + title]) delete h['c:' + title];
      else h['c:' + title] = 1;
      await saveCfg({ hide: h });
      renderGraficosPage(container, ctx);
      render();
    };
  });

  container.querySelectorAll('[data-reset-std]').forEach((btn: any) => {
    btn.onclick = async () => {
      const id = btn.dataset.resetStd;
      const list = (cfg.customCharts || []).filter((c: any) => c.id !== id);
      await saveCfg({ customCharts: list });
      renderGraficosPage(container, ctx);
      render();
      toast('Gráfico padrão restaurado!');
    };
  });

  container.querySelectorAll('[data-edit-cust]').forEach((btn: any) => {
    btn.onclick = () => {
      openChartInEditor(btn.dataset.editCust, ctx);
      renderGraficosPage(container, ctx);
      window.scrollTo({ top: 0, behavior: 'smooth' });
    };
  });

  container.querySelectorAll('[data-dup-cust]').forEach((btn: any) => {
    btn.onclick = async () => {
      const orig = (cfg.customCharts || []).find((c: any) => c.id === btn.dataset.dupCust);
      if (!orig) return;
      const copy = { ...JSON.parse(JSON.stringify(orig)), id: 'c_' + Date.now().toString(36), title: orig.title + ' (Cópia)' };
      await saveCfg({ customCharts: [...(cfg.customCharts || []), copy] });
      renderGraficosPage(container, ctx);
      render();
      toast('Gráfico duplicado!');
    };
  });

  container.querySelectorAll('[data-toggle-cust]').forEach((btn: any) => {
    btn.onclick = async () => {
      const list = (cfg.customCharts || []).map((c: any) => c.id === btn.dataset.toggleCust ? { ...c, showInDashboard: !c.showInDashboard } : c);
      await saveCfg({ customCharts: list });
      renderGraficosPage(container, ctx);
      render();
    };
  });

  container.querySelectorAll('[data-del-cust]').forEach((btn: any) => {
    btn.onclick = async () => {
      const id = btn.dataset.delCust;
      if (!btn.dataset.confirming) {
        btn.dataset.confirming = '1';
        btn.textContent = 'Confirma?';
        btn.style.fontSize = '12px';
        setTimeout(() => {
          if (btn && btn.dataset) {
            delete btn.dataset.confirming;
            btn.textContent = '×';
            btn.style.fontSize = '18px';
          }
        }, 3500);
        return;
      }
      const list = (cfg.customCharts || []).filter((c: any) => c.id !== id);
      if (activeEditId === id) {
        activeEditId = '';
        draftChart = createDefaultDraft();
      }
      await saveCfg({ customCharts: list });
      renderGraficosPage(container, ctx);
      render();
      toast('Gráfico excluído com sucesso!');
    };
  });

  // Spreadsheet upload
  const spInp = container.querySelector('#sp_file_input') as HTMLInputElement;
  container.querySelector('#sp_upload_btn')?.addEventListener('click', () => spInp?.click());
  spInp?.addEventListener('change', async (e: any) => {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: 'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows: any[] = XLSX.utils.sheet_to_json(ws);
      if (!rows.length) return toast('Planilha vazia.');
      const headers = Object.keys(rows[0]);
      const newSp: SpreadsheetSource = {
        id: 'sp_' + Date.now().toString(36),
        name: file.name,
        uploadedAt: new Date().toISOString(),
        headers,
        rows: rows.slice(0, 2000)
      };
      await saveCfg({ spreadsheets: [...(cfg.spreadsheets || []), newSp] });
      renderGraficosPage(container, ctx);
      toast(`Planilha "${file.name}" importada (${rows.length} linhas)!`);
    } catch (err: any) {
      toast('Erro ao ler planilha: ' + err.message);
    }
  });

  container.querySelector('#sp_template_btn')?.addEventListener('click', () => {
    const ws = XLSX.utils.aoa_to_sheet([
      ['Data', 'Origem', 'Indicador', 'Quantidade'],
      ['2026-10-01', 'Instagram', 'Investimento', 500],
      ['2026-10-01', 'Google', 'Leads', 20],
      ['2026-10-08', 'Instagram', 'Investimento', 650]
    ]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Modelo');
    XLSX.writeFile(wb, 'modelo_planilha_kommo.xlsx');
  });

  container.querySelectorAll('[data-del-sp]').forEach((btn: any) => {
    btn.onclick = async () => {
      const id = btn.dataset.delSp;
      const list = (cfg.spreadsheets || []).filter((s: any) => s.id !== id);
      await saveCfg({ spreadsheets: list });
      renderGraficosPage(container, ctx);
      toast('Planilha removida.');
    };
  });
}
