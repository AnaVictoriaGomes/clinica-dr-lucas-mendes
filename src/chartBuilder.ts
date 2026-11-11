import * as XLSX from 'xlsx';

export type StudioAgg = 'SUM' | 'AVG' | 'MEDIAN' | 'COUNT' | 'COUNT_DISTINCT' | 'MIN' | 'MAX' | 'PCT';
export type StudioChartType =
  | 'group'
  | 'stack'
  | 'hbar'
  | 'line'
  | 'area'
  | 'pie'
  | 'donut'
  | 'radar'
  | 'scatter'
  | 'combo'
  | 'kpi';

export interface SeriesItem {
  id: string;
  name: string;
  source: 'semanas' | 'consultas' | 'metas' | 'planilha';
  field: string;
  agg: StudioAgg;
  color: string;
}

export interface StudioChartCfg {
  id: string;
  title: string;
  subtitle?: string;
  type: StudioChartType;
  dataSource: 'semanas' | 'consultas' | 'cruzado' | 'metas' | 'planilha';
  spreadsheetId?: string;
  xAxisField: string;
  xAxisSeries?: string[];
  xAxisGrouping: 'individual' | 'week' | 'month' | 'quarter' | 'year' | 'weekday';
  xAxisSort: 'asc' | 'desc';
  periodFilter: 'all' | 'last4' | 'last8' | 'last12' | 'last' | 'penultimate' | string;
  originFilter: string;
  statusFilter: string;
  series: SeriesItem[];
  showInDashboard: boolean;
  hidden?: boolean;
  isStandardOverride?: boolean;
  standardKey?: string;
}

export type CustomChartDef = StudioChartCfg;

function getApp(): any {
  return (window as any).KommoApp || {};
}

function getThemeColors(): string[] {
  if (typeof document === 'undefined' || typeof getComputedStyle !== 'function') {
    return ['#1f7a6d', '#3d6fa8', '#d99a1c', '#c2455a', '#7a5aa6', '#9aa8a4'];
  }
  const cs = getComputedStyle(document.documentElement);
  const readVar = (v: string, fb: string) => (cs.getPropertyValue(v) || '').trim() || fb;
  return [
    readVar('--g', '#1f7a6d'),
    readVar('--b', '#3d6fa8'),
    readVar('--a', '#d99a1c'),
    readVar('--r', '#c2455a'),
    readVar('--p', '#7a5aa6'),
    readVar('--k', '#9aa8a4')
  ];
}

function niceMax(m: number): number {
  if (m <= 0) return 10;
  const p = Math.pow(10, Math.floor(Math.log10(m)));
  const r = m / p;
  return (r <= 1 ? 1 : r <= 2 ? 2 : r <= 5 ? 5 : 10) * p;
}

function calcMedian(nums: number[]): number {
  if (!nums.length) return 0;
  const sorted = [...nums].sort((a, b) => a - b);
  const mid = Math.floor(sorted.length / 2);
  return sorted.length % 2 !== 0 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2;
}

function aggregateNums(nums: number[], agg: StudioAgg, distinctSet?: Set<string>): number {
  if (agg === 'COUNT') return nums.length;
  if (agg === 'COUNT_DISTINCT') return distinctSet ? distinctSet.size : new Set(nums.map(String)).size;
  if (!nums.length) return 0;
  if (agg === 'AVG') return nums.reduce((x, y) => x + y, 0) / nums.length;
  if (agg === 'MEDIAN') return calcMedian(nums);
  if (agg === 'MAX') return Math.max(...nums);
  if (agg === 'MIN') return Math.min(...nums);
  return nums.reduce((x, y) => x + y, 0);
}

export function computeStudioChartData(c: StudioChartCfg) {
  const app = getApp();
  const WA: any[] = app.W ? app.W() : [];
  const allAppts: any[] = Object.values(app.appts || {});
  const cfg = app.cfg || {};
  const ORIG = app.ORIG || {};
  const STAT = app.STAT || {};
  const MONTH_NAMES: string[] = app.MONTH_NAMES || ['janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho', 'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro'];
  const fd = app.fd || ((s: string) => s);
  const fdf = app.fdf || ((s: string) => s);
  const normIso = app.normIso || ((s: string) => s);
  const addDays = app.addDays || ((s: string) => s);
  const pct = app.pct || ((a: number, b: number) => b ? (a / b * 100) : 0);

  let filteredWeeks = [...WA];
  if (c.periodFilter === 'last4') filteredWeeks = WA.slice(-4);
  else if (c.periodFilter === 'last8') filteredWeeks = WA.slice(-8);
  else if (c.periodFilter === 'last12') filteredWeeks = WA.slice(-12);
  else if (c.periodFilter === 'last') filteredWeeks = WA.slice(-1);
  else if (c.periodFilter === 'penultimate') {
    filteredWeeks = WA.length >= 2 ? [WA[WA.length - 2]] : WA.slice(-1);
  } else if (c.periodFilter && c.periodFilter.startsWith('ym:')) {
    const ym = c.periodFilter.slice(3);
    filteredWeeks = WA.filter(w => w.id && w.id.startsWith(ym + '-'));
  }

  let periodStart = '';
  let periodEnd = '';
  if (c.periodFilter && c.periodFilter !== 'all' && filteredWeeks.length) {
    periodStart = filteredWeeks[0].id;
    periodEnd = addDays(filteredWeeks[filteredWeeks.length - 1].id, 6);
  }

  const filteredAppts = allAppts.filter(a => {
    if (c.originFilter && a.origin !== c.originFilter) return false;
    if (c.statusFilter && a.status !== c.statusFilter) return false;
    if (periodStart && periodEnd) {
      const dt = normIso(a.consult) || normIso(a.schedOn) || normIso(a.newDate);
      if (!dt || dt < periodStart || dt > periodEnd) return false;
    }
    return true;
  });

  const ds = c.dataSource || 'semanas';
  const xField = c.xAxisField || 'id';
  const extraXFields = (c.xAxisSeries || []).filter(f => f && f !== xField);
  const buckets: Record<string, { label: string; fullCat: string; weeks: any[]; appts: any[]; rows: any[] }> = {};

  const getGroupBucket = (iso: string, fallbackPrefix = 'Semana de ') => {
    if (!iso || !/^\d{4}-\d{2}-\d{2}$/.test(iso)) {
      return { key: 'sem_data', label: 'Sem data', fullCat: 'Sem data' };
    }
    const parts = iso.split('-');
    const mIdx = parseInt(parts[1], 10) - 1;
    if (c.xAxisGrouping === 'month') {
      return {
        key: `${parts[0]}-${parts[1]}`,
        label: (MONTH_NAMES[mIdx] || parts[1]).slice(0, 3) + '/' + parts[0].slice(2),
        fullCat: (MONTH_NAMES[mIdx] || parts[1]) + ' de ' + parts[0]
      };
    }
    if (c.xAxisGrouping === 'quarter') {
      const q = Math.floor(mIdx / 3) + 1;
      return {
        key: `${parts[0]}-Q${q}`,
        label: `T${q}/${parts[0].slice(2)}`,
        fullCat: `${q}º Trimestre de ${parts[0]}`
      };
    }
    if (c.xAxisGrouping === 'year') {
      return { key: parts[0], label: parts[0], fullCat: 'Ano ' + parts[0] };
    }
    if (c.xAxisGrouping === 'weekday') {
      const dt = new Date(Number(parts[0]), mIdx, Number(parts[2]));
      const dayNames = ['Dom', 'Seg', 'Ter', 'Qua', 'Qui', 'Sex', 'Sáb'];
      const dIdx = dt.getDay();
      return { key: String(dIdx), label: dayNames[dIdx], fullCat: 'Dia: ' + dayNames[dIdx] };
    }
    if (c.xAxisGrouping === 'week') {
      const dt = new Date(Number(parts[0]), mIdx, Number(parts[2]));
      const diff = dt.getDay() === 6 ? 0 : -(dt.getDay() + 1);
      dt.setDate(dt.getDate() + diff);
      const satIso = `${dt.getFullYear()}-${String(dt.getMonth() + 1).padStart(2, '0')}-${String(dt.getDate()).padStart(2, '0')}`;
      return { key: satIso, label: fd(satIso), fullCat: 'Semana de ' + fdf(satIso) };
    }
    return { key: iso, label: fd(iso), fullCat: fallbackPrefix + fdf(iso) };
  };

  const REG: Record<string, string> = { '': 'Nenhum', sim: 'Conseguimos reagendar', contato: 'Em contato', nao: 'Não conseguiu' };
  const STAT_NEW: Record<string, string> = { '': 'Pendente', compareceu: 'Compareceu', nao_compareceu: 'Não compareceu', aguardando: 'Aguardando data', reagendou: 'Reagendou novamente', cancelou: 'Cancelou' };

  const resolveApptField = (a: any, field: string) => {
    if (field === 'origin') {
      const lbl = ORIG[a.origin] || a.origin || 'Outro';
      return { key: a.origin || 'outro', label: lbl, fullCat: 'Origem: ' + lbl };
    }
    if (field === 'status') {
      const lbl = STAT[a.status] || a.status || 'Agendada';
      return { key: a.status || 'agendada', label: lbl, fullCat: 'Status: ' + lbl };
    }
    if (field === 'regained') {
      const k = a.regained || '';
      const lbl = REG[k] || k || 'Nenhum';
      return { key: k, label: lbl, fullCat: 'Status Contato: ' + lbl };
    }
    if (field === 'statusNewDate') {
      const k = a.statusNewDate || '';
      const lbl = STAT_NEW[k] || k || 'Pendente';
      return { key: k, label: lbl, fullCat: 'Status Nova Data: ' + lbl };
    }
    if (field === 'consult' || field === 'schedOn' || field === 'newDate') {
      const iso = normIso(a[field] || (field === 'newDate' ? '' : (a.consult || a.schedOn)));
      return getGroupBucket(iso, '');
    }
    const v = String(a[field] || '—');
    return { key: v, label: v, fullCat: v };
  };

  if (ds === 'semanas' || ds === 'metas') {
    const goals = app.getUnifiedGoalsList ? app.getUnifiedGoalsList() : [];
    filteredWeeks.forEach(w => {
      const bInfo = getGroupBucket(w.id, 'Semana de ');
      const extraParts = extraXFields.map(ef => String(w[ef] ?? '—'));
      const fullKey = extraParts.length ? `${bInfo.key}__${extraParts.join('__')}` : bInfo.key;
      const fullLabel = extraParts.length ? `${bInfo.label} · ${extraParts.join(' · ')}` : bInfo.label;
      const fullCat = extraParts.length ? `${bInfo.fullCat} · ${extraParts.join(' · ')}` : bInfo.fullCat;
      if (!buckets[fullKey]) buckets[fullKey] = { label: fullLabel, fullCat, weeks: [], appts: [], rows: goals };
      buckets[fullKey].weeks.push(w);
    });
  } else if (ds === 'consultas') {
    filteredAppts.forEach(a => {
      const mainDim = resolveApptField(a, xField);
      const extraDims = extraXFields.map(ef => resolveApptField(a, ef));
      const key = [mainDim.key, ...extraDims.map(d => d.key)].join('__');
      const label = [mainDim.label, ...extraDims.map(d => d.label)].join(' · ');
      const fullCat = [mainDim.fullCat, ...extraDims.map(d => d.fullCat)].join(' · ');

      if (!buckets[key]) buckets[key] = { label, fullCat, weeks: [], appts: [], rows: [] };
      buckets[key].appts.push(a);
    });
  } else if (ds === 'cruzado') {
    filteredWeeks.forEach((w, idx) => {
      const start = w.id;
      let end = addDays(w.id, 6);
      if (filteredWeeks[idx + 1] && filteredWeeks[idx + 1].id <= end) {
        end = addDays(filteredWeeks[idx + 1].id, -1);
      }
      const matchedAppts = filteredAppts.filter(a => {
        const dt = normIso(a.consult) || normIso(a.schedOn);
        return dt >= start && dt <= end;
      });

      const bInfo = getGroupBucket(w.id, 'Semana de ');
      if (extraXFields.length > 0 && matchedAppts.length > 0) {
        matchedAppts.forEach(a => {
          const extraDims = extraXFields.map(ef => resolveApptField(a, ef));
          const key = `${bInfo.key}__${extraDims.map(d => d.key).join('__')}`;
          const label = `${bInfo.label} · ${extraDims.map(d => d.label).join(' · ')}`;
          const fullCat = `${bInfo.fullCat} · ${extraDims.map(d => d.fullCat).join(' · ')}`;
          if (!buckets[key]) buckets[key] = { label, fullCat, weeks: [w], appts: [], rows: [] };
          buckets[key].appts.push(a);
        });
      } else {
        if (!buckets[bInfo.key]) buckets[bInfo.key] = { label: bInfo.label, fullCat: bInfo.fullCat, weeks: [], appts: [], rows: [] };
        buckets[bInfo.key].weeks.push(w);
        buckets[bInfo.key].appts.push(...matchedAppts);
      }
    });
  } else if (ds === 'planilha') {
    const sheets = cfg.spreadsheets || [];
    const sp = sheets.find((s: any) => s.id === c.spreadsheetId) || sheets[0];
    if (sp && Array.isArray(sp.rows)) {
      sp.rows.forEach((r: any, idx: number) => {
        const rawVal = r[xField];
        const baseKey = rawVal !== undefined && rawVal !== null && rawVal !== '' ? String(rawVal) : `Linha ${idx + 1}`;
        const extraVals = extraXFields.map(ef => String(r[ef] ?? '—'));
        const key = extraVals.length ? `${baseKey}__${extraVals.join('__')}` : baseKey;
        const label = extraVals.length ? `${baseKey.slice(0, 14)} · ${extraVals.map(v => v.slice(0, 10)).join(' · ')}` : baseKey.slice(0, 18);
        const fullCat = extraVals.length ? `${baseKey} · ${extraVals.join(' · ')}` : baseKey;
        if (!buckets[key]) buckets[key] = { label, fullCat, weeks: [], appts: [], rows: [] };
        buckets[key].rows.push(r);
      });
    }
  }

  const sortedKeys = Object.keys(buckets).sort((a, b) => {
    return c.xAxisSort === 'desc' ? b.localeCompare(a) : a.localeCompare(b);
  });

  const labels = sortedKeys.map(k => buckets[k].label);
  const categories = sortedKeys.map(k => buckets[k].fullCat);
  const palette = getThemeColors();

  const series = (c.series || []).map((s, idx) => {
    const rawValues = sortedKeys.map(k => {
      const b = buckets[k];

      if (ds === 'planilha' || s.source === 'planilha') {
        if (s.agg === 'COUNT') return b.rows.length;
        if (s.agg === 'COUNT_DISTINCT') {
          const dist = new Set(b.rows.map(r => String(r[s.field] ?? '').trim()).filter(Boolean));
          return dist.size;
        }
        const nums = b.rows.map(r => parseFloat(String(r[s.field] ?? '0').replace(',', '.'))).filter(n => !isNaN(n));
        return aggregateNums(nums, s.agg);
      }

      if (ds === 'metas' && s.field.startsWith('goal_target_')) {
        const gId = s.field.replace('goal_target_', '');
        const goal = (b.rows || []).find((g: any) => g.id === gId);
        return goal ? Number(goal.target) || 0 : 0;
      }

      if (ds === 'consultas' || s.source === 'consultas') {
        let matched = b.appts;
        if (s.field === 'appts_agendada') matched = b.appts.filter(a => a.status === 'agendada');
        else if (s.field === 'appts_realizada') matched = b.appts.filter(a => a.status === 'realizada' || a.statusNewDate === 'compareceu');
        else if (s.field === 'appts_cancelada') matched = b.appts.filter(a => a.status === 'cancelada' || a.statusNewDate === 'cancelou');
        else if (s.field === 'appts_reagendada') matched = b.appts.filter(a => a.status === 'reagendada' || a.statusNewDate === 'reagendou');
        else if (s.field === 'appts_nao_compareceu') matched = b.appts.filter(a => a.status === 'nao_compareceu' || a.statusNewDate === 'nao_compareceu');
        else if (s.field === 'appts_recuperadas') matched = b.appts.filter(a => a.regained === 'sim');
        else if (s.field.startsWith('appts_orig_')) {
          const origKey = s.field.replace('appts_orig_', '');
          matched = b.appts.filter(a => a.origin === origKey);
        }

        if (s.agg === 'COUNT_DISTINCT') {
          const dist = new Set(matched.map(a => String(a.patient || '').trim().toLowerCase()).filter(Boolean));
          return dist.size;
        }
        return matched.length;
      }

      if (s.field === 'msgs') {
        const vals = b.weeks.map(w => (w.fu2 || 0) + (w.fu3 || 0) + (w.fu4 || 0) + (w.fu5 || 0));
        return aggregateNums(vals, s.agg);
      }
      if (s.field === 'taxa_resp') {
        const m = b.weeks.reduce((x, w) => x + (w.fu2 || 0) + (w.fu3 || 0) + (w.fu4 || 0) + (w.fu5 || 0), 0);
        const r = b.weeks.reduce((x, w) => x + (w.resp || 0), 0);
        return pct(r, m);
      }
      if (s.field === 'taxa_conv') {
        const r = b.weeks.reduce((x, w) => x + (w.resp || 0), 0);
        const a = b.weeks.reduce((x, w) => x + (w.agend || 0), 0);
        return pct(a, r);
      }
      if (s.field === 'taxa_agend_geral') {
        const at = b.weeks.reduce((x, w) => x + (w.ativas || 0), 0) / (b.weeks.length || 1);
        const ca = b.weeks.reduce((x, w) => x + (w.cagend || 0), 0);
        return pct(ca, at);
      }

      const vals = b.weeks.map(w => Number(w[s.field]) || 0);
      const dist = new Set(vals.map(String));
      return aggregateNums(vals, s.agg, dist);
    });

    let finalValues = rawValues;
    if (s.agg === 'PCT' && s.field !== 'taxa_resp' && s.field !== 'taxa_conv' && s.field !== 'taxa_agend_geral') {
      const totalSum = rawValues.reduce((a, b) => a + Math.max(0, b), 0);
      finalValues = rawValues.map(v => totalSum > 0 ? (Math.max(0, v) / totalSum) * 100 : 0);
    }

    const isPct = s.agg === 'PCT' || s.field === 'taxa_resp' || s.field === 'taxa_conv' || s.field === 'taxa_agend_geral';
    return {
      n: (s.name && String(s.name).trim()) || s.field || `Série ${idx + 1}`,
      c: s.color || palette[idx % palette.length],
      v: finalValues.map(v => Math.round(v * 100) / 100),
      isPct
    };
  });

  return { labels, series, categories };
}

export const computeChartData = computeStudioChartData;

export function renderStudioSvg(c: StudioChartCfg): { html: string } {
  const app = getApp();
  const esc = app.esc || ((s: any) => String(s ?? ''));
  const pf = app.pf || ((v: number) => v.toFixed(1).replace('.', ',') + '%');
  const { labels, series, categories } = computeStudioChartData(c);

  if (!labels.length || !series.length) {
    return {
      html: `<p class="note" style="padding:16px 0">Sem dados disponíveis para os filtros e dimensões selecionados.</p>`
    };
  }

  const isPctAny = series.some(s => s.isPct);

  if (c.type === 'kpi') {
    const s = series[0];
    const totalVal = s.v.reduce((a, b) => a + b, 0);
    const displayVal = s.isPct
      ? pf(s.v.length ? totalVal / s.v.length : 0)
      : totalVal.toLocaleString('pt-BR');
    return {
      html: `
        <div class="card kpi" style="border-left:4px solid ${s.c}">
          <span>${esc(c.title)}</span>
          <b>${displayVal}</b>
          <small>${esc(s.n)} (${labels.length} ponto(s))</small>
        </div>
      `
    };
  }

  const leg = `<div class="lg" role="list" aria-label="Legenda"><b style="color:var(--ink)">Legenda:</b>${series.map(x => `<span role="listitem"><i style="background:${x.c}"></i>${esc(x.n)}</span>`).join('')}</div>`;

  if (c.type === 'pie' || c.type === 'donut') {
    const s = series[0];
    const total = s.v.reduce((a, b) => a + Math.max(0, b), 0);
    const palette = getThemeColors();
    const cx = 140, cy = 110, r = 85, innerR = c.type === 'donut' ? 46 : 0;
    let startAngle = 0;
    let paths = '';
    const pieLegend: string[] = [];

    labels.forEach((lbl, i) => {
      const val = Math.max(0, s.v[i] || 0);
      const col = palette[i % palette.length];
      const pctVal = total > 0 ? (val / total) * 100 : 0;
      const pctStr = pf(pctVal);
      pieLegend.push(`<span role="listitem"><i style="background:${col}"></i>${esc(lbl)}: <b>${val.toLocaleString('pt-BR')}</b> (${pctStr})</span>`);
      if (val <= 0 || total <= 0) return;

      const angle = (val / total) * 2 * Math.PI;
      const endAngle = startAngle + angle;
      const largeArc = angle > Math.PI ? 1 : 0;
      const x1 = cx + r * Math.cos(startAngle), y1 = cy + r * Math.sin(startAngle);
      const x2 = cx + r * Math.cos(endAngle), y2 = cy + r * Math.sin(endAngle);
      const x3 = cx + innerR * Math.cos(endAngle), y3 = cy + innerR * Math.sin(endAngle);
      const x4 = cx + innerR * Math.cos(startAngle), y4 = cy + innerR * Math.sin(startAngle);

      const d = angle >= 2 * Math.PI - 0.001
        ? `M ${cx - r} ${cy} A ${r} ${r} 0 1 0 ${cx + r} ${cy} A ${r} ${r} 0 1 0 ${cx - r} ${cy} Z`
        : innerR > 0
          ? `M ${x1} ${y1} A ${r} ${r} 0 ${largeArc} 1 ${x2} ${y2} L ${x3} ${y3} A ${innerR} ${innerR} 0 ${largeArc} 0 ${x4} ${y4} Z`
          : `M ${cx} ${cy} L ${x1} ${y1} A ${r} ${r} 0 ${largeArc} 1 ${x2} ${y2} Z`;

      paths += `<path d="${d}" fill="${col}" stroke="var(--card)" stroke-width="1.5"
        data-tip-series="${esc(s.n)}"
        data-tip-color="${col}"
        data-tip-cat="${esc(categories[i] || lbl)}"
        data-tip-val="${val.toLocaleString('pt-BR')} (${pctStr})"></path>`;
      startAngle = endAngle;
    });

    return {
      html: `
        <div class="lg" role="list"><b style="color:var(--ink)">Legenda:</b>${pieLegend.join('')}</div>
        <div class="chart-scroll-wrap">
          <svg viewBox="0 0 280 220" style="max-width:280px;height:auto;margin:auto;display:block">${paths}</svg>
        </div>
      `
    };
  }

  if (c.type === 'hbar') {
    const s = series[0];
    const barH = 22, pad = 12;
    const mx = Math.max(1, ...s.v);
    const totalH = Math.max(180, labels.length * (barH + pad) + 36);
    let rows = '';
    labels.forEach((lbl, i) => {
      const val = s.v[i] || 0;
      const w = (val / mx) * 230;
      const yPos = 20 + i * (barH + pad);
      const valStr = s.isPct ? pf(val) : val.toLocaleString('pt-BR');
      rows += `
        <text x="118" y="${yPos + 15}" text-anchor="end" font-weight="600" fill="var(--ink)">${esc(lbl)}</text>
        <rect x="126" y="${yPos}" width="${Math.max(2, w)}" height="${barH}" fill="${s.c}" rx="3"
          data-tip-series="${esc(s.n)}"
          data-tip-color="${s.c}"
          data-tip-cat="${esc(categories[i] || lbl)}"
          data-tip-val="${valStr}"></rect>
        <text x="${132 + w}" y="${yPos + 15}" font-weight="700" fill="var(--ink)" font-size="10.5px">${valStr}</text>
      `;
    });
    return {
      html: `${leg}<div class="chart-scroll-wrap"><svg viewBox="0 0 440 ${totalH}" style="min-width:440px;height:auto;display:block">${rows}</svg></div>`
    };
  }

  // Radar Chart
  if (c.type === 'radar') {
    const cx = 175, cy = 125, r = 78;
    const count = Math.max(3, labels.length);
    const maxVal = Math.max(1, ...series.flatMap(s => s.v));
    let rings = '';
    for (let ring = 1; ring <= 4; ring++) {
      const rr = (r * ring) / 4;
      const rPts = Array.from({ length: count }, (_, i) => {
        const ang = (Math.PI * 2 * i) / count - Math.PI / 2;
        return `${(cx + rr * Math.cos(ang)).toFixed(1)},${(cy + rr * Math.sin(ang)).toFixed(1)}`;
      }).join(' ');
      rings += `<polygon points="${rPts}" fill="none" stroke="var(--line)" stroke-width="1"/>`;
    }
    let axes = '';
    labels.forEach((lbl, i) => {
      const ang = (Math.PI * 2 * i) / count - Math.PI / 2;
      const ax = cx + r * Math.cos(ang), ay = cy + r * Math.sin(ang);
      const lx = cx + (r + 16) * Math.cos(ang), ly = cy + (r + 16) * Math.sin(ang);
      const anchor = Math.abs(Math.cos(ang)) < 0.2 ? 'middle' : (Math.cos(ang) > 0 ? 'start' : 'end');
      axes += `<line x1="${cx}" y1="${cy}" x2="${ax.toFixed(1)}" y2="${ay.toFixed(1)}" stroke="var(--line)" stroke-width="1"/>`;
      axes += `<text x="${lx.toFixed(1)}" y="${(ly + 3).toFixed(1)}" text-anchor="${anchor}" font-size="9.5px" font-weight="600" fill="var(--ink)">${esc(lbl)}</text>`;
    });
    let polys = '';
    series.forEach(s => {
      const pts = labels.map((_, i) => {
        const val = Math.max(0, s.v[i] || 0);
        const rr = (val / maxVal) * r;
        const ang = (Math.PI * 2 * i) / count - Math.PI / 2;
        return [cx + rr * Math.cos(ang), cy + rr * Math.sin(ang)];
      });
      polys += `<polygon points="${pts.map(p => `${p[0].toFixed(1)},${p[1].toFixed(1)}`).join(' ')}" fill="${s.c}" fill-opacity="0.22" stroke="${s.c}" stroke-width="2"/>`;
      pts.forEach((p, i) => {
        const val = s.v[i] || 0;
        const valStr = s.isPct ? pf(val) : val.toLocaleString('pt-BR');
        polys += `<circle cx="${p[0].toFixed(1)}" cy="${p[1].toFixed(1)}" r="4" fill="${s.c}" stroke="var(--card)" stroke-width="1.5"
          data-tip-series="${esc(s.n)}" data-tip-color="${s.c}" data-tip-cat="${esc(categories[i] || labels[i])}" data-tip-val="${valStr}"></circle>`;
      });
    });
    return {
      html: `${leg}<div class="chart-scroll-wrap"><svg viewBox="0 0 350 250" style="max-width:350px;height:auto;margin:auto;display:block">${rings}${axes}${polys}</svg></div>`
    };
  }

  const n = labels.length || 1;
  const minColWidth = (c.type === 'group' || c.type === 'combo') ? Math.max(86, (series.length * 20) + 26) : 76;
  const L = 38, R = 18, T = 14, B = 28;
  const Wd = Math.max(520, L + R + n * minColWidth);
  const H = 220;
  const iw = Wd - L - R, ih = H - T - B;
  const bw = iw / n;

  let mx = c.type === 'stack'
    ? Math.max(1, ...labels.map((_, i) => series.reduce((a, s) => a + (s.v[i] || 0), 0)))
    : Math.max(1, ...series.flatMap(s => s.v));
  mx = niceMax(mx);
  const y = (v: number) => T + ih - (v / mx) * ih;
  let o = '';

  for (let i = 0; i <= 4; i++) {
    const v = mx * i / 4, yy = y(v);
    o += `<line x1="${L}" x2="${Wd - R}" y1="${yy}" y2="${yy}" stroke="var(--line)" stroke-dasharray="${i === 0 ? '' : '3 3'}" opacity="${i === 0 ? 1 : 0.8}"/><text x="${L - 6}" y="${yy + 3.5}" text-anchor="end">${isPctAny ? Math.round(v) + '%' : Math.round(v)}</text>`;
  }

  labels.forEach((l, i) => {
    o += `<text x="${L + bw * i + bw / 2}" y="${H - 8}" text-anchor="middle" font-weight="600" fill="var(--ink)">${esc(l)}</text>`;
  });

  if (c.type === 'stack') {
    labels.forEach((l, i) => {
      let acc = 0;
      const stackTotal = series.reduce((a, s) => a + (s.v[i] || 0), 0);
      const cat = categories[i] || l;
      series.forEach(s => {
        const v = s.v[i] || 0;
        if (!v) return;
        const h = (v / mx) * ih;
        const valStr = s.isPct ? pf(v) : v.toLocaleString('pt-BR');
        const totStr = stackTotal.toLocaleString('pt-BR');
        o += `<rect x="${L + bw * i + bw * 0.18}" y="${y(acc + v)}" width="${bw * 0.64}" height="${h}" fill="${s.c}" rx="3"
          data-tip-series="${esc(s.n)}" data-tip-color="${s.c}" data-tip-cat="${esc(cat)}" data-tip-val="${valStr}" data-tip-total="${totStr}"></rect>`;
        if (h >= 14) {
          o += `<text x="${L + bw * i + bw * 0.5}" y="${y(acc + v) + h / 2 + 3.5}" text-anchor="middle" fill="#fff" font-size="9.5px" font-weight="700">${valStr}</text>`;
        }
        acc += v;
      });
      if (stackTotal > 0) {
        o += `<text x="${L + bw * i + bw * 0.5}" y="${Math.max(12, y(stackTotal) - 5)}" text-anchor="middle" font-weight="700" fill="var(--ink)" font-size="10.5px">${stackTotal.toLocaleString('pt-BR')}</text>`;
      }
    });
  } else if (c.type === 'group') {
    const k = series.length;
    const w = (bw * 0.78) / k;
    labels.forEach((l, i) => {
      const cat = categories[i] || l;
      series.forEach((s, j) => {
        const v = s.v[i] || 0;
        if (v <= 0) return;
        const h = (v / mx) * ih;
        const valStr = s.isPct ? pf(v) : v.toLocaleString('pt-BR');
        o += `<rect x="${L + bw * i + bw * 0.11 + w * j}" y="${y(v)}" width="${Math.max(2, w - 3)}" height="${h}" fill="${s.c}" rx="3"
          data-tip-series="${esc(s.n)}" data-tip-color="${s.c}" data-tip-cat="${esc(cat)}" data-tip-val="${valStr}"></rect>`;
        o += `<text x="${L + bw * i + bw * 0.11 + w * j + (w - 3) / 2}" y="${Math.max(12, y(v) - 5)}" text-anchor="middle" font-weight="700" font-size="9.5px">${valStr}</text>`;
      });
    });
  } else if (c.type === 'combo') {
    // First series as bars, remaining series as lines
    const barSeries = series.slice(0, Math.max(1, Math.ceil(series.length / 2)));
    const lineSeries = series.slice(barSeries.length);
    const k = barSeries.length;
    const w = (bw * 0.65) / k;
    labels.forEach((l, i) => {
      const cat = categories[i] || l;
      barSeries.forEach((s, j) => {
        const v = s.v[i] || 0;
        if (v <= 0) return;
        const h = (v / mx) * ih;
        const valStr = s.isPct ? pf(v) : v.toLocaleString('pt-BR');
        o += `<rect x="${L + bw * i + bw * 0.175 + w * j}" y="${y(v)}" width="${Math.max(2, w - 3)}" height="${h}" fill="${s.c}" opacity="0.88" rx="3"
          data-tip-series="${esc(s.n)}" data-tip-color="${s.c}" data-tip-cat="${esc(cat)}" data-tip-val="${valStr}"></rect>`;
        o += `<text x="${L + bw * i + bw * 0.175 + w * j + (w - 3) / 2}" y="${Math.max(12, y(v) - 5)}" text-anchor="middle" font-weight="700" font-size="9.5px">${valStr}</text>`;
      });
    });
    lineSeries.forEach(s => {
      const pts = s.v.map((v, i) => [L + bw * i + bw / 2, y(v || 0)]);
      o += `<polyline fill="none" stroke="${s.c}" stroke-width="2.6" stroke-linecap="round" stroke-linejoin="round" points="${pts.map(p => p.join(',')).join(' ')}"/>`;
      pts.forEach((p, i) => {
        const v = s.v[i] || 0;
        const valStr = s.isPct ? pf(v) : v.toLocaleString('pt-BR');
        const cat = categories[i] || labels[i];
        o += `<circle cx="${p[0]}" cy="${p[1]}" r="4.5" fill="${s.c}" stroke="var(--card)" stroke-width="2"
          data-tip-series="${esc(s.n)}" data-tip-color="${s.c}" data-tip-cat="${esc(cat)}" data-tip-val="${valStr}"></circle>`;
      });
    });
  } else if (c.type === 'scatter') {
    series.forEach((s, sIdx) => {
      s.v.forEach((v, i) => {
        const cx = L + bw * i + bw / 2 + (sIdx - (series.length - 1) / 2) * 8;
        const cy = y(v || 0);
        const valStr = s.isPct ? pf(v) : v.toLocaleString('pt-BR');
        const cat = categories[i] || labels[i];
        o += `<circle cx="${cx}" cy="${cy}" r="6" fill="${s.c}" opacity="0.85" stroke="var(--card)" stroke-width="2"
          data-tip-series="${esc(s.n)}" data-tip-color="${s.c}" data-tip-cat="${esc(cat)}" data-tip-val="${valStr}"></circle>`;
        o += `<text x="${cx}" y="${cy - 9}" text-anchor="middle" font-weight="600" font-size="9px">${valStr}</text>`;
      });
    });
  } else {
    series.forEach(s => {
      const pts = s.v.map((v, i) => [L + bw * i + bw / 2, y(v || 0)]);
      if (c.type === 'area' && pts.length) {
        const areaPath = `M ${pts[0][0]} ${T + ih} L ${pts.map(p => p.join(' ')).join(' L ')} L ${pts[pts.length - 1][0]} ${T + ih} Z`;
        o += `<path d="${areaPath}" fill="${s.c}" opacity="0.18"/>`;
      }
      o += `<polyline fill="none" stroke="${s.c}" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round" points="${pts.map(p => p.join(',')).join(' ')}"/>`;
      pts.forEach((p, i) => {
        const v = s.v[i] || 0;
        const valStr = s.isPct ? pf(v) : v.toLocaleString('pt-BR');
        const cat = categories[i] || labels[i];
        o += `<circle cx="${p[0]}" cy="${p[1]}" r="4.5" fill="${s.c}" stroke="var(--card)" stroke-width="2"
          data-tip-series="${esc(s.n)}" data-tip-color="${s.c}" data-tip-cat="${esc(cat)}" data-tip-val="${valStr}"></circle>`;
        o += `<text x="${p[0]}" y="${p[1] - 8}" text-anchor="middle" font-weight="600" font-size="9.5px">${valStr}</text>`;
      });
    });
  }

  return {
    html: `
      ${leg}
      <div class="chart-scroll-wrap" title="Arraste para os lados para visualizar todas as semanas">
        <div style="min-width:${Wd}px;width:100%">
          <svg viewBox="0 0 ${Wd} ${H}" role="img" style="width:100%;height:auto;min-height:220px;display:block">${o}</svg>
        </div>
      </div>
    `
  };
}

export const renderVisualChart = renderStudioSvg;

let editingChartId: string | null = null;
let draftChart: StudioChartCfg | null = null;
let confirmDeleteId: string | null = null;

function createDefaultDraft(): StudioChartCfg {
  const palette = getThemeColors();
  return {
    id: '',
    title: 'Novo Gráfico Personalizado',
    type: 'group',
    dataSource: 'semanas',
    xAxisField: 'id',
    xAxisSeries: [],
    xAxisGrouping: 'individual',
    xAxisSort: 'asc',
    periodFilter: 'all',
    originFilter: '',
    statusFilter: '',
    series: [
      { id: 's1', name: 'Novos leads', source: 'semanas', field: 'novos', agg: 'SUM', color: palette[0] },
      { id: 's2', name: 'Consultas agendadas', source: 'semanas', field: 'cagend', agg: 'SUM', color: palette[1] }
    ],
    showInDashboard: true
  };
}

function getXAxisOptions(ds: string, spreadsheetId?: string): { value: string; label: string }[] {
  const app = getApp();
  const cfg = app.cfg || {};
  if (ds === 'consultas') {
    return [
      { value: 'origin', label: 'Origem do Paciente (Veio de)' },
      { value: 'status', label: 'Status da Consulta' },
      { value: 'consult', label: 'Data da Consulta' },
      { value: 'schedOn', label: 'Data de Agendamento' },
      { value: 'newDate', label: 'Nova Data (Reagendamento)' },
      { value: 'regained', label: 'Status Contato (Reagendamento)' },
      { value: 'statusNewDate', label: 'Status da Nova Data' }
    ];
  }
  if (ds === 'cruzado') {
    return [
      { value: 'id', label: 'Data da Semana (Cronológica)' },
      { value: 'origin', label: 'Origem do Paciente (Veio de)' },
      { value: 'status', label: 'Status da Consulta' }
    ];
  }
  if (ds === 'planilha') {
    const sheets = cfg.spreadsheets || [];
    const sp = sheets.find((s: any) => s.id === spreadsheetId) || sheets[0];
    if (sp && Array.isArray(sp.headers)) {
      return sp.headers.map((h: string) => ({ value: h, label: h }));
    }
    return [{ value: 'col1', label: 'Carregue uma planilha abaixo' }];
  }
  return [
    { value: 'id', label: 'Data da Semana (Cronológica)' }
  ];
}

function getYFieldOptions(ds: string, spreadsheetId?: string): { value: string; label: string; group: string; source: SeriesItem['source'] }[] {
  const app = getApp();
  const cfg = app.cfg || {};
  const FIELDS: [string, string][] = app.FIELDS || [];
  const lab = app.lab || ((k: string) => k);
  const ORIG = app.ORIG || {};
  const opts: { value: string; label: string; group: string; source: SeriesItem['source'] }[] = [];
  const seen = new Set<string>();
  const pushOpt = (o: { value: string; label: string; group: string; source: SeriesItem['source'] }) => {
    if (seen.has(o.value)) return;
    seen.add(o.value);
    opts.push(o);
  };

  if (ds === 'semanas' || ds === 'cruzado' || ds === 'metas') {
    FIELDS.forEach(([k]) => {
      pushOpt({ value: k, label: lab(k), group: 'Campos Semanais', source: 'semanas' });
    });
    pushOpt({ value: 'cprev', label: 'Consultas previstas da semana', group: 'Campos Semanais', source: 'semanas' });
    pushOpt({ value: 'msgs', label: 'Mensagens de follow-up (FU 2 a 5)', group: 'Campos Semanais', source: 'semanas' });
    pushOpt({ value: 'taxa_resp', label: 'Taxa de Resposta do Follow-up (%)', group: 'Taxas Calculadas', source: 'semanas' });
    pushOpt({ value: 'taxa_conv', label: 'Conversão do Follow-up (%)', group: 'Taxas Calculadas', source: 'semanas' });
    pushOpt({ value: 'taxa_agend_geral', label: 'Taxa Geral de Agendamento (%)', group: 'Taxas Calculadas', source: 'semanas' });
  }

  if (ds === 'metas') {
    const goals = app.getUnifiedGoalsList ? app.getUnifiedGoalsList() : [];
    goals.forEach((g: any) => {
      pushOpt({
        value: 'goal_target_' + g.id,
        label: `Meta: ${g.label || g.metric} (${g.target}${g.unit === 'pct' ? '%' : ''})`,
        group: 'Linhas de Meta',
        source: 'semanas'
      });
    });
  }

  if (ds === 'consultas' || ds === 'cruzado') {
    pushOpt({ value: 'appts_total', label: 'Total de Consultas', group: 'Consultas', source: 'consultas' });
    pushOpt({ value: 'appts_agendada', label: 'Consultas Agendadas (Pendentes)', group: 'Consultas', source: 'consultas' });
    pushOpt({ value: 'appts_realizada', label: 'Consultas Realizadas (Compareceu)', group: 'Consultas', source: 'consultas' });
    pushOpt({ value: 'appts_cancelada', label: 'Consultas Canceladas', group: 'Consultas', source: 'consultas' });
    pushOpt({ value: 'appts_reagendada', label: 'Consultas Reagendadas', group: 'Consultas', source: 'consultas' });
    pushOpt({ value: 'appts_nao_compareceu', label: 'Faltou (Não compareceu)', group: 'Consultas', source: 'consultas' });
    pushOpt({ value: 'appts_recuperadas', label: 'Cancelados Recuperados', group: 'Consultas', source: 'consultas' });
    Object.entries(ORIG).forEach(([k, n]) => {
      pushOpt({ value: 'appts_orig_' + k, label: `Consultas de "${n}"`, group: 'Consultas por Origem', source: 'consultas' });
    });
  }

  if (ds === 'planilha') {
    const sheets = cfg.spreadsheets || [];
    const sp = sheets.find((s: any) => s.id === spreadsheetId) || sheets[0];
    if (sp && Array.isArray(sp.headers)) {
      sp.headers.forEach((h: string) => {
        opts.push({ value: h, label: h, group: 'Colunas da Planilha', source: 'planilha' });
      });
    }
  }

  return opts;
}

export function startEditingChart(chartId: string) {
  const app = getApp();
  const cfg = app.cfg || {};
  const list: StudioChartCfg[] = cfg.customCharts || [];
  const found = list.find(c => c.id === chartId);
  if (found) {
    editingChartId = found.id;
    draftChart = JSON.parse(JSON.stringify(found));
  } else {
    const stdMap: Record<string, StudioChartCfg> = getStandardChartDefs();
    if (stdMap[chartId]) {
      editingChartId = chartId;
      draftChart = JSON.parse(JSON.stringify(stdMap[chartId]));
    }
  }
  renderGraficosTabContent();
}

export function getStandardChartDefs(): Record<string, StudioChartCfg> {
  const palette = getThemeColors();
  return {
    std_fu: {
      id: 'std_fu',
      title: 'Mensagens de follow-up enviadas por tentativa (FU 2 a FU 5)',
      type: 'stack',
      dataSource: 'semanas',
      xAxisField: 'id',
      xAxisGrouping: 'individual',
      xAxisSort: 'asc',
      periodFilter: 'all',
      originFilter: '',
      statusFilter: '',
      series: [
        { id: 'fu2', name: 'FU 2', source: 'semanas', field: 'fu2', agg: 'SUM', color: palette[1] },
        { id: 'fu3', name: 'FU 3', source: 'semanas', field: 'fu3', agg: 'SUM', color: palette[0] },
        { id: 'fu4', name: 'FU 4', source: 'semanas', field: 'fu4', agg: 'SUM', color: palette[2] },
        { id: 'fu5', name: 'FU 5', source: 'semanas', field: 'fu5', agg: 'SUM', color: palette[4] }
      ],
      showInDashboard: true,
      isStandardOverride: true,
      standardKey: 'std_fu'
    },
    std_rs: {
      id: 'std_rs',
      title: 'Taxa de resposta e conversão do follow-up (%)',
      type: 'line',
      dataSource: 'semanas',
      xAxisField: 'id',
      xAxisGrouping: 'individual',
      xAxisSort: 'asc',
      periodFilter: 'all',
      originFilter: '',
      statusFilter: '',
      series: [
        { id: 'resp', name: 'Taxa de resposta', source: 'semanas', field: 'taxa_resp', agg: 'PCT', color: palette[1] },
        { id: 'conv', name: 'Respostas que viraram agendamento', source: 'semanas', field: 'taxa_conv', agg: 'PCT', color: palette[0] }
      ],
      showInDashboard: true,
      isStandardOverride: true,
      standardKey: 'std_rs'
    },
    std_st: {
      id: 'std_st',
      title: 'Situação dos pacientes ao fim de cada semana',
      type: 'stack',
      dataSource: 'semanas',
      xAxisField: 'id',
      xAxisGrouping: 'individual',
      xAxisSort: 'asc',
      periodFilter: 'all',
      originFilter: '',
      statusFilter: '',
      series: [
        { id: 'sr', name: 'Sem resposta', source: 'semanas', field: 'semresp', agg: 'SUM', color: palette[5] },
        { id: 'pe', name: 'Pensando ou pesquisando', source: 'semanas', field: 'pensando', agg: 'SUM', color: palette[2] },
        { id: 'ag', name: 'Aguardando cadastro', source: 'semanas', field: 'aguard', agg: 'SUM', color: palette[0] },
        { id: 'na', name: 'Não vão seguir', source: 'semanas', field: 'nao', agg: 'SUM', color: palette[3] }
      ],
      showInDashboard: true,
      isStandardOverride: true,
      standardKey: 'std_st'
    },
    std_ag: {
      id: 'std_ag',
      title: 'Agenda da semana: previstas, realizadas, canceladas e reagendadas',
      type: 'group',
      dataSource: 'semanas',
      xAxisField: 'id',
      xAxisGrouping: 'individual',
      xAxisSort: 'asc',
      periodFilter: 'all',
      originFilter: '',
      statusFilter: '',
      series: [
        { id: 'cp', name: 'Previstas', source: 'semanas', field: 'cprev', agg: 'SUM', color: palette[1] },
        { id: 'cr', name: 'Realizadas', source: 'semanas', field: 'creal', agg: 'SUM', color: palette[0] },
        { id: 'cc', name: 'Canceladas', source: 'semanas', field: 'ccanc', agg: 'SUM', color: palette[3] },
        { id: 'ce', name: 'Reagendadas', source: 'semanas', field: 'creag', agg: 'SUM', color: palette[2] }
      ],
      showInDashboard: true,
      isStandardOverride: true,
      standardKey: 'std_ag'
    },
    std_fl: {
      id: 'std_fl',
      title: 'Funil: leads, respostas e agendamentos',
      type: 'group',
      dataSource: 'semanas',
      xAxisField: 'id',
      xAxisGrouping: 'individual',
      xAxisSort: 'asc',
      periodFilter: 'all',
      originFilter: '',
      statusFilter: '',
      series: [
        { id: 'nv', name: 'Novos leads', source: 'semanas', field: 'novos', agg: 'SUM', color: palette[0] },
        { id: 'rp', name: 'Respostas no follow-up', source: 'semanas', field: 'resp', agg: 'SUM', color: palette[1] },
        { id: 'ag', name: 'Agendamentos pelo follow-up', source: 'semanas', field: 'agend', agg: 'SUM', color: palette[2] }
      ],
      showInDashboard: true,
      isStandardOverride: true,
      standardKey: 'std_fl'
    },
    std_goals: {
      id: 'std_goals',
      title: 'Acompanhamento de Metas por Semana',
      type: 'line',
      dataSource: 'metas',
      xAxisField: 'id',
      xAxisGrouping: 'individual',
      xAxisSort: 'asc',
      periodFilter: 'all',
      originFilter: '',
      statusFilter: '',
      series: [
        { id: 'g_act', name: 'Consultas agendadas na semana', source: 'semanas', field: 'cagend', agg: 'SUM', color: palette[0] }
      ],
      showInDashboard: true,
      isStandardOverride: true,
      standardKey: 'std_goals'
    }
  };
}

export function renderGraficosTabContent() {
  const container = document.getElementById('graf');
  if (!container) return;

  const app = getApp();
  const cfg = app.cfg || {};
  const esc = app.esc || ((s: any) => String(s ?? ''));
  const ORIG = app.ORIG || {};
  const STAT = app.STAT || {};

  if (!draftChart) draftChart = createDefaultDraft();
  const c = draftChart;
  if (!Array.isArray(c.xAxisSeries)) c.xAxisSeries = [];

  const xOptions = getXAxisOptions(c.dataSource, c.spreadsheetId);
  if (!xOptions.some(o => o.value === c.xAxisField)) {
    c.xAxisField = xOptions[0]?.value || 'id';
  }
  const yOptions = getYFieldOptions(c.dataSource, c.spreadsheetId);
  const sheets: any[] = cfg.spreadsheets || [];
  const customCharts: StudioChartCfg[] = cfg.customCharts || [];
  const stdDefs = getStandardChartDefs();

  const previewHtml = renderStudioSvg(c).html;
  const metricHtmlStr = app.metricHtml ? app.metricHtml() : '';

  const allGalleryCharts: { def: StudioChartCfg; isStd: boolean; isOverride: boolean; isVisible: boolean }[] = [
    ...Object.values(stdDefs).map(std => {
      const override = customCharts.find(x => x.id === std.id);
      const activeDef = override || std;
      return {
        def: activeDef,
        isStd: true,
        isOverride: !!override,
        isVisible: activeDef.showInDashboard !== false && !activeDef.hidden
      };
    }),
    ...customCharts.filter(x => !x.isStandardOverride && !String(x.id).startsWith('std_')).map(cc => ({
      def: cc,
      isStd: false,
      isOverride: false,
      isVisible: cc.showInDashboard !== false && !cc.hidden
    }))
  ];

  container.innerHTML = `
    <div class="card" id="gb_studio_card">
      <div style="display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:10px;padding-bottom:12px;border-bottom:1px solid var(--line)">
        <div>
          <h2 style="margin:0">${editingChartId ? '✏️ Editando Gráfico: ' + esc(c.title) : '🪄 Estúdio de Construção Livre de Gráficos'}</h2>
          <p class="note" style="margin:3px 0 0">Configure livremente o Eixo X, o Eixo Y e o Modelo Visual para analisar qualquer cruzamento de dados em tempo real.</p>
        </div>
        <div style="display:flex;gap:8px;flex-wrap:wrap">
          <button type="button" class="b s" id="gb_btn_new">➕ Novo Gráfico</button>
          <button type="button" class="b" id="gb_btn_save">📌 Fixar Este Gráfico no Painel</button>
        </div>
      </div>

      <div class="f" style="margin-top:14px;padding:14px;background:var(--bg);border:1px solid var(--line);border-radius:12px">
        <label style="grid-column:1/-1">Título do Gráfico
          <input type="text" id="gb_title" value="${esc(c.title)}">
        </label>

        <label>1. Domínio de Origem do Dado
          <select id="gb_source">
            <option value="semanas" ${c.dataSource === 'semanas' ? 'selected' : ''}>Semanas Lançadas (Follow-up e Funil)</option>
            <option value="consultas" ${c.dataSource === 'consultas' ? 'selected' : ''}>Consultas &amp; Leads (Individual)</option>
            <option value="cruzado" ${c.dataSource === 'cruzado' ? 'selected' : ''}>Cruzamento: Semanas × Consultas</option>
            <option value="metas" ${c.dataSource === 'metas' ? 'selected' : ''}>Metas e Objetivos</option>
            <option value="planilha" ${c.dataSource === 'planilha' ? 'selected' : ''}>Planilha Importada (.xlsx / .csv)</option>
          </select>
        </label>

        ${c.dataSource === 'planilha' ? `
          <label>Planilha Selecionada
            <select id="gb_sheet_id">
              ${sheets.length ? sheets.map((s: any) => `<option value="${esc(s.id)}" ${c.spreadsheetId === s.id ? 'selected' : ''}>${esc(s.name)} (${s.rows?.length || 0} linhas)</option>`).join('') : '<option value="">Nenhuma planilha importada</option>'}
            </select>
          </label>
        ` : ''}

        <label>2. Dimensão / Eixo X
          <select id="gb_x_axis">
            ${xOptions.map(o => `<option value="${esc(o.value)}" ${c.xAxisField === o.value ? 'selected' : ''}>${esc(o.label)}</option>`).join('')}
          </select>
        </label>

        <label>3. Modelo Visual do Gráfico
          <select id="gb_type">
            <option value="group" ${c.type === 'group' ? 'selected' : ''}>Colunas Verticais Agrupadas</option>
            <option value="stack" ${c.type === 'stack' ? 'selected' : ''}>Colunas Empilhadas</option>
            <option value="hbar" ${c.type === 'hbar' ? 'selected' : ''}>Barras Horizontais</option>
            <option value="line" ${c.type === 'line' ? 'selected' : ''}>Gráfico de Linhas</option>
            <option value="area" ${c.type === 'area' ? 'selected' : ''}>Gráfico de Área</option>
            <option value="pie" ${c.type === 'pie' ? 'selected' : ''}>Gráfico de Pizza</option>
            <option value="donut" ${c.type === 'donut' ? 'selected' : ''}>Gráfico de Rosca</option>
            <option value="radar" ${c.type === 'radar' ? 'selected' : ''}>Gráfico de Radar</option>
            <option value="scatter" ${c.type === 'scatter' ? 'selected' : ''}>Gráfico de Dispersão</option>
            <option value="combo" ${c.type === 'combo' ? 'selected' : ''}>Gráfico Combinado (Colunas + Linha)</option>
            <option value="kpi" ${c.type === 'kpi' ? 'selected' : ''}>Card KPI de Métrica</option>
          </select>
        </label>

        <label>Agrupamento do Eixo X
          <select id="gb_x_group">
            <option value="individual" ${c.xAxisGrouping === 'individual' ? 'selected' : ''}>Individual (por registro/semana)</option>
            <option value="week" ${c.xAxisGrouping === 'week' ? 'selected' : ''}>Agrupar por Semana (Sáb–Sex)</option>
            <option value="month" ${c.xAxisGrouping === 'month' ? 'selected' : ''}>Agrupar por Mês</option>
            <option value="quarter" ${c.xAxisGrouping === 'quarter' ? 'selected' : ''}>Agrupar por Trimestre</option>
            <option value="year" ${c.xAxisGrouping === 'year' ? 'selected' : ''}>Agrupar por Ano</option>
            <option value="weekday" ${c.xAxisGrouping === 'weekday' ? 'selected' : ''}>Agrupar por Dia da Semana</option>
          </select>
        </label>

        <label>Ordenação do Eixo X
          <select id="gb_x_sort">
            <option value="asc" ${c.xAxisSort === 'asc' ? 'selected' : ''}>Crescente (Antiga → Recente / A → Z)</option>
            <option value="desc" ${c.xAxisSort === 'desc' ? 'selected' : ''}>Decrescente (Recente → Antiga / Z → A)</option>
          </select>
        </label>

        <label>Filtro de Período
          <select id="gb_period">
            <option value="all" ${c.periodFilter === 'all' ? 'selected' : ''}>Todas as semanas</option>
            <option value="last4" ${c.periodFilter === 'last4' ? 'selected' : ''}>Últimas 4 semanas</option>
            <option value="last8" ${c.periodFilter === 'last8' ? 'selected' : ''}>Últimas 8 semanas</option>
            <option value="last12" ${c.periodFilter === 'last12' ? 'selected' : ''}>Últimas 12 semanas</option>
            <option value="last" ${c.periodFilter === 'last' ? 'selected' : ''}>Última semana lançada</option>
            <option value="penultimate" ${c.periodFilter === 'penultimate' ? 'selected' : ''}>Penúltima semana lançada</option>
            ${(() => {
              const WA: any[] = app.W ? app.W() : [];
              const MONTH_NAMES: string[] = app.MONTH_NAMES || ['janeiro','fevereiro','março','abril','maio','junho','julho','agosto','setembro','outubro','novembro','dezembro'];
              const ymSet = new Set<string>();
              WA.forEach(w => { if (w.id && w.id.length >= 7) ymSet.add(w.id.slice(0, 7)); });
              return Array.from(ymSet).sort().reverse().map(ym => {
                const [y, m] = ym.split('-');
                const mName = MONTH_NAMES[parseInt(m, 10) - 1] || m;
                const lbl = mName.charAt(0).toUpperCase() + mName.slice(1) + '/' + y;
                return `<option value="ym:${ym}" ${c.periodFilter === 'ym:' + ym ? 'selected' : ''}>Mês: ${lbl}</option>`;
              }).join('');
            })()}
          </select>
        </label>

        ${(c.dataSource === 'consultas' || c.dataSource === 'cruzado') ? `
          <label>Filtrar Origem
            <select id="gb_origin">
              <option value="">Todas as origens</option>
              ${Object.entries(ORIG).map(([k, v]: any) => `<option value="${esc(k)}" ${c.originFilter === k ? 'selected' : ''}>${esc(v)}</option>`).join('')}
            </select>
          </label>
          <label>Filtrar Status
            <select id="gb_status">
              <option value="">Todos os status</option>
              ${Object.entries(STAT).map(([k, v]: any) => `<option value="${esc(k)}" ${c.statusFilter === k ? 'selected' : ''}>${esc(v)}</option>`).join('')}
            </select>
          </label>
        ` : ''}
      </div>

      <!-- Eixo X: Séries / Dimensões adicionais -->
      <div style="margin-top:14px">
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px;flex-wrap:wrap;gap:6px">
          <div>
            <h3 style="margin:0;font-size:14px">Eixo X (Séries e Cruzamento de Dimensões)</h3>
            <span class="note" style="font-size:12px">Adicione dimensões extras no Eixo X para cruzar dados (ex: Semana × Origem ou Origem × Status).</span>
          </div>
          <button type="button" class="b s" id="gb_btn_add_x_series" style="font-size:11.5px;padding:4px 9px">➕ Adicionar Série no Eixo X</button>
        </div>
        ${c.xAxisSeries.length ? `
          <div style="display:flex;flex-direction:column;gap:6px">
            ${c.xAxisSeries.map((xf, idx) => `
              <div class="row" style="margin:0;padding:7px 12px;background:var(--bg);border:1px solid var(--line);border-radius:9px;align-items:center">
                <span style="font-weight:700;font-size:12px;color:var(--mut)">Dimensão X #${idx + 2}</span>
                <select data-gx-idx="${idx}" style="flex:1;min-width:180px">
                  ${xOptions.map(o => `<option value="${esc(o.value)}" ${xf === o.value ? 'selected' : ''}>${esc(o.label)}</option>`).join('')}
                </select>
                <button type="button" class="x" data-gx-del="${idx}" title="Remover dimensão X" style="font-size:18px">×</button>
              </div>
            `).join('')}
          </div>
        ` : ''}
      </div>

      <!-- Eixo Y: Séries -->
      <div style="margin-top:14px">
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px;flex-wrap:wrap;gap:6px">
          <h3 style="margin:0;font-size:14px">4. Métrica / Eixo Y (Séries Independentes)</h3>
          <div style="display:flex;gap:6px">
            <button type="button" class="b s" id="gb_btn_palette" style="font-size:11.5px;padding:4px 9px" title="Aplicar cores padrão do painel">🎨 Aplicar Cores do Site</button>
            <button type="button" class="b s" id="gb_btn_add_series" style="font-size:11.5px;padding:4px 9px">➕ Adicionar Série no Eixo Y</button>
          </div>
        </div>
        <div style="display:flex;flex-direction:column;gap:8px">
          ${c.series.map((s, idx) => `
            <div class="row" style="margin:0;padding:8px 12px;background:var(--bg);border:1px solid var(--line);border-radius:9px;align-items:center">
              <span style="font-weight:700;font-size:12px;color:var(--mut)">#${idx + 1}</span>
              <label style="font-size:11.5px;color:var(--mut);flex:1;min-width:130px">Nome da Série
                <input type="text" data-gs-idx="${idx}" data-gs-k="name" value="${esc(s.name)}">
              </label>
              <label style="font-size:11.5px;color:var(--mut);flex:1.5;min-width:180px">Campo / Métrica (Eixo Y)
                <select data-gs-idx="${idx}" data-gs-k="field">
                  ${yOptions.map(o => `<option value="${esc(o.value)}" data-src="${o.source}" ${s.field === o.value ? 'selected' : ''}>${esc(o.group)}: ${esc(o.label)}</option>`).join('')}
                </select>
              </label>
              <label style="font-size:11.5px;color:var(--mut);min-width:130px">Operação
                <select data-gs-idx="${idx}" data-gs-k="agg">
                  <option value="SUM" ${s.agg === 'SUM' ? 'selected' : ''}>Soma (SUM)</option>
                  <option value="AVG" ${s.agg === 'AVG' ? 'selected' : ''}>Média (AVG)</option>
                  <option value="MEDIAN" ${s.agg === 'MEDIAN' ? 'selected' : ''}>Mediana (MEDIAN)</option>
                  <option value="COUNT" ${s.agg === 'COUNT' ? 'selected' : ''}>Contagem</option>
                  <option value="COUNT_DISTINCT" ${s.agg === 'COUNT_DISTINCT' ? 'selected' : ''}>Contagem Distinta</option>
                  <option value="MAX" ${s.agg === 'MAX' ? 'selected' : ''}>Máximo</option>
                  <option value="MIN" ${s.agg === 'MIN' ? 'selected' : ''}>Mínimo</option>
                  <option value="PCT" ${s.agg === 'PCT' ? 'selected' : ''}>Taxa / Porcentagem (%)</option>
                </select>
              </label>
              <label style="font-size:11.5px;color:var(--mut)">Cor
                <input type="color" data-gs-idx="${idx}" data-gs-k="color" value="${s.color}" style="width:42px;height:34px;padding:2px;cursor:pointer">
              </label>
              ${c.series.length > 1 ? `<button type="button" class="x" data-gs-del="${idx}" title="Remover série" style="font-size:18px">×</button>` : ''}
            </div>
          `).join('')}
        </div>
      </div>

      <!-- Pré-visualização ao vivo -->
      <div style="margin-top:16px;padding:16px;background:var(--bg);border:1px solid var(--line);border-radius:12px">
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px;flex-wrap:wrap;gap:6px">
          <div>
            <h3 id="gb_preview_title" style="margin:0;font-size:15px">${esc(c.title)}</h3>
            <span class="note" style="font-size:11.5px">Passe o mouse sobre os valores para ver detalhes</span>
          </div>
          <span class="tag" style="background:color-mix(in srgb,var(--b) 12%,var(--card));color:var(--b);font-size:11px;font-weight:700">LIVE RENDER</span>
        </div>
        <div id="gb_preview_area">${previewHtml}</div>
      </div>

      <div class="row" style="margin-top:14px;justify-content:space-between;align-items:center">
        <label style="display:inline-flex;align-items:center;gap:6px;font-size:13px;font-weight:600;cursor:pointer">
          <input type="checkbox" id="gb_show_dash" ${c.showInDashboard ? 'checked' : ''} style="width:auto;margin:0">
          <span>Exibir este gráfico no Painel Principal</span>
        </label>
        <div style="display:flex;gap:8px">
          <button type="button" class="b s" id="gb_btn_cancel">${editingChartId ? 'Cancelar Edição' : 'Limpar'}</button>
          <button type="button" class="b" id="gb_btn_save_bottom">💾 Salvar Gráfico no Painel</button>
        </div>
      </div>
    </div>

    <!-- Galeria de Gráficos do Dashboard Executivo (Editar Gráficos do Painel) -->
    <div class="card" style="margin-top:14px">
      <div style="display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:8px;margin-bottom:12px">
        <div>
          <h2 style="margin:0">🎛️ Galeria de Gráficos do Painel Executivo</h2>
          <p class="note" style="margin:2px 0 0">Múltiplos painéis analíticos com dados de consultas, desempenho semanal e metas. Edite, duplique, oculte ou restaure sem afetar os dados.</p>
        </div>
        <div style="display:flex;gap:6px;flex-wrap:wrap">
          <button type="button" class="b s" id="gb_restore_all_std" style="font-size:12px;padding:6px 11px">↺ Restaurar Padrões</button>
          <button type="button" class="b" id="gb_gallery_add" style="font-size:12px;padding:6px 12px">+ Adicionar Painel</button>
        </div>
      </div>

      <div class="grid charts">
        ${allGalleryCharts.map(item => {
          const d = item.def;
          const isConfirm = confirmDeleteId === d.id;
          const miniSvg = renderStudioSvg(d).html;
          return `
            <div style="background:var(--bg);border:1px solid var(--line);border-radius:12px;padding:14px;display:flex;flex-direction:column;gap:10px;min-width:0">
              <div style="display:flex;justify-content:space-between;align-items:flex-start;gap:8px;padding-bottom:8px;border-bottom:1px solid var(--line)">
                <div style="min-width:0">
                  <div style="font-weight:700;font-size:13px;text-transform:uppercase;letter-spacing:.02em;color:var(--ink)">${esc(d.title)}</div>
                  <div style="display:flex;gap:6px;align-items:center;margin-top:4px;flex-wrap:wrap">
                    <span class="tag" style="font-size:10.5px">${d.type}</span>
                    <span class="note" style="font-size:11px">${item.isVisible ? '✓ Visível no Painel' : 'Oculto no Painel'}</span>
                    ${item.isStd ? '<span class="tag" style="font-size:10.5px">Padrão</span>' : '<span class="tag" style="font-size:10.5px;background:color-mix(in srgb,var(--g) 15%,var(--card));color:var(--g)">Personalizado</span>'}
                  </div>
                </div>
                <div style="display:flex;gap:4px;flex-shrink:0;flex-wrap:wrap">
                  <button type="button" class="b s" data-cm-toggle="${d.id}" style="font-size:11px;padding:3px 7px" title="${item.isVisible ? 'Ocultar do Painel' : 'Exibir no Painel'}">${item.isVisible ? 'Ocultar' : 'Exibir'}</button>
                  <button type="button" class="b s" data-cm-dup="${d.id}" style="font-size:11px;padding:3px 7px" title="Duplicar gráfico">Duplicar</button>
                  <button type="button" class="b s" data-cm-edit="${d.id}" style="font-size:11px;padding:3px 7px" title="Editar gráfico">✏️</button>
                  ${item.isStd
                    ? (item.isOverride
                        ? (isConfirm
                            ? `<button type="button" class="b" data-cm-del-confirm="${d.id}" style="font-size:11px;padding:3px 7px;background:var(--r);border-color:var(--r)">Restaurar?</button>`
                            : `<button type="button" class="b s" data-cm-del="${d.id}" style="font-size:11px;padding:3px 7px;color:var(--r)" title="Restaurar padrão">↺</button>`)
                        : '')
                    : (isConfirm
                        ? `<button type="button" class="b" data-cm-del-confirm="${d.id}" style="font-size:11px;padding:3px 7px;background:var(--r);border-color:var(--r)">Excluir?</button>`
                        : `<button type="button" class="b s" data-cm-del="${d.id}" style="font-size:11px;padding:3px 7px;color:var(--r)" title="Excluir gráfico">🗑️</button>`)}
                </div>
              </div>
              <div style="min-width:0">${miniSvg}</div>
            </div>
          `;
        }).join('')}
      </div>
    </div>

    ${metricHtmlStr}

    <div class="card" style="margin-top:14px">
      <h2>Importação de Planilhas Externas (.xlsx e .csv)</h2>
      <p class="note">Importe planilhas para usar como fonte de dados no Construtor de Gráficos acima.</p>
      <div class="row" style="margin-top:8px">
        <input type="file" id="gb_sp_file" accept=".xlsx,.xls,.csv" style="display:none">
        <button type="button" class="b" id="gb_sp_upload">📁 Carregar Planilha (.xlsx / .csv)</button>
        <button type="button" class="b s" id="gb_sp_template">📄 Baixar Modelo Padrão (.xlsx)</button>
      </div>
      ${sheets.length ? `
        <div style="margin-top:12px;display:flex;flex-direction:column;gap:6px">
          ${sheets.map((sp: any) => `
            <div style="display:flex;justify-content:space-between;align-items:center;padding:8px 12px;background:var(--bg);border:1px solid var(--line);border-radius:8px">
              <div>
                <b>${esc(sp.name)}</b>
                <span class="note" style="margin-left:8px;font-size:12px">${sp.rows?.length || 0} linhas · Colunas: ${esc((sp.headers || []).join(', '))}</span>
              </div>
              <button type="button" class="x" data-sp-del="${esc(sp.id)}" title="Remover planilha" style="font-size:18px">×</button>
            </div>
          `).join('')}
        </div>
      ` : ''}
    </div>
  `;

  attachGrafEvents();
  if (typeof app.attachMetricBuilderEvents === 'function') {
    app.attachMetricBuilderEvents();
  }
}

export const renderGrafTab = renderGraficosTabContent;

function attachGrafEvents() {
  const container = document.getElementById('graf');
  if (!container) return;
  const app = getApp();
  const c = draftChart!;

  const refreshPreviewOnly = () => {
    const area = document.getElementById('gb_preview_area');
    if (area) area.innerHTML = renderStudioSvg(c).html;
  };

  container.querySelector('#gb_title')?.addEventListener('input', (e: any) => {
    c.title = e.target.value;
    const h3 = container.querySelector('#gb_preview_title');
    if (h3) h3.textContent = c.title;
  });
  container.querySelector('#gb_type')?.addEventListener('change', (e: any) => { c.type = e.target.value; refreshPreviewOnly(); });
  container.querySelector('#gb_source')?.addEventListener('change', (e: any) => {
    c.dataSource = e.target.value;
    const xOpts = getXAxisOptions(c.dataSource, c.spreadsheetId);
    c.xAxisField = xOpts[0]?.value || 'id';
    c.xAxisSeries = [];
    const yOpts = getYFieldOptions(c.dataSource, c.spreadsheetId);
    if (yOpts.length) {
      c.series = [{ id: 's1', name: yOpts[0].label, source: yOpts[0].source, field: yOpts[0].value, agg: 'SUM', color: getThemeColors()[0] }];
    }
    renderGraficosTabContent();
  });
  container.querySelector('#gb_sheet_id')?.addEventListener('change', (e: any) => {
    c.spreadsheetId = e.target.value;
    renderGraficosTabContent();
  });
  container.querySelector('#gb_x_axis')?.addEventListener('change', (e: any) => { c.xAxisField = e.target.value; refreshPreviewOnly(); });
  container.querySelector('#gb_x_group')?.addEventListener('change', (e: any) => { c.xAxisGrouping = e.target.value; refreshPreviewOnly(); });
  container.querySelector('#gb_x_sort')?.addEventListener('change', (e: any) => { c.xAxisSort = e.target.value; refreshPreviewOnly(); });
  container.querySelector('#gb_period')?.addEventListener('change', (e: any) => { c.periodFilter = e.target.value; refreshPreviewOnly(); });
  container.querySelector('#gb_origin')?.addEventListener('change', (e: any) => { c.originFilter = e.target.value; refreshPreviewOnly(); });
  container.querySelector('#gb_status')?.addEventListener('change', (e: any) => { c.statusFilter = e.target.value; refreshPreviewOnly(); });
  container.querySelector('#gb_show_dash')?.addEventListener('change', (e: any) => { c.showInDashboard = e.target.checked; });

  // X-axis extra series
  container.querySelector('#gb_btn_add_x_series')?.addEventListener('click', () => {
    const xOpts = getXAxisOptions(c.dataSource, c.spreadsheetId);
    const nextOpt = xOpts.find(o => o.value !== c.xAxisField && !(c.xAxisSeries || []).includes(o.value)) || xOpts[0];
    if (!nextOpt) return;
    if (!Array.isArray(c.xAxisSeries)) c.xAxisSeries = [];
    c.xAxisSeries.push(nextOpt.value);
    renderGraficosTabContent();
  });

  container.querySelectorAll('[data-gx-idx]').forEach((el: any) => {
    el.addEventListener('change', () => {
      const idx = parseInt(el.dataset.gxIdx, 10);
      if (Array.isArray(c.xAxisSeries)) {
        c.xAxisSeries[idx] = el.value;
        refreshPreviewOnly();
      }
    });
  });

  container.querySelectorAll('[data-gx-del]').forEach((btn: any) => {
    btn.onclick = () => {
      const idx = parseInt(btn.dataset.gxDel, 10);
      if (Array.isArray(c.xAxisSeries)) {
        c.xAxisSeries.splice(idx, 1);
        renderGraficosTabContent();
      }
    };
  });

  // Y-axis series inputs
  container.querySelectorAll('[data-gs-idx]').forEach((el: any) => {
    const handler = () => {
      const idx = parseInt(el.dataset.gsIdx, 10);
      const k = el.dataset.gsK;
      if (c.series[idx] && k) {
        (c.series[idx] as any)[k] = el.value;
        if (k === 'field' && el.selectedOptions?.[0]?.dataset?.src) {
          c.series[idx].source = el.selectedOptions[0].dataset.src;
          const lbl = el.selectedOptions[0].textContent?.split(': ').slice(1).join(': ');
          if (lbl) {
            c.series[idx].name = lbl;
            const nameInp = container.querySelector(`[data-gs-idx="${idx}"][data-gs-k="name"]`) as HTMLInputElement;
            if (nameInp) nameInp.value = lbl;
          }
        }
        refreshPreviewOnly();
      }
    };
    el.addEventListener('input', handler);
    el.addEventListener('change', handler);
  });

  container.querySelectorAll('[data-gs-del]').forEach((btn: any) => {
    btn.onclick = () => {
      const idx = parseInt(btn.dataset.gsDel, 10);
      c.series.splice(idx, 1);
      renderGraficosTabContent();
    };
  });

  container.querySelector('#gb_btn_add_series')?.addEventListener('click', () => {
    const yOpts = getYFieldOptions(c.dataSource, c.spreadsheetId);
    const pick = yOpts[c.series.length % yOpts.length] || yOpts[0];
    const palette = getThemeColors();
    c.series.push({
      id: 's' + Date.now().toString(36),
      name: pick ? pick.label : 'Nova Série',
      source: pick ? pick.source : 'semanas',
      field: pick ? pick.value : 'novos',
      agg: 'SUM',
      color: palette[c.series.length % palette.length]
    });
    renderGraficosTabContent();
  });

  container.querySelector('#gb_btn_palette')?.addEventListener('click', () => {
    const palette = getThemeColors();
    c.series.forEach((s, i) => { s.color = palette[i % palette.length]; });
    renderGraficosTabContent();
    app.toast?.('Cores padrão do site aplicadas!');
  });

  container.querySelector('#gb_btn_new')?.addEventListener('click', () => {
    editingChartId = null;
    confirmDeleteId = null;
    draftChart = createDefaultDraft();
    renderGraficosTabContent();
  });

  container.querySelector('#gb_btn_cancel')?.addEventListener('click', () => {
    editingChartId = null;
    confirmDeleteId = null;
    draftChart = createDefaultDraft();
    renderGraficosTabContent();
    app.toast?.('Edição cancelada.');
  });

  const saveDraftHandler = async () => {
    if (!c.title.trim()) return app.toast?.('Informe o título do gráfico.');
    const cfg = app.cfg || {};
    const list: StudioChartCfg[] = [...(cfg.customCharts || [])];
    const finalChart: StudioChartCfg = JSON.parse(JSON.stringify(c));
    finalChart.id = editingChartId || ('gc_' + Date.now().toString(36));

    const idx = list.findIndex(x => x.id === finalChart.id);
    if (idx >= 0) list[idx] = finalChart;
    else list.push(finalChart);

    editingChartId = null;
    confirmDeleteId = null;
    draftChart = createDefaultDraft();
    await app.saveCfg({ customCharts: list });
    renderGraficosTabContent();
    app.toast?.('Gráfico salvo no painel com sucesso!');
  };

  container.querySelector('#gb_btn_save')?.addEventListener('click', saveDraftHandler);
  container.querySelector('#gb_btn_save_bottom')?.addEventListener('click', saveDraftHandler);
  container.querySelector('#gb_gallery_add')?.addEventListener('click', () => {
    editingChartId = null;
    confirmDeleteId = null;
    draftChart = createDefaultDraft();
    renderGraficosTabContent();
    window.scrollTo({ top: 0, behavior: 'smooth' });
  });

  container.querySelector('#gb_restore_all_std')?.addEventListener('click', async () => {
    const list: StudioChartCfg[] = (app.cfg?.customCharts || []).filter((x: StudioChartCfg) => !String(x.id).startsWith('std_') && !x.isStandardOverride);
    editingChartId = null;
    confirmDeleteId = null;
    await app.saveCfg({ customCharts: list });
    renderGraficosTabContent();
    app.toast?.('Todos os gráficos padrão foram restaurados (dados preservados).');
  });

  container.querySelectorAll('[data-cm-edit]').forEach((btn: any) => {
    btn.onclick = () => {
      confirmDeleteId = null;
      startEditingChart(btn.dataset.cmEdit);
      window.scrollTo({ top: 0, behavior: 'smooth' });
    };
  });

  container.querySelectorAll('[data-cm-toggle]').forEach((btn: any) => {
    btn.onclick = async () => {
      const id = btn.dataset.cmToggle;
      const list: StudioChartCfg[] = [...(app.cfg?.customCharts || [])];
      const idx = list.findIndex(x => x.id === id);
      if (idx >= 0) {
        list[idx].showInDashboard = !list[idx].showInDashboard;
        list[idx].hidden = !list[idx].showInDashboard;
      } else {
        const std = getStandardChartDefs()[id];
        if (std) {
          std.showInDashboard = false;
          std.hidden = true;
          list.push(std);
        }
      }
      await app.saveCfg({ customCharts: list });
      renderGraficosTabContent();
    };
  });

  container.querySelectorAll('[data-cm-dup]').forEach((btn: any) => {
    btn.onclick = async () => {
      const id = btn.dataset.cmDup;
      const list: StudioChartCfg[] = [...(app.cfg?.customCharts || [])];
      const item = list.find(x => x.id === id) || getStandardChartDefs()[id];
      if (item) {
        const copy: StudioChartCfg = JSON.parse(JSON.stringify(item));
        copy.id = 'gc_' + Date.now().toString(36);
        copy.title = copy.title + ' (Cópia)';
        copy.isStandardOverride = false;
        copy.showInDashboard = true;
        delete copy.hidden;
        delete copy.standardKey;
        list.push(copy);
        await app.saveCfg({ customCharts: list });
        renderGraficosTabContent();
        app.toast?.('Gráfico duplicado!');
      }
    };
  });

  container.querySelectorAll('[data-cm-del]').forEach((btn: any) => {
    btn.onclick = () => {
      confirmDeleteId = btn.dataset.cmDel;
      renderGraficosTabContent();
    };
  });

  container.querySelectorAll('[data-cm-del-confirm]').forEach((btn: any) => {
    btn.onclick = async () => {
      const id = btn.dataset.cmDelConfirm;
      const list: StudioChartCfg[] = (app.cfg?.customCharts || []).filter((x: StudioChartCfg) => x.id !== id);
      if (editingChartId === id) {
        editingChartId = null;
        draftChart = createDefaultDraft();
      }
      confirmDeleteId = null;
      await app.saveCfg({ customCharts: list });
      renderGraficosTabContent();
      app.toast?.('Gráfico removido.');
    };
  });

  const spInp = container.querySelector('#gb_sp_file') as HTMLInputElement;
  container.querySelector('#gb_sp_upload')?.addEventListener('click', () => spInp?.click());
  spInp?.addEventListener('change', async (e: any) => {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: 'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows: any[] = XLSX.utils.sheet_to_json(ws);
      if (!rows.length) return app.toast?.('Planilha vazia ou sem cabeçalhos.');
      const headers = Object.keys(rows[0]);
      const newSp = {
        id: 'sp_' + Date.now().toString(36),
        name: file.name,
        uploadedAt: new Date().toISOString(),
        headers,
        rows
      };
      const sheets = [...(app.cfg?.spreadsheets || []), newSp];
      await app.saveCfg({ spreadsheets: sheets });
      if (draftChart) {
        draftChart.dataSource = 'planilha';
        draftChart.spreadsheetId = newSp.id;
      }
      renderGraficosTabContent();
      app.toast?.(`Planilha "${file.name}" importada (${rows.length} linhas)!`);
    } catch (err: any) {
      app.toast?.('Erro ao ler planilha: ' + (err?.message || 'formato inválido'));
    }
  });

  container.querySelector('#gb_sp_template')?.addEventListener('click', () => {
    const sample = [
      ['Data', 'Origem', 'Indicador', 'Quantidade'],
      ['2026-10-01', 'Instagram', 'Leads', 35],
      ['2026-10-01', 'Google', 'Leads', 20],
      ['2026-10-02', 'Instagram', 'Leads', 42]
    ];
    const ws = XLSX.utils.aoa_to_sheet(sample);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Modelo');
    XLSX.writeFile(wb, 'modelo_planilha_kommo.xlsx');
  });

  container.querySelectorAll('[data-sp-del]').forEach((btn: any) => {
    btn.onclick = async () => {
      const id = btn.dataset.spDel;
      const sheets = (app.cfg?.spreadsheets || []).filter((s: any) => s.id !== id);
      await app.saveCfg({ spreadsheets: sheets });
      renderGraficosTabContent();
      app.toast?.('Planilha removida.');
    };
  });

  container.onclick = async (e: any) => {
    const t = e.target;
    if (t.dataset?.md != null) {
      const idx = +t.dataset.md;
      await app.saveCfg({ metrics: (app.cfg?.metrics || []).filter((_: any, i: number) => i !== idx) });
      renderGraficosTabContent();
      app.toast?.('Métrica removida.');
    }
  };
}
