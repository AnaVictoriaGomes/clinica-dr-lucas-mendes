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
  dashed?: boolean;
}

export type StudioFilterOp = 'eq' | 'neq' | 'contains' | 'gte' | 'lte' | 'between';

export interface StudioFilterRule {
  id: string;
  field: string;
  op: StudioFilterOp;
  value: string;
  value2?: string;
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
  periodFilter: 'all' | 'last4' | 'last8' | 'last12' | 'last' | 'penultimate' | 'past' | 'future' | 'custom_dates' | string;
  dateFrom?: string;
  dateTo?: string;
  originFilter: string;
  statusFilter: string;
  filters?: StudioFilterRule[];
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

function resolveWeekMetricValue(w: any, fieldKey: string, app: any, allAppts: any[], goalObj?: any): number {
  const cfg = app.cfg || {};
  const pct = app.pct || ((a: number, b: number) => b ? (a / b * 100) : 0);
  const normIso = app.normIso || ((s: string) => s);
  const msgsFn = app.msgs || ((x: any) => (x.fu2 || 0) + (x.fu3 || 0) + (x.fu4 || 0) + (x.fu5 || 0));
  const totConsultas = app.weekTotalConsultas
    ? app.weekTotalConsultas(w)
    : Math.max(w.ctot || 0, w.cprev || 0, (w.creal || 0) + (w.creag || 0) + (w.ccanc || 0) + (w.cfaltou || 0));

  const customM = (cfg.metrics || []).find((m: any) => m.id === fieldKey || m.title === fieldKey);
  if (customM && app.ev) {
    try { return Number(app.ev(customM.expr, [w])) || 0; } catch { return 0; }
  }

  const isCprevAlias = fieldKey === 'cprev' || (app.isCprevFieldKey && app.isCprevFieldKey(fieldKey));
  if (isCprevAlias) {
    if (w.cprevEffective != null) return Number(w.cprevEffective) || 0;
    if (app.getEffectiveCprev) return Number(app.getEffectiveCprev(w)) || 0;
    return Number(w.cprev) || 0;
  }
  if (fieldKey === 'msgs') return Number(msgsFn(w)) || 0;
  if (fieldKey === 'tot') return Number(totConsultas) || 0;
  if (fieldKey === 'taxa_resp' || (fieldKey === 'resp' && goalObj?.unit === 'pct')) {
    return pct(w.resp || 0, msgsFn(w));
  }
  if (fieldKey === 'taxa_conv' || fieldKey === 'conv') {
    return pct(w.agend || 0, w.resp || 0);
  }
  if (fieldKey === 'taxa_agend_geral' || fieldKey === 'tg') {
    return pct(w.cagend || 0, w.ativas || 0);
  }
  if (fieldKey === 'convNovos') {
    const start = w.id;
    const end = app.addDays ? app.addDays(w.id, 6) : w.id;
    const novosAg = allAppts.filter(a => {
      const s = normIso(a.schedOn) || normIso(a.consult);
      return a.origin === 'novo' && s >= start && s <= end;
    }).length;
    return pct(novosAg, w.novos || 0);
  }
  if (fieldKey === 'comp') {
    const t = (w.creal || 0) + (w.creag || 0) + (w.ccanc || 0) + (w.cfaltou || 0);
    return pct(w.creal || 0, t);
  }
  if (fieldKey === 'canc') {
    return goalObj?.unit === 'num' ? (Number(w.ccanc) || 0) : pct(w.ccanc || 0, totConsultas);
  }
  if (fieldKey === 'creag' && goalObj?.unit === 'pct') {
    return pct(w.creag || 0, totConsultas);
  }
  return Number(w[fieldKey]) || 0;
}

function matchScalarRule(rawVal: string | number, op: StudioFilterOp, val1: string, val2 = '', isNumeric = false): boolean {
  if (!val1 && op !== 'eq' && op !== 'neq') return true;
  if (isNumeric) {
    const num = Number(rawVal) || 0;
    const n1 = parseFloat(String(val1).replace(',', '.'));
    const n2 = parseFloat(String(val2).replace(',', '.'));
    if (isNaN(n1)) return true;
    if (op === 'eq') return num === n1;
    if (op === 'neq') return num !== n1;
    if (op === 'gte') return num >= n1;
    if (op === 'lte') return num <= n1;
    if (op === 'between') {
      if (!isNaN(n2)) return num >= Math.min(n1, n2) && num <= Math.max(n1, n2);
      return num >= n1;
    }
    return String(num).includes(String(val1));
  }
  const s = String(rawVal ?? '').trim().toLowerCase();
  const v1 = String(val1 ?? '').trim().toLowerCase();
  const v2 = String(val2 ?? '').trim().toLowerCase();
  if (op === 'eq') return s === v1;
  if (op === 'neq') return s !== v1;
  if (op === 'contains') return s.includes(v1);
  if (op === 'gte') return s >= v1;
  if (op === 'lte') return s <= v1;
  if (op === 'between') {
    if (v1 && s < v1) return false;
    if (v2 && s > v2) return false;
    return true;
  }
  return true;
}

const APPT_RULE_FIELDS = new Set(['origin', 'status', 'regained', 'statusNewDate', 'patient']);
const WEEK_TIME_FIELDS = new Set(['year', 'month', 'week_id', 'week_id_in', 'week_time', 'date_range']);

export function computeStudioChartData(c: StudioChartCfg) {
  const app = getApp();
  const ds = c.dataSource || 'semanas';
  const firstGoalSeries = (c.series || []).find(s => s.field && !s.field.startsWith('goal_target_'));
  const isFirstCprev = firstGoalSeries && (firstGoalSeries.field === 'cprev' || (app.isCprevFieldKey && app.isCprevFieldKey(firstGoalSeries.field)));
  const WA: any[] = (ds === 'metas' || (c.type === 'kpi' && isFirstCprev)) && app.getGoalChartWeeks
    ? app.getGoalChartWeeks(isFirstCprev ? 'cprev' : firstGoalSeries?.field)
    : (app.W ? app.W() : []);
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
  const todayStr = app.today ? app.today() : new Date().toISOString().slice(0, 10);
  const curSat = app.getSaturdayOfWeek ? app.getSaturdayOfWeek(todayStr) : todayStr;
  const rules: StudioFilterRule[] = Array.isArray(c.filters) ? c.filters : [];

  let filteredWeeks = [...WA];
  if (c.periodFilter === 'last4') filteredWeeks = WA.slice(-4);
  else if (c.periodFilter === 'last8') filteredWeeks = WA.slice(-8);
  else if (c.periodFilter === 'last12') filteredWeeks = WA.slice(-12);
  else if (c.periodFilter === 'last') filteredWeeks = WA.slice(-1);
  else if (c.periodFilter === 'penultimate') {
    filteredWeeks = WA.length >= 2 ? [WA[WA.length - 2]] : WA.slice(-1);
  } else if (c.periodFilter === 'past') {
    filteredWeeks = WA.filter(w => w.id < curSat && w.id < todayStr);
  } else if (c.periodFilter === 'future') {
    filteredWeeks = WA.filter(w => w.id >= curSat || w.id >= todayStr);
  } else if (c.periodFilter && c.periodFilter.startsWith('ym:')) {
    const ym = c.periodFilter.slice(3);
    filteredWeeks = WA.filter(w => w.id && w.id.startsWith(ym + '-'));
  }

  if (c.dateFrom) {
    filteredWeeks = filteredWeeks.filter(w => w.id >= c.dateFrom!);
  }
  if (c.dateTo) {
    filteredWeeks = filteredWeeks.filter(w => w.id <= c.dateTo!);
  }

  // Apply structured rules to weeks
  const numericWeekFields = new Set([
    ...(app.FIELDS || []).map((f: any) => f[0]),
    'cprev', 'msgs', 'tot', 'tg', 'conv', 'convNovos', 'comp', 'canc',
    'taxa_resp', 'taxa_conv', 'taxa_agend_geral',
    ...(cfg.metrics || []).flatMap((m: any) => [m.id, m.title])
  ]);

  if (rules.length > 0) {
    filteredWeeks = filteredWeeks.filter(w => {
      for (const r of rules) {
        if (!r || !r.field || APPT_RULE_FIELDS.has(r.field)) continue;
        if (ds === 'planilha' && !WEEK_TIME_FIELDS.has(r.field)) continue;
        if (r.field === 'year') {
          if (!r.value) continue;
          const y = String(w.id || '').slice(0, 4);
          if (!matchScalarRule(y, r.op, r.value)) return false;
        } else if (r.field === 'month') {
          if (!r.value) continue;
          const mNum = String(parseInt(String(w.id || '').split('-')[1] || '0', 10));
          const targetM = String(parseInt(r.value, 10));
          if (!matchScalarRule(mNum, r.op, targetM)) return false;
        } else if (r.field === 'week_id') {
          if (!r.value && !r.value2) continue;
          if (!matchScalarRule(w.id || '', r.op, r.value, r.value2)) return false;
        } else if (r.field === 'week_id_in') {
          const list = String(r.value || '').split(',').map(s => s.trim()).filter(Boolean);
          if (list.length) {
            const has = list.includes(w.id);
            if (r.op === 'neq' ? has : !has) return false;
          }
        } else if (r.field === 'week_time') {
          const isFuture = w.id >= curSat || w.id >= todayStr;
          const wantsFuture = r.value === 'future';
          if (r.op === 'neq' ? (isFuture === wantsFuture) : (isFuture !== wantsFuture)) return false;
        } else if (r.field === 'date_range') {
          if (r.value && w.id < r.value) return false;
          if (r.value2 && w.id > r.value2) return false;
        } else if (numericWeekFields.has(r.field)) {
          if (r.value === '' && r.value2 === '') continue;
          const val = resolveWeekMetricValue(w, r.field, app, allAppts);
          if (!matchScalarRule(val, r.op, r.value, r.value2, true)) return false;
        }
      }
      return true;
    });
  }

  let periodStart = '';
  let periodEnd = '';
  const hasWeekRestriction = (c.periodFilter && c.periodFilter !== 'all' && c.periodFilter !== 'custom_dates')
    || !!c.dateFrom
    || !!c.dateTo
    || rules.some(r => WEEK_TIME_FIELDS.has(r.field));
  if (hasWeekRestriction && filteredWeeks.length) {
    periodStart = filteredWeeks[0].id;
    periodEnd = addDays(filteredWeeks[filteredWeeks.length - 1].id, 6);
  }

  const filteredAppts = allAppts.filter(a => {
    if (c.originFilter && a.origin !== c.originFilter) return false;
    if (c.statusFilter && a.status !== c.statusFilter) return false;
    const dt = normIso(a.newDate) || normIso(a.consult) || normIso(a.schedOn);
    if (c.dateFrom && (!dt || dt < c.dateFrom)) return false;
    if (c.dateTo && (!dt || dt > c.dateTo)) return false;
    if (periodStart && periodEnd && !c.dateFrom && !c.dateTo) {
      const anyDt = normIso(a.consult) || normIso(a.schedOn) || normIso(a.newDate);
      if (!anyDt || anyDt < periodStart || anyDt > periodEnd) return false;
    }
    for (const r of rules) {
      if (!r || !r.field) continue;
      if (r.field === 'origin' && r.value !== '') {
        if (!matchScalarRule(a.origin || '', r.op, r.value)) return false;
      } else if (r.field === 'status' && r.value !== '') {
        if (!matchScalarRule(a.status || 'agendada', r.op, r.value)) return false;
      } else if (r.field === 'regained' && r.value !== '') {
        if (!matchScalarRule(a.regained || '', r.op, r.value)) return false;
      } else if (r.field === 'statusNewDate' && r.value !== '') {
        if (!matchScalarRule(a.statusNewDate || '', r.op, r.value)) return false;
      } else if (r.field === 'patient' && r.value !== '') {
        if (!matchScalarRule(a.patient || '', r.op, r.value)) return false;
      } else if (r.field === 'date_range') {
        if (r.value && (!dt || dt < r.value)) return false;
        if (r.value2 && (!dt || dt > r.value2)) return false;
      } else if (r.field === 'year' && r.value) {
        if (!dt || !matchScalarRule(dt.slice(0, 4), r.op, r.value)) return false;
      } else if (r.field === 'month' && r.value) {
        if (!dt) return false;
        const mNum = String(parseInt(dt.split('-')[1] || '0', 10));
        if (!matchScalarRule(mNum, r.op, String(parseInt(r.value, 10)))) return false;
      }
    }
    return true;
  });

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

  const goalsListAll = app.getUnifiedGoalsList ? app.getUnifiedGoalsList() : [];
  let matchedRowCount = 0;

  if (ds === 'semanas' || ds === 'metas') {
    filteredWeeks.forEach(w => {
      const bInfo = getGroupBucket(w.id, 'Semana de ');
      const extraParts = extraXFields.map(ef => String(w[ef] ?? '—'));
      const fullKey = extraParts.length ? `${bInfo.key}__${extraParts.join('__')}` : bInfo.key;
      const fullLabel = extraParts.length ? `${bInfo.label} · ${extraParts.join(' · ')}` : bInfo.label;
      const fullCat = extraParts.length ? `${bInfo.fullCat} · ${extraParts.join(' · ')}` : bInfo.fullCat;
      if (!buckets[fullKey]) buckets[fullKey] = { label: fullLabel, fullCat, weeks: [], appts: [], rows: goalsListAll };
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
      const filteredRows = sp.rows.filter((r: any) => {
        for (const rule of rules) {
          if (!rule || !rule.field || WEEK_TIME_FIELDS.has(rule.field) || APPT_RULE_FIELDS.has(rule.field)) continue;
          if (r[rule.field] === undefined && rule.value === '') continue;
          const cell = r[rule.field];
          const isNum = typeof cell === 'number' || (!isNaN(parseFloat(String(cell))) && isFinite(Number(cell)));
          if (!matchScalarRule(cell ?? '', rule.op, rule.value, rule.value2, isNum && (rule.op === 'gte' || rule.op === 'lte' || rule.op === 'between'))) {
            return false;
          }
        }
        if (c.dateFrom || c.dateTo) {
          const dateHeader = (sp.headers || []).find((h: string) => /data|date|dia|semana/i.test(h));
          if (dateHeader && r[dateHeader]) {
            const iso = normIso(String(r[dateHeader]));
            if (iso && /^\d{4}-\d{2}-\d{2}$/.test(iso)) {
              if (c.dateFrom && iso < c.dateFrom) return false;
              if (c.dateTo && iso > c.dateTo) return false;
            }
          }
        }
        return true;
      });
      matchedRowCount = filteredRows.length;
      filteredRows.forEach((r: any, idx: number) => {
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
    const linkedGoal = s.field.startsWith('goal_target_')
      ? goalsListAll.find((g: any) => g.id === s.field.replace('goal_target_', ''))
      : goalsListAll.find((g: any) => g.metric === s.field);

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

      if ((ds === 'metas' || s.source === 'metas') && s.field.startsWith('goal_target_')) {
        const gId = s.field.replace('goal_target_', '');
        const baseGoal = (b.rows || goalsListAll).find((g: any) => g.id === gId);
        if (!baseGoal) return 0;
        if (b.weeks.length && app.isGoalApplicableToWeek) {
          const targets = b.weeks.map((w: any) => {
            const spec = goalsListAll.find((g: any) => g.metric === baseGoal.metric && (g.scope === 'week' || g.type === 'specific') && g.week === w.id);
            if (spec) return Number(spec.target) || 0;
            const rec = goalsListAll.find((g: any) => g.metric === baseGoal.metric && app.isGoalApplicableToWeek(g, w.id));
            if (rec) return Number(rec.target) || 0;
            return Number(baseGoal.target) || 0;
          });
          return aggregateNums(targets, 'AVG');
        }
        return Number(baseGoal.target) || 0;
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
        const vals = b.weeks.map(w => resolveWeekMetricValue(w, 'msgs', app, allAppts));
        return aggregateNums(vals, s.agg);
      }
      if (s.field === 'taxa_resp') {
        const m = b.weeks.reduce((x, w) => x + resolveWeekMetricValue(w, 'msgs', app, allAppts), 0);
        const r = b.weeks.reduce((x, w) => x + (w.resp || 0), 0);
        return pct(r, m);
      }
      if (s.field === 'taxa_conv' || s.field === 'conv') {
        const r = b.weeks.reduce((x, w) => x + (w.resp || 0), 0);
        const a = b.weeks.reduce((x, w) => x + (w.agend || 0), 0);
        return pct(a, r);
      }
      if (s.field === 'taxa_agend_geral' || s.field === 'tg') {
        const at = b.weeks.reduce((x, w) => x + (w.ativas || 0), 0) / (b.weeks.length || 1);
        const ca = b.weeks.reduce((x, w) => x + (w.cagend || 0), 0);
        return pct(ca, at);
      }

      const vals = b.weeks.map(w => resolveWeekMetricValue(w, s.field, app, allAppts, linkedGoal));
      const dist = new Set(vals.map(String));
      const effectiveAgg = (['comp', 'canc', 'convNovos', 'resp'].includes(s.field) && linkedGoal?.unit === 'pct') ? 'AVG' : s.agg;
      return aggregateNums(vals, effectiveAgg, dist);
    });

    let finalValues = rawValues;
    const isBuiltInPct = ['taxa_resp', 'taxa_conv', 'taxa_agend_geral', 'tg', 'conv', 'convNovos', 'comp'].includes(s.field)
      || (['canc', 'resp', 'creag'].includes(s.field) && linkedGoal?.unit === 'pct')
      || (s.field.startsWith('goal_target_') && linkedGoal?.unit === 'pct')
      || Boolean((cfg.metrics || []).find((cm: any) => (cm.id === s.field || cm.title === s.field) && cm.fmt === 'pct'));

    if (s.agg === 'PCT' && !isBuiltInPct && !s.field.startsWith('goal_target_')) {
      const totalSum = rawValues.reduce((a, b) => a + Math.max(0, b), 0);
      finalValues = rawValues.map(v => totalSum > 0 ? (Math.max(0, v) / totalSum) * 100 : 0);
    }

    const isPct = isBuiltInPct || (s.agg === 'PCT' && !s.field.startsWith('goal_target_'));
    const dashed = Boolean(s.dashed || s.field.startsWith('goal_target_') || /^meta\b/i.test((s.name && String(s.name).trim()) || ''));
    return {
      n: (s.name && String(s.name).trim()) || s.field || `Série ${idx + 1}`,
      c: s.color || palette[idx % palette.length],
      v: finalValues.map(v => Math.round(v * 100) / 100),
      isPct,
      dashed
    };
  });

  return {
    labels,
    series,
    categories,
    meta: {
      weekCount: filteredWeeks.length,
      apptCount: filteredAppts.length,
      rowCount: matchedRowCount,
      weeks: filteredWeeks,
      appts: filteredAppts
    }
  };
}

export const computeChartData = computeStudioChartData;

export function renderStudioSvg(c: StudioChartCfg): { html: string } {
  const app = getApp();
  const esc = app.esc || ((s: any) => String(s ?? ''));
  const pf = app.pf || ((v: number) => v.toFixed(1).replace('.', ',') + '%');
  const { labels, series, categories, meta } = computeStudioChartData(c);

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
    const titleLow = String(c.title || '').toLowerCase();
    const field0 = String(c.series?.[0]?.field || '');
    let kpiKey = '';
    if (field0 === 'creal' || titleLow.includes('consultas realizadas')) kpiKey = 'creal';
    else if (field0 === 'cprev' || (app.isCprevFieldKey && app.isCprevFieldKey(field0)) || titleLow.includes('consultas previstas')) kpiKey = 'cprev';
    else if (field0 === 'cagend' || titleLow.includes('consultas agendadas')) kpiKey = 'cagend';
    else if (field0 === 'ccanc' || titleLow.includes('cancelad')) kpiKey = 'canc';
    else if (field0 === 'creag' || titleLow.includes('reagendad')) kpiKey = 'creag';
    if (kpiKey && c.id) {
      const w = window as any;
      w._studioKpiContexts = w._studioKpiContexts || {};
      w._studioKpiContexts[c.id] = {
        title: c.title,
        field0,
        weeks: meta?.weeks || [],
        appts: meta?.appts || [],
        kpiCount: totalVal
      };
    }
    const clickAttrs = kpiKey ? ` data-kpi-metric="${esc(kpiKey)}" data-kpi-chart-id="${esc(c.id || '')}" style="border-left:4px solid ${s.c};cursor:pointer" title="Clique para ver os pacientes desta métrica"` : ` style="border-left:4px solid ${s.c}"`;
    const clickHint = kpiKey ? `<small style="color:var(--g);font-weight:600;margin-top:4px">Ver pacientes →</small>` : '';
    return {
      html: `
        <div class="card kpi"${clickAttrs}>
          <span>${esc(c.title)}</span>
          <b>${displayVal}</b>
          <small>${esc(s.n)} (${labels.length} ponto(s))</small>
          ${clickHint}
        </div>
      `
    };
  }

  const leg = `<div class="lg" role="list" aria-label="Legenda"><b style="color:var(--ink)">Legenda:</b>${series.map(x => `<span role="listitem">${x.dashed ? `<i style="background:none;border-top:2.5px dashed ${x.c};height:0;width:15px;border-radius:0;vertical-align:middle"></i>` : `<i style="background:${x.c}"></i>`}${esc(x.n)}</span>`).join('')}</div>`;

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
    const orderedSeries = [
      ...series.filter(s => s.dashed),
      ...series.filter(s => !s.dashed)
    ];
    const solidSeriesList = orderedSeries.filter(s => !s.dashed);
    orderedSeries.forEach(s => {
      const isRef = Boolean(s.dashed);
      const solidIdx = isRef ? -1 : solidSeriesList.indexOf(s);
      const isMultiSolid = solidSeriesList.length > 1;
      const lineW = isRef ? '2' : (isMultiSolid && solidIdx === 0 ? '3.3' : (isMultiSolid && solidIdx > 0 ? '2.1' : '2.6'));
      const dotR = isRef ? '3.2' : (isMultiSolid && solidIdx === 0 ? '5.4' : (isMultiSolid && solidIdx > 0 ? '3.3' : '4.5'));
      const pts = s.v.map((v, i) => [L + bw * i + bw / 2, y(v || 0)]);
      if (c.type === 'area' && pts.length && !isRef) {
        const areaPath = `M ${pts[0][0]} ${T + ih} L ${pts.map(p => p.join(' ')).join(' L ')} L ${pts[pts.length - 1][0]} ${T + ih} Z`;
        o += `<path d="${areaPath}" fill="${s.c}" opacity="0.18"/>`;
      }
      o += `<polyline fill="none" stroke="${s.c}" stroke-width="${lineW}" ${isRef ? 'stroke-dasharray="6 4" opacity="0.88"' : ''} stroke-linecap="round" stroke-linejoin="round" points="${pts.map(p => p.join(',')).join(' ')}"/>`;
      pts.forEach((p, i) => {
        const v = s.v[i] || 0;
        const valStr = s.isPct ? pf(v) : v.toLocaleString('pt-BR');
        const cat = categories[i] || labels[i];
        o += `<circle cx="${p[0]}" cy="${p[1]}" r="${dotR}" fill="${s.c}" stroke="var(--card)" stroke-width="${isRef ? '1.5' : '1.8'}"
          data-tip-series="${esc(s.n)}" data-tip-color="${s.c}" data-tip-cat="${esc(cat)}" data-tip-val="${valStr}"></circle>`;
        o += `<circle cx="${p[0]}" cy="${p[1]}" r="15" fill="transparent"
          data-tip-series="${esc(s.n)}" data-tip-color="${s.c}" data-tip-cat="${esc(cat)}" data-tip-val="${valStr}"></circle>`;
      });
    });

    labels.forEach((_, i) => {
      const colPts = orderedSeries.map(s => {
        const isRef = Boolean(s.dashed);
        const v = s.v[i] || 0;
        const isFlatRef = isRef && s.v.every(x => x === s.v[0]);
        return {
          s,
          isRef,
          isFlatRef,
          cx: L + bw * i + bw / 2,
          cy: y(v),
          valStr: s.isPct ? pf(v) : v.toLocaleString('pt-BR')
        };
      });
      const placedLabels: { ly: number; cy: number; valStr: string; isRef: boolean }[] = [];
      const sortedPts = colPts.slice().sort((a, b) => {
        if (a.isRef !== b.isRef) return a.isRef ? 1 : -1;
        return a.cy - b.cy;
      });
      sortedPts.forEach(pt => {
        if (pt.isFlatRef && i > 0 && i < labels.length - 1) {
          const nearOther = colPts.some(oPt => !oPt.isRef && Math.abs(oPt.cy - pt.cy) < 16);
          if (nearOther) return;
        }
        if (!pt.isRef && placedLabels.some(pl => !pl.isRef && Math.abs(pl.cy - pt.cy) <= 2 && pl.valStr === pt.valStr)) {
          return;
        }
        let ly = pt.cy - 8;
        if (ly < 11) ly = pt.cy + 15;
        for (let attempt = 0; attempt < 4; attempt++) {
          const collision = placedLabels.some(pl => Math.abs(pl.ly - ly) < 12);
          if (!collision) break;
          if (attempt === 0) ly = pt.cy + 14;
          else if (attempt === 1) ly = pt.cy - 17;
          else ly = pt.cy + 22;
        }
        ly = Math.max(10, Math.min(T + ih + 10, ly));
        placedLabels.push({ ly, cy: pt.cy, valStr: pt.valStr, isRef: pt.isRef });
        o += `<text x="${pt.cx}" y="${ly}" text-anchor="middle" font-weight="${pt.isRef ? '600' : '700'}" font-size="9.5px" fill="${pt.isRef ? 'var(--mut)' : pt.s.c}" style="paint-order:stroke;stroke:var(--card);stroke-width:3.5px;stroke-linejoin:round">${pt.valStr}</text>`;
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
    dateFrom: '',
    dateTo: '',
    originFilter: '',
    statusFilter: '',
    filters: [],
    series: [
      { id: 's1', name: 'Novos leads', source: 'semanas', field: 'novos', agg: 'SUM', color: palette[0] },
      { id: 's2', name: 'Consultas agendadas', source: 'semanas', field: 'cagend', agg: 'SUM', color: palette[1] }
    ],
    showInDashboard: true
  };
}

function normalizeChartFilters(c: StudioChartCfg): StudioChartCfg {
  if (!Array.isArray(c.filters)) c.filters = [];
  if (c.originFilter && !c.filters.some(f => f.field === 'origin' && f.value === c.originFilter)) {
    c.filters.push({ id: 'mig_orig', field: 'origin', op: 'eq', value: c.originFilter });
  }
  c.originFilter = '';
  if (c.statusFilter && !c.filters.some(f => f.field === 'status' && f.value === c.statusFilter)) {
    c.filters.push({ id: 'mig_stat', field: 'status', op: 'eq', value: c.statusFilter });
  }
  c.statusFilter = '';
  return c;
}

function applyDashboardFilterToDraft(chartId: string, target: StudioChartCfg) {
  const app = getApp();
  if (!Array.isArray(target.filters)) target.filters = [];
  if (chartId === 'std_goals') {
    const gf = app.goalFilter || {};
    if (gf.dateFrom) target.dateFrom = gf.dateFrom;
    if (gf.dateTo) target.dateTo = gf.dateTo;
    if (gf.dateFrom || gf.dateTo) target.periodFilter = 'custom_dates';
    else if (gf.mode === 'future' || gf.mode === 'past') target.periodFilter = gf.mode;
    else if (gf.year && gf.year !== 'all' && gf.month && gf.month !== 'all') {
      target.periodFilter = 'ym:' + gf.year + '-' + String(Number(gf.month) + 1).padStart(2, '0');
    } else {
      if (gf.year && gf.year !== 'all' && !target.filters.some(f => f.field === 'year')) {
        target.filters.push({ id: 'gf_y_' + Date.now().toString(36), field: 'year', op: 'eq', value: String(gf.year) });
      }
      if (gf.month != null && gf.month !== 'all' && !target.filters.some(f => f.field === 'month')) {
        target.filters.push({ id: 'gf_m_' + Date.now().toString(36), field: 'month', op: 'eq', value: String(Number(gf.month) + 1) });
      }
    }
    if (gf.sort === 'desc') target.xAxisSort = 'desc';
  } else if (chartId === 'std_ag') {
    const af = app.agendaFilter || {};
    if (af.mode === 'future' || af.mode === 'past') target.periodFilter = af.mode;
    else if (af.year && af.year !== 'all' && af.month && af.month !== 'all') {
      target.periodFilter = 'ym:' + af.year + '-' + String(Number(af.month) + 1).padStart(2, '0');
    } else {
      if (af.year && af.year !== 'all' && !target.filters.some(f => f.field === 'year')) {
        target.filters.push({ id: 'af_y_' + Date.now().toString(36), field: 'year', op: 'eq', value: String(af.year) });
      }
      if (af.month != null && af.month !== 'all' && !target.filters.some(f => f.field === 'month')) {
        target.filters.push({ id: 'af_m_' + Date.now().toString(36), field: 'month', op: 'eq', value: String(Number(af.month) + 1) });
      }
    }
    if (af.sort === 'desc') target.xAxisSort = 'desc';
  } else if (['std_fu', 'std_rs', 'std_st', 'std_fl'].includes(chartId)) {
    const ef = app.excelFilter || {};
    if (ef.mode === 'last4' || ef.mode === 'last' || ef.mode === 'penultimate') {
      target.periodFilter = ef.mode;
    } else if (ef.mode === 'custom' && Array.isArray(ef.selected) && ef.selected.length > 0) {
      target.filters = target.filters.filter(f => f.field !== 'week_id_in');
      target.filters.push({
        id: 'ef_sel_' + Date.now().toString(36),
        field: 'week_id_in',
        op: 'eq',
        value: ef.selected.join(',')
      });
    }
    if (ef.sort === 'desc') target.xAxisSort = 'desc';
  }
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
    pushOpt({ value: 'cprev', label: 'Consultas previstas da semana', group: 'Campos Semanais', source: 'semanas' });
    FIELDS.forEach(([k]) => {
      pushOpt({ value: k, label: lab(k), group: 'Campos Semanais', source: 'semanas' });
    });
    pushOpt({ value: 'msgs', label: 'Mensagens de follow-up (FU 2 a 5)', group: 'Campos Semanais', source: 'semanas' });
    pushOpt({ value: 'tot', label: 'Total de consultas da semana', group: 'Campos Semanais', source: 'semanas' });
    pushOpt({ value: 'taxa_resp', label: 'Taxa de Resposta do Follow-up (%)', group: 'Taxas Calculadas', source: 'semanas' });
    pushOpt({ value: 'taxa_conv', label: 'Conversão do Follow-up (%)', group: 'Taxas Calculadas', source: 'semanas' });
    pushOpt({ value: 'taxa_agend_geral', label: 'Taxa Geral de Agendamento (%)', group: 'Taxas Calculadas', source: 'semanas' });
    pushOpt({ value: 'convNovos', label: 'Conversão de Novos Leads (%)', group: 'Taxas Calculadas', source: 'semanas' });
    pushOpt({ value: 'comp', label: 'Comparecimento (Check-in %)', group: 'Taxas Calculadas', source: 'semanas' });
    pushOpt({ value: 'canc', label: 'Taxa de Cancelamento (%)', group: 'Taxas Calculadas', source: 'semanas' });
    (cfg.metrics || []).forEach((m: any) => {
      pushOpt({ value: m.id, label: m.title, group: 'Métricas Personalizadas', source: 'semanas' });
    });
  }

  if (ds === 'metas') {
    const goals = app.getUnifiedGoalsList ? app.getUnifiedGoalsList() : [];
    goals.forEach((g: any) => {
      pushOpt({
        value: 'goal_target_' + g.id,
        label: `Meta: ${g.label || g.metric} (${g.target}${g.unit === 'pct' ? '%' : ''})`,
        group: 'Linhas de Meta',
        source: 'metas'
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
    draftChart = normalizeChartFilters(JSON.parse(JSON.stringify(found)));
  } else {
    const stdMap: Record<string, StudioChartCfg> = getStandardChartDefs();
    if (stdMap[chartId]) {
      editingChartId = chartId;
      draftChart = normalizeChartFilters(JSON.parse(JSON.stringify(stdMap[chartId])));
      applyDashboardFilterToDraft(chartId, draftChart);
    }
  }
  renderGraficosTabContent();
}

export function getStandardChartDefs(): Record<string, StudioChartCfg> {
  const palette = getThemeColors();
  const app = getApp();
  const goals = app.getUnifiedGoalsList ? app.getUnifiedGoalsList() : [];
  const selGoalId = app.selectedGoalChartId;
  const activeGoal = (selGoalId && goals.find((g: any) => g.id === selGoalId)) || goals[0];
  const goalSeries: SeriesItem[] = activeGoal
    ? (activeGoal.metric === 'cprev'
        ? [
            {
              id: 'g_act',
              name: activeGoal.label || activeGoal.metric,
              source: 'semanas',
              field: activeGoal.metric,
              agg: activeGoal.unit === 'pct' ? 'PCT' : 'SUM',
              color: palette[1]
            },
            {
              id: 'g_real',
              name: 'Consultas realizadas (de fato)',
              source: 'semanas',
              field: 'creal',
              agg: 'SUM',
              color: palette[0]
            },
            {
              id: 'g_tgt',
              name: `Meta (${activeGoal.target}${activeGoal.unit === 'pct' ? '%' : ''})`,
              source: 'metas',
              field: 'goal_target_' + activeGoal.id,
              agg: 'SUM',
              color: palette[2],
              dashed: true
            }
          ]
        : [
            {
              id: 'g_act',
              name: activeGoal.label || activeGoal.metric,
              source: 'semanas',
              field: activeGoal.metric,
              agg: activeGoal.unit === 'pct' ? 'PCT' : 'SUM',
              color: palette[0]
            },
            {
              id: 'g_tgt',
              name: `Meta (${activeGoal.target}${activeGoal.unit === 'pct' ? '%' : ''})`,
              source: 'metas',
              field: 'goal_target_' + activeGoal.id,
              agg: 'SUM',
              color: palette[2],
              dashed: true
            }
          ])
    : [
        { id: 'g_act', name: 'Consultas previstas da semana', source: 'semanas', field: 'cprev', agg: 'SUM', color: palette[1] },
        { id: 'g_real', name: 'Consultas realizadas (de fato)', source: 'semanas', field: 'creal', agg: 'SUM', color: palette[0] }
      ];

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
      filters: [],
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
      filters: [],
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
      filters: [],
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
      filters: [],
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
      filters: [],
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
      dateFrom: '',
      dateTo: '',
      originFilter: '',
      statusFilter: '',
      filters: [],
      series: goalSeries,
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
  const c = normalizeChartFilters(draftChart);
  if (!Array.isArray(c.xAxisSeries)) c.xAxisSeries = [];
  if (!Array.isArray(c.filters)) c.filters = [];

  const xOptions = getXAxisOptions(c.dataSource, c.spreadsheetId);
  if (!xOptions.some(o => o.value === c.xAxisField)) {
    c.xAxisField = xOptions[0]?.value || 'id';
  }
  const yOptions = getYFieldOptions(c.dataSource, c.spreadsheetId);
  const sheets: any[] = cfg.spreadsheets || [];
  const customCharts: StudioChartCfg[] = cfg.customCharts || [];
  const stdDefs = getStandardChartDefs();
  const WA_ALL: any[] = app.getGoalChartWeeks ? app.getGoalChartWeeks('cprev') : (app.W ? app.W() : []);
  const MONTH_NAMES: string[] = app.MONTH_NAMES || ['janeiro','fevereiro','março','abril','maio','junho','julho','agosto','setembro','outubro','novembro','dezembro'];
  const fdf = app.fdf || ((s: string) => s);

  const computedPreview = computeStudioChartData(c);
  const previewHtml = renderStudioSvg(c).html;
  const metricHtmlStr = app.metricHtml ? app.metricHtml() : '';

  const filterFieldOptions: { value: string; label: string; group: string }[] = [
    { value: 'date_range', label: 'Intervalo de Datas (De / Até)', group: 'Datas e Semanas' },
    { value: 'week_time', label: 'Janela Temporal (Passadas / Futuras)', group: 'Datas e Semanas' },
    { value: 'year', label: 'Ano (AAAA)', group: 'Datas e Semanas' },
    { value: 'month', label: 'Mês', group: 'Datas e Semanas' },
    { value: 'week_id', label: 'Semana específica (Data da semana)', group: 'Datas e Semanas' },
    { value: 'week_id_in', label: 'Lista de Semanas selecionadas', group: 'Datas e Semanas' }
  ];
  if (c.dataSource !== 'planilha') {
    filterFieldOptions.push(
      { value: 'cprev', label: 'Consultas previstas da semana', group: 'Indicadores Semanais' },
      { value: 'cagend', label: 'Consultas agendadas na semana', group: 'Indicadores Semanais' },
      { value: 'creal', label: 'Consultas realizadas na semana', group: 'Indicadores Semanais' },
      { value: 'ccanc', label: 'Consultas canceladas da semana', group: 'Indicadores Semanais' },
      { value: 'creag', label: 'Consultas reagendadas da semana', group: 'Indicadores Semanais' },
      { value: 'novos', label: 'Novos leads', group: 'Indicadores Semanais' },
      { value: 'ativas', label: 'Conversas ativas', group: 'Indicadores Semanais' },
      { value: 'resp', label: 'Respostas recebidas no follow-up', group: 'Indicadores Semanais' },
      { value: 'agend', label: 'Agendamentos pelo follow-up', group: 'Indicadores Semanais' },
      { value: 'origin', label: 'Origem do Paciente (Veio de)', group: 'Consultas' },
      { value: 'status', label: 'Status da Consulta', group: 'Consultas' },
      { value: 'regained', label: 'Status Contato (Reagendamento)', group: 'Consultas' },
      { value: 'statusNewDate', label: 'Status Nova Data', group: 'Consultas' },
      { value: 'patient', label: 'Nome do Paciente', group: 'Consultas' }
    );
  } else {
    const sp = sheets.find((s: any) => s.id === c.spreadsheetId) || sheets[0];
    (sp?.headers || []).forEach((h: string) => {
      filterFieldOptions.push({ value: h, label: h, group: 'Colunas da Planilha' });
    });
  }

  const periodLabels: Record<string, string> = {
    all: 'Todas as semanas',
    custom_dates: 'Intervalo de datas personalizado',
    past: 'Apenas semanas passadas',
    future: 'Apenas semanas futuras / atual',
    last4: 'Últimas 4 semanas',
    last8: 'Últimas 8 semanas',
    last12: 'Últimas 12 semanas',
    last: 'Última semana lançada',
    penultimate: 'Penúltima semana lançada'
  };

  const getRuleSummary = (r: StudioFilterRule) => {
    const fLbl = filterFieldOptions.find(o => o.value === r.field)?.label || r.field;
    const opMap: Record<StudioFilterOp, string> = {
      eq: '=',
      neq: '≠',
      contains: 'contém',
      gte: '≥',
      lte: '≤',
      between: 'entre'
    };
    let valDesc = r.value || '—';
    if (r.field === 'origin') valDesc = ORIG[r.value] || r.value || 'Qualquer';
    else if (r.field === 'status') valDesc = STAT[r.value] || r.value || 'Qualquer';
    else if (r.field === 'month' && r.value) valDesc = MONTH_NAMES[parseInt(r.value, 10) - 1] || r.value;
    else if (r.field === 'week_time') valDesc = r.value === 'future' ? 'Apenas futuras/atual' : 'Apenas passadas';
    else if (r.field === 'week_id' && r.value) valDesc = fdf(r.value);
    else if (r.field === 'week_id_in' && r.value) valDesc = `${r.value.split(',').filter(Boolean).length} semana(s)`;
    else if (r.field === 'date_range') valDesc = `${r.value ? fdf(r.value) : 'Início'} até ${r.value2 ? fdf(r.value2) : 'Hoje'}`;
    if (r.op === 'between' && r.field !== 'date_range' && r.value2) valDesc = `${r.value} e ${r.value2}`;
    return `${fLbl} ${opMap[r.op] || r.op} ${valDesc}`;
  };

  const renderFilterValueInputs = (r: StudioFilterRule, idx: number) => {
    const yearsAvail = Array.from(new Set(WA_ALL.map(w => String(w.id || '').slice(0, 4)).filter(Boolean))).sort();
    if (r.field === 'date_range') {
      return `
        <input type="date" data-gf-idx="${idx}" data-gf-k="value" value="${esc(r.value || '')}" style="min-width:130px" title="Data inicial">
        <span class="note" style="font-size:11px">até</span>
        <input type="date" data-gf-idx="${idx}" data-gf-k="value2" value="${esc(r.value2 || '')}" style="min-width:130px" title="Data final">
      `;
    }
    if (r.field === 'week_time') {
      return `
        <select data-gf-idx="${idx}" data-gf-k="value" style="min-width:160px">
          <option value="past" ${r.value === 'past' ? 'selected' : ''}>Apenas semanas passadas</option>
          <option value="future" ${r.value === 'future' ? 'selected' : ''}>Apenas futuras / atual</option>
        </select>
      `;
    }
    if (r.field === 'year') {
      return `
        <select data-gf-idx="${idx}" data-gf-k="value" style="min-width:120px">
          <option value="">Selecione...</option>
          ${yearsAvail.map(y => `<option value="${y}" ${r.value === y ? 'selected' : ''}>${y}</option>`).join('')}
        </select>
      `;
    }
    if (r.field === 'month') {
      return `
        <select data-gf-idx="${idx}" data-gf-k="value" style="min-width:150px">
          <option value="">Selecione...</option>
          ${MONTH_NAMES.map((m, i) => `<option value="${i + 1}" ${String(r.value) === String(i + 1) ? 'selected' : ''}>${m.charAt(0).toUpperCase() + m.slice(1)}</option>`).join('')}
        </select>
      `;
    }
    if (r.field === 'week_id') {
      if (r.op === 'between') {
        return `
          <input type="date" data-gf-idx="${idx}" data-gf-k="value" value="${esc(r.value || '')}" style="min-width:130px">
          <span class="note" style="font-size:11px">até</span>
          <input type="date" data-gf-idx="${idx}" data-gf-k="value2" value="${esc(r.value2 || '')}" style="min-width:130px">
        `;
      }
      return `
        <select data-gf-idx="${idx}" data-gf-k="value" style="min-width:170px">
          <option value="">Selecione uma semana...</option>
          ${WA_ALL.slice().reverse().map(w => `<option value="${w.id}" ${r.value === w.id ? 'selected' : ''}>${fdf(w.id)}</option>`).join('')}
        </select>
      `;
    }
    if (r.field === 'week_id_in') {
      const selArr = String(r.value || '').split(',').map(s => s.trim()).filter(Boolean);
      return `
        <input type="text" data-gf-idx="${idx}" data-gf-k="value" value="${esc(r.value || '')}" placeholder="Ex: 2026-10-03,2026-10-10" style="flex:1;min-width:180px" title="Datas separadas por vírgula (${selArr.length} selecionada(s))">
      `;
    }
    if (r.field === 'origin') {
      return `
        <select data-gf-idx="${idx}" data-gf-k="value" style="min-width:170px">
          <option value="">Todas</option>
          ${Object.entries(ORIG).map(([k, v]: any) => `<option value="${esc(k)}" ${r.value === k ? 'selected' : ''}>${esc(v)}</option>`).join('')}
        </select>
      `;
    }
    if (r.field === 'status') {
      return `
        <select data-gf-idx="${idx}" data-gf-k="value" style="min-width:170px">
          <option value="">Todos</option>
          ${Object.entries(STAT).map(([k, v]: any) => `<option value="${esc(k)}" ${r.value === k ? 'selected' : ''}>${esc(v)}</option>`).join('')}
        </select>
      `;
    }
    if (r.field === 'regained') {
      const regOpts = [['', 'Nenhum'], ['contato', 'Em contato'], ['sim', 'Conseguimos reagendar'], ['nao', 'Não conseguiu']];
      return `
        <select data-gf-idx="${idx}" data-gf-k="value" style="min-width:170px">
          ${regOpts.map(([k, v]) => `<option value="${k}" ${r.value === k ? 'selected' : ''}>${v}</option>`).join('')}
        </select>
      `;
    }
    if (r.field === 'statusNewDate') {
      const stNewOpts = [['', 'Pendente'], ['compareceu', 'Compareceu'], ['nao_compareceu', 'Não compareceu'], ['aguardando', 'Aguardando data'], ['reagendou', 'Reagendou novamente'], ['cancelou', 'Cancelou']];
      return `
        <select data-gf-idx="${idx}" data-gf-k="value" style="min-width:170px">
          ${stNewOpts.map(([k, v]) => `<option value="${k}" ${r.value === k ? 'selected' : ''}>${v}</option>`).join('')}
        </select>
      `;
    }
    return `
      <input type="text" data-gf-idx="${idx}" data-gf-k="value" value="${esc(r.value || '')}" placeholder="Valor..." style="min-width:110px;flex:1">
      ${r.op === 'between' ? `<span class="note" style="font-size:11px">e</span><input type="text" data-gf-idx="${idx}" data-gf-k="value2" value="${esc(r.value2 || '')}" placeholder="Até..." style="min-width:100px;flex:1">` : ''}
    `;
  };

  const activeBadges: { id: string; text: string; kind: 'period' | 'dateFrom' | 'dateTo' | 'rule'; idx?: number }[] = [];
  if (c.periodFilter && c.periodFilter !== 'all' && c.periodFilter !== 'custom_dates') {
    const pTxt = c.periodFilter.startsWith('ym:') ? `Mês: ${c.periodFilter.slice(3)}` : (periodLabels[c.periodFilter] || c.periodFilter);
    activeBadges.push({ id: 'b_period', text: `Período: ${pTxt}`, kind: 'period' });
  }
  if (c.dateFrom) activeBadges.push({ id: 'b_from', text: `Data inicial ≥ ${fdf(c.dateFrom)}`, kind: 'dateFrom' });
  if (c.dateTo) activeBadges.push({ id: 'b_to', text: `Data final ≤ ${fdf(c.dateTo)}`, kind: 'dateTo' });
  c.filters.forEach((r, idx) => {
    activeBadges.push({ id: 'b_r_' + idx, text: getRuleSummary(r), kind: 'rule', idx });
  });

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
          <p class="note" style="margin:3px 0 0">Configure livremente o Eixo X, o Eixo Y, o Modelo Visual e os Filtros exatos para analisar os dados em tempo real.</p>
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
      </div>

      <!-- Painel Preciso de Filtros Reais Aplicados e Filtro de Datas -->
      <div style="margin-top:14px;padding:14px;background:var(--bg);border:1px solid var(--line);border-radius:12px">
        <div style="display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:8px;margin-bottom:10px">
          <div>
            <h3 style="margin:0;font-size:14px">🎯 Filtros Aplicados, Datas e Regras</h3>
            <span class="note" style="font-size:12px">Visualize exatamente quais filtros estão ativos, edite o intervalo de datas, adicione novas condições ou remova filtros.</span>
          </div>
          <div style="display:flex;gap:6px;flex-wrap:wrap">
            ${editingChartId && String(editingChartId).startsWith('std_') ? `<button type="button" class="b s" id="gb_btn_sync_dash_filter" style="font-size:11.5px;padding:4px 9px" title="Puxar o filtro que está aplicado no Painel neste momento">🔄 Importar Filtro do Painel</button>` : ''}
            <button type="button" class="b s" id="gb_btn_clear_filters" style="font-size:11.5px;padding:4px 9px">🧹 Limpar Todos</button>
            <button type="button" class="b" id="gb_btn_add_filter" style="font-size:11.5px;padding:4px 10px">➕ Adicionar Filtro</button>
          </div>
        </div>

        <div id="gb_active_filter_badges" style="display:flex;flex-wrap:wrap;gap:6px;margin-bottom:12px;padding:8px 10px;background:var(--card);border:1px dashed var(--line);border-radius:8px;align-items:center">
          <span style="font-size:11.5px;font-weight:700;color:var(--mut);margin-right:4px">Filtros reais ativos:</span>
          ${activeBadges.length ? activeBadges.map(b => `
            <span class="tag" style="display:inline-flex;align-items:center;gap:5px;font-size:11.5px;background:color-mix(in srgb,var(--g) 12%,var(--card));color:var(--ink);border:1px solid color-mix(in srgb,var(--g) 30%,var(--line))">
              <span>${esc(b.text)}</span>
              <button type="button" class="x" data-gchip-kind="${b.kind}" ${b.idx != null ? `data-gchip-idx="${b.idx}"` : ''} style="font-size:14px;padding:0 3px;line-height:1" title="Excluir este filtro">×</button>
            </span>
          `).join('') : `<span class="note" style="font-size:12px">Nenhum filtro restritivo (exibindo todo o histórico disponível).</span>`}
        </div>

        <div class="f" style="margin-bottom:10px">
          <label>Período Rápido
            <select id="gb_period">
              <option value="all" ${c.periodFilter === 'all' ? 'selected' : ''}>Todas as semanas / datas</option>
              <option value="custom_dates" ${c.periodFilter === 'custom_dates' ? 'selected' : ''}>Intervalo de datas personalizado (De / Até)</option>
              <option value="past" ${c.periodFilter === 'past' ? 'selected' : ''}>Apenas semanas passadas</option>
              <option value="future" ${c.periodFilter === 'future' ? 'selected' : ''}>Apenas semanas futuras / atual</option>
              <option value="last4" ${c.periodFilter === 'last4' ? 'selected' : ''}>Últimas 4 semanas</option>
              <option value="last8" ${c.periodFilter === 'last8' ? 'selected' : ''}>Últimas 8 semanas</option>
              <option value="last12" ${c.periodFilter === 'last12' ? 'selected' : ''}>Últimas 12 semanas</option>
              <option value="last" ${c.periodFilter === 'last' ? 'selected' : ''}>Última semana lançada</option>
              <option value="penultimate" ${c.periodFilter === 'penultimate' ? 'selected' : ''}>Penúltima semana lançada</option>
              ${(() => {
                const ymSet = new Set<string>();
                WA_ALL.forEach(w => { if (w.id && w.id.length >= 7) ymSet.add(w.id.slice(0, 7)); });
                return Array.from(ymSet).sort().reverse().map(ym => {
                  const [y, m] = ym.split('-');
                  const mName = MONTH_NAMES[parseInt(m, 10) - 1] || m;
                  const lbl = mName.charAt(0).toUpperCase() + mName.slice(1) + '/' + y;
                  return `<option value="ym:${ym}" ${c.periodFilter === 'ym:' + ym ? 'selected' : ''}>Mês: ${lbl}</option>`;
                }).join('');
              })()}
            </select>
          </label>

          <label>Data Inicial (De)
            <input type="date" id="gb_date_from" value="${esc(c.dateFrom || '')}">
          </label>

          <label>Data Final (Até)
            <div style="display:flex;gap:6px;align-items:center">
              <input type="date" id="gb_date_to" value="${esc(c.dateTo || '')}" style="flex:1">
              ${(c.dateFrom || c.dateTo) ? `<button type="button" class="b s" id="gb_btn_clear_dates" style="font-size:11px;padding:5px 8px" title="Limpar datas">✕</button>` : ''}
            </div>
          </label>
        </div>

        ${c.filters.length ? `
          <div style="display:flex;flex-direction:column;gap:8px;margin-top:8px">
            ${c.filters.map((r, idx) => `
              <div class="row" style="margin:0;padding:8px 10px;background:var(--card);border:1px solid var(--line);border-radius:9px;align-items:center;gap:8px;flex-wrap:wrap">
                <span style="font-weight:700;font-size:11.5px;color:var(--mut)">Filtro #${idx + 1}</span>
                <select data-gf-idx="${idx}" data-gf-k="field" style="min-width:190px;flex:1">
                  ${(() => {
                    const groups: Record<string, { value: string; label: string }[]> = {};
                    filterFieldOptions.forEach(o => {
                      groups[o.group] = groups[o.group] || [];
                      groups[o.group].push(o);
                    });
                    return Object.entries(groups).map(([grp, items]) =>
                      `<optgroup label="${esc(grp)}">${items.map(o => `<option value="${esc(o.value)}" ${r.field === o.value ? 'selected' : ''}>${esc(o.label)}</option>`).join('')}</optgroup>`
                    ).join('');
                  })()}
                </select>
                ${r.field === 'date_range' ? `<span class="tag" style="font-size:11px">entre</span>` : `
                  <select data-gf-idx="${idx}" data-gf-k="op" style="min-width:130px">
                    <option value="eq" ${r.op === 'eq' ? 'selected' : ''}>Igual a (=)</option>
                    <option value="neq" ${r.op === 'neq' ? 'selected' : ''}>Diferente de (≠)</option>
                    <option value="gte" ${r.op === 'gte' ? 'selected' : ''}>Maior ou igual (≥)</option>
                    <option value="lte" ${r.op === 'lte' ? 'selected' : ''}>Menor ou igual (≤)</option>
                    <option value="between" ${r.op === 'between' ? 'selected' : ''}>Entre (De / Até)</option>
                    <option value="contains" ${r.op === 'contains' ? 'selected' : ''}>Contém texto</option>
                  </select>
                `}
                ${renderFilterValueInputs(r, idx)}
                <button type="button" class="x" data-gf-del="${idx}" title="Excluir este filtro" style="font-size:16px;padding:3px 7px;color:var(--r)">🗑️</button>
              </div>
            `).join('')}
          </div>
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
            <span class="note" id="gb_preview_counts" style="font-size:11.5px">Exibindo ${c.dataSource === 'planilha' ? `${computedPreview.meta.rowCount} linha(s)` : `${computedPreview.meta.weekCount} semana(s) · ${computedPreview.meta.apptCount} consulta(s)`} · Passe o mouse sobre os valores para ver detalhes</span>
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
    const comp = computeStudioChartData(c);
    const area = document.getElementById('gb_preview_area');
    if (area) area.innerHTML = renderStudioSvg(c).html;
    const cnt = document.getElementById('gb_preview_counts');
    if (cnt) {
      cnt.textContent = `Exibindo ${c.dataSource === 'planilha' ? `${comp.meta.rowCount} linha(s)` : `${comp.meta.weekCount} semana(s) · ${comp.meta.apptCount} consulta(s)`} · Passe o mouse sobre os valores para ver detalhes`;
    }
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
  container.querySelector('#gb_period')?.addEventListener('change', (e: any) => {
    c.periodFilter = e.target.value;
    if (c.periodFilter !== 'custom_dates' && c.periodFilter !== 'all') {
      // Keep dateFrom/dateTo if user only switched periodFilter? Clear dates if selecting a preset like last4/ym
      c.dateFrom = '';
      c.dateTo = '';
    }
    renderGraficosTabContent();
  });
  const onDateChange = (k: 'dateFrom' | 'dateTo', val: string) => {
    c[k] = val || '';
    if (c.dateFrom || c.dateTo) {
      c.periodFilter = 'custom_dates';
    } else if (c.periodFilter === 'custom_dates') {
      c.periodFilter = 'all';
    }
    renderGraficosTabContent();
  };
  container.querySelector('#gb_date_from')?.addEventListener('change', (e: any) => onDateChange('dateFrom', e.target.value));
  container.querySelector('#gb_date_to')?.addEventListener('change', (e: any) => onDateChange('dateTo', e.target.value));
  container.querySelector('#gb_btn_clear_dates')?.addEventListener('click', () => {
    c.dateFrom = '';
    c.dateTo = '';
    if (c.periodFilter === 'custom_dates') c.periodFilter = 'all';
    renderGraficosTabContent();
  });

  container.querySelector('#gb_btn_add_filter')?.addEventListener('click', () => {
    if (!Array.isArray(c.filters)) c.filters = [];
    const defaultField = c.dataSource === 'consultas' ? 'origin' : (c.dataSource === 'planilha' ? ((app.cfg?.spreadsheets?.[0]?.headers?.[0]) || 'date_range') : 'cprev');
    const defaultOp: StudioFilterOp = defaultField === 'cprev' ? 'gte' : 'eq';
    c.filters.push({
      id: 'f_' + Date.now().toString(36),
      field: defaultField,
      op: defaultOp,
      value: defaultField === 'cprev' ? '1' : '',
      value2: ''
    });
    renderGraficosTabContent();
  });

  container.querySelector('#gb_btn_clear_filters')?.addEventListener('click', () => {
    c.periodFilter = 'all';
    c.dateFrom = '';
    c.dateTo = '';
    c.originFilter = '';
    c.statusFilter = '';
    c.filters = [];
    renderGraficosTabContent();
    app.toast?.('Todos os filtros removidos.');
  });

  container.querySelector('#gb_btn_sync_dash_filter')?.addEventListener('click', () => {
    if (!editingChartId) return;
    c.periodFilter = 'all';
    c.dateFrom = '';
    c.dateTo = '';
    c.originFilter = '';
    c.statusFilter = '';
    c.filters = [];
    applyDashboardFilterToDraft(editingChartId, c);
    renderGraficosTabContent();
    app.toast?.('Filtros atuais do Painel importados!');
  });

  container.querySelectorAll('[data-gchip-kind]').forEach((btn: any) => {
    btn.onclick = () => {
      const kind = btn.dataset.gchipKind;
      if (kind === 'period') c.periodFilter = 'all';
      else if (kind === 'dateFrom') c.dateFrom = '';
      else if (kind === 'dateTo') c.dateTo = '';
      else if (kind === 'rule' && btn.dataset.gchipIdx != null && Array.isArray(c.filters)) {
        c.filters.splice(parseInt(btn.dataset.gchipIdx, 10), 1);
      }
      if (!c.dateFrom && !c.dateTo && c.periodFilter === 'custom_dates') c.periodFilter = 'all';
      renderGraficosTabContent();
    };
  });

  container.querySelectorAll('[data-gf-idx]').forEach((el: any) => {
    const idx = parseInt(el.dataset.gfIdx, 10);
    const k = el.dataset.gfK;
    const isSelect = el.tagName === 'SELECT';
    const isDate = el.type === 'date';
    el.addEventListener(isSelect || isDate ? 'change' : 'input', () => {
      if (!Array.isArray(c.filters) || !c.filters[idx] || !k) return;
      (c.filters[idx] as any)[k] = el.value;
      if (k === 'field') {
        c.filters[idx].value = el.value === 'week_time' ? 'past' : '';
        c.filters[idx].value2 = '';
        c.filters[idx].op = el.value === 'date_range' ? 'between' : (['cprev', 'cagend', 'creal', 'ccanc', 'creag', 'novos', 'ativas', 'resp', 'agend'].includes(el.value) ? 'gte' : 'eq');
        renderGraficosTabContent();
        return;
      }
      if (k === 'op') {
        renderGraficosTabContent();
        return;
      }
      if (isSelect || isDate) {
        renderGraficosTabContent();
      } else {
        refreshPreviewOnly();
      }
    });
  });

  container.querySelectorAll('[data-gf-del]').forEach((btn: any) => {
    btn.onclick = () => {
      const idx = parseInt(btn.dataset.gfDel, 10);
      if (Array.isArray(c.filters)) {
        c.filters.splice(idx, 1);
        renderGraficosTabContent();
      }
    };
  });

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
      const list = [...(app.cfg?.metrics || [])];
      const delM = list[idx];
      const newList = list.filter((_: any, i: number) => i !== idx);
      const h = { ...(app.cfg?.hide || {}) };
      if (delM) {
        if (!newList.some((x: any) => x.title === delM.title)) {
          delete h['k:' + delM.title];
          delete h['c:' + delM.title];
        }
        delete h['k:' + delM.id];
        delete h['c:' + delM.id];
      }
      const curGoals = app.getUnifiedGoalsList ? app.getUnifiedGoalsList() : (app.cfg?.goalsList || []);
      const newGoalsList = curGoals.filter((g: any) => !delM || (g.metric !== delM.id && g.metric !== delM.title && g.label !== delM.title));
      const newGoalsMap: Record<string, any> = {};
      newGoalsList.forEach((g: any) => {
        if ((g.scope === 'recurrent' || g.week === 'all' || !g.week) && !g.startWeek && !g.endWeek) {
          newGoalsMap[g.metric] = g.target;
        }
      });
      await app.saveCfg({ metrics: newList, hide: h, goalsList: newGoalsList, goals: newGoalsMap });
      renderGraficosTabContent();
      app.toast?.('Métrica removida.');
    }
  };
}
