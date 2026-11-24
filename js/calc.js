/**
 * calc.js - pure functions (no DOM, no network) that turn the data model + filters into KPI / chart series.
 * Unit-tested in tests/run-tests.mjs against the figures in the reference dashboard.
 */

const EMPTY = new Set();
const asSet = (v) => (v instanceof Set ? v : new Set(v || []));

/**
 * @param model   result of parseWorkbook()
 * @param filters { factories:[], mrps:[], from:'YYYY-MM-DD', to:'YYYY-MM-DD' }  (empty list = all)
 * @param rules   CONFIG.rules
 */
export function compute(model, filters = {}, rules = {}) {
  const factories = asSet(filters.factories);
  const mrps = asSet(filters.mrps);
  const from = filters.from || model.dates[0];
  const to = filters.to || model.dates[model.dates.length - 1];
  const rollingDist = rules.rollingDistribution || 'dps';

  const passFM = (o) => (factories.size === 0 || factories.has(o.factory)) && (mrps.size === 0 || mrps.has(o.mrp));
  const inRange = (d) => d >= from && d <= to;
  const axis = model.dates.filter(inRange);
  const last = model.lastActualDate;                 // last day that has actual production (data "as of" date)

  /* ---- actual PV (Daily) ---- */
  const actByDay = {};
  const actByMrp = {};
  let pv = 0;
  for (const r of model.daily) {
    if (!passFM(r) || !inRange(r.date)) continue;
    pv += r.amt;
    actByDay[r.date] = (actByDay[r.date] || 0) + r.amt;
    actByMrp[r.mrp] = (actByMrp[r.mrp] || 0) + r.amt;
  }

  /* ---- DPS plan value per day (plan qty x UP x Check) ---- */
  const dpsByDay = {};
  const dpsAllDays = {};                             // unfiltered, used as fallback weights
  for (const p of model.dps) {
    for (const [d, v] of Object.entries(p.pv)) {
      dpsAllDays[d] = (dpsAllDays[d] || 0) + v;
      if (passFM(p)) dpsByDay[d] = (dpsByDay[d] || 0) + v;
    }
  }

  /* ---- Rolling AMT and its spread over the month ---- */
  const rollRows = model.rolling.filter(passFM);
  const rollingByMrp = {};
  let rollingMonth = 0;
  const rollPerDay = {};                             // spread over the WHOLE month (not only the range)
  const qtyByModel = new Map();
  for (const p of model.dps) {
    if (!qtyByModel.has(p.model)) qtyByModel.set(p.model, {});
    const q = qtyByModel.get(p.model);
    for (const [d, v] of Object.entries(p.qty)) q[d] = (q[d] || 0) + v;
  }
  const allDays = model.dates;
  const fallbackW = weights(allDays, dpsAllDays);
  for (const r of rollRows) {
    rollingMonth += r.rollingAmt;
    rollingByMrp[r.mrp] = (rollingByMrp[r.mrp] || 0) + r.rollingAmt;
    if (!r.rollingAmt) continue;
    let w;
    if (rollingDist === 'linear') w = allDays.map(() => 1 / allDays.length);
    else {
      w = weights(allDays, qtyByModel.get(r.model) || {});
      if (!w) w = fallbackW || allDays.map(() => 1 / allDays.length);   // model has no DPS plan
    }
    allDays.forEach((d, i) => { rollPerDay[d] = (rollPerDay[d] || 0) + r.rollingAmt * w[i]; });
  }
  const rollingInRange = axis.reduce((s, d) => s + (rollPerDay[d] || 0), 0);
  const rolling = rules.rollingFollowsDateRange ? rollingInRange : rollingMonth;

  /* ---- KPI ---- */
  const kpi = { pv, rolling, ach: rolling ? pv / rolling : null, gap: pv - rolling };

  /* ---- cumulative chart ---- */
  const cumulative = { labels: axis, forecastRolling: [], pvDps: [], pvAct: [], forecastPv: [] };
  let cR = 0, cD = 0, cA = 0, run = 0;
  for (const d of axis) {
    cR += rollPerDay[d] || 0;
    cD += dpsByDay[d] || 0;
    cumulative.forecastRolling.push(cR);
    cumulative.pvDps.push(cD);
    if (!last || d <= last) { cA += actByDay[d] || 0; run = cA; cumulative.pvAct.push(cA); }
    else { run += dpsByDay[d] || 0; cumulative.pvAct.push(null); }
    cumulative.forecastPv.push(last && d < last ? null : run);
  }

  /* ---- by MRP (comparison + GAP) ---- */
  const byMrp = [];
  for (const m of model.mrpList) {
    if (mrps.size && !mrps.has(m)) continue;
    const a = actByMrp[m] || 0, ro = rollingByMrp[m] || 0;
    if (!a && !ro) continue;
    byMrp.push({ mrp: m, pvAct: a, rolling: ro, gap: a - ro });
  }

  /* ---- daily chart ---- */
  const dailyChart = { labels: axis, pvAct: [], dpsWith: [], dpsWithout: [] };
  for (const d of axis) {
    const act = actByDay[d];
    const dps = dpsByDay[d];
    dailyChart.pvAct.push(act ? act : null);
    const isPast = !!last && d <= last;
    dailyChart.dpsWith.push(isPast && dps ? dps : null);
    dailyChart.dpsWithout.push(!isPast && dps ? dps : null);
  }

  return { kpi, cumulative, byMrp, daily: dailyChart, range: { from, to }, lastActualDate: last };
}

/** share of each day, proportional to qtyByDay; null when there is no quantity at all */
function weights(days, qtyByDay) {
  const arr = days.map((d) => qtyByDay[d] || 0);
  const s = arr.reduce((a, b) => a + b, 0);
  return s > 0 ? arr.map((v) => v / s) : null;
}

/* ---------- formatting helpers shared by UI and tests ---------- */
/** Power BI style: always millions ("0.1M", "-1.8M") */
export function fmtM(v, digits = 1) {
  if (v === null || v === undefined || Number.isNaN(v)) return '-';
  const t = (v / 1e6).toFixed(digits);
  return (Number(t) === 0 ? (0).toFixed(digits) : t) + 'M';
}
export const fmtPct = (v, d = 2) => (v === null || v === undefined ? '-' : (v * 100).toFixed(d) + '%');
export const fmtUsd = (v) => (v === null || v === undefined ? '-' : v.toLocaleString('en-US', { maximumFractionDigits: 0 }));
