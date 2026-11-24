/** charts.js - Chart.js views. Each render* function creates the chart once and then only updates data. */
import { fmtM, fmtUsd } from './calc.js';

export const COLORS = {
  blue: '#0078D4', teal: '#5ED6D6', tealLight: '#BDEFF0', green: '#A8D94F', greenDark: '#7FB02A',
  grey: '#9A9EA3', pos: '#4CAF50', neg: '#F2727B', text: '#3b4450', grid: '#e8ecf0',
};

const Chart = window.Chart;
Chart.defaults.font.family = "'Segoe UI', 'Noto Sans Thai', Arial, sans-serif";
Chart.defaults.color = COLORS.text;
Chart.defaults.animation.duration = 250;

/** tiny plugin: value labels at bar ends (so no extra library is needed) */
const barLabels = {
  id: 'barLabels',
  afterDatasetsDraw(chart, _a, opts) {
    if (!opts || !opts.enabled) return;
    const { ctx } = chart; const horizontal = chart.options.indexAxis === 'y';
    ctx.save(); ctx.font = `${opts.size || 11}px ${Chart.defaults.font.family}`; ctx.fillStyle = COLORS.text; ctx.textBaseline = 'middle';
    chart.data.datasets.forEach((ds, di) => {
      if (!chart.isDatasetVisible(di)) return;
      chart.getDatasetMeta(di).data.forEach((bar, i) => {
        const v = ds.data[i]; if (v === null || v === undefined || v === 0) return;
        const t = fmtM(v);
        if (horizontal) { ctx.textAlign = v >= 0 ? 'left' : 'right'; ctx.fillText(t, bar.x + (v >= 0 ? 4 : -4), bar.y); }
        else if (opts.vertical !== false && bar.width > 9) { ctx.textAlign = 'center'; ctx.textBaseline = 'bottom'; ctx.fillText(t, bar.x, bar.y - 2); ctx.textBaseline = 'middle'; }
      });
    });
    ctx.restore();
  },
};
Chart.register(barLabels);

const dayLabel = (iso) => { const [, m, d] = iso.split('-'); return `${['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][+m - 1]} ${d}`; };
const moneyTick = (v) => fmtM(v, Math.abs(v) % 1e6 === 0 ? 0 : 1);
const tooltipMoney = { callbacks: { label: (c) => `${c.dataset.label}: ${fmtUsd(c.parsed.x ?? c.parsed.y)} USD (${fmtM(c.parsed.x ?? c.parsed.y, 2)})` } };

const instances = {};
function upsert(id, config) {
  if (instances[id]) { const c = instances[id]; c.data = config.data; c.update(); return c; }
  instances[id] = new Chart(document.getElementById(id), config);
  return instances[id];
}

export function renderCumulative(s) {
  const ds = (label, data, color, extra = {}) => ({ label, data, borderColor: color, backgroundColor: color, pointBackgroundColor: color, borderWidth: 3, pointRadius: 3, pointHoverRadius: 5, tension: 0.15, spanGaps: false, ...extra });
  return upsert('chartCumulative', {
    type: 'line',
    data: {
      labels: s.labels.map(dayLabel),
      datasets: [
        ds('Forecast Cumulative Rolling', s.forecastRolling, COLORS.grey),
        ds('Cumulative PV(DPS)', s.pvDps, COLORS.teal),
        ds('Cumulative PV(Act)', s.pvAct, COLORS.green),
        ds('Forecast Cumulative PV', s.forecastPv, COLORS.greenDark, { borderDash: [2, 6], borderWidth: 3, backgroundColor: COLORS.green }),
      ],
    },
    options: {
      responsive: true, maintainAspectRatio: false, interaction: { mode: 'index', intersect: false },
      plugins: { legend: { position: 'top', align: 'start', labels: { usePointStyle: true, boxWidth: 8 } }, tooltip: tooltipMoney },
      scales: { y: { beginAtZero: true, ticks: { callback: moneyTick }, grid: { color: COLORS.grid } }, x: { grid: { display: false }, ticks: { maxTicksLimit: 8, maxRotation: 0 } } },
    },
  });
}

export function renderComparison(rows, onClickMrp) {
  return upsert('chartComparison', {
    type: 'bar',
    data: {
      labels: rows.map((r) => r.mrp),
      datasets: [
        { label: 'PV(Act)', data: rows.map((r) => r.pvAct), backgroundColor: COLORS.blue, barPercentage: 0.9, categoryPercentage: 0.8 },
        { label: 'Sum of Rolling AMT', data: rows.map((r) => r.rolling), backgroundColor: COLORS.teal, barPercentage: 0.9, categoryPercentage: 0.8 },
      ],
    },
    options: {
      indexAxis: 'y', responsive: true, maintainAspectRatio: false, layout: { padding: { right: 36 } },
      onClick: (_e, els) => { if (els.length && onClickMrp) onClickMrp(rows[els[0].index].mrp); },
      plugins: { legend: { position: 'top', align: 'start', labels: { usePointStyle: true, boxWidth: 8 } }, tooltip: tooltipMoney, barLabels: { enabled: true, size: 10 } },
      scales: { x: { beginAtZero: true, ticks: { callback: moneyTick }, grid: { color: COLORS.grid } }, y: { grid: { display: false }, reverse: false } },
    },
  });
}

export function renderGap(rows, onClickMrp) {
  const vals = rows.map((r) => r.gap);
  const c = upsert('chartGap', {
    type: 'bar',
    data: { labels: rows.map((r) => r.mrp), datasets: [{ label: 'GAP', data: vals, backgroundColor: vals.map((v) => (v >= 0 ? COLORS.pos : COLORS.neg)), barPercentage: 0.8, categoryPercentage: 0.85 }] },
    options: {
      indexAxis: 'y', responsive: true, maintainAspectRatio: false, layout: { padding: { left: 36, right: 36 } },
      onClick: (_e, els) => { if (els.length && onClickMrp) onClickMrp(rows[els[0].index].mrp); },
      plugins: { legend: { display: false }, tooltip: tooltipMoney, barLabels: { enabled: true, size: 10 } },
      scales: { x: { ticks: { callback: moneyTick }, grid: { color: COLORS.grid } }, y: { grid: { display: false }, ticks: { padding: 4 } } },
    },
  });
  c.data.datasets[0].backgroundColor = vals.map((v) => (v >= 0 ? COLORS.pos : COLORS.neg));   // recolour on update
  c.update();
  return c;
}

export function renderDaily(s) {
  return upsert('chartDaily', {
    type: 'bar',
    data: {
      labels: s.labels.map(dayLabel),
      datasets: [
        { label: 'PV(Act)', data: s.pvAct, backgroundColor: COLORS.blue, barPercentage: 1, categoryPercentage: 0.8 },
        { label: 'DPS_With_PV', data: s.dpsWith, backgroundColor: COLORS.teal, barPercentage: 1, categoryPercentage: 0.8 },
        { label: 'DPS_Without_PV', data: s.dpsWithout, backgroundColor: COLORS.tealLight, barPercentage: 1, categoryPercentage: 0.8 },
      ],
    },
    options: {
      responsive: true, maintainAspectRatio: false, layout: { padding: { top: 16 } },
      plugins: { legend: { position: 'top', align: 'start', labels: { usePointStyle: true, boxWidth: 8 } }, tooltip: { mode: 'index', ...tooltipMoney }, barLabels: { enabled: true, size: 10 } },
      scales: { y: { beginAtZero: true, ticks: { callback: moneyTick }, grid: { color: COLORS.grid } }, x: { grid: { display: false }, ticks: { maxTicksLimit: 8, maxRotation: 0 } } },
    },
  });
}
