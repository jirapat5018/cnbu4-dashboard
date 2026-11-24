import { CONFIG } from './config.js';
import { parseWorkbook } from './dataModel.js';
import { compute, fmtM, fmtPct } from './calc.js';
import { renderCumulative, renderComparison, renderGap, renderDaily } from './charts.js';
import { SampleConnector } from './connectors/sampleConnector.js';
import { SharePointConnector } from './connectors/sharepointConnector.js';
import { Auth } from './auth.js';
import { LocalFileConnector } from './connectors/localFileConnector.js';

const $ = (id) => document.getElementById(id);
const state = { model: null, meta: null, loadedAt: null, filters: { factories: new Set(), mrps: new Set(), from: '', to: '' }, busy: false };
let connector, auth, timer;

/* ------------------------------ UI helpers ------------------------------ */
const setLoading = (on, text = 'Loading…') => {
  state.busy = on; $('loadStatus').hidden = !on; $('loadStatus').lastChild.textContent = text; $('btnRefresh').disabled = on;
};
const showError = (msg) => { $('errorBox').textContent = msg; $('errorBox').hidden = !msg; };
const fmtDT = (d) => new Date(d).toLocaleString('th-TH', { dateStyle: 'medium', timeStyle: 'short' });

function renderStatus() {
  const m = state.meta || {};
  const file = m.lastModified ? `ไฟล์แก้ไขล่าสุด ${fmtDT(m.lastModified)}` : '';
  const loaded = state.loadedAt ? `โหลดเมื่อ ${fmtDT(state.loadedAt)}` : '';
  $('lastUpdated').textContent = `Last updated: ${[file, loaded].filter(Boolean).join(' | ') || '-'}`;
}

function chipRow(el, values, set, onChange) {
  el.innerHTML = '';
  for (const v of values) {
    const b = document.createElement('button');
    b.type = 'button'; b.className = 'chip' + (set.has(v) ? ' on' : ''); b.textContent = v; b.dataset.v = v;
    b.setAttribute('aria-pressed', set.has(v));
    b.onclick = () => { set.has(v) ? set.delete(v) : set.add(v); onChange(); };
    el.appendChild(b);
  }
}

/* ------------------------------ rendering ------------------------------ */
function buildFilters() {
  const m = state.model, f = state.filters;
  // drop selections that no longer exist in the new file
  f.factories = new Set([...f.factories].filter((x) => m.factories.includes(x)));
  f.mrps = new Set([...f.mrps].filter((x) => m.mrpList.includes(x)));
  chipRow($('factoryBtns'), m.factories, f.factories, render);
  chipRow($('mrpBtns'), m.mrpList, f.mrps, render);
  const first = m.dates[0], last = m.dates[m.dates.length - 1];
  for (const k of ['from', 'to']) if (!f[k] || f[k] < first || f[k] > last) f[k] = k === 'from' ? first : last;
  $('dFrom').min = $('dTo').min = first; $('dFrom').max = $('dTo').max = last;
  $('dFrom').value = f.from; $('dTo').value = f.to;
}

function render() {
  if (!state.model) return;
  const f = state.filters;
  document.querySelectorAll('#factoryBtns .chip').forEach((b) => { b.classList.toggle('on', f.factories.has(b.dataset.v)); b.setAttribute('aria-pressed', f.factories.has(b.dataset.v)); });
  document.querySelectorAll('#mrpBtns .chip').forEach((b) => { b.classList.toggle('on', f.mrps.has(b.dataset.v)); b.setAttribute('aria-pressed', f.mrps.has(b.dataset.v)); });
  const r = compute(state.model, { factories: f.factories, mrps: f.mrps, from: f.from, to: f.to }, CONFIG.rules);
  state.result = r;
  $('kpiPv').textContent = fmtM(r.kpi.pv, 2);
  $('kpiRolling').textContent = fmtM(r.kpi.rolling, 2);
  const ach = $('kpiAch'); ach.textContent = fmtPct(r.kpi.ach);
  ach.className = 'kpi-val ' + (r.kpi.ach !== null && r.kpi.ach >= CONFIG.rules.achGoodThreshold ? 'good' : 'bad');
  const g = $('gapTotal'); g.textContent = fmtM(r.kpi.gap, 2); g.classList.toggle('pos', r.kpi.gap >= 0);
  const toggleMrp = (m) => { f.mrps.has(m) ? f.mrps.delete(m) : f.mrps.add(m); render(); };   // click a bar = cross-filter
  renderCumulative(r.cumulative); renderComparison(r.byMrp, toggleMrp); renderGap(r.byMrp, toggleMrp); renderDaily(r.daily);
}

/* ------------------------------ data loading ------------------------------ */
async function loadData({ silent = false } = {}) {
  if (state.busy) return;
  showError('');
  if (!silent) setLoading(true, 'กำลังดึงข้อมูลล่าสุด…');
  try {
    const { buffer, meta } = await connector.fetchFile();
    if (!silent) setLoading(true, 'กำลังประมวลผลข้อมูล…');
    await new Promise((r) => setTimeout(r, 0));                   // let the spinner paint before parsing
    const model = parseWorkbook(buffer);
    state.model = model; state.meta = meta; state.loadedAt = new Date();
    const w = $('warnBox'); w.hidden = !model.warnings.length; w.textContent = model.warnings.join(' | ');
    $('dash').hidden = false; $('signinPanel').hidden = true; $('filePanel').hidden = true;
    buildFilters(); render(); renderStatus();
  } catch (e) {
    console.error(e);
    showError(`ไม่สามารถโหลดข้อมูลได้: ${e.message}`);
    if (e.status === 401 && auth) { $('signinPanel').hidden = false; }
  } finally { setLoading(false); }
}

/** cheap check: only re-download when the file on SharePoint really changed */
async function checkForUpdate() {
  if (state.busy || document.hidden) return;
  try { const m = await connector.getMeta(); if (!state.meta || m.version !== state.meta.version) await loadData({ silent: true }); }
  catch (e) { console.warn('auto-refresh check failed', e); }
}

/* ------------------------------ start ------------------------------ */
async function start() {
  document.title = CONFIG.title; $('title').textContent = CONFIG.title;
  if (CONFIG.logoUrl) { $('logo').src = CONFIG.logoUrl; $('logo').hidden = false; }

  $('btnRefresh').onclick = () => loadData();
  $('btnFull').onclick = () => (document.fullscreenElement ? document.exitFullscreen() : document.documentElement.requestFullscreen?.());
  $('btnClear').onclick = () => { state.filters.mrps.clear(); state.filters.factories.clear(); render(); };
  const dateChange = () => { state.filters.from = $('dFrom').value; state.filters.to = $('dTo').value; if (state.filters.from > state.filters.to) { [state.filters.from, state.filters.to] = [state.filters.to, state.filters.from]; $('dFrom').value = state.filters.from; $('dTo').value = state.filters.to; } render(); };
  $('dFrom').onchange = $('dTo').onchange = dateChange;
  window.addEventListener('resize', () => {});                    // Chart.js handles responsive resize itself

  if (CONFIG.dataSource === 'sharepoint') {
    try {
      auth = new Auth(CONFIG.sharepoint);
      await auth.init();
    } catch (e) { showError(`เริ่มระบบ Login ไม่ได้: ${e.message} — ตรวจสอบ tenantId/clientId/redirectUri ใน config.js`); return; }
    $('btnSignIn').onclick = () => auth.signIn();
    $('btnSignOut').onclick = () => auth.signOut();
    if (!auth.account) { $('signinPanel').hidden = false; return; }
    $('userBox').hidden = false; $('userName').textContent = auth.userName;
    connector = new SharePointConnector(CONFIG.sharepoint, () => auth.getToken());
  } else if (CONFIG.dataSource === 'file' || location.protocol === 'file:') {
    $('modeBadge').hidden = false; $('modeBadge').textContent = 'LOCAL FILE MODE';
    connector = new LocalFileConnector($('fileInput'));
    $('btnPick').onclick = () => loadData();
    $('filePanel').hidden = false;
    return;                                                          // wait for the user to pick a file (needs a click)
  } else {
    $('modeBadge').hidden = false; $('modeBadge').textContent = 'SAMPLE FILE MODE';
    connector = new SampleConnector(CONFIG.sample);
  }

  await loadData();
  if (CONFIG.autoRefreshMinutes > 0) timer = setInterval(checkForUpdate, CONFIG.autoRefreshMinutes * 60000);
}

window.__dash = { state, loadData, checkForUpdate, render };      // for tests / debugging
start();
