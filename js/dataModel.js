/**
 * dataModel.js - turns the Excel workbook into a clean in-memory model.
 *
 * Column names / sheet names below are the REAL ones found in Rolling_Estimation_2611.xlsx:
 *   Rolling : Model, X, Customer, MRP, PC, Profit, MFG, UP, Rolling QTY, Rolling AMT, Actual Q'ty, Actual AMT, ...
 *   Daily   : Plnt, Profit Ctr, Material, Material Description, Order, ' Qty in UnE', MvT, SLoc, Date, Pallet,
 *             Check, UP, AMT, GR QTY            (table tDaily = A:N; columns P:W are helper formulas -> ignored)
 *   core    : tDate (Date, Index) A:B | tModel (Model, Profit, Factory, Customer) D:G | tMRP (MRP) I
 *   DPS     : MRP, Model, Profit, Check, UP, 1..30  (daily DPS plan, thousand pcs, index = day of month)
 *
 * Needs the SheetJS global `XLSX` (browser: vendor/xlsx.full.min.js; node tests: set globalThis.XLSX).
 * Excel formulas are NOT re-calculated here - the cached values saved by Excel are used
 * (same behaviour as Power BI reading the file).
 */

export class DataError extends Error {
  constructor(message, details) { super(message); this.name = 'DataError'; this.details = details; }
}

const SHEETS = { rolling: 'Rolling', daily: 'Daily', core: 'core', dps: 'DPS' };

const num = (v) => {
  if (typeof v === 'number') return Number.isFinite(v) ? v : 0;
  if (v === undefined || v === null || v === '') return 0;
  const n = Number(String(v).replace(/,/g, ''));
  return Number.isFinite(n) ? n : 0;            // '#N/A' etc. -> 0
};
const str = (v) => (v === undefined || v === null ? '' : String(v).trim());

/** Excel serial date -> 'YYYY-MM-DD' (no time-zone maths, so it is identical for every user) */
export function serialToISO(serial) {
  const ms = Math.round((serial - 25569) * 86400000);
  return new Date(ms).toISOString().slice(0, 10);
}

function sheetRows(wb, name, maxCol) {
  const ws = wb.Sheets[name];
  if (!ws) throw new DataError(`ไม่พบ Sheet "${name}" ใน Excel (Sheet ที่พบ: ${wb.SheetNames.join(', ')})`);
  const ref = XLSX.utils.decode_range(ws['!ref'] || 'A1:A1');
  const range = { s: { r: 0, c: 0 }, e: { r: ref.e.r, c: Math.min(ref.e.c, maxCol) } };
  return XLSX.utils.sheet_to_json(ws, { header: 1, range, raw: true, defval: undefined, blankrows: false });
}

function headerIndex(headerRow, sheet, required) {
  const idx = {};
  headerRow.forEach((h, i) => { const k = str(h); if (k && !(k in idx)) idx[k] = i; });
  const missing = required.filter((c) => !(c in idx));
  if (missing.length) throw new DataError(`Sheet "${sheet}" ไม่พบคอลัมน์: ${missing.join(', ')}`, { found: Object.keys(idx) });
  return idx;
}

/** @param {ArrayBuffer|Uint8Array} buffer */
export function parseWorkbook(buffer) {
  let wb;
  try {
    wb = XLSX.read(buffer, { type: 'array', cellFormula: false, cellStyles: false, cellDates: false, sheets: Object.values(SHEETS) });
  } catch (e) {
    throw new DataError('อ่านไฟล์ Excel ไม่ได้ (ไฟล์เสียหรือไม่ใช่ .xlsx)', e);
  }

  /* ---------------- core: calendar, model->factory, MRP list ---------------- */
  const coreRows = sheetRows(wb, SHEETS.core, 8);              // columns A..I only (L:S is another block)
  const ch = headerIndex(coreRows[0] || [], 'core', ['Date', 'Index', 'Model', 'Factory', 'MRP']);
  const dateByIndex = new Map();
  const dates = [];
  const modelFactory = new Map();
  const mrpList = [];
  for (let r = 1; r < coreRows.length; r++) {
    const row = coreRows[r];
    const d = row[ch.Date];
    if (typeof d === 'number') { const iso = serialToISO(d); dates.push(iso); dateByIndex.set(num(row[ch.Index]), iso); }
    const m = str(row[ch.Model]);
    if (m && !modelFactory.has(m)) modelFactory.set(m, str(row[ch.Factory]));
    const mrp = str(row[ch.MRP]);
    if (mrp && !mrpList.includes(mrp)) mrpList.push(mrp);
  }
  if (!dates.length) throw new DataError('Sheet "core" ไม่มีตารางวันที่ (tDate: Date/Index)');

  /* ---------------- Rolling ---------------- */
  const rRows = sheetRows(wb, SHEETS.rolling, 14);
  const rh = headerIndex(rRows[0] || [], 'Rolling', ['Model', 'MRP', 'MFG', 'UP', 'Rolling QTY', 'Rolling AMT', 'Actual AMT']);
  const rolling = [];
  const modelMrp = new Map();
  const modelMfgFactory = (mfg) => (mfg.toUpperCase() === 'CPSBU' ? 'CPSBU' : 'CNBU4');
  for (let r = 1; r < rRows.length; r++) {
    const row = rRows[r];
    const model = str(row[rh.Model]);
    if (!model) continue;
    const mrp = str(row[rh.MRP]);
    const factory = modelFactory.get(model) || modelMfgFactory(str(row[rh.MFG]));
    rolling.push({
      model, mrp, factory,
      customer: str(row[rh.Customer]), mfg: str(row[rh.MFG]),
      up: num(row[rh.UP]), rollingQty: num(row[rh['Rolling QTY']]), rollingAmt: num(row[rh['Rolling AMT']]),
      excelActualAmt: num(row[rh['Actual AMT']]),
    });
    modelMrp.set(model, { mrp, factory });
  }
  if (!rolling.length) throw new DataError('Sheet "Rolling" ไม่มีข้อมูล');

  /* ---------------- Daily (actual production) ---------------- */
  const dRows = sheetRows(wb, SHEETS.daily, 13);               // A..N only
  const dh = headerIndex(dRows[0] || [], 'Daily', ['Material', 'Date', 'Check', 'AMT', 'GR QTY']);
  const daily = [];
  let missingAmt = 0, unknownModels = new Set();
  for (let r = 1; r < dRows.length; r++) {
    const row = dRows[r];
    const material = str(row[dh.Material]);
    const dv = row[dh.Date];
    if (!material || typeof dv !== 'number') continue;
    if (row[dh.AMT] === undefined) missingAmt++;
    const info = modelMrp.get(material);
    if (!info) unknownModels.add(material);
    daily.push({
      date: serialToISO(dv), model: material,
      mrp: info ? info.mrp : '', factory: info ? info.factory : (modelFactory.get(material) || ''),
      amt: num(row[dh.AMT]), qty: num(row[dh['GR QTY']]), check: num(row[dh.Check]),
    });
  }
  if (daily.length && missingAmt / daily.length > 0.5) {
    throw new DataError('คอลัมน์ AMT/Check ใน Sheet "Daily" ไม่มีค่า (ไฟล์ไม่ได้ผ่านการคำนวณสูตรของ Excel) — เปิดไฟล์ด้วย Excel แล้ว Save ก่อนอัปโหลด');
  }

  /* ---------------- DPS (daily plan) ---------------- */
  const pRows = sheetRows(wb, SHEETS.dps, 40);
  const ph = headerIndex(pRows[0] || [], 'DPS', ['MRP', 'Model', 'Check', 'UP']);
  const dayCols = [];                                          // [{col, iso}]
  (pRows[0] || []).forEach((h, i) => {
    const n = Number(str(h));
    if (Number.isInteger(n) && n >= 1 && n <= 31 && dateByIndex.has(n)) dayCols.push({ col: i, iso: dateByIndex.get(n) });
  });
  if (!dayCols.length) throw new DataError('Sheet "DPS" ไม่พบคอลัมน์วันที่ (1..30)');
  const dps = [];
  for (let r = 1; r < pRows.length; r++) {
    const row = pRows[r];
    const model = str(row[ph.Model]);
    if (!model) continue;
    const info = modelMrp.get(model);
    const up = num(row[ph.UP]);
    const check = num(row[ph.Check]);
    const qty = {};                                            // iso -> pieces (sheet is in thousand pcs)
    for (const { col, iso } of dayCols) { const q = num(row[col]) * 1000; if (q) qty[iso] = q; }
    dps.push({
      mrp: str(row[ph.MRP]) || (info ? info.mrp : ''), model,
      factory: modelFactory.get(model) || (info ? info.factory : ''),
      up, check, qty,
      // value of the plan row per day (UP x pieces x Check)
      pv: Object.fromEntries(Object.entries(qty).map(([d, q]) => [d, q * up * check])),
    });
  }

  const actualDates = daily.filter((d) => d.amt !== 0 || d.qty !== 0).map((d) => d.date).sort();
  const lastActualDate = actualDates.length ? actualDates[actualDates.length - 1] : null;

  return {
    dates, mrpList, rolling, daily, dps, lastActualDate,
    factories: [...new Set(rolling.map((r) => r.factory).filter(Boolean))].sort(),
    warnings: unknownModels.size ? [`Daily มี Material ที่ไม่อยู่ใน Rolling ${unknownModels.size} รายการ (นับในยอดรวม แต่ไม่มี MRP)`] : [],
    stats: { rolling: rolling.length, daily: daily.length, dps: dps.length, days: dates.length },
  };
}
