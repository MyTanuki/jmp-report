// Smoke checks for the UI layer (no browser needed):
//   1. index.html loads the scripts in the required order.
//   2. Every "#id" that ui.js / app.js look up exists in index.html or in a
//      template string that ui.js renders.
//   3. ui.js and app.js load into a vm sandbox with a stub document, and
//      JMP.app.loadArrayBuffer() runs the real workbook through the full
//      load → classify → screen switch path without throwing.
//
// Usage: node test/ui-smoke.mjs

import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';
import { fileURLToPath } from 'node:url';
import XLSX from 'xlsx';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(__dirname, '..');
const XLSX_FILE = path.join(ROOT, 'report-01_01_2017-30_09_2026.xlsx');

let passed = 0, failed = 0;
function check(name, cond, detail) {
  if (cond) { passed++; console.log(`  ok   ${name}`); }
  else { failed++; console.log(`  FAIL ${name}${detail ? ` -> ${detail}` : ''}`); }
}

const html = fs.readFileSync(path.join(ROOT, 'index.html'), 'utf8');
const ui = fs.readFileSync(path.join(ROOT, 'js', 'ui.js'), 'utf8');
const app = fs.readFileSync(path.join(ROOT, 'js', 'app.js'), 'utf8');

// 1. script order --------------------------------------------------------------
console.log('\nindex.html');
const scripts = [...html.matchAll(/<script src="([^"]+)"/g)].map((m) => m[1]);
check('script order', JSON.stringify(scripts) === JSON.stringify([
  'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js',
  'https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js',
  'js/default-settings.js', 'js/settings.js', 'js/parser.js', 'js/aggregate.js', 'js/ui.js', 'js/app.js',
]), scripts.join(', '));
check('no ES module scripts', !/type="module"/.test(html));
check('stylesheet linked', /href="css\/style\.css"/.test(html));

// 2. DOM ids -------------------------------------------------------------------
console.log('\nDOM ids');
const definedIds = new Set();
for (const m of html.matchAll(/\sid="([^"]+)"/g)) definedIds.add(m[1]);
for (const m of ui.matchAll(/\sid="([^"]+)"/g)) definedIds.add(m[1]);           // rendered by ui.js
for (const m of ui.matchAll(/\sid="' \+ \(opts\.tableId \|\| '([^']+)'\)/g)) definedIds.add(m[1]);
for (const m of ui.matchAll(/tableId: '([^']+)'/g)) definedIds.add(m[1]);

const usedIds = new Set();
for (const src of [ui, app]) {
  for (const m of src.matchAll(/\$\(['"]#([\w-]+)['"]/g)) usedIds.add(m[1]);
  for (const m of src.matchAll(/\$all\(['"]#([\w-]+)/g)) usedIds.add(m[1]);
  for (const m of src.matchAll(/querySelector(?:All)?\(['"]#([\w-]+)/g)) usedIds.add(m[1]);
  for (const m of src.matchAll(/getElementById\(['"]([\w-]+)['"]\)/g)) usedIds.add(m[1]);
}
const missing = [...usedIds].filter((id) => !definedIds.has(id));
check(`every looked-up id is defined (${usedIds.size} ids)`, missing.length === 0, missing.join(', '));
for (const id of ['screen-load', 'screen-mapping', 'screen-dashboard', 'screen-settings', 'step-nav', 'tabs', 'tab-body', 'toast']) {
  check(`index.html has #${id}`, definedIds.has(id));
}
const screensInApp = [...app.matchAll(/'(load|mapping|dashboard|settings)'/g)].map((m) => m[1]);
check('app.js references all four screens', ['load', 'mapping', 'dashboard', 'settings'].every((s) => screensInApp.includes(s)));
const tabsInHtml = [...html.matchAll(/data-tab="(\w+)"/g)].map((m) => m[1]);
check('index.html tabs match ui.js tab keys', JSON.stringify(tabsInHtml) === JSON.stringify(['roomGroup', 'pivot', 'trend', 'invoices']));

// 3. vm load with stub document -------------------------------------------------
console.log('\nvm load');
function stubElement() {
  return {
    classList: { toggle() {}, add() {}, remove() {}, contains() { return false; } },
    addEventListener() {}, setAttribute() {}, getAttribute() { return null; },
    textContent: '', innerHTML: '', className: '', disabled: false, title: '', value: '',
    querySelector() { return null; }, querySelectorAll() { return []; },
    scrollIntoView() {},
  };
}
const document = {
  readyState: 'complete',
  addEventListener() {},
  querySelector() { return null; },
  querySelectorAll() { return []; },
  createElement() { return stubElement(); },
  body: { appendChild() {}, removeChild() {} },
};
const sandbox = {
  console, XLSX, document,
  setTimeout, clearTimeout,
  localStorage: (() => { const m = new Map(); return { getItem: (k) => m.get(k) ?? null, setItem: (k, v) => m.set(k, String(v)), removeItem: (k) => m.delete(k) }; })(),
  scrollTo() {}, confirm() { return true; },
  URL: { createObjectURL() { return 'blob:x'; }, revokeObjectURL() {} },
  Blob: class { constructor(parts) { this.parts = parts; } },
};
sandbox.window = sandbox;
sandbox.globalThis = sandbox;
vm.createContext(sandbox);
let loadError = null;
try {
  for (const f of ['settings.js', 'parser.js', 'aggregate.js', 'ui.js', 'app.js']) {
    vm.runInContext(fs.readFileSync(path.join(ROOT, 'js', f), 'utf8'), sandbox, { filename: f });
  }
} catch (e) { loadError = e; }
check('all scripts load without throwing', !loadError, loadError && loadError.stack);
const JMP = sandbox.window.JMP;
check('JMP.ui defined with render functions', !!JMP.ui && ['renderNav', 'renderLoad', 'renderMapping', 'renderDashboard', 'renderSettings'].every((k) => typeof JMP.ui[k] === 'function'));
check('JMP.app.loadArrayBuffer exposed', !!JMP.app && typeof JMP.app.loadArrayBuffer === 'function');
check('app initialised on load screen with settings', JMP.app.state.screen === 'load' && !!JMP.app.state.settings);
check('fmt uses th-TH 2 decimals', JMP.ui.fmt(-1234.5) === '-1,234.50', JMP.ui.fmt(-1234.5));
check('esc escapes html', JMP.ui.esc('<a href="x">&</a>') === '&lt;a href=&quot;x&quot;&gt;&amp;&lt;/a&gt;');

if (fs.existsSync(XLSX_FILE)) {
  const buf = fs.readFileSync(XLSX_FILE);
  const ab = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
  let err = null, rep = null;
  try { rep = JMP.app.loadArrayBuffer(ab, 'smoke.xlsx'); } catch (e) { err = e; }
  check('loadArrayBuffer runs without throwing', !err, err && err.stack);
  check('report stored in state (1783 invoices)', rep && JMP.app.state.report && JMP.app.state.report.invoices.length === 1783);
  check('stays on the load screen so the summary/warnings are shown', JMP.app.state.screen === 'load', JMP.app.state.screen);
  check('colMap covers all headers', Object.keys(JMP.app.state.colMap).length === 83);
  check('default statuses exclude แก้ไขบิล', JSON.stringify(JMP.app.state.filter.statuses.slice().sort()) === JSON.stringify(['ค้างชำระ', 'ชำระเงินแล้ว'].sort()), JSON.stringify(JMP.app.state.filter.statuses));

  // Actions that do not need a DOM
  let actErr = null;
  try {
    JMP.app.setScreen('mapping');
    JMP.app.actions.confirmAllSuggested();
    JMP.app.actions.setFilter({ year: 2024 });
    JMP.app.actions.setFilter({ statuses: [] });
    JMP.app.actions.saveStatusDefault(); // refused (empty) — must not persist an empty set
    JMP.app.actions.setFilter({ statuses: ['ชำระเงินแล้ว'] });
    JMP.app.actions.saveStatusDefault();
    JMP.app.actions.setTab('pivot');
    JMP.app.actions.showRoomInvoices('311');
    JMP.app.setScreen('settings');
    JMP.app.setScreen('dashboard');
  } catch (e) { actErr = e; }
  check('actions run without throwing', !actErr, actErr && actErr.stack);
  const st = JMP.app.state;
  // 82 suggested raw headers, but two of them are nikhahit spelling variants of one normalized key.
  check('confirmAllSuggested confirmed 81 normalized columns', Object.keys(st.settings.columns).length === 81, Object.keys(st.settings.columns).length);
  check('settings persisted to localStorage', !!sandbox.localStorage.getItem(JMP.settings.STORAGE_KEY));
  check('filter/tab state updated', st.filter.year === 2024 && st.tab === 'invoices' && st.invoiceRoom === '311');
  check('saveStatusDefault persisted the chosen (non-empty) set', JSON.stringify(st.settings.statusInclude) === JSON.stringify(['ชำระเงินแล้ว', 'ยืนยันชำระเงินโดยพนักงาน']), JSON.stringify(st.settings.statusInclude));
  check('no getState duplicate export', JMP.app.getState === undefined);
  check('filtered invoices computed for 2024', Array.isArray(st.filtered) && st.filtered.length > 0 && st.filtered.every((i) => i.year === 2024));
} else {
  check('real xlsx present', false, 'file missing');
}

console.log(`\n${passed} passed, ${failed} failed`);
if (failed) process.exit(1);
