// Node test runner for the plain browser scripts (settings.js / parser.js / aggregate.js).
// Loads them into a vm context with a fake `window`, `localStorage` and the npm `xlsx`
// build as the global `XLSX`, then checks them against the real invoice export.
//
// Usage: npm test   (run from the project root; requires report-01_01_2017-30_09_2026.xlsx)

import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';
import { fileURLToPath } from 'node:url';
import XLSX from 'xlsx';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(__dirname, '..');
const XLSX_FILE = path.join(ROOT, 'report-01_01_2017-30_09_2026.xlsx');
const EXPECTED_FILE = path.join(__dirname, 'expected-headers.json');

// ---------------------------------------------------------------------------
// Tiny assertion helpers
// ---------------------------------------------------------------------------
let passed = 0;
let failed = 0;
const failures = [];

function check(name, cond, detail) {
  if (cond) {
    passed++;
    console.log(`  ok   ${name}`);
  } else {
    failed++;
    failures.push(name + (detail !== undefined ? ` -> ${detail}` : ''));
    console.log(`  FAIL ${name}${detail !== undefined ? ` -> ${detail}` : ''}`);
  }
}

function eq(name, actual, expected) {
  const ok = JSON.stringify(actual) === JSON.stringify(expected);
  check(name, ok, ok ? undefined : `expected ${JSON.stringify(expected)}, got ${JSON.stringify(actual)}`);
}

function near(name, actual, expected, eps = 0.01) {
  const ok = Math.abs(actual - expected) <= eps;
  check(name, ok, ok ? undefined : `expected ${expected}, got ${actual}`);
}

function section(title) {
  console.log(`\n${title}`);
}

// ---------------------------------------------------------------------------
// Load plain scripts into a vm sandbox
// ---------------------------------------------------------------------------
function makeStorage() {
  const store = new Map();
  return {
    getItem: (k) => (store.has(k) ? store.get(k) : null),
    setItem: (k, v) => { store.set(k, String(v)); },
    removeItem: (k) => { store.delete(k); },
    clear: () => store.clear(),
    _store: store,
  };
}

function loadSandbox() {
  const sandbox = { console, XLSX, localStorage: makeStorage() };
  sandbox.window = sandbox;
  sandbox.globalThis = sandbox;
  vm.createContext(sandbox);
  for (const f of ['settings.js', 'parser.js', 'aggregate.js']) {
    const file = path.join(ROOT, 'js', f);
    const code = fs.readFileSync(file, 'utf8');
    vm.runInContext(code, sandbox, { filename: file });
  }
  return sandbox;
}

const sandbox = loadSandbox();
const JMP = sandbox.window.JMP;
const S = JMP.settings;
const P = JMP.parser;
const A = JMP.aggregate;

section('module loading');
check('window.JMP.settings exists', !!S && typeof S.classify === 'function');
check('window.JMP.parser exists', !!P && typeof P.parseWorkbook === 'function');
check('window.JMP.aggregate exists', !!A && typeof A.roomByGroup === 'function');
for (const f of ['settings.js', 'parser.js', 'aggregate.js']) {
  const code = fs.readFileSync(path.join(ROOT, 'js', f), 'utf8');
  check(`${f} starts with window.JMP guard`, code.startsWith('window.JMP = window.JMP || {};'));
  check(`${f} has no ES module syntax`, !/^\s*(import|export)\s/m.test(code));
}

// ---------------------------------------------------------------------------
// settings.js
// ---------------------------------------------------------------------------
section('settings: normalizeHeader');
eq('strip (non-vat) + trim', S.normalizeHeader('ค่าเช่าห้อง (non-vat)'), 'ค่าเช่าห้อง');
eq('case-insensitive / spacing', S.normalizeHeader('  ค่าเช่าห้อง   ( NON-VAT )  '), 'ค่าเช่าห้อง');
eq('collapse inner whitespace', S.normalizeHeader('ส่วนลด   พิเศษ\tครั้งที่ 1 (non-vat)'), 'ส่วนลด พิเศษ ครั้งที่ 1');
eq('nikhahit+sara aa → sara am', S.normalizeHeader('สำหรับ'), S.normalizeHeader('สําหรับ'));
eq('null → empty', S.normalizeHeader(null), '');
check('default rule patterns contain no nikhahit spelling',
  S.DEFAULT_SETTINGS.rules.every((r) => !r.pattern.includes('ํา')));
check('all default rules compile', S.DEFAULT_SETTINGS.rules.every((r) => { const re = S.compileRule(r); return re !== null && typeof re.test === 'function'; }));
eq('default groups count', S.DEFAULT_SETTINGS.groups.length, 17);
eq('default rules count', S.DEFAULT_SETTINGS.rules.length, 15);
eq('default statusInclude', S.DEFAULT_SETTINGS.statusInclude, ['ชำระเงินแล้ว', 'ค้างชำระ', 'ยืนยันชำระเงินโดยพนักงาน']);

section('settings: classify');
{
  const s = S.defaults();
  eq('suggested via rule', S.classify('ค่าเช่าห้อง', s).state, 'suggested');
  eq('suggested rule index (room_rent = rule 10)', S.classify('ค่าเช่าห้อง', s).rule.index, 9);
  eq('prepayment deduction → exclude before the generic ^หัก rule', S.classify('หัก ค่าเช่าห้อง (จ่ายล่วงหน้า)', s).group, 'exclude');
  eq('other หัก headers still → discount_other', S.classify('หัก ค่าน้ำ', s).group, 'discount_other');
  eq('unassigned when no rule matches', S.classify('ส่วนต่างค่าเช่าห้อง 201 และ 203 (ย้ายวันที่ 9/8/2020)', s), { group: null, state: 'unassigned' });
  s.columns['ค่าเช่าห้อง 12 วัน (1-12 มิถุนายน 2566)'] = 'discount_days';
  eq('confirmed wins over rule', S.classify('ค่าเช่าห้อง 12 วัน (1-12 มิถุนายน 2566)', s), { group: 'discount_days', state: 'confirmed' });
  s.columns['x'] = 'no_such_group';
  eq('confirmed to unknown group falls back to rules', S.classify('x', s).state, 'unassigned');
  s.rules.unshift({ pattern: '[', group: 'other' });
  eq('invalid regex rule is skipped', S.classify('ค่าเช่าห้อง', s).group, 'room_rent');
  const cm = S.columnMap(['ค่าน้ำ (non-vat)', 'zzz (non-vat)'], S.defaults());
  eq('columnMap keyed by raw header', cm, { 'ค่าน้ำ (non-vat)': 'water', 'zzz (non-vat)': null });
}

section('settings: persistence + import/export');
{
  sandbox.localStorage.clear();
  eq('load() without storage entry → defaults', S.load(), S.DEFAULT_SETTINGS);
  const s = S.defaults();
  s.columns['ค่าเช่าห้อง'] = 'room_rent';
  check('save() returns true', S.save(s) === true);
  check('storage key', sandbox.localStorage._store.has(S.STORAGE_KEY));
  eq('load() round-trips', S.load().columns, { 'ค่าเช่าห้อง': 'room_rent' });
  sandbox.localStorage.setItem(S.STORAGE_KEY, '{not json');
  eq('load() with corrupted storage → defaults', S.load().columns, {});
  const r = S.reset();
  eq('reset() → defaults and persists', S.load(), r);

  const text = S.exportJSON(s);
  check('exportJSON is pretty JSON', text.startsWith('{\n'));
  eq('importJSON round-trip', S.importJSON(text), s);

  const bad = [
    ['not json', '{'],
    ['array', '[]'],
    ['missing groups', '{"rules":[]}'],
    ['group without id', '{"groups":[{"label":"x","kind":"revenue"}],"rules":[]}'],
    ['bad kind', '{"groups":[{"id":"a","label":"x","kind":"nope"}],"rules":[]}'],
    ['duplicate id', '{"groups":[{"id":"a","kind":"revenue"},{"id":"a","kind":"revenue"}],"rules":[]}'],
    ['rule → unknown group', '{"groups":[{"id":"a","kind":"revenue"}],"rules":[{"pattern":"x","group":"b"}]}'],
    ['rule bad regex', '{"groups":[{"id":"a","kind":"revenue"}],"rules":[{"pattern":"[","group":"a"}]}'],
    ['column → unknown group', '{"groups":[{"id":"a","kind":"revenue"}],"rules":[],"columns":{"h":"zzz"}}'],
    ['parent missing', '{"groups":[{"id":"a","kind":"revenue","parent":"p"}],"rules":[]}'],
    ['parent self', '{"groups":[{"id":"a","kind":"revenue","parent":"a"}],"rules":[]}'],
    ['parent 2-cycle', '{"groups":[{"id":"a","kind":"discount","parent":"b"},{"id":"b","kind":"discount","parent":"a"}],"rules":[]}'],
    ['parent 3-cycle', '{"groups":[{"id":"a","kind":"discount","parent":"b"},{"id":"b","kind":"discount","parent":"c"},{"id":"c","kind":"discount","parent":"a"}],"rules":[]}'],
    ['parent not top-level', '{"groups":[{"id":"a","kind":"discount"},{"id":"b","kind":"discount","parent":"a"},{"id":"c","kind":"discount","parent":"b"}],"rules":[]}'],
    ['parent = other bucket', '{"groups":[{"id":"other","kind":"other"},{"id":"b","kind":"other","parent":"other"}],"rules":[]}'],
    ['rule → parent group', '{"groups":[{"id":"d","kind":"discount"},{"id":"c","kind":"discount","parent":"d"}],"rules":[{"pattern":"x","group":"d"}]}'],
    ['column → parent group', '{"groups":[{"id":"d","kind":"discount"},{"id":"c","kind":"discount","parent":"d"}],"rules":[],"columns":{"h":"d"}}'],
  ];
  for (const [name, txt] of bad) {
    let threw = false;
    try { S.importJSON(txt); } catch (e) { threw = /ไม่ใช่ JSON|ไม่ถูกต้อง/.test(e.message); }
    check(`importJSON rejects: ${name}`, threw);
  }
  const minimal = S.importJSON('{"groups":[{"id":"a","kind":"revenue"}],"rules":[],"columns":{"h (non-vat)":"a"}}');
  eq('importJSON fills defaults (label, order, statusInclude) and normalizes column keys',
    [minimal.groups[0].label, minimal.groups[0].order, minimal.statusInclude, minimal.columns],
    ['a', 10, S.DEFAULT_SETTINGS.statusInclude, { h: 'a' }]);
  eq('importJSON adds the missing "other" bucket group', minimal.groups.map((g) => [g.id, g.kind]), [['a', 'revenue'], ['other', 'other']]);
  const tied = S.importJSON('{"groups":[{"id":"a","kind":"revenue","order":0},{"id":"b","kind":"revenue","order":0},{"id":"c","kind":"revenue","order":0}],"rules":[]}');
  check('moveGroup with tied orders moves on the first call', S.moveGroup(tied, 'b', -1) === true);
  eq('moveGroup with tied orders: b before a', S.sortedGroups(tied).map((g) => g.id), ['b', 'a', 'c', 'other']);
}

section('settings: shipped defaults (js/default-settings.js)');
{
  sandbox.localStorage.clear();
  eq('no shipped file → shippedDefaults() null', S.shippedDefaults(), null);

  const shipped = S.importJSON(S.exportJSON(S.DEFAULT_SETTINGS));
  shipped.columns = { 'ค่าเช่าห้อง 12 วัน (1-12 มิถุนายน 2566)': 'discount_days' };
  shipped.statusInclude = ['ชำระเงินแล้ว'];
  sandbox.JMP_DEFAULT_SETTINGS = shipped;
  eq('defaults() returns the shipped settings', S.defaults(), shipped);
  S.defaults().columns.x = 'room_rent';
  eq('defaults() hands out a fresh copy each call', S.defaults().columns, shipped.columns);
  eq('load() without storage entry → shipped defaults', S.load(), shipped);
  const mine = S.defaults();
  mine.columns = { 'ค่าน้ำ': 'water' };
  S.save(mine);
  eq('saved settings win over shipped defaults', S.load().columns, { 'ค่าน้ำ': 'water' });
  eq('reset() → shipped defaults and persists', [S.reset(), S.load()], [shipped, shipped]);
  eq('DEFAULT_SETTINGS (built-in) is untouched', [S.DEFAULT_SETTINGS.columns, S.DEFAULT_SETTINGS.statusInclude.length], [{}, 3]);

  // defaultSettingsJS() output is a script that reproduces the settings exactly.
  const js = S.defaultSettingsJS(mine, '2026-10-01T00:00:00.000Z */ x');
  const probe = {}; probe.window = probe; vm.createContext(probe);
  vm.runInContext(js, probe);
  eq('defaultSettingsJS round-trips through a script', JSON.parse(JSON.stringify(probe.JMP_DEFAULT_SETTINGS)), mine);
  check('defaultSettingsJS notes the save time', js.includes('Saved: 2026-10-01T00:00:00.000Z'));

  // An invalid shipped file must not break the app.
  const realWarn = console.warn;
  let warned = 0;
  console.warn = () => { warned++; };
  sandbox.JMP_DEFAULT_SETTINGS = { groups: [{ id: 'a', kind: 'nope' }], rules: [] };
  sandbox.localStorage.clear();
  eq('invalid shipped file → built-in defaults', S.load(), S.DEFAULT_SETTINGS);
  console.warn = realWarn;
  check('invalid shipped file is reported on the console', warned > 0);

  delete sandbox.JMP_DEFAULT_SETTINGS;
  sandbox.localStorage.clear();
  eq('shipped file removed → built-in defaults again', S.defaults(), S.DEFAULT_SETTINGS);

  // The file actually committed in js/ must be valid.
  const file = path.join(ROOT, 'js', 'default-settings.js');
  if (fs.existsSync(file)) {
    const box = { console }; box.window = box; box.globalThis = box; vm.createContext(box);
    vm.runInContext(fs.readFileSync(file, 'utf8'), box, { filename: file });
    const v = S.validate(JSON.parse(JSON.stringify(box.JMP_DEFAULT_SETTINGS)));
    check('js/default-settings.js is valid settings', v.ok === true, (v.errors || []).join('; '));
    check('js/default-settings.js keeps the "other" bucket and at least one rule',
      v.ok && v.settings.groups.some((g) => g.id === S.UNASSIGNED_GROUP) && v.settings.rules.length > 0);
  }
}

section('settings: group helpers');
{
  const s = S.defaults();
  check('discount parent is not assignable', !S.assignableGroups(s).some((g) => g.id === 'discount'));
  eq('assignable count', S.assignableGroups(s).length, 16);
  eq('childGroups(discount)', S.childGroups(s, 'discount').map((g) => g.id),
    ['discount_new', 'discount_special', 'discount_days', 'discount_fridge', 'discount_tv', 'discount_other']);
  const g = S.addGroup(s, { label: 'ค่าที่จอดรถ', kind: 'revenue' });
  check('addGroup generates unique id', !!g.id && S.groupById(s, g.id) === g);
  const g2 = S.addGroup(s, { id: g.id, label: 'dup', kind: 'other' });
  check('addGroup avoids id collision', g2.id !== g.id);
  S.renameGroup(s, g.id, 'ที่จอดรถ');
  eq('renameGroup', S.groupById(s, g.id).label, 'ที่จอดรถ');
  S.updateGroup(s, g.id, { kind: 'other', parent: 'discount' });
  eq('updateGroup kind/parent', [S.groupById(s, g.id).kind, S.groupById(s, g.id).parent], ['other', 'discount']);
  let threw = false;
  try { S.updateGroup(s, g.id, { parent: g.id }); } catch (e) { threw = true; }
  check('updateGroup rejects self parent', threw);
  const throws = (fn) => { try { fn(); return false; } catch (e) { return /วนซ้ำ|ระดับบนสุด|กลุ่มแม่|กลุ่มย่อย|ลบกลุ่ม/.test(e.message); } };
  check('updateGroup rejects a parent cycle (discount → discount_new)', throws(() => S.updateGroup(s, 'discount', { parent: 'discount_new' })));
  check('updateGroup rejects a non-top-level parent', throws(() => S.updateGroup(s, 'room_rent', { parent: 'discount_new' })));
  check('updateGroup rejects giving a container a parent', throws(() => S.updateGroup(s, 'discount', { parent: 'room_rent' })));
  check('updateGroup rejects the other bucket as parent', throws(() => S.updateGroup(s, 'room_rent', { parent: 'other' })));
  check('setColumn rejects a parent group', throws(() => S.setColumn(s, 'zzz', 'discount')));
  check('addRule rejects a parent group', throws(() => S.addRule(s, { pattern: 'zzz', group: 'discount' })));
  check('updateRule rejects a parent group', throws(() => S.updateRule(s, 0, { group: 'discount' })));
  check('deleteGroup refuses the other bucket', throws(() => S.deleteGroup(s, 'other')));
  check('other bucket still present', !!S.groupById(s, 'other'));
  {
    // A group that gains its first child hands its columns and rules to that child.
    const s2 = S.defaults();
    S.setColumn(s2, 'ค่าเช่าห้อง', 'room_rent');
    const child = S.addGroup(s2, { label: 'rent monthly', kind: 'revenue', parent: 'room_rent' });
    eq('addGroup under a leaf moves its columns to the child', s2.columns, { 'ค่าเช่าห้อง': child.id });
    check('addGroup under a leaf moves its rules to the child', s2.rules.some((r) => r.group === child.id) && !s2.rules.some((r) => r.group === 'room_rent'));
    check('room_rent is now a container', S.isParent(s2, 'room_rent') && !S.assignableGroups(s2).some((g) => g.id === 'room_rent'));
    check('validate accepts the result', S.validate(JSON.parse(S.exportJSON(s2))).ok);
  }
  S.setColumn(s, 'colA', g.id);
  S.setColumn(s, 'colB', 'room_rent');
  S.addRule(s, { pattern: 'จอดรถ', group: g.id }, 0);
  const res = S.deleteGroup(s, g.id);
  eq('deleteGroup un-maps its columns', res.removedColumns, ['colA']);
  eq('deleteGroup keeps other columns', s.columns, { colB: 'room_rent' });
  eq('deleteGroup removes rules targeting it', res.removedRules, 1);
  check('deleteGroup removes the group', !S.groupById(s, g.id));
  S.clearColumn(s, 'colB');
  eq('clearColumn', s.columns, {});
  const before = S.sortedGroups(s).map((x) => x.id);
  S.moveGroup(s, 'furniture', -1);
  const after = S.sortedGroups(s).map((x) => x.id);
  eq('moveGroup up swaps with previous sibling', [after[0], after[1]], [before[1], before[0]]);
  check('moveGroup at top edge returns false', S.moveGroup(s, after[0], -1) === false);
  S.moveRule(s, 0, 1);
  eq('moveRule', s.rules[1].group, 'water');
  S.updateRule(s, 1, { pattern: '^ค่าน้ำ' });
  eq('updateRule', s.rules[1].pattern, '^ค่าน้ำ');
  check('deleteRule', S.deleteRule(s, 1) && s.rules.length === 14);
  eq('testRule', S.testRule({ pattern: '^ค่าน้ำ' }, ['ค่าน้ำ (non-vat)', 'ค่าไฟฟ้า (non-vat)']).matches, ['ค่าน้ำ']);
}

// ---------------------------------------------------------------------------
// parser.js — unit
// ---------------------------------------------------------------------------
section('parser: parseDate / parseRoom');
eq('dd/mm/yyyy', P.parseDate('04/03/2019'), { date: '2019-03-04', year: 2019, month: 3, ym: '2019-03' });
eq('d/m/yyyy', P.parseDate('4/3/2019'), { date: '2019-03-04', year: 2019, month: 3, ym: '2019-03' });
eq('yyyy-mm-dd', P.parseDate('2019-03-04'), { date: '2019-03-04', year: 2019, month: 3, ym: '2019-03' });
eq('Excel serial 43528 = 2019-03-04', P.parseDate(43528), { date: '2019-03-04', year: 2019, month: 3, ym: '2019-03' });
eq('JS Date', P.parseDate(new Date(2019, 2, 4, 12)), { date: '2019-03-04', year: 2019, month: 3, ym: '2019-03' });
eq('Buddhist year converted', P.parseDate('04/03/2562').year, 2019);
eq('invalid string → null', P.parseDate('รวม'), null);
eq('empty → null', P.parseDate(null), null);
eq('parseRoom leading digits', P.parseRoom('311 5ฟุต แอร์ ตู้เย็น  '), { room: '311', roomDesc: '311 5ฟุต แอร์ ตู้เย็น' });
eq('parseRoom without digits keeps text', P.parseRoom('ห้องรวม'), { room: 'ห้องรวม', roomDesc: 'ห้องรวม' });
eq('parseRoom blank → sentinel', P.parseRoom(''), { room: P.ROOM_UNKNOWN, roomDesc: '' });
eq('parseRoom null → sentinel', P.parseRoom(null).room, '(ไม่ระบุ)');
eq('distinctRooms: numeric first, text rooms last', A.distinctRooms([{ room: P.ROOM_UNKNOWN }, { room: '201' }, { room: '101' }, { room: 'ห้องรวม' }]),
  ['101', '201', P.ROOM_UNKNOWN, 'ห้องรวม']);
eq('toNumber', [P.toNumber('1,234.5'), P.toNumber(null), P.toNumber('x'), P.toNumber(7)], [1234.5, 0, 0, 7]);

section('parser: parseRows synthetic');
{
  const rows = [
    ['title'], [],
    ['วันที่', 'เลขที่เอกสาร', 'ห้อง', 'ชื่อลูกค้า', 'เลขมิเตอร์น้ำก่อนหน้า', 'เลขมิเตอร์น้ำล่าสุด', 'เลขมิเตอร์ไฟฟ้าก่อนหน้า', 'เลขมิเตอร์ไฟฟ้าล่าสุด',
      'ค่าเช่าห้อง (non-vat)', 'ค่าน้ำ (non-vat)', 'รวม', 'ภาษีมูลค่าเพิ่ม', 'รวมสุทธิ', 'สถานะ', 'หมายเหตุ'],
    ['01/01/2024', 'INV1', '101 แอร์', 'A', 1, 2, 3, 4, 2000, 0, 2000, 0, 2000, 'ชำระเงินแล้ว', null],
    [43528, 'INV2', '102', 'B', 1, 2, 3, 4, 2500, 30, 2500, 0, 2500, 'ค้างชำระ', 'x'],  // mismatch: 2530 vs 2500
    ['bad', 'INV3', '103', 'C', 1, 2, 3, 4, 1, 1, 2, 0, 2, 'ชำระเงินแล้ว', null],          // skipped
    [null, null, null, null, null, null, null, 'รวม', 4500],
    ['01/02/2024', 'FOOTER', '999', 'Z', 0, 0, 0, 0, 1, 1, 2, 0, 2, 'ชำระเงินแล้ว', null],
  ];
  const r = P.parseRows(rows, 'test');
  eq('headerRowIndex detected', r.headerRowIndex, 2);
  eq('stops at first empty date (footer ignored)', r.invoices.length, 2);
  eq('itemHeaders', r.itemHeaders, ['ค่าเช่าห้อง (non-vat)', 'ค่าน้ำ (non-vat)']);
  eq('only non-zero items kept', r.invoices[0].items, { 'ค่าเช่าห้อง (non-vat)': 2000 });
  eq('serial date row parsed', r.invoices[1].date, '2019-03-04');
  eq('mismatchRows', r.mismatchRows.map((m) => [m.docNo, m.diff]), [['INV2', 30]]);
  eq('skippedRows', r.skippedRows.map((m) => m.row), [6]);
  eq('columnStats', r.columnStats, { 'ค่าเช่าห้อง (non-vat)': { count: 2, total: 4500 }, 'ค่าน้ำ (non-vat)': { count: 1, total: 30 } });
  eq('statusCounts', r.statusCounts, { 'ชำระเงินแล้ว': 1, 'ค้างชำระ': 1 });
  let threw = null;
  try { P.parseRows([['x'], ['y']], 't'); } catch (e) { threw = e.message; }
  check('throws when header row missing', /ไม่พบแถวหัวตาราง/.test(threw || ''));
}

// ---------------------------------------------------------------------------
// parser.js — real workbook
// ---------------------------------------------------------------------------
section('parser: real workbook');
if (!fs.existsSync(XLSX_FILE)) {
  check(`real xlsx present at ${XLSX_FILE}`, false, 'file missing — cannot run data assertions');
} else {
  const buf = fs.readFileSync(XLSX_FILE);
  const ab = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
  const t0 = Date.now();
  const rep = P.parseWorkbook(ab);
  console.log(`  (parsed in ${Date.now() - t0} ms; warnings: ${JSON.stringify(rep.warnings)})`);

  eq('sheetName', rep.sheetName, 'รายงานใบแจ้งหนี้');
  eq('headerRowIndex', rep.headerRowIndex, 4);
  eq('1783 invoices', rep.invoices.length, 1783);
  eq('83 item headers', rep.itemHeaders.length, 83);
  check('all item headers end with (non-vat)', rep.itemHeaders.every((h) => /\(non-vat\)$/i.test(h)));
  eq('31 rooms', rep.rooms.length, 31);
  eq('rooms list', rep.rooms, ['101', '102', '103', '104', '105', '106', '107', '109', '110',
    '201', '202', '203', '204', '205', '206', '207', '208', '209', '210', '211',
    '301', '302', '303', '304', '305', '306', '307', '308', '309', '310', '311']);
  eq('status counts', rep.statusCounts, { 'ชำระเงินแล้ว': 1290, 'แก้ไขบิล': 53, 'ค้างชำระ': 440 });
  eq('date range', rep.dateRange, { min: '2019-03-04', max: '2026-09-07' });
  check('every invoice has a numeric room', rep.invoices.every((i) => /^\d+$/.test(i.room)));
  check('every invoice has ym matching date', rep.invoices.every((i) => i.date.startsWith(i.ym)));
  const first = rep.invoices[0];
  eq('first invoice basics', [first.date, first.docNo, first.room, first.status, first.total, first.net],
    ['2019-03-04', 'INV201903000001', '311', 'ชำระเงินแล้ว', 3653, 3653]);
  eq('first invoice items', first.items, {
    'ค่าเช่าห้อง (non-vat)': 2000, 'ค่าน้ำ (non-vat)': 81, 'ค่าไฟฟ้า (non-vat)': 572,
    'ค่าเช่าเฟอร์นิเจอร์ (non-vat)': 1000, 'ค่าสาธารูปโภคส่วนกลาง (non-vat)': 500, 'ส่วนลดผู้เช่าใหม่ (non-vat)': -500,
  });
  const matchRatio = (rep.invoices.length - rep.mismatchRows.length) / rep.invoices.length;
  console.log(`  (item-sum == รวม for ${rep.invoices.length - rep.mismatchRows.length}/${rep.invoices.length} rows = ${(matchRatio * 100).toFixed(2)}%)`);
  check('item-sum == รวม for >= 99% rows', matchRatio >= 0.99, `${(matchRatio * 100).toFixed(2)}%`);
  near('columnStats ค่าเช่าห้อง total = 3,529,100', rep.columnStats['ค่าเช่าห้อง (non-vat)'].total, 3529100);
  eq('columnStats ค่าเช่าห้อง count', rep.columnStats['ค่าเช่าห้อง (non-vat)'].count, 1738);
  near('columnStats ส่วนลดพิเศษ total = -764,100', rep.columnStats['ส่วนลดพิเศษ (non-vat)'].total, -764100);
  eq('columnStats ส่วนลดพิเศษ count', rep.columnStats['ส่วนลดพิเศษ (non-vat)'].count, 1004);

  // -------------------------------------------------------------------------
  // default-rule classification of all 83 headers
  // -------------------------------------------------------------------------
  section('settings: default classification of every header');
  const expected = JSON.parse(fs.readFileSync(EXPECTED_FILE, 'utf8'));
  delete expected._comment;
  const expectedKeys = Object.keys(expected);
  eq('expected table has 83 entries', expectedKeys.length, 83);
  const missing = rep.itemHeaders.filter((h) => !(h in expected));
  eq('every file header has an expected entry', missing, []);
  const extra = expectedKeys.filter((k) => !rep.itemHeaders.includes(k));
  eq('no stale expected entries', extra, []);
  const defaults = S.defaults();
  const cls = S.classifyAll(rep.itemHeaders, defaults);
  const wrong = cls.filter((c) => (c.group ?? null) !== (expected[c.raw] ?? null))
    .map((c) => `${c.raw} → ${c.group} (expected ${expected[c.raw]})`);
  eq('classification matches expected table', wrong, []);
  const states = cls.reduce((m, c) => { m[c.state] = (m[c.state] || 0) + 1; return m; }, {});
  eq('states: 82 suggested + 1 unassigned, 0 confirmed', states, { suggested: 82, unassigned: 1 });
  eq('unassigned header', cls.filter((c) => c.state === 'unassigned').map((c) => c.raw),
    ['ส่วนต่างค่าเช่าห้อง 201 และ 203 (ย้ายวันที่ 9/8/2020) (non-vat)']);
  const distinctNormalized = new Set(cls.map((c) => c.normalized));
  eq('nikhahit variant collapses: 82 distinct normalized headers', distinctNormalized.size, 82);

  // Sign hints from columnSummary
  const summary = A.columnSummary(rep.itemHeaders, rep.columnStats, defaults);
  const hinted = summary.filter((c) => c.signHint).map((c) => c.normalized).sort();
  eq('sign hints on the known tricky headers', hinted, [
    'คืน ส่วนลดผู้เช่าใหม่',
    'ค่าเช่าห้อง 12 วัน (1-12 มิถุนายน 2566)',
    'ยกเลิก ส่วนลดผู้เช่าใหม่ (New Customer) เดือน 3/2019',
  ]);

  // -------------------------------------------------------------------------
  // aggregate.js against the real data
  // -------------------------------------------------------------------------
  section('aggregate: real data');
  const colMap = S.columnMap(rep.itemHeaders, defaults);
  const groups = defaults.groups;
  const all = rep.invoices;

  // The spec figures are single-column totals; check them through the aggregate
  // path with a colMap that maps only that column.
  const onlyRent = { 'ค่าเช่าห้อง (non-vat)': 'room_rent' };
  const rbg1 = A.roomByGroup(all, onlyRent, groups);
  near('room_rent (column ค่าเช่าห้อง only, all statuses/years) = 3,529,100', rbg1.totals.groups.room_rent, 3529100);
  const onlySpecial = { 'ส่วนลดพิเศษ (non-vat)': 'discount_special' };
  near('discount_special (column ส่วนลดพิเศษ only) = -764,100', A.roomByGroup(all, onlySpecial, groups).totals.groups.discount_special, -764100);

  // With the default rules several one-off columns also land in these groups.
  const rbg = A.roomByGroup(all, colMap, groups);
  const sumCols = (gid) => A.round2(rep.itemHeaders.filter((h) => colMap[h] === gid).reduce((s, h) => s + rep.columnStats[h].total, 0));
  // 3,529,100 + 800 + 2,000 + 2,000 + 66,000 - 1,400 (the five one-off "ค่าเช่าห้อง ..." columns)
  near('room_rent (default rules) = 3,598,500', rbg.totals.groups.room_rent, 3598500);
  near('room_rent (default rules) equals sum of its columns', rbg.totals.groups.room_rent, sumCols('room_rent'));
  near('discount_special (default rules) = -775,900', rbg.totals.groups.discount_special, -775900);
  near('discount_special (default rules) equals sum of its columns', rbg.totals.groups.discount_special, sumCols('discount_special'));
  near('water = sum of its columns', rbg.totals.groups.water, sumCols('water'));
  near('electric = sum of its columns', rbg.totals.groups.electric, sumCols('electric'));
  near('discount parent = sum of children', rbg.totals.groups.discount,
    A.round2(['discount_new', 'discount_special', 'discount_days', 'discount_fridge', 'discount_tv', 'discount_other']
      .reduce((s, g) => s + (rbg.totals.groups[g] || 0), 0)));
  eq('roomByGroup rows = 31 rooms', rbg.rows.length, 31);
  eq('roomByGroup row order numeric', rbg.rows.map((r) => r.room), rep.rooms);
  eq('roomByGroup total count', rbg.totals.count, 1783);
  near('roomByGroup totals.invoiceNet = sum of net', rbg.totals.invoiceNet, A.round2(all.reduce((s, i) => s + i.net, 0)));
  near('unassignedTotal = ส่วนต่าง column', rbg.totals.unassignedTotal, 141.94);
  near('unassigned lands in other', rbg.totals.groups.other, 141.94);
  near('revenue+discount+other = net', rbg.totals.net, A.round2(rbg.totals.revenue + rbg.totals.discount + rbg.totals.other));
  near('exclude excluded from net', rbg.totals.exclude, sumCols('exclude'));
  const allItems = A.round2(rep.itemHeaders.reduce((s, h) => s + rep.columnStats[h].total, 0));
  near('net + utility + exclude = all item totals', A.round2(rbg.totals.net + rbg.totals.utility + rbg.totals.exclude), allItems);
  eq('groupColumns order: revenue, discount(with children), other, utility, exclude',
    rbg.columns.map((c) => c.id),
    ['room_rent', 'furniture', 'common_utility', 'aircon', 'tv', 'fridge', 'discount', 'other', 'water', 'electric', 'exclude']);
  eq('groupColumns discount children', rbg.columns.find((c) => c.id === 'discount').children.length, 6);

  section('aggregate: filters');
  const paid = A.applyFilters(all, { statuses: ['ชำระเงินแล้ว'] });
  eq('status filter', paid.length, 1290);
  const defaultStatus = A.applyFilters(all, { statuses: defaults.statusInclude });
  eq('default statusInclude excludes แก้ไขบิล', defaultStatus.length, 1730);
  const y2024 = A.applyFilters(all, { year: 2024 });
  check('year filter', y2024.length > 0 && y2024.every((i) => i.year === 2024));
  const q1 = A.applyFilters(all, { year: '2024', monthFrom: 1, monthTo: 3 });
  check('month range filter', q1.length > 0 && q1.every((i) => i.year === 2024 && i.month <= 3));
  eq('year all', A.applyFilters(all, { year: 'all' }).length, 1783);
  const r311 = A.applyFilters(all, { rooms: ['311'] });
  check('room filter', r311.length > 0 && r311.every((i) => i.room === '311'));
  const cust = A.applyFilters(all, { customer: first.customer.slice(0, 6).toLowerCase() });
  check('customer substring filter (case-insensitive)', cust.length > 0 && cust.some((i) => i.docNo === first.docNo));
  eq('date range filter', A.applyFilters(all, { dateFrom: '2019-03-04', dateTo: '2019-03-04' }).length,
    all.filter((i) => i.date === '2019-03-04').length);
  eq('distinctYears', A.distinctYears(all), [2019, 2020, 2021, 2022, 2023, 2024, 2025, 2026]);
  eq('distinctRooms', A.distinctRooms(all), rep.rooms);
  eq('distinctStatuses sorted by count', A.distinctStatuses(all).map((s) => s.status), ['ชำระเงินแล้ว', 'ค้างชำระ', 'แก้ไขบิล']);

  section('aggregate: invoiceGroupAmounts');
  const a0 = A.invoiceGroupAmounts(first, colMap, groups);
  eq('first invoice groups', a0.groups, { room_rent: 2000, water: 81, electric: 572, furniture: 1000, common_utility: 500, discount_new: -500, discount: -500 });
  eq('first invoice kinds', [a0.revenue, a0.discount, a0.other, a0.utility, a0.exclude, a0.net, a0.unassignedTotal], [3500, -500, 0, 653, 0, 3000, 0]);
  const a1 = A.invoiceGroupAmounts(first, {}, groups);
  eq('empty colMap → everything unassigned → other', [a1.groups.other, a1.unassignedTotal, a1.other, a1.net], [3653, 3653, 3653, 3653]);
  eq('valueOf selectors', [
    A.valueOf(first, colMap, groups, A.VALUE_KEYS.REVENUE),
    A.valueOf(first, colMap, groups, A.VALUE_KEYS.NET),
    A.valueOf(first, colMap, groups, A.VALUE_KEYS.INVOICE_NET),
    A.valueOf(first, colMap, groups, A.VALUE_KEYS.UTILITY),
    A.valueOf(first, colMap, groups, A.VALUE_KEYS.COUNT),
    A.valueOf(first, colMap, groups, 'room_rent'),
    A.valueOf(first, colMap, groups, 'discount'),
    A.valueOf(first, colMap, groups, 'tv'),
  ], [3500, 3000, 3653, 653, 1, 2000, -500, 0]);
  eq('valueOptions starts with the 5 built-ins then groups', A.valueOptions(groups).slice(0, 5).map((o) => o.key),
    ['__revenue', '__net', '__invoiceNet', '__utility', '__count']);
  eq('valueOptions length', A.valueOptions(groups).length, 5 + 17);

  section('aggregate: metrics');
  const m = A.metrics(all, colMap, groups);
  eq('metrics count', m.count, 1783);
  near('metrics revenue = roomByGroup revenue', m.revenue, rbg.totals.revenue);
  near('metrics net', m.net, rbg.totals.net);
  near('metrics utility = water + electric', m.utility, A.round2(rbg.totals.groups.water + rbg.totals.groups.electric));
  near('metrics unpaid = sum net of ค้างชำระ', m.unpaid, A.round2(all.filter((i) => i.status === 'ค้างชำระ').reduce((s, i) => s + i.net, 0)));
  eq('metrics unpaidCount', m.unpaidCount, 440);
  near('metrics discount negative', Math.sign(m.discount), -1, 0);

  section('aggregate: pivotMonthly');
  {
    // Several value keys are summed per invoice (discounts are stored negative).
    const V = A.VALUE_KEYS;
    const sample = all.find((i) => { const a = A.invoiceGroupAmounts(i, colMap, groups); return a.revenue > 0 && a.discount < 0; });
    const am = A.invoiceGroupAmounts(sample, colMap, groups);
    near('revenue + discount = revenue minus discount amount', A.valueOfMany(sample, colMap, groups, [V.REVENUE, 'discount']), am.revenue + am.discount);
    near('revenue + discount = NET when no other', A.valueOfMany(sample, colMap, groups, [V.REVENUE, 'discount']), am.revenue + am.discount);
    near('overlap counted once: revenue + room_rent = revenue', A.valueOfMany(sample, colMap, groups, [V.REVENUE, 'room_rent']), am.revenue);
    near('parent + child counted once', A.valueOfMany(sample, colMap, groups, ['discount', 'discount_special']), am.discount);
    near('revenue + utility', A.valueOfMany(sample, colMap, groups, [V.REVENUE, V.UTILITY]), am.revenue + am.utility);
    eq('exclusive key stands alone', A.valueOfMany(sample, colMap, groups, [V.REVENUE, V.COUNT]), 1);
    near('single key = valueOf', A.valueOfMany(sample, colMap, groups, ['room_rent']), A.valueOf(sample, colMap, groups, 'room_rent'));
    const multi = A.pivotMonthly(y2024, colMap, { year: 2024, valueKeys: [V.REVENUE, 'discount'], groups });
    const net2024 = A.pivotMonthly(y2024, colMap, { year: 2024, valueKey: V.NET, groups });
    const rev2024 = A.pivotMonthly(y2024, colMap, { year: 2024, valueKey: V.REVENUE, groups });
    const disc2024 = A.pivotMonthly(y2024, colMap, { year: 2024, valueKey: 'discount', groups });
    near('pivot revenue+discount grand total = revenue + discount', multi.grandTotal.total, rev2024.grandTotal.total + disc2024.grandTotal.total);
    eq('pivot label joins the selection', multi.valueLabel, 'รายได้หลัก + ส่วนลด');
    check('pivot valueKey alias still works', net2024.valueKeys.length === 1 && net2024.valueKey === V.NET);
  }
  const pv = A.pivotMonthly(y2024, colMap, { year: 2024, valueKey: 'room_rent', groups });
  eq('12 month columns', pv.columns.map((c) => c.label), A.MONTHS_TH);
  eq('column keys', pv.columns[0].key, '2024-01');
  eq('rows = rooms with invoices in 2024', pv.rows.length, A.distinctRooms(y2024).length);
  check('every row has 12 values', pv.rows.every((r) => r.values.length === 12));
  near('grand total = sum of room_rent in 2024', pv.grandTotal.total,
    A.round2(y2024.reduce((s, i) => s + A.invoiceGroupAmounts(i, colMap, groups).groups.room_rent || 0, 0)));
  near('grand total = sum of column values', pv.grandTotal.total, A.round2(pv.grandTotal.values.reduce((s, x) => s + x, 0)));
  near('row totals sum to grand total', A.round2(pv.rows.reduce((s, r) => s + r.total, 0)), pv.grandTotal.total);
  eq('valueLabel', pv.valueLabel, 'ค่าเช่าห้อง');
  const pvc = A.pivotMonthly(y2024, colMap, { year: 2024, valueKey: A.VALUE_KEYS.COUNT, byCustomer: true, groups });
  eq('count pivot grand total = invoices', pvc.grandTotal.total, y2024.length);
  check('byCustomer adds level-1 rows', pvc.rows.some((r) => r.level === 1 && r.customer));
  const roomRow = pvc.rows.find((r) => r.level === 0);
  const custRows = pvc.rows.filter((r) => r.level === 1 && r.room === roomRow.room);
  near('room subtotal = sum of its customer rows', roomRow.total, custRows.reduce((s, r) => s + r.total, 0));
  {
    // Customers within a room are ordered by their first invoice date (oldest first), using all invoices.
    const pvAllCust = A.pivotMonthly(all, colMap, { year: null, valueKey: A.VALUE_KEYS.COUNT, byCustomer: true, groups });
    const first = {};
    all.forEach((i) => { const k = i.room + '|' + i.customer; if (!(k in first) || i.date < first[k]) first[k] = i.date; });
    let ordered = true;
    const byRoom = {};
    pvAllCust.rows.filter((r) => r.level === 1).forEach((r) => { (byRoom[r.room] = byRoom[r.room] || []).push(first[r.room + '|' + r.customer]); });
    for (const dates of Object.values(byRoom)) for (let i = 1; i < dates.length; i++) if (dates[i] < dates[i - 1]) ordered = false;
    check('pivot customers per room ordered by first invoice date', ordered && Object.values(byRoom).some((d) => d.length > 1));
    // A filtered subset keeps the order that all invoices define.
    const r103 = all.filter((i) => i.room === '103');
    const sub = A.pivotMonthly(r103.filter((i) => i.year === 2026), colMap, { year: 2026, valueKey: A.VALUE_KEYS.COUNT, byCustomer: true, groups, allInvoices: all });
    const subDates = sub.rows.filter((r) => r.level === 1).map((r) => first['103|' + r.customer]);
    check('order uses allInvoices when filtered', subDates.every((d, i) => i === 0 || d >= subDates[i - 1]));
  }
  const pvAll = A.pivotMonthly(all, colMap, { year: null, valueKey: A.VALUE_KEYS.INVOICE_NET, groups });
  check('year=all → first column is the first month', pvAll.columns.length > 12 && pvAll.columns[0].key === '2019-03' && /2019$/.test(pvAll.columns[0].label));
  eq('year=all → every calendar month from first to last (91 columns 2019-03..2026-09)', pvAll.columns.length, 91);
  check('year=all → vacant month 2021-02 is present with 0', (() => {
    const i = pvAll.columns.findIndex((c) => c.key === '2021-02');
    return i > 0 && pvAll.grandTotal.values[i] === 0 && pvAll.columns[i - 1].key === '2021-01' && pvAll.columns[i + 1].key === '2021-03';
  })());
  near('year=all invoiceNet grand total = sum net', pvAll.grandTotal.total, A.round2(all.reduce((s, i) => s + i.net, 0)));
  const pv311 = A.pivotMonthly(r311, colMap, { year: null, valueKey: A.VALUE_KEYS.COUNT, groups, range: rep.dateRange });
  eq('year=all with report range → room filter keeps the full axis', pv311.columns.length, 91);
  eq('monthColumns: single-month data + no range → 1 column', A.monthColumns([{ ym: '2024-05' }], null).length, 1);
  eq('monthColumns: no data, no range → []', A.monthColumns([], 'all'), []);

  section('aggregate: monthlySeries');
  const ms = A.monthlySeries(y2024, colMap, groups, { year: 2024 });
  eq('12 labels', ms.labels, A.MONTHS_TH);
  eq('series = assignable revenue/discount/other groups', ms.series.map((s) => s.id),
    ['room_rent', 'furniture', 'common_utility', 'aircon', 'tv', 'fridge', 'discount_new', 'discount_special', 'discount_days', 'discount_fridge', 'discount_tv', 'discount_other', 'other']);
  check('series data length 12', ms.series.every((s) => s.data.length === 12));
  near('net = sum of series per month (Jan)', ms.net[0], A.round2(ms.series.reduce((s, x) => s + x.data[0], 0)));
  eq('count per month sums to year count', ms.count.reduce((s, x) => s + x, 0), y2024.length);
  const msAll = A.monthlySeries(all, colMap, groups, {});
  check('year=all labels include year', msAll.keys.length > 12 && msAll.labels[0] === 'มี.ค. 2019');
  eq('year=all series cover every month (gaps filled)', msAll.keys.length, 91);
  check('columnSummary carries the group label for sorting', summary.every((c) => typeof c.groupLabel === 'string') &&
    summary.find((c) => c.group === 'room_rent').groupLabel === 'ค่าเช่าห้อง');

  section('aggregate: invoiceRows');
  const ir = A.invoiceRows(all.slice(0, 3), colMap, groups);
  eq('invoiceRows shape', Object.keys(ir[0]).sort(), ['customer', 'date', 'discount', 'docNo', 'exclude', 'groups', 'net', 'netRevenue', 'other', 'revenue', 'room', 'roomDesc', 'status', 'total', 'unassignedTotal', 'utility', 'vat', 'ym']);
}

// ---------------------------------------------------------------------------
console.log(`\n${passed} passed, ${failed} failed`);
if (failed) {
  console.log('\nFailures:');
  for (const f of failures) console.log('  - ' + f);
  process.exit(1);
}
