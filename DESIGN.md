# JMP Revenue Report — Design Spec

Static web app (GitHub Pages) that reads the invoice export
`report-01_01_2017-30_09_2026.xlsx` (and future exports with the same layout),
lets the user group the many overlapping item columns into revenue groups,
saves that grouping as editable settings, and shows per-room revenue summaries.

All processing happens in the browser. No backend. The xlsx is never uploaded
anywhere and is NOT committed to the repo (user picks the file every time).

## Decisions (confirmed by owner, 2026-09-30)

1. ค่าน้ำ / ค่าไฟฟ้า → separate groups of kind `utility`, shown apart from main revenue.
2. Status filter default: `ชำระเงินแล้ว` + `ค้างชำระ` (+ `ยืนยันชำระเงินโดยพนักงาน` if present). `แก้ไขบิล` excluded by default. Toggle in UI.
3. Group `ไม่นับ` (kind `exclude`) exists for non-revenue columns (prepayments, bad debt, etc.).
4. File is chosen via file picker every load (privacy). Settings persist in `localStorage` + JSON export/import.
5. Room = leading number only (`"311 5ฟุต แอร์ ..."` → `"311"`); keep full text as `roomDesc`.

## Input file layout

- First worksheet (`รายงานใบแจ้งหนี้`). Rows 1–4 are title lines; header row is the
  first row whose column A equals `วันที่` (row 5 in current file — detect, don't hardcode).
- Fixed columns (by header text, not index):
  `วันที่`, `เลขที่เอกสาร`, `ห้อง`, `ชื่อลูกค้า`,
  `เลขมิเตอร์น้ำก่อนหน้า`, `เลขมิเตอร์น้ำล่าสุด`, `เลขมิเตอร์ไฟฟ้าก่อนหน้า`, `เลขมิเตอร์ไฟฟ้าล่าสุด`,
  then N item columns, then `รวม`, `ภาษีมูลค่าเพิ่ม`, `รวมสุทธิ`, `สถานะ`, `หมายเหตุ`.
- Item columns = every header between `เลขมิเตอร์ไฟฟ้าล่าสุด` and `รวม`. Current file: 83 of them,
  all ending in ` (non-vat)`. Many are one-off (used in 1 invoice). Examples:
  `ค่าเช่าห้อง (non-vat)`, `ค่าสาธารูปโภคส่วนกลาง (non-vat)` (typo) and
  `ค่าสาธารณูปโภคส่วนกลาง (non-vat)`, `ส่วนลดพิเศษ ครั้งที่ 3 (non-vat)`,
  `หัก ค่าเช่า 17 วัน (1-17 มิย. 65) (non-vat)`, `ชำระล่วงหน้า วันที่ 25/3/2019 (non-vat)`,
  `ลูกค้าย้ายออกแต่ไม่ชำระค่าเช่า (คิดเป็นหนี้สูญ) (non-vat)`.
- Data rows: 1,783 invoices, dates `dd/mm/yyyy` as text (CE years, e.g. `04/03/2019`),
  range 2019-03-04 → 2026-09-07. Row ends when column A (date) is empty.
  Below data there are footer rows (COUNTIF formulas) — must be ignored.
- Status values: `ชำระเงินแล้ว` (1290), `ค้างชำระ` (440), `แก้ไขบิล` (53); footer also references
  `ยืนยันชำระเงินโดยพนักงาน` — treat as paid-like.
- Sanity check: for each row, sum(item columns) should equal `รวม`. Report row count of mismatches
  in the load summary (do not block).
- Read with SheetJS `XLSX.read(arrayBuffer, {type:'array', cellDates:false})` and
  `sheet_to_json(ws, {header:1, raw:true, defval:null})`. Dates may still arrive as JS Date or
  Excel serial in other exports — handle string `dd/mm/yyyy`, Date, and number.
- 31 rooms: 101–107, 109, 110, 201–211, 301–311.

## Data model (in-memory)

```js
Invoice = {
  date: 'YYYY-MM-DD', year: 2019, month: 3, ym: '2019-03',
  docNo, room: '311', roomDesc, customer, status,
  items: { [rawHeader]: number },   // only non-zero items
  total, vat, net,                  // from รวม / ภาษีมูลค่าเพิ่ม / รวมสุทธิ
  itemSum,                          // computed sum of items
}
```

## Grouping

### Settings shape (persisted, JSON export/import)

```js
Settings = {
  version: 1,
  groups: [ { id, label, kind, parent?, order } ],
  //   kind: 'revenue' | 'discount' | 'utility' | 'other' | 'exclude'
  rules:  [ { pattern, flags?, group } ],           // regex, ordered, first match wins
  columns: { [normalizedHeader]: groupId },          // explicit, confirmed by user; wins over rules
  statusInclude: ['ชำระเงินแล้ว','ค้างชำระ','ยืนยันชำระเงินโดยพนักงาน'],
}
```

### Default groups

| id | label | kind | parent |
|---|---|---|---|
| room_rent | ค่าเช่าห้อง | revenue | |
| furniture | ค่าเช่าเฟอร์นิเจอร์ | revenue | |
| common_utility | ค่าสาธารณูปโภคส่วนกลาง | revenue | |
| aircon | ค่าเช่าเครื่องปรับอากาศ | revenue | |
| tv | ค่าเช่าโทรทัศน์ | revenue | |
| fridge | ค่าเช่าตู้เย็น | revenue | |
| discount | ส่วนลด | discount | (parent, not assignable directly) |
| discount_new | ส่วนลดผู้เช่าใหม่ | discount | discount |
| discount_special | ส่วนลดพิเศษ | discount | discount |
| discount_days | ส่วนลดตามจำนวนวัน (เข้าพักไม่ตรงต้นเดือน) | discount | discount |
| discount_fridge | ส่วนลดค่าเช่าตู้เย็น | discount | discount |
| discount_tv | ส่วนลดค่าเช่าโทรทัศน์ | discount | discount |
| discount_other | ส่วนลดอื่นๆ | discount | discount |
| other | อื่นๆ | other | (receives unassigned columns; cannot be deleted or used as a parent) |
| water | ค่าน้ำ | utility | |
| electric | ค่าไฟฟ้า | utility | |
| exclude | ไม่นับ | exclude | |

Users can add / rename / delete groups (any kind, optional parent). Deleting a group
that has columns mapped moves those columns back to "unassigned". Nesting is one level
deep: a parent must be top-level, a group with children cannot get a parent, and
columns / rules may only target groups without children. When a group gains its first
child, its confirmed columns and rules move to that child. Import validation enforces
the same invariants (no cycles, no parent targets) and adds the `other` group if missing.

### Header normalization (key for `settings.columns` and for rules)

1. Unicode NFC.
2. Replace `ํา` (nikhahit + sara aa) with `ำ` (sara am) — the file contains both spellings
   (`สำหรับ` vs `สําหรับ`).
3. Strip trailing `(non-vat)` (case-insensitive, optional surrounding spaces).
4. Collapse whitespace, trim.

Rules and the mapping table operate on the normalized header. The raw header stays as key in `invoice.items`.

### Default rules (ordered; first match wins). Regex on normalized header, flags `i`

```
1  ^ค่าน้ำ|water rate                                     → water
2  ^ค่าไฟ|electrical rate                                 → electric
3  ส่วนลด.*ตู้เย็น                                        → discount_fridge
4  ส่วนลด.*(โทรทัศน์|ทีวี)                                 → discount_tv
5  (หัก|ส่วนลด).*(\d+\s*วัน|ครึ่งเดือน|\d+\s*-\s*\d+)       → discount_days
6  ส่วนลด(ผู้เช่า|ลูกค้า)ใหม่|new\s*cu?s?t?omer            → discount_new
7  ส่วนลดพิเศษ                                            → discount_special
8  ชำระล่วงหน้า|จ่ายล่วงหน้า|หนี้สูญ|ค้างชำระ               → exclude
9  ^ส่วนลด|^หัก                                            → discount_other
10 ^ค่าเช่าห้อง|room\s*rate                                → room_rent
11 เฟอร์นิเจอร์|furniture                                  → furniture
12 สาธาร.*ส่วนกลาง                                         → common_utility
13 เครื่องปรับอากาศ                                         → aircon
14 โทรทัศน์|\btv\b                                          → tv
15 ตู้เย็น|refrigerator                                     → fridge
```

Rule 8 (exclude) sits before the generic `^หัก` rule so that
`หัก ค่าเช่าห้อง (จ่ายล่วงหน้า)` is excluded like the other prepayment headers
(`ชำระล่วงหน้า วันที่ …`), per decision 3.

Rule matches are shown as **suggested** (badge "เดาให้"); explicit `settings.columns`
entries are **confirmed**. Headers matching no rule are **unassigned** and highlighted.
Known tricky headers (rules above intentionally leave some for the user):

- `ค่าเช่าห้อง 12 วัน (1-12 มิถุนายน 2566)` total is −1,400 → rule 10 says room_rent but it is a
  proration discount. Show a hint when a header's total sign disagrees with its group kind
  (revenue group but negative total, or discount group but positive total).
- `ยกเลิก ส่วนลดผู้เช่าใหม่ ...` (+500) and `คืน ส่วนลดผู้เช่าใหม่` (+1500) are discount reversals;
  discount_new is correct (positive offsets). Hint shown but fine.
- `ส่วนต่างค่าเช่าห้อง 201 และ 203 (ย้ายวันที่ 9/8/2020)` → unassigned; user decides.
- `ค่าเช่าค้างชำระ` (+276) → rule 8 → exclude (user may override).

### Mapping workflow (screen 2)

After a file loads the app stays on screen 1 showing the load summary (row range, status counts,
parser warnings, mismatch rows) with buttons to continue. Screen 2: table of every item column with normalized name, invoice count, total,
current group (select), state badge (confirmed / suggested / unassigned). Sort: unassigned first,
then suggested, then confirmed; secondary by |total| desc. Search box. Bulk: "ยืนยันทั้งหมดที่เดาให้".
"ไปที่สรุป" button enabled always; unassigned columns are treated as `other` in the dashboard with
a warning count. Confirming writes `settings.columns[normalized] = groupId` and persists.
Next file load: only columns without a confirmed mapping are surfaced first.

### Settings screen

- Groups editor (add / rename / kind / parent / reorder / delete).
- Rules editor (add / edit pattern & target / reorder / delete / test against loaded headers).
- Columns mapping table (same as screen 2 but all columns, plus "ล้างการยืนยัน" per row).
- Export settings JSON (download `jmp-settings.json`), Import JSON (file picker, validate shape),
  Reset to defaults (confirm).
- `localStorage` key: `jmp-report.settings.v1`.

## Dashboard (screen 3)

Global filters (top bar): year (all / single), month range within year, rooms (multi, default all),
status set (checkboxes, default per decision 2; session-only, with a "จำเป็นค่าเริ่มต้น" button
that writes the checked set into `settings.statusInclude`), search customer.

Metrics row:
- รายได้หลัก = sum of kind `revenue`
- ส่วนลด = sum of kind `discount` (negative)
- รายได้สุทธิ = revenue + discount + other
- สาธารณูปโภค (น้ำ+ไฟ) = kind `utility`
- ค้างชำระ = sum of `net` for invoices with status `ค้างชำระ` within filter
- จำนวนบิล

Views (tabs, all respect global filters):

1. **ห้อง × กลุ่ม** — rows = room, columns = every assignable group (revenue groups, then ส่วนลด
   collapsed into one column expandable to subgroups, then อื่นๆ, then น้ำ / ไฟ, then สุทธิ).
   Totals row. Column header click = sort.
2. **Pivot รายเดือน** (matches the owner's existing Excel pivot) — rows = room, optional
   sub-rows = customer; columns = months (ม.ค. … ธ.ค.) of selected year + Grand Total; value =
   selector: any single group | รายได้หลัก | รายได้สุทธิ | รวมสุทธิ(บิล) | count.
   Empty cells show `0.00`-style formatted zero like the Excel pivot. Grand total row.
   When year = all, columns = every calendar month of the file's date range (vacant months = 0),
   regardless of the room / status / customer filter.
3. **แนวโน้ม** — Chart.js stacked bar per month (groups as series, discounts negative) with net line;
   when year = all, x-axis = every month of the file's date range (same rule as the pivot).
4. **รายละเอียดบิล** — invoice table (date, docNo, room, customer, status, per-group amounts, net).
   Click a room cell anywhere → this tab filtered to that room.

Export button on each table view → `.xlsx` via SheetJS (`XLSX.utils.aoa_to_sheet`), filename
`jmp-<view>-<filter>.xlsx`.

Number format: `toLocaleString('th-TH', {minimumFractionDigits:2, maximumFractionDigits:2})`.
Negative numbers shown in red. Thai month short names: ม.ค. ก.พ. มี.ค. เม.ย. พ.ค. มิ.ย. ก.ค. ส.ค. ก.ย. ต.ค. พ.ย. ธ.ค.

## Tech constraints

- Plain HTML/CSS/JS, **no ES modules, no build step** (must work when opened via `file://` and on
  GitHub Pages). Multiple `<script>` files sharing a `window.JMP` namespace is fine.
- Libraries via CDN only: SheetJS `xlsx` 0.18.5 and Chart.js 4 from `cdnjs.cloudflare.com`.
- UI language: Thai. Font stack: `"Segoe UI", Tahoma, sans-serif` (system Thai fonts OK).
- Responsive down to ~900px; tables scroll horizontally inside their container.
- Light theme only is acceptable; keep colors in CSS variables.
- No frameworks. Keep total JS ≲ 2,500 lines, readable, commented where non-obvious.

## File layout

```
index.html
css/style.css
js/settings.js     defaults, normalizeHeader, load/save/import/export, rule engine (classify)
js/parser.js       xlsx → Invoice[]  (+ column stats, sanity check)
js/aggregate.js    pure functions: filter, group sums, room×group, pivot, monthly series
js/ui.js           screens, tables, chart, export
js/app.js          bootstrap + state
test/               node tests for settings.js / parser.js / aggregate.js against the real xlsx
README.md          usage + GitHub Pages deploy steps (Thai)
.gitignore         *.xlsx, node_modules
```

Node tests: `test/run.mjs` loads the plain scripts with `vm` into a fake `window`, installs
`xlsx` from npm as a devDependency (`package.json`), reads the real xlsx from the repo root, and
asserts: 1783 invoices, 31 rooms, status counts (1290/440/53), item-sum == รวม for ≥ 99% rows,
default-rule classification of every one of the 83 headers matches an expected table, and
aggregate totals for `room_rent` equal 3,529,100 (paid+unpaid+edited, all years) and
`ส่วนลดพิเศษ` column total −764,100.
