window.JMP = window.JMP || {};

/*
 * JMP.parser
 * ----------
 * Reads the invoice export workbook (SheetJS `XLSX` global) into Invoice[]
 * plus per-column statistics and the item-sum sanity check.
 *
 * Layout (see DESIGN.md): first worksheet, header row = first row whose column A
 * is "วันที่", fixed columns located by header text, item columns = everything
 * between "เลขมิเตอร์ไฟฟ้าล่าสุด" and "รวม". Data ends at the first row with an
 * empty date; footer rows below are ignored.
 */
(function (global) {
  'use strict';

  var FIXED_COLUMNS = {
    date: 'วันที่',
    docNo: 'เลขที่เอกสาร',
    room: 'ห้อง',
    customer: 'ชื่อลูกค้า',
    waterPrev: 'เลขมิเตอร์น้ำก่อนหน้า',
    waterLast: 'เลขมิเตอร์น้ำล่าสุด',
    elecPrev: 'เลขมิเตอร์ไฟฟ้าก่อนหน้า',
    elecLast: 'เลขมิเตอร์ไฟฟ้าล่าสุด',
    total: 'รวม',
    vat: 'ภาษีมูลค่าเพิ่ม',
    net: 'รวมสุทธิ',
    status: 'สถานะ',
    note: 'หมายเหตุ'
  };

  var EPSILON = 0.005; // item-sum vs รวม tolerance (2 decimals)

  function round2(n) {
    return Math.round(n * 100) / 100;
  }

  function pad2(n) {
    return (n < 10 ? '0' : '') + n;
  }

  function cellText(v) {
    if (v === null || v === undefined) return '';
    return String(v).replace(/\s+/g, ' ').trim();
  }

  function isEmpty(v) {
    return v === null || v === undefined || cellText(v) === '';
  }

  /** Coerce a cell to a number ("1,234.50" and blanks accepted). Non-numeric → 0. */
  function toNumber(v) {
    if (typeof v === 'number') return isFinite(v) ? v : 0;
    if (v === null || v === undefined || v === '') return 0;
    var n = parseFloat(String(v).replace(/,/g, '').trim());
    return isFinite(n) ? n : 0;
  }

  function makeDate(y, m, d) {
    if (!(y >= 1900 && y <= 2200 && m >= 1 && m <= 12 && d >= 1 && d <= 31)) return null;
    return { date: y + '-' + pad2(m) + '-' + pad2(d), year: y, month: m, ym: y + '-' + pad2(m) };
  }

  /**
   * Parse a date cell. Accepts "dd/mm/yyyy" (also d/m/yyyy, dd-mm-yyyy, yyyy-mm-dd),
   * a JS Date, or an Excel serial number. Buddhist-era years (>= 2400) are
   * converted to CE. Returns {date, year, month, ym} or null.
   */
  function parseDate(value) {
    if (value === null || value === undefined || value === '') return null;
    var y, m, d;
    if (Object.prototype.toString.call(value) === '[object Date]') { // realm-safe instanceof Date
      if (isNaN(value.getTime())) return null;
      y = value.getFullYear(); m = value.getMonth() + 1; d = value.getDate();
    } else if (typeof value === 'number') {
      // Excel serial (1900 date system): day 0 = 1899-12-30
      if (!isFinite(value) || value < 1) return null;
      var ms = Math.round((value - 25569) * 86400 * 1000);
      var dt = new Date(ms);
      y = dt.getUTCFullYear(); m = dt.getUTCMonth() + 1; d = dt.getUTCDate();
    } else {
      var s = String(value).trim();
      var mt = s.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{4})(?:\s.*)?$/);
      if (mt) {
        d = +mt[1]; m = +mt[2]; y = +mt[3];
      } else {
        mt = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T\s].*)?$/);
        if (!mt) return null;
        y = +mt[1]; m = +mt[2]; d = +mt[3];
      }
    }
    if (y >= 2400) y -= 543; // Buddhist era → CE
    return makeDate(y, m, d);
  }

  var ROOM_UNKNOWN = '(ไม่ระบุ)'; // room key for invoices whose "ห้อง" cell is blank

  /** "311 5ฟุต แอร์ ..." → { room: "311", roomDesc: "311 5ฟุต แอร์ ..." }; blank cell → ROOM_UNKNOWN. */
  function parseRoom(raw) {
    var desc = cellText(raw);
    var mt = desc.match(/^\d+/);
    return { room: mt ? mt[0] : (desc || ROOM_UNKNOWN), roomDesc: desc };
  }

  function findHeaderRow(rows) {
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      if (r && cellText(r[0]) === FIXED_COLUMNS.date) return i;
    }
    return -1;
  }

  /**
   * Parse an array-of-arrays (as returned by sheet_to_json header:1) into the
   * result structure. Exposed separately from parseWorkbook for testing.
   */
  function parseRows(rows, sheetName) {
    var warnings = [];
    var headerRowIndex = findHeaderRow(rows);
    if (headerRowIndex < 0) {
      throw new Error('ไม่พบแถวหัวตาราง (คอลัมน์ A = "' + FIXED_COLUMNS.date + '")');
    }
    var header = rows[headerRowIndex].map(cellText);

    var col = {};
    Object.keys(FIXED_COLUMNS).forEach(function (key) {
      col[key] = header.indexOf(FIXED_COLUMNS[key]);
    });
    ['date', 'docNo', 'room', 'customer', 'elecLast', 'total', 'net', 'status'].forEach(function (key) {
      if (col[key] < 0) throw new Error('ไม่พบคอลัมน์ "' + FIXED_COLUMNS[key] + '" ในแถวหัวตาราง');
    });
    if (col.vat < 0) warnings.push('ไม่พบคอลัมน์ "' + FIXED_COLUMNS.vat + '" — ใช้ค่า 0');
    if (col.total <= col.elecLast) throw new Error('ลำดับคอลัมน์ไม่ถูกต้อง: "รวม" ต้องอยู่หลัง "เลขมิเตอร์ไฟฟ้าล่าสุด"');

    // Item columns: everything strictly between เลขมิเตอร์ไฟฟ้าล่าสุด and รวม (skip blank headers).
    var itemCols = [];
    var itemHeaders = [];
    var seen = {};
    for (var c = col.elecLast + 1; c < col.total; c++) {
      var h = header[c];
      if (h === '') { warnings.push('คอลัมน์ที่ ' + (c + 1) + ' ไม่มีชื่อหัวคอลัมน์ — ข้าม'); continue; }
      if (seen[h]) {
        warnings.push('หัวคอลัมน์ซ้ำ: "' + h + '" (คอลัมน์ที่ ' + (c + 1) + ') — ยอดจะถูกรวมกัน');
      } else {
        seen[h] = true;
        itemHeaders.push(h);
      }
      itemCols.push(c);
    }
    if (itemHeaders.length === 0) warnings.push('ไม่พบคอลัมน์รายการสินค้า');

    var columnStats = {};
    itemHeaders.forEach(function (h) { columnStats[h] = { count: 0, total: 0 }; });

    var invoices = [];
    var mismatchRows = [];
    var skippedRows = [];
    var beYearWarned = false;
    var noRoomWarned = 0;
    var lastRowIndex = headerRowIndex;

    for (var i = headerRowIndex + 1; i < rows.length; i++) {
      var r = rows[i] || [];
      var rawDate = r[col.date];
      if (isEmpty(rawDate)) break; // end of data; footer rows follow
      lastRowIndex = i;
      var excelRow = i + 1;

      var dt = parseDate(rawDate);
      if (!dt) {
        skippedRows.push({ row: excelRow, reason: 'วันที่อ่านไม่ได้: ' + cellText(rawDate) });
        continue;
      }
      if (typeof rawDate === 'string' && /\d{4}/.test(rawDate) && +rawDate.match(/(\d{4})/)[1] >= 2400 && !beYearWarned) {
        warnings.push('พบปี พ.ศ. ในคอลัมน์วันที่ — แปลงเป็น ค.ศ. ให้อัตโนมัติ');
        beYearWarned = true;
      }

      var roomInfo = parseRoom(r[col.room]);
      if (!/^\d+$/.test(roomInfo.room)) noRoomWarned++;

      var items = {};
      var itemSum = 0;
      for (var k = 0; k < itemCols.length; k++) {
        var h2 = header[itemCols[k]];
        var v = toNumber(r[itemCols[k]]);
        if (v === 0) continue;
        items[h2] = round2((items[h2] || 0) + v);
        itemSum += v;
        var st = columnStats[h2];
        st.count += 1;
        st.total += v;
      }
      itemSum = round2(itemSum);

      var total = round2(toNumber(r[col.total]));
      var vat = col.vat >= 0 ? round2(toNumber(r[col.vat])) : 0;
      var net = round2(toNumber(r[col.net]));

      var inv = {
        row: excelRow,
        date: dt.date,
        year: dt.year,
        month: dt.month,
        ym: dt.ym,
        docNo: cellText(r[col.docNo]),
        room: roomInfo.room,
        roomDesc: roomInfo.roomDesc,
        customer: cellText(r[col.customer]),
        status: cellText(r[col.status]),
        note: col.note >= 0 ? cellText(r[col.note]) : '',
        items: items,
        itemSum: itemSum,
        total: total,
        vat: vat,
        net: net
      };
      invoices.push(inv);

      if (Math.abs(itemSum - total) > EPSILON) {
        mismatchRows.push({ row: excelRow, docNo: inv.docNo, date: inv.date, room: inv.room,
                            itemSum: itemSum, total: total, diff: round2(itemSum - total) });
      }
    }

    itemHeaders.forEach(function (h) { columnStats[h].total = round2(columnStats[h].total); });

    if (skippedRows.length) warnings.push('ข้ามแถวที่อ่านวันที่ไม่ได้ ' + skippedRows.length + ' แถว');
    if (noRoomWarned) warnings.push('มี ' + noRoomWarned + ' บิลที่ช่อง "ห้อง" ไม่ได้ขึ้นต้นด้วยเลขห้อง');
    if (mismatchRows.length) warnings.push('ผลรวมรายการไม่ตรงกับช่อง "รวม" ' + mismatchRows.length + ' แถว');

    var statusCounts = {};
    var rooms = {};
    var minDate = null, maxDate = null;
    invoices.forEach(function (inv) {
      statusCounts[inv.status] = (statusCounts[inv.status] || 0) + 1;
      rooms[inv.room] = true;
      if (!minDate || inv.date < minDate) minDate = inv.date;
      if (!maxDate || inv.date > maxDate) maxDate = inv.date;
    });

    return {
      sheetName: sheetName || '',
      headerRowIndex: headerRowIndex,
      firstDataRow: headerRowIndex + 2,   // 1-based Excel row of first invoice
      lastDataRow: lastRowIndex + 1,      // 1-based Excel row of last invoice
      invoices: invoices,
      itemHeaders: itemHeaders,
      columnStats: columnStats,
      mismatchRows: mismatchRows,
      skippedRows: skippedRows,
      statusCounts: statusCounts,
      rooms: Object.keys(rooms).sort(),
      dateRange: { min: minDate, max: maxDate },
      warnings: warnings
    };
  }

  /**
   * Parse an xlsx ArrayBuffer (or Uint8Array) using the global XLSX (SheetJS).
   * @returns {{invoices, itemHeaders, columnStats, mismatchRows, sheetName, headerRowIndex, warnings, ...}}
   */
  function parseWorkbook(arrayBuffer, options) {
    var XLSX = (options && options.XLSX) || global.XLSX;
    if (!XLSX) throw new Error('ไม่พบไลบรารี SheetJS (XLSX)');
    var data = arrayBuffer instanceof Uint8Array ? arrayBuffer : new Uint8Array(arrayBuffer);
    var wb = XLSX.read(data, { type: 'array', cellDates: false });
    if (!wb.SheetNames || !wb.SheetNames.length) throw new Error('ไฟล์ไม่มีชีต');
    var sheetName = wb.SheetNames[0];
    var ws = wb.Sheets[sheetName];
    var rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
    var result = parseRows(rows, sheetName);
    result.sheetNames = wb.SheetNames.slice();
    return result;
  }

  global.JMP.parser = {
    FIXED_COLUMNS: FIXED_COLUMNS,
    ROOM_UNKNOWN: ROOM_UNKNOWN,
    parseWorkbook: parseWorkbook,
    parseRows: parseRows,
    parseDate: parseDate,
    parseRoom: parseRoom,
    toNumber: toNumber,
    round2: round2
  };
})(typeof window !== 'undefined' ? window : globalThis);
