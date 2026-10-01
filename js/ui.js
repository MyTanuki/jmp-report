window.JMP = window.JMP || {};

/*
 * JMP.ui
 * ------
 * Renders the four screens (load, column mapping, dashboard, settings) into
 * the containers defined in index.html. All rendering is string-based
 * (innerHTML) with event delegation; state lives in JMP.app and is passed in.
 *
 * Every render function takes (state, actions). `actions` is the callback
 * object created by app.js (see app.js for the list). Settings-screen editors
 * mutate state.settings through JMP.settings helpers and then call
 * actions.settingsChanged() so app.js can persist and re-render.
 */
(function (global) {
  'use strict';

  var S = global.JMP.settings;
  var A = global.JMP.aggregate;

  var KIND_LABELS = {
    revenue: 'รายได้หลัก', discount: 'ส่วนลด', utility: 'สาธารณูปโภค', other: 'อื่นๆ', exclude: 'ไม่นับ'
  };
  var STATE_LABELS = { confirmed: 'ยืนยันแล้ว', suggested: 'เดาให้', unassigned: 'ยังไม่กำหนด' };
  var STATE_RANK = { unassigned: 0, suggested: 1, confirmed: 2 };

  var CHART_COLORS = {
    revenue: ['#2563eb', '#0891b2', '#059669', '#7c3aed', '#0d9488', '#4f46e5', '#0284c7', '#65a30d'],
    discount: ['#dc2626', '#ea580c', '#d97706', '#e11d48', '#c026d3', '#b91c1c', '#f97316'],
    other: ['#6b7280', '#9ca3af']
  };

  var chart = null; // current Chart.js instance (trend tab)
  var toastTimer = null;

  // ---------------------------------------------------------------------------
  // Small helpers
  // ---------------------------------------------------------------------------

  function $(sel, root) { return (root || document).querySelector(sel); }
  function $all(sel, root) { return Array.prototype.slice.call((root || document).querySelectorAll(sel)); }

  function esc(s) {
    if (s === null || s === undefined) return '';
    return String(s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }

  /** 1234.5 → "1,234.50" (th-TH). */
  function fmt(n) {
    if (n === null || n === undefined || isNaN(n)) return '';
    return Number(n).toLocaleString('th-TH', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }

  function fmtInt(n) {
    if (n === null || n === undefined || isNaN(n)) return '';
    return Number(n).toLocaleString('th-TH', { maximumFractionDigits: 0 });
  }

  /** Numeric table cell with red negatives and dimmed zeros. */
  function numTd(n, extraClass) {
    var cls = 'num';
    if (n < 0) cls += ' neg';
    if (n === 0) cls += ' zero';
    if (extraClass) cls += ' ' + extraClass;
    return '<td class="' + cls + '">' + fmt(n) + '</td>';
  }

  function toast(msg, kind) {
    var el = $('#toast');
    if (!el) return;
    el.textContent = msg;
    el.className = 'toast' + (kind ? ' ' + kind : '');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function () { el.classList.add('hidden'); }, kind === 'error' ? 6000 : 3500);
  }

  function downloadBlob(blob, fileName) {
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(function () { URL.revokeObjectURL(url); }, 1000);
  }

  /** Write an array-of-arrays to an .xlsx file via SheetJS. */
  function exportAoa(aoa, sheetName, fileName) {
    var XLSX = global.XLSX;
    if (!XLSX) { toast('ไม่พบไลบรารี SheetJS จึงส่งออกไฟล์ไม่ได้', 'error'); return; }
    var ws = XLSX.utils.aoa_to_sheet(aoa);
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, (sheetName || 'Sheet1').slice(0, 31));
    XLSX.writeFile(wb, fileName);
    toast('ส่งออก ' + fileName + ' แล้ว', 'ok');
  }

  /** "2024", "2024-m1-3", "all" plus room tag, used in export file names. */
  function filterTag(filter) {
    var tag = filter.year && filter.year !== 'all' ? String(filter.year) : 'all';
    if (tag !== 'all' && (filter.monthFrom !== 1 || filter.monthTo !== 12)) tag += '-m' + filter.monthFrom + '-' + filter.monthTo;
    if (Array.isArray(filter.rooms)) tag += '-' + filter.rooms.length + 'rooms';
    return tag;
  }

  function optionsHtml(items, selected) {
    return items.map(function (it) {
      return '<option value="' + esc(it.value) + '"' + (String(it.value) === String(selected) ? ' selected' : '') + '>' + esc(it.label) + '</option>';
    }).join('');
  }

  /** <select> options for assignable groups, grouped by kind. */
  function groupOptionsHtml(settings, selected, includeEmpty) {
    var html = includeEmpty ? '<option value=""' + (!selected ? ' selected' : '') + '>— ยังไม่กำหนด —</option>' : '';
    var byKind = {};
    S.assignableGroups(settings).forEach(function (g) {
      (byKind[g.kind] = byKind[g.kind] || []).push(g);
    });
    S.KINDS.forEach(function (kind) {
      if (!byKind[kind]) return;
      html += '<optgroup label="' + esc(KIND_LABELS[kind]) + '">';
      html += byKind[kind].map(function (g) {
        return '<option value="' + esc(g.id) + '"' + (g.id === selected ? ' selected' : '') + '>' + esc(g.label) + '</option>';
      }).join('');
      html += '</optgroup>';
    });
    return html;
  }

  function groupLabel(settings, id) {
    var g = S.groupById(settings, id);
    return g ? g.label : id;
  }

  function countStates(rows) {
    var c = { confirmed: 0, suggested: 0, unassigned: 0 };
    rows.forEach(function (r) { c[r.state] = (c[r.state] || 0) + 1; });
    return c;
  }

  // ---------------------------------------------------------------------------
  // Navigation
  // ---------------------------------------------------------------------------

  function renderNav(state) {
    var hasReport = !!state.report;
    $all('#step-nav .step').forEach(function (btn) {
      var screen = btn.getAttribute('data-screen');
      btn.classList.toggle('active', screen === state.screen);
      if (screen === 'mapping' || screen === 'dashboard') btn.disabled = !hasReport;
    });
    var chip = $('#nav-file');
    if (chip) {
      chip.textContent = state.fileName || '';
      chip.classList.toggle('hidden', !state.fileName);
    }
    var badge = $('#nav-mapping-badge');
    if (badge) {
      var open = 0;
      if (hasReport) {
        S.classifyAll(state.report.itemHeaders, state.settings).forEach(function (c) {
          if (c.state !== 'confirmed') open++;
        });
      }
      badge.textContent = String(open);
      badge.classList.toggle('hidden', !open);
      badge.title = 'คอลัมน์ที่ยังไม่ได้ยืนยัน';
    }
    $all('.screen').forEach(function (sec) {
      sec.classList.toggle('hidden', sec.id !== 'screen-' + state.screen);
    });
  }

  // ---------------------------------------------------------------------------
  // Screen 1: load
  // ---------------------------------------------------------------------------

  function setLoadStatus(msg, kind) {
    var el = $('#load-status');
    if (!el) return;
    if (!msg) { el.classList.add('hidden'); return; }
    el.textContent = msg;
    el.className = 'status' + (kind ? ' ' + kind : '');
  }

  function renderLoad(state, actions) {
    var root = $('#load-summary');
    if (!root) return;
    var rep = state.report;
    if (!rep) { root.innerHTML = ''; return; }

    var cls = S.classifyAll(rep.itemHeaders, state.settings);
    var counts = countStates(cls);
    var statusList = Object.keys(rep.statusCounts).map(function (s) {
      return esc(s) + ' ' + fmtInt(rep.statusCounts[s]);
    }).join(' · ');

    var html = '<div class="card">';
    html += '<div class="section-head"><h2>สรุปการโหลดไฟล์</h2><span class="muted">' + esc(state.fileName) + '</span></div>';
    html += '<dl class="kv">';
    html += kv('ชีต', esc(rep.sheetName));
    html += kv('แถวหัวตาราง', 'แถวที่ ' + (rep.headerRowIndex + 1));
    html += kv('แถวข้อมูล', 'แถวที่ ' + rep.firstDataRow + ' – ' + rep.lastDataRow);
    html += kv('จำนวนบิล', fmtInt(rep.invoices.length));
    html += kv('ช่วงวันที่', (rep.dateRange.min || '-') + ' ถึง ' + (rep.dateRange.max || '-'));
    html += kv('จำนวนห้อง', fmtInt(rep.rooms.length));
    html += kv('คอลัมน์รายการ', fmtInt(rep.itemHeaders.length) + ' คอลัมน์');
    html += kv('สถานะบิล', statusList);
    html += kv('ผลรวมรายการไม่ตรงกับช่อง "รวม"', fmtInt(rep.mismatchRows.length) + ' แถว');
    html += kv('การจับคู่คอลัมน์', 'ยืนยันแล้ว ' + counts.confirmed + ' · เดาให้ ' + counts.suggested + ' · ยังไม่กำหนด ' + counts.unassigned);
    html += '</dl>';

    if (rep.warnings.length) {
      html += '<div class="notice warn" style="margin-top:12px"><strong>ข้อสังเกต</strong><ul>' +
        rep.warnings.map(function (w) { return '<li>' + esc(w) + '</li>'; }).join('') + '</ul></div>';
    }
    if (rep.mismatchRows.length) {
      html += '<details><summary>ดูแถวที่ผลรวมรายการไม่ตรง (' + rep.mismatchRows.length + ' แถว แสดงสูงสุด 50)</summary>';
      html += '<div class="table-wrap short" style="margin-top:8px"><table class="data"><thead><tr><th>แถว</th><th>เลขที่เอกสาร</th><th>วันที่</th><th>ห้อง</th><th class="num">ผลรวมรายการ</th><th class="num">รวม</th><th class="num">ผลต่าง</th></tr></thead><tbody>';
      rep.mismatchRows.slice(0, 50).forEach(function (m) {
        html += '<tr><td>' + m.row + '</td><td>' + esc(m.docNo) + '</td><td>' + esc(m.date) + '</td><td>' + esc(m.room) + '</td>' +
          numTd(m.itemSum) + numTd(m.total) + numTd(m.diff) + '</tr>';
      });
      html += '</tbody></table></div></details>';
    }

    // Primary action = the natural next step: mapping while columns are open, else the dashboard.
    var open = counts.unassigned + counts.suggested;
    html += '<div class="actions">';
    html += '<button type="button"' + (open ? ' class="btn-primary"' : '') + ' data-act="mapping">' +
      (open ? 'ไปจับคู่คอลัมน์ (' + open + ' รายการรอตรวจ)' : 'ดูการจับคู่คอลัมน์') + '</button>';
    html += '<button type="button"' + (open ? '' : ' class="btn-primary"') + ' data-act="dashboard">ไปที่สรุปผล</button>';
    html += '</div></div>';
    root.innerHTML = html;

    $all('[data-act]', root).forEach(function (btn) {
      btn.addEventListener('click', function () { actions.setScreen(btn.getAttribute('data-act')); });
    });
  }

  /** `value` is HTML: callers escape file-derived strings themselves. */
  function kv(label, value) {
    return '<div><dt>' + esc(label) + '</dt><dd>' + value + '</dd></div>';
  }

  // ---------------------------------------------------------------------------
  // Column mapping table (shared by screen 2 and the settings screen)
  // ---------------------------------------------------------------------------

  /**
   * rows: output of JMP.aggregate.columnSummary (plus optional `missing: true`
   * for confirmed columns that are not in the loaded file).
   */
  function sortMappingRows(rows, sort) {
    var key = sort.key || 'state', dir = sort.dir || 1;
    return rows.slice().sort(function (a, b) {
      var d = 0;
      if (key === 'state') {
        d = STATE_RANK[a.state] - STATE_RANK[b.state];
        if (d === 0) d = Math.abs(b.total) - Math.abs(a.total);
      } else if (key === 'name') d = a.normalized.localeCompare(b.normalized, 'th');
      else if (key === 'count') d = a.count - b.count;
      else if (key === 'total') d = a.total - b.total;
      else if (key === 'group') {
        // Sort by the visible Thai label; unassigned rows always last.
        if (!a.group !== !b.group) return a.group ? -1 : 1;
        d = String(a.groupLabel || '').localeCompare(String(b.groupLabel || ''), 'th');
      }
      return d * dir;
    });
  }

  function mappingTableHtml(rows, settings, sort, opts) {
    opts = opts || {};
    function th(key, label, cls) {
      var ind = sort.key === key ? '<span class="sort-ind">' + (sort.dir > 0 ? '▲' : '▼') + '</span>' : '';
      return '<th class="sortable ' + (cls || '') + '" data-sort="' + key + '">' + esc(label) + ind + '</th>';
    }
    var html = '<div class="table-wrap"><table class="data sticky-first" id="' + (opts.tableId || 'mapping-table') + '"><thead><tr>';
    html += th('name', 'คอลัมน์') + th('count', 'จำนวนบิล', 'num') + th('total', 'ยอดรวม', 'num') +
      th('group', 'กลุ่ม') + th('state', 'สถานะ') + '<th>หมายเหตุ</th><th></th></tr></thead><tbody>';
    if (!rows.length) {
      html += '<tr><td colspan="7" class="muted">ไม่มีรายการ</td></tr>';
    }
    rows.forEach(function (r) {
      var kindLabel = r.kind ? KIND_LABELS[r.kind] : '';
      html += '<tr class="state-' + r.state + '" data-norm="' + esc(r.normalized) + '">';
      html += '<td><div>' + esc(r.normalized) + '</div>' +
        (r.raw && r.raw !== r.normalized ? '<div class="raw">' + esc(r.raw) + '</div>' : '') +
        (r.missing && opts.missingLabel ? '<div class="raw">' + esc(opts.missingLabel) + '</div>' : '') + '</td>';
      html += '<td class="num">' + (r.missing ? '' : fmtInt(r.count)) + '</td>';
      html += r.missing ? '<td></td>' : numTd(r.total);
      html += '<td><select class="select-sm" data-act="set" data-norm="' + esc(r.normalized) + '">' +
        groupOptionsHtml(settings, r.group, true) + '</select>' +
        (kindLabel ? ' <span class="badge kind">' + esc(kindLabel) + '</span>' : '') + '</td>';
      html += '<td><span class="badge ' + r.state + '">' + STATE_LABELS[r.state] + '</span>' +
        (r.rule ? ' <span class="muted small" title="' + esc(r.rule.pattern) + '">กฎ #' + (r.rule.index + 1) + '</span>' : '') + '</td>';
      var hint = '';
      if (r.signHint) {
        hint = r.kind === 'revenue'
          ? 'ยอดรวมติดลบแต่อยู่ในกลุ่มรายได้ — อาจเป็นส่วนลด'
          : 'ยอดรวมเป็นบวกแต่อยู่ในกลุ่มส่วนลด — อาจเป็นการยกเลิกส่วนลด';
      } else if (r.state === 'unassigned') {
        hint = 'ไม่มีกฎที่ตรง — จะถูกนับในกลุ่ม "อื่นๆ" จนกว่าจะกำหนด';
      }
      html += '<td>' + (hint ? '<span class="hint">' + esc(hint) + '</span>' : '') + '</td>';
      html += '<td class="nowrap">';
      if (r.state === 'suggested') html += '<button type="button" class="btn-sm" data-act="confirm" data-norm="' + esc(r.normalized) + '" data-group="' + esc(r.group) + '">ยืนยัน</button> ';
      if (r.state === 'confirmed') html += '<button type="button" class="btn-sm" data-act="clear" data-norm="' + esc(r.normalized) + '">' + (opts.clearLabel || 'ล้าง') + '</button>';
      html += '</td></tr>';
    });
    html += '</tbody></table></div>';
    return html;
  }

  /** Delegated events for a mapping table: select change, confirm, clear, sort. */
  function bindMappingTable(root, actions, onSort) {
    root.addEventListener('change', function (ev) {
      var t = ev.target;
      if (t.getAttribute('data-act') !== 'set') return;
      var norm = t.getAttribute('data-norm');
      if (t.value) actions.setColumn(norm, t.value);
      else actions.clearColumn(norm);
    });
    root.addEventListener('click', function (ev) {
      var btn = ev.target.closest('[data-act]');
      if (btn) {
        var act = btn.getAttribute('data-act');
        if (act === 'confirm') actions.setColumn(btn.getAttribute('data-norm'), btn.getAttribute('data-group'));
        else if (act === 'clear') actions.clearColumn(btn.getAttribute('data-norm'));
        return;
      }
      var th = ev.target.closest('th[data-sort]');
      if (th && onSort) onSort(th.getAttribute('data-sort'));
    });
  }

  function filterMappingRows(rows, search, onlyOpen) {
    var q = (search || '').trim().toLowerCase();
    return rows.filter(function (r) {
      if (onlyOpen && r.state === 'confirmed') return false;
      if (q && r.normalized.toLowerCase().indexOf(q) < 0 && String(r.raw || '').toLowerCase().indexOf(q) < 0) return false;
      return true;
    });
  }

  // ---------------------------------------------------------------------------
  // Screen 2: mapping
  // ---------------------------------------------------------------------------

  function renderMapping(state, actions) {
    var root = $('#mapping-root');
    if (!root) return;
    if (!state.report) {
      root.innerHTML = '<div class="card"><p class="muted">ยังไม่ได้โหลดไฟล์</p></div>';
      return;
    }
    var m = state.mapping;
    var html = '<div class="card">';
    html += '<div class="section-head"><h2>จับคู่คอลัมน์รายการกับกลุ่มรายได้</h2>' +
      '<div class="actions" style="margin:0"><button type="button" class="btn-primary" data-act="dashboard">ไปที่สรุปผล</button></div></div>';
    html += '<p class="muted">คอลัมน์ที่ "เดาให้" มาจากกฎอัตโนมัติ กดยืนยันหรือเลือกกลุ่มเองเพื่อบันทึกถาวร คอลัมน์ที่ยังไม่กำหนดจะถูกนับในกลุ่ม "อื่นๆ" ในหน้าสรุป</p>';
    html += '<div class="toolbar">';
    html += '<input type="search" id="map-search" placeholder="ค้นหาชื่อคอลัมน์" value="' + esc(m.search) + '" style="min-width:240px">';
    html += '<label><input type="checkbox" id="map-only-open"' + (m.onlyOpen ? ' checked' : '') + '> แสดงเฉพาะที่ยังไม่ยืนยัน</label>';
    html += '<span class="spacer"></span>';
    html += '<span id="map-counts" class="muted small"></span>';
    html += '<button type="button" id="map-confirm-all">ยืนยันทั้งหมดที่เดาให้</button>';
    html += '</div>';
    html += '<div id="map-table"></div>';
    html += '</div>';
    root.innerHTML = html;

    $('[data-act="dashboard"]', root).addEventListener('click', function () { actions.setScreen('dashboard'); });
    $('#map-search', root).addEventListener('input', function (ev) {
      state.mapping.search = ev.target.value;
      renderMappingTable(state, actions);
    });
    $('#map-only-open', root).addEventListener('change', function (ev) {
      state.mapping.onlyOpen = ev.target.checked;
      renderMappingTable(state, actions);
    });
    $('#map-confirm-all', root).addEventListener('click', function () { actions.confirmAllSuggested(); });
    bindMappingTable($('#map-table', root), actions, function (key) {
      var s = state.mapping.sort;
      if (s.key === key) s.dir = -s.dir; else { s.key = key; s.dir = 1; }
      renderMappingTable(state, actions);
    });
    renderMappingTable(state, actions);
  }

  /** Re-render only the table (keeps the search box focus). */
  function renderMappingTable(state, actions) {
    var holder = $('#map-table');
    if (!holder || !state.report) return;
    var rep = state.report;
    var all = A.columnSummary(rep.itemHeaders, rep.columnStats, state.settings);
    var counts = countStates(all);
    var rows = sortMappingRows(filterMappingRows(all, state.mapping.search, state.mapping.onlyOpen), state.mapping.sort);
    holder.innerHTML = mappingTableHtml(rows, state.settings, state.mapping.sort, { tableId: 'mapping-table' });
    var c = $('#map-counts');
    if (c) c.textContent = 'ยืนยันแล้ว ' + counts.confirmed + ' · เดาให้ ' + counts.suggested + ' · ยังไม่กำหนด ' + counts.unassigned +
      ' · แสดง ' + rows.length + '/' + all.length;
    var bulk = $('#map-confirm-all');
    if (bulk) bulk.disabled = counts.suggested === 0;
  }

  // ---------------------------------------------------------------------------
  // Screen 3: dashboard
  // ---------------------------------------------------------------------------

  function renderFilterBar(state, actions) {
    var root = $('#filter-bar');
    if (!root || !state.report) return;
    var rep = state.report, f = state.filter;
    var years = A.distinctYears(rep.invoices);
    var rooms = A.distinctRooms(rep.invoices);
    var statuses = A.distinctStatuses(rep.invoices);
    var months = A.MONTHS_TH.map(function (label, i) { return { value: i + 1, label: label }; });
    var yearDisabled = f.year === 'all';

    var html = '';
    html += '<label class="field"><span>ปี</span><select id="f-year">' +
      optionsHtml([{ value: 'all', label: 'ทุกปี' }].concat(years.map(function (y) { return { value: y, label: String(y) }; })), f.year) +
      '</select></label>';
    html += '<label class="field"><span>ตั้งแต่เดือน</span><select id="f-month-from"' + (yearDisabled ? ' disabled' : '') + '>' + optionsHtml(months, f.monthFrom) + '</select></label>';
    html += '<label class="field"><span>ถึงเดือน</span><select id="f-month-to"' + (yearDisabled ? ' disabled' : '') + '>' + optionsHtml(months, f.monthTo) + '</select></label>';

    var roomSel = Array.isArray(f.rooms) ? f.rooms : null;
    var roomLabel = roomSel ? (roomSel.length === 0 ? 'ไม่เลือกห้อง' : 'เลือก ' + roomSel.length + ' ห้อง') : 'ทุกห้อง (' + rooms.length + ')';
    html += '<div class="field"><span>ห้อง</span><details class="dropdown" id="f-rooms"><summary id="f-rooms-label">' + esc(roomLabel) + '</summary>' +
      '<div class="dropdown-panel"><div class="actions"><button type="button" class="btn-sm" data-act="rooms-all">ทั้งหมด</button>' +
      '<button type="button" class="btn-sm" data-act="rooms-none">ไม่เลือก</button></div><div class="grid">' +
      rooms.map(function (r) {
        var checked = !roomSel || roomSel.indexOf(r) >= 0;
        return '<label><input type="checkbox" data-room="' + esc(r) + '"' + (checked ? ' checked' : '') + '> ' + esc(r) + '</label>';
      }).join('') + '</div></div></details></div>';

    html += '<div class="field"><span>สถานะ</span><div class="status-group" id="f-status">' +
      statuses.map(function (s) {
        var checked = f.statuses.indexOf(s.status) >= 0;
        return '<label><input type="checkbox" data-status="' + esc(s.status) + '"' + (checked ? ' checked' : '') + '> ' +
          esc(s.status) + ' <span class="muted">(' + fmtInt(s.count) + ')</span></label>';
      }).join('') +
      '<button type="button" class="btn-sm" id="f-status-save" title="บันทึกสถานะที่เลือกไว้เป็นค่าเริ่มต้นสำหรับการโหลดไฟล์ครั้งต่อไป (เก็บใน localStorage และไฟล์ตั้งค่า JSON)">จำเป็นค่าเริ่มต้น</button>' +
      '</div></div>';

    html += '<label class="field"><span>ค้นหาลูกค้า</span><input type="search" id="f-customer" placeholder="ชื่อลูกค้า" value="' + esc(f.customer) + '"></label>';
    root.innerHTML = html;

    $('#f-year', root).addEventListener('change', function (ev) {
      var v = ev.target.value;
      actions.setFilter({ year: v === 'all' ? 'all' : +v });
      var dis = v === 'all';
      $('#f-month-from', root).disabled = dis;
      $('#f-month-to', root).disabled = dis;
    });
    $('#f-month-from', root).addEventListener('change', function (ev) {
      var from = +ev.target.value, to = state.filter.monthTo;
      if (to < from) { to = from; $('#f-month-to', root).value = String(to); }
      actions.setFilter({ monthFrom: from, monthTo: to });
    });
    $('#f-month-to', root).addEventListener('change', function (ev) {
      var to = +ev.target.value, from = state.filter.monthFrom;
      if (from > to) { from = to; $('#f-month-from', root).value = String(from); }
      actions.setFilter({ monthFrom: from, monthTo: to });
    });

    var roomsBox = $('#f-rooms', root);
    function readRooms() {
      var boxes = $all('input[data-room]', roomsBox);
      var chosen = boxes.filter(function (b) { return b.checked; }).map(function (b) { return b.getAttribute('data-room'); });
      var all = chosen.length === boxes.length;
      $('#f-rooms-label', root).textContent = all ? 'ทุกห้อง (' + boxes.length + ')' : (chosen.length === 0 ? 'ไม่เลือกห้อง' : 'เลือก ' + chosen.length + ' ห้อง');
      actions.setFilter({ rooms: all ? null : chosen });
    }
    roomsBox.addEventListener('change', function (ev) { if (ev.target.hasAttribute('data-room')) readRooms(); });
    roomsBox.addEventListener('click', function (ev) {
      var btn = ev.target.closest('[data-act]');
      if (!btn) return;
      var on = btn.getAttribute('data-act') === 'rooms-all';
      $all('input[data-room]', roomsBox).forEach(function (b) { b.checked = on; });
      readRooms();
    });
    // Closing the dropdown on an outside click is handled once in app.js.

    $('#f-status', root).addEventListener('change', function () {
      var chosen = $all('#f-status input[data-status]', root).filter(function (b) { return b.checked; })
        .map(function (b) { return b.getAttribute('data-status'); });
      actions.setFilter({ statuses: chosen });
    });
    $('#f-status-save', root).addEventListener('click', function () { actions.saveStatusDefault(); });

    var custTimer = null;
    $('#f-customer', root).addEventListener('input', function (ev) {
      clearTimeout(custTimer);
      var v = ev.target.value;
      custTimer = setTimeout(function () { actions.setFilter({ customer: v }); }, 250);
    });
  }

  function renderDashboard(state, actions) {
    if (!state.report) return;
    renderFilterBar(state, actions);
    renderDashboardBody(state, actions);
  }

  /** Metrics + warnings + the active tab. Called on every filter/tab change. */
  function renderDashboardBody(state, actions) {
    if (!state.report) return;
    var rep = state.report, groups = state.settings.groups, colMap = state.colMap;
    var filtered = A.applyFilters(rep.invoices, state.filter);
    state.filtered = filtered;

    // Warnings
    var warn = $('#dash-warning');
    if (warn) {
      var cls = S.classifyAll(rep.itemHeaders, state.settings);
      var unassigned = cls.filter(function (c) { return c.state === 'unassigned'; }).length;
      var suggested = cls.filter(function (c) { return c.state === 'suggested'; }).length;
      var items = [];
      if (unassigned) items.push('มี ' + unassigned + ' คอลัมน์ที่ยังไม่กำหนดกลุ่ม (ถูกนับในกลุ่ม "อื่นๆ")');
      if (suggested) items.push('มี ' + suggested + ' คอลัมน์ที่ใช้กลุ่มจากการเดา ยังไม่ได้ยืนยัน');
      if (rep.mismatchRows.length) items.push('ผลรวมรายการไม่ตรงกับช่อง "รวม" ' + rep.mismatchRows.length + ' แถว (ดูได้ในหน้าโหลดไฟล์)');
      warn.innerHTML = items.length
        ? '<div class="notice warn">' + items.map(esc).join(' · ') +
          (unassigned || suggested ? ' <span class="link" data-act="mapping">ไปจับคู่คอลัมน์</span>' : '') + '</div>'
        : '';
      var lnk = $('[data-act="mapping"]', warn);
      if (lnk) lnk.addEventListener('click', function () { actions.setScreen('mapping'); });
    }

    // Metrics
    var m = A.metrics(filtered, colMap, groups);
    var metricsEl = $('#metrics');
    if (metricsEl) {
      metricsEl.innerHTML = [
        metric('รายได้หลัก', m.revenue),
        metric('ส่วนลด', m.discount),
        metric('รายได้สุทธิ', m.net, 'รายได้หลัก + ส่วนลด + อื่นๆ'),
        metric('สาธารณูปโภค (น้ำ+ไฟ)', m.utility),
        metric('ค้างชำระ', m.unpaid, fmtInt(m.unpaidCount) + ' บิล'),
        '<div class="metric"><div class="label">จำนวนบิล</div><div class="value">' + fmtInt(m.count) + '</div>' +
          '<div class="sub">รวมสุทธิ (บิล) ' + fmt(m.invoiceNet) + '</div></div>'
      ].join('');
    }

    $all('#tabs .tab').forEach(function (t) { t.classList.toggle('active', t.getAttribute('data-tab') === state.tab); });
    var body = $('#tab-body');
    if (!body) return;
    if (state.tab !== 'trend' && chart) { chart.destroy(); chart = null; }
    if (state.tab === 'roomGroup') renderRoomGroupTab(body, state, actions, filtered);
    else if (state.tab === 'pivot') renderPivotTab(body, state, actions, filtered);
    else if (state.tab === 'trend') renderTrendTab(body, state, actions, filtered);
    else renderInvoicesTab(body, state, actions, filtered);
  }

  function metric(label, value, sub) {
    return '<div class="metric"><div class="label">' + esc(label) + '</div>' +
      '<div class="value' + (value < 0 ? ' neg' : '') + '">' + fmt(value) + '</div>' +
      (sub ? '<div class="sub">' + esc(sub) + '</div>' : '') + '</div>';
  }

  /** Column list for room × group and invoice tables. */
  function groupTableColumns(settings, opts) {
    opts = opts || {};
    var cols = [];
    A.groupColumns(settings.groups).forEach(function (gc) {
      if (gc.kind === 'exclude' && !opts.showExclude) return;
      if (gc.children.length && opts.expandParents) {
        gc.children.forEach(function (c) { cols.push({ key: c.id, label: c.label, kind: c.kind, child: true }); });
        cols.push({ key: gc.id, label: gc.label + ' (รวม)', kind: gc.kind, parent: true });
      } else {
        cols.push({ key: gc.id, label: gc.label, kind: gc.kind });
      }
    });
    return cols;
  }

  // --- Tab 1: room × group -------------------------------------------------

  function renderRoomGroupTab(body, state, actions, filtered) {
    var rg = state.roomGroup;
    var data = A.roomByGroup(filtered, state.colMap, state.settings.groups);
    var gcols = groupTableColumns(state.settings, { expandParents: rg.expandDiscount, showExclude: rg.showExclude });
    var cols = [{ key: 'room', label: 'ห้อง' }, { key: 'count', label: 'จำนวนบิล', num: true }]
      .concat(gcols.map(function (c) { c.num = true; c.group = true; return c; }))
      .concat([{ key: 'net', label: 'รายได้สุทธิ', num: true, title: 'รายได้หลัก + ส่วนลด + อื่นๆ (ไม่รวมน้ำ/ไฟ)' }]);

    function val(row, c) {
      if (c.key === 'room') return row.room;
      if (c.key === 'count') return row.count;
      if (c.key === 'net') return row.net;
      return row.groups[c.key] || 0;
    }
    // The sorted-on group may have been deleted in settings since the last render.
    if (!cols.some(function (c) { return c.key === rg.sort.key; })) rg.sort = { key: 'room', dir: 1 };
    var rows = data.rows.slice();
    if (rg.sort.key !== 'room' || rg.sort.dir !== 1) {
      var sc = cols.filter(function (c) { return c.key === rg.sort.key; })[0];
      rows.sort(function (a, b) {
        var x = val(a, sc), y = val(b, sc), d;
        if (sc.key === 'room') d = (parseInt(x, 10) || 0) - (parseInt(y, 10) || 0) || String(x).localeCompare(String(y));
        else d = x - y;
        return d * rg.sort.dir;
      });
    }

    var html = '<div class="toolbar">';
    html += '<label><input type="checkbox" id="rg-expand"' + (rg.expandDiscount ? ' checked' : '') + '> แยกส่วนลดเป็นกลุ่มย่อย</label>';
    html += '<label><input type="checkbox" id="rg-exclude"' + (rg.showExclude ? ' checked' : '') + '> แสดงกลุ่ม "ไม่นับ"</label>';
    html += '<span class="spacer"></span><span class="muted small">' + fmtInt(rows.length) + ' ห้อง · ' + fmtInt(data.totals.count) + ' บิล · คลิกหัวคอลัมน์เพื่อเรียง คลิกเลขห้องเพื่อดูบิล</span>';
    html += '<button type="button" id="rg-export">ส่งออก .xlsx</button></div>';
    html += '<div class="table-wrap"><table class="data sticky-first" id="room-group-table"><thead><tr>';
    cols.forEach(function (c) {
      var ind = rg.sort.key === c.key ? '<span class="sort-ind">' + (rg.sort.dir > 0 ? '▲' : '▼') + '</span>' : '';
      html += '<th class="sortable' + (c.num ? ' num' : '') + (c.child ? ' child-col' : '') + (c.group ? ' group-col' : '') + '" data-sort="' + esc(c.key) + '"' +
        (c.title ? ' title="' + esc(c.title) + '"' : '') + '>' + esc(c.label) + ind + '</th>';
    });
    html += '</tr></thead><tbody>';
    rows.forEach(function (r) {
      html += '<tr>';
      cols.forEach(function (c) {
        if (c.key === 'room') html += '<td class="room-cell" data-room="' + esc(r.room) + '">' + esc(r.room) + '</td>';
        else if (c.key === 'count') html += '<td class="num">' + fmtInt(r.count) + '</td>';
        else html += numTd(val(r, c));
      });
      html += '</tr>';
    });
    html += '<tr class="total">';
    cols.forEach(function (c) {
      if (c.key === 'room') html += '<td>รวม</td>';
      else if (c.key === 'count') html += '<td class="num">' + fmtInt(data.totals.count) + '</td>';
      else html += numTd(val(data.totals, c));
    });
    html += '</tr></tbody></table></div>';
    body.innerHTML = html;

    $('#rg-expand', body).addEventListener('change', function (ev) { rg.expandDiscount = ev.target.checked; renderDashboardBody(state, actions); });
    $('#rg-exclude', body).addEventListener('change', function (ev) { rg.showExclude = ev.target.checked; renderDashboardBody(state, actions); });
    $('#rg-export', body).addEventListener('click', function () {
      var aoa = [cols.map(function (c) { return c.label; })];
      rows.concat([data.totals]).forEach(function (r) {
        aoa.push(cols.map(function (c) { return c.key === 'room' ? r.room : val(r, c); }));
      });
      exportAoa(aoa, 'ห้อง x กลุ่ม', 'jmp-room-group-' + filterTag(state.filter) + '.xlsx');
    });
    bindTableClicks($('#room-group-table', body), actions, function (key) {
      if (rg.sort.key === key) rg.sort.dir = -rg.sort.dir; else { rg.sort.key = key; rg.sort.dir = key === 'room' ? 1 : -1; }
      renderDashboardBody(state, actions);
    });
  }

  /** Room-cell click → invoices tab; header click → sort callback. */
  function bindTableClicks(table, actions, onSort) {
    if (!table) return;
    table.addEventListener('click', function (ev) {
      var cell = ev.target.closest('td.room-cell');
      if (cell) { actions.showRoomInvoices(cell.getAttribute('data-room')); return; }
      var th = ev.target.closest('th[data-sort]');
      if (th && onSort) onSort(th.getAttribute('data-sort'));
    });
  }

  // --- Tab 2: monthly pivot ------------------------------------------------

  function renderPivotTab(body, state, actions, filtered) {
    var pv = state.pivot;
    var groups = state.settings.groups;
    var year = state.filter.year === 'all' ? null : state.filter.year;
    var valueOpts = A.valueOptions(groups);
    // The selected group may have been deleted in settings since the last render.
    if (!valueOpts.some(function (o) { return o.key === pv.valueKey; })) pv.valueKey = A.VALUE_KEYS.REVENUE;
    var data = A.pivotMonthly(filtered, state.colMap, {
      year: year, range: state.report.dateRange, valueKey: pv.valueKey, byCustomer: pv.byCustomer, groups: groups
    });
    var isCount = pv.valueKey === A.VALUE_KEYS.COUNT;
    var cell = function (v) { return isCount ? '<td class="num' + (v === 0 ? ' zero' : '') + '">' + fmtInt(v) + '</td>' : numTd(v); };

    var opts = valueOpts.map(function (o) {
      var label = o.label;
      // Browsers collapse leading spaces in <option>, so indent children with a marker.
      if (o.kind) label = (o.parent ? '   ↳ ' : '') + label + ' [' + KIND_LABELS[o.kind] + ']';
      return { value: o.key, label: label };
    });

    var html = '<div class="toolbar">';
    html += '<label>ค่าที่แสดง <select id="pv-value">' + optionsHtml(opts, pv.valueKey) + '</select></label>';
    html += '<label><input type="checkbox" id="pv-customer"' + (pv.byCustomer ? ' checked' : '') + '> แยกตามลูกค้า</label>';
    html += '<span class="spacer"></span><span class="muted small">' + (year ? 'ปี ' + year : 'ทุกปี (คอลัมน์ = ทุกเดือนในช่วงข้อมูลของไฟล์)') + ' · ' + esc(data.valueLabel) + '</span>';
    html += '<button type="button" id="pv-export">ส่งออก .xlsx</button></div>';
    html += '<div class="table-wrap"><table class="data sticky-first" id="pivot-table"><thead><tr><th>' + (pv.byCustomer ? 'ห้อง / ลูกค้า' : 'ห้อง') + '</th>';
    data.columns.forEach(function (c) { html += '<th class="num">' + esc(c.label) + '</th>'; });
    html += '<th class="num">Grand Total</th></tr></thead><tbody>';
    data.rows.forEach(function (r) {
      html += '<tr class="level-' + r.level + '">';
      if (r.level === 0) html += '<td class="room-cell" data-room="' + esc(r.room) + '">' + esc(r.room) + '</td>';
      else html += '<td>' + esc(r.customer || '(ไม่ระบุ)') + '</td>';
      r.values.forEach(function (v) { html += cell(v); });
      html += cell(r.total) + '</tr>';
    });
    html += '<tr class="total"><td>Grand Total</td>';
    data.grandTotal.values.forEach(function (v) { html += cell(v); });
    html += cell(data.grandTotal.total) + '</tr></tbody></table></div>';
    body.innerHTML = html;

    $('#pv-value', body).addEventListener('change', function (ev) { pv.valueKey = ev.target.value; renderDashboardBody(state, actions); });
    $('#pv-customer', body).addEventListener('change', function (ev) { pv.byCustomer = ev.target.checked; renderDashboardBody(state, actions); });
    $('#pv-export', body).addEventListener('click', function () {
      var head = ['ห้อง'].concat(pv.byCustomer ? ['ลูกค้า'] : []).concat(data.columns.map(function (c) { return c.label; })).concat(['Grand Total']);
      var aoa = [head];
      data.rows.forEach(function (r) {
        var lead = pv.byCustomer ? [r.level === 0 ? r.room : '', r.level === 0 ? '' : r.customer] : [r.room];
        aoa.push(lead.concat(r.values).concat([r.total]));
      });
      aoa.push(['Grand Total'].concat(pv.byCustomer ? [''] : []).concat(data.grandTotal.values).concat([data.grandTotal.total]));
      exportAoa(aoa, 'Pivot', 'jmp-pivot-' + (isCount ? 'count' : pv.valueKey.replace(/^__/, '')) + '-' + filterTag(state.filter) + '.xlsx');
    });
    bindTableClicks($('#pivot-table', body), actions, null);
  }

  // --- Tab 3: trend chart --------------------------------------------------

  function renderTrendTab(body, state, actions, filtered) {
    var year = state.filter.year === 'all' ? null : state.filter.year;
    var ms = A.monthlySeries(filtered, state.colMap, state.settings.groups, { year: year, range: state.report.dateRange });
    body.innerHTML = '<div class="toolbar"><span class="muted small">แท่ง = แต่ละกลุ่ม (ส่วนลดติดลบ) · เส้น = รายได้สุทธิ' +
      (year ? ' · ปี ' + year : ' · ทุกปี (ทุกเดือนในช่วงข้อมูลของไฟล์ เดือนที่ไม่มีบิล = 0)') + '</span></div>' +
      '<div class="chart-wrap"><canvas id="trend-chart"></canvas></div>';
    if (!global.Chart) {
      body.innerHTML += '<div class="notice error">ไม่พบไลบรารี Chart.js (ต้องเชื่อมต่ออินเทอร์เน็ตเพื่อโหลดจาก CDN)</div>';
      return;
    }
    if (!ms.keys.length) {
      body.innerHTML += '<p class="muted">ไม่มีข้อมูลตามตัวกรองที่เลือก</p>';
      return;
    }
    var counters = { revenue: 0, discount: 0, other: 0 };
    var datasets = ms.series.filter(function (s) {
      return s.data.some(function (v) { return v !== 0; });
    }).map(function (s) {
      var palette = CHART_COLORS[s.kind] || CHART_COLORS.other;
      var color = palette[counters[s.kind]++ % palette.length];
      return { type: 'bar', label: s.label, data: s.data, backgroundColor: color, stack: 'groups', order: 2 };
    });
    datasets.unshift({
      type: 'line', label: 'รายได้สุทธิ', data: ms.net, borderColor: '#111827', backgroundColor: '#111827',
      borderWidth: 2, pointRadius: 3, tension: 0.2, order: 1
    });
    if (chart) chart.destroy();
    chart = new global.Chart($('#trend-chart', body).getContext('2d'), {
      data: { labels: ms.labels, datasets: datasets },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        interaction: { mode: 'index', intersect: false },
        scales: {
          x: { stacked: true },
          y: { stacked: true, ticks: { callback: function (v) { return fmtInt(v); } } }
        },
        plugins: {
          legend: { position: 'bottom', labels: { boxWidth: 12, font: { family: 'Segoe UI, Tahoma, sans-serif' } } },
          tooltip: {
            callbacks: {
              label: function (ctx) { return ctx.dataset.label + ': ' + fmt(ctx.parsed.y); },
              footer: function (items) {
                var i = items.length ? items[0].dataIndex : -1;
                return i >= 0 ? 'จำนวนบิล ' + fmtInt(ms.count[i]) : '';
              }
            }
          }
        }
      }
    });
  }

  // --- Tab 4: invoice detail -----------------------------------------------

  function renderInvoicesTab(body, state, actions, filtered) {
    var invs = filtered;
    if (state.invoiceRoom) invs = invs.filter(function (i) { return i.room === state.invoiceRoom; });
    var rows = A.invoiceRows(invs, state.colMap, state.settings.groups);
    var gcols = groupTableColumns(state.settings, { expandParents: false, showExclude: true });
    var fixed = [
      { key: 'date', label: 'วันที่' }, { key: 'docNo', label: 'เลขที่เอกสาร' }, { key: 'room', label: 'ห้อง' },
      { key: 'customer', label: 'ลูกค้า' }, { key: 'status', label: 'สถานะ' }
    ];
    var tail = [
      { key: 'netRevenue', label: 'รายได้สุทธิ' }, { key: 'net', label: 'รวมสุทธิ (บิล)' }
    ];

    var html = '<div class="toolbar">';
    if (state.invoiceRoom) html += '<span class="chip">ห้อง ' + esc(state.invoiceRoom) + ' <button type="button" id="inv-clear-room" title="ยกเลิกตัวกรองห้อง">×</button></span>';
    html += '<span class="muted small">' + fmtInt(rows.length) + ' บิล</span><span class="spacer"></span>';
    html += '<button type="button" id="inv-export">ส่งออก .xlsx</button></div>';
    html += '<div class="table-wrap"><table class="data" id="invoice-table"><thead><tr>';
    fixed.forEach(function (c) { html += '<th>' + esc(c.label) + '</th>'; });
    gcols.forEach(function (c) { html += '<th class="num">' + esc(c.label) + '</th>'; });
    tail.forEach(function (c) { html += '<th class="num">' + esc(c.label) + '</th>'; });
    html += '</tr></thead><tbody>';
    rows.forEach(function (r) {
      html += '<tr' + (r.unassignedTotal ? ' class="state-unassigned" title="มีรายการที่ยังไม่กำหนดกลุ่ม"' : '') + '>';
      html += '<td class="nowrap">' + esc(r.date) + '</td><td class="nowrap">' + esc(r.docNo) + '</td>' +
        '<td title="' + esc(r.roomDesc) + '">' + esc(r.room) + '</td><td>' + esc(r.customer) + '</td><td class="nowrap">' + esc(r.status) + '</td>';
      gcols.forEach(function (c) { html += numTd(r.groups[c.key] || 0); });
      html += numTd(r.netRevenue) + numTd(r.net) + '</tr>';
    });
    if (!rows.length) html += '<tr><td colspan="' + (fixed.length + gcols.length + tail.length) + '" class="muted">ไม่มีบิลตามตัวกรอง</td></tr>';
    html += '</tbody></table></div>';
    body.innerHTML = html;

    var clr = $('#inv-clear-room', body);
    if (clr) clr.addEventListener('click', function () { actions.showRoomInvoices(null); });
    $('#inv-export', body).addEventListener('click', function () {
      var head = fixed.concat(gcols).concat(tail).map(function (c) { return c.label; });
      var aoa = [head];
      rows.forEach(function (r) {
        aoa.push([r.date, r.docNo, r.room, r.customer, r.status]
          .concat(gcols.map(function (c) { return r.groups[c.key] || 0; }))
          .concat([r.netRevenue, r.net]));
      });
      exportAoa(aoa, 'รายละเอียดบิล', 'jmp-invoices-' + filterTag(state.filter) + (state.invoiceRoom ? '-room' + state.invoiceRoom : '') + '.xlsx');
    });
  }

  // ---------------------------------------------------------------------------
  // Settings screen
  // ---------------------------------------------------------------------------

  function renderSettings(state, actions) {
    var root = $('#settings-root');
    if (!root) return;
    var settings = state.settings;
    var html = '';

    // Toolbar
    html += '<div class="card"><div class="section-head"><h2>ตั้งค่า</h2><div class="actions" style="margin:0">' +
      '<button type="button" id="btn-export-settings">ส่งออกการตั้งค่า (JSON)</button>' +
      '<button type="button" id="btn-import-settings">นำเข้าการตั้งค่า</button>' +
      '<input type="file" id="import-input" accept=".json,application/json" hidden>' +
      '<button type="button" id="btn-save-default">บันทึกเป็นค่าเริ่มต้น</button>' +
      '<button type="button" class="btn-danger" id="btn-reset-settings">คืนค่าเริ่มต้น</button></div></div>' +
      '<p class="muted">การตั้งค่าถูกเก็บในเบราว์เซอร์นี้ (localStorage) ส่งออกเป็นไฟล์ JSON เพื่อสำรองหรือย้ายไปเครื่องอื่น</p>' +
      '<p class="muted small">เบราว์เซอร์ที่ยังไม่เคยตั้งค่าจะเริ่มจากค่าเริ่มต้นในไฟล์ <code>js/default-settings.js</code> ' +
      (S.shippedDefaults() ? ''
        : global.JMP_DEFAULT_SETTINGS == null ? '(ไม่พบไฟล์นี้หรือโหลดไม่ได้ จึงใช้ค่าเริ่มต้นในตัว) '
        : '(ไฟล์นี้มีข้อมูลไม่ถูกต้อง จึงใช้ค่าเริ่มต้นในตัว ดูรายละเอียดใน console) ') +
      'กด "บันทึกเป็นค่าเริ่มต้น" เพื่อดาวน์โหลดไฟล์นี้จากการตั้งค่าปัจจุบัน แล้ววางทับในโฟลเดอร์ <code>js/</code> และ commit</p></div>';

    // Groups
    html += '<div class="card settings-section"><h3>กลุ่มรายได้</h3>' +
      '<p class="muted small">กลุ่มที่มีกลุ่มย่อย (เช่น ส่วนลด) ใช้เป็นหัวรวมเท่านั้น จับคู่คอลัมน์ได้เฉพาะกลุ่มย่อย</p>';
    html += '<div class="table-wrap short"><table class="data" id="groups-table"><thead><tr><th>ชื่อกลุ่ม</th><th>ประเภท</th><th>กลุ่มแม่</th><th>id</th><th>ลำดับ</th><th></th></tr></thead><tbody>';
    S.sortedGroups(settings).forEach(function (g) {
      var isBucket = g.id === S.UNASSIGNED_GROUP;
      html += '<tr data-id="' + esc(g.id) + '"' + (g.parent ? ' class="level-1"' : '') + '>';
      html += '<td><input type="text" class="input-sm" data-field="label" value="' + esc(g.label) + '" style="width:100%;min-width:180px"></td>';
      html += '<td><select class="select-sm" data-field="kind">' + optionsHtml(S.KINDS.map(function (k) { return { value: k, label: KIND_LABELS[k] }; }), g.kind) + '</select></td>';
      html += '<td><select class="select-sm" data-field="parent"' + (S.isParent(settings, g.id) ? ' disabled title="กลุ่มนี้มีกลุ่มย่อย"' : '') + '>' +
        groupParentOptionsHtml(settings, g.id, g.parent) + '</select></td>';
      html += '<td class="muted small group-id">' + groupIdCellHtml(settings, g) + '</td>';
      html += '<td class="nowrap"><button type="button" class="btn-icon" data-act="group-up" title="เลื่อนขึ้น">▲</button> <button type="button" class="btn-icon" data-act="group-down" title="เลื่อนลง">▼</button></td>';
      html += '<td><button type="button" class="btn-sm btn-danger" data-act="group-delete"' +
        (isBucket ? ' disabled title="กลุ่มนี้ใช้รับคอลัมน์ที่ยังไม่กำหนด ลบไม่ได้"' : '') + '>ลบ</button></td></tr>';
    });
    html += '</tbody></table></div>';
    html += '<div class="inline-form"><label class="field"><span>ชื่อกลุ่มใหม่</span><input type="text" id="new-group-label" placeholder="เช่น ค่าที่จอดรถ"></label>' +
      '<label class="field"><span>ประเภท</span><select id="new-group-kind">' + optionsHtml(S.KINDS.map(function (k) { return { value: k, label: KIND_LABELS[k] }; }), 'revenue') + '</select></label>' +
      '<label class="field"><span>กลุ่มแม่</span><select id="new-group-parent">' + groupParentOptionsHtml(settings, null, '') + '</select></label>' +
      '<button type="button" id="btn-add-group">เพิ่มกลุ่ม</button></div></div>';

    // Rules
    var groupSel = function (sel) { return groupOptionsHtml(settings, sel, false); };
    html += '<div class="card settings-section"><h3>กฎจับคู่อัตโนมัติ</h3>' +
      '<p class="muted small">regex (ไม่สนตัวพิมพ์ใหญ่-เล็ก) ทดสอบกับชื่อคอลัมน์ที่ตัด "(non-vat)" ออกแล้ว เรียงจากบนลงล่าง กฎแรกที่ตรงจะถูกใช้ การยืนยันคอลัมน์เองมีผลเหนือกฎ</p>';
    html += '<div class="table-wrap short"><table class="data" id="rules-table"><thead><tr><th>#</th><th>pattern</th><th>flags</th><th>กลุ่ม</th><th>ลำดับ</th><th></th></tr></thead><tbody>';
    settings.rules.forEach(function (r, i) {
      html += '<tr data-index="' + i + '"><td class="muted">' + (i + 1) + '</td>';
      html += '<td><input type="text" class="input-sm mono" data-field="pattern" value="' + esc(r.pattern) + '" style="width:100%;min-width:260px"></td>';
      html += '<td><input type="text" class="input-sm mono" data-field="flags" value="' + esc(r.flags === undefined ? 'i' : r.flags) + '" style="width:50px"></td>';
      html += '<td><select class="select-sm" data-field="group">' + groupSel(r.group) + '</select></td>';
      html += '<td class="nowrap"><button type="button" class="btn-icon" data-act="rule-up">▲</button> <button type="button" class="btn-icon" data-act="rule-down">▼</button></td>';
      html += '<td class="nowrap"><button type="button" class="btn-sm" data-act="rule-test"' +
        (state.report ? '' : ' disabled title="โหลดไฟล์ก่อนจึงจะทดสอบกับชื่อคอลัมน์ได้"') + '>ทดสอบ</button> ' +
        '<button type="button" class="btn-sm btn-danger" data-act="rule-delete">ลบ</button></td></tr>';
    });
    html += '</tbody></table></div>';
    html += '<div class="inline-form"><label class="field" style="flex:1"><span>pattern ใหม่</span><input type="text" id="new-rule-pattern" class="mono" placeholder="เช่น ^ค่าที่จอดรถ" style="width:100%"></label>' +
      '<label class="field"><span>กลุ่ม</span><select id="new-rule-group">' + groupSel('') + '</select></label>' +
      '<label class="field"><span>ตำแหน่ง</span><select id="new-rule-pos"><option value="end">ต่อท้าย</option><option value="start">แทรกบนสุด</option></select></label>' +
      '<button type="button" id="btn-add-rule">เพิ่มกฎ</button></div></div>';

    // Columns
    html += '<div class="card settings-section"><div class="section-head"><h3>การจับคู่คอลัมน์</h3>' +
      '<input type="search" id="cols-search" placeholder="ค้นหาชื่อคอลัมน์" style="min-width:220px"></div>' +
      '<p class="muted small">' + (state.report ? 'แสดงทุกคอลัมน์ในไฟล์ที่โหลด รวมทั้งคอลัมน์ที่เคยยืนยันไว้แต่ไม่มีในไฟล์นี้' : 'ยังไม่ได้โหลดไฟล์ — แสดงเฉพาะคอลัมน์ที่เคยยืนยันไว้') + '</p>' +
      '<div id="cols-table"></div></div>';

    root.innerHTML = html;
    bindSettingsEvents(root, state, actions);
    renderSettingsColumns(state, actions);
  }

  /** Parent <select> options: top-level groups except `excludeId` (the row's own group). */
  function groupParentOptionsHtml(settings, excludeId, selected) {
    var opts = [{ value: '', label: '—' }].concat(S.sortedGroups(settings)
      .filter(function (x) { return !x.parent && x.id !== excludeId; })
      .map(function (x) { return { value: x.id, label: x.label }; }));
    return optionsHtml(opts, selected || '');
  }

  function groupIdCellHtml(settings, g) {
    var mappedCount = S.referencesTo(settings, g.id).columns.length;
    return esc(g.id) + (mappedCount ? '<br>' + mappedCount + ' คอลัมน์' : '');
  }

  /**
   * Refresh the parts of the settings screen that depend on group labels /
   * kinds / parents without replacing the editors, so a field edit keeps
   * keyboard focus, the rule-test row and the column search text.
   */
  function renderSettingsFields(state, actions) {
    var root = $('#settings-root');
    if (!root) return;
    var settings = state.settings;
    $all('#groups-table tr[data-id]', root).forEach(function (tr) {
      var g = S.groupById(settings, tr.getAttribute('data-id'));
      if (!g) return;
      tr.classList.toggle('level-1', !!g.parent);
      var sel = $('select[data-field="parent"]', tr);
      sel.innerHTML = groupParentOptionsHtml(settings, g.id, g.parent);
      var container = S.isParent(settings, g.id);
      sel.disabled = container;
      sel.title = container ? 'กลุ่มนี้มีกลุ่มย่อย' : '';
      $('td.group-id', tr).innerHTML = groupIdCellHtml(settings, g);
    });
    var newParent = $('#new-group-parent', root);
    newParent.innerHTML = groupParentOptionsHtml(settings, null, newParent.value);
    $all('#rules-table tr[data-index]', root).forEach(function (tr) {
      var r = settings.rules[+tr.getAttribute('data-index')];
      if (r) $('select[data-field="group"]', tr).innerHTML = groupOptionsHtml(settings, r.group, false);
    });
    var newRule = $('#new-rule-group', root);
    newRule.innerHTML = groupOptionsHtml(settings, newRule.value, false);
    renderSettingsColumns(state, actions);
  }

  function settingsColumnRows(state) {
    var settings = state.settings;
    var rows = [];
    var seen = {};
    if (state.report) {
      rows = A.columnSummary(state.report.itemHeaders, state.report.columnStats, settings);
      rows.forEach(function (r) { seen[r.normalized] = true; });
    }
    Object.keys(settings.columns || {}).forEach(function (norm) {
      if (seen[norm]) return;
      var gid = settings.columns[norm];
      var g = S.groupById(settings, gid);
      rows.push({ raw: norm, normalized: norm, group: gid, groupLabel: g ? g.label : '', state: 'confirmed', rule: null,
                  kind: g ? g.kind : null, count: 0, total: 0, signHint: false, missing: true });
    });
    return rows;
  }

  function renderSettingsColumns(state, actions) {
    var holder = $('#cols-table');
    if (!holder) return;
    var search = ($('#cols-search') || {}).value || '';
    var rows = sortMappingRows(filterMappingRows(settingsColumnRows(state), search, false), state.settingsSort);
    holder.innerHTML = mappingTableHtml(rows, state.settings, state.settingsSort, {
      tableId: 'settings-columns-table', clearLabel: 'ล้างการยืนยัน',
      // Without a file every stored column is "missing"; the note above the table already says so.
      missingLabel: state.report ? 'ไม่มีในไฟล์ที่โหลด' : ''
    });
  }

  function bindSettingsEvents(root, state, actions) {
    var settings = state.settings;

    /** Run a settings mutation; `scope` = 'fields' keeps the editors' DOM (see app.js afterSettingsChange). */
    function guard(fn, scope) {
      try { fn(); actions.settingsChanged(scope); }
      catch (e) { toast(e.message || String(e), 'error'); renderSettings(state, actions); }
    }

    /** Making `parentId` a container moves its columns/rules to the new child; tell the user. */
    function toastRetarget(parentId, before) {
      if (!before || (!before.columns.length && !before.rules.length)) return;
      toast('กลุ่ม "' + groupLabel(settings, parentId) + '" กลายเป็นกลุ่มแม่ — ย้ายคอลัมน์ ' + before.columns.length +
        ' และกฎ ' + before.rules.length + ' ไปยังกลุ่มย่อยใหม่', 'ok');
    }
    function refsIfNewParent(parentId) {
      return parentId && !S.isParent(settings, parentId) ? S.referencesTo(settings, parentId) : null;
    }

    // Toolbar
    $('#btn-export-settings', root).addEventListener('click', function () {
      downloadBlob(new Blob([S.exportJSON(settings)], { type: 'application/json' }), 'jmp-settings.json');
      toast('ส่งออก jmp-settings.json แล้ว', 'ok');
    });
    $('#btn-import-settings', root).addEventListener('click', function () { $('#import-input', root).click(); });
    $('#import-input', root).addEventListener('change', function (ev) {
      var file = ev.target.files && ev.target.files[0];
      if (!file) return;
      var reader = new FileReader();
      reader.onload = function () {
        try {
          var next = S.importJSON(String(reader.result));
          actions.replaceSettings(next);
          toast('นำเข้าการตั้งค่าแล้ว', 'ok');
        } catch (e) { toast(e.message, 'error'); }
      };
      reader.onerror = function () { toast('อ่านไฟล์ไม่ได้', 'error'); };
      reader.readAsText(file);
      ev.target.value = '';
    });
    $('#btn-save-default', root).addEventListener('click', function () {
      var js = S.defaultSettingsJS(settings, new Date().toISOString());
      downloadBlob(new Blob([js], { type: 'text/javascript' }), 'default-settings.js');
      toast('ดาวน์โหลด default-settings.js แล้ว วางทับไฟล์ในโฟลเดอร์ js/ แล้ว commit เพื่อใช้เป็นค่าเริ่มต้น', 'ok');
    });
    $('#btn-reset-settings', root).addEventListener('click', function () {
      if (!global.confirm('คืนค่าเริ่มต้นทั้งหมด? กลุ่ม กฎ และการยืนยันคอลัมน์ในเบราว์เซอร์นี้จะถูกแทนด้วยค่าเริ่มต้นของแอป (js/default-settings.js)')) return;
      actions.replaceSettings(S.reset());
      toast('คืนค่าเริ่มต้นแล้ว', 'ok');
    });

    // Groups
    var gt = $('#groups-table', root);
    gt.addEventListener('change', function (ev) {
      var tr = ev.target.closest('tr[data-id]');
      if (!tr) return;
      var id = tr.getAttribute('data-id'), field = ev.target.getAttribute('data-field'), value = ev.target.value;
      guard(function () {
        if (field === 'label') S.renameGroup(settings, id, value);
        else if (field === 'kind') S.updateGroup(settings, id, { kind: value });
        else if (field === 'parent') {
          var before = refsIfNewParent(value);
          S.updateGroup(settings, id, { parent: value || null });
          toastRetarget(value, before);
        }
      }, 'fields');
    });
    gt.addEventListener('click', function (ev) {
      var btn = ev.target.closest('[data-act]');
      if (!btn) return;
      var id = btn.closest('tr[data-id]').getAttribute('data-id');
      var act = btn.getAttribute('data-act');
      if (act === 'group-up') guard(function () { S.moveGroup(settings, id, -1); });
      else if (act === 'group-down') guard(function () { S.moveGroup(settings, id, 1); });
      else if (act === 'group-delete') {
        var g = S.groupById(settings, id);
        var msg = 'ลบกลุ่ม "' + (g ? g.label : id) + '"? คอลัมน์ที่ยืนยันไว้กับกลุ่มนี้จะกลับเป็น "ยังไม่กำหนด" และกฎที่ชี้มากลุ่มนี้จะถูกลบ';
        if (!global.confirm(msg)) return;
        guard(function () {
          var res = S.deleteGroup(settings, id);
          toast('ลบกลุ่มแล้ว (คอลัมน์ ' + res.removedColumns.length + ' · กฎ ' + res.removedRules + ')', 'ok');
        });
      }
    });
    $('#btn-add-group', root).addEventListener('click', function () {
      var label = $('#new-group-label', root).value.trim();
      if (!label) { toast('กรุณาใส่ชื่อกลุ่ม', 'error'); return; }
      var parent = $('#new-group-parent', root).value || undefined;
      guard(function () {
        var before = refsIfNewParent(parent);
        S.addGroup(settings, { label: label, kind: $('#new-group-kind', root).value, parent: parent });
        toastRetarget(parent, before);
      });
    });

    // Rules
    var rt = $('#rules-table', root);
    rt.addEventListener('change', function (ev) {
      var tr = ev.target.closest('tr[data-index]');
      if (!tr) return;
      var idx = +tr.getAttribute('data-index'), field = ev.target.getAttribute('data-field');
      var patch = {};
      patch[field] = ev.target.value;
      guard(function () { S.updateRule(settings, idx, patch); }, 'fields');
    });
    rt.addEventListener('click', function (ev) {
      var btn = ev.target.closest('[data-act]');
      if (!btn) return;
      var tr = btn.closest('tr[data-index]');
      var idx = +tr.getAttribute('data-index');
      var act = btn.getAttribute('data-act');
      if (act === 'rule-up') guard(function () { S.moveRule(settings, idx, -1); });
      else if (act === 'rule-down') guard(function () { S.moveRule(settings, idx, 1); });
      else if (act === 'rule-delete') guard(function () { S.deleteRule(settings, idx); });
      else if (act === 'rule-test') {
        if (!state.report) { toast('โหลดไฟล์ก่อนจึงจะทดสอบกฎกับชื่อคอลัมน์ได้', 'error'); return; }
        var rule = { pattern: $('[data-field="pattern"]', tr).value, flags: $('[data-field="flags"]', tr).value };
        var res = S.testRule(rule, state.report.itemHeaders);
        var old = tr.nextElementSibling;
        if (old && old.classList.contains('rule-test')) old.parentNode.removeChild(old);
        var row = document.createElement('tr');
        row.className = 'rule-test';
        row.innerHTML = '<td></td><td colspan="5">' + (!res.valid ? '<span class="neg">regex ไม่ถูกต้อง</span>' :
          (res.matches.length ? 'ตรง ' + res.matches.length + ' คอลัมน์ (ก่อนพิจารณากฎอื่นและการยืนยัน):<ul class="rule-matches">' +
            res.matches.map(function (m) { return '<li>' + esc(m) + '</li>'; }).join('') + '</ul>' : 'ไม่ตรงกับคอลัมน์ใดเลย')) + '</td>';
        tr.parentNode.insertBefore(row, tr.nextSibling);
      }
    });
    $('#btn-add-rule', root).addEventListener('click', function () {
      var pattern = $('#new-rule-pattern', root).value.trim();
      if (!pattern) { toast('กรุณาใส่ pattern', 'error'); return; }
      var pos = $('#new-rule-pos', root).value;
      guard(function () {
        S.addRule(settings, { pattern: pattern, group: $('#new-rule-group', root).value }, pos === 'start' ? 0 : undefined);
      });
    });

    // Columns
    $('#cols-search', root).addEventListener('input', function () { renderSettingsColumns(state, actions); });
    bindMappingTable($('#cols-table', root), actions, function (key) {
      var s = state.settingsSort;
      if (s.key === key) s.dir = -s.dir; else { s.key = key; s.dir = 1; }
      renderSettingsColumns(state, actions);
    });
  }

  global.JMP.ui = {
    KIND_LABELS: KIND_LABELS,
    STATE_LABELS: STATE_LABELS,
    esc: esc,
    fmt: fmt,
    fmtInt: fmtInt,
    toast: toast,
    exportAoa: exportAoa,
    renderNav: renderNav,
    setLoadStatus: setLoadStatus,
    renderLoad: renderLoad,
    renderMapping: renderMapping,
    renderMappingTable: renderMappingTable,
    renderDashboard: renderDashboard,
    renderFilterBar: renderFilterBar,
    renderDashboardBody: renderDashboardBody,
    renderSettings: renderSettings,
    renderSettingsFields: renderSettingsFields,
    renderSettingsColumns: renderSettingsColumns
  };
})(typeof window !== 'undefined' ? window : globalThis);
