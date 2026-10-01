window.JMP = window.JMP || {};

/*
 * JMP.app
 * -------
 * Application state, screen routing and event wiring. The DOM is rendered by
 * JMP.ui; data comes from JMP.parser / JMP.settings / JMP.aggregate.
 *
 * window.JMP.app.loadArrayBuffer(arrayBuffer, fileName) loads a workbook
 * programmatically (used by smoke checks).
 */
(function (global) {
  'use strict';

  var S = global.JMP.settings;
  var P = global.JMP.parser;
  var A = global.JMP.aggregate;
  var UI = global.JMP.ui;

  var state = {
    settings: null,
    report: null,
    fileName: '',
    screen: 'load',
    colMap: {},
    filtered: [],
    filter: { year: 'all', monthFrom: 1, monthTo: 12, rooms: null, statuses: [], customer: '' },
    tab: 'roomGroup',
    invoiceRoom: null,
    mapping: { search: '', onlyOpen: false, sort: { key: 'state', dir: 1 } },
    settingsSort: { key: 'state', dir: 1 },
    roomGroup: { sort: { key: 'room', dir: 1 }, expandDiscount: false, showExclude: false },
    pivot: { valueKey: A.VALUE_KEYS.REVENUE, byCustomer: false }
  };

  function $(sel) { return document.querySelector(sel); }

  // ---------------------------------------------------------------------------
  // Derived data
  // ---------------------------------------------------------------------------

  function recomputeColMap() {
    state.colMap = state.report ? S.columnMap(state.report.itemHeaders, state.settings) : {};
  }

  /**
   * Default status set = statuses present in the file ∩ settings.statusInclude;
   * falls back to the built-in default set, then (unknown vocabulary) to every status present.
   */
  function defaultStatuses() {
    if (!state.report) return [];
    var present = A.distinctStatuses(state.report.invoices).map(function (s) { return s.status; });
    function pick(list) { return present.filter(function (s) { return list.indexOf(s) >= 0; }); }
    var chosen = pick(state.settings.statusInclude || []);
    if (!chosen.length) chosen = pick(S.DEFAULT_SETTINGS.statusInclude);
    return chosen.length ? chosen : present;
  }

  function resetFilters() {
    state.filter = { year: 'all', monthFrom: 1, monthTo: 12, rooms: null, statuses: defaultStatuses(), customer: '' };
    state.invoiceRoom = null;
  }

  // ---------------------------------------------------------------------------
  // Rendering
  // ---------------------------------------------------------------------------

  function renderCurrent() {
    UI.renderNav(state);
    if (state.screen === 'load') UI.renderLoad(state, actions);
    else if (state.screen === 'mapping') UI.renderMapping(state, actions);
    else if (state.screen === 'dashboard') UI.renderDashboard(state, actions);
    else if (state.screen === 'settings') UI.renderSettings(state, actions);
  }

  function setScreen(name) {
    if ((name === 'mapping' || name === 'dashboard') && !state.report) {
      UI.toast('กรุณาโหลดไฟล์ก่อน', 'error');
      name = 'load';
    }
    state.screen = name;
    renderCurrent();
    window.scrollTo(0, 0);
  }

  /**
   * Persist settings and refresh whatever depends on them.
   * `scope` = 'columns' (column mapping changed) | 'fields' (a group/rule field was
   * edited in place on the settings screen; keeps the editors' DOM) | 'all'.
   */
  function afterSettingsChange(scope) {
    S.save(state.settings);
    recomputeColMap();
    UI.renderNav(state);
    if (state.screen === 'mapping') {
      if (scope === 'columns') UI.renderMappingTable(state, actions); else UI.renderMapping(state, actions);
    } else if (state.screen === 'settings') {
      if (scope === 'columns') UI.renderSettingsColumns(state, actions);
      else if (scope === 'fields') UI.renderSettingsFields(state, actions);
      else UI.renderSettings(state, actions);
    } else if (state.screen === 'dashboard') {
      UI.renderDashboardBody(state, actions);
    } else {
      UI.renderLoad(state, actions);
    }
  }

  // ---------------------------------------------------------------------------
  // Loading
  // ---------------------------------------------------------------------------

  /**
   * Parse a workbook already in memory and show the load summary (row range,
   * status counts, parser warnings, mismatch rows) with buttons to continue.
   * @returns {object|null} the parsed report, or null on failure
   */
  function loadArrayBuffer(arrayBuffer, fileName) {
    if (!global.XLSX) {
      UI.setLoadStatus('ไม่พบไลบรารี SheetJS — ต้องเชื่อมต่ออินเทอร์เน็ตเพื่อโหลดจาก CDN', 'error');
      return null;
    }
    UI.setLoadStatus('กำลังอ่านไฟล์ ' + (fileName || '') + ' ...');
    var report;
    try {
      report = P.parseWorkbook(arrayBuffer);
    } catch (e) {
      UI.setLoadStatus('อ่านไฟล์ไม่สำเร็จ: ' + (e && e.message ? e.message : e), 'error');
      return null;
    }
    state.report = report;
    state.fileName = fileName || '';
    recomputeColMap();
    resetFilters();
    state.mapping.search = '';
    UI.setLoadStatus('โหลดสำเร็จ: ' + report.invoices.length.toLocaleString('th-TH') + ' บิล, ' +
      report.itemHeaders.length + ' คอลัมน์รายการ', 'ok');

    var open = S.classifyAll(report.itemHeaders, state.settings).filter(function (c) { return c.state !== 'confirmed'; }).length;
    // Stay on the load screen so the summary and parser warnings are seen.
    state.screen = 'load';
    renderCurrent();
    if (open === 0) UI.toast('ทุกคอลัมน์ยืนยันแล้ว — ไปที่สรุปผลได้เลย', 'ok');
    else UI.toast('โหลดสำเร็จ — มี ' + open + ' คอลัมน์รอตรวจในขั้นที่ 2', 'ok');
    var summary = $('#load-summary');
    if (summary && summary.scrollIntoView) summary.scrollIntoView({ behavior: 'smooth', block: 'start' });
    return report;
  }

  function loadFile(file) {
    if (!file) return;
    if (!/\.(xlsx|xlsm|xls)$/i.test(file.name)) {
      UI.setLoadStatus('กรุณาเลือกไฟล์ Excel (.xlsx)', 'error');
      return;
    }
    UI.setLoadStatus('กำลังอ่านไฟล์ ' + file.name + ' ...');
    var reader = new FileReader();
    reader.onload = function () {
      // Defer so the status text paints before the (synchronous) parse.
      setTimeout(function () { loadArrayBuffer(reader.result, file.name); }, 20);
    };
    reader.onerror = function () { UI.setLoadStatus('อ่านไฟล์ไม่ได้', 'error'); };
    reader.readAsArrayBuffer(file);
  }

  // ---------------------------------------------------------------------------
  // Actions exposed to JMP.ui
  // ---------------------------------------------------------------------------

  var actions = {
    setScreen: setScreen,
    loadFile: loadFile,

    setColumn: function (normalized, groupId) {
      try {
        S.setColumn(state.settings, normalized, groupId);
        afterSettingsChange('columns');
      } catch (e) { UI.toast(e.message, 'error'); }
    },
    clearColumn: function (normalized) {
      S.clearColumn(state.settings, normalized);
      afterSettingsChange('columns');
    },
    confirmAllSuggested: function () {
      if (!state.report) return;
      var n = 0;
      S.classifyAll(state.report.itemHeaders, state.settings).forEach(function (c) {
        if (c.state === 'suggested' && c.group) { S.setColumn(state.settings, c.normalized, c.group); n++; }
      });
      afterSettingsChange('columns');
      UI.toast('ยืนยันแล้ว ' + n + ' คอลัมน์', 'ok');
    },

    /** Dashboard filters are session-only; see saveStatusDefault for persisting the status set. */
    setFilter: function (patch) {
      Object.keys(patch).forEach(function (k) { state.filter[k] = patch[k]; });
      UI.renderDashboardBody(state, actions);
    },
    /** Persist the checked status set as settings.statusInclude (explicit user action). */
    saveStatusDefault: function () {
      if (!state.report) return;
      if (!state.filter.statuses.length) { UI.toast('เลือกอย่างน้อย 1 สถานะก่อนบันทึกเป็นค่าเริ่มต้น', 'error'); return; }
      // Keep entries for statuses absent from this file.
      var present = A.distinctStatuses(state.report.invoices).map(function (s) { return s.status; });
      var kept = state.settings.statusInclude.filter(function (s) { return present.indexOf(s) < 0; });
      state.settings.statusInclude = state.filter.statuses.concat(kept);
      S.save(state.settings);
      UI.toast('บันทึกสถานะที่เลือกเป็นค่าเริ่มต้นแล้ว', 'ok');
    },
    setTab: function (tab) {
      state.tab = tab;
      UI.renderDashboardBody(state, actions);
    },
    /** Room cell click anywhere → invoice tab filtered to that room (null clears). */
    showRoomInvoices: function (room) {
      state.invoiceRoom = (room === null || room === undefined) ? null : String(room);
      state.tab = 'invoices';
      UI.renderDashboardBody(state, actions);
      var body = $('#tab-body');
      if (body && state.invoiceRoom !== null) body.scrollIntoView({ behavior: 'smooth', block: 'start' });
    },

    settingsChanged: function (scope) { afterSettingsChange(scope || 'all'); },
    replaceSettings: function (next) {
      state.settings = next;
      state.filter.statuses = defaultStatuses();
      afterSettingsChange('all');
    }
  };

  // ---------------------------------------------------------------------------
  // Bootstrap
  // ---------------------------------------------------------------------------

  function bindStatic() {
    document.querySelectorAll('#step-nav .step').forEach(function (btn) {
      btn.addEventListener('click', function () { setScreen(btn.getAttribute('data-screen')); });
    });

    var zone = $('#drop-zone');
    var input = $('#file-input');
    if (zone && input) {
      zone.addEventListener('click', function () { input.click(); });
      zone.addEventListener('keydown', function (ev) {
        if (ev.key === 'Enter' || ev.key === ' ') { ev.preventDefault(); input.click(); }
      });
      ['dragenter', 'dragover'].forEach(function (evt) {
        zone.addEventListener(evt, function (ev) { ev.preventDefault(); zone.classList.add('dragover'); });
      });
      ['dragleave', 'drop'].forEach(function (evt) {
        zone.addEventListener(evt, function (ev) { ev.preventDefault(); zone.classList.remove('dragover'); });
      });
      zone.addEventListener('drop', function (ev) {
        var files = ev.dataTransfer && ev.dataTransfer.files;
        if (files && files.length) loadFile(files[0]);
      });
      input.addEventListener('change', function () {
        if (input.files && input.files.length) loadFile(input.files[0]);
        input.value = ''; // allow re-selecting the same file
      });
    }
    // Dropping a file anywhere else must not navigate away from the page.
    document.addEventListener('dragover', function (ev) { ev.preventDefault(); });
    document.addEventListener('drop', function (ev) { ev.preventDefault(); });

    document.querySelectorAll('#tabs .tab').forEach(function (btn) {
      btn.addEventListener('click', function () { actions.setTab(btn.getAttribute('data-tab')); });
    });

    // Close any open dropdown (room picker) when clicking outside it.
    document.addEventListener('click', function (ev) {
      document.querySelectorAll('details.dropdown[open]').forEach(function (d) {
        if (!d.contains(ev.target)) d.open = false;
      });
    });
  }

  function init() {
    state.settings = S.load();
    bindStatic();
    renderCurrent();
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
  else init();

  global.JMP.app = {
    init: init,
    state: state,
    actions: actions,
    setScreen: setScreen,
    loadArrayBuffer: loadArrayBuffer,
    loadFile: loadFile
  };
})(typeof window !== 'undefined' ? window : globalThis);
