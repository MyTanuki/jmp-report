window.JMP = window.JMP || {};

/*
 * JMP.aggregate
 * -------------
 * Pure functions over Invoice[] (from JMP.parser) and a column map
 * `colMap = { [rawHeader]: groupId | null }` (from JMP.settings.columnMap).
 *
 * Group semantics:
 *   kind 'revenue'  → รายได้หลัก
 *   kind 'discount' → ส่วนลด (negative amounts)
 *   kind 'other'    → อื่นๆ (counted in net revenue)
 *   kind 'utility'  → ค่าน้ำ / ค่าไฟ (reported apart from revenue)
 *   kind 'exclude'  → ไม่นับ (never part of any revenue figure)
 *   unassigned column (null) → added to group 'other' and tracked in unassignedTotal
 *
 * รายได้สุทธิ (net) = revenue + discount + other.
 */
(function (global) {
  'use strict';

  var MONTHS_TH = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];
  // Shared with JMP.settings, which guarantees the group exists and cannot be deleted.
  var UNASSIGNED_GROUP = (global.JMP.settings && global.JMP.settings.UNASSIGNED_GROUP) || 'other';
  var STATUS_UNPAID = 'ค้างชำระ';

  // Value selector keys that are not group ids (prefixed to avoid collisions).
  var VALUE_KEYS = {
    REVENUE: '__revenue',       // รายได้หลัก (kind revenue)
    NET: '__net',               // รายได้สุทธิ (revenue + discount + other)
    INVOICE_NET: '__invoiceNet',// รวมสุทธิ of the bill (column รวมสุทธิ)
    UTILITY: '__utility',       // น้ำ + ไฟ
    COUNT: '__count'            // number of invoices
  };

  var VALUE_LABELS = {
    __revenue: 'รายได้หลัก',
    __net: 'รายได้สุทธิ',
    __invoiceNet: 'รวมสุทธิ (บิล)',
    __utility: 'สาธารณูปโภค (น้ำ+ไฟ)',
    __count: 'จำนวนบิล'
  };

  function round2(n) {
    return Math.round(n * 100) / 100;
  }

  function pad2(n) {
    return (n < 10 ? '0' : '') + n;
  }

  // ---------------------------------------------------------------------------
  // Group index
  // ---------------------------------------------------------------------------

  /**
   * Build a lookup over settings.groups. Accepts either the groups array or a
   * full settings object. Cheap enough to rebuild on every call (≤ dozens of groups).
   */
  function groupIndex(groups) {
    if (groups && groups.groups) groups = groups.groups;
    groups = groups || [];
    var byId = {};
    var sorted = groups.slice().sort(function (a, b) { return (a.order || 0) - (b.order || 0); });
    sorted.forEach(function (g) { byId[g.id] = g; });
    var parents = {};
    sorted.forEach(function (g) { if (g.parent) parents[g.parent] = true; });
    return {
      byId: byId,
      sorted: sorted,
      assignable: sorted.filter(function (g) { return !parents[g.id]; }),
      // Unknown ids (e.g. a deleted group still referenced) count as 'other'.
      kindOf: function (id) {
        var g = byId[id];
        return g ? g.kind : 'other';
      },
      parentOf: function (id) {
        var g = byId[id];
        return g && g.parent ? g.parent : null;
      },
      isParent: function (id) { return !!parents[id]; }
    };
  }

  /**
   * Column layout for the ห้อง × กลุ่ม table: revenue groups, then discount
   * parents (with children), then other, then utility. Exclude is appended
   * last (kind exclude) so the UI can decide whether to show it.
   * @returns {Array<{id,label,kind,children:Array<{id,label,kind}>}>}
   */
  function groupColumns(groups) {
    var gi = groupIndex(groups);
    var kindOrder = { revenue: 0, discount: 1, other: 2, utility: 3, exclude: 4 };
    var top = gi.sorted.filter(function (g) { return !g.parent; });
    top.sort(function (a, b) {
      var ka = kindOrder[a.kind], kb = kindOrder[b.kind];
      if (ka !== kb) return ka - kb;
      return (a.order || 0) - (b.order || 0);
    });
    return top.map(function (g) {
      return {
        id: g.id, label: g.label, kind: g.kind,
        children: gi.sorted.filter(function (c) { return c.parent === g.id; })
          .map(function (c) { return { id: c.id, label: c.label, kind: c.kind }; })
      };
    });
  }

  // ---------------------------------------------------------------------------
  // Filters
  // ---------------------------------------------------------------------------

  /**
   * filter = {
   *   year?: number|null,          null / undefined / 'all' → all years
   *   monthFrom?: 1..12, monthTo?: 1..12   (only meaningful with a year)
   *   rooms?: string[]|null,       null → all
   *   statuses?: string[]|null,    null → all
   *   customer?: string,           case-insensitive substring
   *   dateFrom?: 'YYYY-MM-DD', dateTo?: 'YYYY-MM-DD'
   * }
   */
  function applyFilters(invoices, filter) {
    filter = filter || {};
    var year = (filter.year === 'all' || filter.year === '' || filter.year === undefined) ? null : filter.year;
    if (year !== null) year = +year;
    var mFrom = filter.monthFrom ? +filter.monthFrom : 1;
    var mTo = filter.monthTo ? +filter.monthTo : 12;
    var rooms = Array.isArray(filter.rooms) ? filter.rooms.map(String) : null;
    var roomSet = rooms ? rooms.reduce(function (m, r) { m[r] = true; return m; }, {}) : null;
    var statusSet = Array.isArray(filter.statuses)
      ? filter.statuses.reduce(function (m, s) { m[s] = true; return m; }, {}) : null;
    var cust = filter.customer ? String(filter.customer).trim().toLowerCase() : '';
    var dFrom = filter.dateFrom || null;
    var dTo = filter.dateTo || null;

    return invoices.filter(function (inv) {
      if (year !== null) {
        if (inv.year !== year) return false;
        if (inv.month < mFrom || inv.month > mTo) return false;
      }
      if (roomSet && !roomSet[inv.room]) return false;
      if (statusSet && !statusSet[inv.status]) return false;
      if (cust && inv.customer.toLowerCase().indexOf(cust) < 0) return false;
      if (dFrom && inv.date < dFrom) return false;
      if (dTo && inv.date > dTo) return false;
      return true;
    });
  }

  // ---------------------------------------------------------------------------
  // Per-invoice group amounts
  // ---------------------------------------------------------------------------

  /**
   * Sum one invoice's items per group.
   * @returns {{
   *   groups: {[groupId]: number},   // includes parent totals (children rolled up)
   *   byKind: {revenue, discount, other, utility, exclude},
   *   revenue, discount, other, utility, exclude, net, unassignedTotal
   * }}
   */
  function invoiceGroupAmounts(inv, colMap, groups) {
    var gi = groupIndex(groups || []);
    var out = {};
    var byKind = { revenue: 0, discount: 0, other: 0, utility: 0, exclude: 0 };
    var unassigned = 0;
    var items = inv.items || {};
    for (var raw in items) {
      if (!Object.prototype.hasOwnProperty.call(items, raw)) continue;
      var v = items[raw];
      var gid = colMap ? colMap[raw] : null;
      if (!gid) { gid = UNASSIGNED_GROUP; unassigned += v; }
      out[gid] = (out[gid] || 0) + v;
      var kind = gi.kindOf(gid);
      byKind[kind] = (byKind[kind] || 0) + v;
      // roll up into parent chain
      var p = gi.parentOf(gid), guard = 0;
      while (p && guard++ < 10) { out[p] = (out[p] || 0) + v; p = gi.parentOf(p); }
    }
    for (var k in out) out[k] = round2(out[k]);
    for (var kk in byKind) byKind[kk] = round2(byKind[kk]);
    return {
      groups: out,
      byKind: byKind,
      revenue: byKind.revenue,
      discount: byKind.discount,
      other: byKind.other,
      utility: byKind.utility,
      exclude: byKind.exclude,
      net: round2(byKind.revenue + byKind.discount + byKind.other),
      unassignedTotal: round2(unassigned)
    };
  }

  /** Value selector: groupId | VALUE_KEYS.* → number for one invoice. */
  function valueOf(inv, colMap, groups, valueKey) {
    if (valueKey === VALUE_KEYS.COUNT) return 1;
    if (valueKey === VALUE_KEYS.INVOICE_NET) return inv.net;
    var a = invoiceGroupAmounts(inv, colMap, groups);
    if (valueKey === VALUE_KEYS.REVENUE) return a.revenue;
    if (valueKey === VALUE_KEYS.NET) return a.net;
    if (valueKey === VALUE_KEYS.UTILITY) return a.utility;
    return a.groups[valueKey] || 0;
  }

  /** Options for the pivot value selector: [{key,label,kind?}] */
  function valueOptions(groups) {
    var gi = groupIndex(groups);
    var opts = [
      { key: VALUE_KEYS.REVENUE, label: VALUE_LABELS.__revenue },
      { key: VALUE_KEYS.NET, label: VALUE_LABELS.__net },
      { key: VALUE_KEYS.INVOICE_NET, label: VALUE_LABELS.__invoiceNet },
      { key: VALUE_KEYS.UTILITY, label: VALUE_LABELS.__utility },
      { key: VALUE_KEYS.COUNT, label: VALUE_LABELS.__count }
    ];
    gi.sorted.forEach(function (g) {
      opts.push({ key: g.id, label: g.label, kind: g.kind, parent: g.parent || null });
    });
    return opts;
  }

  function valueLabel(valueKey, groups) {
    if (VALUE_LABELS[valueKey]) return VALUE_LABELS[valueKey];
    var g = groupIndex(groups).byId[valueKey];
    return g ? g.label : String(valueKey);
  }

  // ---------------------------------------------------------------------------
  // Room × group
  // ---------------------------------------------------------------------------

  function newAccumulator(key) {
    return { key: key, count: 0, groups: {}, revenue: 0, discount: 0, other: 0, utility: 0, exclude: 0,
             net: 0, invoiceNet: 0, unassignedTotal: 0 };
  }

  function accumulate(acc, amounts, inv) {
    acc.count += 1;
    for (var g in amounts.groups) acc.groups[g] = (acc.groups[g] || 0) + amounts.groups[g];
    acc.revenue += amounts.revenue;
    acc.discount += amounts.discount;
    acc.other += amounts.other;
    acc.utility += amounts.utility;
    acc.exclude += amounts.exclude;
    acc.net += amounts.net;
    acc.invoiceNet += inv.net;
    acc.unassignedTotal += amounts.unassignedTotal;
  }

  function finalize(acc) {
    for (var g in acc.groups) acc.groups[g] = round2(acc.groups[g]);
    ['revenue', 'discount', 'other', 'utility', 'exclude', 'net', 'invoiceNet', 'unassignedTotal'].forEach(function (k) {
      acc[k] = round2(acc[k]);
    });
    return acc;
  }

  /** Numeric rooms first (by number), then text rooms such as "(ไม่ระบุ)". */
  function sortRooms(a, b) {
    var na = parseInt(a, 10), nb = parseInt(b, 10);
    var an = !isNaN(na), bn = !isNaN(nb);
    if (an && bn && na !== nb) return na - nb;
    if (an !== bn) return an ? -1 : 1;
    return a < b ? -1 : a > b ? 1 : 0;
  }

  /**
   * rows = one per room (sorted numerically), each:
   *   { room, count, groups:{[groupId]:sum}, revenue, discount, other, utility, exclude, net, invoiceNet, unassignedTotal }
   * totals = same shape with room = 'รวม'. columns = groupColumns(groups).
   */
  function roomByGroup(invoices, colMap, groups) {
    var acc = {};
    var totals = newAccumulator('รวม');
    invoices.forEach(function (inv) {
      var a = invoiceGroupAmounts(inv, colMap, groups);
      if (!acc[inv.room]) acc[inv.room] = newAccumulator(inv.room);
      accumulate(acc[inv.room], a, inv);
      accumulate(totals, a, inv);
    });
    var rows = Object.keys(acc).sort(sortRooms).map(function (room) {
      var r = finalize(acc[room]);
      r.room = room;
      delete r.key;
      return r;
    });
    var t = finalize(totals);
    t.room = 'รวม';
    delete t.key;
    return { rows: rows, totals: t, columns: groupColumns(groups) };
  }

  // ---------------------------------------------------------------------------
  // Monthly pivot
  // ---------------------------------------------------------------------------

  /**
   * Month columns: the 12 months of `year`, or (year = all) every calendar
   * month from the first to the last month with data, gaps included so vacant
   * months show as zero. `range` = { min, max } ('YYYY-MM' or 'YYYY-MM-DD',
   * e.g. report.dateRange) widens the span so the axis stays the same across
   * room / status / customer filters.
   */
  function monthColumns(invoices, year, range) {
    if (year !== null && year !== undefined && year !== 'all' && year !== '') {
      var y = +year;
      return MONTHS_TH.map(function (label, i) {
        return { key: y + '-' + pad2(i + 1), label: label, year: y, month: i + 1 };
      });
    }
    var from = null, to = null;
    function widen(ym) {
      if (!ym) return;
      ym = String(ym).slice(0, 7);
      if (!from || ym < from) from = ym;
      if (!to || ym > to) to = ym;
    }
    invoices.forEach(function (inv) { widen(inv.ym); });
    if (range) { widen(range.min); widen(range.max); }
    if (!from) return [];
    var cols = [];
    var cy = +from.slice(0, 4), cm = +from.slice(5, 7);
    var ey = +to.slice(0, 4), em = +to.slice(5, 7);
    while ((cy < ey || (cy === ey && cm <= em)) && cols.length < 1200) {
      cols.push({ key: cy + '-' + pad2(cm), label: MONTHS_TH[cm - 1] + ' ' + cy, year: cy, month: cm });
      if (++cm > 12) { cm = 1; cy++; }
    }
    return cols;
  }

  /**
   * Excel-style pivot: rows = room (optionally room → customer sub-rows),
   * columns = months of `year` (or every month in range when year is null,
   * see monthColumns / opts.range), value = valueKey (groupId | VALUE_KEYS.*).
   *
   * @returns {{
   *   columns: [{key:'2024-01', label:'ม.ค.', year, month}],
   *   rows: [{ room, customer?: string, level: 0|1, values: number[], total }],
   *   grandTotal: { values: number[], total },
   *   valueKey, valueLabel
   * }}
   * Room rows are level 0; when byCustomer, each room row is followed by its
   * customer rows (level 1) and the room row holds the room subtotal.
   * Customer rows are ordered by the date of the customer's first invoice in that
   * room (oldest tenant first), taken from opts.allInvoices when given (so the
   * order does not change with filters) else from `invoices`; ties sort by name.
   */
  function pivotMonthly(invoices, colMap, opts) {
    opts = opts || {};
    var groups = opts.groups || [];
    var valueKey = opts.valueKey || VALUE_KEYS.REVENUE;
    var columns = monthColumns(invoices, opts.year, opts.range);
    var colIdx = {};
    columns.forEach(function (c, i) { colIdx[c.key] = i; });
    var n = columns.length;

    var firstDate = {};
    (opts.allInvoices || invoices).forEach(function (inv) {
      var k = inv.room + '\u0000' + inv.customer;
      if (firstDate[k] === undefined || inv.date < firstDate[k]) firstDate[k] = inv.date;
    });

    var rooms = {};
    var grand = new Array(n).fill(0);
    invoices.forEach(function (inv) {
      var ci = colIdx[inv.ym];
      if (ci === undefined) return;
      var v = valueOf(inv, colMap, groups, valueKey);
      var r = rooms[inv.room];
      if (!r) r = rooms[inv.room] = { values: new Array(n).fill(0), customers: {} };
      r.values[ci] += v;
      grand[ci] += v;
      if (opts.byCustomer) {
        var c = r.customers[inv.customer];
        if (!c) c = r.customers[inv.customer] = new Array(n).fill(0);
        c[ci] += v;
      }
    });

    function sum(arr) { return round2(arr.reduce(function (s, x) { return s + x; }, 0)); }
    function roundAll(arr) { return arr.map(round2); }

    var rows = [];
    Object.keys(rooms).sort(sortRooms).forEach(function (room) {
      var r = rooms[room];
      rows.push({ room: room, level: 0, values: roundAll(r.values), total: sum(r.values) });
      if (opts.byCustomer) {
        Object.keys(r.customers).sort(function (a, b) {
          var da = firstDate[room + '\u0000' + a] || '', db = firstDate[room + '\u0000' + b] || '';
          return da < db ? -1 : da > db ? 1 : (a < b ? -1 : a > b ? 1 : 0);
        }).forEach(function (cust) {
          rows.push({ room: room, customer: cust, level: 1, values: roundAll(r.customers[cust]), total: sum(r.customers[cust]) });
        });
      }
    });

    return {
      columns: columns,
      rows: rows,
      grandTotal: { values: roundAll(grand), total: sum(grand) },
      valueKey: valueKey,
      valueLabel: valueLabel(valueKey, groups)
    };
  }

  // ---------------------------------------------------------------------------
  // Monthly series (chart)
  // ---------------------------------------------------------------------------

  /**
   * Data for the stacked bar + net line chart.
   * @param opts {year?: number|null, range?: {min, max}}   (see monthColumns)
   * @returns {{
   *   keys: ['2024-01', ...], labels: ['ม.ค.', ...] or ['ม.ค. 2024', ...],
   *   series: [{ id, label, kind, data: number[] }],   // assignable groups of kind revenue/discount/other
   *   net: number[], revenue: number[], discount: number[], utility: number[], count: number[],
   *   unassignedTotal: number[]
   * }}
   */
  function monthlySeries(invoices, colMap, groups, opts) {
    opts = opts || {};
    var gi = groupIndex(groups);
    var columns = monthColumns(invoices, opts.year, opts.range);
    var colIdx = {};
    columns.forEach(function (c, i) { colIdx[c.key] = i; });
    var n = columns.length;
    var seriesGroups = gi.assignable.filter(function (g) {
      return g.kind === 'revenue' || g.kind === 'discount' || g.kind === 'other';
    });
    var series = seriesGroups.map(function (g) {
      return { id: g.id, label: g.label, kind: g.kind, data: new Array(n).fill(0) };
    });
    var sIdx = {};
    series.forEach(function (s, i) { sIdx[s.id] = i; });
    var net = new Array(n).fill(0), revenue = new Array(n).fill(0), discount = new Array(n).fill(0);
    var utility = new Array(n).fill(0), count = new Array(n).fill(0), unassigned = new Array(n).fill(0);

    invoices.forEach(function (inv) {
      var ci = colIdx[inv.ym];
      if (ci === undefined) return;
      var a = invoiceGroupAmounts(inv, colMap, groups);
      for (var gid in a.groups) {
        var si = sIdx[gid];
        if (si !== undefined) series[si].data[ci] += a.groups[gid];
      }
      net[ci] += a.net;
      revenue[ci] += a.revenue;
      discount[ci] += a.discount;
      utility[ci] += a.utility;
      unassigned[ci] += a.unassignedTotal;
      count[ci] += 1;
    });

    series.forEach(function (s) { s.data = s.data.map(round2); });
    return {
      keys: columns.map(function (c) { return c.key; }),
      labels: columns.map(function (c) { return c.label; }),
      columns: columns,
      series: series,
      net: net.map(round2),
      revenue: revenue.map(round2),
      discount: discount.map(round2),
      utility: utility.map(round2),
      unassignedTotal: unassigned.map(round2),
      count: count
    };
  }

  // ---------------------------------------------------------------------------
  // Metrics row
  // ---------------------------------------------------------------------------

  /**
   * @returns {{ revenue, discount, other, net, utility, exclude, unpaid, unpaidCount, count,
   *             invoiceNet, unassignedTotal, groups: {[groupId]: sum} }}
   * unpaid = sum of invoice `net` where status === 'ค้างชำระ'.
   */
  function metrics(invoices, colMap, groups) {
    var acc = newAccumulator('all');
    var unpaid = 0, unpaidCount = 0;
    invoices.forEach(function (inv) {
      accumulate(acc, invoiceGroupAmounts(inv, colMap, groups), inv);
      if (inv.status === STATUS_UNPAID) { unpaid += inv.net; unpaidCount += 1; }
    });
    finalize(acc);
    return {
      revenue: acc.revenue,
      discount: acc.discount,
      other: acc.other,
      net: acc.net,
      utility: acc.utility,
      exclude: acc.exclude,
      unpaid: round2(unpaid),
      unpaidCount: unpaidCount,
      count: acc.count,
      invoiceNet: acc.invoiceNet,
      unassignedTotal: acc.unassignedTotal,
      groups: acc.groups
    };
  }

  // ---------------------------------------------------------------------------
  // Invoice detail rows
  // ---------------------------------------------------------------------------

  /**
   * Flat rows for the รายละเอียดบิล table.
   * @returns [{ date, docNo, room, roomDesc, customer, status, groups:{[groupId]:sum}, revenue, discount, other, utility, netRevenue, net, unassignedTotal }]
   */
  function invoiceRows(invoices, colMap, groups) {
    return invoices.map(function (inv) {
      var a = invoiceGroupAmounts(inv, colMap, groups);
      return {
        date: inv.date, ym: inv.ym, docNo: inv.docNo, room: inv.room, roomDesc: inv.roomDesc,
        customer: inv.customer, status: inv.status,
        groups: a.groups, revenue: a.revenue, discount: a.discount, other: a.other, utility: a.utility,
        exclude: a.exclude, netRevenue: a.net, net: inv.net, total: inv.total, vat: inv.vat,
        unassignedTotal: a.unassignedTotal
      };
    });
  }

  /** Distinct helpers for filter widgets. */
  function distinctYears(invoices) {
    var s = {};
    invoices.forEach(function (inv) { s[inv.year] = true; });
    return Object.keys(s).map(Number).sort(function (a, b) { return a - b; });
  }

  function distinctRooms(invoices) {
    var s = {};
    invoices.forEach(function (inv) { s[inv.room] = true; });
    return Object.keys(s).sort(sortRooms);
  }

  function distinctStatuses(invoices) {
    var s = {};
    invoices.forEach(function (inv) { s[inv.status] = (s[inv.status] || 0) + 1; });
    return Object.keys(s).map(function (k) { return { status: k, count: s[k] }; })
      .sort(function (a, b) { return b.count - a.count; });
  }

  /**
   * Column-level statistics with grouping info for the mapping table:
   * [{ raw, normalized, group, groupLabel, state, kind, count, total, signHint }]
   * signHint = true when total sign disagrees with the group kind.
   */
  function columnSummary(itemHeaders, columnStats, settings) {
    var S = global.JMP.settings;
    var gi = groupIndex(settings.groups);
    return itemHeaders.map(function (raw) {
      var normalized = S.normalizeHeader(raw);
      var c = S.classify(normalized, settings);
      var st = columnStats[raw] || { count: 0, total: 0 };
      var g = c.group ? gi.byId[c.group] : null;
      var kind = g ? g.kind : (c.group ? gi.kindOf(c.group) : null);
      var signHint = (kind === 'revenue' && st.total < 0) || (kind === 'discount' && st.total > 0);
      return { raw: raw, normalized: normalized, group: c.group, groupLabel: g ? g.label : '', state: c.state,
               rule: c.rule || null, kind: kind, count: st.count, total: st.total, signHint: signHint };
    });
  }

  global.JMP.aggregate = {
    MONTHS_TH: MONTHS_TH,
    VALUE_KEYS: VALUE_KEYS,
    VALUE_LABELS: VALUE_LABELS,
    UNASSIGNED_GROUP: UNASSIGNED_GROUP,
    STATUS_UNPAID: STATUS_UNPAID,
    round2: round2,
    groupIndex: groupIndex,
    groupColumns: groupColumns,
    monthColumns: monthColumns,
    applyFilters: applyFilters,
    invoiceGroupAmounts: invoiceGroupAmounts,
    valueOf: valueOf,
    valueOptions: valueOptions,
    valueLabel: valueLabel,
    roomByGroup: roomByGroup,
    pivotMonthly: pivotMonthly,
    monthlySeries: monthlySeries,
    metrics: metrics,
    invoiceRows: invoiceRows,
    distinctYears: distinctYears,
    distinctRooms: distinctRooms,
    distinctStatuses: distinctStatuses,
    columnSummary: columnSummary
  };
})(typeof window !== 'undefined' ? window : globalThis);
