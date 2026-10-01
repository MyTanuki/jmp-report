window.JMP = window.JMP || {};

/*
 * JMP.settings
 * ------------
 * Default groups / rules, header normalization, the rule engine (classify),
 * persistence (localStorage + JSON export/import with shape validation) and
 * group helpers used by the mapping screen and the settings screen.
 *
 * Plain script (no modules). Works in the browser and inside a node `vm`
 * context whose global exposes `window`.
 */
(function (global) {
  'use strict';

  var STORAGE_KEY = 'jmp-report.settings.v1';
  var SETTINGS_VERSION = 1;
  var KINDS = ['revenue', 'discount', 'utility', 'other', 'exclude'];
  // Group that receives columns without a mapping (see JMP.aggregate). It must
  // always exist, cannot be deleted and cannot be used as a parent.
  var UNASSIGNED_GROUP = 'other';

  var DEFAULT_GROUPS = [
    { id: 'room_rent',        label: 'ค่าเช่าห้อง',                                   kind: 'revenue',  order: 10 },
    { id: 'furniture',        label: 'ค่าเช่าเฟอร์นิเจอร์',                            kind: 'revenue',  order: 20 },
    { id: 'common_utility',   label: 'ค่าสาธารณูปโภคส่วนกลาง',                         kind: 'revenue',  order: 30 },
    { id: 'aircon',           label: 'ค่าเช่าเครื่องปรับอากาศ',                        kind: 'revenue',  order: 40 },
    { id: 'tv',               label: 'ค่าเช่าโทรทัศน์',                               kind: 'revenue',  order: 50 },
    { id: 'fridge',           label: 'ค่าเช่าตู้เย็น',                                kind: 'revenue',  order: 60 },
    { id: 'discount',         label: 'ส่วนลด',                                        kind: 'discount', order: 100 },
    { id: 'discount_new',     label: 'ส่วนลดผู้เช่าใหม่',                              kind: 'discount', parent: 'discount', order: 110 },
    { id: 'discount_special', label: 'ส่วนลดพิเศษ',                                   kind: 'discount', parent: 'discount', order: 120 },
    { id: 'discount_days',    label: 'ส่วนลดตามจำนวนวัน (เข้าพักไม่ตรงต้นเดือน)',       kind: 'discount', parent: 'discount', order: 130 },
    { id: 'discount_fridge',  label: 'ส่วนลดค่าเช่าตู้เย็น',                           kind: 'discount', parent: 'discount', order: 140 },
    { id: 'discount_tv',      label: 'ส่วนลดค่าเช่าโทรทัศน์',                          kind: 'discount', parent: 'discount', order: 150 },
    { id: 'discount_other',   label: 'ส่วนลดอื่นๆ',                                   kind: 'discount', parent: 'discount', order: 160 },
    { id: 'other',            label: 'อื่นๆ',                                         kind: 'other',    order: 200 },
    { id: 'water',            label: 'ค่าน้ำ',                                        kind: 'utility',  order: 300 },
    { id: 'electric',         label: 'ค่าไฟฟ้า',                                      kind: 'utility',  order: 310 },
    { id: 'exclude',          label: 'ไม่นับ',                                        kind: 'exclude',  order: 400 }
  ];

  // Ordered; first match wins. Patterns are matched against the normalized header.
  var DEFAULT_RULES = [
    { pattern: '^ค่าน้ำ|water rate',                                          group: 'water' },
    { pattern: '^ค่าไฟ|electrical rate',                                      group: 'electric' },
    { pattern: 'ส่วนลด.*ตู้เย็น',                                             group: 'discount_fridge' },
    { pattern: 'ส่วนลด.*(โทรทัศน์|ทีวี)',                                      group: 'discount_tv' },
    { pattern: '(หัก|ส่วนลด).*(\\d+\\s*วัน|ครึ่งเดือน|\\d+\\s*-\\s*\\d+)',      group: 'discount_days' },
    { pattern: 'ส่วนลด(ผู้เช่า|ลูกค้า)ใหม่|new\\s*cu?s?t?omer',                 group: 'discount_new' },
    { pattern: 'ส่วนลดพิเศษ',                                                 group: 'discount_special' },
    // Prepayments / bad debt must win over the generic "^หัก" discount rule.
    { pattern: 'ชำระล่วงหน้า|จ่ายล่วงหน้า|หนี้สูญ|ค้างชำระ',                    group: 'exclude' },
    { pattern: '^ส่วนลด|^หัก',                                                 group: 'discount_other' },
    { pattern: '^ค่าเช่าห้อง|room\\s*rate',                                    group: 'room_rent' },
    { pattern: 'เฟอร์นิเจอร์|furniture',                                       group: 'furniture' },
    { pattern: 'สาธาร.*ส่วนกลาง',                                              group: 'common_utility' },
    { pattern: 'เครื่องปรับอากาศ',                                              group: 'aircon' },
    { pattern: 'โทรทัศน์|\\btv\\b',                                            group: 'tv' },
    { pattern: 'ตู้เย็น|refrigerator',                                          group: 'fridge' }
  ];

  var DEFAULT_STATUS_INCLUDE = ['ชำระเงินแล้ว', 'ค้างชำระ', 'ยืนยันชำระเงินโดยพนักงาน'];

  var DEFAULT_SETTINGS = {
    version: SETTINGS_VERSION,
    groups: DEFAULT_GROUPS,
    rules: DEFAULT_RULES,
    columns: {},
    statusInclude: DEFAULT_STATUS_INCLUDE
  };

  function clone(obj) {
    return JSON.parse(JSON.stringify(obj));
  }

  /**
   * Defaults shipped with the app in js/default-settings.js
   * (window.JMP_DEFAULT_SETTINGS). Returns a validated copy, or null when the
   * file is absent or invalid (the built-in defaults are used instead).
   */
  function shippedDefaults() {
    var raw = global.JMP_DEFAULT_SETTINGS;
    if (raw === undefined || raw === null) return null;
    var v = validate(raw);
    if (v.ok) return v.settings;
    if (global.console && typeof global.console.warn === 'function') {
      global.console.warn('js/default-settings.js ไม่ถูกต้อง จึงใช้ค่าเริ่มต้นในตัว: ' + v.errors.join('; '));
    }
    return null;
  }

  /**
   * Settings for a first run (nothing in localStorage) and for reset():
   * the shipped defaults when present and valid, else the built-in ones.
   */
  function defaults() {
    return shippedDefaults() || clone(DEFAULT_SETTINGS);
  }

  /**
   * Source text of js/default-settings.js holding `settings` as the shipped
   * defaults. A static page cannot write to the repo, so the settings screen
   * offers this as a download that replaces the file.
   */
  function defaultSettingsJS(settings, savedAt) {
    return '/*\n' +
      ' * Default settings shipped with the app: groups, rules, confirmed column\n' +
      ' * mappings and the default status filter. Used when the browser has no saved\n' +
      ' * settings yet, and by "คืนค่าเริ่มต้น" on the settings screen.\n' +
      ' *\n' +
      ' * Generated by the settings screen ("บันทึกเป็นค่าเริ่มต้น") — replace this file\n' +
      ' * with the downloaded one and commit. Deleting it falls back to the built-in\n' +
      ' * defaults in js/settings.js.\n' +
      (savedAt ? ' *\n * Saved: ' + String(savedAt).replace(/[*\/\r\n]/g, '') + '\n' : '') +
      ' */\n' +
      'window.JMP_DEFAULT_SETTINGS = ' + JSON.stringify(settings || defaults(), null, 2) + ';\n';
  }

  // ---------------------------------------------------------------------------
  // Header normalization
  // ---------------------------------------------------------------------------

  var NIKHAHIT_SARA_AA = /ํา/g; // "ํา" typed as two code points
  var SARA_AM = 'ำ';                 // "ำ"

  /**
   * Normalize an item column header so that spelling variants share one key.
   * 1. Unicode NFC  2. ํา → ำ  3. strip trailing "(non-vat)"  4. collapse whitespace.
   */
  function normalizeHeader(raw) {
    if (raw === null || raw === undefined) return '';
    var s = String(raw);
    if (typeof s.normalize === 'function') s = s.normalize('NFC');
    s = s.replace(NIKHAHIT_SARA_AA, SARA_AM);
    s = s.replace(/\s*\(\s*non-vat\s*\)\s*$/i, '');
    s = s.replace(/\s+/g, ' ').trim();
    return s;
  }

  /** Same Unicode fixes as normalizeHeader but for regex patterns (no stripping). */
  function normalizePattern(pattern) {
    var s = String(pattern);
    if (typeof s.normalize === 'function') s = s.normalize('NFC');
    return s.replace(NIKHAHIT_SARA_AA, SARA_AM);
  }

  // ---------------------------------------------------------------------------
  // Rule engine
  // ---------------------------------------------------------------------------

  /** Compile a rule into a RegExp; returns null when the pattern is invalid. */
  function compileRule(rule) {
    if (!rule || typeof rule.pattern !== 'string' || rule.pattern === '') return null;
    var flags = typeof rule.flags === 'string' ? rule.flags : 'i';
    try {
      return new RegExp(normalizePattern(rule.pattern), flags);
    } catch (e) {
      return null;
    }
  }

  /**
   * Decide the group of one normalized header.
   * Explicit settings.columns wins ("confirmed"); otherwise the first matching
   * rule ("suggested"); otherwise "unassigned".
   *
   * @returns {{group: string|null, state: 'confirmed'|'suggested'|'unassigned', rule?: {index:number, pattern:string, group:string}}}
   */
  function classify(normalizedHeader, settings) {
    settings = settings || defaults();
    var key = normalizedHeader;
    var columns = settings.columns || {};
    if (Object.prototype.hasOwnProperty.call(columns, key) && columns[key]) {
      if (groupById(settings, columns[key])) {
        return { group: columns[key], state: 'confirmed' };
      }
    }
    var rules = settings.rules || [];
    for (var i = 0; i < rules.length; i++) {
      var re = compileRule(rules[i]);
      if (!re) continue;
      re.lastIndex = 0;
      if (re.test(key) && groupById(settings, rules[i].group)) {
        return {
          group: rules[i].group,
          state: 'suggested',
          rule: { index: i, pattern: rules[i].pattern, group: rules[i].group }
        };
      }
    }
    return { group: null, state: 'unassigned' };
  }

  /**
   * Classify a list of raw headers.
   * @returns {Array<{raw:string, normalized:string, group:string|null, state:string, rule?:object}>}
   */
  function classifyAll(rawHeaders, settings) {
    return (rawHeaders || []).map(function (raw) {
      var normalized = normalizeHeader(raw);
      var c = classify(normalized, settings);
      return { raw: raw, normalized: normalized, group: c.group, state: c.state, rule: c.rule };
    });
  }

  /**
   * Build the column map consumed by JMP.aggregate: { [rawHeader]: groupId | null }.
   */
  function columnMap(rawHeaders, settings) {
    var map = {};
    (rawHeaders || []).forEach(function (raw) {
      map[raw] = classify(normalizeHeader(raw), settings).group;
    });
    return map;
  }

  /** Test one rule against a list of raw headers (settings screen "ทดสอบ"). */
  function testRule(rule, rawHeaders) {
    var re = compileRule(rule);
    if (!re) return { valid: false, matches: [] };
    var matches = [];
    (rawHeaders || []).forEach(function (raw) {
      var n = normalizeHeader(raw);
      re.lastIndex = 0;
      if (re.test(n)) matches.push(n);
    });
    return { valid: true, matches: matches };
  }

  // ---------------------------------------------------------------------------
  // Validation
  // ---------------------------------------------------------------------------

  function isPlainObject(v) {
    return v !== null && typeof v === 'object' && !Array.isArray(v);
  }

  /**
   * Check whether `parentId` may be the parent of group `id`. Only one level of
   * nesting is supported (the UI offers top-level groups as parents), so a parent
   * must itself be top-level; this also rules out cycles, which are reported
   * with their own message. `lookup(id)` returns the group or null.
   * @returns {string|null} Thai error message, or null when valid
   */
  function parentError(lookup, id, parentId) {
    if (parentId === id) return 'กลุ่ม ' + id + ' เป็น parent ของตัวเอง';
    var p = lookup(parentId);
    if (!p) return 'กลุ่ม ' + id + ' อ้างถึง parent ที่ไม่มี: ' + parentId;
    if (parentId === UNASSIGNED_GROUP) return 'กลุ่ม "' + UNASSIGNED_GROUP + '" ใช้เป็นกลุ่มแม่ไม่ได้';
    var seen = {};
    seen[id] = true;
    var cur = p, guard = 0;
    while (cur && guard++ < 1000) {
      if (seen[cur.id]) return 'กลุ่ม ' + id + ' มี parent วนซ้ำ: ' + cur.id;
      seen[cur.id] = true;
      cur = cur.parent ? lookup(cur.parent) : null;
    }
    if (p.parent) return 'กลุ่ม ' + id + ' อ้างถึง parent ที่ไม่ใช่กลุ่มระดับบนสุด: ' + parentId;
    return null;
  }

  /**
   * Validate a parsed settings object. Returns { ok, errors, settings } where
   * `settings` is a cleaned copy (only when ok).
   */
  function validate(obj) {
    var errors = [];
    if (!isPlainObject(obj)) return { ok: false, errors: ['ไฟล์ตั้งค่าต้องเป็น JSON object'] };

    if (!Array.isArray(obj.groups) || obj.groups.length === 0) errors.push('ไม่พบรายการกลุ่ม (groups)');
    if (!Array.isArray(obj.rules)) errors.push('ไม่พบรายการกฎ (rules)');
    if (obj.columns !== undefined && !isPlainObject(obj.columns)) errors.push('columns ต้องเป็น object');
    if (obj.statusInclude !== undefined && !Array.isArray(obj.statusInclude)) errors.push('statusInclude ต้องเป็น array');
    if (errors.length) return { ok: false, errors: errors };

    var ids = {};
    var groups = obj.groups.map(function (g, i) {
      if (!isPlainObject(g) || typeof g.id !== 'string' || g.id === '') {
        errors.push('กลุ่มลำดับที่ ' + (i + 1) + ' ไม่มี id');
        return null;
      }
      if (ids[g.id]) errors.push('id กลุ่มซ้ำ: ' + g.id);
      ids[g.id] = true;
      if (KINDS.indexOf(g.kind) < 0) errors.push('กลุ่ม ' + g.id + ' มี kind ไม่ถูกต้อง: ' + g.kind);
      var out = {
        id: g.id,
        label: typeof g.label === 'string' && g.label !== '' ? g.label : g.id,
        kind: g.kind
      };
      if (typeof g.parent === 'string' && g.parent !== '') out.parent = g.parent;
      out.order = typeof g.order === 'number' ? g.order : (i + 1) * 10;
      return out;
    });
    var byId = {};
    groups.forEach(function (g) { if (g) byId[g.id] = g; });
    var lookup = function (id) { return byId[id] || null; };
    var parents = {};
    groups.forEach(function (g) {
      if (!g || !g.parent) return;
      var err = parentError(lookup, g.id, g.parent);
      if (err) errors.push(err); else parents[g.parent] = true;
    });
    // The unassigned bucket must exist; add the default one when it is missing.
    if (!ids[UNASSIGNED_GROUP]) {
      var maxOrder = groups.reduce(function (m, g) { return g ? Math.max(m, g.order) : m; }, 0);
      groups.push({ id: UNASSIGNED_GROUP, label: 'อื่นๆ', kind: 'other', order: maxOrder + 10 });
      ids[UNASSIGNED_GROUP] = true;
    }

    var rules = obj.rules.map(function (r, i) {
      if (!isPlainObject(r) || typeof r.pattern !== 'string' || typeof r.group !== 'string') {
        errors.push('กฎลำดับที่ ' + (i + 1) + ' ต้องมี pattern และ group');
        return null;
      }
      if (!ids[r.group]) errors.push('กฎลำดับที่ ' + (i + 1) + ' อ้างถึงกลุ่มที่ไม่มี: ' + r.group);
      else if (parents[r.group]) errors.push('กฎลำดับที่ ' + (i + 1) + ' อ้างถึงกลุ่มแม่ (จับคู่ได้เฉพาะกลุ่มย่อย): ' + r.group);
      if (!compileRule(r)) errors.push('กฎลำดับที่ ' + (i + 1) + ' เป็น regex ที่ไม่ถูกต้อง: ' + r.pattern);
      var out = { pattern: r.pattern, group: r.group };
      if (typeof r.flags === 'string') out.flags = r.flags;
      return out;
    });

    var columns = {};
    if (obj.columns) {
      Object.keys(obj.columns).forEach(function (k) {
        var v = obj.columns[k];
        if (typeof v !== 'string' || v === '') return; // ignore empty mapping
        if (!ids[v]) { errors.push('คอลัมน์ "' + k + '" อ้างถึงกลุ่มที่ไม่มี: ' + v); return; }
        if (parents[v]) { errors.push('คอลัมน์ "' + k + '" อ้างถึงกลุ่มแม่ (จับคู่ได้เฉพาะกลุ่มย่อย): ' + v); return; }
        columns[normalizeHeader(k)] = v;
      });
    }

    var statusInclude = Array.isArray(obj.statusInclude)
      ? obj.statusInclude.filter(function (s) { return typeof s === 'string' && s !== ''; })
      : clone(DEFAULT_STATUS_INCLUDE);

    if (errors.length) return { ok: false, errors: errors };
    return {
      ok: true,
      errors: [],
      settings: {
        version: SETTINGS_VERSION,
        groups: groups,
        rules: rules,
        columns: columns,
        statusInclude: statusInclude
      }
    };
  }

  // ---------------------------------------------------------------------------
  // Persistence
  // ---------------------------------------------------------------------------

  function storage() {
    try {
      var ls = global.localStorage;
      if (ls && typeof ls.getItem === 'function') return ls;
    } catch (e) { /* access denied (privacy mode) */ }
    return null;
  }

  /**
   * Load settings from localStorage; falls back to defaults() (shipped
   * js/default-settings.js, else built-in) when missing/invalid.
   */
  function load() {
    var ls = storage();
    if (!ls) return defaults();
    var text = null;
    try { text = ls.getItem(STORAGE_KEY); } catch (e) { return defaults(); }
    if (!text) return defaults();
    try {
      var v = validate(JSON.parse(text));
      if (v.ok) return v.settings;
    } catch (e) { /* corrupted */ }
    return defaults();
  }

  /** Persist settings. Returns true when written. */
  function save(settings) {
    var ls = storage();
    if (!ls) return false;
    try {
      ls.setItem(STORAGE_KEY, exportJSON(settings));
      return true;
    } catch (e) {
      return false;
    }
  }

  /** Reset to defaults (also persists). Returns the fresh settings. */
  function reset() {
    var s = defaults();
    save(s);
    return s;
  }

  function exportJSON(settings) {
    return JSON.stringify(settings || defaults(), null, 2);
  }

  /**
   * Parse and validate a settings JSON text. Throws Error (Thai message) when invalid.
   */
  function importJSON(text) {
    var obj;
    try {
      obj = JSON.parse(text);
    } catch (e) {
      throw new Error('ไฟล์ไม่ใช่ JSON ที่ถูกต้อง');
    }
    var v = validate(obj);
    if (!v.ok) throw new Error('รูปแบบไฟล์ตั้งค่าไม่ถูกต้อง: ' + v.errors.join('; '));
    return v.settings;
  }

  // ---------------------------------------------------------------------------
  // Group helpers (mutate the settings object passed in; caller saves)
  // ---------------------------------------------------------------------------

  function sortedGroups(settings) {
    return (settings.groups || []).slice().sort(function (a, b) {
      return (a.order || 0) - (b.order || 0);
    });
  }

  function groupById(settings, id) {
    var gs = (settings && settings.groups) || [];
    for (var i = 0; i < gs.length; i++) if (gs[i].id === id) return gs[i];
    return null;
  }

  function isParent(settings, id) {
    return (settings.groups || []).some(function (g) { return g.parent === id; });
  }

  function childGroups(settings, parentId) {
    return sortedGroups(settings).filter(function (g) { return g.parent === parentId; });
  }

  /** Groups a column may be mapped to (parents that have children are containers only). */
  function assignableGroups(settings) {
    return sortedGroups(settings).filter(function (g) { return !isParent(settings, g.id); });
  }

  /** Derive a safe ascii-ish id from a label; guarantees uniqueness within settings. */
  function makeGroupId(settings, label) {
    var base = String(label || 'group').toLowerCase()
      .replace(/[^a-z0-9]+/g, '_').replace(/^_+|_+$/g, '');
    if (!base) base = 'group';
    var id = base, n = 2;
    while (groupById(settings, id)) id = base + '_' + (n++);
    return id;
  }

  /** Columns and rules currently targeting a group. */
  function referencesTo(settings, id) {
    var columns = Object.keys(settings.columns || {}).filter(function (k) { return settings.columns[k] === id; });
    var rules = [];
    (settings.rules || []).forEach(function (r, i) { if (r.group === id) rules.push(i); });
    return { columns: columns, rules: rules };
  }

  /**
   * A group that gains its first child becomes a container and is no longer
   * assignable, so its confirmed columns and rules move to that child.
   */
  function retargetReferences(settings, fromId, toId) {
    var refs = referencesTo(settings, fromId);
    refs.columns.forEach(function (k) { settings.columns[k] = toId; });
    refs.rules.forEach(function (i) { settings.rules[i].group = toId; });
    return refs;
  }

  function addGroup(settings, def) {
    def = def || {};
    if (KINDS.indexOf(def.kind) < 0) throw new Error('kind ไม่ถูกต้อง');
    var id = def.id && !groupById(settings, def.id) ? def.id : makeGroupId(settings, def.id || def.label);
    if (def.parent) {
      var err = parentError(function (pid) { return groupById(settings, pid); }, id, def.parent);
      if (err) throw new Error(err);
    }
    var maxOrder = settings.groups.reduce(function (m, g) { return Math.max(m, g.order || 0); }, 0);
    var g = {
      id: id,
      label: def.label || id,
      kind: def.kind,
      order: typeof def.order === 'number' ? def.order : maxOrder + 10
    };
    if (def.parent) {
      var wasParent = isParent(settings, def.parent);
      g.parent = def.parent;
      if (!wasParent) retargetReferences(settings, def.parent, id);
    }
    settings.groups.push(g);
    return g;
  }

  function renameGroup(settings, id, label) {
    var g = groupById(settings, id);
    if (!g) return null;
    g.label = String(label || '').trim() || g.label;
    return g;
  }

  /** Patch kind / parent / order / label of a group. parent: null removes it. */
  function updateGroup(settings, id, patch) {
    var g = groupById(settings, id);
    if (!g) return null;
    patch = patch || {};
    if (patch.label !== undefined) renameGroup(settings, id, patch.label);
    if (patch.kind !== undefined) {
      if (KINDS.indexOf(patch.kind) < 0) throw new Error('kind ไม่ถูกต้อง');
      g.kind = patch.kind;
    }
    if (patch.parent !== undefined) {
      if (patch.parent === null || patch.parent === '') {
        delete g.parent;
      } else if (patch.parent !== g.parent) {
        var err = parentError(function (pid) { return groupById(settings, pid); }, id, patch.parent);
        if (err) throw new Error(err);
        if (isParent(settings, id)) throw new Error('กลุ่ม ' + id + ' มีกลุ่มย่อยจึงมี parent ไม่ได้');
        var wasParent = isParent(settings, patch.parent);
        g.parent = patch.parent;
        if (!wasParent) retargetReferences(settings, patch.parent, id);
      }
    }
    if (typeof patch.order === 'number') g.order = patch.order;
    return g;
  }

  /**
   * Delete a group. Columns confirmed to it become unassigned, rules targeting
   * it are removed, children lose their parent.
   * @returns {{removedColumns: string[], removedRules: number}}
   */
  function deleteGroup(settings, id) {
    if (id === UNASSIGNED_GROUP) throw new Error('ลบกลุ่ม "' + groupLabelOf(settings, id) + '" ไม่ได้ — ใช้รับคอลัมน์ที่ยังไม่กำหนดกลุ่ม');
    var removedColumns = [];
    Object.keys(settings.columns || {}).forEach(function (k) {
      if (settings.columns[k] === id) { delete settings.columns[k]; removedColumns.push(k); }
    });
    var before = settings.rules.length;
    settings.rules = settings.rules.filter(function (r) { return r.group !== id; });
    settings.groups.forEach(function (g) { if (g.parent === id) delete g.parent; });
    settings.groups = settings.groups.filter(function (g) { return g.id !== id; });
    return { removedColumns: removedColumns, removedRules: before - settings.rules.length };
  }

  /** Move a group up (-1) or down (+1) among its siblings (same parent). Renumbers `order`. */
  function moveGroup(settings, id, delta) {
    var g = groupById(settings, id);
    if (!g) return false;
    var all = sortedGroups(settings);
    // Imported settings may carry duplicate orders; swapping equal values would
    // be a no-op, so renumber (keeping the current order) before swapping.
    var tied = all.some(function (x, i) { return i > 0 && (x.order || 0) === (all[i - 1].order || 0); });
    if (tied) all.forEach(function (x, i) { x.order = (i + 1) * 10; });
    var siblings = all.filter(function (x) { return (x.parent || null) === (g.parent || null); });
    var idx = siblings.indexOf(g);
    var target = idx + (delta < 0 ? -1 : 1);
    if (target < 0 || target >= siblings.length) return false;
    var other = siblings[target];
    var tmp = g.order; g.order = other.order; other.order = tmp;
    return true;
  }

  function groupLabelOf(settings, id) {
    var g = groupById(settings, id);
    return g ? g.label : id;
  }

  /** Throw when `groupId` cannot receive columns/rules (unknown or a container). */
  function assertAssignable(settings, groupId) {
    if (!groupById(settings, groupId)) throw new Error('ไม่พบกลุ่ม ' + groupId);
    if (isParent(settings, groupId)) throw new Error('กลุ่ม "' + groupLabelOf(settings, groupId) + '" เป็นกลุ่มแม่ จับคู่ได้เฉพาะกลุ่มย่อย');
  }

  /** Confirm a column mapping (normalized header → group). */
  function setColumn(settings, normalizedHeader, groupId) {
    assertAssignable(settings, groupId);
    settings.columns = settings.columns || {};
    settings.columns[normalizedHeader] = groupId;
  }

  /** Remove a confirmed mapping (column falls back to rules). */
  function clearColumn(settings, normalizedHeader) {
    if (settings.columns) delete settings.columns[normalizedHeader];
  }

  // ---------------------------------------------------------------------------
  // Rule helpers
  // ---------------------------------------------------------------------------

  function addRule(settings, rule, index) {
    if (!compileRule(rule)) throw new Error('regex ไม่ถูกต้อง');
    assertAssignable(settings, rule.group);
    var r = { pattern: rule.pattern, group: rule.group };
    if (typeof rule.flags === 'string') r.flags = rule.flags;
    if (typeof index === 'number' && index >= 0 && index <= settings.rules.length) settings.rules.splice(index, 0, r);
    else settings.rules.push(r);
    return r;
  }

  function updateRule(settings, index, patch) {
    var r = settings.rules[index];
    if (!r) return null;
    var next = { pattern: patch.pattern !== undefined ? patch.pattern : r.pattern,
                 group: patch.group !== undefined ? patch.group : r.group };
    var flags = patch.flags !== undefined ? patch.flags : r.flags;
    if (typeof flags === 'string') next.flags = flags;
    if (!compileRule(next)) throw new Error('regex ไม่ถูกต้อง');
    assertAssignable(settings, next.group);
    settings.rules[index] = next;
    return next;
  }

  function deleteRule(settings, index) {
    if (index < 0 || index >= settings.rules.length) return false;
    settings.rules.splice(index, 1);
    return true;
  }

  function moveRule(settings, index, delta) {
    var target = index + (delta < 0 ? -1 : 1);
    if (index < 0 || index >= settings.rules.length || target < 0 || target >= settings.rules.length) return false;
    var r = settings.rules.splice(index, 1)[0];
    settings.rules.splice(target, 0, r);
    return true;
  }

  global.JMP.settings = {
    STORAGE_KEY: STORAGE_KEY,
    SETTINGS_VERSION: SETTINGS_VERSION,
    KINDS: KINDS,
    UNASSIGNED_GROUP: UNASSIGNED_GROUP,
    DEFAULT_SETTINGS: DEFAULT_SETTINGS,
    defaults: defaults,
    shippedDefaults: shippedDefaults,
    defaultSettingsJS: defaultSettingsJS,
    normalizeHeader: normalizeHeader,
    compileRule: compileRule,
    classify: classify,
    classifyAll: classifyAll,
    columnMap: columnMap,
    testRule: testRule,
    validate: validate,
    load: load,
    save: save,
    reset: reset,
    exportJSON: exportJSON,
    importJSON: importJSON,
    sortedGroups: sortedGroups,
    groupById: groupById,
    isParent: isParent,
    childGroups: childGroups,
    assignableGroups: assignableGroups,
    makeGroupId: makeGroupId,
    referencesTo: referencesTo,
    addGroup: addGroup,
    renameGroup: renameGroup,
    updateGroup: updateGroup,
    deleteGroup: deleteGroup,
    moveGroup: moveGroup,
    setColumn: setColumn,
    clearColumn: clearColumn,
    addRule: addRule,
    updateRule: updateRule,
    deleteRule: deleteRule,
    moveRule: moveRule
  };
})(typeof window !== 'undefined' ? window : globalThis);
