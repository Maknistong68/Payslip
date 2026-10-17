/*
 * Payslip Generator - the safety net. Keeps the current work in IndexedDB so a refresh, a crash or a
 * browser restart loses nothing, keeps the delivery journal append-only so two tabs or a restored
 * session never undo each other's ticks, and remembers whether this computer has seen the app before.
 * The pure functions (journal merge, fingerprints, CSV restore) also run in Node for the tests.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.PayslipVault = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var DB_NAME = 'payslip-generator';
  var STORE = 'kv';
  var VISIT_KEY = 'payslip-generator-visit';
  var BACKUP_KEY = 'payslip-generator-backup';
  var PREFS_KEY = 'payslip-generator-prefs';
  var LS_PREFIX = 'payslip-';

  // ---------------------------------------------------------------------------
  // Key-value store: IndexedDB, or memory when the browser blocks it (private windows, some file:// setups).
  // ---------------------------------------------------------------------------
  var memory = {};
  var dbPromise = null;

  function openDb() {
    if (dbPromise) return dbPromise;
    dbPromise = new Promise(function (resolve) {
      try {
        if (typeof indexedDB === 'undefined') return resolve(null);
        var req = indexedDB.open(DB_NAME, 1);
        req.onupgradeneeded = function () { req.result.createObjectStore(STORE); };
        req.onsuccess = function () { resolve(req.result); };
        req.onerror = function () { resolve(null); };
        req.onblocked = function () { resolve(null); };
      } catch (e) { resolve(null); }
    });
    return dbPromise;
  }

  function tx(mode, fn) {
    return openDb().then(function (db) {
      if (!db) return fn(null);
      return new Promise(function (resolve, reject) {
        var t = db.transaction(STORE, mode);
        var result;
        t.oncomplete = function () { resolve(result); };
        t.onerror = function () { reject(t.error); };
        t.onabort = function () { reject(t.error); };
        result = fn(t.objectStore(STORE));
        if (result && typeof result === 'object' && 'onsuccess' in result) {
          var req = result;
          req.onsuccess = function () { result = req.result; };
        }
      });
    }).catch(function () { return fn(null); });
  }

  function get(key) {
    return tx('readonly', function (store) {
      if (!store) return memory[key];
      return store.get(key);
    });
  }

  function set(key, value) {
    return tx('readwrite', function (store) {
      if (!store) { memory[key] = value; return; }
      store.put(value, key);
    });
  }

  function del(key) {
    return tx('readwrite', function (store) {
      if (!store) { delete memory[key]; return; }
      store.delete(key);
    });
  }

  function keys() {
    return tx('readonly', function (store) {
      if (!store) return Object.keys(memory);
      return store.getAllKeys();
    });
  }

  // Erases everything this app keeps in this browser.
  function eraseAll() {
    memory = {};
    try {
      Object.keys(localStorage).forEach(function (k) { if (k.indexOf(LS_PREFIX) === 0) localStorage.removeItem(k); });
    } catch (e) { /* storage blocked */ }
    return tx('readwrite', function (store) { if (store) store.clear(); });
  }

  // Asks the browser not to evict our data when the disk gets full. Silent in Chrome/Edge.
  function persist() {
    try {
      if (navigator.storage && navigator.storage.persist) {
        return navigator.storage.persisted().then(function (yes) { return yes || navigator.storage.persist(); }).catch(function () { return false; });
      }
    } catch (e) { /* ignore */ }
    return Promise.resolve(false);
  }

  function persisted() {
    try {
      if (navigator.storage && navigator.storage.persisted) return navigator.storage.persisted().catch(function () { return false; });
    } catch (e) { /* ignore */ }
    return Promise.resolve(false);
  }

  // ---------------------------------------------------------------------------
  // Small JSON records in localStorage (synchronous, so they are written before the tab can die).
  // ---------------------------------------------------------------------------
  function readJson(key, fallback) {
    try {
      var raw = localStorage.getItem(key);
      return raw ? JSON.parse(raw) : fallback;
    } catch (e) { return fallback; }
  }

  function writeJson(key, value) {
    try { localStorage.setItem(key, JSON.stringify(value)); return true; } catch (e) { return false; }
  }

  // First visit or returning? Counted once per page load.
  function touchVisit(now) {
    var v = readJson(VISIT_KEY, null);
    var first = !v;
    v = v || { firstSeen: now, visits: 0, tourDone: false };
    v.visits = (v.visits || 0) + 1;
    v.previousSeen = v.lastSeen || null;
    v.lastSeen = now;
    writeJson(VISIT_KEY, v);
    return { first: first, visit: v };
  }

  function visit() { return readJson(VISIT_KEY, {}); }
  function updateVisit(patch) {
    var v = readJson(VISIT_KEY, {});
    Object.keys(patch).forEach(function (k) { v[k] = patch[k]; });
    writeJson(VISIT_KEY, v);
    return v;
  }

  function backup() { return readJson(BACKUP_KEY, {}); }
  function setBackup(rec) { writeJson(BACKUP_KEY, rec); return rec; }

  function prefs() {
    var p = readJson(PREFS_KEY, {});
    if (p.remember == null) p.remember = true;
    return p;
  }
  function setPrefs(patch) {
    var p = prefs();
    Object.keys(patch).forEach(function (k) { p[k] = patch[k]; });
    writeJson(PREFS_KEY, p);
    return p;
  }

  // ---------------------------------------------------------------------------
  // Fingerprints
  // ---------------------------------------------------------------------------
  // JSON with sorted keys, so the same settings always give the same text.
  function stableStringify(v) {
    if (v === null || typeof v !== 'object') return JSON.stringify(v === undefined ? null : v);
    if (Array.isArray(v)) return '[' + v.map(stableStringify).join(',') + ']';
    return '{' + Object.keys(v).sort().filter(function (k) { return v[k] !== undefined; }).map(function (k) {
      return JSON.stringify(k) + ':' + stableStringify(v[k]);
    }).join(',') + '}';
  }

  // FNV-1a, 2 x 32 bit: fast, synchronous, plenty for "did anything change".
  function hashText(s) {
    var h1 = 0x811c9dc5, h2 = 0x01000193 ^ s.length;
    for (var i = 0; i < s.length; i++) {
      var c = s.charCodeAt(i);
      h1 = Math.imul(h1 ^ c, 16777619);
      h2 = Math.imul(h2 ^ c, 2246822519);
    }
    return ('0000000' + (h1 >>> 0).toString(16)).slice(-8) + ('0000000' + (h2 >>> 0).toString(16)).slice(-8);
  }

  // Settings that change every month (period, dates) do not make a backup stale.
  var MONTHLY_KEYS = ['payPeriod', 'periodStart', 'periodEnd', 'paymentDate'];
  function settingsFingerprint(obj) {
    var copy = {};
    Object.keys(obj || {}).forEach(function (k) { if (k !== 'savedAt' && k !== 'version') copy[k] = obj[k]; });
    if (copy.company) {
      var c = {};
      Object.keys(copy.company).forEach(function (k) { if (MONTHLY_KEYS.indexOf(k) < 0 && copy.company[k] !== '' && copy.company[k] != null) c[k] = copy.company[k]; });
      copy.company = c;
    }
    return hashText(stableStringify(copy));
  }

  // ---------------------------------------------------------------------------
  // Delivery journal: append-only. "sent", "undo" (one employee) and "clear" (everyone) entries;
  // merging two copies never loses a tick or brings back one that was undone.
  // ---------------------------------------------------------------------------
  function entryId(e) {
    return e.id || [e.timestamp, e.employeeId, e.action || 'sent', e.channel || ''].join('|');
  }

  // Accepts today's journal or the older plain list (where undone rows were deleted).
  function normalizeLog(raw) {
    var list = Array.isArray(raw) ? raw : raw && Array.isArray(raw.entries) ? raw.entries : [];
    return list.filter(function (e) { return e && e.timestamp && (e.employeeId != null || e.action === 'clear') && !e.undone; }).map(function (e) {
      var out = {};
      Object.keys(e).forEach(function (k) { out[k] = e[k]; });
      out.action = e.action || 'sent';
      out.id = entryId(out);
      return out;
    });
  }

  function mergeLogs(a, b) {
    var seen = {};
    var out = [];
    normalizeLog(a).concat(normalizeLog(b)).forEach(function (e) {
      if (seen[e.id]) return;
      seen[e.id] = true;
      out.push(e);
    });
    out.sort(function (x, y) { return x.timestamp < y.timestamp ? -1 : x.timestamp > y.timestamp ? 1 : 0; });
    return out;
  }

  // employeeId -> the "sent" entry that currently counts.
  function sentMap(entries) {
    var map = {};
    normalizeLog(entries).sort(function (x, y) { return x.timestamp < y.timestamp ? -1 : x.timestamp > y.timestamp ? 1 : 0; }).forEach(function (e) {
      if (e.action === 'clear') map = {};
      else if (e.action === 'undo') delete map[e.employeeId];
      else map[e.employeeId] = e;
    });
    return map;
  }

  // Reads a delivery log CSV back (after a browser reset, or on another computer).
  function parseCsv(text) {
    var rows = [], row = [], cell = '', q = false;
    text = String(text || '').replace(/^﻿/, '');
    for (var i = 0; i < text.length; i++) {
      var ch = text[i];
      if (q) {
        if (ch === '"' && text[i + 1] === '"') { cell += '"'; i++; }
        else if (ch === '"') q = false;
        else cell += ch;
      } else if (ch === '"') q = true;
      else if (ch === ',') { row.push(cell); cell = ''; }
      else if (ch === '\n' || ch === '\r') {
        if (ch === '\r' && text[i + 1] === '\n') i++;
        row.push(cell); rows.push(row); row = []; cell = '';
      } else cell += ch;
    }
    if (cell !== '' || row.length) { row.push(cell); rows.push(row); }
    return rows.filter(function (r) { return r.some(function (c) { return c !== ''; }); });
  }

  function logFromCsv(text) {
    var rows = parseCsv(text);
    if (!rows.length) return [];
    var head = rows[0].map(function (h) { return h.trim().toLowerCase(); });
    var col = function (name) { return head.indexOf(name); };
    if (col('timestamp') < 0 || col('employee_id') < 0) throw new Error('This is not a delivery log (needs "timestamp" and "employee_id" columns).');
    return normalizeLog(rows.slice(1).map(function (r) {
      var get = function (name) { var i = col(name); return i >= 0 ? (r[i] || '') : ''; };
      return { timestamp: get('timestamp'), employeeId: get('employee_id'), name: get('name'), channel: get('channel'), contact: get('contact'), file: get('file'), hash: get('sha256'), note: get('note'), action: get('action') || 'sent', source: get('source') };
    }));
  }

  // ---------------------------------------------------------------------------
  // Current work (file + where the user was), restored after a refresh or browser restart.
  // ---------------------------------------------------------------------------
  var SESSION_FILE = 'session-file';
  var SESSION_STATE = 'session-state';
  var SESSION_TTL_DAYS = 30;

  function saveSessionFile(rec) { return set(SESSION_FILE, rec); }
  function saveSessionState(rec) { return set(SESSION_STATE, rec); }

  function loadSession(now) {
    return Promise.all([get(SESSION_FILE), get(SESSION_STATE)]).then(function (r) {
      var file = r[0], st = r[1];
      if (!file || (!file.bytes && file.text == null)) return null;
      var age = (now || Date.now()) - Date.parse((st && st.savedAt) || file.savedAt || 0);
      if (!(age >= 0) || age > SESSION_TTL_DAYS * 864e5) { clearSession(); return null; }
      return { file: file, state: st || {}, ageMs: age };
    }).catch(function () { return null; });
  }

  function clearSession() {
    return Promise.all([del(SESSION_FILE), del(SESSION_STATE)]);
  }

  function timeAgo(iso, now) {
    var ms = (now || Date.now()) - Date.parse(iso);
    if (!(ms >= 0)) return '';
    var m = Math.round(ms / 60000);
    if (m < 1) return 'just now';
    if (m < 60) return m + ' min ago';
    var h = Math.round(m / 60);
    if (h < 24) return h + ' hour' + (h === 1 ? '' : 's') + ' ago';
    var d = Math.round(h / 24);
    return d + ' day' + (d === 1 ? '' : 's') + ' ago';
  }

  return {
    get: get,
    set: set,
    del: del,
    keys: keys,
    eraseAll: eraseAll,
    persist: persist,
    persisted: persisted,
    readJson: readJson,
    writeJson: writeJson,
    touchVisit: touchVisit,
    visit: visit,
    updateVisit: updateVisit,
    backup: backup,
    setBackup: setBackup,
    prefs: prefs,
    setPrefs: setPrefs,
    stableStringify: stableStringify,
    hashText: hashText,
    settingsFingerprint: settingsFingerprint,
    normalizeLog: normalizeLog,
    mergeLogs: mergeLogs,
    sentMap: sentMap,
    parseCsv: parseCsv,
    logFromCsv: logFromCsv,
    saveSessionFile: saveSessionFile,
    saveSessionState: saveSessionState,
    loadSession: loadSession,
    clearSession: clearSession,
    timeAgo: timeAgo,
    SESSION_TTL_DAYS: SESSION_TTL_DAYS
  };
});
