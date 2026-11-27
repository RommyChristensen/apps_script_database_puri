/**
 * Minimal in-memory stand-ins for the Apps Script services used by Code.gs and
 * Auth.gs (SpreadsheetApp, LockService, Session, CacheService, Utilities).
 * Used by dev/build-preview.js and dev/test.js only — never deployed.
 *
 * opts.sheets adds more tabs after the database tab: { 'Authentication': [[...], ...] }.
 *
 * Works in Node (module.exports) and in the browser (window.createMockServices).
 */
(function (root) {
  function createMockServices(values, opts) {
    opts = opts || {};
    var formulas = opts.formulas || {}; // { 'row,col': '=FORMULA' } (1-based)
    var writes = [];

    function makeSheet(name, id, grid) {
      var sheet = {
        _grid: grid,
        getName: function () { return name; },
        getSheetId: function () { return id; },
        getLastRow: function () { return grid.length; },
        getLastColumn: function () { return grid.reduce(function (m, r) { return Math.max(m, r.length); }, 0); },
        getDataRange: function () { return sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()); },
        getRange: function (r, c, nr, nc) {
          if (typeof r === 'string') return { setFontWeight: function () { return this; } };
          nr = nr || 1; nc = nc || 1;
          function cells(fn) {
            var out = [];
            for (var i = 0; i < nr; i++) {
              var row = [];
              for (var j = 0; j < nc; j++) row.push(fn(r + i, c + j));
              out.push(row);
            }
            return out;
          }
          function raw(ri, ci) { var row = grid[ri - 1] || []; return row[ci - 1] === undefined ? '' : row[ci - 1]; }
          return {
            getValues: function () { return cells(raw); },
            getDisplayValues: function () {
              return cells(function (ri, ci) { var v = raw(ri, ci); return typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : String(v); });
            },
            getFormulas: function () { return cells(function (ri, ci) { return formulas[ri + ',' + ci] || ''; }); },
            setNumberFormat: function () { return this; },
            setFontWeight: function () { return this; },
            copyTo: function (dest) { return this; },
            setValues: function (vals) {
              for (var i = 0; i < vals.length; i++) {
                while (grid.length < r + i) grid.push([]);
                for (var j = 0; j < vals[i].length; j++) grid[r + i - 1][c + j - 1] = vals[i][j];
              }
              writes.push({ sheet: name, row: r, col: c, values: vals });
              return this;
            },
            setValue: function (v) {
              while (grid.length < r) grid.push([]);
              grid[r - 1][c - 1] = v;
              writes.push({ sheet: name, row: r, col: c, value: v });
              return this;
            },
          };
        },
        appendRow: function (row) { grid.push(row.slice()); return sheet; },
        deleteRow: function (n) { grid.splice(n - 1, 1); return sheet; },
        insertRowAfter: function (n) { grid.splice(n, 0, []); return sheet; },
        setFrozenRows: function () { return sheet; },
      };
      return sheet;
    }

    var sheets = [makeSheet(opts.sheetName || 'Database', 0, values)];
    Object.keys(opts.sheets || {}).forEach(function (name, i) { sheets.push(makeSheet(name, i + 1, opts.sheets[name])); });

    var cacheStore = {};
    var cache = {
      get: function (k) { return Object.prototype.hasOwnProperty.call(cacheStore, k) ? cacheStore[k] : null; },
      put: function (k, v) { cacheStore[k] = String(v); },
      remove: function (k) { delete cacheStore[k]; },
    };
    var uuidCounter = 0;
    var propStore = {};
    var props = {
      getProperty: function (k) { return Object.prototype.hasOwnProperty.call(propStore, k) ? propStore[k] : null; },
      setProperty: function (k, v) { propStore[k] = String(v); return props; },
      deleteProperty: function (k) { delete propStore[k]; return props; },
    };
    var ss = {
      getSheets: function () { return sheets.slice(); },
      getSheetByName: function (n) { return sheets.filter(function (s) { return s.getName() === n; })[0] || null; },
      insertSheet: function (n, index) {
        var s = makeSheet(n, sheets.length + 100, []);
        sheets.splice(index == null ? sheets.length : index, 0, s);
        return s;
      },
      getUrl: function () { return 'https://docs.google.com/spreadsheets/d/demo/edit'; },
    };

    var api = {
      writes: writes,
      spreadsheet: ss,
      SpreadsheetApp: {
        openById: function () { return ss; },
        getActiveSpreadsheet: function () { return ss; },
        flush: function () {},
      },
      LockService: { getScriptLock: function () { return { tryLock: function () { return true; }, releaseLock: function () {} }; } },
      Session: { getActiveUser: function () { return { getEmail: function () { return opts.email || 'preview@example.com'; } }; } },
      HtmlService: {},
      cacheStore: cacheStore,
      geocodeCalls: [],
      fetchCalls: [],
      PropertiesService: { getScriptProperties: function () { return props; } },
      // Fake geocoder: opts.geocode maps a lower-cased query to {lat, lng}; anything else is "not found".
      Maps: {
        newGeocoder: function () {
          var g = {
            setRegion: function () { return g; }, setLanguage: function () { return g; }, setBounds: function () { return g; },
            geocode: function (q) {
              api.geocodeCalls.push(q);
              var key = String(q).toLowerCase().replace(/, indonesia$/, '');
              var hit = (opts.geocode || {})[key];
              if (!hit && opts.geocodeAny) {
                // Preview only: a stable made-up point around West Jakarta for any text.
                var h = 0;
                for (var i = 0; i < key.length; i++) h = (h * 31 + key.charCodeAt(i)) >>> 0;
                hit = { lat: -6.24 + (h % 1000) / 1000 * 0.16, lng: 106.66 + ((h >>> 10) % 1000) / 1000 * 0.16 };
              }
              return hit ? { status: 'OK', results: [{ geometry: { location: hit } }] } : { status: 'ZERO_RESULTS', results: [] };
            },
          };
          return g;
        },
      },
      // Fake fetch: opts.redirects maps a URL to the Location it redirects to.
      UrlFetchApp: {
        fetch: function (url) {
          api.fetchCalls.push(url);
          var loc = (opts.redirects || {})[url];
          return { getHeaders: function () { return loc ? { Location: loc } : {}; } };
        },
      },
      CacheService: { getScriptCache: function () { return cache; } },
      Utilities: {
        DigestAlgorithm: { SHA_256: 'SHA_256' },
        Charset: { UTF_8: 'UTF_8' },
        getUuid: function () {
          var hex = '';
          for (var i = 0; i < 32; i++) hex += Math.floor(Math.random() * 16).toString(16);
          uuidCounter++;
          return hex.slice(0, 8) + '-' + hex.slice(8, 12) + '-4' + hex.slice(13, 16) + '-a' + hex.slice(17, 20) + '-' + hex.slice(20, 32);
        },
        // Enough of Utilities.formatDate for 'yyyy-MM-dd HH:mm' and 'yyMMddHHmmss' (local time).
        formatDate: function (d, tz, fmt) {
          var p = function (n) { return ('0' + n).slice(-2); };
          var map = { yyyy: String(d.getFullYear()), yy: String(d.getFullYear()).slice(-2), MM: p(d.getMonth() + 1),
            dd: p(d.getDate()), HH: p(d.getHours()), mm: p(d.getMinutes()), ss: p(d.getSeconds()) };
          return fmt.replace(/yyyy|yy|MM|dd|HH|mm|ss/g, function (k) { return map[k]; });
        },
        // Like Apps Script: an array of signed bytes.
        computeDigest: function (alg, text) {
          return sha256Bytes(unescape(encodeURIComponent(text))).map(function (b) { return b > 127 ? b - 256 : b; });
        },
      },
    };
    return api;
  }

  /** Plain SHA-256 over a binary string, returning byte values 0..255. */
  function sha256Bytes(str) {
    var K = [], H = [], i, j;
    var isPrime = function (n) { for (var f = 2; f * f <= n; f++) if (n % f === 0) return false; return true; };
    var frac = function (x) { return ((x - Math.floor(x)) * 4294967296) >>> 0; };
    for (var n = 2, c = 0; c < 64; n++) if (isPrime(n)) { if (c < 8) H[c] = frac(Math.pow(n, 1 / 2)); K[c++] = frac(Math.pow(n, 1 / 3)); }
    var bytes = [];
    for (i = 0; i < str.length; i++) bytes.push(str.charCodeAt(i) & 0xff);
    var bitLen = bytes.length * 8;
    bytes.push(0x80);
    while (bytes.length % 64 !== 56) bytes.push(0);
    for (i = 7; i >= 0; i--) bytes.push(i >= 4 ? 0 : (bitLen >>> (i * 8)) & 0xff);
    var rotr = function (x, r) { return (x >>> r) | (x << (32 - r)); };
    for (i = 0; i < bytes.length; i += 64) {
      var w = [];
      for (j = 0; j < 16; j++) w[j] = (bytes[i + 4 * j] << 24) | (bytes[i + 4 * j + 1] << 16) | (bytes[i + 4 * j + 2] << 8) | bytes[i + 4 * j + 3];
      for (j = 16; j < 64; j++) {
        var s0 = rotr(w[j - 15], 7) ^ rotr(w[j - 15], 18) ^ (w[j - 15] >>> 3);
        var s1 = rotr(w[j - 2], 17) ^ rotr(w[j - 2], 19) ^ (w[j - 2] >>> 10);
        w[j] = (w[j - 16] + s0 + w[j - 7] + s1) | 0;
      }
      var a = H[0], b = H[1], cc = H[2], d = H[3], e = H[4], f = H[5], g = H[6], h = H[7];
      for (j = 0; j < 64; j++) {
        var t1 = (h + (rotr(e, 6) ^ rotr(e, 11) ^ rotr(e, 25)) + ((e & f) ^ (~e & g)) + K[j] + w[j]) | 0;
        var t2 = ((rotr(a, 2) ^ rotr(a, 13) ^ rotr(a, 22)) + ((a & b) ^ (a & cc) ^ (b & cc))) | 0;
        h = g; g = f; f = e; e = (d + t1) | 0; d = cc; cc = b; b = a; a = (t1 + t2) | 0;
      }
      H = [H[0] + a, H[1] + b, H[2] + cc, H[3] + d, H[4] + e, H[5] + f, H[6] + g, H[7] + h].map(function (x) { return x | 0; });
    }
    var out = [];
    H.forEach(function (x) { out.push((x >>> 24) & 0xff, (x >>> 16) & 0xff, (x >>> 8) & 0xff, x & 0xff); });
    return out;
  }

  createMockServices.sha256Bytes = sha256Bytes;
  if (typeof module !== 'undefined' && module.exports) module.exports = createMockServices;
  else root.createMockServices = createMockServices;
})(this);
