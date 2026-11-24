/**
 * Minimal in-memory stand-ins for the Apps Script services used by Code.gs
 * (SpreadsheetApp, LockService, Session). Used by dev/build-preview.js and
 * dev/test.js only — never deployed.
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
            setValue: function (v) {
              while (grid.length < r) grid.push([]);
              grid[r - 1][c - 1] = v;
              writes.push({ sheet: name, row: r, col: c, value: v });
              return this;
            },
          };
        },
        appendRow: function (row) { grid.push(row.slice()); return sheet; },
        setFrozenRows: function () { return sheet; },
      };
      return sheet;
    }

    var sheets = [makeSheet(opts.sheetName || 'Database', 0, values)];
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

    return {
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
    };
  }

  if (typeof module !== 'undefined' && module.exports) module.exports = createMockServices;
  else root.createMockServices = createMockServices;
})(this);
