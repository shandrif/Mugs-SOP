/*
 * Branch mapping: reads the HM "Master_Location_Mapping.xlsx" (or a CSV with
 * the same headers) and links each POS store name ("B99, Mec Haram") to its
 * city, zone, area manager, ROM and operations chain by branch code.
 *
 * The mapping is uploaded by the user and kept in the browser (or the cloud
 * database); it is never committed to the repository.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.MugsMapping = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var FIELDS = {
    name: 'BRANCHES_CODES_NAME', city: 'CITY', region: 'REGION', zone: 'ZONE_NAME', subZone: 'SUB_ZONE_NAME',
    type: 'LOCATION_TYPE', driveThru: 'DRIVE_THRU_TYPE', am: 'AREA_MANAGER', rom: 'ROM', om: 'ACTING_OM',
    doo: 'DOO', coo: 'COO', opened: 'OPENING_DATE'
  };

  function clean(v) { return v === null || v === undefined ? '' : String(v).replace(/\s+/g, ' ').trim(); }
  // "B05 Rajhi Bank", "B121- Madina Haram", "b99" -> "B5", "B121", "B99"
  function codeOf(s) {
    var m = clean(s).match(/^B\s*0*(\d+)/i);
    return m ? 'B' + parseInt(m[1], 10) : '';
  }
  function cleanPerson(v) {
    v = clean(v);
    return v === '—' || v === '-' ? '' : v;
  }

  function sheetRows(XLSX, ws) { return XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '' }); }
  function headerIndex(rows, must) {
    for (var r = 0; r < Math.min(rows.length, 15); r++) {
      var h = rows[r].map(function (c) { return clean(c).toUpperCase(); });
      if (must.every(function (m) { return h.indexOf(m) >= 0; })) return { row: r, cols: h };
    }
    return null;
  }

  function parseWorkbook(XLSX, wb) {
    var branches = {}, mainCount = 0, extraCount = 0;
    // 1. Main sheet: Location_Mapping (any sheet that has SYRVE_CODE + AREA_MANAGER)
    var mainName = wb.SheetNames.filter(function (n) { return /location.?mapping/i.test(n); })[0];
    var order = mainName ? [mainName].concat(wb.SheetNames.filter(function (n) { return n !== mainName; })) : wb.SheetNames;
    var found = false;
    order.forEach(function (name) {
      if (found) return;
      var rows = sheetRows(XLSX, wb.Sheets[name]);
      var h = headerIndex(rows, ['SYRVE_CODE', 'AREA_MANAGER']);
      if (!h) return;
      found = true;
      var col = {}; Object.keys(FIELDS).forEach(function (k) { col[k] = h.cols.indexOf(FIELDS[k]); });
      var codeCol = h.cols.indexOf('SYRVE_CODE');
      for (var i = h.row + 1; i < rows.length; i++) {
        var code = codeOf(rows[i][codeCol]);
        if (!code) continue;
        var b = { code: code, source: 'mapping' };
        Object.keys(col).forEach(function (k) { b[k] = col[k] >= 0 ? (k === 'am' || k === 'rom' || k === 'om' || k === 'doo' || k === 'coo' ? cleanPerson(rows[i][col[k]]) : clean(rows[i][col[k]])) : ''; });
        branches[code] = b; mainCount++;
      }
    });
    if (!found) throw new Error('This file has no sheet with SYRVE_CODE and AREA_MANAGER columns. Use the Master Location Mapping workbook.');

    // 2. "Excluded" sheet: fills gaps for new/unonboarded branches that have a B-code (e.g. B121 Madina Haram)
    var ex = wb.SheetNames.filter(function (n) { return /excluded/i.test(n); })[0];
    if (ex) {
      var rows = sheetRows(XLSX, wb.Sheets[ex]);
      var h = headerIndex(rows, ['SOURCE_CODE', 'AREA_MANAGER']);
      if (h) {
        var c = function (k) { return h.cols.indexOf(k); };
        for (var i = h.row + 1; i < rows.length; i++) {
          var code = codeOf(rows[i][c('SOURCE_CODE')]);
          if (!code || branches[code]) continue;
          var status = clean(rows[i][c('STATUS')]);
          if (/closed/i.test(status) && !/held out|new/i.test(status)) continue;
          branches[code] = {
            code: code, source: 'excluded', name: clean(rows[i][c('SOURCE_NAME')]), city: '', region: '', zone: '', subZone: '',
            type: clean(rows[i][c('CATEGORY')]), driveThru: '', am: cleanPerson(rows[i][c('AREA_MANAGER')]), rom: cleanPerson(rows[i][c('ROM')]),
            om: cleanPerson(rows[i][c('ACTING_OM')]), doo: cleanPerson(rows[i][c('DOO')]), coo: cleanPerson(rows[i][c('COO')]),
            opened: clean(rows[i][c('OPENING_DATE')])
          };
          extraCount++;
        }
      }
    }
    return { branches: branches, count: mainCount + extraCount, mainCount: mainCount, extraCount: extraCount };
  }

  function lookup(mapping, storeName) {
    if (!mapping || !mapping.branches) return null;
    return mapping.branches[codeOf(storeName)] || null;
  }

  return { parseWorkbook: parseWorkbook, lookup: lookup, codeOf: codeOf };
});
