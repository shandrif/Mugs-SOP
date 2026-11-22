/*
 * Branded Excel report (HM theme) built with ExcelJS in the browser.
 * MugsReport.build(result, meta) -> Promise<Blob>
 *   result: output of MugsEngine.analyse()
 *   meta:   { fileName, reportDate }
 */
(function (root) {
  'use strict';

  /* ---------- brand palette (ARGB for Excel, hex for canvas) ---------- */
  var C = {
    green: '0B3D2C', green2: '0F4A36', sand: 'E9DCC0', sand2: 'D9C7A3', cream: 'FFFDF8', paper: 'F7F1E6',
    zebra: 'FBF7EF', gold: 'B8862C', ink: '15211B', ink2: '44514A', muted: '6F786F', line: 'E6DFD2',
    direct: '0A7D55', split: 'B8862C',
    barDirect: 'A9D8C2', barSplit: 'E8D09C', barHigh: 'F1B3AA',
    High: 'B3261E', HighBg: 'FBE8E4', Medium: '8F6210', MediumBg: 'F7EDD6',
    Low: '3B5D8A', LowBg: 'E7EDF6', Clean: '0A6E4B', CleanBg: 'E2F1E9'
  };
  var A = function (hex) { return 'FF' + hex; };
  var H = function (hex) { return '#' + hex; };
  var SERIF = 'Georgia', SANS = 'Calibri';

  function fill(hex) { return { type: 'pattern', pattern: 'solid', fgColor: { argb: A(hex) } }; }
  function font(o) {
    return { name: o.serif ? SERIF : SANS, size: o.size || 10, bold: !!o.bold, italic: !!o.italic, color: { argb: A(o.color || C.ink) } };
  }
  function xlDate(ms) { return new Date(ms); } // engine times are naive (UTC fields); ExcelJS writes UTC
  function fmtDate(ms) {
    return new Date(ms).toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric', timeZone: 'UTC' });
  }
  function pad(n) { return (n < 10 ? '0' : '') + n; }
  function stamp(ms) {
    var d = new Date(ms);
    return d.getUTCFullYear() + '-' + pad(d.getUTCMonth() + 1) + '-' + pad(d.getUTCDate()) + ' ' + pad(d.getUTCHours()) + ':' + pad(d.getUTCMinutes());
  }
  var nf0 = new Intl.NumberFormat('en-US', { maximumFractionDigits: 0 });

  /* ======================================================================
     Charts (drawn on canvas at 2x, embedded as PNG)
     ====================================================================== */
  var FONT = '"Segoe UI", Manrope, Calibri, Arial, sans-serif';

  function canvas(w, h) {
    var c = document.createElement('canvas');
    c.width = w * 2; c.height = h * 2;
    var g = c.getContext('2d');
    g.scale(2, 2);
    g.fillStyle = H(C.cream); g.fillRect(0, 0, w, h);
    return { c: c, g: g, w: w, h: h };
  }
  function niceMax(v) {
    if (v <= 10) return 10;
    var p = Math.pow(10, Math.floor(Math.log10(v))), m = v / p;
    return (m <= 2 ? 2 : m <= 2.5 ? 2.5 : m <= 5 ? 5 : 10) * p;
  }
  function roundBarRight(g, x, y, w, h, r) {
    r = Math.min(r, w, h / 2);
    g.beginPath();
    g.moveTo(x, y); g.lineTo(x + w - r, y); g.quadraticCurveTo(x + w, y, x + w, y + r);
    g.lineTo(x + w, y + h - r); g.quadraticCurveTo(x + w, y + h, x + w - r, y + h);
    g.lineTo(x, y + h); g.closePath(); g.fill();
  }
  function roundBarTop(g, x, y, w, h, r) {
    r = Math.min(r, w / 2, h);
    g.beginPath();
    g.moveTo(x, y + h); g.lineTo(x, y + r); g.quadraticCurveTo(x, y, x + r, y);
    g.lineTo(x + w - r, y); g.quadraticCurveTo(x + w, y, x + w, y + r);
    g.lineTo(x + w, y + h); g.closePath(); g.fill();
  }
  function legend(g, x, y) {
    var items = [['Direct violations', C.direct], ['Split-invoice runs', C.split]];
    g.font = '500 12px ' + FONT; g.textBaseline = 'middle';
    items.forEach(function (it) {
      g.fillStyle = H(it[1]); g.beginPath(); g.roundRect ? g.roundRect(x, y - 6, 12, 12, 3) : g.rect(x, y - 6, 12, 12); g.fill();
      g.fillStyle = H(C.ink2); g.fillText(it[0], x + 18, y);
      x += 18 + g.measureText(it[0]).width + 18;
    });
  }
  function ellipsize(g, text, maxW) {
    if (g.measureText(text).width <= maxW) return text;
    while (text.length > 1 && g.measureText(text + '…').width > maxW) text = text.slice(0, -1);
    return text + '…';
  }

  // Top branches by mugs over the limit (stacked horizontal bars)
  function chartBranches(stores) {
    var data = stores.filter(function (s) { return s.flaggedExcess > 0; })
      .sort(function (a, b) { return b.flaggedExcess - a.flaggedExcess; }).slice(0, 10);
    var K = canvas(600, 390), g = K.g;
    legend(g, 16, 18);
    if (!data.length) { g.fillStyle = H(C.muted); g.font = '14px ' + FONT; g.fillText('No mugs over the limit.', 16, 60); return K.c.toDataURL('image/png'); }
    var left = 200, right = 48, top = 44, bottom = 30;
    var plotW = K.w - left - right, rowH = (K.h - top - bottom) / data.length, barH = Math.min(18, rowH * 0.55);
    var max = niceMax(data[0].flaggedExcess);
    // grid
    g.strokeStyle = 'rgba(21,33,27,0.08)'; g.lineWidth = 1; g.fillStyle = H(C.muted); g.font = '11px ' + FONT; g.textAlign = 'center'; g.textBaseline = 'top';
    for (var i = 0; i <= 4; i++) {
      var gx = Math.round(left + plotW * i / 4) + 0.5;
      g.beginPath(); g.moveTo(gx, top - 4); g.lineTo(gx, K.h - bottom); g.stroke();
      g.fillText(nf0.format(max * i / 4), gx, K.h - bottom + 8);
    }
    data.forEach(function (s, idx) {
      var cy = top + rowH * idx + rowH / 2, y = cy - barH / 2;
      g.textAlign = 'right'; g.textBaseline = 'middle';
      g.fillStyle = H(C.ink); g.font = '600 12px ' + FONT;
      g.fillText(ellipsize(g, s.store, left - 24), left - 12, cy - 6);
      g.fillStyle = H(s.status === 'High' ? C.High : C.muted); g.font = '11px ' + FONT;
      g.fillText(s.status + ' risk', left - 12, cy + 8);
      var x = left, wD = plotW * s.directExcess / max, wS = plotW * s.runExcess / max;
      if (wD > 0) { g.fillStyle = H(C.direct); if (wS > 0) g.fillRect(x, y, Math.max(1, wD - 1), barH); else roundBarRight(g, x, y, Math.max(2, wD), barH, 4); x += wD; }
      if (wS > 0) { g.fillStyle = H(C.split); roundBarRight(g, x + (wD > 0 ? 1 : 0), y, Math.max(2, wS - (wD > 0 ? 1 : 0)), barH, 4); x += wS; }
      g.textAlign = 'left'; g.fillStyle = H(C.ink); g.font = '700 12px ' + FONT;
      g.fillText(nf0.format(s.flaggedExcess), x + 8, cy);
    });
    return K.c.toDataURL('image/png');
  }

  // Mugs over the limit by hour of day (stacked columns)
  function chartHours(result) {
    var d = new Array(24).fill(0), s = new Array(24).fill(0);
    result.direct.forEach(function (v) { d[new Date(v.time).getUTCHours()] += v.excess; });
    result.runs.forEach(function (r) { s[new Date(r.start).getUTCHours()] += r.excess; });
    var K = canvas(600, 390), g = K.g;
    legend(g, 16, 18);
    var left = 46, right = 14, top = 44, bottom = 34;
    var plotW = K.w - left - right, plotH = K.h - top - bottom;
    var max = niceMax(Math.max.apply(null, d.map(function (v, i) { return v + s[i]; })) || 1);
    g.strokeStyle = 'rgba(21,33,27,0.08)'; g.lineWidth = 1; g.fillStyle = H(C.muted); g.font = '11px ' + FONT;
    for (var i = 0; i <= 4; i++) {
      var gy = Math.round(top + plotH - plotH * i / 4) + 0.5;
      g.beginPath(); g.moveTo(left, gy); g.lineTo(K.w - right, gy); g.stroke();
      g.textAlign = 'right'; g.textBaseline = 'middle'; g.fillText(nf0.format(max * i / 4), left - 8, gy);
    }
    var slot = plotW / 24, bw = Math.min(16, slot * 0.62), peak = 0;
    for (var h = 0; h < 24; h++) if (d[h] + s[h] > d[peak] + s[peak]) peak = h;
    for (h = 0; h < 24; h++) {
      var x = left + slot * h + (slot - bw) / 2, base = top + plotH;
      var hD = plotH * d[h] / max, hS = plotH * s[h] / max;
      if (hD > 0) { g.fillStyle = H(C.direct); if (hS > 0) g.fillRect(x, base - hD, bw, hD); else roundBarTop(g, x, base - hD, bw, hD, 4); }
      if (hS > 0) { g.fillStyle = H(C.split); roundBarTop(g, x, base - hD - hS - (hD > 0 ? 2 : 0), bw, hS, 4); }
      if (h % 3 === 0) { g.fillStyle = H(C.muted); g.textAlign = 'center'; g.textBaseline = 'top'; g.fillText(pad(h) + ':00', x + bw / 2, base + 8); }
    }
    // label the peak hour only
    var tot = d[peak] + s[peak];
    if (tot > 0) {
      var px = left + slot * peak + slot / 2, py = top + plotH - plotH * tot / max - (d[peak] && s[peak] ? 2 : 0);
      g.fillStyle = H(C.ink); g.font = '700 12px ' + FONT; g.textAlign = 'center'; g.textBaseline = 'bottom';
      g.fillText('Peak ' + pad(peak) + ':00 · ' + nf0.format(tot), Math.min(Math.max(px, left + 60), K.w - right - 60), py - 6);
    }
    return K.c.toDataURL('image/png');
  }

  /* ======================================================================
     Sheet helpers
     ====================================================================== */
  function badge(cell, level) {
    cell.value = level;
    cell.fill = fill(C[level + 'Bg'] || C.paper);
    cell.font = font({ bold: true, color: C[level] || C.ink, size: 10 });
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
  }

  // A styled data table: green header, zebra rows, filters, frozen header.
  function table(ws, cols, rows, opts) {
    opts = opts || {};
    ws.columns = cols.map(function (c) { return { key: c.key, width: c.width || 14 }; });
    var hdr = ws.getRow(1);
    cols.forEach(function (c, i) {
      var cell = hdr.getCell(i + 1);
      cell.value = c.header;
      cell.fill = fill(C.green);
      cell.font = font({ bold: true, color: C.sand, size: 10 });
      cell.alignment = { vertical: 'middle', horizontal: c.num ? 'right' : (c.center ? 'center' : 'left'), wrapText: true };
      cell.border = { bottom: { style: 'medium', color: { argb: A(C.gold) } } };
    });
    hdr.height = 30;
    rows.forEach(function (r, ri) {
      var row = ws.getRow(ri + 2);
      cols.forEach(function (c, ci) {
        var cell = row.getCell(ci + 1), v = r[c.key];
        if (c.badge) { badge(cell, v); }
        else {
          cell.value = v === undefined ? null : v;
          cell.font = font({ size: 10, bold: !!c.bold, color: c.color || C.ink });
          cell.alignment = { vertical: 'middle', horizontal: c.num ? 'right' : (c.center ? 'center' : 'left'), wrapText: !!c.wrap };
          if (ri % 2 === 1) cell.fill = fill(C.zebra);
        }
        if (c.fmt) cell.numFmt = c.fmt;
        cell.border = { bottom: { style: 'thin', color: { argb: A(C.line) } } };
      });
      row.height = opts.rowHeight || 18;
    });
    ws.views = [{ state: 'frozen', ySplit: 1, xSplit: opts.freezeCols || 0, showGridLines: false, zoomScale: 100 }];
    if (rows.length) {
      ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: rows.length + 1, column: cols.length } };
      cols.forEach(function (c, ci) {
        if (!c.bar) return;
        var col = ws.getColumn(ci + 1).letter;
        ws.addConditionalFormatting({
          ref: col + '2:' + col + (rows.length + 1),
          rules: [{ type: 'dataBar', minLength: 0, maxLength: 100, gradient: false, border: false,
            cfvo: [{ type: 'num', value: 0 }, c.barMax !== undefined ? { type: 'num', value: c.barMax } : { type: 'max' }],
            color: { argb: A(c.bar) } }]
        });
      });
    } else {
      ws.getCell('A2').value = 'Nothing flagged with these rules.';
      ws.getCell('A2').font = font({ italic: true, color: C.muted });
    }
    ws.pageSetup = { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0, margins: { left: 0.4, right: 0.4, top: 0.5, bottom: 0.5, header: 0.2, footer: 0.2 } };
    ws.headerFooter = { oddFooter: '&L&8Half Million · Mugs SOP Monitor&R&8Page &P of &N' };
  }

  /* ======================================================================
     Dashboard sheet
     ====================================================================== */
  function dashboard(wb, ws, R, meta, images) {
    var T = R.totals, o = R.options;
    // grid: A margin, B..M = 12 content columns, N margin
    ws.columns = [{ width: 2.5 }].concat(new Array(12).fill(0).map(function () { return { width: 14.5 }; })).concat([{ width: 2.5 }]);
    ws.views = [{ showGridLines: false, zoomScale: 90 }];
    ws.properties.defaultRowHeight = 18;

    function band(r1, r2, hex) {
      for (var r = r1; r <= r2; r++) for (var c = 1; c <= 14; c++) ws.getCell(r, c).fill = fill(hex);
    }
    function put(addr, value, f, align) {
      var cell = ws.getCell(addr);
      cell.value = value; if (f) cell.font = f;
      cell.alignment = Object.assign({ vertical: 'middle', horizontal: 'left' }, align || {});
      return cell;
    }
    function merge(range, value, f, align, bg) {
      ws.mergeCells(range);
      var cell = put(range.split(':')[0], value, f, align);
      if (bg) cell.fill = fill(bg);
      return cell;
    }

    /* --- title band --- */
    band(1, 7, C.green);
    [1, 7].forEach(function (r) { ws.getRow(r).height = 10; });
    ws.getRow(2).height = 18; ws.getRow(3).height = 36; ws.getRow(4).height = 20; ws.getRow(5).height = 18; ws.getRow(6).height = 18;
    merge('B2:I2', 'HALF MILLION  ·  RETAIL COMPLIANCE', font({ bold: true, size: 9, color: C.sand2 }));
    merge('B3:M3', 'Mugs SOP report  ·  ' + fmtDate(T.start), font({ serif: true, size: 24, color: 'F6EEDC' }));
    merge('B4:M4', meta.fileName + '  ·  ' + T.stores + ' branches  ·  ' + stamp(T.start) + ' → ' + stamp(T.end), font({ size: 10, color: C.sand }));
    merge('B5:M5', 'Rules: max ' + o.limit + ' per ' + (o.groupBy === 'colour' ? 'colour' : 'item/SKU') + ' per customer  ·  invoices ≤ ' + o.gapMinutes +
      ' min apart form a run  ·  min ' + o.minInvoices + ' invoices per run', font({ size: 10, color: C.sand2 }));
    merge('J2:M2', 'Generated ' + stamp(Date.now() - new Date().getTimezoneOffset() * 60000), font({ size: 9, color: C.sand2 }), { horizontal: 'right' });
    ws.getRow(8).height = 12;

    /* --- KPI tiles: 6 tiles x 2 columns, rows 9-12 --- */
    var highStores = R.stores.filter(function (s) { return s.status === 'High'; }).length;
    var tiles = [
      { label: 'HIGH-RISK BRANCHES', value: highStores, note: 'of ' + T.stores + ' branches need follow-up', feature: true, accent: C.gold },
      { label: 'DIRECT VIOLATIONS', value: T.directLines, note: nf0.format(T.directExcess) + ' mugs over limit · ' + T.directStores + ' branches', accent: C.High, labelColor: C.High },
      { label: 'SPLIT-INVOICE RUNS', value: T.runs, note: T.runsHigh + ' high · ' + T.runsMedium + ' medium · ' + T.runStores + ' branches', accent: C.split, labelColor: C.Medium },
      { label: 'MUGS SOLD', value: T.mugs, note: nf0.format(T.invoices) + ' invoices', accent: C.line },
      { label: 'NET SALES (SAR)', value: T.sales, fmt: '#,##0', note: 'SAR ' + T.sales.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }), accent: C.line },
      { label: 'ZERO-QTY LINES', value: T.zero, note: 'voids or returns to review', accent: C.line }
    ];
    ws.getRow(9).height = 22; ws.getRow(10).height = 24; ws.getRow(11).height = 24; ws.getRow(12).height = 30;
    tiles.forEach(function (t, i) {
      var c1 = 2 + i * 2, c2 = c1 + 1;
      var L = function (c) { return ws.getColumn(c).letter; };
      var bg = t.feature ? C.green : C.cream, fg = t.feature ? C.sand : C.ink, sub = t.feature ? C.sand2 : C.muted;
      for (var r = 9; r <= 12; r++) for (var c = c1; c <= c2; c++) {
        var cell = ws.getCell(r, c);
        cell.fill = fill(bg);
        cell.border = {
          top: r === 9 ? { style: 'thin', color: { argb: A(t.feature ? C.green : C.line) } } : undefined,
          left: c === c1 ? { style: 'thin', color: { argb: A(t.feature ? C.green : C.line) } } : undefined,
          right: c === c2 ? { style: 'thin', color: { argb: A(t.feature ? C.green : C.line) } } : undefined,
          bottom: r === 12 ? { style: 'thick', color: { argb: A(t.accent) } } : undefined
        };
      }
      merge(L(c1) + '9:' + L(c2) + '9', '  ' + t.label, font({ bold: true, size: 9, color: t.feature ? C.sand2 : (t.labelColor || C.muted) }));
      var v = merge(L(c1) + '10:' + L(c2) + '11', t.value, font({ serif: true, size: 26, color: fg }), { indent: 1 });
      v.numFmt = t.fmt || '#,##0';
      merge(L(c1) + '12:' + L(c2) + '12', t.note, font({ size: 9, color: sub }), { indent: 1, wrapText: true, vertical: 'top' });
    });
    ws.getRow(13).height = 16;

    /* --- charts --- */
    merge('B14:G14', 'Where the excess mugs are', font({ bold: true, size: 13, color: C.ink }));
    merge('B15:G15', 'Mugs sold over the limit, top 10 branches', font({ size: 9, color: C.muted }));
    merge('H14:M14', 'When it happens', font({ bold: true, size: 13, color: C.ink }));
    merge('H15:M15', 'Mugs over the limit by hour of day (run start time for split runs)', font({ size: 9, color: C.muted }));
    var img1 = wb.addImage({ base64: images.branches, extension: 'png' });
    var img2 = wb.addImage({ base64: images.hours, extension: 'png' });
    ws.addImage(img1, { tl: { col: 1.05, row: 15.3 }, ext: { width: 600, height: 390 } });
    ws.addImage(img2, { tl: { col: 7.1, row: 15.3 }, ext: { width: 600, height: 390 } });
    for (var r = 16; r <= 36; r++) ws.getRow(r).height = 18.75;
    var row = 38;

    /* --- top split runs --- */
    function sectionTitle(r, text, sub) {
      ws.getRow(r).height = 24;
      merge('B' + r + ':M' + r, text, font({ bold: true, size: 13, color: C.ink }));
      if (sub) { ws.getRow(r + 1).height = 16; merge('B' + (r + 1) + ':M' + (r + 1), sub, font({ size: 9, color: C.muted })); }
      return r + (sub ? 2 : 1);
    }
    function headerRow(r, spec) {
      ws.getRow(r).height = 22;
      spec.forEach(function (s) {
        if (s.range.indexOf(':') > 0) ws.mergeCells(s.range.replace(/#/g, r));
        var cell = ws.getCell(s.range.split(':')[0].replace(/#/g, r));
        cell.value = s.h;
        cell.font = font({ bold: true, size: 9, color: C.sand });
        cell.alignment = { vertical: 'middle', horizontal: s.num ? 'right' : (s.center ? 'center' : 'left'), indent: s.num ? 0 : 1 };
      });
      for (var c = 2; c <= 13; c++) {
        ws.getCell(r, c).fill = fill(C.green);
        ws.getCell(r, c).border = { bottom: { style: 'medium', color: { argb: A(C.gold) } } };
      }
    }
    function bodyRow(r, spec, rec, zebra) {
      ws.getRow(r).height = 20;
      spec.forEach(function (s) {
        if (s.range.indexOf(':') > 0) ws.mergeCells(s.range.replace(/#/g, r));
        var cell = ws.getCell(s.range.split(':')[0].replace(/#/g, r));
        if (s.badge) { badge(cell, rec[s.k]); return; }
        cell.value = rec[s.k];
        cell.font = font({ size: 10, bold: !!s.bold });
        cell.alignment = { vertical: 'middle', horizontal: s.num ? 'right' : (s.center ? 'center' : 'left'), indent: s.num ? 0 : 1 };
        if (s.fmt) cell.numFmt = s.fmt;
      });
      var badgeCols = spec.filter(function (x) { return x.badge; }).map(function (x) { return ws.getCell(x.range.replace(/#/g, r)).col; });
      for (var c = 2; c <= 13; c++) {
        var cl = ws.getCell(r, c);
        if (zebra && badgeCols.indexOf(c) < 0) cl.fill = fill(C.zebra);
        cl.border = { bottom: { style: 'thin', color: { argb: A(C.line) } } };
      }
    }

    row = sectionTitle(row, 'Top split-invoice runs', 'Same item rung on back-to-back invoices, each within the limit, adding up to more. Full list on the "Split runs" sheet.');
    var runSpec = [
      { range: 'B#', h: 'Risk', k: 'level', badge: true, center: true },
      { range: 'C#', h: 'Score', k: 'score', num: true, fmt: '0' },
      { range: 'D#:E#', h: 'Branch', k: 'store', bold: true },
      { range: 'F#:G#', h: 'Item / colour', k: 'label' },
      { range: 'H#', h: 'From', k: 'start', fmt: 'hh:mm:ss', center: true },
      { range: 'I#', h: 'To', k: 'end', fmt: 'hh:mm:ss', center: true },
      { range: 'J#', h: 'Invoices', k: 'invoices', num: true },
      { range: 'K#', h: 'Mugs', k: 'mugs', num: true },
      { range: 'L#', h: 'Over limit', k: 'excess', num: true },
      { range: 'M#', h: '× normal pace', k: 'pace', num: true, fmt: '[>=100]"100+";0.0' }
    ];
    headerRow(row, runSpec); row++;
    var topRuns = R.runs.slice(0, 10);
    if (!topRuns.length) { merge('B' + row + ':M' + row, 'No split-invoice runs with these rules.', font({ italic: true, color: C.muted })); row++; }
    topRuns.forEach(function (r, i) {
      bodyRow(row++, runSpec, { level: r.level, score: r.score, store: r.store, label: r.label, start: xlDate(r.start), end: xlDate(r.end),
        invoices: r.invoices, mugs: r.mugs, excess: r.excess, pace: Math.min(r.burst, 100) }, i % 2 === 1);
    });
    row++;

    /* --- top direct violations --- */
    row = sectionTitle(row, 'Largest direct violations', 'Single invoices over the limit. Full list on the "Direct violations" sheet.');
    var dirSpec = [
      { range: 'B#:D#', h: 'Branch', k: 'store', bold: true },
      { range: 'E#:G#', h: 'Item / colour', k: 'label' },
      { range: 'H#', h: 'Time', k: 'time', fmt: 'hh:mm:ss', center: true },
      { range: 'I#', h: 'Qty', k: 'qty', num: true },
      { range: 'J#', h: 'Over limit', k: 'excess', num: true },
      { range: 'K#', h: 'SAR', k: 'sales', num: true, fmt: '#,##0.00' },
      { range: 'L#:M#', h: 'Items on invoice', k: 'lines', center: true }
    ];
    headerRow(row, dirSpec); row++;
    var topDirect = R.direct.slice(0, 10);
    if (!topDirect.length) { merge('B' + row + ':M' + row, 'No direct violations with these rules.', font({ italic: true, color: C.muted })); row++; }
    topDirect.forEach(function (d, i) {
      bodyRow(row++, dirSpec, { store: d.store, label: d.label, time: xlDate(d.time), qty: d.qty, excess: d.excess, sales: d.sales,
        lines: d.basket ? d.basket.split(' + ').length : 1 }, i % 2 === 1);
    });
    row++;

    /* --- branches by risk level --- */
    row = sectionTitle(row, 'Branches by risk level', null);
    var counts = { High: 0, Medium: 0, Low: 0, Clean: 0 };
    R.stores.forEach(function (s) { counts[s.status]++; });
    var first = row;
    ['High', 'Medium', 'Low', 'Clean'].forEach(function (lv) {
      ws.getRow(row).height = 20;
      badge(ws.getCell('B' + row), lv);
      ws.mergeCells('C' + row + ':D' + row);
      var c = ws.getCell('C' + row);
      c.value = counts[lv]; c.numFmt = '#,##0 "branches"'; c.font = font({ bold: true, size: 10 }); c.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
      ws.mergeCells('E' + row + ':M' + row);
      var b = ws.getCell('E' + row); b.value = counts[lv];
      b.font = font({ size: 1, color: C.cream });
      row++;
    });
    ['High', 'Medium', 'Low', 'Clean'].forEach(function (lv, i) {
      ws.addConditionalFormatting({
        ref: 'E' + (first + i) + ':M' + (first + i),
        rules: [{ type: 'dataBar', minLength: 0, maxLength: 100, gradient: false, border: false,
          cfvo: [{ type: 'num', value: 0 }, { type: 'num', value: T.stores }], color: { argb: A(C[lv]) } }]
      });
    });
    row++;

    /* --- footer note --- */
    ws.getRow(row).height = 46;
    merge('B' + row + ':M' + row,
      'How to read this: invoices are rebuilt from branch + closing time. A split run is likely, not proven, until the POS export includes a card, ' +
      'customer or cashier ID; check High runs with the branch, POS logs and CCTV. Risk score = invoices (≤30) + mugs over limit (≤25) + burst vs ' +
      'the branch’s normal pace (≤30) + identical baskets (≤15). High ≥ 60, Medium 35–59, Low < 35.',
      font({ size: 9, italic: true, color: C.muted }), { wrapText: true, vertical: 'top' });

    ws.pageSetup = { orientation: 'landscape', paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0,
      margins: { left: 0.3, right: 0.3, top: 0.4, bottom: 0.4, header: 0.2, footer: 0.2 }, printArea: 'A1:N' + row };
    ws.headerFooter = { oddFooter: '&L&8Half Million · Mugs SOP Monitor&R&8Page &P of &N' };
  }

  /* ======================================================================
     Build workbook
     ====================================================================== */
  function build(R, meta) {
    var ExcelJS = root.ExcelJS;
    var wb = new ExcelJS.Workbook();
    wb.creator = 'Half Million · Mugs SOP Monitor';
    wb.created = new Date();
    wb.title = 'Mugs SOP report ' + fmtDate(R.totals.start);

    var images = { branches: chartBranches(R.stores), hours: chartHours(R) };

    var tab = function (name, color) { return wb.addWorksheet(name, { properties: { tabColor: { argb: A(color) } } }); };
    dashboard(wb, tab('Dashboard', C.green), R, meta, images);

    table(tab('Branches', C.gold), [
      { key: 'status', header: 'Risk', width: 10, badge: true, center: true },
      { key: 'store', header: 'Branch', width: 34, bold: true },
      { key: 'invoices', header: 'Invoices', width: 10, num: true, fmt: '#,##0' },
      { key: 'mugs', header: 'Mugs', width: 9, num: true, fmt: '#,##0' },
      { key: 'sales', header: 'Net sales SAR', width: 14, num: true, fmt: '#,##0.00' },
      { key: 'directLines', header: 'Direct violations', width: 11, num: true },
      { key: 'directExcess', header: 'Mugs over (direct)', width: 13, num: true, bar: C.barDirect },
      { key: 'runs', header: 'Split runs', width: 10, num: true },
      { key: 'runsHigh', header: 'High runs', width: 9, num: true },
      { key: 'runsMedium', header: 'Medium runs', width: 10, num: true },
      { key: 'runMugs', header: 'Mugs in split runs', width: 12, num: true },
      { key: 'runExcess', header: 'Mugs over (split)', width: 13, num: true, bar: C.barSplit },
      { key: 'pctOver', header: '% of mugs over limit', width: 14, num: true, fmt: '0.0%', bar: C.barHigh, barMax: 1 },
      { key: 'zero', header: 'Zero-qty lines', width: 10, num: true }
    ], R.stores, { freezeCols: 2 });

    table(tab('Split runs', C.split), [
      { key: 'n', header: 'Run #', width: 7, num: true },
      { key: 'level', header: 'Risk', width: 10, badge: true, center: true },
      { key: 'score', header: 'Score', width: 9, num: true, bar: C.barDirect, barMax: 100 },
      { key: 'store', header: 'Branch', width: 30, bold: true },
      { key: 'label', header: 'Item / colour', width: 26 },
      { key: 'start', header: 'From', width: 18, fmt: 'yyyy-mm-dd hh:mm:ss' },
      { key: 'end', header: 'To', width: 18, fmt: 'yyyy-mm-dd hh:mm:ss' },
      { key: 'dur', header: 'Duration (min)', width: 10, num: true, fmt: '0.00' },
      { key: 'invoices', header: 'Invoices', width: 9, num: true },
      { key: 'mugs', header: 'Mugs', width: 8, num: true },
      { key: 'excess', header: 'Over limit', width: 9, num: true, bar: C.barSplit },
      { key: 'sales', header: 'SAR', width: 11, num: true, fmt: '#,##0.00' },
      { key: 'pace', header: '× normal pace', width: 11, num: true, fmt: '[>=100]"100+";0.0' },
      { key: 'same', header: 'Same basket', width: 11, num: true, fmt: '0%' },
      { key: 'basket', header: 'Most common basket', width: 60, wrap: true }
    ], R.runs.map(function (r, i) {
      return { n: i + 1, level: r.level, score: r.score, store: r.store, label: r.label, start: xlDate(r.start), end: xlDate(r.end),
        dur: r.durationMin, invoices: r.invoices, mugs: r.mugs, excess: r.excess, sales: r.sales, pace: Math.min(r.burst, 100),
        same: r.invoices >= 3 ? r.identical : null, basket: r.topBasket };
    }), { freezeCols: 4 });

    var inv = [];
    R.runs.forEach(function (r, i) {
      r.detail.forEach(function (d, j) {
        inv.push({ n: i + 1, level: r.level, store: r.store, label: r.label, time: xlDate(d.time),
          gap: j ? (d.time - r.detail[j - 1].time) / 60000 : null, qty: d.qty, iq: d.invoiceQty, sales: d.sales, basket: d.basket });
      });
    });
    table(tab('Split run invoices', C.split), [
      { key: 'n', header: 'Run #', width: 7, num: true },
      { key: 'level', header: 'Risk', width: 10, badge: true, center: true },
      { key: 'store', header: 'Branch', width: 30, bold: true },
      { key: 'label', header: 'Item / colour', width: 26 },
      { key: 'time', header: 'Time', width: 18, fmt: 'yyyy-mm-dd hh:mm:ss' },
      { key: 'gap', header: 'Gap (min)', width: 9, num: true, fmt: '0.00' },
      { key: 'qty', header: 'This item', width: 9, num: true },
      { key: 'iq', header: 'Invoice qty', width: 10, num: true },
      { key: 'sales', header: 'Invoice SAR', width: 11, num: true, fmt: '#,##0.00' },
      { key: 'basket', header: 'Full invoice', width: 70, wrap: true }
    ], inv, { freezeCols: 3 });

    table(tab('Direct violations', C.High), [
      { key: 'store', header: 'Branch', width: 30, bold: true },
      { key: 'label', header: 'Item / colour', width: 26 },
      { key: 'items', header: 'Items', width: 30 },
      { key: 'time', header: 'Time', width: 18, fmt: 'yyyy-mm-dd hh:mm:ss' },
      { key: 'qty', header: 'Qty', width: 8, num: true, bar: C.barHigh },
      { key: 'excess', header: 'Over limit', width: 9, num: true },
      { key: 'sales', header: 'SAR', width: 11, num: true, fmt: '#,##0.00' },
      { key: 'basket', header: 'Full invoice', width: 70, wrap: true }
    ], R.direct.map(function (d) {
      return { store: d.store, label: d.label, items: d.items, time: xlDate(d.time), qty: d.qty, excess: d.excess, sales: d.sales, basket: d.basket };
    }), { freezeCols: 1 });

    table(tab('Zero-qty lines', C.muted), [
      { key: 'store', header: 'Branch', width: 30, bold: true },
      { key: 'item', header: 'Item', width: 36 },
      { key: 'time', header: 'Time', width: 18, fmt: 'yyyy-mm-dd hh:mm:ss' },
      { key: 'qty', header: 'Qty', width: 8, num: true },
      { key: 'sales', header: 'SAR', width: 10, num: true, fmt: '#,##0.00' }
    ], R.zero.map(function (z) { return { store: z.store, item: z.item, time: xlDate(z.time), qty: z.qty, sales: z.sales }; }));

    var o = R.options;
    table(tab('Settings', C.muted), [
      { key: 'k', header: 'Setting', width: 34, bold: true },
      { key: 'v', header: 'Value', width: 70, wrap: true }
    ], [
      { k: 'Source file', v: meta.fileName },
      { k: 'Report date', v: meta.reportDate || fmtDate(R.totals.start) },
      { k: 'Period covered', v: stamp(R.totals.start) + ' → ' + stamp(R.totals.end) },
      { k: 'Limit per customer', v: o.limit },
      { k: 'Max gap between invoices (min)', v: o.gapMinutes },
      { k: 'Min invoices in a run', v: o.minInvoices },
      { k: '“Same colour” means', v: o.groupBy === 'colour' ? 'Same colour across items' : 'Same item / SKU' },
      { k: 'Invoice key', v: 'Branch + closing time (to the millisecond); the export has no invoice number' },
      { k: 'Risk score', v: 'Invoices in run (≤30) + mugs over limit (≤25) + burst vs branch normal pace (≤30) + identical baskets (≤15). High ≥ 60, Medium 35–59, Low < 35.' },
      { k: 'Limitations', v: 'Without card, customer or cashier ID a split run is likely, not proven. Verify High runs with the branch, POS logs and CCTV.' }
    ], { rowHeight: 30 });

    return wb.xlsx.writeBuffer().then(function (buf) {
      return new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    });
  }

  root.MugsReport = { build: build };
})(typeof self !== 'undefined' ? self : this);
