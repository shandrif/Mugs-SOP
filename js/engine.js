/*
 * Mugs SOP detection engine.
 * Pure functions, no DOM: used by the browser app and by the Node test harness.
 *
 * Input: rows from the POS "Sales By Product" export (array of arrays, as
 * produced by SheetJS sheet_to_json(ws, {header: 1, raw: true})).
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.MugsEngine = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var DEFAULTS = {
    limit: 2,              // max mugs of one colour/SKU per customer
    gapMinutes: 3,         // max minutes between invoices to chain them into one run
    minInvoices: 2,        // min invoices in a run before it can be flagged
    groupBy: 'sku'         // 'sku' or 'colour'
  };

  // Colour words recognised in item names (English part). Order matters: first match wins.
  var COLOURS = [
    ['olive green', 'Olive Green'], ['beige', 'Beige'], ['green', 'Green'], ['black', 'Black'],
    ['yellow', 'Yellow'], ['orange', 'Orange'], ['blue', 'Blue'], ['espresso', 'Espresso'],
    ['cherry', 'Cherry'], ['jacaranda', 'Jacaranda'], ['khozama', 'Khozama'], ['diriyah', 'Diriyah'],
    ['play', 'Play']
  ];

  function clean(s) {
    if (s === null || s === undefined) return '';
    return String(s).replace(/[​-‍﻿]/g, '').replace(/\s+/g, ' ').trim();
  }

  function englishPart(item) {
    var i = item.indexOf('/');
    return clean(i >= 0 ? item.slice(0, i) : item);
  }

  function colourOf(item) {
    var en = englishPart(item).toLowerCase();
    for (var i = 0; i < COLOURS.length; i++) if (en.indexOf(COLOURS[i][0]) >= 0) return COLOURS[i][1];
    return englishPart(item) || item;
  }

  // Excel serial (1900 system) or string/Date -> epoch ms, treated as naive local time (UTC fields).
  function toMs(v) {
    if (v === null || v === undefined || v === '') return null;
    if (typeof v === 'number') return Math.round((v - 25569) * 86400000);
    if (v instanceof Date) return v.getTime() - v.getTimezoneOffset() * 60000;
    var m = String(v).trim().match(/^(\d{4})-(\d{1,2})-(\d{1,2})[ T](\d{1,2}):(\d{2})(?::(\d{2})(?:\.(\d+))?)?/);
    if (m) {
      var ms = m[7] ? Math.round(parseFloat('0.' + m[7]) * 1000) : 0;
      return Date.UTC(+m[1], +m[2] - 1, +m[3], +m[4], +m[5], +(m[6] || 0), ms);
    }
    m = String(v).trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})[ ,]+(\d{1,2}):(\d{2})(?::(\d{2})(?:\.(\d+))?)?\s*(AM|PM)?/i);
    if (m) {
      var y = +m[3] < 100 ? 2000 + +m[3] : +m[3], h = +m[4];
      if (m[8]) { var pm = m[8].toUpperCase() === 'PM'; if (h === 12) h = pm ? 12 : 0; else if (pm) h += 12; }
      var ms2 = m[7] ? Math.round(parseFloat('0.' + m[7]) * 1000) : 0;
      return Date.UTC(y, +m[1] - 1, +m[2], h, +m[5], +(m[6] || 0), ms2);
    }
    return null;
  }

  // Accounting day -> "YYYY-MM-DD" (Excel serial, Date, "2026-11-18 00:00:00", "11/18/2026" or "18.11.2026").
  function toDay(v) {
    if (v === null || v === undefined || v === '') return '';
    if (typeof v === 'number' && isFinite(v)) return new Date(Math.round((v - 25569) * 86400000)).toISOString().slice(0, 10);
    if (v instanceof Date) return new Date(v.getTime() - v.getTimezoneOffset() * 60000).toISOString().slice(0, 10);
    var t = String(v).trim(), m;
    if ((m = t.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/))) return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
    if ((m = t.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/))) { var y = +m[3] < 100 ? 2000 + +m[3] : +m[3]; return y + '-' + ('0' + m[1]).slice(-2) + '-' + ('0' + m[2]).slice(-2); }
    if ((m = t.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})/))) return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
    return '';
  }

  function fmtTime(ms, withDate) {
    var d = new Date(ms), p = function (n) { return (n < 10 ? '0' : '') + n; };
    var t = p(d.getUTCHours()) + ':' + p(d.getUTCMinutes()) + ':' + p(d.getUTCSeconds());
    return withDate ? d.getUTCFullYear() + '-' + p(d.getUTCMonth() + 1) + '-' + p(d.getUTCDate()) + ' ' + t : t;
  }

  /* ---------- 1. Parse the export into clean sale lines ---------- */
  function parseRows(rows) {
    var hdrIdx = -1, col = {};
    for (var r = 0; r < Math.min(rows.length, 40); r++) {
      var row = (rows[r] || []).map(clean);
      var lower = row.map(function (c) { return c.toLowerCase(); });
      if (lower.indexOf('store') >= 0 && lower.some(function (c) { return c.indexOf('closing time') >= 0; })) {
        hdrIdx = r;
        lower.forEach(function (c, i) {
          if (c === 'store') col.store = i;
          else if (c.indexOf('accounting day') >= 0) col.day = i;
          else if (c.indexOf('item code') >= 0) col.code = i;
          else if (c === 'item') col.item = i;
          else if (c.indexOf('closing time') >= 0) col.close = i;
          else if (c.indexOf('net sales') >= 0) col.sales = i;
          else if (c.indexOf('number of') >= 0 || c === 'qty' || c === 'quantity') col.qty = i;
          else if (c.indexOf('invoice') >= 0 || c.indexOf('receipt') >= 0 || c.indexOf('check') >= 0) col.invoice = i;
          else if (c.indexOf('cashier') >= 0 || c.indexOf('employee') >= 0) col.cashier = i;
          else if (c.indexOf('card') >= 0 || c.indexOf('customer') >= 0 || c.indexOf('mobile') >= 0) col.customer = i;
        });
        break;
      }
    }
    if (hdrIdx < 0) throw new Error('Could not find the header row (expected columns "Store" and "Closing time").');
    ['store', 'item', 'close', 'qty'].forEach(function (k) {
      if (col[k] === undefined) throw new Error('Missing required column: ' + k);
    });

    var lines = [], warnings = [], cur = { store: '', day: null, code: '', item: '' };
    for (var i = hdrIdx + 1; i < rows.length; i++) {
      var x = rows[i] || [];
      var st = clean(x[col.store]);
      if (/grand total/i.test(st)) continue;
      if (st && !/total$/i.test(st)) cur.store = st;
      if (col.day !== undefined && x[col.day] !== undefined && x[col.day] !== null && x[col.day] !== '') cur.day = x[col.day];
      var code = col.code !== undefined ? clean(x[col.code]) : '';
      var item = clean(x[col.item]);
      if (code) cur.code = code;
      if (item) { cur.item = item; if (!code && col.code !== undefined) cur.code = ''; }
      var closeMs = toMs(x[col.close]);
      var qty = Number(x[col.qty]);
      if (closeMs === null || !isFinite(qty)) continue; // subtotal / blank rows
      if (qty <= 0) continue; // zero-quantity lines (voids, comps, returns) are not sales
      if (!cur.store || !cur.item) { warnings.push('Row ' + (i + 1) + ': sale without store/item, skipped'); continue; }
      lines.push({
        row: i + 1,
        store: cur.store,
        code: cur.code,
        item: cur.item,
        itemEn: englishPart(cur.item),
        colour: colourOf(cur.item),
        time: closeMs,
        day: toDay(cur.day),
        sales: Number(x[col.sales]) || 0,
        qty: qty,
        invoice: col.invoice !== undefined ? clean(x[col.invoice]) : '',
        cashier: col.cashier !== undefined ? clean(x[col.cashier]) : '',
        customer: col.customer !== undefined ? clean(x[col.customer]) : ''
      });
    }
    var reportDate = null;
    for (var k = 0; k < hdrIdx; k++) {
      var t = (rows[k] || []).map(clean).join(' ');
      var m = t.match(/date:\s*([0-9\/\-]+)/i);
      if (m) reportDate = m[1];
    }
    // Accounting days in the file (a weekly export holds several).
    var dayset = {};
    lines.forEach(function (l) { if (l.day) dayset[l.day] = 1; });
    var days = Object.keys(dayset).sort();
    return { lines: lines, warnings: warnings, reportDate: reportDate, days: days, hasInvoiceCol: col.invoice !== undefined };
  }

  /* ---------- 2. Analyse ---------- */
  function analyse(lines, opts) {
    // Ignore zero-quantity lines, including in data saved before they were dropped at upload.
    lines = lines.filter(function (l) { return l.qty > 0; });
    var o = {};
    Object.keys(DEFAULTS).forEach(function (k) { o[k] = (opts && opts[k] !== undefined) ? opts[k] : DEFAULTS[k]; });
    var gapMs = o.gapMinutes * 60000;
    var keyOf = function (l) { return o.groupBy === 'colour' ? l.colour : (l.code || l.itemEn); };
    var labelOf = function (l) { return o.groupBy === 'colour' ? l.colour : l.itemEn; };

    // Invoice key: explicit invoice number if present, else store + closing time.
    var invoices = {}, invList = [];
    lines.forEach(function (l) {
      l.invKey = l.store + '|' + (l.invoice || l.time);
      var inv = invoices[l.invKey];
      if (!inv) {
        inv = invoices[l.invKey] = { key: l.invKey, store: l.store, time: l.time, invoice: l.invoice, lines: [], qty: 0, sales: 0 };
        invList.push(inv);
      }
      inv.lines.push(l); inv.qty += l.qty; inv.sales += l.sales;
    });
    invList.forEach(function (inv) {
      inv.basket = inv.lines.filter(function (l) { return l.qty > 0; })
        .map(function (l) { return labelOf(l) + ' x' + l.qty; }).sort().join(' + ');
    });

    var tMin = Infinity, tMax = -Infinity;
    lines.forEach(function (l) { if (l.time < tMin) tMin = l.time; if (l.time > tMax) tMax = l.time; });
    var spanMin = Math.max(60, (tMax - tMin) / 60000);

    // Aggregate per invoice + group key (two SKUs of the same colour merge in colour mode).
    var cells = {};
    lines.forEach(function (l) {
      var k = l.store + '||' + keyOf(l) + '||' + l.invKey;
      var c = cells[k];
      if (!c) c = cells[k] = { store: l.store, key: keyOf(l), label: labelOf(l), invKey: l.invKey, time: l.time, qty: 0, sales: 0, items: {} };
      c.qty += l.qty; c.sales += l.sales; c.items[l.itemEn] = true;
    });
    var cellList = Object.keys(cells).map(function (k) { return cells[k]; });

    // Rule 1: direct violations.
    var direct = cellList.filter(function (c) { return c.qty > o.limit; }).map(function (c) {
      return {
        store: c.store, label: c.label, items: Object.keys(c.items).join(', '), time: c.time,
        qty: c.qty, excess: c.qty - o.limit, sales: c.sales, basket: invoices[c.invKey].basket
      };
    }).sort(function (a, b) { return b.qty - a.qty || a.time - b.time; });

    // Rule 2: split-invoice runs.
    var byGroup = {};
    cellList.forEach(function (c) {
      if (c.qty <= 0) return;
      var g = c.store + '||' + c.key;
      (byGroup[g] = byGroup[g] || []).push(c);
    });
    var runs = [];
    Object.keys(byGroup).forEach(function (g) {
      var arr = byGroup[g].sort(function (a, b) { return a.time - b.time; });
      var groupInvoices = arr.length;
      var chain = [arr[0]];
      var flush = function () {
        if (chain.length >= o.minInvoices) {
          var total = 0, mx = 0;
          chain.forEach(function (c) { total += c.qty; if (c.qty > mx) mx = c.qty; });
          if (total > o.limit && mx <= o.limit) runs.push(buildRun(chain, total, groupInvoices));
        }
      };
      for (var i = 1; i < arr.length; i++) {
        if (arr[i].time - arr[i - 1].time <= gapMs) chain.push(arr[i]);
        else { flush(); chain = [arr[i]]; }
      }
      flush();
    });

    function buildRun(chain, total, groupInvoices) {
      var start = chain[0].time, end = chain[chain.length - 1].time;
      var n = chain.length;
      var durMin = (end - start) / 60000;
      // Store-specific baseline: this store's invoices for this colour/SKU outside the run, spread over the day.
      var others = Math.max(groupInvoices - n, 0.5);
      var expected = others / spanMin * (durMin + o.gapMinutes);
      var burst = Math.min(1000, n / Math.max(expected, 0.001));
      // Identical baskets: share of invoices in the run with the most common full-basket signature.
      var sig = {}, top = 0, topSig = '';
      chain.forEach(function (c) {
        var b = invoices[c.invKey].basket; sig[b] = (sig[b] || 0) + 1;
        if (sig[b] > top) { top = sig[b]; topSig = b; }
      });
      var identical = top / n;
      var excess = total - o.limit;
      var pts = {
        invoices: Math.min(30, (n - 1) * 5),
        excess: Math.min(25, excess * 2.5),
        burst: Math.max(0, Math.min(30, 15 * Math.log10(burst))),
        identical: n >= 3 ? Math.round(15 * identical) : 0
      };
      var score = Math.round(pts.invoices + pts.excess + pts.burst + pts.identical);
      var level = score >= 60 ? 'High' : score >= 35 ? 'Medium' : 'Low';
      return {
        store: chain[0].store, label: chain[0].label,
        items: Object.keys(chain.reduce(function (a, c) { Object.assign(a, c.items); return a; }, {})).join(', '),
        start: start, end: end, durationMin: durMin, invoices: n, mugs: total, excess: excess,
        sales: chain.reduce(function (a, c) { return a + c.sales; }, 0),
        expected: expected, burst: burst, identical: identical, topBasket: topSig,
        score: score, level: level, points: pts,
        detail: chain.map(function (c) {
          var inv = invoices[c.invKey];
          return { time: c.time, invoice: inv.invoice, qty: c.qty, basket: inv.basket, invoiceQty: inv.qty, sales: inv.sales };
        })
      };
    }
    runs.sort(function (a, b) { return b.score - a.score || b.mugs - a.mugs; });

    // Store league table.
    var stores = {};
    var S = function (name) {
      return stores[name] = stores[name] || {
        store: name, invoices: 0, lines: 0, mugs: 0, sales: 0,
        directLines: 0, directExcess: 0, runs: 0, runsHigh: 0, runsMedium: 0, runMugs: 0, runExcess: 0, maxScore: 0
      };
    };
    invList.forEach(function (inv) { var s = S(inv.store); s.invoices++; s.mugs += inv.qty; s.sales += inv.sales; s.lines += inv.lines.length; });
    direct.forEach(function (d) { var s = S(d.store); s.directLines++; s.directExcess += d.excess; });
    runs.forEach(function (r) {
      var s = S(r.store); s.runs++; s.runMugs += r.mugs; s.runExcess += r.excess;
      if (r.level === 'High') s.runsHigh++; else if (r.level === 'Medium') s.runsMedium++;
      if (r.score > s.maxScore) s.maxScore = r.score;
    });
    var storeList = Object.keys(stores).map(function (k) {
      var s = stores[k];
      s.flaggedExcess = s.directExcess + s.runExcess;
      s.pctOver = s.mugs > 0 ? s.flaggedExcess / s.mugs : 0;
      s.status = (s.runsHigh > 0 || s.directExcess >= 20) ? 'High'
        : (s.runsMedium > 0 || s.directLines > 0) ? 'Medium'
        : (s.runs > 0) ? 'Low' : 'Clean';
      return s;
    }).sort(function (a, b) {
      var rank = { High: 0, Medium: 1, Low: 2, Clean: 3 };
      return rank[a.status] - rank[b.status] || b.flaggedExcess - a.flaggedExcess;
    });

    var totals = {
      lines: lines.length, invoices: invList.length, stores: storeList.length,
      mugs: lines.reduce(function (a, l) { return a + l.qty; }, 0),
      sales: lines.reduce(function (a, l) { return a + l.sales; }, 0),
      directLines: direct.length, directExcess: direct.reduce(function (a, d) { return a + d.excess; }, 0),
      directStores: storeList.filter(function (s) { return s.directLines > 0; }).length,
      runs: runs.length, runsHigh: runs.filter(function (r) { return r.level === 'High'; }).length,
      runsMedium: runs.filter(function (r) { return r.level === 'Medium'; }).length,
      runMugs: runs.reduce(function (a, r) { return a + r.mugs; }, 0),
      runStores: storeList.filter(function (s) { return s.runs > 0; }).length,
      start: tMin, end: tMax
    };

    return { options: o, totals: totals, stores: storeList, direct: direct, runs: runs };
  }

  return { DEFAULTS: DEFAULTS, toDay: toDay, parseRows: parseRows, analyse: analyse, colourOf: colourOf, toMs: toMs, fmtTime: fmtTime };
});
