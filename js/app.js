/* Mugs SOP Monitor – browser UI */
(function () {
  'use strict';
  var E = window.MugsEngine;
  var $ = function (id) { return document.getElementById(id); };

  var state = {
    parsed: null, result: null, fileName: '',
    tab: 'stores', sort: {}, open: {},
    filter: { store: '', level: '', q: '' }
  };

  /* ---------- helpers ---------- */
  function esc(s) {
    return String(s === null || s === undefined ? '' : s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }
  var nf0 = new Intl.NumberFormat('en-US', { maximumFractionDigits: 0 });
  var nf2 = new Intl.NumberFormat('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  function n0(v) { return nf0.format(v); }
  function sar(v) { return nf2.format(v); }
  function pct(v) { return (v * 100).toFixed(1) + '%'; }
  function t(ms) { return E.fmtTime(ms, false); }
  function dt(ms) { return E.fmtTime(ms, true); }
  function mins(m) { return m < 1 ? Math.round(m * 60) + 's' : m.toFixed(1) + ' min'; }
  function badge(level) { return '<span class="badge ' + esc(level) + '">' + esc(level) + '</span>'; }

  /* ---------- file loading ---------- */
  var drop = $('drop'), fileInput = $('file');
  $('btnPick').addEventListener('click', function () { fileInput.click(); });
  drop.addEventListener('keydown', function (e) { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); fileInput.click(); } });
  fileInput.addEventListener('change', function () { if (fileInput.files[0]) load(fileInput.files[0]); });
  ['dragenter', 'dragover'].forEach(function (ev) {
    drop.addEventListener(ev, function (e) { e.preventDefault(); drop.classList.add('over'); });
  });
  ['dragleave', 'drop'].forEach(function (ev) {
    drop.addEventListener(ev, function (e) { e.preventDefault(); drop.classList.remove('over'); });
  });
  drop.addEventListener('drop', function (e) { var f = e.dataTransfer.files[0]; if (f) load(f); });

  function showError(msg) { var el = $('error'); el.textContent = msg; el.hidden = !msg; }

  function load(file) {
    showError('');
    var reader = new FileReader();
    reader.onload = function (e) {
      try {
        var wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array', raw: false, cellDates: false });
        var ws = wb.Sheets[wb.SheetNames[0]];
        var rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
        var parsed = E.parseRows(rows);
        if (!parsed.lines.length) throw new Error('No sale lines found in this file.');
        state.parsed = parsed; state.fileName = file.name; state.open = {};
        state.filter = { store: '', level: '', q: '' };
        $('fStore').value = ''; $('fLevel').value = ''; $('fSearch').value = '';
        run();
      } catch (err) {
        console.error(err);
        showError('Could not read “' + file.name + '”: ' + err.message);
      }
    };
    reader.readAsArrayBuffer(file);
  }

  /* ---------- settings ---------- */
  function readOptions() {
    var lim = parseInt($('optLimit').value, 10), gap = parseFloat($('optGap').value), mi = parseInt($('optMinInv').value, 10);
    return {
      limit: isFinite(lim) && lim > 0 ? lim : 2,
      gapMinutes: isFinite(gap) && gap > 0 ? gap : 3,
      minInvoices: isFinite(mi) && mi >= 2 ? mi : 2,
      groupBy: $('optGroup').value
    };
  }
  try {
    var saved = JSON.parse(localStorage.getItem('mugsSopOptions') || 'null');
    if (saved) {
      if (saved.limit) $('optLimit').value = saved.limit;
      if (saved.gapMinutes) $('optGap').value = saved.gapMinutes;
      if (saved.minInvoices) $('optMinInv').value = saved.minInvoices;
      if (saved.groupBy) $('optGroup').value = saved.groupBy;
    }
  } catch (e) { /* storage unavailable */ }
  ['optLimit', 'optGap', 'optMinInv', 'optGroup'].forEach(function (id) {
    $(id).addEventListener('change', function () {
      try { localStorage.setItem('mugsSopOptions', JSON.stringify(readOptions())); } catch (e) { /* ignore */ }
      $('hdrLimit').textContent = readOptions().limit;
      if (state.parsed) { state.open = {}; run(); }
    });
  });
  $('hdrLimit').textContent = readOptions().limit;

  /* ---------- analysis ---------- */
  function run() {
    var res = E.analyse(state.parsed.lines, readOptions());
    state.result = res;
    drop.classList.add('compact');
    $('empty').hidden = true;
    $('results').hidden = false;
    $('btnExport').disabled = false;
    var T = res.totals;
    $('fileInfo').hidden = false;
    $('fileInfo').textContent = state.fileName + ' · report date ' + (state.parsed.reportDate || '—') +
      ' · ' + dt(T.start) + ' → ' + dt(T.end) +
      (state.parsed.warnings.length ? ' · ' + state.parsed.warnings.length + ' rows skipped' : '');

    var fs = $('fStore');
    fs.innerHTML = '<option value="">All branches</option>' + res.stores.map(function (s) {
      return '<option value="' + esc(s.store) + '">' + esc(s.store) + '</option>';
    }).join('');
    fs.value = state.filter.store;

    renderKpis();
    $('cntRuns').textContent = T.runs;
    $('cntDirect').textContent = T.directLines;
    $('cntZero').textContent = T.zero;
    renderTable();
  }

  function renderKpis() {
    var T = state.result.totals;
    var highStores = state.result.stores.filter(function (s) { return s.status === 'High'; }).length;
    var k = [
      { label: 'Mugs sold', value: n0(T.mugs), note: n0(T.invoices) + ' invoices · ' + T.stores + ' branches' },
      { label: 'Net sales (SAR)', value: sar(T.sales), note: n0(T.lines) + ' sale lines' },
      { label: 'Direct violations', value: n0(T.directLines), note: n0(T.directExcess) + ' mugs over limit · ' + T.directStores + ' branches', cls: 'high' },
      { label: 'Split-invoice runs', value: n0(T.runs), note: T.runsHigh + ' high · ' + T.runsMedium + ' medium · ' + T.runStores + ' branches', cls: 'med' },
      { label: 'High-risk branches', value: n0(highStores), note: 'need follow-up', cls: 'high' },
      { label: 'Zero-qty lines', value: n0(T.zero), note: 'voids / returns to review' }
    ];
    $('kpis').innerHTML = k.map(function (x) {
      return '<div class="kpi ' + (x.cls || '') + '"><div class="label">' + esc(x.label) + '</div><div class="value">' +
        esc(x.value) + '</div><div class="note">' + esc(x.note) + '</div></div>';
    }).join('');
  }

  /* ---------- tables ---------- */
  var COLS = {
    stores: [
      { k: 'status', h: 'Risk', f: function (r) { return badge(r.status); }, sortVal: function (r) { return { High: 3, Medium: 2, Low: 1, Clean: 0 }[r.status]; } },
      { k: 'store', h: 'Branch' },
      { k: 'invoices', h: 'Invoices', num: 1, f: function (r) { return n0(r.invoices); } },
      { k: 'mugs', h: 'Mugs', num: 1, f: function (r) { return n0(r.mugs); } },
      { k: 'directLines', h: 'Direct viol.', num: 1 },
      { k: 'directExcess', h: 'Mugs over (direct)', num: 1 },
      { k: 'runs', h: 'Split runs', num: 1 },
      { k: 'runsHigh', h: 'High runs', num: 1 },
      { k: 'runExcess', h: 'Mugs over (split)', num: 1 },
      { k: 'pctOver', h: '% of mugs over limit', num: 1, f: function (r) {
        return pct(r.pctOver) + '<span class="bar" style="width:' + Math.round(Math.min(1, r.pctOver * 2) * 60) + 'px"></span>'; } },
      { k: 'zero', h: 'Zero-qty', num: 1 }
    ],
    runs: [
      { k: 'level', h: 'Risk', f: function (r) { return badge(r.level); }, sortVal: function (r) { return r.score; } },
      { k: 'score', h: 'Score', num: 1 },
      { k: 'store', h: 'Branch' },
      { k: 'label', h: 'Item / colour' },
      { k: 'start', h: 'From', f: function (r) { return t(r.start); } },
      { k: 'end', h: 'To', f: function (r) { return t(r.end); } },
      { k: 'durationMin', h: 'Duration', num: 1, f: function (r) { return mins(r.durationMin); } },
      { k: 'invoices', h: 'Invoices', num: 1 },
      { k: 'mugs', h: 'Mugs', num: 1 },
      { k: 'excess', h: 'Over limit', num: 1 },
      { k: 'burst', h: '× normal pace', num: 1, f: function (r) { return r.burst >= 100 ? '100+' : r.burst.toFixed(1); } },
      { k: 'identical', h: 'Same basket', num: 1, f: function (r) { return r.invoices >= 3 ? pct(r.identical) : '—'; } }
    ],
    direct: [
      { k: 'store', h: 'Branch' },
      { k: 'label', h: 'Item / colour' },
      { k: 'time', h: 'Time', f: function (r) { return t(r.time); } },
      { k: 'qty', h: 'Qty', num: 1 },
      { k: 'excess', h: 'Over limit', num: 1 },
      { k: 'sales', h: 'SAR', num: 1, f: function (r) { return sar(r.sales); } },
      { k: 'basket', h: 'Full invoice', cls: 'wrap-cell' }
    ],
    zero: [
      { k: 'store', h: 'Branch' },
      { k: 'item', h: 'Item' },
      { k: 'time', h: 'Time', f: function (r) { return t(r.time); } },
      { k: 'qty', h: 'Qty', num: 1 },
      { k: 'sales', h: 'SAR', num: 1, f: function (r) { return sar(r.sales); } }
    ]
  };
  var DEFAULT_SORT = {
    stores: null, runs: null,
    direct: { k: 'qty', dir: -1 }, zero: { k: 'time', dir: 1 }
  };

  function dataFor(tab) {
    var R = state.result, F = state.filter, q = F.q.toLowerCase();
    var rows = tab === 'stores' ? R.stores : tab === 'runs' ? R.runs : tab === 'direct' ? R.direct : R.zero;
    return rows.filter(function (r) {
      if (F.store && r.store !== F.store) return false;
      if (F.level) {
        var lv = tab === 'stores' ? r.status : tab === 'runs' ? r.level : null;
        if (lv !== null && lv !== F.level) return false;
      }
      if (q) {
        var hay = (r.store + ' ' + (r.label || '') + ' ' + (r.item || '') + ' ' + (r.items || '') + ' ' + (r.basket || '')).toLowerCase();
        if (hay.indexOf(q) < 0) return false;
      }
      return true;
    });
  }

  function renderTable() {
    var tab = state.tab, cols = COLS[tab];
    var rows = dataFor(tab).slice();
    var s = state.sort[tab] || DEFAULT_SORT[tab];
    if (s) {
      var col = cols.filter(function (c) { return c.k === s.k; })[0];
      var get = col && col.sortVal ? col.sortVal : function (r) { return r[s.k]; };
      rows.sort(function (a, b) {
        var x = get(a), y = get(b);
        if (typeof x === 'string') return s.dir * x.localeCompare(y);
        return s.dir * (x - y);
      });
    }
    $('fLevel').disabled = (tab === 'direct' || tab === 'zero');
    $('rowCount').textContent = rows.length + ' rows' + (tab === 'stores' ? ' · click a branch to see its runs' : tab === 'runs' ? ' · click a run for invoice detail' : '');

    var html = '<table><thead><tr>' + cols.map(function (c) {
      var arrow = s && s.k === c.k ? (s.dir > 0 ? ' ▲' : ' ▼') : '';
      return '<th data-k="' + c.k + '" class="' + (c.num ? 'num' : '') + '">' + esc(c.h) + '<span class="arrow">' + arrow + '</span></th>';
    }).join('') + '</tr></thead><tbody>';

    var MAX = 2000;
    rows.slice(0, MAX).forEach(function (r, i) {
      var id = tab === 'runs' ? r.store + '|' + r.label + '|' + r.start : tab === 'stores' ? r.store : '';
      var clickable = tab === 'runs' || tab === 'stores';
      html += '<tr' + (clickable ? ' class="clickable" data-id="' + esc(id) + '"' : '') + '>' + cols.map(function (c) {
        var v = c.f ? c.f(r) : esc(r[c.k]);
        return '<td class="' + (c.num ? 'num ' : '') + (c.cls || '') + '">' + v + '</td>';
      }).join('') + '</tr>';
      if (tab === 'runs' && state.open[id]) html += '<tr class="detail"><td colspan="' + cols.length + '">' + runDetail(r) + '</td></tr>';
    });
    if (!rows.length) html += '<tr><td colspan="' + cols.length + '" class="muted">Nothing to show for these filters.</td></tr>';
    html += '</tbody></table>';
    if (rows.length > MAX) html += '<p class="muted" style="padding:8px 12px">Showing first ' + MAX + ' rows. Use filters or export to Excel for all.</p>';
    $('tableWrap').innerHTML = html;
  }

  function runDetail(r) {
    var o = state.result.options;
    var span = Math.max(1, r.end - r.start);
    var bars = r.detail.map(function (d) {
      return '<span title="' + esc(t(d.time) + ' · ' + d.qty) + '" style="left:calc(' + ((d.time - r.start) / span * 100).toFixed(2) +
        '% - 2px);height:' + Math.max(25, d.qty / o.limit * 100) + '%"></span>';
    }).join('');
    var why = [];
    why.push(r.invoices + ' invoices in ' + mins(r.durationMin) + ' (+' + r.points.invoices + ' pts)');
    why.push(r.excess + ' mugs over the limit of ' + o.limit + ' (+' + r.points.excess + ' pts)');
    why.push('normally ~' + r.expected.toFixed(2) + ' invoices of this item in such a window at this branch → ' +
      (r.burst >= 100 ? 'over 100' : r.burst.toFixed(1)) + '× normal pace (+' + Math.round(r.points.burst) + ' pts)');
    if (r.invoices >= 3) why.push(pct(r.identical) + ' of invoices have the identical basket “' + r.topBasket + '” (+' + r.points.identical + ' pts)');
    return '<div class="detail-grid"><span>Score <b>' + r.score + '</b></span><span>SAR <b>' + sar(r.sales) + '</b></span>' +
      '<span>Items: <b>' + esc(r.items) + '</b></span></div>' +
      '<ul style="margin:0 0 8px 18px;padding:0">' + why.map(function (w) { return '<li>' + esc(w) + '</li>'; }).join('') + '</ul>' +
      '<div class="timeline" aria-hidden="true">' + bars + '</div>' +
      '<div class="timeline-axis"><span>' + t(r.start) + '</span><span>' + t(r.end) + '</span></div>' +
      '<div class="inner"><table><thead><tr><th>#</th><th>Time</th><th>Gap</th><th class="num">This item</th><th class="num">Invoice qty</th><th class="num">Invoice SAR</th><th>Full invoice</th></tr></thead><tbody>' +
      r.detail.map(function (d, i) {
        var gap = i ? mins((d.time - r.detail[i - 1].time) / 60000) : '—';
        return '<tr><td>' + (i + 1) + '</td><td>' + t(d.time) + '</td><td>' + gap + '</td><td class="num">' + d.qty +
          '</td><td class="num">' + d.invoiceQty + '</td><td class="num">' + sar(d.sales) + '</td><td class="wrap-cell">' + esc(d.basket) + '</td></tr>';
      }).join('') + '</tbody></table></div>';
  }

  $('tableWrap').addEventListener('click', function (e) {
    var th = e.target.closest('th[data-k]');
    if (th && th.closest('.inner') === null) {
      var k = th.getAttribute('data-k'), cur = state.sort[state.tab] || DEFAULT_SORT[state.tab];
      state.sort[state.tab] = { k: k, dir: cur && cur.k === k ? -cur.dir : -1 };
      renderTable(); return;
    }
    var tr = e.target.closest('tr.clickable');
    if (!tr) return;
    var id = tr.getAttribute('data-id');
    if (state.tab === 'stores') {
      state.filter.store = id; $('fStore').value = id;
      switchTab('runs');
    } else {
      state.open[id] = !state.open[id]; renderTable();
    }
  });

  function switchTab(tab) {
    state.tab = tab;
    Array.prototype.forEach.call(document.querySelectorAll('.tabs button'), function (b) {
      b.classList.toggle('active', b.getAttribute('data-tab') === tab);
      b.setAttribute('aria-selected', b.getAttribute('data-tab') === tab);
    });
    renderTable();
  }
  Array.prototype.forEach.call(document.querySelectorAll('.tabs button'), function (b) {
    b.addEventListener('click', function () { switchTab(b.getAttribute('data-tab')); });
  });
  $('fStore').addEventListener('change', function () { state.filter.store = this.value; renderTable(); });
  $('fLevel').addEventListener('change', function () { state.filter.level = this.value; renderTable(); });
  $('fSearch').addEventListener('input', function () { state.filter.q = this.value; renderTable(); });

  /* ---------- export ---------- */
  $('btnExport').addEventListener('click', function () {
    var R = state.result, o = R.options, wb = XLSX.utils.book_new();
    var add = function (name, rows) { XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows.length ? rows : [{ '': 'none' }]), name); };
    add('Branches', R.stores.map(function (s) {
      return { Risk: s.status, Branch: s.store, Invoices: s.invoices, Mugs: s.mugs, 'Net sales SAR': +s.sales.toFixed(2),
        'Direct violations': s.directLines, 'Mugs over (direct)': s.directExcess, 'Split runs': s.runs, 'High runs': s.runsHigh,
        'Medium runs': s.runsMedium, 'Mugs in split runs': s.runMugs, 'Mugs over (split)': s.runExcess,
        '% mugs over limit': +(s.pctOver * 100).toFixed(2), 'Zero-qty lines': s.zero };
    }));
    add('Split runs', R.runs.map(function (r, i) {
      return { 'Run #': i + 1, Risk: r.level, Score: r.score, Branch: r.store, 'Item / colour': r.label, Items: r.items,
        From: dt(r.start), To: dt(r.end), 'Duration min': +r.durationMin.toFixed(2), Invoices: r.invoices, Mugs: r.mugs,
        'Over limit': r.excess, 'SAR': +r.sales.toFixed(2), 'x normal pace': +r.burst.toFixed(1),
        'Same basket %': r.invoices >= 3 ? +(r.identical * 100).toFixed(1) : null, 'Most common basket': r.topBasket };
    }));
    var inv = [];
    R.runs.forEach(function (r, i) {
      r.detail.forEach(function (d) {
        inv.push({ 'Run #': i + 1, Risk: r.level, Branch: r.store, 'Item / colour': r.label, Time: dt(d.time),
          'This item qty': d.qty, 'Invoice qty': d.invoiceQty, 'Invoice SAR': +d.sales.toFixed(2), 'Full invoice': d.basket });
      });
    });
    add('Split run invoices', inv);
    add('Direct violations', R.direct.map(function (d) {
      return { Branch: d.store, 'Item / colour': d.label, Items: d.items, Time: dt(d.time), Qty: d.qty,
        'Over limit': d.excess, SAR: +d.sales.toFixed(2), 'Full invoice': d.basket };
    }));
    add('Zero-qty lines', R.zero.map(function (z) {
      return { Branch: z.store, Item: z.item, Time: dt(z.time), Qty: z.qty, SAR: z.sales };
    }));
    add('Settings', [
      { Setting: 'Source file', Value: state.fileName },
      { Setting: 'Report date', Value: state.parsed.reportDate || '' },
      { Setting: 'Limit per customer', Value: o.limit },
      { Setting: 'Max gap between invoices (min)', Value: o.gapMinutes },
      { Setting: 'Min invoices in a run', Value: o.minInvoices },
      { Setting: 'Grouping', Value: o.groupBy === 'colour' ? 'Same colour across items' : 'Same item / SKU' }
    ]);
    var base = state.fileName.replace(/\.[^.]+$/, '');
    XLSX.writeFile(wb, 'SOP-flags_' + base + '.xlsx');
  });

  /* ---------- help ---------- */
  $('btnHelp').addEventListener('click', function () {
    var d = $('help');
    if (d.showModal) d.showModal(); else d.setAttribute('open', '');
  });
})();
