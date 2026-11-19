/* Mugs SOP Monitor – browser UI */
(function () {
  'use strict';
  var E = window.MugsEngine;
  var $ = function (id) { return document.getElementById(id); };
  var reduceMotion = window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches;

  var state = {
    parsed: null, result: null, fileName: '',
    tab: 'stores', sort: {}, open: {},
    filter: { store: '', level: '', q: '' },
    kpiValues: {}
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
  function plural(n, one, many) { return n0(n) + ' ' + (n === 1 ? one : many); }
  function longDate(ms) {
    return new Date(ms).toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric', timeZone: 'UTC' });
  }

  /* ---------- theme ---------- */
  $('btnTheme').addEventListener('click', function () {
    var root = document.documentElement;
    var cur = root.getAttribute('data-theme') ||
      (window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light');
    var next = cur === 'dark' ? 'light' : 'dark';
    root.setAttribute('data-theme', next);
    try { localStorage.setItem('mugsSopTheme', next); } catch (e) { /* ignore */ }
  });

  /* ---------- nav: solid once the hero is scrolled past ---------- */
  var nav = $('nav');
  function onScroll() { nav.classList.toggle('scrolled', window.scrollY > 24); }
  window.addEventListener('scroll', onScroll, { passive: true });
  onScroll();

  /* ---------- scroll reveal (fires once per element) ---------- */
  var revealObserver = 'IntersectionObserver' in window ? new IntersectionObserver(function (entries) {
    entries.forEach(function (en) {
      if (en.isIntersecting) { en.target.classList.add('in'); revealObserver.unobserve(en.target); }
    });
  }, { rootMargin: '0px 0px -60px 0px', threshold: 0.08 }) : null;
  function observeReveals(scope) {
    Array.prototype.forEach.call((scope || document).querySelectorAll('.reveal:not(.in)'), function (el) {
      if (revealObserver) revealObserver.observe(el); else el.classList.add('in');
    });
  }
  observeReveals(document);

  /* ---------- file loading ---------- */
  var drop = $('drop'), fileInput = $('file');
  drop.addEventListener('click', function () { fileInput.click(); });
  drop.addEventListener('keydown', function (e) { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); fileInput.click(); } });
  fileInput.addEventListener('click', function (e) { e.stopPropagation(); });
  fileInput.addEventListener('change', function () { if (fileInput.files[0]) load(fileInput.files[0]); fileInput.value = ''; });
  ['dragenter', 'dragover'].forEach(function (ev) {
    document.addEventListener(ev, function (e) { e.preventDefault(); drop.classList.add('over'); });
  });
  ['dragleave', 'drop'].forEach(function (ev) {
    document.addEventListener(ev, function (e) {
      e.preventDefault();
      if (ev === 'drop' || !e.relatedTarget) drop.classList.remove('over');
    });
  });
  document.addEventListener('drop', function (e) { var f = e.dataTransfer && e.dataTransfer.files[0]; if (f) load(f); });

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
        run(true);
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
      if (state.parsed) { state.open = {}; run(false); }
    });
  });
  $('hdrLimit').textContent = readOptions().limit;

  /* ---------- analysis ---------- */
  function run(isNewFile) {
    var res = E.analyse(state.parsed.lines, readOptions());
    state.result = res;
    var T = res.totals;
    document.body.classList.add('has-data');
    $('empty').hidden = true;
    $('results').hidden = false;
    $('btnExport').disabled = false;

    $('reportTitle').textContent = longDate(T.start);
    $('fileInfo').textContent = state.fileName + ' · ' + plural(T.stores, 'branch', 'branches') + ' · ' +
      dt(T.start) + ' → ' + dt(T.end) +
      (state.parsed.warnings.length ? ' · ' + plural(state.parsed.warnings.length, 'row', 'rows') + ' skipped' : '');

    var fs = $('fStore');
    fs.innerHTML = '<option value="">All branches</option>' + res.stores.map(function (s) {
      return '<option value="' + esc(s.store) + '">' + esc(s.store) + '</option>';
    }).join('');
    fs.value = state.filter.store;

    renderKpis(isNewFile);
    renderChart(isNewFile);
    $('cntRuns').textContent = n0(T.runs);
    $('cntDirect').textContent = n0(T.directLines);
    $('cntZero').textContent = n0(T.zero);
    renderTable();
    moveInk();
    if (isNewFile) {
      observeReveals($('results'));
      window.scrollTo({ top: 0, behavior: reduceMotion ? 'auto' : 'smooth' });
    }
  }

  /* ---------- KPIs: numbers tween from their previous value when results change ---------- */
  function renderKpis(isNewFile) {
    var T = state.result.totals;
    var highStores = state.result.stores.filter(function (s) { return s.status === 'High'; }).length;
    var k = [
      { id: 'hi', label: 'High-risk branches', value: highStores, note: 'of ' + plural(T.stores, 'branch', 'branches') + ' need follow-up', cls: 'feature' },
      { id: 'dv', label: 'Direct violations', value: T.directLines, note: n0(T.directExcess) + ' mugs over limit · ' + plural(T.directStores, 'branch', 'branches'), cls: 'high' },
      { id: 'sr', label: 'Split-invoice runs', value: T.runs, note: T.runsHigh + ' high · ' + T.runsMedium + ' medium · ' + plural(T.runStores, 'branch', 'branches'), cls: 'med' },
      { id: 'mu', label: 'Mugs sold', value: T.mugs, note: plural(T.invoices, 'invoice', 'invoices') },
      { id: 'sa', label: 'Net sales', value: T.sales, money: true, note: 'SAR · ' + plural(T.lines, 'sale line', 'sale lines') },
      { id: 'zq', label: 'Zero-qty lines', value: T.zero, note: 'voids or returns to review' }
    ];
    var wrap = $('kpis');
    if (!wrap.children.length) {
      wrap.innerHTML = k.map(function (x) {
        return '<div class="kpi ' + (x.cls || '') + '" data-k="' + x.id + '"><div class="label">' +
          (x.cls === 'high' || x.cls === 'med' ? '<i class="dot" aria-hidden="true"></i>' : '') + esc(x.label) +
          '</div><div class="value">0</div><div class="note"></div></div>';
      }).join('');
    }
    k.forEach(function (x) {
      var el = wrap.querySelector('[data-k="' + x.id + '"]');
      el.querySelector('.note').textContent = x.note;
      var from = isNewFile ? 0 : (state.kpiValues[x.id] || 0);
      tween(el.querySelector('.value'), from, x.value, x.money, isNewFile ? 900 : 400);
      state.kpiValues[x.id] = x.value;
    });
  }
  function tween(node, from, to, money, dur) {
    var fmt = function (v) { return money ? nf0.format(v) : n0(Math.round(v)); };
    if (reduceMotion || from === to) { node.textContent = money ? sar(to) : fmt(to); return; }
    var start = null;
    node._tw = (node._tw || 0) + 1;
    var token = node._tw;
    function step(ts) {
      if (token !== node._tw) return; // a newer tween took over
      if (start === null) start = ts;
      var p = Math.min(1, (ts - start) / dur);
      var e = 1 - Math.pow(1 - p, 4); // ease-out quart
      node.textContent = fmt(from + (to - from) * e);
      if (p < 1) requestAnimationFrame(step); else node.textContent = money ? sar(to) : fmt(to);
    }
    requestAnimationFrame(step);
  }

  /* ---------- chart: excess mugs by branch, stacked direct + split ---------- */
  var chartObserver = null;
  function niceMax(v) {
    if (v <= 10) return 10;
    var p = Math.pow(10, Math.floor(Math.log10(v))), m = v / p;
    return (m <= 2 ? 2 : m <= 2.5 ? 2.5 : m <= 5 ? 5 : 10) * p;
  }
  function renderChart(animate) {
    var data = state.result.stores.filter(function (s) { return s.flaggedExcess > 0; })
      .sort(function (a, b) { return b.flaggedExcess - a.flaggedExcess; }).slice(0, 10);
    var el = $('chart');
    if (!data.length) { el.innerHTML = '<p class="muted small">No mugs over the limit with these rules.</p>'; return; }
    var max = niceMax(data[0].flaggedExcess);
    var html = data.map(function (s, i) {
      var segs = '';
      if (s.directExcess > 0) segs += '<span class="seg direct" style="--g:' + s.directExcess + '" data-series="direct"></span>';
      if (s.runExcess > 0) segs += '<span class="seg split" style="--g:' + s.runExcess + '" data-series="split"></span>';
      // non-uniform stagger: shrinking gaps, total well under the bar duration
      var d = Math.round(220 * Math.sqrt(i / 10));
      return '<button type="button" class="bar-row" data-store="' + esc(s.store) + '" style="--d:' + d + 'ms">' +
        '<span class="bar-label">' + esc(s.store) + '<small>' + badgeText(s.status) + '</small></span>' +
        '<span class="bar-track"><span class="bar-stack" style="--f:' + (s.flaggedExcess / max).toFixed(4) + '">' + segs + '</span>' +
        '<span class="bar-val">' + n0(s.flaggedExcess) + '</span></span></button>';
    }).join('');
    html += '<div class="chart-axis" aria-hidden="true"><span></span><div><span>0</span><span>' + n0(max / 4) + '</span><span>' +
      n0(max / 2) + '</span><span>' + n0(max * 3 / 4) + '</span><span>' + n0(max) + '</span></div></div>';
    el.innerHTML = html;

    if (animate && !reduceMotion && 'IntersectionObserver' in window) {
      el.classList.remove('play'); el.classList.add('armed');
      if (chartObserver) chartObserver.disconnect();
      chartObserver = new IntersectionObserver(function (entries) {
        if (entries[0].isIntersecting) {
          chartObserver.disconnect();
          el.getBoundingClientRect(); // commit the armed state before playing
          requestAnimationFrame(function () { el.classList.add('play'); el.classList.remove('armed'); });
        }
      }, { rootMargin: '0px 0px -80px 0px', threshold: 0.2 });
      chartObserver.observe(el);
    } else if (animate) {
      el.classList.remove('armed');
    }
  }
  function badgeText(level) { return level === 'Clean' ? 'Clean' : level + ' risk'; }

  var tip = $('tip');
  $('chart').addEventListener('pointermove', function (e) {
    var row = e.target.closest('.bar-row');
    if (!row) { tip.hidden = true; return; }
    var s = state.result.stores.filter(function (x) { return x.store === row.getAttribute('data-store'); })[0];
    if (!s) return;
    var hl = e.target.closest('.seg') ? e.target.closest('.seg').getAttribute('data-series') : '';
    tip.innerHTML = '<div style="font-weight:700;margin-bottom:4px">' + esc(s.store) + '</div>' +
      '<div class="t-row"' + (hl === 'split' ? ' style="opacity:.6"' : '') + '><i class="sw sw-direct"></i>Direct: <b>' + n0(s.directExcess) + '</b> over, ' + plural(s.directLines, 'invoice', 'invoices') + '</div>' +
      '<div class="t-row"' + (hl === 'direct' ? ' style="opacity:.6"' : '') + '><i class="sw sw-split"></i>Split: <b>' + n0(s.runExcess) + '</b> over, ' + plural(s.runs, 'run', 'runs') + '</div>' +
      '<div style="margin-top:4px;opacity:.75">' + pct(s.pctOver) + ' of ' + n0(s.mugs) + ' mugs over the limit</div>';
    tip.hidden = false;
    var x = Math.min(e.clientX + 14, window.innerWidth - tip.offsetWidth - 12);
    var y = e.clientY + 16 + tip.offsetHeight > window.innerHeight ? e.clientY - tip.offsetHeight - 12 : e.clientY + 16;
    tip.style.left = x + 'px'; tip.style.top = y + 'px';
  });
  $('chart').addEventListener('pointerleave', function () { tip.hidden = true; });
  $('chart').addEventListener('click', function (e) {
    var row = e.target.closest('.bar-row');
    if (!row) return;
    tip.hidden = true;
    openStore(row.getAttribute('data-store'));
  });

  function openStore(name) {
    state.filter.store = name; $('fStore').value = name;
    var s = state.result.stores.filter(function (x) { return x.store === name; })[0];
    switchTab(s && s.runs === 0 && s.directLines > 0 ? 'direct' : 'runs');
    var panel = document.querySelector('.table-panel');
    var top = panel.getBoundingClientRect().top + window.scrollY - 84;
    window.scrollTo({ top: top, behavior: reduceMotion ? 'auto' : 'smooth' });
  }

  /* ---------- tables ---------- */
  var COLS = {
    stores: [
      { k: 'status', h: 'Risk', f: function (r) { return badge(r.status); }, sortVal: function (r) { return { High: 3, Medium: 2, Low: 1, Clean: 0 }[r.status]; } },
      { k: 'store', h: 'Branch', cls: 'strong' },
      { k: 'invoices', h: 'Invoices', num: 1, f: function (r) { return n0(r.invoices); } },
      { k: 'mugs', h: 'Mugs', num: 1, f: function (r) { return n0(r.mugs); } },
      { k: 'directLines', h: 'Direct', num: 1 },
      { k: 'directExcess', h: 'Over (direct)', num: 1 },
      { k: 'runs', h: 'Split runs', num: 1 },
      { k: 'runsHigh', h: 'High runs', num: 1 },
      { k: 'runExcess', h: 'Over (split)', num: 1 },
      { k: 'pctOver', h: '% over limit', num: 1, f: function (r) {
        return pct(r.pctOver) + '<span class="meter' + (r.pctOver >= 0.25 ? ' warn' : '') + '"><i style="width:' + Math.round(Math.min(1, r.pctOver) * 100) + '%"></i></span>'; } },
      { k: 'zero', h: 'Zero-qty', num: 1 }
    ],
    runs: [
      { k: 'level', h: 'Risk', f: function (r) { return badge(r.level); }, sortVal: function (r) { return r.score; } },
      { k: 'score', h: 'Score', num: 1, f: function (r) { return r.score + '<span class="meter' + (r.score >= 60 ? ' warn' : '') + '"><i style="width:' + r.score + '%"></i></span>'; } },
      { k: 'store', h: 'Branch', cls: 'strong' },
      { k: 'label', h: 'Item / colour' },
      { k: 'start', h: 'From', f: function (r) { return t(r.start); } },
      { k: 'end', h: 'To', f: function (r) { return t(r.end); } },
      { k: 'durationMin', h: 'Duration', num: 1, f: function (r) { return mins(r.durationMin); } },
      { k: 'invoices', h: 'Invoices', num: 1 },
      { k: 'mugs', h: 'Mugs', num: 1 },
      { k: 'excess', h: 'Over', num: 1 },
      { k: 'burst', h: '× pace', num: 1, f: function (r) { return r.burst >= 100 ? '100+' : r.burst.toFixed(1); } },
      { k: 'identical', h: 'Same basket', num: 1, f: function (r) { return r.invoices >= 3 ? pct(r.identical) : '—'; } }
    ],
    direct: [
      { k: 'store', h: 'Branch', cls: 'strong' },
      { k: 'label', h: 'Item / colour' },
      { k: 'time', h: 'Time', f: function (r) { return t(r.time); } },
      { k: 'qty', h: 'Qty', num: 1 },
      { k: 'excess', h: 'Over', num: 1 },
      { k: 'sales', h: 'SAR', num: 1, f: function (r) { return sar(r.sales); } },
      { k: 'basket', h: 'Full invoice', cls: 'wrap-cell' }
    ],
    zero: [
      { k: 'store', h: 'Branch', cls: 'strong' },
      { k: 'item', h: 'Item' },
      { k: 'time', h: 'Time', f: function (r) { return t(r.time); } },
      { k: 'qty', h: 'Qty', num: 1 },
      { k: 'sales', h: 'SAR', num: 1, f: function (r) { return sar(r.sales); } }
    ]
  };
  var DEFAULT_SORT = { stores: null, runs: null, direct: { k: 'qty', dir: -1 }, zero: { k: 'time', dir: 1 } };

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
    $('rowCount').textContent = plural(rows.length, 'row', 'rows') +
      (tab === 'stores' ? ' · click a branch to open it' : tab === 'runs' ? ' · click a run for invoice detail' : '');

    var html = '<table><thead><tr>' + cols.map(function (c) {
      var arrow = s && s.k === c.k ? (s.dir > 0 ? ' ↑' : ' ↓') : '';
      return '<th data-k="' + c.k + '" class="' + (c.num ? 'num' : '') + '" aria-sort="' +
        (s && s.k === c.k ? (s.dir > 0 ? 'ascending' : 'descending') : 'none') + '">' + esc(c.h) + '<span class="arrow">' + arrow + '</span></th>';
    }).join('') + '</tr></thead><tbody>';

    var MAX = 2000;
    rows.slice(0, MAX).forEach(function (r) {
      var id = tab === 'runs' ? r.store + '|' + r.label + '|' + r.start : tab === 'stores' ? r.store : '';
      var clickable = tab === 'runs' || tab === 'stores';
      var open = tab === 'runs' && state.open[id];
      html += '<tr' + (clickable ? ' class="clickable' + (open ? ' open' : '') + '" data-id="' + esc(id) + '"' : '') + '>' + cols.map(function (c) {
        var v = c.f ? c.f(r) : esc(r[c.k]);
        return '<td class="' + (c.num ? 'num ' : '') + (c.cls || '') + '">' + v + '</td>';
      }).join('') + '</tr>';
      if (open) html += '<tr class="detail"><td colspan="' + cols.length + '"><div class="detail-inner">' + runDetail(r) + '</div></td></tr>';
    });
    if (!rows.length) html += '<tr><td colspan="' + cols.length + '" class="empty-row">Nothing matches these filters.</td></tr>';
    html += '</tbody></table>';
    if (rows.length > MAX) html += '<p class="muted small" style="padding:10px 14px">Showing the first ' + MAX + ' rows. Narrow the filters or export to Excel for all.</p>';
    $('tableWrap').innerHTML = html;
  }

  function runDetail(r) {
    var o = state.result.options;
    var span = Math.max(1, r.end - r.start);
    var bars = r.detail.map(function (d) {
      return '<span title="' + esc(t(d.time) + ' · ' + d.qty) + '" style="left:calc(' + ((d.time - r.start) / span * 100).toFixed(2) +
        '% - 2px);height:' + Math.max(25, d.qty / o.limit * 100) + '%"></span>';
    }).join('');
    var why = [
      [r.points.invoices, r.invoices + ' invoices in ' + mins(r.durationMin)],
      [r.points.excess, r.excess + ' mugs over the limit of ' + o.limit],
      [Math.round(r.points.burst), (r.burst >= 100 ? 'Over 100' : r.burst.toFixed(1)) + '× this branch’s normal pace (normally ~' + r.expected.toFixed(2) + ' invoices of this item in a window this long)']
    ];
    if (r.invoices >= 3) why.push([r.points.identical, pct(r.identical) + ' of invoices have the identical basket “' + r.topBasket + '”']);
    return '<div class="detail-grid"><span>Score <b>' + r.score + '</b> / 100</span><span>Value <b>SAR ' + sar(r.sales) + '</b></span>' +
      '<span>Items <b>' + esc(r.items) + '</b></span></div>' +
      '<ul class="why">' + why.map(function (w) { return '<li><span class="pts">+' + w[0] + '</span><span>' + esc(w[1]) + '</span></li>'; }).join('') + '</ul>' +
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
    if (state.tab === 'stores') openStore(id);
    else { state.open[id] = !state.open[id]; renderTable(); }
  });

  /* ---------- tabs with a sliding indicator ---------- */
  function moveInk() {
    var active = document.querySelector('.tabs button.active'), ink = $('tabInk');
    if (!active) return;
    ink.style.transform = 'translateX(' + active.offsetLeft + 'px) scaleX(' + (active.offsetWidth / 100) + ')';
  }
  function switchTab(tab) {
    state.tab = tab;
    Array.prototype.forEach.call(document.querySelectorAll('.tabs button'), function (b) {
      var on = b.getAttribute('data-tab') === tab;
      b.classList.toggle('active', on);
      b.setAttribute('aria-selected', on);
    });
    moveInk();
    renderTable();
  }
  Array.prototype.forEach.call(document.querySelectorAll('.tabs button'), function (b) {
    b.addEventListener('click', function () { switchTab(b.getAttribute('data-tab')); });
  });
  window.addEventListener('resize', moveInk);
  if (document.fonts && document.fonts.ready) document.fonts.ready.then(moveInk);
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
  $('help').addEventListener('click', function (e) { if (e.target === this) this.close(); });
})();
