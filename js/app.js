/* Mugs SOP Monitor – browser UI */
(function () {
  'use strict';
  var E = window.MugsEngine;
  var $ = function (id) { return document.getElementById(id); };
  var reduceMotion = window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches;

  var state = {
    reviews: {}, cloudReport: null, mapping: null, scope: { am: '', rom: '', city: '', zone: '' },
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
  function mins(m) { return (m < 1 ? m.toFixed(2) : m.toFixed(1)) + ' min'; }
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

  /* ---------- remember the last file in this browser (IndexedDB) ---------- */
  // Stays on this computer only. Restored on refresh; "Clear" removes it.
  var store = (function () {
    var DB = 'mugs-sop', OS = 'files', KEY = 'last';
    function open() {
      return new Promise(function (res, rej) {
        if (!('indexedDB' in window)) return rej(new Error('no indexedDB'));
        var r = indexedDB.open(DB, 1);
        r.onupgradeneeded = function () { r.result.createObjectStore(OS); };
        r.onsuccess = function () { res(r.result); };
        r.onerror = function () { rej(r.error); };
      });
    }
    function tx(mode, fn) {
      return open().then(function (db) {
        return new Promise(function (res, rej) {
          var t = db.transaction(OS, mode), req = fn(t.objectStore(OS));
          t.oncomplete = function () { db.close(); res(req && req.result); };
          t.onerror = t.onabort = function () { db.close(); rej(t.error); };
        });
      });
    }
    return {
      save: function (rec) { return tx('readwrite', function (s) { return s.put(rec, KEY); }).catch(function () {}); },
      get: function () { return tx('readonly', function (s) { return s.get(KEY); }).catch(function () { return null; }); },
      clear: function () { return tx('readwrite', function (s) { return s.delete(KEY); }).catch(function () {}); },
      saveKey: function (k, rec) { return tx('readwrite', function (s) { return s.put(rec, k); }).catch(function () {}); },
      getKey: function (k) { return tx('readonly', function (s) { return s.get(k); }).catch(function () { return null; }); },
      clearKey: function (k) { return tx('readwrite', function (s) { return s.delete(k); }).catch(function () {}); }
    };
  })();

  function parseBuffer(buf) {
    var wb = XLSX.read(new Uint8Array(buf), { type: 'array', raw: false, cellDates: false });
    var ws = wb.Sheets[wb.SheetNames[0]];
    var rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
    var parsed = E.parseRows(rows);
    if (!parsed.lines.length) throw new Error('No sale lines found in this file.');
    return parsed;
  }
  // Show a parsed report. meta: { savedAt } for browser-saved, { cloud } for cloud reports.
  function useParsed(parsed, name, meta) {
    meta = meta || {};
    state.parsed = parsed; state.fileName = name; state.savedAt = meta.savedAt || null;
    state.cloudReport = meta.cloud || null; state.reviews = meta.reviews || {};
    state.open = {};
    state.filter = { store: '', level: '', q: '' };
    $('fStore').value = ''; $('fLevel').value = ''; $('fSearch').value = '';
    run(true);
  }
  function process(buf, name, savedAt) { useParsed(parseBuffer(buf), name, { savedAt: savedAt }); }

  function load(file) {
    showError('');
    var reader = new FileReader();
    reader.onload = function (e) {
      var buf = e.target.result, now = Date.now(), parsed;
      try {
        parsed = parseBuffer(buf);
      } catch (err) {
        console.error(err);
        showError('Could not read “' + file.name + '”: ' + err.message);
        return;
      }
      if (cloud.on) {
        useParsed(parsed, file.name, {});
        cloud.save(buf, file, parsed);
      } else {
        useParsed(parsed, file.name, { savedAt: now });
        store.save({ name: file.name, bytes: buf, savedAt: now });
      }
    };
    reader.readAsArrayBuffer(file);
  }

  $('btnClear').addEventListener('click', function (e) {
    e.stopPropagation();
    if (!window.confirm('Remove the saved report from this browser?')) return;
    store.clear().then(function () { location.reload(); });
  });

  /* ---------- cloud mode (Cloudflare Worker API) ---------- */
  // Active only when /api/me answers (the Cloudflare deployment). On GitHub Pages
  // the app stays local-only and keeps the last file in this browser.
  var REVIEW_LABELS = { open: 'Open', genuine: 'Genuine (family/group)', violation: 'Violation', false_positive: 'False positive' };
  function runKey(r) { return 'run|' + r.store + '|' + r.label + '|' + r.start; }
  function directKey(d) { return 'direct|' + d.store + '|' + d.label + '|' + d.time; }
  function isoDate(parsed, T) {
    var m = parsed.reportDate && String(parsed.reportDate).match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (m) return m[3] + '-' + ('0' + m[1]).slice(-2) + '-' + ('0' + m[2]).slice(-2);
    return new Date(T.start).toISOString().slice(0, 10);
  }
  function toast(msg, kind) {
    var el = $('toast');
    el.textContent = msg; el.className = 'toast show' + (kind ? ' ' + kind : '');
    clearTimeout(el._t);
    if (kind !== 'busy') el._t = setTimeout(function () { el.className = 'toast'; }, 4200);
  }
  function gzip(text) {
    return new Response(new Blob([text]).stream().pipeThrough(new CompressionStream('gzip'))).arrayBuffer();
  }
  function api(path, opts) {
    opts = opts || {};
    var init = { method: opts.method || 'GET', headers: {}, credentials: 'same-origin' };
    if (opts.json !== undefined) { init.body = JSON.stringify(opts.json); init.headers['content-type'] = 'application/json'; }
    if (opts.body !== undefined) { init.body = opts.body; if (opts.type) init.headers['content-type'] = opts.type; }
    return fetch(path, init).then(function (res) {
      if (opts.raw && res.ok) return res;
      return res.text().then(function (txt) {
        var data = null; try { data = txt ? JSON.parse(txt) : null; } catch (e) { /* not json */ }
        if (!res.ok) throw new Error((data && data.error) || ('HTTP ' + res.status));
        return data;
      });
    });
  }

  var cloud = {
    on: false, me: null, reports: [],
    init: function () {
      var ctrl = 'AbortController' in window ? new AbortController() : null;
      var timer = ctrl ? setTimeout(function () { ctrl.abort(); }, 5000) : null;
      return fetch('/api/me', { credentials: 'same-origin', signal: ctrl ? ctrl.signal : undefined })
        .then(function (res) {
          clearTimeout(timer);
          var type = res.headers.get('content-type') || '';
          if (type.indexOf('application/json') < 0) return false;
          return res.json().then(function (me) {
            if (res.status === 401 && me && me.auth === 'app') { auth.show(); return 'login'; }
            if (!res.ok || !me || !me.email) return false;
            cloud.on = true; cloud.me = me;
            cloud.admin = me.role === 'admin';
            document.body.classList.add('cloud', cloud.admin ? 'admin' : 'viewer');
            $('userChip').hidden = false;
            $('userEmail').textContent = me.name || me.email;
            $('userRole').textContent = cloud.admin ? 'Admin' : 'Area manager';
            $('retentionDays').textContent = me.retentionDays;
            if (cloud.admin) users.badge();
            else {
              $('viewerName').textContent = me.areaManager || '—';
              if (!me.areaManager) toast('No area manager is linked to your account yet. Ask an admin.', 'error');
            }
            return true;
          });
        }).catch(function () { clearTimeout(timer); return false; });
    },
    refresh: function () {
      return api('/api/reports').then(function (d) { cloud.reports = d.reports || []; renderHistory(); return cloud.reports; });
    },
    open: function (id) {
      toast('Opening report…', 'busy');
      var meta;
      return api('/api/reports/' + id).then(function (m) {
        meta = m;
        return Promise.all([
          api('/api/reports/' + id + '/lines', { raw: true }).then(function (res) {
            return new Response(res.body.pipeThrough(new DecompressionStream('gzip'))).json();
          }),
          api('/api/reports/' + id + '/reviews')
        ]);
      }).then(function (out) {
        var data = out[0], reviews = {};
        (out[1].reviews || []).forEach(function (r) { reviews[r.flagKey] = r; });
        if (!cloud.admin && meta.rules) applyRules(meta.rules);
        useParsed({ lines: data.lines, warnings: [], reportDate: data.reportDate }, meta.fileName, { cloud: meta, reviews: reviews });
        renderHistory();
        toast('Opened ' + longDate(meta.periodStart || Date.parse(meta.reportDate)) + ' from HM cloud');
      }).catch(function (err) { console.error(err); toast('Could not open report: ' + err.message, 'error'); });
    },
    save: function (buf, file, parsed) {
      var R = E.analyse(parsed.lines, readOptions()), T = R.totals, id;
      toast('Saving to HM cloud…', 'busy');
      var info = function (store, k) { var m = branchInfo(store); return m ? (m[k] || '') : ''; };
      if (!state.mapping) toast('Saving without a branch mapping: area managers will not see this day until you upload the mapping and re-upload the file.', 'error');
      var body = {
        reportDate: isoDate(parsed, T), fileName: file.name, fileSize: buf.byteLength,
        periodStart: T.start, periodEnd: T.end, rules: R.options,
        totals: {
          mugs: T.mugs, sales: Math.round(T.sales * 100) / 100, invoices: T.invoices, stores: T.stores,
          directLines: T.directLines, directExcess: T.directExcess, runs: T.runs, runsHigh: T.runsHigh,
          runsMedium: T.runsMedium, zero: T.zero, highStores: R.stores.filter(function (s) { return s.status === 'High'; }).length
        },
        stores: R.stores.map(function (s) {
          return { store: s.store, status: s.status, invoices: s.invoices, mugs: s.mugs, sales: Math.round(s.sales * 100) / 100,
            directLines: s.directLines, directExcess: s.directExcess, runs: s.runs, runsHigh: s.runsHigh,
            runExcess: s.runExcess, pctOver: Math.round(s.pctOver * 10000) / 10000,
            am: info(s.store, 'am'), rom: info(s.store, 'rom'), city: info(s.store, 'city'), zone: info(s.store, 'zone') };
        })
      };
      var byAm = {};
      var lines = parsed.lines.map(function (l) {
        return { row: l.row, store: l.store, code: l.code, item: l.item, itemEn: l.itemEn, colour: l.colour, time: l.time,
          sales: l.sales, qty: l.qty, invoice: l.invoice, cashier: l.cashier, customer: l.customer };
      });
      lines.forEach(function (l) { var am = info(l.store, 'am'); if (am) (byAm[am] = byAm[am] || []).push(l); });
      return api('/api/reports', { method: 'POST', json: body }).then(function (d) {
        id = d.id;
        return Promise.all([
          api('/api/reports/' + id + '/file', { method: 'PUT', body: buf, type: file.type || 'application/octet-stream' }),
          gzip(JSON.stringify({ v: 1, reportDate: parsed.reportDate, lines: lines })).then(function (gz) {
            return api('/api/reports/' + id + '/lines', { method: 'PUT', body: gz, type: 'application/gzip' });
          })
        ].concat(Object.keys(byAm).map(function (am) {
          // Each area manager gets a file with only their branches; the server serves it by login.
          return gzip(JSON.stringify({ v: 1, reportDate: parsed.reportDate, lines: byAm[am] })).then(function (gz) {
            return api('/api/reports/' + id + '/lines?am=' + encodeURIComponent(am), { method: 'PUT', body: gz, type: 'application/gzip' });
          });
        })));
      }).then(function () {
        return api('/api/reports/' + id + '/complete', { method: 'POST' });
      }).then(function (res) {
        return api('/api/reports/' + id).then(function (meta) {
          state.cloudReport = meta; state.reviews = {};
          run(false);
          toast(res.replaced ? 'Saved to HM cloud · replaced the earlier upload for this day' : 'Saved to HM cloud · your team can see it now');
          cloud.refresh(); renderTrends();
        });
      }).catch(function (err) {
        console.error(err);
        toast('Not saved to the cloud: ' + err.message + '. The report is still shown here.', 'error');
      });
    },
    remove: function (id) {
      return api('/api/reports/' + id, { method: 'DELETE' }).then(function () {
        toast('Report deleted');
        if (state.cloudReport && state.cloudReport.id === id) { location.reload(); return; }
        cloud.refresh(); renderTrends();
      }).catch(function (err) { toast('Could not delete: ' + err.message, 'error'); });
    },
    saveReview: function (key, kind, store, status, note) {
      if (!state.cloudReport) { toast('Save the report to the cloud before reviewing', 'error'); return Promise.resolve(); }
      return api('/api/reviews', { method: 'PUT', json: { flagKey: key, reportId: state.cloudReport.id, kind: kind, store: store, status: status, note: note } })
        .then(function (r) {
          state.reviews[key] = { flagKey: key, kind: kind, store: store, status: status, note: note, reviewedBy: r.reviewedBy, reviewedAt: r.reviewedAt };
          toast('Review saved: ' + REVIEW_LABELS[status]);
          renderTable();
        }).catch(function (err) { toast('Review not saved: ' + err.message, 'error'); });
    }
  };

  function reviewBadge(key) {
    var r = state.reviews && state.reviews[key];
    var st = r ? r.status : 'open';
    return '<span class="rv rv-' + st + '">' + esc(r ? REVIEW_LABELS[st] : 'Not reviewed') + '</span>';
  }
  function reviewBar(key, kind, store) {
    if (!cloud.on) return '';
    var r = (state.reviews && state.reviews[key]) || {};
    if (!cloud.admin) {
      return '<div class="review-bar"><span class="rv-title">Review</span>' + reviewBadge(key) +
        (r.note ? '<span class="small">' + esc(r.note) + '</span>' : '') +
        (r.reviewedBy ? '<span class="rv-meta">By ' + esc(r.reviewedBy) + ' · ' +
          new Date(r.reviewedAt).toLocaleString('en-GB', { day: 'numeric', month: 'short', hour: '2-digit', minute: '2-digit' }) + '</span>' : '') + '</div>';
    }
    var opts = Object.keys(REVIEW_LABELS).map(function (k) {
      return '<option value="' + k + '"' + ((r.status || 'open') === k ? ' selected' : '') + '>' + REVIEW_LABELS[k] + '</option>';
    }).join('');
    return '<div class="review-bar" data-key="' + esc(key) + '" data-kind="' + kind + '" data-store="' + esc(store) + '">' +
      '<span class="rv-title">Review</span>' +
      '<select class="rv-status" aria-label="Review status">' + opts + '</select>' +
      '<input class="rv-note" type="text" maxlength="500" placeholder="Note (e.g. family of 5, manager approved)" value="' + esc(r.note || '') + '" aria-label="Review note">' +
      '<button type="button" class="btn btn-primary btn-sm rv-save">Save</button>' +
      (r.reviewedBy ? '<span class="rv-meta">Last saved by ' + esc(r.reviewedBy) + ' · ' +
        new Date(r.reviewedAt).toLocaleString('en-GB', { day: 'numeric', month: 'short', hour: '2-digit', minute: '2-digit' }) + '</span>' : '') +
      '</div>';
  }

  function renderHistory() {
    var el = $('historyBody');
    if (!el) return;
    var cur = state.cloudReport && state.cloudReport.id;
    if (!cloud.reports.length) { el.innerHTML = '<tr><td colspan="8" class="empty-row">No reports saved yet. Upload a day to start the history.</td></tr>'; return; }
    el.innerHTML = cloud.reports.map(function (r) {
      var T = r.totals || {};
      return '<tr' + (r.id === cur ? ' class="current"' : '') + '>' +
        '<td class="strong">' + esc(longDate(Date.parse(r.reportDate + 'T00:00:00Z'))) + (r.id === cur ? ' <span class="pill">open</span>' : '') + '</td>' +
        '<td>' + esc(r.fileName) + '</td>' +
        '<td>' + esc(r.uploadedBy || '') + '<br><span class="muted small">' + new Date(r.uploadedAt).toLocaleString('en-GB', { day: 'numeric', month: 'short', hour: '2-digit', minute: '2-digit' }) + '</span></td>' +
        '<td class="num">' + n0(T.highStores || 0) + '</td><td class="num">' + n0(T.directLines || 0) + '</td>' +
        '<td class="num">' + n0(T.runs || 0) + '</td><td class="num">' + n0(T.mugs || 0) + '</td>' +
        '<td class="actions">' +
          (r.id === cur ? '' : '<button type="button" class="btn btn-ghost btn-sm" data-act="open" data-id="' + r.id + '">Open</button>') +
          (cloud.admin ? '<a class="btn btn-ghost btn-sm" href="/api/reports/' + r.id + '/file">Original</a>' +
          '<button type="button" class="btn btn-ghost btn-sm danger" data-act="delete" data-id="' + r.id + '">Delete</button>' : '') +
        '</td></tr>';
    }).join('');
  }
  $('historyBody').addEventListener('click', function (e) {
    var b = e.target.closest('button[data-act]');
    if (!b) return;
    var id = b.getAttribute('data-id');
    if (b.getAttribute('data-act') === 'open') { cloud.open(id); window.scrollTo({ top: 0, behavior: reduceMotion ? 'auto' : 'smooth' }); }
    else if (window.confirm('Delete this day for everyone? The uploaded file and its reviews are removed.')) cloud.remove(id);
  });

  function renderTrends() {
    if (!cloud.on) return;
    var days = parseInt($('trendDays').value, 10) || 30;
    api('/api/trends?days=' + days).then(function (d) {
      state.trend = d;
      renderDaily();
      var dates = d.dates || [], by = {};
      var rowsIn = d.rows.filter(trendInScope);
      rowsIn.forEach(function (r) {
        var s = by[r.store] = by[r.store] || { store: r.store, days: {}, high: 0, med: 0, excess: 0, last: '' };
        s.days[r.report_date] = r.status;
        if (r.status === 'High') { s.high++; s.last = r.report_date > s.last ? r.report_date : s.last; }
        if (r.status === 'Medium') s.med++;
        s.excess += (r.direct_excess || 0) + (r.run_excess || 0);
      });
      var list = Object.keys(by).map(function (k) { return by[k]; }).filter(function (s) { return s.high || s.med; })
        .sort(function (a, b) { return b.high - a.high || b.med - a.med || b.excess - a.excess; }).slice(0, 25);
      $('trendNote').textContent = dates.length
        ? plural(dates.length, 'day', 'days') + ' of history · branches with any High or Medium day, worst first'
        : 'No history yet.';
      if (!list.length) { $('trendWrap').innerHTML = '<p class="muted small">Nothing flagged in this period.</p>'; return; }
      var short = function (iso) { var p = iso.split('-'); return +p[2] + '/' + +p[1]; };
      var html = '<table class="heat"><thead><tr><th>Branch</th>' +
        dates.map(function (dd) { return '<th class="day" title="' + dd + '"><span>' + short(dd) + '</span></th>'; }).join('') +
        '<th class="num">High days</th><th class="num">Medium days</th><th class="num">Mugs over</th></tr></thead><tbody>' +
        list.map(function (s) {
          return '<tr><td class="strong">' + esc(s.store) + '</td>' + dates.map(function (dd) {
            var st = s.days[dd] || 'none';
            return '<td class="cell"><i class="hc hc-' + st + '" title="' + esc(s.store + ' · ' + dd + ' · ' + (st === 'none' ? 'no data' : st)) + '"></i></td>';
          }).join('') + '<td class="num strong">' + s.high + '</td><td class="num">' + s.med + '</td><td class="num">' + n0(s.excess) + '</td></tr>';
        }).join('') + '</tbody></table>';
      $('trendWrap').innerHTML = html;
    }).catch(function (err) { $('trendWrap').innerHTML = '<p class="muted small">Could not load trends: ' + esc(err.message) + '</p>'; });
  }
  $('trendDays').addEventListener('change', renderTrends);
  $('btnOriginal').addEventListener('click', function (e) {
    e.stopPropagation();
    if (state.cloudReport) window.location.href = '/api/reports/' + state.cloudReport.id + '/file';
  });

  // Start: cloud if available (open the latest day), otherwise restore from this browser.
  cloud.init().then(function (on) {
    if (on === 'login') return;
    var mapReady = on
      ? api('/api/mapping').then(function (m) { if (m && m.branches) { state.mapping = m; renderMapStatus(); } }).catch(function () {})
      : store.getKey('mapping').then(function (m) { if (m && m.branches) { state.mapping = m; renderMapStatus(); if (state.parsed) run(false); } });
    if (on) {
      mapReady.then(function () { if (state.parsed) run(false); });
      cloud.refresh().then(function (list) {
        renderTrends();
        if (list.length && !state.parsed) cloud.open(list[0].id);
        else if (!list.length && !cloud.admin) { $('empty').hidden = true; $('viewerEmpty').hidden = false; }
      }).catch(function (err) { toast('Could not load reports: ' + err.message, 'error'); });
      return;
    }
    store.get().then(function (rec) {
      if (!rec || !rec.bytes || state.parsed) return;
      try { process(rec.bytes, rec.name, rec.savedAt); }
      catch (err) { console.error(err); store.clear(); }
    });
  });

  /* ---------- branch mapping (area manager, ROM, city, zone) ---------- */
  var MAP = window.MugsMapping;
  var SCOPE_FIELDS = [['am', 'Area manager'], ['rom', 'ROM'], ['city', 'City'], ['zone', 'Zone']];
  function branchInfo(store) { return state.mapping ? MAP.lookup(state.mapping, store) : null; }
  function scopeActive() { return !!(state.mapping && (state.scope.am || state.scope.rom || state.scope.city || state.scope.zone)); }
  function inScope(store) {
    var m = branchInfo(store);
    if (!m) return false;
    return SCOPE_FIELDS.every(function (f) { return !state.scope[f[0]] || m[f[0]] === state.scope[f[0]]; });
  }
  function scopeLabel() {
    if (!scopeActive()) return '';
    return SCOPE_FIELDS.filter(function (f) { return state.scope[f[0]]; }).map(function (f) { return f[1] + ': ' + state.scope[f[0]]; }).join(' · ');
  }
  function renderScope() {
    var bar = $('scopeBar');
    bar.hidden = !state.mapping;
    if (!state.mapping || !state.parsed) return;
    // options come from the branches present in this file, narrowed by the other active choices
    var stores = {};
    state.parsed.lines.forEach(function (l) { stores[l.store] = 1; });
    var infos = Object.keys(stores).map(branchInfo).filter(Boolean);
    SCOPE_FIELDS.forEach(function (f) {
      var key = f[0], sel = $('scope_' + key);
      var vals = {};
      infos.forEach(function (m) {
        var ok = SCOPE_FIELDS.every(function (g) { return g[0] === key || !state.scope[g[0]] || m[g[0]] === state.scope[g[0]]; });
        if (ok && m[key]) vals[m[key]] = (vals[m[key]] || 0) + 1;
      });
      sel.innerHTML = '<option value="">All ' + (key === 'city' ? 'cities' : key === 'zone' ? 'zones' : f[1] + 's') + '</option>' +
        Object.keys(vals).sort().map(function (v) {
          return '<option value="' + esc(v) + '"' + (state.scope[key] === v ? ' selected' : '') + '>' + esc(v) + ' (' + vals[v] + ')</option>';
        }).join('');
    });
    var total = Object.keys(stores).length, mapped = infos.length, shown = state.result ? state.result.stores.length : total;
    $('scopeNote').textContent = scopeActive()
      ? 'Showing ' + plural(shown, 'branch', 'branches') + ' of ' + total
      : mapped + ' of ' + total + ' branches matched to the mapping' + (mapped < total ? ' · ' + (total - mapped) + ' not in mapping' : '');
    $('scopeReset').hidden = !scopeActive();
  }
  SCOPE_FIELDS.forEach(function (f) {
    $('scope_' + f[0]).addEventListener('change', function () {
      state.scope[f[0]] = this.value;
      state.filter.store = ''; state.open = {};
      if (state.parsed) run(false);
    });
  });
  $('scopeReset').addEventListener('click', function () {
    state.scope = { am: '', rom: '', city: '', zone: '' };
    if (state.parsed) run(false);
  });

  function renderMapStatus() {
    var m = state.mapping;
    $('mapStatus').innerHTML = m
      ? '<b>' + n0(m.count) + ' branches</b> from ' + esc(m.fileName || 'mapping file') +
        (m.manual && Object.keys(m.manual).length ? ' + ' + Object.keys(m.manual).length + ' added here' : '') +
        (m.savedAt ? '<br><span class="muted small">Loaded ' + new Date(m.savedAt).toLocaleString('en-GB', { day: 'numeric', month: 'short', hour: '2-digit', minute: '2-digit' }) +
          (m.savedBy ? ' by ' + esc(m.savedBy) : '') + (cloud.on ? ' · shared with your team' : ' · kept in this browser') + '</span>' : '')
      : 'No mapping loaded. Upload the Master Location Mapping file to sort and filter by area manager, ROM, city and zone.';
    $('btnMapClear').hidden = !m;
    $('btnMap').textContent = m ? 'Replace mapping' : 'Upload mapping';
    renderMissing();
  }
  // Branches in the sales file that the mapping file does not cover: pick the area manager here.
  function renderMissing() {
    var box = $('mapMissing'), m = state.mapping;
    if (!m || !state.parsed) { box.innerHTML = ''; return; }
    var names = {};
    state.parsed.lines.forEach(function (l) { names[l.store] = 1; });
    var missing = Object.keys(names).filter(function (s) { return MAP.codeOf(s) && !m.branches[MAP.codeOf(s)]; }).sort();
    if (!missing.length) { box.innerHTML = ''; return; }
    var ams = {};
    Object.keys(m.branches).forEach(function (k) { if (m.branches[k].am) ams[m.branches[k].am] = 1; });
    var html = '<p class="small strong-label">Branches not in the mapping file</p><datalist id="amList">' +
      Object.keys(ams).sort().map(function (a) { return '<option value="' + esc(a) + '">'; }).join('') + '</datalist>';
    var todo = missing.filter(function (s) { return !(m.manual && m.manual[MAP.codeOf(s)]); });
    var done = missing.filter(function (s) { return m.manual && m.manual[MAP.codeOf(s)]; });
    if (!todo.length) html = '';
    var rowHtml = function (s) {
      var code = MAP.codeOf(s), man = m.manual && m.manual[code];
      if (man) {
        return '<div class="miss-row done"><span><b>' + esc(s) + '</b><br><span class="muted small">' + esc(man.am) + ' · ' + esc(man.rom || '—') +
          ' · ' + esc(man.city || '—') + '</span></span><button type="button" class="btn btn-ghost btn-sm" data-unmap="' + esc(code) + '" aria-label="Remove">Edit</button></div>';
      }
      return '<div class="miss-row"><span class="small"><b>' + esc(s) + '</b></span>' +
        '<span class="miss-input"><input list="amList" placeholder="Area manager" data-store="' + esc(s) + '" aria-label="Area manager for ' + esc(s) + '">' +
        '<button type="button" class="btn btn-primary btn-sm" data-map="' + esc(s) + '">Save</button></span></div>';
    };
    html += todo.map(rowHtml).join('');
    if (done.length) {
      html += '<details class="miss-done"' + (todo.length ? '' : '') + '><summary class="small">Added here (' + done.length + ') · click to view or edit</summary>' +
        done.map(rowHtml).join('') + '</details>';
    }
    if (!todo.length) html = '<datalist id="amList">' + Object.keys(ams).sort().map(function (a) { return '<option value="' + esc(a) + '">'; }).join('') + '</datalist>' + html;
    box.innerHTML = html;
  }
  $('mapMissing').addEventListener('click', function (e) {
    var save = e.target.closest('button[data-map]'), un = e.target.closest('button[data-unmap]');
    if (!state.mapping) return;
    var m = state.mapping;
    if (save) {
      var store = save.getAttribute('data-map');
      var am = save.parentNode.querySelector('input').value.trim();
      if (!am) { toast('Type or pick the area manager first', 'error'); return; }
      m.manual = m.manual || {};
      var b = MAP.inferBranch(m, store, am);
      m.manual[b.code] = b;
      setMapping(m, true);
      toast(store + ' → ' + b.am + (b.rom ? ' · ROM ' + b.rom : ''));
    } else if (un) {
      delete m.manual[un.getAttribute('data-unmap')];
      setMapping(m, true);
    }
  });
  function setMapping(m, persist) {
    state.mapping = m;
    renderMapStatus();
    if (persist) {
      if (cloud.on) api('/api/mapping', { method: 'PUT', json: m }).then(function (r) {
        m.savedBy = r.savedBy; renderMapStatus(); toast('Branch mapping saved for your team');
      }).catch(function (err) { toast('Mapping not saved to the cloud: ' + err.message, 'error'); });
      else store.saveKey('mapping', m);
    }
    if (state.parsed) { state.open = {}; run(false); }
  }
  $('btnMap').addEventListener('click', function () { $('mapFile').click(); });
  $('mapFile').addEventListener('change', function () {
    var f = this.files[0]; this.value = '';
    if (!f) return;
    var reader = new FileReader();
    reader.onload = function (e) {
      try {
        var wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
        var m = MAP.parseWorkbook(XLSX, wb);
        m.fileName = f.name; m.savedAt = Date.now();
        if (state.mapping && state.mapping.manual) {
          m.manual = {};
          Object.keys(state.mapping.manual).forEach(function (k) { if (!m.branches[k]) m.manual[k] = state.mapping.manual[k]; });
        }
        setMapping(m, true);
        if (!cloud.on) toast('Mapping loaded: ' + m.count + ' branches');
      } catch (err) { console.error(err); toast('Could not read the mapping: ' + err.message, 'error'); }
    };
    reader.readAsArrayBuffer(f);
  });
  $('btnMapClear').addEventListener('click', function () {
    if (!window.confirm(cloud.on ? 'Remove the branch mapping for everyone?' : 'Remove the branch mapping from this browser?')) return;
    state.scope = { am: '', rom: '', city: '', zone: '' };
    if (cloud.on) api('/api/mapping', { method: 'DELETE' }).catch(function () {});
    else store.clearKey('mapping');
    setMapping(null, false);
  });
  renderMapStatus();

  /* ---------- accounts: sign in / create account ---------- */
  var auth = {
    show: function (tab) {
      document.body.classList.add('auth-open');
      $('authScreen').hidden = false;
      auth.tab(tab || 'login');
    },
    tab: function (t) {
      ['login', 'signup'].forEach(function (k) {
        $(k + 'Form').hidden = k !== t;
        $('authTab_' + k).classList.toggle('active', k === t);
        $('authTab_' + k).setAttribute('aria-selected', k === t);
      });
      auth.msg('');
      var first = $(t + 'Form').querySelector('input');
      if (first) setTimeout(function () { first.focus(); }, 50);
    },
    msg: function (text, kind) { var el = $('authMsg'); el.textContent = text; el.className = 'auth-msg' + (kind ? ' ' + kind : ''); el.hidden = !text; }
  };
  ['login', 'signup'].forEach(function (k) { $('authTab_' + k).addEventListener('click', function () { auth.tab(k); }); });
  function formBusy(form, busy) { Array.prototype.forEach.call(form.querySelectorAll('button, input'), function (el) { el.disabled = busy; }); }
  $('loginForm').addEventListener('submit', function (e) {
    e.preventDefault();
    var f = this; formBusy(f, true); auth.msg('Signing in…');
    api('/api/auth/login', { method: 'POST', json: { email: f.email.value, password: f.password.value } })
      .then(function () { location.reload(); })
      .catch(function (err) { formBusy(f, false); auth.msg(err.message, 'error'); });
  });
  $('signupForm').addEventListener('submit', function (e) {
    e.preventDefault();
    var f = this;
    if (f.password.value !== f.password2.value) { auth.msg('The two passwords do not match', 'error'); return; }
    formBusy(f, true); auth.msg('Creating your account…');
    api('/api/auth/signup', { method: 'POST', json: { name: f.name.value, email: f.email.value, password: f.password.value } })
      .then(function (r) {
        if (r.status === 'active') { location.reload(); return; }
        formBusy(f, false); f.reset();
        auth.tab('login');
        auth.msg('Account created. An admin will approve it and link it to your area. You can sign in after that.', 'ok');
      })
      .catch(function (err) { formBusy(f, false); auth.msg(err.message, 'error'); });
  });

  /* ---------- my account: change password, sign out ---------- */
  $('userChip').addEventListener('click', function () {
    if (!cloud.me) return;
    $('acctName').textContent = cloud.me.name || '';
    $('acctEmail').textContent = cloud.me.email;
    $('acctRole').textContent = cloud.admin ? 'Admin' : 'Area manager · ' + (cloud.me.areaManager || 'not linked yet');
    $('pwForm').reset(); $('pwMsg').hidden = true;
    $('accountDlg').showModal();
  });
  $('pwForm').addEventListener('submit', function (e) {
    e.preventDefault();
    var f = this, msg = $('pwMsg');
    api('/api/me/password', { method: 'POST', json: { current: f.current.value, password: f.password.value } })
      .then(function () { f.reset(); msg.textContent = 'Password changed.'; msg.className = 'auth-msg ok'; msg.hidden = false; })
      .catch(function (err) { msg.textContent = err.message; msg.className = 'auth-msg error'; msg.hidden = false; });
  });
  $('btnSignOut').addEventListener('click', function () {
    api('/api/auth/logout', { method: 'POST' }).then(function () { location.reload(); }, function () { location.reload(); });
  });

  /* ---------- admin: manage users ---------- */
  var users = {
    list: [],
    load: function () {
      return api('/api/users').then(function (d) { users.list = d.users || []; users.render(); users.badge(true); });
    },
    badge: function (skipLoad) {
      if (!skipLoad) return users.load();
      var n = users.list.filter(function (u) { return u.status === 'pending'; }).length;
      $('usersBadge').textContent = n; $('usersBadge').hidden = !n;
    },
    amNames: function () {
      var set = {};
      if (state.mapping) [state.mapping.branches, state.mapping.manual].forEach(function (src) {
        Object.keys(src || {}).forEach(function (k) { if (src[k].am) set[src[k].am] = 1; });
      });
      users.list.forEach(function (u) { if (u.areaManager) set[u.areaManager] = 1; });
      return Object.keys(set).sort();
    },
    render: function () {
      var dl = '<datalist id="userAmList">' + users.amNames().map(function (a) { return '<option value="' + esc(a) + '">'; }).join('') + '</datalist>';
      var pending = users.list.filter(function (u) { return u.status === 'pending'; }).length;
      $('usersNote').textContent = plural(users.list.length, 'account', 'accounts') + (pending ? ' · ' + pending + ' waiting for approval' : '');
      $('usersBody').innerHTML = dl + users.list.map(function (u) {
        var me = cloud.me && u.id === cloud.me.id;
        return '<tr data-id="' + u.id + '" class="' + (u.status === 'pending' ? 'pending' : '') + '">' +
          '<td><input class="u-name" value="' + esc(u.name) + '" aria-label="Name"><div class="muted small">' + esc(u.email) + '</div></td>' +
          '<td><select class="u-role" aria-label="Role"' + (me ? ' disabled' : '') + '><option value="viewer"' + (u.role === 'viewer' ? ' selected' : '') + '>Area manager</option>' +
            '<option value="admin"' + (u.role === 'admin' ? ' selected' : '') + '>Admin</option></select></td>' +
          '<td><input class="u-am" list="userAmList" value="' + esc(u.areaManager) + '" placeholder="Pick area manager" aria-label="Area manager"></td>' +
          '<td><select class="u-status" aria-label="Status"' + (me ? ' disabled' : '') + '>' +
            ['pending', 'active', 'disabled'].map(function (s) { return '<option value="' + s + '"' + (u.status === s ? ' selected' : '') + '>' + s[0].toUpperCase() + s.slice(1) + '</option>'; }).join('') +
          '</select></td>' +
          '<td class="small muted">' + (u.lastLogin ? new Date(u.lastLogin).toLocaleString('en-GB', { day: 'numeric', month: 'short', hour: '2-digit', minute: '2-digit' }) : 'Never') + '</td>' +
          '<td class="actions">' +
            (u.status === 'pending' ? '<button type="button" class="btn btn-primary btn-sm" data-u="approve">Approve</button>' : '<button type="button" class="btn btn-primary btn-sm" data-u="save">Save</button>') +
            '<button type="button" class="btn btn-ghost btn-sm" data-u="reset">Reset password</button>' +
            (me ? '' : '<button type="button" class="btn btn-ghost btn-sm danger" data-u="delete">Delete</button>') +
          '</td></tr>';
      }).join('');
    },
    patch: function (id, body, okMsg) {
      return api('/api/users/' + id, { method: 'PATCH', json: body })
        .then(function () { toast(okMsg || 'Saved'); return users.load(); })
        .catch(function (err) { toast(err.message, 'error'); });
    }
  };
  $('btnUsers').addEventListener('click', function () { users.load(); $('usersDlg').showModal(); });
  $('usersBody').addEventListener('click', function (e) {
    var b = e.target.closest('button[data-u]');
    if (!b) return;
    var tr = b.closest('tr'), id = tr.getAttribute('data-id'), act = b.getAttribute('data-u');
    var body = { name: tr.querySelector('.u-name').value, role: tr.querySelector('.u-role').value,
      areaManager: tr.querySelector('.u-am').value, status: tr.querySelector('.u-status').value };
    if (act === 'approve') { body.status = 'active'; users.patch(id, body, 'Account approved'); }
    else if (act === 'save') users.patch(id, body, 'Account updated');
    else if (act === 'reset') {
      var pw = window.prompt('New temporary password for this person (at least 8 characters). Share it with them privately; they can change it after signing in.');
      if (pw) users.patch(id, { password: pw }, 'Password reset');
    } else if (act === 'delete' && window.confirm('Delete this account? They will no longer be able to sign in.')) {
      api('/api/users/' + id, { method: 'DELETE' }).then(function () { toast('Account deleted'); users.load(); }).catch(function (err) { toast(err.message, 'error'); });
    }
  });

  /* ---------- rules stored with a report (area managers see the admin's rules) ---------- */
  function applyRules(r) {
    if (r.limit) $('optLimit').value = r.limit;
    if (r.gapMinutes) $('optGap').value = r.gapMinutes;
    if (r.minInvoices) $('optMinInv').value = r.minInvoices;
    if (r.groupBy) $('optGroup').value = r.groupBy;
    $('hdrLimit').textContent = r.limit || 2;
  }

  /* ---------- admin: leaderboard by area manager / ROM / city / zone ---------- */
  var GROUPS = { am: 'Area manager', rom: 'ROM', city: 'City', zone: 'Zone' };
  function renderLeaderboard() {
    var panel = $('leaderPanel');
    var show = !!(state.mapping && state.fullResult && (!cloud.on || cloud.admin));
    panel.hidden = !show;
    if (!show) return;
    var key = $('groupBy').value, groups = {};
    state.fullResult.stores.forEach(function (s) {
      var m = branchInfo(s.store), name = m ? (m[key] || '—') : 'Not in mapping';
      var g = groups[name] = groups[name] || { name: name, branches: 0, high: 0, direct: 0, directExcess: 0, runs: 0, runsHigh: 0, runExcess: 0, mugs: 0 };
      g.branches++; if (s.status === 'High') g.high++;
      g.direct += s.directLines; g.directExcess += s.directExcess; g.runs += s.runs; g.runsHigh += s.runsHigh; g.runExcess += s.runExcess; g.mugs += s.mugs;
    });
    var rows = Object.keys(groups).map(function (k) { return groups[k]; })
      .sort(function (a, b) { return (b.directExcess + b.runExcess) - (a.directExcess + a.runExcess) || b.direct - a.direct; });
    var maxD = Math.max.apply(null, rows.map(function (r) { return r.direct; }).concat([1]));
    var maxR = Math.max.apply(null, rows.map(function (r) { return r.runs; }).concat([1]));
    $('leaderHead').textContent = GROUPS[key];
    $('leaderBody').innerHTML = rows.map(function (r) {
      var active = state.scope[key] === r.name;
      var clickable = r.name !== 'Not in mapping' && r.name !== '—';
      return '<tr class="' + (clickable ? 'clickable' : '') + (active ? ' current' : '') + '" data-g="' + esc(r.name) + '">' +
        '<td class="strong">' + esc(r.name) + (active ? ' <span class="pill">viewing</span>' : '') + '</td>' +
        '<td class="num">' + r.branches + '</td><td class="num">' + (r.high ? '<span class="badge High">' + r.high + '</span>' : '0') + '</td>' +
        '<td class="num bar-cell"><span class="mini"><i class="d" style="width:' + Math.round(r.direct / maxD * 100) + '%"></i></span><b>' + n0(r.direct) + '</b></td>' +
        '<td class="num">' + n0(r.directExcess) + '</td>' +
        '<td class="num bar-cell"><span class="mini"><i class="s" style="width:' + Math.round(r.runs / maxR * 100) + '%"></i></span><b>' + n0(r.runs) + '</b></td>' +
        '<td class="num">' + n0(r.runsHigh) + '</td><td class="num">' + n0(r.runExcess) + '</td>' +
        '<td class="num">' + pct(r.mugs ? (r.directExcess + r.runExcess) / r.mugs : 0) + '</td></tr>';
    }).join('');
  }
  $('groupBy').addEventListener('change', renderLeaderboard);
  $('leaderBody').addEventListener('click', function (e) {
    var tr = e.target.closest('tr.clickable');
    if (!tr) return;
    var key = $('groupBy').value, name = tr.getAttribute('data-g');
    state.scope = { am: '', rom: '', city: '', zone: '' };
    state.scope[key] = state.scope[key] === name ? '' : name;
    state.filter.store = ''; state.open = {};
    run(false);
    window.scrollTo({ top: 0, behavior: reduceMotion ? 'auto' : 'smooth' });
  });

  /* ---------- daily trend: direct violations and split runs per day ---------- */
  function trendInScope(r) {
    if (!scopeActive()) return true;
    return SCOPE_FIELDS.every(function (f) { return !state.scope[f[0]] || r[f[0]] === state.scope[f[0]]; });
  }
  function renderDaily() {
    var panel = $('dailyPanel'), d = state.trend;
    panel.hidden = !(cloud.on && d && d.dates && d.dates.length);
    if (panel.hidden) return;
    var per = {};
    d.dates.forEach(function (x) { per[x] = { direct: 0, runs: 0, dEx: 0, rEx: 0 }; });
    d.rows.filter(trendInScope).forEach(function (r) {
      var p = per[r.report_date]; if (!p) return;
      p.direct += r.direct_lines || 0; p.runs += r.runs || 0; p.dEx += r.direct_excess || 0; p.rEx += r.run_excess || 0;
    });
    var max = Math.max.apply(null, d.dates.map(function (x) { return Math.max(per[x].direct, per[x].runs); }).concat([1]));
    var tot = d.dates.reduce(function (a, x) { a.d += per[x].direct; a.r += per[x].runs; return a; }, { d: 0, r: 0 });
    $('dailyNote').textContent = plural(d.dates.length, 'day', 'days') + (scopeLabel() ? ' · ' + scopeLabel() : '') +
      ' · ' + n0(tot.d) + ' direct violations · ' + n0(tot.r) + ' split runs';
    var short = function (iso) { var p = iso.split('-'); return +p[2] + '/' + +p[1]; };
    $('dailyChart').innerHTML = d.dates.map(function (x) {
      var p = per[x];
      return '<div class="day-col" title="' + esc(longDate(Date.parse(x + 'T00:00:00Z')) + ': ' + p.direct + ' direct violations (' + p.dEx + ' mugs over), ' + p.runs + ' split runs (' + p.rEx + ' mugs over)') + '">' +
        '<div class="day-bars"><i class="d" style="height:' + (p.direct / max * 100).toFixed(1) + '%"></i><i class="s" style="height:' + (p.runs / max * 100).toFixed(1) + '%"></i></div>' +
        '<span class="day-lbl">' + short(x) + '</span></div>';
    }).join('');
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
    var lines = state.parsed.lines;
    if (scopeActive()) {
      var scoped = lines.filter(function (l) { return inScope(l.store); });
      if (scoped.length) lines = scoped;
      else { state.scope = { am: '', rom: '', city: '', zone: '' }; toast('No branches in this file match that view, showing all branches', 'error'); }
    }
    var res = E.analyse(lines, readOptions());
    state.fullResult = lines === state.parsed.lines ? res : E.analyse(state.parsed.lines, readOptions());
    res.stores.forEach(function (s) {
      var m = branchInfo(s.store);
      s.city = m ? m.city : ''; s.zone = m ? m.zone : ''; s.am = m ? m.am : ''; s.rom = m ? m.rom : ''; s.om = m ? m.om : '';
      s.mapped = !!m;
    });
    state.result = res;
    renderScope();
    renderMissing();
    renderLeaderboard();
    renderDaily();
    if (cloud.on && !isNewFile) renderTrends();
    if (cloud.on && !cloud.admin) {
      $('viewerBranches').textContent = plural(res.stores.length, 'branch', 'branches');
    }
    var T = res.totals;
    document.body.classList.add('has-data');
    $('empty').hidden = true;
    $('results').hidden = false;
    $('btnExport').disabled = false;

    $('reportTitle').textContent = longDate(T.start);
    $('fileInfo').textContent = state.fileName + ' · ' + plural(T.stores, 'branch', 'branches') + ' · ' +
      dt(T.start) + ' → ' + dt(T.end) +
      (state.parsed.warnings.length ? ' · ' + plural(state.parsed.warnings.length, 'row', 'rows') + ' skipped' : '') +
      (state.cloudReport ? ' · Saved to HM cloud by ' + state.cloudReport.uploadedBy + ', ' +
        new Date(state.cloudReport.uploadedAt).toLocaleString('en-GB', { day: 'numeric', month: 'short', hour: '2-digit', minute: '2-digit' }) :
       state.savedAt ? ' · Saved in this browser ' + new Date(state.savedAt).toLocaleString('en-GB', { day: 'numeric', month: 'short', hour: '2-digit', minute: '2-digit' }) : '');
    $('btnOriginal').hidden = !state.cloudReport;

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
      { id: 'hi', label: 'High-risk branches', value: highStores, note: 'of ' + plural(T.stores, 'branch', 'branches') + ' need follow-up', cls: 'feature', go: 'stores:High' },
      { id: 'dv', label: 'Direct violations', value: T.directLines, note: n0(T.directExcess) + ' mugs over limit · ' + plural(T.directStores, 'branch', 'branches'), cls: 'high', go: 'direct:' },
      { id: 'sr', label: 'Split-invoice runs', value: T.runs, note: T.runsHigh + ' high · ' + T.runsMedium + ' medium · ' + plural(T.runStores, 'branch', 'branches'), cls: 'med', go: 'runs:' },
      { id: 'mu', label: 'Mugs sold', value: T.mugs, note: plural(T.invoices, 'invoice', 'invoices') },
      { id: 'sa', label: 'Net sales', value: T.sales, unit: 'SAR', note: plural(T.lines, 'sale line', 'sale lines') + ' · SAR ' + sar(T.sales) },
      { id: 'zq', label: 'Zero-qty lines', value: T.zero, note: 'voids or returns to review', go: 'zero:' }
    ];
    var wrap = $('kpis');
    if (!wrap.children.length) {
      wrap.innerHTML = k.map(function (x) {
        return '<div class="kpi ' + (x.cls || '') + '" data-k="' + x.id + '"' +
          (x.go ? ' data-go="' + x.go + '" role="button" tabindex="0" aria-label="Open ' + esc(x.label) + '"' : '') + '>' +
          (x.go ? '<svg class="go" viewBox="0 0 24 24" aria-hidden="true"><path d="M7 17 17 7M9 7h8v8"/></svg>' : '') +
          '<div class="label">' + (x.cls === 'high' || x.cls === 'med' ? '<i class="dot" aria-hidden="true"></i>' : '') + esc(x.label) +
          '</div><div class="value"><span class="num">0</span>' + (x.unit ? '<span class="unit">' + x.unit + '</span>' : '') +
          '</div><div class="note"></div></div>';
      }).join('');
    }
    k.forEach(function (x) {
      var el = wrap.querySelector('[data-k="' + x.id + '"]');
      el.querySelector('.note').textContent = x.note;
      var from = isNewFile ? 0 : (state.kpiValues[x.id] || 0);
      tween(el.querySelector('.num'), from, x.value, isNewFile ? 900 : 400);
      state.kpiValues[x.id] = x.value;
    });
    replay(wrap, isNewFile ? 'enter' : 'recalc');
  }
  // restart a one-shot CSS animation class
  function replay(el, cls) {
    el.classList.remove('enter', 'recalc');
    void el.offsetWidth;
    el.classList.add(cls);
  }
  function kpiGo(el) {
    var parts = el.getAttribute('data-go').split(':');
    state.filter.store = ''; $('fStore').value = '';
    state.filter.level = parts[1] || ''; $('fLevel').value = state.filter.level;
    switchTab(parts[0]);
    var panel = document.querySelector('.table-panel');
    window.scrollTo({ top: panel.getBoundingClientRect().top + window.scrollY - 84, behavior: reduceMotion ? 'auto' : 'smooth' });
  }
  $('kpis').addEventListener('click', function (e) { var c = e.target.closest('.kpi[data-go]'); if (c) kpiGo(c); });
  $('kpis').addEventListener('keydown', function (e) {
    var c = e.target.closest('.kpi[data-go]');
    if (c && (e.key === 'Enter' || e.key === ' ')) { e.preventDefault(); kpiGo(c); }
  });
  function tween(node, from, to, dur) {
    var fmt = function (v) { return n0(Math.round(v)); };
    if (reduceMotion || from === to) { node.textContent = fmt(to); return; }
    var start = null;
    node._tw = (node._tw || 0) + 1;
    var token = node._tw;
    function step(ts) {
      if (token !== node._tw) return; // a newer tween took over
      if (start === null) start = ts;
      var p = Math.min(1, (ts - start) / dur);
      var e = 1 - Math.pow(1 - p, 4); // ease-out quart
      node.textContent = fmt(from + (to - from) * e);
      if (p < 1) requestAnimationFrame(step); else node.textContent = fmt(to);
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
    } else {
      el.classList.remove('recalc'); void el.offsetWidth; el.classList.add('recalc');
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
      { k: 'city', h: 'City', map: 1, f: function (r) { return r.mapped ? esc(r.city) : '<span class="muted">Not in mapping</span>'; } },
      { k: 'am', h: 'Area manager', map: 1, f: function (r) { return esc(r.am || ''); } },
      { k: 'rom', h: 'ROM', map: 1, f: function (r) { return esc(r.rom || ''); } },
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
      { k: 'am', h: 'Area manager', map: 1, f: function (r) { var m = branchInfo(r.store); return esc(m ? m.am : ''); }, sortVal: function (r) { var m = branchInfo(r.store); return m ? m.am : ''; } },
      { k: 'label', h: 'Item / colour' },
      { k: 'start', h: 'From', f: function (r) { return t(r.start); } },
      { k: 'end', h: 'To', f: function (r) { return t(r.end); } },
      { k: 'durationMin', h: 'Duration', num: 1, f: function (r) { return mins(r.durationMin); } },
      { k: 'invoices', h: 'Invoices', num: 1 },
      { k: 'mugs', h: 'Mugs', num: 1 },
      { k: 'excess', h: 'Over', num: 1 },
      { k: 'burst', h: '× pace', num: 1, f: function (r) { return r.burst >= 100 ? '100+' : r.burst.toFixed(1); } },
      { k: 'identical', h: 'Same basket', num: 1, f: function (r) { return r.invoices >= 3 ? pct(r.identical) : '—'; } },
      { k: 'review', h: 'Review', cloud: 1, f: function (r) { return reviewBadge(runKey(r)); }, sortVal: function (r) { var v = state.reviews[runKey(r)]; return v ? v.status : 'open'; } }
    ],
    direct: [
      { k: 'store', h: 'Branch', cls: 'strong' },
      { k: 'am', h: 'Area manager', map: 1, f: function (r) { var m = branchInfo(r.store); return esc(m ? m.am : ''); }, sortVal: function (r) { var m = branchInfo(r.store); return m ? m.am : ''; } },
      { k: 'label', h: 'Item / colour' },
      { k: 'time', h: 'Time', f: function (r) { return t(r.time); } },
      { k: 'qty', h: 'Qty', num: 1 },
      { k: 'excess', h: 'Over', num: 1 },
      { k: 'sales', h: 'SAR', num: 1, f: function (r) { return sar(r.sales); } },
      { k: 'basket', h: 'Full invoice', cls: 'wrap-cell' },
      { k: 'review', h: 'Review', cloud: 1, f: function (r) { return reviewBadge(directKey(r)); }, sortVal: function (r) { var v = state.reviews[directKey(r)]; return v ? v.status : 'open'; } }
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
        var mi = branchInfo(r.store);
        var hay = (r.store + ' ' + (r.label || '') + ' ' + (r.item || '') + ' ' + (r.items || '') + ' ' + (r.basket || '') +
          (mi ? ' ' + mi.city + ' ' + mi.zone + ' ' + mi.am + ' ' + mi.rom : '')).toLowerCase();
        if (hay.indexOf(q) < 0) return false;
      }
      return true;
    });
  }

  function renderTable() {
    var tab = state.tab, cols = COLS[tab].filter(function (c) { return (!c.cloud || cloud.on) && (!c.map || state.mapping); });
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
      (tab === 'stores' ? ' · click a branch to open it' : tab === 'runs' ? ' · click a run for invoice detail' + (cloud.on ? ' and review' : '') :
        tab === 'direct' && cloud.on ? ' · click a row to review it' : '');

    var html = '<table><thead><tr>' + cols.map(function (c) {
      var arrow = s && s.k === c.k ? (s.dir > 0 ? ' ↑' : ' ↓') : '';
      return '<th data-k="' + c.k + '" class="' + (c.num ? 'num' : '') + '" aria-sort="' +
        (s && s.k === c.k ? (s.dir > 0 ? 'ascending' : 'descending') : 'none') + '">' + esc(c.h) + '<span class="arrow">' + arrow + '</span></th>';
    }).join('') + '</tr></thead><tbody>';

    var MAX = 2000;
    rows.slice(0, MAX).forEach(function (r) {
      var id = tab === 'runs' ? runKey(r) : tab === 'direct' ? directKey(r) : tab === 'stores' ? r.store : '';
      var clickable = tab === 'runs' || tab === 'stores' || (tab === 'direct' && cloud.on);
      var open = (tab === 'runs' || tab === 'direct') && state.open[id];
      html += '<tr' + (clickable ? ' class="clickable' + (open ? ' open' : '') + '" data-id="' + esc(id) + '"' : '') + '>' + cols.map(function (c) {
        var v = c.f ? c.f(r) : esc(r[c.k]);
        return '<td class="' + (c.num ? 'num ' : '') + (c.cls || '') + '">' + v + '</td>';
      }).join('') + '</tr>';
      if (open) html += '<tr class="detail"><td colspan="' + cols.length + '"><div class="detail-inner">' +
        (tab === 'runs' ? runDetail(r) : directDetail(r)) + '</div></td></tr>';
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
    return reviewBar(runKey(r), 'run', r.store) +
      '<div class="detail-grid"><span>Score <b>' + r.score + '</b> / 100</span><span>Value <b>SAR ' + sar(r.sales) + '</b></span>' +
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

  function directDetail(d) {
    return reviewBar(directKey(d), 'direct', d.store) +
      '<div class="detail-grid"><span>Qty <b>' + d.qty + '</b></span><span>Over limit <b>' + d.excess + '</b></span>' +
      '<span>Value <b>SAR ' + sar(d.sales) + '</b></span><span>Time <b>' + t(d.time) + '</b></span></div>' +
      '<p class="small" style="margin:0">Full invoice: <b>' + esc(d.basket) + '</b></p>';
  }

  $('tableWrap').addEventListener('click', function (e) {
    var save = e.target.closest('.rv-save');
    if (save) {
      var bar = save.closest('.review-bar');
      save.disabled = true;
      cloud.saveReview(bar.getAttribute('data-key'), bar.getAttribute('data-kind'), bar.getAttribute('data-store'),
        bar.querySelector('.rv-status').value, bar.querySelector('.rv-note').value.trim()).then(function () { save.disabled = false; });
      return;
    }
    if (e.target.closest('.review-bar')) return;
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
  // The branded workbook needs ExcelJS (~0.9 MB); load it only when someone exports.
  var excelJsReady = null;
  function loadExcelJs() {
    if (window.ExcelJS) return Promise.resolve();
    if (!excelJsReady) {
      excelJsReady = new Promise(function (res, rej) {
        var s = document.createElement('script');
        s.src = 'vendor/exceljs.min.js';
        s.onload = res;
        s.onerror = function () { excelJsReady = null; rej(new Error('Could not load the Excel library.')); };
        document.head.appendChild(s);
      });
    }
    return excelJsReady;
  }
  var btnExport = $('btnExport');
  btnExport.addEventListener('click', function () {
    if (!state.result || btnExport.classList.contains('busy')) return;
    var label = btnExport.querySelector('span'), prev = label.textContent;
    btnExport.classList.add('busy'); btnExport.setAttribute('aria-busy', 'true'); label.textContent = 'Preparing…';
    loadExcelJs().then(function () {
      return window.MugsReport.build(state.result, { fileName: state.fileName, reportDate: state.parsed.reportDate,
        reviews: cloud.on ? state.reviews : null, reviewLabels: REVIEW_LABELS, runKey: runKey, directKey: directKey,
        branchInfo: state.mapping ? branchInfo : null, scopeLabel: scopeLabel() });
    }).then(function (blob) {
      var d = new Date(state.result.totals.start);
      var day = d.getUTCFullYear() + '-' + ('0' + (d.getUTCMonth() + 1)).slice(-2) + '-' + ('0' + d.getUTCDate()).slice(-2);
      var a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = 'HM Mugs SOP report ' + day + '.xlsx';
      document.body.appendChild(a); a.click(); a.remove();
      setTimeout(function () { URL.revokeObjectURL(a.href); }, 4000);
    }).catch(function (err) {
      console.error(err);
      showError('Export failed: ' + err.message);
    }).then(function () {
      btnExport.classList.remove('busy'); btnExport.removeAttribute('aria-busy'); label.textContent = prev;
    });
  });

  /* ---------- help ---------- */
  $('btnHelp').addEventListener('click', function () {
    var d = $('help');
    if (d.showModal) d.showModal(); else d.setAttribute('open', '');
  });
  $('help').addEventListener('click', function (e) { if (e.target === this) this.close(); });
})();
