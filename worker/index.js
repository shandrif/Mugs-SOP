/*
 * Mugs SOP Monitor – Cloudflare Worker API
 *
 * Static site is served from assets; this Worker handles /api/* only.
 * The browser parses the Excel export and runs the detection engine
 * (keeps Worker CPU tiny); the Worker stores and serves:
 *   R2  FILES : original uploads + gzipped parsed sale lines
 *   D1  DB    : reports, per-branch daily summary, review decisions
 * Every API call is authenticated with Cloudflare Access (JWT verified here).
 */

const MAX_FILE_BYTES = 25 * 1024 * 1024;   // original Excel upload
const MAX_LINES_BYTES = 20 * 1024 * 1024;  // gzipped JSON of sale lines
const REVIEW_STATUSES = ['open', 'genuine', 'violation', 'false_positive'];
const RISK = ['High', 'Medium', 'Low', 'Clean'];

/* ---------------------------------------------------------------- helpers */

function json(data, status = 200, headers = {}) {
  return new Response(JSON.stringify(data), {
    status,
    headers: { 'content-type': 'application/json; charset=utf-8', 'cache-control': 'no-store', ...headers }
  });
}
class HttpError extends Error {
  constructor(status, message) { super(message); this.status = status; }
}
const bad = (msg) => new HttpError(400, msg);

function log(level, msg, fields = {}) {
  console[level === 'error' ? 'error' : 'log'](JSON.stringify({ level, msg, ...fields }));
}

async function readJson(request, maxBytes = 1024 * 1024) {
  const len = Number(request.headers.get('content-length') || 0);
  if (len > maxBytes) throw new HttpError(413, 'Request too large');
  const text = await request.text();
  if (text.length > maxBytes) throw new HttpError(413, 'Request too large');
  try { return JSON.parse(text); } catch { throw bad('Invalid JSON'); }
}

const isDate = (s) => typeof s === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(s);
const str = (v, max = 300) => (typeof v === 'string' ? v.slice(0, max) : '');
const num = (v) => (Number.isFinite(Number(v)) ? Number(v) : 0);
const safeName = (n) => str(n, 180).replace(/[^\w.\- ()]+/g, '_') || 'upload.xlsx';

/* ------------------------------------------------------ Cloudflare Access */

// Public signing keys are cached per isolate (not request state).
let jwksCache = { url: '', keys: null, expires: 0 };

function b64urlToBytes(s) {
  s = s.replace(/-/g, '+').replace(/_/g, '/');
  const pad = s.length % 4 ? 4 - (s.length % 4) : 0;
  const bin = atob(s + '='.repeat(pad));
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}
const b64urlJson = (s) => JSON.parse(new TextDecoder().decode(b64urlToBytes(s)));

async function accessKeys(team) {
  const url = `${team}/cdn-cgi/access/certs`;
  if (jwksCache.url === url && jwksCache.keys && Date.now() < jwksCache.expires) return jwksCache.keys;
  const res = await fetch(url);
  if (!res.ok) throw new HttpError(503, 'Could not load Access signing keys');
  const { keys } = await res.json();
  jwksCache = { url, keys, expires: Date.now() + 10 * 60 * 1000 };
  return keys;
}

function cookie(request, name) {
  const c = request.headers.get('cookie') || '';
  const m = c.match(new RegExp('(?:^|;\\s*)' + name + '=([^;]+)'));
  return m ? m[1] : null;
}

async function authenticate(request, env) {
  const mode = env.AUTH_MODE || 'access';
  if (mode === 'dev') return { email: env.DEV_USER || 'dev@local', mode };

  const team = (env.ACCESS_TEAM_DOMAIN || '').replace(/\/+$/, '');
  const aud = env.ACCESS_AUD || '';
  if (!team || !aud) throw new HttpError(503, 'Login is not configured yet (ACCESS_TEAM_DOMAIN / ACCESS_AUD).');

  const token = request.headers.get('cf-access-jwt-assertion') || cookie(request, 'CF_Authorization');
  if (!token) throw new HttpError(401, 'Not signed in');
  const parts = token.split('.');
  if (parts.length !== 3) throw new HttpError(401, 'Invalid login token');

  let header, payload;
  try { header = b64urlJson(parts[0]); payload = b64urlJson(parts[1]); } catch { throw new HttpError(401, 'Invalid login token'); }
  if (header.alg !== 'RS256') throw new HttpError(401, 'Invalid login token');

  let keys = await accessKeys(team);
  let jwk = keys.find((k) => k.kid === header.kid);
  if (!jwk) { jwksCache.expires = 0; keys = await accessKeys(team); jwk = keys.find((k) => k.kid === header.kid); }
  if (!jwk) throw new HttpError(401, 'Unknown login key');

  const key = await crypto.subtle.importKey('jwk', jwk, { name: 'RSASSA-PKCS1-v1_5', hash: 'SHA-256' }, false, ['verify']);
  const ok = await crypto.subtle.verify('RSASSA-PKCS1-v1_5', key, b64urlToBytes(parts[2]),
    new TextEncoder().encode(parts[0] + '.' + parts[1]));
  if (!ok) throw new HttpError(401, 'Invalid login token');

  const now = Math.floor(Date.now() / 1000);
  const auds = Array.isArray(payload.aud) ? payload.aud : [payload.aud];
  if (!auds.includes(aud)) throw new HttpError(403, 'Token is for another application');
  if (payload.exp && payload.exp < now) throw new HttpError(401, 'Login expired');
  if (payload.nbf && payload.nbf > now + 60) throw new HttpError(401, 'Login not yet valid');
  if (payload.iss && payload.iss.replace(/\/+$/, '') !== team) throw new HttpError(401, 'Token from another team');

  return { email: payload.email || payload.common_name || 'unknown', mode };
}

/* ----------------------------------------------------------------- storage */

async function getReport(env, id) {
  const r = await env.DB.prepare('SELECT * FROM reports WHERE id = ?').bind(id).first();
  if (!r) throw new HttpError(404, 'Report not found');
  return r;
}

function publicReport(r) {
  return {
    id: r.id, reportDate: r.report_date, periodStart: r.period_start, periodEnd: r.period_end,
    fileName: r.file_name, fileSize: r.file_size, status: r.status,
    rules: JSON.parse(r.rules || '{}'), totals: JSON.parse(r.totals || '{}'),
    uploadedBy: r.uploaded_by, uploadedAt: r.uploaded_at
  };
}

async function deleteReports(env, rows) {
  if (!rows.length) return 0;
  const keys = rows.flatMap((r) => [r.file_key, r.lines_key]).filter(Boolean);
  for (let i = 0; i < keys.length; i += 1000) await env.FILES.delete(keys.slice(i, i + 1000));
  const stmts = [];
  for (const r of rows) {
    stmts.push(env.DB.prepare('DELETE FROM reviews WHERE report_id = ?').bind(r.id));
    stmts.push(env.DB.prepare('DELETE FROM store_daily WHERE report_id = ?').bind(r.id));
    stmts.push(env.DB.prepare('DELETE FROM reports WHERE id = ?').bind(r.id));
  }
  await env.DB.batch(stmts);
  return rows.length;
}

/* ------------------------------------------------------------------ routes */

async function createReport(request, env, user) {
  const b = await readJson(request, 2 * 1024 * 1024);
  if (!isDate(b.reportDate)) throw bad('reportDate must be YYYY-MM-DD');
  if (!b.fileName) throw bad('fileName is required');
  if (!Array.isArray(b.stores) || b.stores.length > 2000) throw bad('stores must be an array (max 2000)');
  if (num(b.fileSize) > MAX_FILE_BYTES) throw new HttpError(413, 'File is larger than 25 MB');

  const id = crypto.randomUUID();
  const fileKey = `files/${b.reportDate}/${id}/${safeName(b.fileName)}`;
  const linesKey = `lines/${id}.json.gz`;
  const stmts = [
    env.DB.prepare(`INSERT INTO reports (id, report_date, period_start, period_end, file_name, file_size, file_key, lines_key,
        status, rules, totals, uploaded_by, uploaded_at) VALUES (?,?,?,?,?,?,?,?, 'uploading', ?,?,?,?)`)
      .bind(id, b.reportDate, num(b.periodStart) || null, num(b.periodEnd) || null, safeName(b.fileName), num(b.fileSize),
        fileKey, linesKey, JSON.stringify(b.rules || {}).slice(0, 4000), JSON.stringify(b.totals || {}).slice(0, 8000),
        user.email, Date.now())
  ];
  for (const s of b.stores) {
    const status = RISK.includes(s.status) ? s.status : 'Clean';
    stmts.push(env.DB.prepare(`INSERT INTO store_daily (report_id, report_date, store, status, invoices, mugs, sales,
        direct_lines, direct_excess, runs, runs_high, run_excess, pct_over) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)`)
      .bind(id, b.reportDate, str(s.store, 200), status, num(s.invoices), num(s.mugs), num(s.sales),
        num(s.directLines), num(s.directExcess), num(s.runs), num(s.runsHigh), num(s.runExcess), num(s.pctOver)));
  }
  for (let i = 0; i < stmts.length; i += 100) await env.DB.batch(stmts.slice(i, i + 100));
  log('info', 'report.created', { id, reportDate: b.reportDate, stores: b.stores.length, by: user.email });
  return json({ id }, 201);
}

async function putObject(request, env, id, which) {
  const r = await getReport(env, id);
  if (r.status !== 'uploading') throw new HttpError(409, 'Report is already complete');
  const len = Number(request.headers.get('content-length') || 0);
  const max = which === 'file' ? MAX_FILE_BYTES : MAX_LINES_BYTES;
  if (!len) throw new HttpError(411, 'Content-Length is required');
  if (len > max) throw new HttpError(413, 'Upload too large');
  if (!request.body) throw bad('Empty body');
  const key = which === 'file' ? r.file_key : r.lines_key;
  const contentType = which === 'file'
    ? (request.headers.get('content-type') || 'application/octet-stream')
    : 'application/gzip';
  // Stream straight into R2 (no buffering in the Worker).
  await env.FILES.put(key, request.body, { httpMetadata: { contentType } });
  return json({ ok: true });
}

async function completeReport(env, id) {
  const r = await getReport(env, id);
  const [f, l] = await Promise.all([env.FILES.head(r.file_key), env.FILES.head(r.lines_key)]);
  if (!f || !l) throw new HttpError(409, 'Upload is not finished (file or data missing)');
  await env.DB.prepare("UPDATE reports SET status = 'ready', file_size = ? WHERE id = ?").bind(f.size, id).run();
  // Same day uploaded again: keep only the newest.
  const { results: older } = await env.DB.prepare('SELECT * FROM reports WHERE report_date = ? AND id != ?')
    .bind(r.report_date, id).all();
  const replaced = await deleteReports(env, older);
  log('info', 'report.ready', { id, reportDate: r.report_date, replaced });
  return json({ ok: true, replaced });
}

async function listReports(env) {
  const { results } = await env.DB.prepare(
    "SELECT * FROM reports WHERE status = 'ready' ORDER BY report_date DESC, uploaded_at DESC LIMIT 500").all();
  return json({ reports: results.map(publicReport) });
}

async function streamObject(env, key, filename, type) {
  const obj = await env.FILES.get(key);
  if (!obj) throw new HttpError(404, 'File not found');
  const headers = new Headers();
  obj.writeHttpMetadata(headers);
  headers.set('etag', obj.httpEtag);
  headers.set('cache-control', 'private, no-store');
  if (type) headers.set('content-type', type);
  if (filename) headers.set('content-disposition', `attachment; filename="${filename.replace(/"/g, '')}"`);
  return new Response(obj.body, { headers });
}

async function listReviews(env, id) {
  const { results } = await env.DB.prepare(
    'SELECT flag_key, kind, store, status, note, reviewed_by, reviewed_at FROM reviews WHERE report_id = ?').bind(id).all();
  return json({ reviews: results.map((x) => ({
    flagKey: x.flag_key, kind: x.kind, store: x.store, status: x.status, note: x.note, reviewedBy: x.reviewed_by, reviewedAt: x.reviewed_at
  })) });
}

async function saveReview(request, env, user) {
  const b = await readJson(request, 32 * 1024);
  if (!b.flagKey || !b.reportId) throw bad('flagKey and reportId are required');
  if (!['run', 'direct'].includes(b.kind)) throw bad('kind must be run or direct');
  if (!REVIEW_STATUSES.includes(b.status)) throw bad('Invalid status');
  await getReport(env, str(b.reportId, 64));
  const now = Date.now();
  await env.DB.prepare(`INSERT INTO reviews (flag_key, report_id, kind, store, status, note, reviewed_by, reviewed_at)
      VALUES (?,?,?,?,?,?,?,?)
      ON CONFLICT(flag_key) DO UPDATE SET status = excluded.status, note = excluded.note,
        reviewed_by = excluded.reviewed_by, reviewed_at = excluded.reviewed_at`)
    .bind(str(b.flagKey, 600), str(b.reportId, 64), b.kind, str(b.store, 200), b.status, str(b.note, 2000), user.email, now).run();
  return json({ ok: true, reviewedBy: user.email, reviewedAt: now });
}

async function trends(env, url) {
  const days = Math.min(Math.max(parseInt(url.searchParams.get('days') || '30', 10) || 30, 1), 365);
  const { results: dates } = await env.DB.prepare(
    "SELECT DISTINCT report_date FROM reports WHERE status = 'ready' ORDER BY report_date DESC LIMIT ?").bind(days).all();
  if (!dates.length) return json({ dates: [], rows: [] });
  const from = dates[dates.length - 1].report_date;
  const { results: rows } = await env.DB.prepare(
    `SELECT d.report_date, d.store, d.status, d.mugs, d.direct_excess, d.run_excess, d.runs_high
       FROM store_daily d JOIN reports r ON r.id = d.report_id
      WHERE r.status = 'ready' AND d.report_date >= ?
      ORDER BY d.report_date`).bind(from).all();
  return json({ dates: dates.map((d) => d.report_date).reverse(), rows });
}

async function handleApi(request, env, url) {
  const user = await authenticate(request, env);
  const p = url.pathname.replace(/\/+$/, '');
  const m = request.method;
  let mm;

  if (p === '/api/me' && m === 'GET') {
    return json({ email: user.email, authMode: user.mode, retentionDays: Number(env.RETENTION_DAYS || 90) });
  }
  if (p === '/api/reports' && m === 'GET') return listReports(env);
  if (p === '/api/reports' && m === 'POST') return createReport(request, env, user);
  if (p === '/api/trends' && m === 'GET') return trends(env, url);
  if (p === '/api/reviews' && m === 'PUT') return saveReview(request, env, user);

  if ((mm = p.match(/^\/api\/reports\/([0-9a-f-]{36})(?:\/(file|lines|complete|reviews))?$/))) {
    const [, id, sub] = mm;
    if (!sub && m === 'GET') return json(publicReport(await getReport(env, id)));
    if (!sub && m === 'DELETE') {
      const r = await getReport(env, id);
      await deleteReports(env, [r]);
      log('info', 'report.deleted', { id, by: user.email });
      return json({ ok: true });
    }
    if (sub === 'file' && m === 'PUT') return putObject(request, env, id, 'file');
    if (sub === 'lines' && m === 'PUT') return putObject(request, env, id, 'lines');
    if (sub === 'complete' && m === 'POST') return completeReport(env, id);
    if (sub === 'file' && m === 'GET') { const r = await getReport(env, id); return streamObject(env, r.file_key, r.file_name); }
    if (sub === 'lines' && m === 'GET') { const r = await getReport(env, id); return streamObject(env, r.lines_key, null, 'application/gzip'); }
    if (sub === 'reviews' && m === 'GET') return listReviews(env, id);
  }
  throw new HttpError(404, 'Not found');
}

/* ------------------------------------------------------------------ export */

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    if (!url.pathname.startsWith('/api/')) return env.ASSETS.fetch(request);
    try {
      return await handleApi(request, env, url);
    } catch (err) {
      const status = err instanceof HttpError ? err.status : 500;
      if (status >= 500) log('error', 'api.error', { path: url.pathname, method: request.method, error: String(err && err.stack || err) });
      return json({ error: status >= 500 && !(err instanceof HttpError) ? 'Server error' : err.message }, status);
    }
  },

  // Daily clean-up: reports past retention, and uploads that never completed.
  async scheduled(controller, env, ctx) {
    const days = Number(env.RETENTION_DAYS || 90);
    const cutoff = new Date(Date.now() - days * 86400000).toISOString().slice(0, 10);
    const { results: old } = await env.DB.prepare(
      "SELECT * FROM reports WHERE report_date < ? OR (status = 'uploading' AND uploaded_at < ?)")
      .bind(cutoff, Date.now() - 86400000).all();
    const n = await deleteReports(env, old);
    log('info', 'retention.run', { cutoff, deleted: n });
  }
};
