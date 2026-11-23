/*
 * Mugs SOP Monitor – Cloudflare Worker API
 *
 * Static site is served from assets; this Worker handles /api/* only.
 * The browser parses the Excel export and runs the detection engine
 * (keeps Worker CPU tiny); the Worker stores and serves:
 *   R2  FILES : original uploads, gzipped sale lines (all + one per area manager), branch mapping
 *   D1  DB    : users, sessions, reports, per-branch daily summary, review decisions
 *
 * Accounts are managed in the app:
 *   admin  – upload data and mapping, see everything, review flags, manage users
 *   viewer – an area manager; read-only, and the server only returns their own branches
 */

const MAX_FILE_BYTES = 90 * 1024 * 1024;   // original Excel upload (a month or more of data)
const MAX_LINES_BYTES = 20 * 1024 * 1024;  // gzipped JSON of sale lines
const REVIEW_STATUSES = ['open', 'genuine', 'violation', 'false_positive'];
const RISK = ['High', 'Medium', 'Low', 'Clean'];
const MAPPING_KEY = 'mapping/current.json';   // branch → city / zone / area manager / ROM
const SESSION_COOKIE = 'mugs_session';
const SESSION_DAYS = 14;
// PBKDF2 work factor: kept modest so a login fits the Workers free-plan CPU budget;
// combined with per-account lockout after repeated failures.
const PBKDF2_ITER = 20000;
const MAX_FAILED = 8;
const LOCK_MS = 15 * 60 * 1000;

/* ---------------------------------------------------------------- helpers */

function json(data, status = 200, headers = {}) {
  return new Response(JSON.stringify(data), {
    status,
    headers: { 'content-type': 'application/json; charset=utf-8', 'cache-control': 'no-store', ...headers }
  });
}
class HttpError extends Error {
  constructor(status, message, extra) { super(message); this.status = status; this.extra = extra; }
}
const bad = (msg) => new HttpError(400, msg);
const forbidden = () => new HttpError(403, 'Only an admin can do this');

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
const enc = new TextEncoder();
const b64 = (buf) => btoa(String.fromCharCode(...new Uint8Array(buf)));
const unb64 = (s) => Uint8Array.from(atob(s), (c) => c.charCodeAt(0));
const hex = (buf) => [...new Uint8Array(buf)].map((x) => x.toString(16).padStart(2, '0')).join('');
async function sha256Hex(s) { return hex(await crypto.subtle.digest('SHA-256', enc.encode(s))); }

/* ------------------------------------------------------------- passwords */

async function hashPassword(password, saltB64, iter) {
  const salt = saltB64 ? unb64(saltB64) : crypto.getRandomValues(new Uint8Array(16));
  const key = await crypto.subtle.importKey('raw', enc.encode(password), 'PBKDF2', false, ['deriveBits']);
  const bits = await crypto.subtle.deriveBits({ name: 'PBKDF2', hash: 'SHA-256', salt, iterations: iter }, key, 256);
  return { hash: b64(bits), salt: b64(salt), iter };
}
// Constant-time comparison: hash both sides to a fixed size first.
async function sameSecret(a, b) {
  const [x, y] = await Promise.all([crypto.subtle.digest('SHA-256', enc.encode(a)), crypto.subtle.digest('SHA-256', enc.encode(b))]);
  return crypto.subtle.timingSafeEqual(x, y);
}
function checkPassword(p) {
  if (typeof p !== 'string' || p.length < 8) throw bad('Password must be at least 8 characters');
  if (p.length > 200) throw bad('Password is too long');
}
function checkEmail(e) {
  e = str(e, 200).trim().toLowerCase();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e)) throw bad('Enter a valid email address');
  return e;
}

/* --------------------------------------------------------------- sessions */

function cookie(request, name) {
  const c = request.headers.get('cookie') || '';
  const m = c.match(new RegExp('(?:^|;\\s*)' + name + '=([^;]+)'));
  return m ? m[1] : null;
}
function sessionCookie(token, url, maxAge) {
  const secure = url.protocol === 'https:' ? '; Secure' : '';
  return `${SESSION_COOKIE}=${token}; Path=/; HttpOnly; SameSite=Lax; Max-Age=${maxAge}${secure}`;
}
async function createSession(env, userId) {
  const token = hex(crypto.getRandomValues(new Uint8Array(32)));
  const now = Date.now();
  await env.DB.prepare('INSERT INTO sessions (token_hash, user_id, created_at, expires_at) VALUES (?,?,?,?)')
    .bind(await sha256Hex(token), userId, now, now + SESSION_DAYS * 86400000).run();
  return token;
}
function publicUser(u) {
  return { id: u.id, email: u.email, name: u.name, role: u.role, areaManager: u.area_manager || '', status: u.status,
    createdAt: u.created_at, lastLogin: u.last_login };
}

async function currentUser(request, env) {
  if (env.AUTH_MODE === 'dev') {
    return { id: 'dev', email: env.DEV_USER || 'dev@local', name: 'Developer', role: env.DEV_ROLE || 'admin',
      area_manager: env.DEV_AM || '', status: 'active' };
  }
  const token = cookie(request, SESSION_COOKIE);
  if (!token || !/^[0-9a-f]{64}$/.test(token)) return null;
  const row = await env.DB.prepare(
    `SELECT u.* , s.expires_at FROM sessions s JOIN users u ON u.id = s.user_id WHERE s.token_hash = ?`)
    .bind(await sha256Hex(token)).first();
  if (!row || row.expires_at < Date.now() || row.status !== 'active') return null;
  return row;
}

/* ------------------------------------------------------------ auth routes */

async function signup(request, env, url) {
  const b = await readJson(request, 16 * 1024);
  const email = checkEmail(b.email);
  const name = str(b.name, 100).trim();
  if (!name) throw bad('Enter your name');
  checkPassword(b.password);
  const exists = await env.DB.prepare('SELECT id FROM users WHERE email = ?').bind(email).first();
  if (exists) throw new HttpError(409, 'An account with this email already exists. Sign in instead.');
  const pending = await env.DB.prepare("SELECT COUNT(*) AS n FROM users WHERE status = 'pending'").first();
  if (pending.n >= 200) throw new HttpError(429, 'Too many accounts are waiting for approval. Ask an admin.');

  // Bootstrap: the configured admin email (or, if none is configured, the very first account) becomes an active admin.
  const admins = await env.DB.prepare("SELECT COUNT(*) AS n FROM users WHERE role = 'admin'").first();
  const adminEmails = String(env.ADMIN_EMAIL || '').toLowerCase().split(',').map((s) => s.trim()).filter(Boolean);
  const isAdmin = adminEmails.includes(email) || (!adminEmails.length && admins.n === 0);

  const pw = await hashPassword(b.password, null, PBKDF2_ITER);
  const id = crypto.randomUUID();
  await env.DB.prepare(`INSERT INTO users (id, email, name, pw_hash, pw_salt, pw_iter, role, area_manager, status, created_at)
      VALUES (?,?,?,?,?,?,?,?,?,?)`)
    .bind(id, email, name, pw.hash, pw.salt, pw.iter, isAdmin ? 'admin' : 'viewer', str(b.areaManager, 120) || null,
      isAdmin ? 'active' : 'pending', Date.now()).run();
  log('info', 'user.signup', { id, admin: isAdmin });
  if (!isAdmin) return json({ ok: true, status: 'pending' }, 201);
  const token = await createSession(env, id);
  return json({ ok: true, status: 'active' }, 201, { 'set-cookie': sessionCookie(token, url, SESSION_DAYS * 86400) });
}

async function login(request, env, url) {
  const b = await readJson(request, 16 * 1024);
  const email = checkEmail(b.email);
  if (typeof b.password !== 'string') throw bad('Enter your password');
  const u = await env.DB.prepare('SELECT * FROM users WHERE email = ?').bind(email).first();
  const fail = () => new HttpError(401, 'Email or password is incorrect');
  if (!u) { await hashPassword(b.password, null, PBKDF2_ITER); throw fail(); } // same work either way
  if (u.locked_until > Date.now()) throw new HttpError(429, 'Too many attempts. Try again in 15 minutes.');
  const pw = await hashPassword(b.password, u.pw_salt, u.pw_iter);
  if (!(await sameSecret(pw.hash, u.pw_hash))) {
    const failed = u.failed_logins + 1;
    await env.DB.prepare('UPDATE users SET failed_logins = ?, locked_until = ? WHERE id = ?')
      .bind(failed >= MAX_FAILED ? 0 : failed, failed >= MAX_FAILED ? Date.now() + LOCK_MS : 0, u.id).run();
    throw fail();
  }
  if (u.status === 'pending') throw new HttpError(403, 'Your account is waiting for an admin to approve it.');
  if (u.status !== 'active') throw new HttpError(403, 'Your account is disabled. Contact an admin.');
  await env.DB.prepare('UPDATE users SET failed_logins = 0, locked_until = 0, last_login = ? WHERE id = ?').bind(Date.now(), u.id).run();
  const token = await createSession(env, u.id);
  log('info', 'user.login', { id: u.id });
  return json({ ok: true, user: publicUser(u) }, 200, { 'set-cookie': sessionCookie(token, url, SESSION_DAYS * 86400) });
}

async function logout(request, env, url) {
  const token = cookie(request, SESSION_COOKIE);
  if (token && /^[0-9a-f]{64}$/.test(token)) await env.DB.prepare('DELETE FROM sessions WHERE token_hash = ?').bind(await sha256Hex(token)).run();
  return json({ ok: true }, 200, { 'set-cookie': sessionCookie('', url, 0) });
}

async function changeOwnPassword(request, env, user) {
  const b = await readJson(request, 16 * 1024);
  const pw = await hashPassword(String(b.current || ''), user.pw_salt, user.pw_iter);
  if (!(await sameSecret(pw.hash, user.pw_hash))) throw new HttpError(401, 'Current password is incorrect');
  checkPassword(b.password);
  const next = await hashPassword(b.password, null, PBKDF2_ITER);
  await env.DB.prepare('UPDATE users SET pw_hash = ?, pw_salt = ?, pw_iter = ? WHERE id = ?').bind(next.hash, next.salt, next.iter, user.id).run();
  return json({ ok: true });
}

/* ------------------------------------------------------------ user admin */

async function listUsers(env) {
  const { results } = await env.DB.prepare('SELECT * FROM users ORDER BY status = \'pending\' DESC, name').all();
  return json({ users: results.map(publicUser) });
}

async function updateUser(request, env, admin, id) {
  const b = await readJson(request, 16 * 1024);
  const u = await env.DB.prepare('SELECT * FROM users WHERE id = ?').bind(id).first();
  if (!u) throw new HttpError(404, 'User not found');
  const role = b.role !== undefined ? b.role : u.role;
  const status = b.status !== undefined ? b.status : u.status;
  if (!['admin', 'viewer'].includes(role)) throw bad('Invalid role');
  if (!['pending', 'active', 'disabled'].includes(status)) throw bad('Invalid status');
  const am = b.areaManager !== undefined ? (str(b.areaManager, 120).trim() || null) : u.area_manager;
  const name = b.name !== undefined ? (str(b.name, 100).trim() || u.name) : u.name;
  if (role === 'viewer' && status === 'active' && !am) throw bad('Choose which area manager this person is before activating them');
  // Never leave the app without an active admin.
  if (u.role === 'admin' && (role !== 'admin' || status !== 'active')) {
    const n = await env.DB.prepare("SELECT COUNT(*) AS n FROM users WHERE role = 'admin' AND status = 'active' AND id != ?").bind(id).first();
    if (n.n === 0) throw bad('There must be at least one active admin');
  }
  const stmts = [env.DB.prepare('UPDATE users SET role = ?, status = ?, area_manager = ?, name = ? WHERE id = ?').bind(role, status, am, name, id)];
  if (b.password) {
    checkPassword(b.password);
    const pw = await hashPassword(b.password, null, PBKDF2_ITER);
    stmts.push(env.DB.prepare('UPDATE users SET pw_hash = ?, pw_salt = ?, pw_iter = ?, failed_logins = 0, locked_until = 0 WHERE id = ?')
      .bind(pw.hash, pw.salt, pw.iter, id));
  }
  // Role, status or password changes sign the person out everywhere.
  if (b.password || role !== u.role || status !== 'active' || am !== u.area_manager) stmts.push(env.DB.prepare('DELETE FROM sessions WHERE user_id = ?').bind(id));
  await env.DB.batch(stmts);
  log('info', 'user.updated', { id, by: admin.id, role, status });
  const fresh = await env.DB.prepare('SELECT * FROM users WHERE id = ?').bind(id).first();
  return json({ user: publicUser(fresh) });
}

async function deleteUser(env, admin, id) {
  if (id === admin.id) throw bad('You cannot delete your own account');
  const u = await env.DB.prepare('SELECT * FROM users WHERE id = ?').bind(id).first();
  if (!u) throw new HttpError(404, 'User not found');
  await env.DB.batch([
    env.DB.prepare('DELETE FROM sessions WHERE user_id = ?').bind(id),
    env.DB.prepare('DELETE FROM users WHERE id = ?').bind(id)
  ]);
  log('info', 'user.deleted', { id, by: admin.id });
  return json({ ok: true });
}

/* ----------------------------------------------------------------- storage */

async function getReport(env, id) {
  const r = await env.DB.prepare('SELECT * FROM reports WHERE id = ?').bind(id).first();
  if (!r) throw new HttpError(404, 'Report not found');
  return r;
}

// Area-manager totals for one report, from the per-branch summary.
async function amTotals(env, reportId, am) {
  const t = await env.DB.prepare(`SELECT COUNT(*) AS stores, SUM(status = 'High') AS highStores, SUM(invoices) AS invoices, SUM(mugs) AS mugs,
      SUM(sales) AS sales, SUM(direct_lines) AS directLines, SUM(direct_excess) AS directExcess, SUM(runs) AS runs,
      SUM(runs_high) AS runsHigh, SUM(run_excess) AS runExcess FROM store_daily WHERE report_id = ? AND am = ?`).bind(reportId, am).first();
  return Object.fromEntries(Object.entries(t || {}).map(([k, v]) => [k, v || 0]));
}

function publicReport(r, totalsOverride) {
  return {
    id: r.id, reportDate: r.report_date, periodStart: r.period_start, periodEnd: r.period_end,
    fileName: r.file_name, fileSize: r.file_size, status: r.status,
    rules: JSON.parse(r.rules || '{}'), totals: totalsOverride || JSON.parse(r.totals || '{}'),
    uploadedBy: r.uploaded_by, uploadedAt: r.uploaded_at
  };
}

const amKey = async (reportId, am) => `lines/${reportId}/am/${(await sha256Hex(am.toUpperCase())).slice(0, 24)}.json.gz`;

async function deleteReports(env, rows) {
  if (!rows.length) return 0;
  const keys = [];
  const ids = rows.map((r) => r.id);
  for (const r of rows) {
    // An original file shared by several days is removed only when no other day uses it.
    const others = await env.DB.prepare(`SELECT COUNT(*) AS n FROM reports WHERE file_key = ? AND id NOT IN (${ids.map(() => '?').join(',')})`)
      .bind(r.file_key, ...ids).first();
    if (!others || others.n === 0) keys.push(r.file_key);
    keys.push(r.lines_key);
    let cursor;
    do {
      const page = await env.FILES.list({ prefix: `lines/${r.id}/`, cursor });
      page.objects.forEach((o) => keys.push(o.key));
      cursor = page.truncated ? page.cursor : undefined;
    } while (cursor);
  }
  const clean = keys.filter(Boolean);
  for (let i = 0; i < clean.length; i += 1000) await env.FILES.delete(clean.slice(i, i + 1000));
  const stmts = [];
  for (const r of rows) {
    stmts.push(env.DB.prepare('DELETE FROM reviews WHERE report_id = ?').bind(r.id));
    stmts.push(env.DB.prepare('DELETE FROM store_daily WHERE report_id = ?').bind(r.id));
    stmts.push(env.DB.prepare('DELETE FROM reports WHERE id = ?').bind(r.id));
  }
  await env.DB.batch(stmts);
  return rows.length;
}

/* ---------------------------------------------------------- report routes */

async function createReport(request, env, user) {
  const b = await readJson(request, 2 * 1024 * 1024);
  if (!isDate(b.reportDate)) throw bad('reportDate must be YYYY-MM-DD');
  if (!b.fileName) throw bad('fileName is required');
  if (!Array.isArray(b.stores) || b.stores.length > 2000) throw bad('stores must be an array (max 2000)');
  if (num(b.fileSize) > MAX_FILE_BYTES) throw new HttpError(413, 'File is larger than 90 MB');

  const id = crypto.randomUUID();
  // Several days from one upload share the original file (uploaded once with the first day).
  let fileKey = `files/${b.reportDate}/${id}/${safeName(b.fileName)}`;
  if (b.fileFrom) {
    const src = await env.DB.prepare('SELECT file_key FROM reports WHERE id = ?').bind(str(b.fileFrom, 64)).first();
    if (!src || !src.file_key) throw bad('fileFrom report not found');
    fileKey = src.file_key;
  }
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
        direct_lines, direct_excess, runs, runs_high, run_excess, pct_over, am, rom, city, zone) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)`)
      .bind(id, b.reportDate, str(s.store, 200), status, num(s.invoices), num(s.mugs), num(s.sales),
        num(s.directLines), num(s.directExcess), num(s.runs), num(s.runsHigh), num(s.runExcess), num(s.pctOver),
        str(s.am, 120) || null, str(s.rom, 120) || null, str(s.city, 80) || null, str(s.zone, 120) || null));
  }
  for (let i = 0; i < stmts.length; i += 100) await env.DB.batch(stmts.slice(i, i + 100));
  log('info', 'report.created', { id, reportDate: b.reportDate, stores: b.stores.length, by: user.email });
  return json({ id }, 201);
}

async function putObject(request, env, id, key, contentType) {
  const r = await getReport(env, id);
  if (r.status !== 'uploading') throw new HttpError(409, 'Report is already complete');
  const len = Number(request.headers.get('content-length') || 0);
  const max = contentType === 'application/gzip' ? MAX_LINES_BYTES : MAX_FILE_BYTES;
  if (!len) throw new HttpError(411, 'Content-Length is required');
  if (len > max) throw new HttpError(413, 'Upload too large');
  if (!request.body) throw bad('Empty body');
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

async function listReports(env, user) {
  if (user.role === 'admin') {
    const { results } = await env.DB.prepare(
      "SELECT * FROM reports WHERE status = 'ready' ORDER BY report_date DESC, uploaded_at DESC LIMIT 500").all();
    return json({ reports: results.map((r) => publicReport(r)) });
  }
  // Area manager: only days that include their branches, with their own totals.
  const { results } = await env.DB.prepare(`SELECT r.*, COUNT(d.store) AS stores, SUM(d.status = 'High') AS highStores,
      SUM(d.mugs) AS mugs, SUM(d.direct_lines) AS directLines, SUM(d.direct_excess) AS directExcess, SUM(d.runs) AS runs,
      SUM(d.runs_high) AS runsHigh, SUM(d.run_excess) AS runExcess
      FROM reports r JOIN store_daily d ON d.report_id = r.id AND d.am = ?
      WHERE r.status = 'ready' GROUP BY r.id ORDER BY r.report_date DESC LIMIT 500`).bind(user.area_manager || '').all();
  return json({ reports: results.map((r) => publicReport(r, {
    stores: r.stores || 0, highStores: r.highStores || 0, mugs: r.mugs || 0, directLines: r.directLines || 0,
    directExcess: r.directExcess || 0, runs: r.runs || 0, runsHigh: r.runsHigh || 0, runExcess: r.runExcess || 0
  })) });
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

async function listReviews(env, user, id) {
  const { results } = await env.DB.prepare(
    'SELECT flag_key, kind, store, status, note, reviewed_by, reviewed_at FROM reviews WHERE report_id = ?').bind(id).all();
  let rows = results;
  if (user.role !== 'admin') {
    const { results: mine } = await env.DB.prepare('SELECT store FROM store_daily WHERE report_id = ? AND am = ?').bind(id, user.area_manager || '').all();
    const set = new Set(mine.map((m) => m.store));
    rows = rows.filter((x) => set.has(x.store));
  }
  return json({ reviews: rows.map((x) => ({
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
    .bind(str(b.flagKey, 600), str(b.reportId, 64), b.kind, str(b.store, 200), b.status, str(b.note, 2000), user.name || user.email, now).run();
  return json({ ok: true, reviewedBy: user.name || user.email, reviewedAt: now });
}

async function trends(env, url, user) {
  const days = Math.min(Math.max(parseInt(url.searchParams.get('days') || '30', 10) || 30, 1), 365);
  const am = user.role === 'admin' ? null : (user.area_manager || '');
  const { results: dates } = am === null
    ? await env.DB.prepare("SELECT DISTINCT report_date FROM reports WHERE status = 'ready' ORDER BY report_date DESC LIMIT ?").bind(days).all()
    : await env.DB.prepare(`SELECT DISTINCT r.report_date FROM reports r JOIN store_daily d ON d.report_id = r.id AND d.am = ?
        WHERE r.status = 'ready' ORDER BY r.report_date DESC LIMIT ?`).bind(am, days).all();
  if (!dates.length) return json({ dates: [], rows: [] });
  const from = dates[dates.length - 1].report_date;
  const sql = `SELECT d.report_date, d.store, d.status, d.mugs, d.direct_lines, d.direct_excess, d.runs, d.run_excess, d.runs_high,
        d.am, d.rom, d.city, d.zone
       FROM store_daily d JOIN reports r ON r.id = d.report_id
      WHERE r.status = 'ready' AND d.report_date >= ?` + (am === null ? '' : ' AND d.am = ?') + ' ORDER BY d.report_date';
  const stmt = env.DB.prepare(sql);
  const { results: rows } = am === null ? await stmt.bind(from).all() : await stmt.bind(from, am).all();
  return json({ dates: dates.map((d) => d.report_date).reverse(), rows });
}

async function getMapping(env, user) {
  const obj = await env.FILES.get(MAPPING_KEY);
  if (!obj) return json(null);
  if (user.role === 'admin') {
    return new Response(obj.body, { headers: { 'content-type': 'application/json; charset=utf-8', 'cache-control': 'no-store' } });
  }
  // Area managers only get their own branches.
  const m = await obj.json();
  const pick = (src) => Object.fromEntries(Object.entries(src || {}).filter(([, b]) => b && b.am && b.am.toUpperCase() === String(user.area_manager || '').toUpperCase()));
  return json({ branches: pick(m.branches), manual: pick(m.manual), count: 0, fileName: m.fileName, savedAt: m.savedAt });
}

async function saveMapping(env, b, savedBy) {
  if (!b || typeof b.branches !== 'object' || !Object.keys(b.branches).length) throw bad('branches is required');
  const rec = { branches: b.branches, manual: (b.manual && typeof b.manual === 'object') ? b.manual : {}, count: num(b.count),
    mainCount: num(b.mainCount), extraCount: num(b.extraCount), fileName: str(b.fileName, 200), savedAt: Date.now(), savedBy };
  await env.FILES.put(MAPPING_KEY, JSON.stringify(rec), { httpMetadata: { contentType: 'application/json' } });
  log('info', 'mapping.saved', { count: rec.count, by: savedBy });
  return rec;
}

async function syncMapping(request, env) {
  const key = String(env.MAPPING_SYNC_TOKEN || '');
  const auth = request.headers.get('authorization') || '';
  const given = auth.startsWith('Bearer ') ? auth.slice(7).trim() : '';
  if (key.length < 24) throw new HttpError(503, 'Mapping sync is not configured (MAPPING_SYNC_TOKEN)');
  if (!given || !(await sameSecret(given, key))) throw new HttpError(401, 'Invalid sync key');
  const b = await readJson(request, 2 * 1024 * 1024);
  // Keep branches an admin added by hand, unless the new file now covers them.
  const prevObj = await env.FILES.get(MAPPING_KEY);
  const prev = prevObj ? await prevObj.json() : null;
  const manual = {};
  Object.entries((prev && prev.manual) || {}).forEach(([code, v]) => { if (!b.branches[code]) manual[code] = v; });
  const rec = await saveMapping(env, { ...b, manual }, 'Weekly SharePoint sync');
  return json({ ok: true, count: rec.count, keptManual: Object.keys(manual).length, savedAt: rec.savedAt });
}

/* ------------------------------------------------------------------ router */

async function handleApi(request, env, url) {
  const p = url.pathname.replace(/\/+$/, '');
  const m = request.method;
  let mm;

  // Public endpoints
  if (p === '/api/auth/signup' && m === 'POST') return signup(request, env, url);
  if (p === '/api/auth/login' && m === 'POST') return login(request, env, url);
  if (p === '/api/auth/logout' && m === 'POST') return logout(request, env, url);

  // Weekly automation: replaces the mapping with a sync key instead of a login.
  if (p === '/api/mapping/sync' && m === 'PUT') return syncMapping(request, env);

  const user = await currentUser(request, env);
  if (!user) throw new HttpError(401, 'Not signed in', { auth: 'app' });
  const admin = user.role === 'admin';
  const needAdmin = () => { if (!admin) throw forbidden(); };

  if (p === '/api/me' && m === 'GET') {
    return json({ ...publicUser(user), retentionDays: Number(env.RETENTION_DAYS || 90), authMode: env.AUTH_MODE === 'dev' ? 'dev' : 'app' });
  }
  if (p === '/api/me/password' && m === 'POST') return changeOwnPassword(request, env, user);

  if (p === '/api/users' && m === 'GET') { needAdmin(); return listUsers(env); }
  if ((mm = p.match(/^\/api\/users\/([0-9a-f-]{36})$/))) {
    needAdmin();
    if (m === 'PATCH') return updateUser(request, env, user, mm[1]);
    if (m === 'DELETE') return deleteUser(env, user, mm[1]);
  }

  if (p === '/api/reports' && m === 'GET') return listReports(env, user);
  if (p === '/api/reports' && m === 'POST') { needAdmin(); return createReport(request, env, user); }
  if (p === '/api/trends' && m === 'GET') return trends(env, url, user);
  if (p === '/api/reviews' && m === 'PUT') { needAdmin(); return saveReview(request, env, user); }
  if (p === '/api/mapping' && m === 'GET') return getMapping(env, user);
  if (p === '/api/mapping' && m === 'PUT') {
    needAdmin();
    const rec = await saveMapping(env, await readJson(request, 2 * 1024 * 1024), user.name || user.email);
    return json({ ok: true, savedBy: rec.savedBy, savedAt: rec.savedAt });
  }
  if (p === '/api/mapping' && m === 'DELETE') { needAdmin(); await env.FILES.delete(MAPPING_KEY); return json({ ok: true }); }

  if ((mm = p.match(/^\/api\/reports\/([0-9a-f-]{36})(?:\/(file|lines|complete|reviews))?$/))) {
    const [, id, sub] = mm;
    if (!sub && m === 'GET') {
      const r = await getReport(env, id);
      return json(publicReport(r, admin ? null : await amTotals(env, id, user.area_manager || '')));
    }
    if (!sub && m === 'DELETE') {
      needAdmin();
      const r = await getReport(env, id);
      await deleteReports(env, [r]);
      log('info', 'report.deleted', { id, by: user.id });
      return json({ ok: true });
    }
    if (sub === 'file' && m === 'PUT') {
      needAdmin();
      const r = await getReport(env, id);
      return putObject(request, env, id, r.file_key, request.headers.get('content-type') || 'application/octet-stream');
    }
    if (sub === 'lines' && m === 'PUT') {
      needAdmin();
      const r = await getReport(env, id);
      const am = url.searchParams.get('am');
      return putObject(request, env, id, am ? await amKey(id, am) : r.lines_key, 'application/gzip');
    }
    if (sub === 'complete' && m === 'POST') { needAdmin(); return completeReport(env, id); }
    if (sub === 'file' && m === 'GET') { needAdmin(); const r = await getReport(env, id); return streamObject(env, r.file_key, r.file_name); }
    if (sub === 'lines' && m === 'GET') {
      const r = await getReport(env, id);
      if (admin) return streamObject(env, r.lines_key, null, 'application/gzip');
      if (!user.area_manager) throw new HttpError(403, 'No area manager is linked to your account yet');
      return streamObject(env, await amKey(id, user.area_manager), null, 'application/gzip');
    }
    if (sub === 'reviews' && m === 'GET') return listReviews(env, user, id);
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
      return json({ error: status >= 500 && !(err instanceof HttpError) ? 'Server error' : err.message, ...(err.extra || {}) }, status);
    }
  },

  // Daily clean-up: reports past retention, unfinished uploads, expired sessions.
  async scheduled(controller, env) {
    const days = Number(env.RETENTION_DAYS || 90);
    const cutoff = new Date(Date.now() - days * 86400000).toISOString().slice(0, 10);
    const { results: old } = await env.DB.prepare(
      "SELECT * FROM reports WHERE report_date < ? OR (status = 'uploading' AND uploaded_at < ?)")
      .bind(cutoff, Date.now() - 86400000).all();
    const n = await deleteReports(env, old);
    await env.DB.prepare('DELETE FROM sessions WHERE expires_at < ?').bind(Date.now()).run();
    log('info', 'retention.run', { cutoff, deleted: n });
  }
};
