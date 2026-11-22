# Cloudflare app: handover and plan

This file lets a new Claude session continue building the Cloudflare version of the Mugs SOP Monitor. Read it fully before changing anything.

## Goal

Turn the static GitHub Pages prototype into a full app on Cloudflare, so the team can upload the daily POS **"Sales By Product"** export, keep history in a database, and see the same reports from any computer behind a staff login.

The existing site (index.html, css/, js/engine.js, js/app.js, js/report.js, vendor/) stays the front end. Detection logic lives in `js/engine.js` and must not change behaviour (tests: `node tests/engine.test.js`; on the 18 Nov 2026 file it gives 701 direct violations, 235 split runs, 14 high-risk branches).

## Already done (owner did these in the Cloudflare dashboard)

| Item | Value |
|---|---|
| Cloudflare account ID | `9ca5e366700c58875109a636ebf1a04e` (in `wrangler.jsonc`, not secret) |
| R2 bucket | `mugs-sop-files` (location Automatic, Standard class) |
| D1 database | `mugs-sop`, ID `d49eb71f-940d-42fd-be94-67331f949226` (in `wrangler.jsonc`) |
| API token | Name `mugs-sop-deploy`. Permissions: Workers template (account part) + D1 Edit, Workers R2 Storage Edit, Access: Apps and Policies Edit, Access: Organizations, Identity Providers, and Groups Edit. Zone policy removed (no domain; deploy to workers.dev). |
| Environment | Owner added `CLOUDFLARE_API_TOKEN` and `CLOUDFLARE_ACCOUNT_ID` as environment variables of the Claude cloud environment. Never print, log or commit them. |

The account also contains unrelated resources (`hm-mixes-db`, `hm-mixes-order`, `baldy-card-tracker`). **Do not touch them.**

## Already in the repo (branch `claude/sleepy-albattani-kztzzl`)

- `wrangler.jsonc`: Worker `mugs-sop`, `main: worker/index.js`, static assets from `.` with `run_worker_first: ["/api/*"]`, D1 binding `DB`, R2 binding `FILES`, daily cron `17 3 * * *`, vars `RETENTION_DAYS=90`, `AUTH_MODE=access`, `ACCESS_TEAM_DOMAIN`, `ACCESS_AUD` (empty until Access is set up), observability + traces on.
- `.assetsignore`: keeps worker/, migrations/, tests/, config, node_modules, README and any xlsx/csv out of the public assets.
- `migrations/0001_init.sql`: tables `reports`, `store_daily`, `reviews` (see file).
- `package.json` with Wrangler 4 (`npm install` first). `.dev.vars.example` for local dev (`AUTH_MODE=dev`).
- `worker/index.js` (API, Access JWT check, retention cron) and the front-end cloud mode are **written and tested locally** (wrangler dev, full browser flow).
- Deploy runs from GitHub Actions on push to `main` (`.github/workflows/deploy-cloudflare.yml`) using the repository secret `CLOUDFLARE_API_TOKEN` (added by the owner). The token is NOT in the Claude environment of the original session.

## Design decisions (agreed)

1. **Free-plan friendly:** Workers free plan has ~10 ms CPU per request. So the **browser** keeps parsing the Excel file (SheetJS) and running `engine.js`. The Worker only stores and serves data; it must not parse xlsx or JSON-parse the large lines payload.
2. **Storage:**
   - R2 `files/<YYYY-MM-DD>/<id>/<original-name>`: the original upload (for audit / re-download).
   - R2 `lines/<id>.json.gz`: parsed sale lines (`parsed.lines`), gzipped in the browser with `CompressionStream('gzip')`; served back as-is and decompressed in the browser with `DecompressionStream`.
   - D1 `reports`: one row per day, with `rules` and `totals` JSON. Re-uploading the same `report_date` replaces the old report once the new one is complete.
   - D1 `store_daily`: per-branch summary per day (computed in the browser at upload with the uploader's rules) for history / repeat offenders.
   - D1 `reviews`: review decision per flag (`open | genuine | violation | false_positive`, note, reviewer email, time).
3. **Upload flow** (keeps each request small and streams files):
   1. `POST /api/reports` JSON `{reportDate, fileName, fileSize, periodStart, periodEnd, rules, totals, stores[]}` → inserts report (`status='uploading'`) + `store_daily` rows in one `DB.batch`; returns `{id}`.
   2. `PUT /api/reports/:id/file` raw body → stream to R2 (reject > 25 MB via Content-Length).
   3. `PUT /api/reports/:id/lines` raw gzip body → stream to R2.
   4. `POST /api/reports/:id/complete` → check both R2 objects exist (`head`), set `status='ready'`, delete older reports with the same date (R2 objects + rows).
4. **Other API routes**
   - `GET /api/me` → `{email, authMode, retentionDays}`. The front end uses this to detect cloud mode: if it fails (e.g. on GitHub Pages) the app stays in the current local-only mode.
   - `GET /api/reports` → ready reports, newest first (id, report_date, file_name, uploaded_by, uploaded_at, totals, rules).
   - `GET /api/reports/:id`, `GET /api/reports/:id/lines` (stream), `GET /api/reports/:id/file` (stream, `Content-Disposition: attachment`), `DELETE /api/reports/:id`.
   - `GET /api/reports/:id/reviews`; `PUT /api/reviews` JSON `{flagKey, reportId, kind, store, status, note}`. Flag key: `run|<store>|<label>|<startMs>` or `direct|<store>|<label>|<timeMs>`.
   - `GET /api/trends?days=30` → `store_daily` rows for the last N ready report dates + the list of dates.
   - JSON errors `{error}` with proper status codes; structured JSON logs.
5. **Auth: Cloudflare Access** in front of the whole app (email one-time PIN; allow list from the owner, e.g. `@halfm.sa`). The Worker **also verifies** the `Cf-Access-Jwt-Assertion` JWT on every `/api/*` call (RS256 via Web Crypto; keys from `<ACCESS_TEAM_DOMAIN>/cdn-cgi/access/certs`, cached per isolate; check `aud` = `ACCESS_AUD`, `exp`, `iss`). Email from the token = `uploaded_by` / `reviewed_by`. `AUTH_MODE=dev` (only via `.dev.vars`) skips this and uses `DEV_USER`.
6. **Retention cron:** delete reports older than `RETENTION_DAYS` (90) and any `status='uploading'` older than 1 day, from both R2 and D1.
7. **Front end, cloud mode** (only when `/api/me` succeeds):
   - After a file loads and is analysed, save it to the cloud automatically (flow above) and show "Saved to HM cloud · visible to your team". Keep IndexedDB restore for local mode only.
   - On open, load the latest ready report from the cloud.
   - **History panel:** list of days (date, file, uploaded by, high-risk / direct / runs) with Open, Download original, Delete.
   - **Trends panel:** "risk calendar" heatmap (branches × last 30 days, coloured by status) + repeat-offender table (High days, Medium days, mugs over limit, last seen).
   - **Review:** in the split-run detail (and a new direct-violation detail) a status selector + note + Save; a Review column in those tables.
   - Excel export keeps working; optionally add review status columns.
   - Keep the HM theme and motion rules already in `css/styles.css` (no animation on tables/filters).

## Steps for the next session

1. `npm install`; confirm `npx wrangler whoami` works with the env token (do not echo the token).
2. Write `worker/index.js` per the API above; run `npx wrangler types` if useful.
3. Local test: copy `.dev.vars.example` → `.dev.vars`; `npm run db:migrate:local`; `npx wrangler dev`; drive the full flow with Playwright (upload the sample file from the owner, list, open, delete, review). The sample xlsx is not in the repo (git-ignored); ask the owner to upload it again if needed.
4. Build the front-end cloud features; re-test local mode still works when `/api/me` is unavailable.
5. Deploy: `npm run db:migrate:remote`, then `npx wrangler deploy` → `https://mugs-sop.<subdomain>.workers.dev`.
6. Access: create the Zero Trust organisation if missing (the owner may need to pick the free plan in the dashboard), enable the One-time PIN login method, create a self-hosted Access application for the workers.dev hostname with an Allow policy (emails / domain from the owner), then set `ACCESS_TEAM_DOMAIN` and `ACCESS_AUD` in `wrangler.jsonc` vars and redeploy. Verify an unauthenticated `curl` to `/api/me` is refused.
7. Test live end to end, commit, push, and only open/merge PRs when the owner says so (they say "merge it").

## Owner preferences

- Plain, non-technical explanations; step-by-step instructions one at a time when they need to click things in a dashboard.
- Never ask them to paste secrets in chat.
- HM brand: deep green `#0b3d2c`, sand `#e9dcc0`, gold `#b8862c`.
