# Mugs SOP Monitor

A browser-based tool that flags branches breaking the mug SOP:
**no more than 2 mugs of the same colour per customer**.

Open the page, drop the daily POS **“Sales By Product”** export (.xlsx), and it shows:

| Check | What it finds |
|---|---|
| **Direct violations** | An invoice with more than the limit of one colour/item. |
| **Split-invoice runs** | A branch ringing the same colour on back-to-back invoices, each within the limit, that together go over it (the “2 per invoice, many invoices” workaround). |
| **Zero-qty lines** | 0-quantity / SAR 0 lines (voids, returns) for review. |

Each branch gets a risk level (High / Medium / Low / Clean). Click a branch to see its runs, click a run to see every invoice in it, and export everything to Excel.

**Privacy:** the file is read in your browser. No data is uploaded anywhere, and sales exports are git-ignored so they are never committed.

## Branch mapping

Upload the **Master Location Mapping** workbook (Detection rules panel → *Branch mapping*) to link each store to its city, zone, area manager and ROM by branch code (`B99, Mec Haram` → `B99`). It reads the `Location_Mapping` sheet and fills gaps for new branches from the `Excluded` sheet (e.g. B121 Madina Haram).

- A **View** bar scopes the whole report (KPIs, chart, tables, Excel export) to one area manager, ROM, city or zone.
- Branch, split-run and direct-violation tables gain City / Area manager / ROM columns (sortable, searchable).
- The mapping contains staff names, so it is **never committed**: it is kept in the browser (local mode) or in R2 behind the login (cloud mode, `GET/PUT/DELETE /api/mapping`).

### Weekly mapping sync

A scheduled Claude task runs every **Sunday morning (Riyadh)**: it reads `Master_Location_Mapping.xlsx` from SharePoint through the Microsoft 365 connector, parses it with `scripts/sync-mapping.mjs` (same rules as the in-app upload) and sends it to `PUT /api/mapping/sync` with the `MAPPING_SYNC_TOKEN` key. Branches an admin added by hand are kept unless the file now covers them. The key lives in two private places only: the `MAPPING_SYNC_TOKEN` GitHub secret (copied to the Worker at deploy) and the Claude environment variable of the same name.

## Excel report

**Export report** downloads `HM Mugs SOP report <date>.xlsx`:

- **Dashboard:** HM title band with date, file and rules; six KPI tiles; charts of excess mugs by branch (top 10) and by hour of day; top 10 split runs and direct violations; branches by risk level; how-to-read note. Set up to print on one landscape page width.
- **Branches, Split runs, Split run invoices, Direct violations, Zero-qty lines:** green header, coloured risk labels, bars inside cells for the key numbers, filters, frozen headers, real dates and number formats.
- **Settings:** rules used and method notes.

## How detection works

1. **Invoices** – the export has no invoice number, so lines with the same *branch + closing time (ms)* form one invoice. If a later export adds an invoice/receipt column it is picked up automatically.
2. **Direct violation** – invoice quantity of one item (or colour) > limit.
3. **Split-invoice run** – per branch and item, invoices in time order are chained while the gap is ≤ *max gap* (default 3 min). A run is flagged if it has ≥ *min invoices* (default 2), every invoice is within the limit, and the total is over the limit.
4. **Risk score 0–100** for each run:
   - invoices in the run (up to 30)
   - mugs over the limit (up to 25)
   - **burst vs normal pace** (up to 30): how much faster the branch sold this item during the run than over the rest of the day. This keeps very busy branches (e.g. Haram) from being over-flagged for a genuine queue, while a quiet branch suddenly ringing 20 invoices in 10 minutes scores high.
   - identical baskets (up to 15): invoices with exactly the same contents.

   High ≥ 60, Medium 35–59, Low < 35.

All thresholds (limit, gap, min invoices, “same colour” = same SKU or same colour across items) are adjustable in the page.

### Limits of Phase 1

The current export has no customer, card or cashier field, so a split run is *likely*, not proven. Phase 2 needs the POS to export: invoice number, cashier ID, payment method, card token / masked card + terminal, customer mobile/loyalty ID, and voids/returns. With these the same rules can link invoices to one payer and separate genuine family purchases from splits.

## Cloudflare app (team version)

The same site runs on Cloudflare with logins and a shared database: **https://mugs-sop.shelfy-help.workers.dev**

| Piece | What it does |
|---|---|
| Worker `mugs-sop` (`worker/index.js`) | Serves the site and the `/api/*` endpoints |
| D1 database `mugs-sop` (`migrations/`) | Users, sessions, reports, per-branch daily summary (with area manager / ROM / city / zone), review decisions |
| R2 bucket `mugs-sop-files` | Original uploads, sale lines (all + one file per area manager), branch mapping |
| Daily cron | Deletes reports older than `RETENTION_DAYS` (90), unfinished uploads and expired sessions |

### Logins and roles

- People **create their own account** (name, work email, password). New accounts are *pending* until an admin approves them.
- **Admin:** uploads the daily file and the branch mapping, sees every branch with the View filters, the **Who needs attention** leaderboard (group by area manager, ROM, city or zone; click to drill in), the **Day by day** chart, repeat offenders, report history (open / original / delete), saves review decisions, and manages users (approve, role, area manager, status, reset password, delete).
- **Area manager:** read-only dashboard of **their own branches only**, focused on direct violations and split-invoice runs (headline cards, day-by-day trend, branch table, run and violation lists with review status). The server enforces the scope: each upload stores a separate data file per area manager, and summaries, trends, reviews and mapping are filtered by the signed-in account. Original files are admin-only.
- The first admin: sign up with the email set in the `ADMIN_EMAIL` GitHub secret (passed to the Worker at deploy time). If that secret is empty, the very first account created becomes the admin.
- Passwords are hashed with PBKDF2-SHA256 (salted); accounts lock for 15 minutes after 8 failed attempts; sessions are HttpOnly cookies valid for 14 days and are revoked when an admin changes the account.

On GitHub Pages (no server) the page stays in local-only mode.

**Deploy:** every push to `main` runs `.github/workflows/deploy-cloudflare.yml` (tests → D1 migrations → `wrangler deploy`) using the `CLOUDFLARE_API_TOKEN` and `ADMIN_EMAIL` repository secrets.

**Local development:**

```bash
npm install
cp .dev.vars.example .dev.vars   # AUTH_MODE=dev skips login; remove it to test the login screens
npm run db:migrate:local
npx wrangler dev
```

## Branch mapping

Upload the **Master Location Mapping** workbook (Detection rules panel → *Branch mapping*) to link each store to its city, zone, area manager and ROM by branch code (`B99, Mec Haram` → `B99`). It reads the `Location_Mapping` sheet and fills gaps for new branches from the `Excluded` sheet (e.g. B121 Madina Haram).

- A **View** bar scopes the whole report (KPIs, chart, tables, Excel export) to one area manager, ROM, city or zone.
- Branch, split-run and direct-violation tables gain City / Area manager / ROM columns (sortable, searchable).
- The mapping contains staff names, so it is **never committed**: it is kept in the browser (local mode) or in R2 behind the login (cloud mode, `GET/PUT/DELETE /api/mapping`).

## Excel report

**Export report** downloads `HM Mugs SOP report <date>.xlsx`:

- **Dashboard:** HM title band with date, file and rules; six KPI tiles; charts of excess mugs by branch (top 10) and by hour of day; top 10 split runs and direct violations; branches by risk level; how-to-read note. Set up to print on one landscape page width.
- **Branches, Split runs, Split run invoices, Direct violations, Zero-qty lines:** green header, coloured risk labels, bars inside cells for the key numbers, filters, frozen headers, real dates and number formats.
- **Settings:** rules used and method notes.

## How detection works

1. **Invoices** – the export has no invoice number, so lines with the same *branch + closing time (ms)* form one invoice. If a later export adds an invoice/receipt column it is picked up automatically.
2. **Direct violation** – invoice quantity of one item (or colour) > limit.
3. **Split-invoice run** – per branch and item, invoices in time order are chained while the gap is ≤ *max gap* (default 3 min). A run is flagged if it has ≥ *min invoices* (default 2), every invoice is within the limit, and the total is over the limit.
4. **Risk score 0–100** for each run:
   - invoices in the run (up to 30)
   - mugs over the limit (up to 25)
   - **burst vs normal pace** (up to 30): how much faster the branch sold this item during the run than over the rest of the day. This keeps very busy branches (e.g. Haram) from being over-flagged for a genuine queue, while a quiet branch suddenly ringing 20 invoices in 10 minutes scores high.
   - identical baskets (up to 15): invoices with exactly the same contents.

   High ≥ 60, Medium 35–59, Low < 35.

All thresholds (limit, gap, min invoices, “same colour” = same SKU or same colour across items) are adjustable in the page.

### Limits of Phase 1

The current export has no customer, card or cashier field, so a split run is *likely*, not proven. Phase 2 needs the POS to export: invoice number, cashier ID, payment method, card token / masked card + terminal, customer mobile/loyalty ID, and voids/returns. With these the same rules can link invoices to one payer and separate genuine family purchases from splits.

## Cloudflare app (team version)

The same site also runs on Cloudflare with a shared database, so every upload is saved and visible to the team:

| Piece | What it does |
|---|---|
| Worker `mugs-sop` (`worker/index.js`) | Serves the site and the `/api/*` endpoints |
| D1 database `mugs-sop` (`migrations/`) | Reports, per-branch daily summary, review decisions |
| R2 bucket `mugs-sop-files` | Original uploaded files and parsed sale lines |
| Cloudflare Access | Staff login (email one-time code); the Worker also verifies the login token on every API call |
| Daily cron | Deletes reports older than `RETENTION_DAYS` (90) |

In cloud mode the page saves each upload automatically, opens the latest day on load, and adds **Repeat offenders** (risk calendar across days), **Report history** (open, download original, delete) and **review status** on split runs and direct violations (Open / Genuine / Violation / False positive, with a note and who saved it). The Excel export includes the review columns. On GitHub Pages, or until the login is configured, the page stays in local-only mode.

**Deploy:** every push to `main` runs `.github/workflows/deploy-cloudflare.yml` (tests → D1 migrations → `wrangler deploy`) using the `CLOUDFLARE_API_TOKEN` repository secret.

**Local development:**

```bash
npm install
cp .dev.vars.example .dev.vars
npm run db:migrate:local
npx wrangler dev
```

## Design

Half Million theme (deep green, sand and gold) with light and dark mode. Brand colours are CSS variables at the top of `css/styles.css` (`--hm-*`); replace them with the official HM brand codes if they differ. Chart colours were checked for colour-blind separation in both themes.

Motion follows Emil Kowalski's design-engineering guidelines (github.com/emilkowalski/skills): a one-time hero entrance, sections that reveal once on scroll, chart bars that grow once when first in view, a sliding tab indicator, 0.97 press feedback on buttons, and no animation on tables, filters or other actions used many times a day. `prefers-reduced-motion` turns movement off.

## Project layout

```
index.html           page
css/styles.css       styles (light + dark)
js/engine.js         parsing + detection (pure JS, also runs in Node)
js/app.js            UI: upload, tables, drill-down, export button
js/report.js         branded Excel report (Dashboard + styled sheets, charts as images)
vendor/xlsx.full.min.js  SheetJS 0.18.5 (Apache-2.0), reads the export
vendor/exceljs.min.js    ExcelJS 4.4.0 (MIT), writes the branded report; loaded only on export
tests/engine.test.js unit tests: `node tests/engine.test.js`
```

## Run locally

Open `index.html` directly in a browser, or serve the folder:

```bash
python3 -m http.server 8000
```

## Publish on GitHub Pages

Settings → Pages → *Build and deployment* → Source: **Deploy from a branch** → select the branch and `/ (root)` → Save. The site appears at `https://<owner>.github.io/<repo>/`.

Note: on a free GitHub plan, Pages needs the repository to be public. The code holds no sales data, but if the repo must stay private, use a paid plan or Cloudflare Pages instead (same static files, no build step).
