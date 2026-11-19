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

## Project layout

```
index.html           page
css/styles.css       styles (light + dark)
js/engine.js         parsing + detection (pure JS, also runs in Node)
js/app.js            UI: upload, tables, drill-down, Excel export
vendor/xlsx.full.min.js  SheetJS 0.18.5 (Apache-2.0), bundled so no CDN is needed
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
