#!/usr/bin/env node
/*
 * Weekly mapping sync.
 *
 * Input: the text the Microsoft 365 connector returns when it reads
 * Master_Location_Mapping.xlsx (tab-separated rows under "## Sheet: <name>" headings).
 * It is parsed with the same rules as the in-app upload (js/mapping.js) and sent to
 * the app with the sync key from the MAPPING_SYNC_TOKEN environment variable.
 *
 *   node scripts/sync-mapping.mjs <connector-text-file> [--dry-run] [--url https://mugs-sop.shelfy-help.workers.dev]
 */
import { readFileSync } from 'node:fs';
import { createRequire } from 'node:module';

const require = createRequire(import.meta.url);
const MAP = require('../js/mapping.js');

const args = process.argv.slice(2);
const file = args.find((a) => !a.startsWith('--'));
const dry = args.includes('--dry-run');
const urlIdx = args.indexOf('--url');
const base = (urlIdx >= 0 ? args[urlIdx + 1] : process.env.MUGS_APP_URL || 'https://mugs-sop.shelfy-help.workers.dev').replace(/\/+$/, '');
if (!file) { console.error('Usage: node scripts/sync-mapping.mjs <connector-text-file> [--dry-run] [--url <app>]'); process.exit(2); }

// The connector may wrap the text in a JSON result; unwrap if needed.
let text = readFileSync(file, 'utf8');
try { const j = JSON.parse(text); text = typeof j === 'string' ? j : (j.text || j.content || (Array.isArray(j) && j.map((x) => x.text || '').join('\n')) || text); } catch { /* plain text */ }

// "## Sheet: Location_Mapping — 119 rows × 43 columns" then header + rows (tab-separated)
const lines = text.split(/\r?\n/);
const sheets = {}, order = [];
let cur = null;
for (const line of lines) {
  const h = line.match(/^## Sheet:\s*(.+?)\s+—/);
  if (h) { cur = h[1].trim(); sheets[cur] = { rows: [] }; order.push(cur); continue; }
  if (!cur) continue;
  if (/^(Formulas|## )/.test(line)) { cur = null; continue; }
  if (line.trim()) sheets[cur].rows.push(line.split('\t'));
}
if (!order.length) { console.error('No "## Sheet:" sections found in ' + file + '. Is this the connector text of Master_Location_Mapping.xlsx?'); process.exit(1); }

const fakeXLSX = { utils: { sheet_to_json: (ws) => ws.rows } };
const m = MAP.parseWorkbook(fakeXLSX, { SheetNames: order, Sheets: sheets });
const ams = new Set(Object.values(m.branches).map((b) => b.am).filter(Boolean));
const roms = new Set(Object.values(m.branches).map((b) => b.rom).filter(Boolean));
console.log(`Parsed ${m.count} branches (${m.mainCount} from Location_Mapping, ${m.extraCount} from Excluded) · ${ams.size} area managers · ${roms.size} ROMs`);
if (m.count < 50) { console.error('Refusing to sync: fewer than 50 branches parsed, the file looks incomplete.'); process.exit(1); }
if (dry) { console.log('Dry run: nothing sent.'); process.exit(0); }

const token = process.env.MAPPING_SYNC_TOKEN;
if (!token) { console.error('MAPPING_SYNC_TOKEN is not set in this environment.'); process.exit(1); }
const res = await fetch(base + '/api/mapping/sync', {
  method: 'PUT',
  headers: { 'content-type': 'application/json', authorization: 'Bearer ' + token },
  body: JSON.stringify({ branches: m.branches, count: m.count, mainCount: m.mainCount, extraCount: m.extraCount,
    fileName: 'Master_Location_Mapping.xlsx (SharePoint, weekly sync)' })
});
const out = await res.text();
if (!res.ok) { console.error(`Sync failed: HTTP ${res.status} ${out}`); process.exit(1); }
const r = JSON.parse(out);
console.log(`Synced to ${base}: ${r.count} branches saved, ${r.keptManual} hand-added branches kept, ${r.keptExceptions || 0} area manager exceptions kept.`);
