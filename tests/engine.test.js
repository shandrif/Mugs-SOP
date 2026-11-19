// Run: node tests/engine.test.js   (no dependencies)
const assert = require('assert');
const E = require('../js/engine.js');

// Build a synthetic export in the same shape as the POS "Sales By Product" report.
const H = ['Store', 'Accounting day', 'Item group', 'Item code', 'Item', 'Sold with item', 'Closing time', 'Net Sales (after discount), SAR.', 'Number of items'];
const ts = (h, m, s) => `2026-11-18 ${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}:${String(s).padStart(2, '0')}.000`;
const rows = [['Sales By Product'], ['Restaurant name: Test'], ['Date: 11/18/2026'], [], H];
// Store A: one direct violation (5 beige) + a split run of 4 invoices x 2 black, 30s apart.
rows.push(['A Store', '2026-11-18', 'RETAIL', 'B1', 'Beige Mug/كوب بيج', null, ts(9, 0, 0), 100, 5]);
rows.push([null, null, null, 'K1', 'Black Mug/كوب أسود', null, ts(10, 0, 0), 40, 2]);
rows.push([null, null, null, null, null, null, ts(10, 0, 30), 40, 2]);
rows.push([null, null, null, null, null, null, ts(10, 1, 0), 40, 2]);
rows.push([null, null, null, null, null, null, ts(10, 1, 30), 40, 2]);
rows.push([null, null, null, null, null, null, ts(15, 0, 0), 20, 1]);
rows.push([null, null, null, null, null, ' Total', null, 180, 9]);
// Store B: two normal sales far apart (no flags) and one zero-qty line.
rows.push(['B Store', '2026-11-18', 'RETAIL', 'B1', 'Beige Mug/كوب بيج', null, ts(8, 0, 0), 40, 2]);
rows.push([null, null, null, null, null, null, ts(12, 0, 0), 40, 2]);
rows.push([null, null, null, null, null, null, ts(13, 0, 0), 0, 0]);
rows.push(['Grand Total', null, null, null, null, null, null, 500, 20]);

const p = E.parseRows(rows);
assert.strictEqual(p.lines.length, 9);
assert.strictEqual(p.reportDate, '11/18/2026');
assert.strictEqual(E.colourOf('Hot Mug – Beige / كوب ساخن – بيج'), 'Beige');

const a = E.analyse(p.lines, { limit: 2, gapMinutes: 3, minInvoices: 2, groupBy: 'sku' });
assert.strictEqual(a.direct.length, 1);
assert.strictEqual(a.direct[0].excess, 3);
assert.strictEqual(a.runs.length, 1);
assert.strictEqual(a.runs[0].store, 'A Store');
assert.strictEqual(a.runs[0].invoices, 4);
assert.strictEqual(a.runs[0].mugs, 8);
assert.strictEqual(a.zero.length, 1);
assert.strictEqual(a.stores.find(s => s.store === 'B Store').status, 'Clean');
assert.strictEqual(a.stores.find(s => s.store === 'A Store').status === 'Clean', false);

// Excel serial numbers keep millisecond precision.
assert.strictEqual(E.fmtTime(E.toMs(46344.41298423611), true), '2026-11-18 09:54:41');

// Optional real-file check: node tests/engine.test.js path/to/export.xlsx (needs `npm i xlsx`).
if (process.argv[2]) {
  const XLSX = require('xlsx');
  const ws = XLSX.readFile(process.argv[2]).Sheets;
  const first = ws[Object.keys(ws)[0]];
  const r = E.analyse(E.parseRows(XLSX.utils.sheet_to_json(first, { header: 1, raw: true, defval: null })).lines);
  console.log(r.totals);
}
console.log('All engine tests passed.');
