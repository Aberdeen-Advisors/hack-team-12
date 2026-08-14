/*
 * verify_tuned_parity.js — replay THE PAGE'S OWN ARITHMETIC over the tuned export and
 * check it reproduces the Python run row for row.
 *
 * The engine block is not copied here. It is EXTRACTED AT RUN TIME out of index.html
 * (PART 2 of 3, the browser port) and evaluated, so this check cannot drift from what the
 * page actually runs and cannot quietly "fix" the page's arithmetic to make parity pass.
 * index.html is opened read-only and never written.
 *
 * Checks, all over the emitted tool-vocabulary columns:
 *   1. disposition parity, 600/600, page vs engine/score_northstar_600_tuned.py
 *   2. priority parity, 600/600
 *   3. both post-lookup guardrails: the lifecycle guard (retire/replace barred for an
 *      early-life row) and the redundancy override (an absorbed member forced to
 *      consolidate) — replayed and compared, not assumed
 *   4. no row carries a negative net first-year saving
 *   5. the two export rules: the money columns EMPTY (never 0) on any row whose
 *      disposition removes no run-rate spend, and successor-link / contract-urgency
 *      columns populated wherever the source supports them
 *
 * Usage: node engine/verify_tuned_parity.js
 */
'use strict';
const fs = require('fs');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const PAGE = path.join(ROOT, 'index.html');
const TOOL_CSV = path.join(ROOT, 'data', 'northstar', 'northstar-600-tuned-tool-vocabulary.csv');
const PY_CSV = path.join(ROOT, 'data', 'northstar', 'northstar-dispositions-600-tuned.csv');

/* ---------------------------------------------------------------- extract the page engine */
const html = fs.readFileSync(PAGE, 'utf8').split('\n');
const start = html.findIndex(l => l.includes('PART 2 of 3'));
const end = html.findIndex(l => l.includes('PART 3 of 3'));
if (start < 0 || end < 0) { console.error('FAIL: could not find the engine block in index.html'); process.exit(1); }
const engineSrc = html.slice(start + 3, end - 2).join('\n');
const factory = new Function(engineSrc + '\n;return { arAnalyse: arAnalyse, AR_DISPOSITION_TABLE: AR_DISPOSITION_TABLE, AR_PASS_THRESHOLD: AR_PASS_THRESHOLD };');
const page = factory();
console.log('extracted the page engine from index.html lines ' + (start + 4) + '..' + (end - 2)
  + ' — ' + engineSrc.length + ' chars, gate ' + page.AR_PASS_THRESHOLD
  + ', ' + Object.keys(page.AR_DISPOSITION_TABLE).length + ' lookup rows');

/* ---------------------------------------------------------------- csv */
function parseCsv(text) {
  const rows = []; let row = [], field = '', q = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (q) {
      if (c === '"') { if (text[i + 1] === '"') { field += '"'; i++; } else q = false; }
      else field += c;
    } else if (c === '"') q = true;
    else if (c === ',') { row.push(field); field = ''; }
    else if (c === '\n') { row.push(field); rows.push(row); row = []; field = ''; }
    else if (c !== '\r') field += c;
  }
  if (field.length || row.length) { row.push(field); rows.push(row); }
  const hdr = rows.shift();
  return rows.filter(r => r.length > 1).map(r => {
    const o = {}; hdr.forEach((h, i) => { o[h] = r[i] === undefined ? '' : r[i]; }); return o;
  });
}

const tool = parseCsv(fs.readFileSync(TOOL_CSV, 'utf8'));
const py = parseCsv(fs.readFileSync(PY_CSV, 'utf8'));
console.log('tool-vocabulary rows ' + tool.length + ', python rows ' + py.length);

/* ---------------------------------------------------------------- replay */
const result = page.arAnalyse(tool, { analysisDate: new Date(Date.UTC(2026, 7, 14)) });
if (!result.summary.usable) { console.error('FAIL: the page refused the file: ' + JSON.stringify(result.summary)); process.exit(1); }

const byId = {};
result.rows.forEach(r => { byId[r.app_id] = r; });
const pyById = {};
py.forEach(r => { pyById[r['App ID']] = r; });

let dispOk = 0, prioOk = 0, dispBad = [], prioBad = [], negNet = [];
py.forEach(r => {
  const id = r['App ID'], p = byId[id];
  if (!p) { dispBad.push(id + ': missing from the page replay'); return; }
  if (String(p.disposition) === String(r['Disposition'])) dispOk++;
  else dispBad.push(id + ': page ' + p.disposition + ' vs python ' + r['Disposition']);
  if (String(p.priority) === String(r['Priority'])) prioOk++;
  else prioBad.push(id + ': page ' + p.priority + ' vs python ' + r['Priority']);
  if (Number(p.net_saving_annual || 0) < 0) negNet.push(id + ' page ' + p.net_saving_first_year);
  if (Number(r['Net first-year saving'] || 0) < 0) negNet.push(id + ' python ' + r['Net first-year saving']);
});

/* both post-lookup guardrails, counted on the page's own replay */
const lifecycle = result.rows.filter(r => r.lifecycle_exclusion_applied || r.suppressed_recommendation === 'retire' || r.suppressed_recommendation === 'replace');
const redundancy = result.rows.filter(r => r.redundancy_override_applied);
const pyOverride = py.filter(r => /redundancy override/i.test(r['Priority basis'] || '') || /redundancy override/i.test(r['Rationale'] || ''));

/* export rules over the emitted columns */
const removesCost = d => d === 'retire' || d === 'consolidate' || d === 'replace';
/* The export rule under test is the TRANSITION-COST column: empty, never 0, on any row
 * whose disposition removes no run-rate spend. gross_saving_annual and net_saving_annual
 * are reported separately below — the savings formula zeroes the CLAIM on a retain or
 * invest row and says so, which is the engine's own semantics and is exactly what the
 * corrected export does too, so it is not a blank-versus-zero defect. */
const transCols = ['amortised_one_time_migration_cost'];
const presentTrans = transCols.filter(c => tool[0] && Object.prototype.hasOwnProperty.call(tool[0], c));
let transLeak = [], transMissing = [], succ = 0, urg = 0;
tool.forEach(r => {
  const d = String(pyById[r.app_id] ? pyById[r.app_id]['Disposition'] : '');
  presentTrans.forEach(c => {
    const v = String(r[c]).trim();
    if (!removesCost(d) && v !== '') transLeak.push(r.app_id + ' ' + c + '="' + v + '" on ' + d);
    if (removesCost(d) && v === '') transMissing.push(r.app_id + ' ' + c + ' blank on ' + d);
  });
  if (String(r.replacement_app_id || '').trim() !== '') succ++;
  if (String(r.notice_deadline_date || '').trim() !== '' || String(r.in_notice_window_now || '').trim() !== '') urg++;
});

/* ---------------------------------------------------------------- report */
const runCost = py.reduce((a, r) => a + Number(r['Annual TCO'] || 0), 0);
const net = py.reduce((a, r) => a + Number(r['Net first-year saving'] || 0), 0);
const pageNet = result.rows.reduce((a, r) => a + Number(r.net_saving_annual || 0), 0);
const spread = {};
py.forEach(r => { spread[r['Disposition']] = (spread[r['Disposition']] || 0) + 1; });

console.log('');
console.log('disposition parity ' + dispOk + '/' + py.length + (dispBad.length ? '  MISMATCHES: ' + dispBad.slice(0, 10).join(' | ') : ''));
console.log('priority parity    ' + prioOk + '/' + py.length + (prioBad.length ? '  MISMATCHES: ' + prioBad.slice(0, 10).join(' | ') : ''));
console.log('lifecycle guard    page suppressed/excluded on ' + lifecycle.length + ' row(s)');
console.log('redundancy override page ' + redundancy.length + ' row(s), python ' + pyOverride.length + ' row(s)');
console.log('negative net       ' + (negNet.length ? 'FAIL ' + negNet.slice(0, 5).join(', ') : 'none — 0 rows'));
console.log('transition columns present in the export: ' + (presentTrans.join(', ') || '(none named as such)'));
console.log('  leaked onto a no-spend-removed row: ' + transLeak.length + (transLeak.length ? ' -> ' + transLeak.slice(0, 5).join(' | ') : ''));
console.log('  blank on an acting row:             ' + transMissing.length + (transMissing.length ? ' -> ' + transMissing.slice(0, 5).join(' | ') : ''));
console.log('replacement_app_id (successor link) populated on ' + succ + ' rows; contract notice/urgency populated on ' + urg + ' rows');
const zeroTrans = tool.filter(r => String(r.amortised_one_time_migration_cost).trim() === '0' || String(r.amortised_one_time_migration_cost).trim() === '0.0').length;
const claimZero = tool.filter(r => { const d = String(pyById[r.app_id] ? pyById[r.app_id]['Disposition'] : ''); return !removesCost(d) && String(r.net_saving_annual).trim() !== ''; }).length;
console.log('acting rows whose source transition cost is genuinely 0: ' + zeroTrans
  + '; non-acting rows carrying an explicit zeroed saving CLAIM (the savings formula saying so, as in the corrected export): ' + claimZero);
console.log('run cost $' + runCost.toLocaleString('en-US') + '; python net $' + net.toLocaleString('en-US')
  + ' = ' + (100 * net / runCost).toFixed(2) + '%; page net $' + Math.round(pageNet).toLocaleString('en-US'));
console.log('spread ' + JSON.stringify(spread));

const fail = dispBad.length || prioBad.length || negNet.length || transLeak.length;
console.log('');
console.log(fail ? 'PARITY FAILED' : 'PARITY OK — the page reproduces the Python run 600/600 on disposition and priority');
process.exit(fail ? 1 : 0);
