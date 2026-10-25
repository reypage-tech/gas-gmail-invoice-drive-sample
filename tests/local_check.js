// Local logic check with Node and stubbed Gmail, Drive and Sheets objects (not real Google).
// Run: node tests/local_check.js   (exits 1 if any assertion fails)
const fs = require('fs'), vm = require('vm'), assert = require('assert');
const csv = f => fs.readFileSync(f, 'utf8').trim().split('\n').map(l => l.split(','));
const tabs = { 'Config': csv('sample_data/config.csv'), 'Price list': csv('sample_data/price_list.csv') };
const sheet = n => ({ getDataRange: () => ({ getValues: () => tabs[n].map(r => r.slice()) }), getLastRow: () => tabs[n].length, appendRow: r => tabs[n].push(r) });
const ss = { getSheetByName: n => tabs[n] ? sheet(n) : null, insertSheet: n => { tabs[n] = []; return sheet(n); } };
const emails = JSON.parse(fs.readFileSync('sample_data/emails.json', 'utf8'));
const labels = {}, threadLabels = {};
function folder(path) {
  const kids = {}, files = [];
  return { path, kids, files,
    getFoldersByName: name => { let done = !kids[name]; return { hasNext: () => !done, next: () => { done = true; return kids[name]; } }; },
    createFolder: name => (kids[name] = folder(path + '/' + name)),
    createFile: blob => { const f = { name: blob, setName(x) { this.name = x; return this; }, getUrl() { return 'https://drive.example/' + path + '/' + this.name; } }; files.push(f); return f; } };
}
const root = folder('ROOT');
function search(q) {
  const ids = [...new Set(emails.map(e => e.thread))].filter(t => !(threadLabels[t] || []).includes('invoices-filed'));
  return ids.map(t => ({ getMessages: () => emails.filter(e => e.thread === t).map(e => ({
      getId: () => e.id, getFrom: () => e.from, getSubject: () => e.subject, getPlainBody: () => e.body, getDate: () => new Date(e.date),
      getAttachments: () => e.attachments.map(a => ({ getName: () => a, copyBlob: () => a })) })),
    addLabel: l => (threadLabels[t] = (threadLabels[t] || []).concat(l.name)) }));
}
const ctx = { console,
  SpreadsheetApp: { getActiveSpreadsheet: () => ss },
  GmailApp: { search, getUserLabelByName: n => labels[n] || null, createLabel: n => (labels[n] = { name: n }) },
  DriveApp: { getFolderById: id => { if (id !== 'FAKE_ROOT_FOLDER_ID') throw new Error('no folder'); return root; } },
  Session: { getScriptTimeZone: () => 'UTC' },
  Utilities: { formatDate: d => d.toISOString().slice(0, 7) } };
vm.createContext(ctx); vm.runInContext(fs.readFileSync('src/Code.gs', 'utf8'), ctx);

const r1 = ctx.processInvoices();
for (const n of ['Invoice log', 'Mismatch log']) { console.log('== ' + n); tabs[n].forEach(r => console.log(r.map(x => x instanceof Date ? x.toISOString().slice(0, 10) : x).join(' | '))); }
const tree = (f, d = 0) => { f.files.forEach(x => console.log('  '.repeat(d) + '- ' + x.name)); Object.values(f.kids).forEach(k => { console.log('  '.repeat(d) + k.path.split('/').pop() + '/'); tree(k, d + 1); }); };
console.log('== Drive'); tree(root);
let pass = 0; const t = (c, m) => { assert.ok(c, m); pass++; console.log('PASS', m); };
t(r1.emails === 3 && r1.rows === 5 && r1.files === 4, 'run 1: 3 emails, 5 rows, 4 files');
const checks = tabs['Invoice log'].slice(1).map(r => r[7]);
t(JSON.stringify(checks) === JSON.stringify(['OK', 'MISMATCH', 'MISMATCH', 'OK', 'NOT IN PRICE LIST']), 'checks: OK, MISMATCH, MISMATCH (14.60 vs 14.00 over 2%), OK, NOT IN PRICE LIST');
t(tabs['Mismatch log'].length === 3 && tabs['Mismatch log'][1][6] === 0.3, 'mismatch log: 2 rows, difference 0.30');
t(root.kids.Maplesupply && root.kids.Maplesupply.kids['2026-09'] && root.kids.Maplesupply.kids['2026-10'], 'files split by vendor and month');
const r2 = ctx.processInvoices();
t(r2.emails === 0 && tabs['Invoice log'].length === 6, 'run 2: labeled threads not processed again');
t(tabs['Run log'].length === 3, 'run log: one row per run');
console.log(`ALL ${pass} CHECKS PASSED`);
