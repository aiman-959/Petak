/**
 * Reference results from LibreOffice for every example formula (Phase 4).
 *   node tests/functions/libreoffice.mjs      → tests/functions/libreoffice-results.json
 * Writes an .xlsx with the data block and one formula per function (each twice: as written and
 * with Excel's `_xlfn.` prefix, which newer functions need in a file), lets headless LibreOffice
 * calculate it, and reads the results back. The JSON is committed, so the unit test does not need
 * LibreOffice.
 */
import * as XLSX from 'xlsx';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { DATA, EXAMPLES } from './examples.mjs';

const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'petak-lo-'));
const ws = {};
const put = (r, c, v) => {
  const ref = XLSX.utils.encode_cell({ r, c });
  if (v === null || v === undefined) return;
  if (typeof v === 'string' && v.startsWith('=')) ws[ref] = { t: 'n', f: v.slice(1) };
  else if (typeof v === 'number') ws[ref] = { t: 'n', v };
  else if (typeof v === 'boolean') ws[ref] = { t: 'b', v };
  else ws[ref] = { t: 's', v: String(v) };
};
DATA.forEach((row, r) => row.forEach((v, c) => put(r, c, v)));
const COL = 23; // X: as written, Y: the function under test with the _xlfn. prefix newer Excel functions need
const esc = (t) => t.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
EXAMPLES.forEach(([fn, f], i) => {
  put(i, COL, f);
  put(i, COL + 1, f.replace(new RegExp(`(^=|[^A-Z0-9._])${esc(fn)}\\(`, 'g'), `$1_xlfn.${fn}(`));
});
ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: Math.max(DATA.length, EXAMPLES.length), c: COL + 1 } });
const input = path.join(dir, 'functions.xlsx');
fs.writeFileSync(input, XLSX.write({ SheetNames: ['S'], Sheets: { S: ws } }, { type: 'buffer', bookType: 'xlsx' }));
// its own profile: a first start sets the profile up and may return before converting
const profile = 'file://' + path.join(dir, 'profile');
const converted = path.join(dir, 'out', 'functions.xlsx');
for (let attempt = 0; attempt < 3 && !fs.existsSync(converted); attempt++) {
  execFileSync('soffice', [`-env:UserInstallation=${profile}`, '--headless', '--calc', '--convert-to', 'xlsx:Calc MS Excel 2007 XML', '--outdir', path.join(dir, 'out'), input], { stdio: 'ignore', timeout: 180000 });
}
if (!fs.existsSync(converted)) throw new Error('LibreOffice did not convert the file');
const out = XLSX.read(fs.readFileSync(converted), { type: 'buffer' });
const s = out.Sheets[out.SheetNames[0]];
const unknown = (c, fn) => (c?.t === 'e' && c.w === '#NAME?') || String(c?.f ?? '').includes(`${fn.toLowerCase()}(`);
const val = (c) => (!c ? { t: 'z' } : c.t === 'e' ? { t: 'e', v: c.w } : c.t === 'b' ? { t: 'b', v: !!c.v } : c.t === 'n' ? { t: 'n', v: c.v } : { t: 's', v: String(c.v ?? '') });
const results = {};
EXAMPLES.forEach(([fn], i) => {
  // LibreOffice writes a function it did not recognise back in lower case
  const x = s[XLSX.utils.encode_cell({ r: i, c: COL })];
  const y = s[XLSX.utils.encode_cell({ r: i, c: COL + 1 })];
  results[fn] = val(unknown(x, fn) && !unknown(y, fn) ? y : x);
});
const version = execFileSync('soffice', ['--version'], { encoding: 'utf8' }).trim();
fs.writeFileSync(new URL('./libreoffice-results.json', import.meta.url), JSON.stringify({ version, results }, null, 1) + '\n');
const names = Object.values(results).filter((r) => r.t === 'e' && r.v === '#NAME?').length;
console.log(`${EXAMPLES.length} formulas, LibreOffice does not know ${names} of them (${version})`);
