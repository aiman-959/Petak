/**
 * Phase 4 round trips. For every sample file (xlsx from Excel and from LibreOffice, xlsm, xls,
 * ods, csv), on a phone:
 *   1. open it in Petak; LibreOffice reads the same original: the values must agree (import)
 *   2. edit one cell, save as .xlsx, open that again: values, formulas, styles, merges, frozen
 *      panes, column widths and row heights must be what was loaded plus the edit (round trip)
 *   3. LibreOffice opens Petak's file and recalculates every formula: values must agree (export)
 *   BASE_URL=http://localhost:4173/ node tests/e2e/roundtrip.mjs
 * Needs `soffice` (LibreOffice) on the PATH.
 */
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import * as XLSX from 'xlsx';
import { unzipSync, strFromU8 } from 'fflate';
import { start, BASE } from './lib.mjs';

const t = await start('roundtrip');
const { page, step, ev, idle, fresh, openFile } = t;
console.log('Petak e2e: round trips —', BASE);
await fresh();

// ------------------------------------------------------------ LibreOffice
const lo = fs.mkdtempSync(path.join(os.tmpdir(), 'petak-rt-lo-'));
const profile = path.join(lo, 'profile');
// always recalculate formulas when loading .xlsx / .ods, so cached values are not simply echoed
fs.mkdirSync(path.join(profile, 'user'), { recursive: true });
fs.writeFileSync(
  path.join(profile, 'user', 'registrymodifications.xcu'),
  `<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop></item>
<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="ODFRecalcMode" oor:op="fuse"><value>0</value></prop></item>
</oor:items>
`,
);
const profileUrl = 'file://' + profile;
function loConvert(file, to = 'xlsx:Calc MS Excel 2007 XML') {
  const out = path.join(lo, 'out-' + path.basename(file).replace(/\W/g, '_'));
  fs.mkdirSync(out, { recursive: true });
  const ext = to.split(':')[0];
  const want = path.join(out, path.basename(file).replace(/\.[^.]+$/, '') + '.' + ext);
  fs.rmSync(want, { force: true });
  const args = [`-env:UserInstallation=${profileUrl}`, '--headless', '--calc', '--convert-to', to, '--outdir', out];
  if (file.endsWith('.csv')) args.push('--infilter=CSV:44,34,76,1');
  for (let i = 0; i < 3 && !fs.existsSync(want); i++) execFileSync('soffice', [...args, file], { stdio: 'ignore', timeout: 180000 });
  if (!fs.existsSync(want)) throw new Error('LibreOffice did not convert ' + file);
  return want;
}
/** What LibreOffice computed: { sheetName: { A1: value } } */
function loValues(file) {
  const wb = XLSX.read(fs.readFileSync(loConvert(file)), { type: 'buffer' });
  const out = {};
  for (const name of wb.SheetNames) {
    const ws = wb.Sheets[name];
    const cells = {};
    for (const [ref, c] of Object.entries(ws)) {
      if (ref.startsWith('!')) continue;
      if (c.t === 'z' || c.v === undefined || c.v === '') continue;
      cells[ref] = c.t === 'e' ? c.w : c.v;
    }
    out[name] = cells;
  }
  return out;
}

// ------------------------------------------------------------- Petak side
/** Everything a round trip must keep, per sheet. */
const snapshot = () =>
  ev(() => {
    const wb = window.__doc.wb;
    const styles = wb.getWorkbook().getStyles();
    const enc = (r, c) => {
      let s = '';
      for (let n = c + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
      return s + (r + 1);
    };
    const rgb = (x) => (x?.rgb ? String(x.rgb).toLowerCase() : undefined);
    // the workbook's default font (name or size) named on a cell is the same as none on it
    const defFont = wb.getWorkbook().getSnapshot().defaultStyle?.ff;
    const defSize = wb.getWorkbook().getSnapshot().defaultStyle?.fs;
    return wb.getSheets().map((ws) => {
      const m = ws.getSheet();
      const cells = {};
      let maxR = 0;
      let maxC = 0;
      m.getCellMatrix().forValue((r, c, cell) => {
        if (!cell) return;
        const st = (typeof cell.s === 'string' ? styles.get(cell.s) : cell.s) ?? {};
        const bd = st.bd ? Object.keys(st.bd).filter((k) => st.bd[k]).sort().join('') : '';
        const e = {
          v: cell.v ?? null,
          f: cell.f || undefined,
          b: st.bl ? 1 : undefined,
          i: st.it ? 1 : undefined,
          bg: rgb(st.bg),
          cl: rgb(st.cl),
          ff: st.ff && st.ff !== defFont ? st.ff : undefined,
          fs: st.fs && st.fs !== defSize ? st.fs : undefined,
          n: st.n?.pattern && st.n.pattern !== 'General' ? st.n.pattern : undefined,
          ht: st.ht || undefined,
          bd: bd || undefined,
        };
        for (const k of Object.keys(e)) if (e[k] === undefined) delete e[k];
        if (e.v === null && Object.keys(e).length === 1) return;
        cells[enc(r, c)] = e;
        maxR = Math.max(maxR, r);
        maxC = Math.max(maxC, c);
      });
      const cols = [];
      const rows = [];
      for (let c = 0; c <= maxC; c++) cols.push(Math.round(m.getColumnWidth(c)));
      for (let r = 0; r <= maxR; r++) rows.push(Math.round(m.getRowHeight(r)));
      return {
        name: ws.getSheetName(),
        cells,
        merges: ws.getMergedRanges().map((x) => x.getA1Notation()).sort(),
        frozen: [ws.getFrozenRows(), ws.getFrozenColumns()],
        cols,
        rows,
      };
    });
  });
/** Values Petak computed, in the same shape as loValues. */
const values = (snap) => Object.fromEntries(snap.map((s) => [s.name, Object.fromEntries(Object.entries(s.cells).filter(([, c]) => c.v !== null && c.v !== '').map(([ref, c]) => [ref, c.v]))]));

const same = (a, b) => {
  if (typeof a === 'number' && typeof b === 'number') return Math.abs(a - b) <= 1e-9 + 1e-9 * Math.abs(b);
  if (typeof a === 'boolean' || typeof b === 'boolean') return Number(a) === Number(b);
  return String(a) === String(b);
};
/** Cells whose values differ between Petak and LibreOffice (ignoring sheets / cells in `skip`). */
function compare(petak, other, skip = () => false) {
  const diffs = [];
  for (const [sheet, cells] of Object.entries(petak)) {
    const o = other[sheet];
    if (!o) {
      diffs.push(`${sheet}: missing in LibreOffice`);
      continue;
    }
    const refs = new Set([...Object.keys(cells), ...Object.keys(o)]);
    for (const ref of refs) {
      if (skip(sheet, ref)) continue;
      const a = cells[ref] ?? null;
      const b = o[ref] ?? null;
      if (a === null && b === null) continue;
      if (!same(a, b)) diffs.push(`${sheet}!${ref}: Petak ${JSON.stringify(a)}, LibreOffice ${JSON.stringify(b)}`);
    }
  }
  return diffs;
}

/** XML parts of a package that are not well formed (every element that opens closes, in order) — F-08. */
function malformedParts(file) {
  const bad = [];
  for (const [name, bytes] of Object.entries(unzipSync(new Uint8Array(fs.readFileSync(file))))) {
    if (!/\.(xml|rels|vml)$/.test(name)) continue;
    const xml = strFromU8(bytes).replace(/<\?[^>]*\?>|<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>/g, '');
    const stack = [];
    let ok = true;
    for (const m of xml.matchAll(/<(\/?)([\w:.-]+)(?:\s[^>]*?)?(\/?)>/g)) {
      if (m[3]) continue;
      if (!m[1]) stack.push(m[2]);
      else if (stack.pop() !== m[2]) ok = false;
    }
    if (!ok || stack.length || /<[^<>]*</.test(xml.replace(/<[^>]*>/g, ''))) bad.push(name);
  }
  return bad;
}

/** Saves the workbook as .xlsx through the app's own writer. */
const saveXlsx = async (name) => {
  const f = path.join(t.tmp, name);
  fs.writeFileSync(f, Buffer.from(await ev(async () => Array.from(new Uint8Array((await window.__doc.serialize('xlsx')).data)))));
  return f;
};

const F = (f) => path.resolve('tests/fixtures/files', f);
// LibreOffice makes the .ods sample from the Excel roster
const ods = loConvert(F('roster.xlsx'), 'ods');

/**
 * What may differ, and why. Each entry: [fixture, (sheet, ref) => boolean, reason].
 */
const KNOWN_IMPORT = [
  ['people.csv', (sheet, ref) => ref === 'D5', 'L-01: Petak keeps the identifier 0012 as text (LibreOffice makes it the number 12)'],
];

const fixtures = [
  ['roster.xlsx', F('roster.xlsx')],
  ['roster-lo.xlsx', F('roster-lo.xlsx')],
  ['qm-report.xlsx', F('qm-report.xlsx')],
  ['qm-report-lo.xlsx', F('qm-report-lo.xlsx')],
  // the test file's VBA part is a placeholder LibreOffice refuses to open; the workbook is qm-report.xlsx
  ['macro.xlsm', F('macro.xlsm'), F('qm-report.xlsx')],
  ['legacy.xls', F('legacy.xls')],
  ['roster.ods', ods],
  ['looks.ods', F('looks.ods')],
  ['people.csv', F('people.csv')],
];
const summary = [];

for (const [label, file, loSource] of fixtures) {
  await step(`${label}: open, save, reopen, and LibreOffice agrees`, async () => {
    await openFile(file);
    await page.waitForTimeout(300);
    const first = await snapshot();
    assert.ok(first.length > 0 && Object.keys(first[0].cells).length > 0, 'something was loaded');
    const skip = KNOWN_IMPORT.filter(([f]) => f === label).map(([, fn]) => fn);
    const skipFn = (s, r) => skip.some((fn) => fn(s, r));

    // 1. import: LibreOffice's reading of the original
    const loOrig = loValues(loSource ?? file);
    if (label.endsWith('.csv')) {
      // a CSV has one sheet; LibreOffice names it after the file
      const only = Object.values(loOrig)[0];
      for (const k of Object.keys(loOrig)) delete loOrig[k];
      loOrig[first[0].name] = only;
    }
    const importDiffs = compare(values(first), loOrig, skipFn);
    assert.deepEqual(importDiffs, [], 'import differs from LibreOffice:\n' + importDiffs.slice(0, 15).join('\n'));

    // 2. edit one free cell, save, reopen
    const free = await ev(() => {
      const ws = window.__doc.wb.getSheets()[0];
      const r = ws.getLastRow() + 2;
      window.__doc.wb.setActiveSheet(ws);
      ws.getRange(r, 0).setValue({ f: '=1+1' });
      return r;
    });
    await idle();
    const edited = await snapshot();
    const saved = await saveXlsx(label.replace(/\.[^.]+$/, '') + '-petak.xlsx');
    assert.deepEqual(malformedParts(saved), [], 'the saved package has XML that is not well formed');
    await openFile(saved);
    await page.waitForTimeout(300);
    const again = await snapshot();
    assert.equal(again.length, edited.length, 'same number of sheets');
    for (let i = 0; i < edited.length; i++) {
      const a = edited[i];
      const b = again[i];
      assert.equal(b.name, a.name);
      const cellDiffs = [];
      for (const ref of new Set([...Object.keys(a.cells), ...Object.keys(b.cells)])) {
        if (JSON.stringify(a.cells[ref]) !== JSON.stringify(b.cells[ref])) cellDiffs.push(`${a.name}!${ref}: ${JSON.stringify(a.cells[ref])} → ${JSON.stringify(b.cells[ref])}`);
      }
      assert.deepEqual(cellDiffs, [], `${a.name}: cells changed by the round trip (${cellDiffs.length}):\n${cellDiffs.slice(0, 12).join('\n')}`);
      assert.deepEqual(b.merges, a.merges, `${a.name}: merges`);
      assert.deepEqual(b.frozen, a.frozen, `${a.name}: frozen panes`);
      const near = (x, y) => x.length === y.length && x.every((v, k) => Math.abs(v - y[k]) <= 1);
      assert.ok(near(b.cols, a.cols), `${a.name}: column widths ${a.cols} → ${b.cols}`);
      assert.ok(near(b.rows, a.rows), `${a.name}: row heights ${a.rows} → ${b.rows}`);
    }
    const editRef = 'A' + (free + 1);
    assert.equal(again[0].cells[editRef]?.v, 2, 'the edit is in the saved file');

    // 3. export: LibreOffice recalculates Petak's file
    const loSaved = loValues(saved);
    const exportDiffs = compare(values(edited), loSaved, skipFn);
    assert.deepEqual(exportDiffs, [], 'LibreOffice computes something else from the saved file:\n' + exportDiffs.slice(0, 15).join('\n'));

    const n = edited.reduce((s, x) => s + Object.keys(x.cells).length, 0);
    const f = edited.reduce((s, x) => s + Object.values(x.cells).filter((c) => c.f).length, 0);
    summary.push({ file: label, sheets: edited.length, cells: n, formulas: f });
  });
}

/**
 * L-03: an .ods file looks as the workbook it came from. Each pair: the .ods, and the reference
 * Petak reads as .xlsx — the Excel original for the roster (LibreOffice made the .ods from it),
 * LibreOffice's own .xlsx for looks.ods.
 */
await step('.ods files look like their .xlsx: fonts, fills, colours, borders, alignment, merges, sizes', async () => {
  // LibreOffice 24.2 writes the Default style's theme colour instead of these cells' own font colour
  // into its .xlsx (tests/unit/ods.test.ts checks them against the colours the file was made with)
  const LO_COLOUR = new Set(['Looks!G6', 'Looks!L6', 'Looks!E8', 'Looks!G8']);
  for (const [odsFile, refFile] of [
    [ods, F('roster.xlsx')],
    [F('looks.ods'), F('looks-lo.xlsx')],
  ]) {
    const view = async (file) => {
      await openFile(file);
      await page.waitForTimeout(300);
      const def = await ev(() => window.__doc.wb.getWorkbook().getSnapshot().defaultStyle ?? {});
      return { snap: await snapshot(), def };
    };
    const a = await view(refFile);
    const b = await view(odsFile);
    assert.equal(b.snap.length, a.snap.length, 'same sheets');
    const pick = (c, def) => {
      if (!c) return {};
      const o = { v: c.v ?? null, b: c.b, i: c.i, bg: c.bg, cl: c.cl === '#000000' ? undefined : c.cl, ff: c.ff, fs: c.fs === def.fs ? undefined : c.fs, ht: c.ht, bd: c.bd };
      for (const k of Object.keys(o)) if (o[k] === undefined) delete o[k];
      return o;
    };
    const diffs = [];
    for (let i = 0; i < a.snap.length; i++) {
      const x = a.snap[i];
      const y = b.snap[i];
      assert.equal(y.name, x.name);
      for (const ref of new Set([...Object.keys(x.cells), ...Object.keys(y.cells)])) {
        const p = pick(x.cells[ref], a.def);
        const q = pick(y.cells[ref], b.def);
        if (LO_COLOUR.has(`${x.name}!${ref}`)) delete p.cl, delete q.cl;
        if (typeof p.v === 'number' && typeof q.v === 'number' && Math.abs(p.v - q.v) < 1e-9) q.v = p.v;
        if (JSON.stringify(p) !== JSON.stringify(q) && (p.v !== null || Object.keys(p).length > 1 || Object.keys(q).length > 1)) diffs.push(`${x.name}!${ref}: xlsx ${JSON.stringify(p)} · ods ${JSON.stringify(q)}`);
      }
      assert.deepEqual(y.merges, x.merges, `${x.name}: merges`);
      // LibreOffice measures the default font's characters a little wider than Excel's 7 px
      x.cols.forEach((w, c) => assert.ok(Math.abs((y.cols[c] ?? w) / w - 1) < 0.1, `${x.name} column ${c + 1}: xlsx ${w} px, ods ${y.cols[c]} px`));
    }
    assert.deepEqual(diffs, [], `${path.basename(odsFile)} looks different:\n${diffs.slice(0, 12).join('\n')}`);
  }
  // the toast says what an .ods does not bring along
  const notice = await ev(() => window.__doc.loadWarnings.join(' '));
  assert.match(notice, /OpenDocument file: charts, pictures, comments and conditional formats are not shown/);
});

fs.mkdirSync('tests/e2e/out', { recursive: true });
fs.writeFileSync('tests/e2e/out/roundtrip.json', JSON.stringify(summary, null, 1) + '\n');
await t.finish();
