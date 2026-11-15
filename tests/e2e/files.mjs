/**
 * Files: Save / Don't save / Cancel before a workbook is replaced (H-06), clear messages when
 * a file cannot be opened, recent files.
 *   BASE_URL=http://localhost:4173/ node tests/e2e/files.mjs
 */
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { start, BASE } from './lib.mjs';

const t = await start('files');
const { page, step, ev, fresh, typeIn, val } = t;
console.log('Petak e2e files —', BASE);
await fresh();
await page.waitForTimeout(300);
const F = (f) => path.resolve('tests/fixtures/files', f);
const name = () => ev(() => window.__doc.snapshot.file.name);
const dirty = () => ev(() => window.__doc.snapshot.dirty);
/** the toast text; with `re` waits until a toast matching it is shown (an earlier toast may still be up) */
const toast = async (re) => {
  if (!re) return (await page.waitForSelector('.toast', { timeout: 15000 })).textContent();
  await page.waitForFunction((src) => new RegExp(src).test(document.querySelector('.toast')?.textContent ?? ''), re.source, { timeout: 15000 });
  return page.textContent('.toast');
};
const pick = async (click, file) => {
  const [chooser] = await Promise.all([page.waitForEvent('filechooser'), click()]);
  await chooser.setFiles(file);
};

await step('unsaved changes: the question offers Save, Don\'t save and Cancel; Cancel changes nothing', async () => {
  await typeIn('A1', 'precious');
  assert.equal(await dirty(), true);
  await page.click('[data-testid=open]');
  await page.waitForSelector('[data-testid=dlg-ask]');
  assert.match(await page.textContent('[data-testid=ask-text]'), /“Untitled\.xlsx” has unsaved changes/);
  for (const id of ['ask-save', 'ask-discard', 'ask-cancel']) assert.ok(await page.isVisible(`[data-testid=${id}]`), id);
  await page.click('[data-testid=ask-cancel]');
  assert.equal(await page.$('[data-testid=dlg-ask]'), null);
  assert.equal(await val('A1'), 'precious');
  assert.equal(await dirty(), true);
});

await step('Save: the workbook is saved first, then the other file opens', async () => {
  await page.click('[data-testid=open]');
  const [dl] = await Promise.all([page.waitForEvent('download'), pick(() => page.click('[data-testid=ask-save]'), F('roster.xlsx'))]);
  const out = path.join(t.tmp, dl.suggestedFilename());
  await dl.saveAs(out);
  assert.equal(dl.suggestedFilename(), 'Untitled.xlsx');
  assert.ok(fs.statSync(out).size > 1000);
  await page.waitForFunction(() => window.__doc.snapshot.file.name === 'roster.xlsx', null, { timeout: 30000 });
  assert.equal(await dirty(), false);
  // the saved file really has the text
  const { unzipSync, strFromU8 } = await import('fflate');
  const z = unzipSync(new Uint8Array(fs.readFileSync(out)));
  assert.ok(Object.keys(z).some((k) => strFromU8(z[k]).includes('precious')), 'saved workbook contains the typed text');
});

await step('Don\'t save: opens the other file without saving', async () => {
  await typeIn('A30', 'throw away');
  await page.click('[data-testid=open]');
  await pick(() => page.click('[data-testid=ask-discard]'), F('qm-report.xlsx'));
  await page.waitForFunction(() => window.__doc.snapshot.file.name === 'qm-report.xlsx', null, { timeout: 30000 });
  assert.equal(await dirty(), false);
});

await step('New workbook asks too; without changes nothing is asked', async () => {
  await typeIn('H1', 'x');
  await page.click('[data-testid=menu-more]');
  await page.click('text=New');
  await page.waitForSelector('[data-testid=dlg-ask]');
  await page.click('[data-testid=ask-discard]');
  await page.waitForFunction(() => window.__doc.snapshot.file.name === 'Untitled.xlsx');
  // clean workbook: straight to the picker
  await pick(() => page.click('[data-testid=open]'), F('people.csv'));
  await page.waitForFunction(() => window.__doc.snapshot.file.name === 'people.csv', null, { timeout: 30000 });
});

await step('a damaged file: clear message, the open workbook stays', async () => {
  const bad = path.join(t.tmp, 'broken.xlsx');
  const src = fs.readFileSync(F('roster.xlsx'));
  fs.writeFileSync(bad, src.subarray(0, Math.floor(src.length / 3))); // cut off: still starts like a zip
  await pick(() => page.click('[data-testid=open]'), bad);
  const msg = await toast(/Could not open/);
  assert.match(msg, /Could not open “broken\.xlsx”/);
  assert.match(msg, /damaged or incomplete/);
  assert.match(msg, /unchanged/);
  assert.equal(await name(), 'people.csv');
  assert.ok(await ev(() => window.__doc.wb.getActiveSheet().getRange('A1').getValue()), 'the workbook is still there');
  await page.waitForSelector('.toast', { state: 'detached', timeout: 15000 });
});

await step('an empty file says so', async () => {
  const empty = path.join(t.tmp, 'empty.xlsx');
  fs.writeFileSync(empty, '');
  await pick(() => page.click('[data-testid=open]'), empty);
  assert.match(await toast(/empty\.xlsx/), /Could not open “empty\.xlsx”: the file is empty/);
  assert.equal(await name(), 'people.csv');
  await page.waitForSelector('.toast', { state: 'detached', timeout: 15000 });
});

await step('recent files: listed newest first; one that is gone is explained and removed', async () => {
  await ev(async () => { await window.__recent.clear(); await window.__recent.add({ name: 'old.xlsx', uri: 'content://x/old' }); await window.__recent.add({ name: 'gone.xlsx', uri: 'content://x/gone' }); await window.__recent.add({ name: 'old.xlsx', uri: 'content://x/old' }); return 0; });
  await page.click('[data-testid=menu-more]');
  await page.click('[data-testid=menu-recent]');
  await page.waitForSelector('[data-testid=dlg-recent]');
  await page.waitForSelector('[data-testid=recent-item]');
  const items = await page.$$eval('[data-testid=recent-item] b', (els) => els.map((e) => e.textContent));
  assert.deepEqual(items, ['old.xlsx', 'gone.xlsx']);
  await page.click('[data-testid=recent-item] >> nth=1');
  const msg = await toast(/gone\.xlsx/);
  assert.match(msg, /“gone\.xlsx” could not be opened again/);
  assert.deepEqual(await ev(() => window.__recent.list().then((l) => l.map((x) => x.name))), ['old.xlsx']);
  assert.equal(await name(), 'people.csv');
});

await step('Back closes the top-most dialog or sheet, one at a time; with nothing open it is left to Android', async () => {
  await page.waitForSelector('.toast', { state: 'detached', timeout: 15000 }).catch(() => {});
  assert.equal(await ev(() => window.__petakBack()), false);
  await page.click('[data-testid=menu-more]');
  await page.click('[data-testid=menu-recent]');
  await page.waitForSelector('[data-testid=dlg-recent]');
  assert.equal(await ev(() => window.__petakBack()), true);
  assert.equal(await page.$('[data-testid=dlg-recent]'), null);
  // the unsaved-changes question: Back is Cancel
  await typeIn('J1', 'x');
  await page.click('[data-testid=open]');
  await page.waitForSelector('[data-testid=dlg-ask]');
  assert.equal(await ev(() => window.__petakBack()), true);
  assert.equal(await page.$('[data-testid=dlg-ask]'), null);
  assert.equal(await val('J1'), 'x');
  assert.equal(await ev(() => window.__petakBack()), false);
  await ev(() => { window.__doc.emit?.({ dirty: false }); return 0; });
});

await step('an old .xls opens with its fonts, colours, borders and sizes; Save writes a new .xlsx', async () => {
  await page.waitForSelector('.toast', { state: 'detached', timeout: 15000 }).catch(() => {});
  await pick(() => page.click('[data-testid=open]'), F('legacy.xls'));
  await page.waitForFunction(() => window.__doc.snapshot.file.name === 'legacy.xls', null, { timeout: 30000 });
  assert.match(await toast(/legacy\.xls/), /Old \.xls format/);
  const info = await ev(() => {
    const ws = window.__doc.wb.getActiveSheet();
    const st = (ref) => ws.getRange(ref).getCellStyleData();
    return { a1: st('A1'), a3: st('A3'), c4: st('C4'), colA: ws.getColumnWidth(0), colB: ws.getColumnWidth(1), row1: ws.getRowHeight(0), row10: ws.getRowHeight(9) };
  });
  assert.equal(info.a1.bg.rgb, '#14532d');
  assert.equal(info.a1.fs, 22);
  assert.equal(info.a1.bl, 1);
  assert.equal(info.a3.bg.rgb, '#fbbf24');
  assert.ok(info.a3.bd?.b?.s, 'header border');
  assert.ok(info.c4.bg?.rgb && info.c4.bd?.l, 'shift cell fill and border');
  assert.ok(info.colB > info.colA * 2.5, JSON.stringify(info));
  assert.equal(info.row1, 56);
  await page.screenshot({ path: '/tmp/claude-0/-home-user-Petak/d2438ac6-d1d7-5083-8327-ceef5a77a0b0/scratchpad/xls.png' });
  // Save never overwrites the .xls: it offers a new .xlsx
  await typeIn('H20', 'edited');
  const [dl] = await Promise.all([page.waitForEvent('download'), page.click('[data-testid=save]')]);
  assert.equal(dl.suggestedFilename(), 'legacy.xlsx');
});

await step('rich text in a cell is shown as formatted runs; typing over it replaces it', async () => {
  const { unzipSync, zipSync, strFromU8, strToU8 } = await import('fflate');
  const z = unzipSync(new Uint8Array(fs.readFileSync(F('roster.xlsx'))));
  const xml = strFromU8(z['xl/sharedStrings.xml']);
  const first = /<si><t[^>]*>([^<]*)<\/t><\/si>/.exec(xml);
  z['xl/sharedStrings.xml'] = strToU8(xml.replace(first[0], '<si><r><t xml:space="preserve">Plain </t></r><r><rPr><b/><color rgb="FFFF0000"/></rPr><t xml:space="preserve">bold red</t></r><r><rPr><i/></rPr><t xml:space="preserve"> note</t></r></si>'));
  const rich = path.join(t.tmp, 'rich.xlsx');
  fs.writeFileSync(rich, zipSync(z));
  await t.openFile(rich);
  const where = await ev(() => {
    const ws = window.__doc.wb.getActiveSheet();
    for (let r = 0; r < 20; r++) for (let c = 0; c < 12; c++) { const d = ws.getRange(r, c).getCellData(); if (d?.p) return { r, c, runs: d.p.body.textRuns.length, text: d.p.body.dataStream, ref: ws.getRange(r, c).getA1Notation() }; }
    return null;
  });
  assert.ok(where, 'a cell with a rich-text document');
  assert.equal(where.text, 'Plain bold red note\r\n');
  assert.equal(where.runs, 3);
  await page.waitForTimeout(300);
  await page.screenshot({ path: '/tmp/claude-0/-home-user-Petak/d2438ac6-d1d7-5083-8327-ceef5a77a0b0/scratchpad/rich.png', clip: { x: 0, y: 56, width: 412, height: 260 } });
  await typeIn(where.ref, 'retyped');
  const after = await ev(([r, c]) => window.__doc.wb.getActiveSheet().getRange(r, c).getCellData(), [where.r, where.c]);
  assert.equal(after.v, 'retyped');
  assert.ok(!after.p, 'the old runs are gone: ' + JSON.stringify(after));
});

await t.finish();
