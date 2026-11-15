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
const toast = async () => (await page.waitForSelector('.toast', { timeout: 15000 })).textContent();
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
  const msg = await toast();
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
  assert.match(await toast(), /Could not open “empty\.xlsx”: the file is empty/);
  assert.equal(await name(), 'people.csv');
  await page.waitForSelector('.toast', { state: 'detached', timeout: 15000 });
});

await step('recent files: listed newest first; one that is gone is explained and removed', async () => {
  await ev(async () => { await window.__recent.clear(); await window.__recent.add({ name: 'old.xlsx', uri: 'content://x/old' }); await window.__recent.add({ name: 'gone.xlsx', uri: 'content://x/gone' }); await window.__recent.add({ name: 'old.xlsx', uri: 'content://x/old' }); return 0; });
  await page.click('[data-testid=menu-more]');
  await page.click('[data-testid=menu-recent]');
  await page.waitForSelector('[data-testid=dlg-recent]');
  const items = await page.$$eval('[data-testid=recent-item] b', (els) => els.map((e) => e.textContent));
  assert.deepEqual(items, ['old.xlsx', 'gone.xlsx']);
  await page.click('[data-testid=recent-item] >> nth=1');
  const msg = await toast();
  assert.match(msg, /“gone\.xlsx” could not be opened again/);
  assert.deepEqual(await ev(() => window.__recent.list().then((l) => l.map((x) => x.name))), ['old.xlsx']);
  assert.equal(await name(), 'people.csv');
});

await t.finish();
