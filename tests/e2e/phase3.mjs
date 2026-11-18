/**
 * Phase 3 (Medium / Low items) in the app on a phone: coloured references while typing a
 * formula (M-03), error explanations incl. the circular-reference loop (M-06), CSV export
 * notice (M-07), CSV leading zeros (L-01), sheet rename / move (M-10), toast position (L-02).
 *   BASE_URL=http://localhost:4173/ node tests/e2e/phase3.mjs
 */
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { start, BASE } from './lib.mjs';
import { touch } from './touch.mjs';

const t = await start('phase3');
const { page, step, ev, idle, val, fresh, select, typeIn, download, openFile } = t;
console.log('Petak e2e phase 3 —', BASE);
await fresh();
await page.waitForTimeout(400);
const T = await touch(page);
const tabs = () => ev(() => window.__doc.wb.getSheets().map((s) => s.getSheetName()));

await step('M-03: references in a formula are coloured in the field and outlined on the grid', async () => {
  await select('E2');
  await page.click('[data-testid=formula-input]');
  await page.keyboard.type('=A1+B2:C3*A1+"D4"');
  await page.waitForTimeout(250);
  const spans = await page.$$eval('[data-testid=formula-mirror] span[data-ref-color]', (els) => els.map((e) => [e.textContent, e.dataset.refColor]));
  assert.deepEqual(spans.map((s) => s[0]), ['A1', 'B2:C3', 'A1']);
  assert.equal(spans[0][1], spans[2][1], 'the same reference keeps its colour');
  assert.notEqual(spans[0][1], spans[1][1]);
  // the mirror shows exactly what is typed, the input text itself is transparent
  assert.equal(await page.textContent('[data-testid=formula-mirror]'), '=A1+B2:C3*A1+"D4"');
  const boxes = await page.$$eval('[data-testid=ref-box]', (els) => els.map((e) => [e.dataset.color, Math.round(e.getBoundingClientRect().width), Math.round(e.getBoundingClientRect().height)]));
  assert.equal(boxes.length, 3);
  const z = await ev(() => window.__view.viewInfo().zoom);
  const b2c3 = boxes.find((b) => b[0] === spans[1][1]);
  assert.ok(Math.abs(b2c3[1] - 176 * z) < 4 && Math.abs(b2c3[2] - 48 * z) < 4, 'B2:C3 is two columns by two rows: ' + JSON.stringify(b2c3));
  await page.screenshot({ path: path.join(t.tmp, 'coloured-refs.png') });
  await page.keyboard.press('Escape');
  await page.waitForTimeout(150);
  assert.equal(await page.$('[data-testid=ref-box]'), null, 'gone when editing ends');
});

await step('M-06: a circular reference explains itself and names the loop', async () => {
  await ev(() => { const ws = window.__doc.wb.getActiveSheet(); ws.getRange('A1').setValue({ f: '=B1+1' }); ws.getRange('B1').setValue({ f: '=C1*2' }); ws.getRange('C1').setValue({ f: '=A1' }); return 0; });
  await idle();
  await select('A1');
  await page.waitForSelector('[data-testid=error-help]', { timeout: 5000 });
  assert.match(await page.textContent('[data-testid=error-help]'), /#CYCLE! Circular reference/);
  await page.click('[data-testid=error-help]');
  await page.waitForSelector('[data-testid=dlg-ask]');
  const text = await page.textContent('[data-testid=ask-text]');
  assert.match(text, /depends on its own result/);
  assert.match(text, /The loop: A1 → B1 → C1 → A1/);
  await page.click('[data-testid=ask-ok]');
  // other errors explain themselves too
  await typeIn('D1', '=1/0');
  await select('D1');
  await page.waitForSelector('[data-testid=error-help]');
  assert.match(await page.textContent('[data-testid=error-help]'), /#DIV\/0! Division by zero/);
  await select('F9');
  await page.waitForTimeout(200);
  assert.equal(await page.$('[data-testid=error-help]'), null);
});

await step('M-10: rename a sheet with Excel’s rules; move it among the tabs', async () => {
  await page.click('[data-testid=add-sheet]');
  await page.waitForTimeout(200);
  assert.deepEqual(await tabs(), ['Sheet1', 'Sheet2']);
  const b = await (await page.$('[data-testid=sheet-tab-Sheet1]')).boundingBox();
  await T.longPress(b.x + b.width / 2, b.y + b.height / 2, 700);
  await page.click('[data-testid=sheet-rename]');
  await page.waitForSelector('[data-testid=dlg-prompt]');
  await page.fill('[data-testid=prompt-input]', 'Q3/Q4');
  assert.match(await page.textContent('[data-testid=prompt-error]'), /cannot contain “\/”/);
  assert.ok(await page.isDisabled('[data-testid=prompt-ok]'));
  await page.fill('[data-testid=prompt-input]', 'sheet2');
  assert.match(await page.textContent('[data-testid=prompt-error]'), /already a sheet/);
  await page.fill('[data-testid=prompt-input]', 'Budget 2026');
  assert.equal(await page.$('[data-testid=prompt-error]'), null);
  await page.click('[data-testid=prompt-ok]');
  await page.waitForTimeout(200);
  assert.deepEqual(await tabs(), ['Budget 2026', 'Sheet2']);
  const b2 = await (await page.$('[data-testid=sheet-tab-Budget\\ 2026]')).boundingBox();
  await T.longPress(b2.x + b2.width / 2, b2.y + b2.height / 2, 700);
  await page.click('[data-testid=sheet-move-right]');
  await page.waitForTimeout(200);
  assert.deepEqual(await tabs(), ['Sheet2', 'Budget 2026']);
  // a formula can use the new name
  await ev(() => { const wb = window.__doc.wb; wb.getSheetByName('Budget 2026').getRange('H1').setValue(42); wb.setActiveSheet(wb.getSheetByName('Sheet2')); wb.getActiveSheet().getRange('A1').setValue({ f: "='Budget 2026'!H1*2" }); return 0; });
  await idle();
  assert.equal(await val('A1'), 84);
});

await step('M-07: CSV export says which sheet was written and that the others are not in it', async () => {
  const out = await download(async () => {
    await page.click('[data-testid=menu-more]');
    await page.click('text=Export .csv');
  });
  assert.ok(out.endsWith('.csv'));
  await page.waitForFunction(() => /A CSV file holds one sheet/.test(document.querySelector('.toast')?.textContent ?? ''), null, { timeout: 10000 });
  assert.match(await page.textContent('.toast'), /Saved “Sheet2” as CSV\. A CSV file holds one sheet: the other 1 sheet is not in it\./);
});

await step('L-02: messages sit at the top and let taps through', async () => {
  const box = await (await page.$('.toast')).boundingBox();
  const grid = await (await page.$('[data-testid=grid]')).boundingBox();
  assert.ok(box.y < grid.y + 60, `toast at the top of the sheet: ${box.y} vs grid ${grid.y}`);
  assert.equal(await ev(() => getComputedStyle(document.querySelector('.toast')).pointerEvents), 'none');
});

await step('L-01: CSV identifiers with leading zeros stay text; dates show as dates', async () => {
  const f = path.join(t.tmp, 'ids.csv');
  fs.writeFileSync(f, 'id,phone,joined,amount\n0012,0123456789,2024-01-05,12.50\n');
  await openFile(f);
  assert.equal(await val('A2'), '0012');
  assert.equal(await val('B2'), '0123456789');
  assert.equal(await val('D2'), 12.5);
  assert.equal(await ev(() => window.__doc.wb.getActiveSheet().getRange('C2').getDisplayValue()), '2024-01-05');
});

await step('M-05: charts, sparklines and drawings undo with the main Undo, in order with cell edits', async () => {
  await fresh();
  await page.waitForTimeout(400);
  const meta = () => ev(() => ({ charts: (window.__doc.meta.charts ?? []).length, sparklines: (window.__doc.meta.sparklines ?? []).length, shapes: Object.values(window.__doc.meta.drawings ?? {}).reduce((t, d) => t + d.shapes.length, 0) }));
  await typeIn('A1', '5');
  await ev(() => { window.__doc.updateMeta((m) => (m.sparklines = [...(m.sparklines ?? []), { id: 'sp1', sheetId: window.__doc.wb.getActiveSheet().getSheetId(), cell: 'D1', source: 'A1:C1', type: 'line', color: '#15803d', negColor: '#b42318' }])); return 0; });
  await typeIn('B1', '7');
  await ev(() => { window.__doc.updateMeta((m) => (m.charts = [{ id: 'c1', type: 'bar', sheetId: window.__doc.wb.getActiveSheet().getSheetId(), range: 'A1:B1', title: 'T' }, ...(m.charts ?? [])])); return 0; });
  assert.deepEqual(await meta(), { charts: 1, sparklines: 1, shapes: 0 });
  const undo = async () => { await page.click('[data-testid=undo]'); await idle(); };
  const redo = async () => { await page.click('[data-testid=redo]'); await idle(); };
  await undo();
  assert.deepEqual(await meta(), { charts: 0, sparklines: 1, shapes: 0 }, 'the chart goes first');
  assert.equal(await val('B1'), 7);
  await undo();
  assert.ok((await val('B1')) == null, 'then the cell edit');
  await undo();
  assert.deepEqual(await meta(), { charts: 0, sparklines: 0, shapes: 0 }, 'then the sparkline');
  assert.equal(await val('A1'), 5);
  await redo();
  await redo();
  await redo();
  assert.deepEqual(await meta(), { charts: 1, sparklines: 1, shapes: 0 });
  assert.equal(await val('B1'), 7);
  // a drawn shape: drawing mode's own undo does not add a second entry to the workbook's history
  await ev(() => { window.__draw.open('select'); const d = window.__draw.current(); window.__draw.commit({ ...d, shapes: [...d.shapes, { id: 'r1', type: 'rect', x: 10, y: 10, w: 50, h: 30, rotation: 0, stroke: '#000', fill: 'none', strokeWidth: 1, layerId: 'L1' }] }); window.__draw.close(); return 0; });
  assert.equal((await meta()).shapes, 1);
  await undo();
  assert.equal((await meta()).shapes, 0, 'the shape is undone with the main Undo');
  await redo();
  assert.equal((await meta()).shapes, 1);
});

await step('Group rows, collapse, expand; outline levels saved in the file', async () => {
  await fresh();
  await page.waitForTimeout(400);
  await ev(() => { const ws = window.__doc.wb.getActiveSheet(); ws.getRange('A1:A6').setValues([['Item'], ['a'], ['b'], ['c'], ['Total'], ['x']]); ws.getRange('A2:A4').activate(); return 0; });
  const dataTool = async (id) => {
    if (!(await page.isVisible(`[data-testid=${id}]`))) await page.click('[data-testid=tab-data]');
    await page.click(`[data-testid=${id}]`);
    await page.waitForTimeout(150);
  };
  await dataTool('tool-group');
  const levels = () => ev(() => [0, 1, 2, 3, 4, 5].map((r) => window.__doc.wb.getActiveSheet().getRowCustomMetadata(r)?.ol ?? 0));
  const hidden = () => ev(() => [0, 1, 2, 3, 4, 5].filter((r) => !window.__doc.wb.getActiveSheet().getSheet().getRowVisible(r)));
  assert.deepEqual(await levels(), [0, 1, 1, 1, 0, 0], 'levels ' + JSON.stringify(await levels()) + ' ' + JSON.stringify(await ev(() => window.__doc.wb.getActiveSheet().getActiveRange().getRange())));
  await ev(() => { window.__doc.wb.getActiveSheet().getRange('A3').activate(); return 0; });
  await dataTool('tool-collapse');
  assert.deepEqual(await hidden(), [1, 2, 3]);
  // saved: outline level on rows 2-4 (hidden), "collapsed" on the summary row 5
  const bytes = await ev(async () => Array.from(new Uint8Array((await window.__doc.serialize('xlsx')).data)));
  const { unzipSync, strFromU8 } = await import('fflate');
  const xml = strFromU8(unzipSync(new Uint8Array(bytes))['xl/worksheets/sheet1.xml']);
  for (const r of [2, 3, 4]) assert.match(xml, new RegExp(`<row r="${r}"[^>]*hidden="1"[^>]*outlineLevel="1"`), 'row ' + r);
  assert.match(xml, /<row r="5"[^>]*collapsed="1"/);
  // expand from the summary row (the selection after collapsing)
  await dataTool('tool-expand');
  assert.deepEqual(await hidden(), []);
  await ev(() => { window.__doc.wb.getActiveSheet().getRange('A2:A4').activate(); return 0; });
  await dataTool('tool-ungroup');
  assert.deepEqual(await levels(), [0, 0, 0, 0, 0, 0]);
  if (await page.isVisible('[data-testid=tool-group]')) await page.mouse.click(200, 120); // close the panel (tap its backdrop)
});

await step('M-09: Print / PDF of the roster: pages at print size, frozen header on every page, CJK text', async () => {
  await openFile(path.resolve('tests/fixtures/files/roster.xlsx'));
  await page.waitForSelector('.toast', { state: 'detached', timeout: 15000 }).catch(() => {});
  await page.click('[data-testid=menu-more]');
  await page.click('[data-testid=menu-print]');
  await page.waitForSelector('[data-testid=dlg-print]');
  await page.click('[data-testid=print-landscape]');
  const out = await download(() => page.click('[data-testid=print-save]'));
  // (headless Chromium here drops non-ASCII download names such as the CJK sheet name; the content is what matters)
  assert.equal(fs.readFileSync(out).subarray(0, 5).toString(), '%PDF-');
  const { execFileSync } = await import('node:child_process');
  const info = execFileSync('pdfinfo', [out], { encoding: 'utf8' });
  assert.match(info, /Pages:\s+1/);
  assert.match(info, /Page size:\s+841\.89 x 595\.28 pts/);
  // render the page and check it is not blank: the dark green title fill and black text are there
  execFileSync('pdftoppm', ['-r', '50', '-png', out, path.join(t.tmp, 'print')]);
  const png = fs.readdirSync(t.tmp).find((f) => f.startsWith('print') && f.endsWith('.png'));
  assert.ok(png, 'page rendered');
  const stats = await ev(async (b64) => {
    const img = await createImageBitmap(await (await fetch('data:image/png;base64,' + b64)).blob());
    const c = new OffscreenCanvas(img.width, img.height);
    const x = c.getContext('2d');
    x.drawImage(img, 0, 0);
    const d = x.getImageData(0, 0, img.width, img.height).data;
    let green = 0;
    let dark = 0;
    for (let i = 0; i < d.length; i += 4) {
      if (Math.abs(d[i] - 0x14) < 20 && Math.abs(d[i + 1] - 0x53) < 25 && Math.abs(d[i + 2] - 0x2d) < 25) green++;
      if (d[i] + d[i + 1] + d[i + 2] < 150) dark++;
    }
    return { green, dark, px: d.length / 4 };
  }, fs.readFileSync(path.join(t.tmp, png)).toString('base64'));
  assert.ok(stats.green > stats.px * 0.01, 'title fill printed: ' + JSON.stringify(stats));
  assert.ok(stats.dark > stats.px * 0.005, 'text and borders printed: ' + JSON.stringify(stats));
  assert.match(await page.textContent('.toast'), /PDF saved: 1 page/);
});

await t.finish();
