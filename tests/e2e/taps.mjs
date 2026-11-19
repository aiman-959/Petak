/**
 * Phase 4: how many touches ten everyday tasks take on a phone. Each task is done with real
 * touches only (no shortcuts through code); every touch and every key typed is counted, and the
 * result is checked. Writes tests/e2e/out/taps.json; TEST-REPORT.md quotes it.
 *   BASE_URL=http://localhost:4173/ node tests/e2e/taps.mjs
 * A touch is a tap, a long press or one drag. Closing a tool panel to get back to the sheet
 * counts as a touch.
 */
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { start, BASE } from './lib.mjs';
import { touch } from './touch.mjs';

const t = await start('taps');
const { page, step, ev, idle, val, fresh } = t;
console.log('Petak e2e: touches per everyday task —', BASE);
await fresh();
await page.waitForTimeout(400);
const T = await touch(page);

let touches = 0;
let keys = 0;
let trail = [];
const note = (s) => trail.push(s);
const g = async () => (await page.$('[data-testid=grid]')).boundingBox();
const at = (ref) =>
  ev((ref) => {
    const V = window.__view;
    const v = V.viewInfo();
    const r = window.__doc.wb.getActiveSheet().getRange(ref).getRange();
    const a = (acc, i) => (i <= 0 ? 0 : acc[i - 1]);
    const p0 = V.toScreen(v, a(v.colAcc, r.startColumn), a(v.rowAcc, r.startRow));
    const p1 = V.toScreen(v, a(v.colAcc, r.endColumn + 1), a(v.rowAcc, r.endRow + 1));
    const h = document.querySelector('[data-testid=grid]').getBoundingClientRect();
    return { x: h.left + (p0.x + p1.x) / 2, y: h.top + (p0.y + p1.y) / 2, x1: h.left + p1.x, y1: h.top + p1.y, y0: h.top + p0.y };
  }, ref);
const tapAt = async (x, y, what) => {
  await T.tap(x, y);
  touches++;
  note(what);
  await page.waitForTimeout(150);
};
const tapCell = async (ref) => {
  const p = await at(ref);
  await tapAt(p.x, p.y, `tap ${ref}`);
};
const tapEl = async (sel, what) => {
  const el = await page.waitForSelector(sel);
  const b = await el.boundingBox();
  await tapAt(b.x + b.width / 2, b.y + b.height / 2, what);
};
const doubleTap = async (ref) => {
  const p = await at(ref);
  await T.doubleTap(p.x, p.y);
  touches += 2;
  note(`double tap ${ref}`);
  await page.waitForTimeout(200);
};
const longPress = async (x, y, what) => {
  await T.longPress(x, y);
  touches++;
  note(`long press ${what}`);
  await page.waitForTimeout(150);
};
const type = async (text) => {
  await page.keyboard.type(text);
  keys += text.length;
  note(`type "${text}"`);
};
const enter = async () => {
  await page.keyboard.press('Enter');
  keys++;
  note('Enter');
  await idle();
};
const closePanel = async () => {
  if (await page.$('.sheet-backdrop')) await tapAt(200, 120, 'close the panel');
};
const results = [];
const task = (name, setup, fn, check) =>
  step(name, async () => {
    await fresh();
    await page.waitForTimeout(300);
    if (setup) await setup();
    await idle();
    touches = 0;
    keys = 0;
    trail = [];
    await fn();
    await idle();
    await check();
    results.push({ task: name, touches, keys, how: trail.join(' → ') });
    console.log(`     ${touches} touches${keys ? `, ${keys} keys` : ''}: ${trail.join(' → ')}`);
  });
const data = () =>
  ev(() => {
    const ws = window.__doc.wb.getActiveSheet();
    ws.getRange('A1:C6').setValues([
      ['Item', 'Qty', 'Price'],
      ['pear', 3, 1.2],
      ['apple', 5, 0.8],
      ['fig', 2, 2.5],
      ['kiwi', 7, 0.6],
      ['date', 4, 3.1],
    ]);
    ws.getRange('A1').activate();
    return 0;
  });

await task(
  '1. Open a file from the phone',
  null,
  async () => {
    const [chooser] = await Promise.all([page.waitForEvent('filechooser'), tapEl('[data-testid=open]', 'Open')]);
    note('pick the file in Android’s picker');
    await chooser.setFiles(path.resolve('tests/fixtures/files/roster.xlsx'));
    await page.waitForFunction(() => window.__doc.snapshot.file.name === 'roster.xlsx', null, { timeout: 60000 });
  },
  async () => assert.equal(await ev(() => window.__doc.snapshot.file.name), 'roster.xlsx'),
);

await task(
  '2. Type a number into a cell',
  null,
  async () => {
    await doubleTap('B2');
    await type('42');
    await enter();
  },
  async () => assert.equal(await val('B2'), 42),
);

await task(
  '3. Add up a column with a SUM formula',
  data,
  async () => {
    await doubleTap('B7');
    await type('=SUM(');
    await tapCell('B2');
    await type(':');
    await tapCell('B6');
    await type(')');
    await enter();
  },
  async () => {
    assert.equal(await ev(() => window.__doc.wb.getActiveSheet().getRange('B7').getFormula()), '=SUM(B2:B6)');
    assert.equal(await val('B7'), 21);
  },
);

await task(
  '4. See the total of a column without a formula',
  data,
  async () => {
    const p = await at('B1');
    await tapAt(p.x, p.y0 - 10, 'tap column header B');
  },
  async () => assert.match(await page.textContent('[data-testid=status-bar]'), /Sum\s*21/),
);

await task(
  '5. Make the header row bold',
  data,
  async () => {
    const p = await at('A1');
    await tapAt((await g()).x + 14, p.y, 'tap row header 1');
    await tapEl('[data-testid=tab-format]', 'Format');
    await tapEl('.bsheet [data-testid=fmt-bold]', 'Bold');
    await closePanel();
  },
  async () => assert.equal(await ev(() => window.__doc.wb.getActiveSheet().getRange('C1').getCellStyleData()?.bl), 1),
);

await task(
  '6. Colour a cell',
  data,
  async () => {
    await tapCell('C3');
    await tapEl('[data-testid=tab-format]', 'Format');
    await tapEl('.bsheet [aria-label="Fill #fef3c7"]', 'a fill colour');
    await closePanel();
  },
  async () => assert.equal(await ev(() => window.__doc.wb.getActiveSheet().getRange('C3').getCellStyleData()?.bg?.rgb?.toLowerCase()), '#fef3c7'),
);

await task(
  '7. Insert a row',
  data,
  async () => {
    const p = await at('A3');
    await longPress((await g()).x + 14, p.y, 'row header 3');
    await tapEl('[data-testid=ctx-row-above]', 'Row above');
  },
  async () => {
    assert.equal(await val('A3'), null);
    assert.equal(await val('A4'), 'apple');
  },
);

await task(
  '8. Sort a list by a column',
  data,
  async () => {
    await tapCell('A3');
    await tapEl('[data-testid=tab-data]', 'Data');
    await tapEl('.bsheet [data-testid=tool-sort-asc]', 'Sort A→Z');
    await closePanel();
  },
  async () => { const got = await ev(() => window.__doc.wb.getActiveSheet().getRange('A1:A6').getValues().flat()); assert.deepEqual(got, ['Item', 'apple', 'date', 'fig', 'kiwi', 'pear'], JSON.stringify(got)); },
);

await task(
  '9. Keep the header row in view (freeze)',
  data,
  async () => {
    await tapEl('[data-testid=tab-data]', 'Data');
    await tapEl('.bsheet [data-testid=tool-freeze-row]', 'Top row');
    await closePanel();
  },
  async () => assert.equal(await ev(() => window.__doc.wb.getActiveSheet().getFrozenRows()), 1),
);

await task(
  '10. Make a bar chart of a table',
  data,
  async () => {
    await tapCell('A1');
    // drag the selection's corner handle to C6
    const a = await at('A1');
    const c = await at('C6');
    await T.drag({ x: a.x1, y: a.y1 }, { x: c.x - 10, y: c.y }, { steps: 14, hold: 80, settle: 100 });
    touches++;
    note('drag the selection handle to C6');
    await page.waitForTimeout(200);
    await tapEl('[data-testid=tab-insert]', 'Insert');
    await tapEl('.bsheet [data-testid=tool-chart-bar]', 'Bar chart');
    await tapEl('[data-testid=chart-add]', 'Add');
  },
  async () => {
    const ch = await ev(() => window.__doc.meta.charts.map((c) => [c.type, c.range ?? c.source]));
    assert.equal(ch.length, 1);
    assert.equal(ch[0][0], 'bar');
    assert.match(JSON.stringify(ch[0][1]), /A1:C6/);
  },
);

await task(
  'Copy a cell to another cell',
  data,
  async () => {
    const p = await at('C2');
    await longPress(p.x, p.y, 'C2');
    await tapEl('[data-testid=ctx-copy]', 'Copy');
    const q = await at('D2');
    await longPress(q.x, q.y, 'D2');
    await tapEl('[data-testid=ctx-paste]', 'Paste');
  },
  async () => assert.equal(await val('D2'), 1.2),
);

await task(
  'Share the workbook',
  data,
  async () => {
    const [dl] = await Promise.all([page.waitForEvent('download'), (async () => {
      await tapEl('[data-testid=menu-more]', '⋮');
      await tapEl('.bsheet button.tool:has-text("Share")', 'Share');
    })()]);
    note('pick the app in Android’s share sheet');
    assert.match(dl.suggestedFilename(), /\.xlsx$/);
  },
  async () => {},
);

fs.mkdirSync('tests/e2e/out', { recursive: true });
fs.writeFileSync('tests/e2e/out/taps.json', JSON.stringify(results, null, 1) + '\n');
await t.finish();
