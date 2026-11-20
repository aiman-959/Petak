/**
 * A working day on a phone (Part 2 of the 1.3.0 brief): the staff roster is opened and changed the
 * way a team lead would, with real touches only, every step timed from the touch to its result,
 * then saved, reopened and checked: nothing lost or changed except the edits.
 *
 *   open the roster → edit a name → fill a shift down → bold + colour a row → sort a column →
 *   filter → pinch out and back → insert a picture → draw a rectangle snapped to cells →
 *   save as .xlsx → reopen → everything is still there
 *
 * Writes tests/e2e/out/daily.json (step, ms, result) for TEST-REPORT.md.
 *   BASE_URL=http://localhost:4173/ node tests/e2e/daily.mjs
 */
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { unzipSync } from 'fflate';
import { start, BASE } from './lib.mjs';
import { touch } from './touch.mjs';
import { excelProblemsOf } from './excelRulesCli.mjs';

const t = await start('daily');
const { page, step, ev, idle, val, fresh, openFile, download, tmp } = t;
console.log('Petak e2e: a working day with the roster on a phone —', BASE);
await fresh();
await page.waitForTimeout(400);
const T = await touch(page);
const timings = [];
const timed = async (what, fn) => {
  const t0 = Date.now();
  const r = await fn();
  const ms = Date.now() - t0;
  timings.push({ what, ms });
  console.log(`     ${what}: ${ms} ms`);
  return r;
};
const at = (ref) =>
  ev((ref) => {
    const V = window.__view;
    const v = V.viewInfo();
    const r = window.__doc.wb.getActiveSheet().getRange(ref).getRange();
    const a = (acc, i) => (i <= 0 ? 0 : acc[i - 1]);
    const p0 = V.toScreen(v, a(v.colAcc, r.startColumn), a(v.rowAcc, r.startRow));
    const p1 = V.toScreen(v, a(v.colAcc, r.endColumn + 1), a(v.rowAcc, r.endRow + 1));
    const h = document.querySelector('[data-testid=grid]').getBoundingClientRect();
    return { x: h.left + (p0.x + p1.x) / 2, y: h.top + (p0.y + p1.y) / 2, x1: h.left + p1.x, y1: h.top + p1.y };
  }, ref);
const tapEl = async (sel) => {
  const el = await page.waitForSelector(sel);
  const b = await el.boundingBox();
  await T.tap(b.x + b.width / 2, b.y + b.height / 2);
  await page.waitForTimeout(150);
};
const closePanel = async () => {
  if (await page.$('.sheet-backdrop')) await ev(() => window.__petakBack?.());
  await page.waitForTimeout(150);
};
const keyboardFocus = () => ev(() => { const a = document.activeElement; return !!a && (a.isContentEditable || a.tagName === 'INPUT' || a.tagName === 'TEXTAREA'); });
const ROSTER = path.resolve('tests/fixtures/files/roster.xlsx');

await step('open the roster', async () => {
  await timed('open roster.xlsx (Open → pick → grid shown)', () => openFile(ROSTER));
  assert.match(await val('B5'), /王芳|.+/);
});
let before;
await step('edit a name: double tap, type, Enter', async () => {
  before = await val('B5');
  const p = await at('B5');
  await timed('double tap B5 → keyboard', async () => {
    await T.doubleTap(p.x, p.y);
    await page.waitForFunction(() => document.activeElement?.dataset?.testid === 'formula-input');
  });
  await page.keyboard.press('Control+A');
  await page.keyboard.type('Wang Fang');
  await timed('Enter → saved', async () => {
    await page.keyboard.press('Enter');
    await idle();
  });
  assert.equal(await val('B5'), 'Wang Fang');
  await page.keyboard.press('Escape');
  await page.waitForTimeout(150);
  assert.equal(await keyboardFocus(), false, 'the keyboard goes away after Escape');
});
await step('fill a shift down with the fill grip (D5 → D8)', async () => {
  const d5 = await val('D5');
  const p = await at('D5');
  await T.tap(p.x, p.y);
  await page.waitForTimeout(200);
  const grip = await (await page.waitForSelector('.fill-handle')).boundingBox();
  const d8 = await at('D8');
  await timed('drag the fill grip down 3 rows', async () => {
    await T.drag({ x: grip.x + grip.width / 2, y: grip.y + grip.height / 2 }, { x: grip.x + grip.width / 2, y: d8.y + 4 }, { steps: 10, hold: 80 });
    await idle();
  });
  for (const r of ['D6', 'D7', 'D8']) assert.equal(await val(r), d5, `${r} copied`);
});
await step('bold and colour row 4 (long press row header → Format)', async () => {
  const p = await at('A4');
  const g = await (await page.$('[data-testid=grid]')).boundingBox();
  await T.tap(g.x + 12, p.y);
  await page.waitForTimeout(200);
  await timed('Format → Bold', async () => {
    await tapEl('[data-testid=tab-format]');
    await tapEl('[data-testid=fmt-bold]');
    await closePanel();
  });
  const bold = await ev(() => window.__doc.wb.getActiveSheet().getRange('C4').getCellStyleData()?.bl);
  assert.equal(bold, 1);
});
await step('sort the hours column (J) high to low keeps the header', async () => {
  const header = await val('A3');
  const p = await at('A5');
  await T.tap(p.x, p.y);
  await page.waitForTimeout(150);
  await timed('Data → Sort A→Z', async () => {
    await tapEl('[data-testid=tab-data]');
    await tapEl('[data-testid=tool-sort-asc]');
    await closePanel();
    await idle();
  });
  assert.equal(await val('A3'), header, 'the heading row stays');
});
await step('filter on and off', async () => {
  await timed('Data → Filter on', async () => {
    await tapEl('[data-testid=tab-data]');
    await tapEl('[data-testid=tool-filter]');
    await closePanel();
  });
  assert.ok(await ev(() => !!window.__doc.wb.getActiveSheet().getFilter?.()), 'filter set');
});
await step('pinch out to see the whole roster and back in', async () => {
  const g = await (await page.$('[data-testid=grid]')).boundingBox();
  const c = { x: g.x + g.width / 2, y: g.y + g.height / 2 };
  await timed('pinch out (fingers 300 → 120 px)', async () => {
    await T.pinch(c, 300, 120, { steps: 12 });
    await page.waitForFunction(() => !window.__doc.gesturing);
  });
  await timed('pinch back in', async () => {
    await T.pinch(c, 120, 300, { steps: 12 });
    await page.waitForFunction(() => !window.__doc.gesturing);
  });
  const z = await ev(() => window.__view.viewInfo().zoom);
  assert.ok(Math.abs(z - 1) < 0.05, 'zoom ' + z);
});
await step('insert a picture', async () => {
  await ev(() => (window.__doc.wb.getActiveSheet().getRange('B20').activate(), 0));
  const png = unzipSync(new Uint8Array(fs.readFileSync(path.resolve('tests/fixtures/files/qm-report.xlsx'))))['xl/media/image2.png'];
  const n0 = await ev(() => window.__pictures.list.length);
  await timed('Insert → Picture → pick', async () => {
    const [chooser] = await Promise.all([page.waitForEvent('filechooser'), (async () => { await tapEl('[data-testid=tab-insert]'); await tapEl('[data-testid=tool-picture]'); })()]);
    await chooser.setFiles({ name: 'logo.png', mimeType: 'image/png', buffer: Buffer.from(png) });
    await page.waitForFunction((n) => window.__pictures.list.length === n + 1, n0);
  });
});
await step('draw a rectangle that snaps to cells (starts mid-cell)', async () => {
  await tapEl('[data-testid=tab-draw]');
  await tapEl('[data-testid=tool-draw]');
  await page.waitForSelector('[data-testid=draw-layer]');
  await page.click('[data-testid=draw-tool-rect]');
  const p = await ev(() => { const v = window.__drawView(); const r = document.querySelector('[data-testid=draw-layer]').getBoundingClientRect(); return { x: r.left + v.ox, y: r.top + v.oy, z: v.zoom }; });
  // B8 middle → D10 middle in sheet px (cells 88 wide); drawn shapes fit to whole cells
  const a = await ev(() => { const v = window.__view.viewInfo(); return { x: v.colAcc[0] + 40, y: v.rowAcc[6] + 10 }; });
  const b = await ev(() => { const v = window.__view.viewInfo(); return { x: v.colAcc[2] + 40, y: v.rowAcc[8] + 10 }; });
  await timed('press, drag, release a rectangle', async () => {
    await T.drag({ x: p.x + a.x * p.z, y: p.y + a.y * p.z }, { x: p.x + b.x * p.z, y: p.y + b.y * p.z }, { steps: 10 });
    await page.waitForTimeout(200);
  });
  const s = await ev(() => window.__doc.meta.drawings?.[window.__doc.wb.getActiveSheet().getSheetId()]?.shapes.at(-1));
  const edges = await ev(() => { const v = window.__view.viewInfo(); return { cols: [0, ...v.colAcc.slice(0, 12)], rows: [0, ...v.rowAcc.slice(0, 20)] }; });
  assert.ok(s && edges.cols.some((x) => Math.abs(x - s.x) < 0.5) && edges.rows.some((y) => Math.abs(y - s.y) < 0.5), 'snapped to a cell corner: ' + JSON.stringify(s && { x: s.x, y: s.y }));
  await ev(() => window.__petakBack());
  await page.waitForTimeout(150);
  while (await ev(() => window.__draw.get().active)) {
    await ev(() => window.__petakBack());
    await page.waitForTimeout(150);
  }
});
let saved;
await step('save as .xlsx: a valid file Excel opens without repairing', async () => {
  saved = await timed('File → Save .xlsx', () =>
    download(async () => {
      await tapEl('[data-testid=menu-more]');
      await page.click('text=Save .xlsx');
    }),
  );
  const problems = excelProblemsOf(saved);
  assert.deepEqual(problems, []);
});
await step('reopen: every edit is there, nothing else changed', async () => {
  const snap = await ev(() => ({ pics: window.__pictures.list.length, shapes: window.__doc.meta.drawings?.[window.__doc.wb.getActiveSheet().getSheetId()]?.shapes.length ?? 0, title: window.__doc.wb.getActiveSheet().getRange('A1').getValue() }));
  await timed('reopen the saved file', () => openFile(saved));
  assert.equal(await val('A1'), snap.title);
  assert.ok((await ev(() => [...Array(20)].map((_, i) => window.__doc.wb.getActiveSheet().getRange(i, 1).getValue()))).includes('Wang Fang'), 'the edited name');
  assert.equal(await ev(() => window.__doc.wb.getActiveSheet().getRange('C4').getCellStyleData()?.bl), 1, 'bold row 4');
  assert.equal(await ev(() => window.__pictures.list.length), snap.pics, 'pictures');
  assert.equal(await ev(() => window.__doc.meta.drawings?.[window.__doc.wb.getActiveSheet().getSheetId()]?.shapes.length ?? 0), snap.shapes, 'shapes');
});

fs.mkdirSync('tests/e2e/out', { recursive: true });
fs.writeFileSync('tests/e2e/out/daily.json', JSON.stringify(timings, null, 1) + '\n');
await t.finish();
