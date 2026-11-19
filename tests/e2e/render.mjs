/**
 * Rendering (K-08, K-11): cell borders stay visible on screen at every zoom, also after
 * scrolling; frozen panes do not paint over each other.
 * Measured on the screenshot (what the user sees), not on the canvas: on this profile the grid
 * canvas is stretched 2 → 2.625 screen pixels.
 *   BASE_URL=http://localhost:4173/ node tests/e2e/render.mjs
 */
import assert from 'node:assert/strict';
import path from 'node:path';
import { start, BASE } from './lib.mjs';
import { touch } from './touch.mjs';

const t = await start('render');
const { page, step, ev, fresh, openFile } = t;
console.log('Petak e2e rendering —', BASE);
await fresh();
await page.waitForTimeout(300);
const T = await touch(page);
const gestureEnd = () => page.waitForFunction(() => !window.__doc.gesturing, null, { timeout: 10000 }).then(() => page.waitForTimeout(500));

/** Darkest pixel (r+g+b, 0–765) of small windows across the borders of three cells in the scrolling pane. */
const borders = async () => {
  const shot = (await page.screenshot()).toString('base64');
  return ev(async (b64) => {
    const img = await createImageBitmap(await (await fetch('data:image/png;base64,' + b64)).blob());
    const c = new OffscreenCanvas(img.width, img.height);
    const ctx = c.getContext('2d');
    ctx.drawImage(img, 0, 0);
    const k = img.width / innerWidth;
    const V = window.__view;
    const v = V.viewInfo();
    const host = document.querySelector('[data-testid=grid]').getBoundingClientRect();
    const a = (acc, i) => (i <= 0 ? 0 : acc[i - 1]);
    const o = V.mainOrigin(v);
    const c0 = v.colAcc.findIndex((e) => e > o.x + 30) + 1;
    const r0 = v.rowAcc.findIndex((e) => e > o.y + 12) + 1;
    const min = (x, y, w, h) => {
      const d = ctx.getImageData(Math.round((host.left + x) * k), Math.round((host.top + y) * k), Math.max(1, Math.round(w * k)), Math.max(1, Math.round(h * k))).data;
      let m = 765;
      for (let i = 0; i < d.length; i += 4) m = Math.min(m, d[i] + d[i + 1] + d[i + 2]);
      return m;
    };
    const out = { vertical: [], horizontal: [], cell: [r0 + 1, c0 + 1], zoom: v.zoom };
    for (let i = 0; i < 3; i++) {
      const p = V.toScreen(v, a(v.colAcc, c0 + i), a(v.rowAcc, r0 + i));
      out.vertical.push(min(p.x - 2, p.y + 3, 4, 1)); // across the left edge, 3 px under the top edge
      out.horizontal.push(min(p.x + 4, p.y - 2, 1, 4)); // across the top edge, 4 px right of the left edge
    }
    return out;
  }, shot);
};
const DARK = 165; // with the fix lines measure 0–156; a one-canvas-pixel line smeared by the stretch measures 170–260 and looks grey

await step('roster: borders between coloured cells are visible at 100 %', async () => {
  await openFile(path.resolve('tests/fixtures/files/roster.xlsx'));
  await page.waitForTimeout(400);
  const b = await borders();
  for (const v of [...b.vertical, ...b.horizontal]) assert.ok(v < DARK, JSON.stringify(b));
});

const values = {};
for (const z of [0.6, 0.7, 0.85, 1.15, 1.3, 1]) {
  await step(`zoom ${Math.round(z * 100)} %, after scrolling: every border is still a dark line`, async () => {
    await ev((z) => { window.__doc.wb.getActiveSheet().zoom(z); return 0; }, z);
    await page.waitForTimeout(400);
    await T.drag({ x: 330, y: 500 }, { x: 287, y: 471 }, { steps: 10, settle: 200 });
    await gestureEnd();
    const b = await borders();
    values[z] = b;
    console.log('   ', JSON.stringify(b));
    for (const v of [...b.vertical, ...b.horizontal]) assert.ok(v < DARK, JSON.stringify(b));
  });
}

await step('without the line fix the same borders are faint (the test sees the difference)', async () => {
  const stretched = await ev(() => { const u = window.__view.renderUnit(); const p = Object.getPrototypeOf(u.engine.getCanvas().getContext()); if (!p.__petakOrig) return false; p.setLineWidthByPrecision = p.__petakOrig; return true; });
  assert.ok(stretched, 'this profile stretches the canvas');
  await ev(() => { window.__doc.wb.getActiveSheet().zoom(0.85); return 0; });
  await page.waitForTimeout(400);
  await T.drag({ x: 330, y: 500 }, { x: 287, y: 471 }, { steps: 10, settle: 200 });
  await gestureEnd();
  const b = await borders();
  console.log('    without:', JSON.stringify(b));
  assert.ok([...b.vertical, ...b.horizontal].some((v) => v > DARK), JSON.stringify(b));
  await fresh();
  await page.waitForTimeout(300);
  await openFile(path.resolve('tests/fixtures/files/roster.xlsx'));
  await T.drag({ x: 330, y: 500 }, { x: 287, y: 471 }, { steps: 10, settle: 200 });
  await gestureEnd();
});

await step('frozen panes: the frozen rows and columns keep their own content while the rest scrolls', async () => {
  // rows 1–3 and columns A–B are frozen; the scrolling pane starts further in
  const s = await ev(() => { const V = window.__view; const v = V.viewInfo(); return { o: V.mainOrigin(v), fw: v.frozenW, fh: v.frozenH, hw: v.headerW, hh: v.headerH }; });
  assert.ok(s.fw > 50 && s.fh > 40, JSON.stringify(s));
  assert.ok(s.o.x > 180 && s.o.y > 140, 'scrolled: ' + JSON.stringify(s.o));
  // a tap in the frozen header row still selects that header cell, a tap below selects a scrolled row
  const g = await (await page.$('[data-testid=grid]')).boundingBox();
  await T.tap(g.x + s.hw + s.fw + 60, g.y + s.hh + s.fh - 12);
  await page.waitForTimeout(250);
  const top = await ev(() => window.__doc.wb.getActiveSheet().getActiveRange().getRange());
  assert.equal(top.startRow, 2, 'frozen header row 3');
  await T.tap(g.x + s.hw + s.fw + 60, g.y + s.hh + s.fh + 14);
  await page.waitForTimeout(250);
  const below = await ev(() => window.__doc.wb.getActiveSheet().getActiveRange().getRange());
  assert.ok(below.startRow > 3, 'first visible scrolled row: ' + below.startRow);
  assert.equal(top.startColumn, below.startColumn);
});

await t.finish();
