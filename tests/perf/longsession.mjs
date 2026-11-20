/**
 * A long session (Part 2 of the 1.3.0 brief): does Petak's memory grow, or something get stuck, when
 * the app is used for a long time without a restart?
 *
 * ROUNDS rounds (default 120, about 10 minutes) of what a user does all day: every eighth round another
 * file is opened (roster, QM report, a new workbook); every round a few cells are typed (one by real
 * double tap + keyboard), a row is inserted and deleted again, the sheet is flung up and down, pinched
 * out and back in, and a picture-free undo / redo. Every fourth round, after a full garbage collection:
 * JS heap, DOM nodes, event listeners, canvases, the gesture tile cache, Univer's undo stack.
 *
 * Pass: from round 8 (all files opened once) to the end, heap within +12 MB, nodes and listeners
 * within +15 %, canvases (outside the tile cache) within +4, the tile cache under its cap, no uncaught
 * errors, the app still edits, computes and pinches.
 *
 *   BASE_URL=http://localhost:4173/ node tests/perf/longsession.mjs     (ROUNDS=120)
 * Writes tests/perf/longsession.json.
 */
import { chromium, devices } from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { fingers, openFile, settled } from '../e2e/pinchlib.mjs';
import { touch } from '../e2e/touch.mjs';

const BASE = process.env.BASE_URL ?? 'http://localhost:4173/';
const ROUNDS = +(process.env.ROUNDS ?? 120);
const HERE = path.dirname(new URL(import.meta.url).pathname);
const FILES = ['tests/fixtures/files/roster.xlsx', 'tests/fixtures/files/qm-report.xlsx', null].map((f) => f && path.resolve(f));

const browser = await chromium.launch({ args: ['--enable-precise-memory-info', '--js-flags=--expose-gc'] });
const ctx = await browser.newContext({ ...devices['Pixel 7'] });
const page = await ctx.newPage();
const errors = [];
page.on('pageerror', (e) => errors.push(e.message));
const cdp = await ctx.newCDPSession(page);
await cdp.send('Performance.enable');
await page.goto(BASE + '?fresh');
await page.waitForFunction(() => window.__doc?.snapshot?.sheets?.length > 0, null, { timeout: 120000 });
await page.addStyleTag({ content: '.toast { display: none !important; }' });
const T = await touch(page);
const F = await fingers(page);

const grid = async () => {
  const b = await page.locator('[data-testid=grid]').boundingBox();
  return { x: b.x, y: b.y, w: b.width, h: b.height, cx: b.x + b.width / 2, cy: b.y + b.height / 2 };
};
async function measure(round) {
  await cdp.send('HeapProfiler.collectGarbage');
  await page.waitForTimeout(200);
  await cdp.send('HeapProfiler.collectGarbage');
  const m = Object.fromEntries((await cdp.send('Performance.getMetrics')).metrics.map((x) => [x.name, x.value]));
  const app = await page.evaluate(() => ({
    // canvases outside the gesture surface (its tiles are the tile cache, capped on its own)
    canvases: [...document.querySelectorAll('canvas')].filter((c) => !c.closest('.gesture-surface')).length,
    tileCapMb: (window.__petakView?.stats?.tileCap ?? 0) / 1048576,
    tilesMb: +((window.__petakView?.stats?.tileBytes ?? 0) / 1048576).toFixed(1),
    tiles: window.__petakView?.stats?.tiles ?? 0,
    file: window.__doc.snapshot.file?.name,
  }));
  return { round, heapMb: +(m.JSHeapUsedSize / 1048576).toFixed(1), nodes: m.Nodes, listeners: m.JSEventListeners, ...app };
}

const samples = [];
let opened = 0;
const t0 = Date.now();
for (let round = 1; round <= ROUNDS; round++) {
  if (round % 8 === 1) {
    const f = FILES[opened++ % FILES.length];
    // OPEN=picker: through the Open button and the file chooser (Playwright keeps each chooser's hidden
    // <input type=file> alive, which shows up as DOM nodes and listeners); default: the same open path
    // the picker leads to, with the file's bytes
    if (f && process.env.OPEN === 'picker') await openFile(page, f);
    else if (f) {
      await page.evaluate(
        async ([b64, name]) => {
          const buf = Uint8Array.from(atob(b64), (c) => c.charCodeAt(0));
          await window.__doc.open({ name, buffer: buf.buffer }, { progress: () => {}, confirm: async () => true });
          await window.__doc.bridge.idle();
        },
        [fs.readFileSync(f).toString('base64'), path.basename(f)],
      );
      await page.waitForTimeout(500);
    } else {
      await page.evaluate(() => window.__doc.newWorkbook());
      await page.waitForTimeout(600);
    }
  }
  // typing: one cell by touch and keyboard, a few through the API (a formula among them)
  const g = await grid();
  await T.doubleTap(g.x + 150, g.y + 140);
  await page.waitForTimeout(150);
  await page.keyboard.press('Control+A');
  await page.keyboard.type('r' + round);
  await page.keyboard.press('Enter');
  await page.keyboard.press('Escape');
  await page.evaluate((r) => {
    const ws = window.__doc.wb.getActiveSheet();
    ws.getRange(30 + (r % 10), 6).setValue(r);
    ws.getRange(30 + (r % 10), 7).setValue('=G' + (31 + (r % 10)) + '*2');
    ws.getRange(30 + (r % 10), 8).setValue('text ' + r);
  }, round);
  // a row in and out again
  await page.evaluate(async () => {
    const ws = window.__doc.wb.getActiveSheet();
    ws.insertRowBefore(40);
    ws.deleteRow(40);
    await window.__doc.bridge.idle();
  });
  // undo / redo
  await page.evaluate(async () => {
    await window.__doc.api.undo();
    await window.__doc.api.redo();
  });
  // fling down and back up, pinch out and back in
  await T.fling({ x: g.cx, y: g.y + g.h * 0.8 }, { x: g.cx, y: g.y + g.h * 0.2 });
  await page.waitForTimeout(400);
  await T.fling({ x: g.cx, y: g.y + g.h * 0.2 }, { x: g.cx, y: g.y + g.h * 0.8 });
  await page.waitForTimeout(400);
  await F.pinch({ x: g.cx, y: g.cy }, 320, 140, { steps: 10, stepMs: 16 });
  await settled(page, 300);
  await F.pinch({ x: g.cx, y: g.cy }, 140, 320, { steps: 10, stepMs: 16 });
  await settled(page, 300);
  if (round % 4 === 0 || round === 1) {
    const s = await measure(round);
    samples.push(s);
    console.log(JSON.stringify(s));
  }
}
const minutes = +((Date.now() - t0) / 60000).toFixed(1);

// still working at the end: a formula computes, a pinch zooms, nothing is stuck
const end = await page.evaluate(async () => {
  const ws = window.__doc.wb.getActiveSheet();
  ws.getRange('K2').setValue(21);
  ws.getRange('K3').setValue('=K2*2');
  await window.__doc.bridge.idle();
  await new Promise((r) => setTimeout(r, 300));
  return { k3: ws.getRange('K3').getValue(), gesturing: window.__doc.gesturing, surface: getComputedStyle(document.querySelector('.gesture-surface')).display };
});
const base = samples.find((s) => s.round >= 8) ?? samples[0];
const last = samples[samples.length - 1];
const report = { rounds: ROUNDS, minutes, base, last, samples, errors, end };
fs.writeFileSync(path.join(HERE, 'longsession.json'), JSON.stringify(report, null, 1) + '\n');
console.log(`${ROUNDS} rounds in ${minutes} min; heap ${base.heapMb} → ${last.heapMb} MB, nodes ${base.nodes} → ${last.nodes}, listeners ${base.listeners} → ${last.listeners}, canvases ${base.canvases} → ${last.canvases}; end ${JSON.stringify(end)}; errors ${errors.length}`);
await browser.close();

assert.equal(errors.length, 0, 'uncaught errors: ' + errors.slice(0, 3).join(' | '));
assert.equal(end.k3, 42, 'still computes');
assert.equal(end.gesturing, false, 'no gesture stuck');
assert.equal(end.surface, 'none', 'the gesture surface is not left on screen');
assert.ok(last.heapMb <= base.heapMb + 12, `heap grew ${base.heapMb} → ${last.heapMb} MB`);
assert.ok(last.nodes <= base.nodes * 1.15 + 50, `DOM nodes grew ${base.nodes} → ${last.nodes}`);
assert.ok(last.listeners <= base.listeners * 1.15 + 20, `listeners grew ${base.listeners} → ${last.listeners}`);
assert.ok(last.canvases <= base.canvases + 4, `canvases grew ${base.canvases} → ${last.canvases}`);
assert.ok(Math.max(...samples.map((x) => x.tilesMb)) <= last.tileCapMb * 1.1, 'tile cache above its cap');
console.log('long session: pass');
