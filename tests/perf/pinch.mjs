/**
 * Pinch-zoom benchmark (before / after numbers for the 1.3.0 renderer).
 *
 * For each tier (CPU slowed ÷6 / ÷4 / ÷1 for Low / Mid / High phones) and file (empty sheet at FN1300,
 * staff roster, 50,000 rows), pinch in and out for SECONDS while every animation frame is timed:
 *   fps, dropped frames (a frame that took longer than 1.5 vsync intervals), p95 / worst frame,
 *   main-thread busy time per second, hand-off time (finger lifted → sharp grid shown),
 *   peak canvas memory during the gestures and JS heap after garbage collection.
 * Run twice: once with the browser's normal vsync (60 Hz here: dropped frames), once uncapped
 * (`--disable-frame-rate-limit`: how many frames per second the work allows, i.e. whether 120 Hz fits).
 *
 *   npx vite build && npx vite preview --port 4173 &
 *   BASE_URL=http://localhost:4173/ LABEL=after node tests/perf/pinch.mjs
 * Writes tests/perf/pinch-<LABEL>.json; tests/perf/PINCH.md is written from both files.
 */
import { chromium, devices } from 'playwright';
import fs from 'node:fs';
import path from 'node:path';
import { ROSTER, fingers, makeRows, openFile, place, settled } from '../e2e/pinchlib.mjs';

const BASE = process.env.BASE_URL ?? 'http://localhost:4173/';
const SECONDS = +(process.env.SECONDS ?? 8);
const LABEL = process.env.LABEL ?? 'run';
const TIERS = (process.env.TIERS ?? 'high:1,mid:4,low:6').split(',').map((x) => x.split(':')).map(([t, c]) => [t, +c]);
const MODES = (process.env.MODES ?? 'vsync,uncapped').split(',');
const HERE = path.dirname(new URL(import.meta.url).pathname);
const BIG = await makeRows(50000, 20);

async function launch(tier, throttle, mode) {
  const args = ['--enable-precise-memory-info', '--js-flags=--expose-gc'];
  if (mode === 'uncapped') args.push('--disable-frame-rate-limit', '--disable-gpu-vsync');
  const browser = await chromium.launch({ executablePath: process.env.CHROMIUM_PATH || undefined, args });
  const ctx = await browser.newContext({ ...devices['Pixel 7'] });
  const page = await ctx.newPage();
  const errors = [];
  page.on('pageerror', (e) => errors.push(e.message));
  const cdp = await ctx.newCDPSession(page);
  await cdp.send('Performance.enable');
  await page.goto(BASE + '?fresh&tier=' + tier);
  await page.waitForFunction(() => window.__doc?.snapshot?.sheets?.length > 0, null, { timeout: 120000 });
  await page.addStyleTag({ content: '.toast { display: none !important; }' });
  return { browser, page, cdp, errors, throttle };
}
const metrics = async (cdp) => Object.fromEntries((await cdp.send('Performance.getMetrics')).metrics.map((m) => [m.name, m.value]));

/** Instrument the page: frame times, hand-off times, canvas memory peaks. */
async function instrument(page) {
  await page.evaluate(() => {
    const s = (window.__bench = { frames: [], handoff: [], peakCanvasMb: 0, liftAt: 0, gen: (window.__bench?.gen ?? 0) + 1 });
    const gen = s.gen;
    let last = performance.now();
    const canvasMb = () => {
      let px = 0;
      for (const c of document.querySelectorAll('canvas')) px += c.width * c.height;
      // 1.3.0 keeps its tiles in the DOM (hidden when not shown): they are counted above
      return (px * 4) / 1048576;
    };
    const visibleOverlay = () => {
      for (const sel of ['.grid-compositor', '.gesture-surface']) {
        const el = document.querySelector(sel);
        if (el && getComputedStyle(el).display !== 'none' && getComputedStyle(el).visibility !== 'hidden') return true;
      }
      return false;
    };
    window.addEventListener('touchend', (e) => e.touches.length === 0 && (s.liftAt = performance.now()), true);
    let n = 0;
    const tick = (t) => {
      s.frames.push(t - last);
      last = t;
      if (++n % 10 === 0) s.peakCanvasMb = Math.max(s.peakCanvasMb, canvasMb());
      if (s.liftAt && !window.__doc.gesturing && !visibleOverlay()) {
        // first frame without the gesture picture: the sharp grid is on screen
        s.handoff.push(performance.now() - s.liftAt);
        s.liftAt = 0;
      }
      if (window.__bench?.gen === gen) requestAnimationFrame(tick);
    };
    requestAnimationFrame(tick);
  });
}

async function pinching(page, F, seconds) {
  const box = await page.locator('[data-testid=grid]').boundingBox();
  const t0 = Date.now();
  let n = 0;
  while (Date.now() - t0 < seconds * 1000) {
    // focal point wanders so new parts of the sheet come into view
    const c = { x: box.x + box.width * (0.35 + 0.3 * ((n * 7) % 5) / 4), y: box.y + box.height * (0.35 + 0.3 * ((n * 3) % 4) / 3) };
    await F.pinch(c, 140, 380, { steps: 14, stepMs: 16 });
    await page.waitForTimeout(180);
    await F.pinch(c, 380, 140, { steps: 14, stepMs: 16 });
    await page.waitForTimeout(180);
    n++;
  }
  return n;
}

const results = [];
for (const mode of MODES) {
  for (const [tier, throttle] of TIERS) {
    for (const file of ['empty FN1300', 'roster', '50k rows']) {
      const L = await launch(tier, throttle, mode);
      const { page, cdp } = L;
      try {
        if (file === 'roster') await openFile(page, ROSTER);
        if (file === '50k rows') await openFile(page, BIG, 50000);
        if (file === 'empty FN1300') await place(page, 1, 1299, 169);
        else if (file === '50k rows') await place(page, 1, 20000, 0);
        else await place(page, 1, 0, 0);
        await settled(page, 600);
        if (throttle > 1) await cdp.send('Emulation.setCPUThrottlingRate', { rate: throttle });
        const F = await fingers(page);
        await instrument(page);
        const m0 = await metrics(cdp);
        const t0 = Date.now();
        const gestures = await pinching(page, F, SECONDS);
        const ms = Date.now() - t0;
        await settled(page, 300);
        const m1 = await metrics(cdp);
        const b = await page.evaluate(() => {
          const s = window.__bench;
          s.gen++;
          return { frames: s.frames.slice(2), handoff: s.handoff, peakCanvasMb: s.peakCanvasMb };
        });
        await cdp.send('Emulation.setCPUThrottlingRate', { rate: 1 });
        await cdp.send('HeapProfiler.collectGarbage').catch(() => {});
        const heap = (await metrics(cdp)).JSHeapUsedSize / 1048576;
        const f = b.frames;
        const sorted = [...f].sort((a, b) => a - b);
        const avg = f.reduce((a, x) => a + x, 0) / f.length;
        const vs = sorted[Math.floor(sorted.length * 0.2)]; // the vsync interval (most frames hit it)
        const dropped = f.filter((x) => x > vs * 1.5).length;
        const ho = [...b.handoff].sort((a, b) => a - b);
        const row = {
          mode,
          tier,
          file,
          gestures: gestures * 2,
          frames: f.length,
          fps: +(1000 / avg).toFixed(1),
          dropped: +((dropped / f.length) * 100).toFixed(1),
          p95: +sorted[Math.floor(sorted.length * 0.95)].toFixed(1),
          worst: Math.round(sorted[sorted.length - 1]),
          busyMsPerS: Math.round(((m1.TaskDuration - m0.TaskDuration) * 1000) / (ms / 1000)),
          handoffP50: ho.length ? Math.round(ho[Math.floor(ho.length / 2)]) : null,
          handoffMax: ho.length ? Math.round(ho[ho.length - 1]) : null,
          peakCanvasMb: +b.peakCanvasMb.toFixed(1),
          heapMb: +heap.toFixed(1),
          errors: L.errors.length,
        };
        results.push(row);
        console.log(JSON.stringify(row));
      } catch (e) {
        console.log('FAILED', mode, tier, file, e.message);
        results.push({ mode, tier, file, error: e.message.split('\n')[0] });
      } finally {
        await L.browser.close();
      }
    }
  }
}
fs.writeFileSync(path.join(HERE, `pinch-${LABEL}.json`), JSON.stringify(results, null, 1) + '\n');
console.log('wrote', path.join(HERE, `pinch-${LABEL}.json`));
