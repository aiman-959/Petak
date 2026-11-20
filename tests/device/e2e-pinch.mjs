/**
 * Pinch zoom on a real Android WebView (CI emulator, or a phone over USB).
 *
 * Two fingers are sent to the WebView over DevTools (adb's `input` has no multi-touch); the page
 * is drawn and composited by the real WebView. While the fingers rest mid-gesture the *screen*
 * is captured with `adb screencap` (SurfaceFlinger's buffer: exactly what is on the glass) and
 * checked for solid-black blocks; the WebView's own frame is compared with the grid after the
 * fingers lift (one grid, headers and frozen panes in place). Also: frame times during 6 s of
 * pinching, a DevTools performance trace of one pinch, and graphics memory (dumpsys meminfo).
 *
 *   node tests/device/e2e-pinch.mjs [debug.apk] [fixtures dir]
 * Output: device-out/pinch-*.png, device-out/pinch-trace.json, device-out/device-pinch.json.
 */
import { adb, sh, sleep, attach, launch, report, PKG, OUT } from './lib.mjs';
import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { analyse, fingers, makeRows, place, regionsOf, settled, gridBox } from '../e2e/pinchlib.mjs';

const [apk, fixtures = 'tests/fixtures/files'] = process.argv.slice(2);
const R = report();
if (apk && !process.env.NO_INSTALL) adb('install', '-r', '-g', apk);
sh(`am force-stop ${PKG}`);
launch();
await sleep(6000);
const { page } = await attach();
await page.waitForFunction(() => window.__doc?.snapshot?.sheets?.length > 0, null, { timeout: 90000 });
await sleep(1500);
const F = await fingers(page);

/** Open bytes in the app the way the file picker does. */
async function openBytes(file) {
  const b64 = fs.readFileSync(file).toString('base64');
  await page.evaluate(
    async ([b64, name]) => {
      const bin = atob(b64);
      const buf = new Uint8Array(bin.length);
      for (let i = 0; i < bin.length; i++) buf[i] = bin.charCodeAt(i);
      await window.__doc.open({ name, buffer: buf.buffer }, { progress: () => {}, confirm: async () => true });
      await window.__doc.bridge.idle();
    },
    [b64, path.basename(file)],
  );
  await sleep(1200);
}

/** Screen capture (what is really on the glass), checked for solid black blocks inside the grid. */
async function glassBlack(name) {
  const png = execFileSync('adb', ['exec-out', 'screencap', '-p'], { maxBuffer: 64 << 20 });
  fs.writeFileSync(path.join(OUT, `pinch-${name}.png`), png);
  return page.evaluate(
    async ([b64]) => {
      const img = await createImageBitmap(await (await fetch('data:image/png;base64,' + b64)).blob());
      const c = new OffscreenCanvas(img.width, img.height);
      const ctx = c.getContext('2d', { willReadFrequently: true });
      ctx.drawImage(img, 0, 0);
      // the grid's place on the glass: the WebView fills the screen width; find it from the page's own size
      const k = img.width / innerWidth;
      const r = document.querySelector('[data-testid=grid]').getBoundingClientRect();
      const top = img.height - innerHeight * k; // status bar above the WebView (edge-to-edge: 0)
      const x0 = Math.round(r.left * k);
      const y0 = Math.round(Math.max(0, top) + r.top * k);
      const w = Math.round(r.width * k);
      const h = Math.round(r.height * k);
      const d = ctx.getImageData(x0, y0, w, h).data;
      const B = Math.round(12 * (k / 2.625)) || 12;
      let blocks = 0;
      let black = 0;
      for (let by = 0; by + B <= h; by += B)
        for (let bx = 0; bx + B <= w; bx += B) {
          blocks++;
          let all = true;
          for (let yy = by; yy < by + B && all; yy++)
            for (let xx = bx; xx < bx + B; xx++) {
              const i = (yy * w + xx) * 4;
              if (d[i] + d[i + 1] + d[i + 2] > 60) {
                all = false;
                break;
              }
            }
          if (all) black++;
        }
      return +((black / Math.max(1, blocks)) * 100).toFixed(2);
    },
    [png.toString('base64')],
  );
}

/**
 * Where the page sits on the glass: CDP screenshots on Android show the page as it is laid out, not
 * necessarily what the compositor put on the glass (during a gesture they showed finished tiles the
 * glass did not have yet), so the comparisons use screen captures. Found once by sliding a resting
 * screen capture over a page screenshot of the same moment.
 */
let glassTop = null;
async function calibrateGlass() {
  const g = execFileSync('adb', ['exec-out', 'screencap', '-p'], { maxBuffer: 64 << 20 }).toString('base64');
  const p = (await page.screenshot()).toString('base64');
  glassTop = await page.evaluate(
    async ([g, p]) => {
      const W = 206; // compare at half the CSS width
      const load = async (b64) => {
        const img = await createImageBitmap(await (await fetch('data:image/png;base64,' + b64)).blob());
        const h = Math.round((img.height * W) / img.width);
        const c = new OffscreenCanvas(W, h);
        const ctx = c.getContext('2d', { willReadFrequently: true });
        ctx.drawImage(img, 0, 0, W, h);
        return { d: ctx.getImageData(0, 0, W, h).data, h, scale: img.width / W };
      };
      const [G, P] = await Promise.all([load(g), load(p)]);
      let best = 0;
      let bestErr = Infinity;
      for (let dy = 0; dy + P.h <= G.h; dy++) {
        let err = 0;
        for (let y = 0; y < P.h; y += 3)
          for (let x = 0; x < W; x += 3) {
            const i = (y * W + x) * 4;
            const j = ((y + dy) * W + x) * 4;
            err += Math.abs(P.d[i] - G.d[j]) + Math.abs(P.d[i + 1] - G.d[j + 1]) + Math.abs(P.d[i + 2] - G.d[j + 2]);
          }
        if (err < bestErr) {
          bestErr = err;
          best = dy;
        }
      }
      return Math.round(best * G.scale); // glass px
    },
    [g, p],
  );
  return glassTop;
}
/** The glass now, cut to the page (so page regions apply), as base64 PNG; also saved to device-out. */
async function glassShot(name) {
  const png = execFileSync('adb', ['exec-out', 'screencap', '-p'], { maxBuffer: 64 << 20 });
  fs.writeFileSync(path.join(OUT, `pinch-${name}.png`), png);
  return page.evaluate(
    async ([b64, top]) => {
      const img = await createImageBitmap(await (await fetch('data:image/png;base64,' + b64)).blob());
      const k = img.width / innerWidth;
      const h = Math.round(innerHeight * k);
      const c = new OffscreenCanvas(img.width, h);
      c.getContext('2d').drawImage(img, 0, -top);
      const blob = await c.convertToBlob({ type: 'image/png' });
      const bytes = new Uint8Array(await blob.arrayBuffer());
      let s = '';
      for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
      return btoa(s);
    },
    [png.toString('base64'), glassTop ?? 0],
  );
}

const results = [];
async function pinchCheck(label, z0, z1, row, col) {
  await place(page, z0, row, col);
  const g = await gridBox(page);
  const c = { x: g.x + g.w / 2, y: g.y + g.h / 2 };
  const span = Math.min(g.w, g.h) * 0.9;
  const r = z1 / z0;
  const [d0, d1] = r >= 1 ? [Math.max(40, span / r), span] : [span, Math.max(40, span * r)];
  await F.pinch(c, d0, d1, { steps: 14, stepMs: 16, release: false });
  await sleep(400);
  const blackGlass = await glassBlack(label + '-during');
  const during = await glassShot(label + '-during-page');
  await F.up();
  await settled(page, 600);
  const g1 = await gridBox(page);
  const after = await glassShot(label + '-after-page');
  const blackGlassAfter = await glassBlack(label + '-after');
  // the glass during the gesture against the glass after it, region by region
  const a = await analyse(page, [during, after], regionsOf(g1));
  const row1 = { label, z0, zoom: +g1.zoom.toFixed(3), want: +Math.max(0.1, Math.min(4, z0 * (d1 / d0))).toFixed(3), blackGlass, blackGlassAfter };
  for (const k of Object.keys(a)) if (a[k].diff !== undefined) row1['diff_' + k] = a[k].diff;
  const worst = Math.max(...Object.entries(a).filter(([k]) => g1.zoom >= 0.2 || !/Header/.test(k)).map(([, v]) => v.diff ?? 0));
  row1.ok = blackGlass <= blackGlassAfter + 0.15 && worst <= 16 && Math.abs(row1.zoom - row1.want) / row1.want < 0.03;
  results.push(row1);
  R.note(`pinch ${label}`, row1);
}

await R.check('device', () => page.evaluate(() => ({ ua: navigator.userAgent.replace(/.*(Chrome\/[\d.]+).*/, '$1'), dpr: devicePixelRatio, tier: document.documentElement.dataset.tier })));
await R.check('canvas limits: largest canvas Petak creates (px), tile size, cache cap', () =>
  page.evaluate(() => {
    const all = [...document.querySelectorAll('canvas')].map((c) => [c.width, c.height]);
    const s = window.__petakView?.stats;
    return { domCanvases: all, tileCacheCapMb: s ? s.tileCap / 1048576 : null, tile: 512 };
  }),
);

// 1. empty sheet far from A1
await page.evaluate(() => window.__doc.newWorkbook());
await sleep(1500);
await R.check('where the page is on the glass (screen px from the top)', () => calibrateGlass());
await sleep(1500);
for (const [z0, z1] of [[1, 0.1], [0.1, 0.25], [0.25, 0.5], [1, 2], [2, 4], [4, 1]]) await pinchCheck(`FN1300 ${z0}→${z1}`, z0, z1, 1299, 169);
// 2. the roster (merged cells, fills, large fonts, frozen panes)
await openBytes(path.join(fixtures, 'roster.xlsx'));
for (const [z0, z1] of [[1, 0.25], [0.25, 1], [1, 2], [2, 4], [4, 1]]) await pinchCheck(`roster ${z0}→${z1}`, z0, z1, 0, 0);
// 3. 50,000 rows
await openBytes(await makeRows(50000, 20));
for (const [z0, z1] of [[1, 0.1], [0.1, 0.5], [0.5, 2], [2, 4], [4, 1]]) await pinchCheck(`50k ${z0}→${z1}`, z0, z1, 24000, 2);

// frame times while pinching for 6 s, graphics memory before / after
const meminfo = () => {
  try {
    const m = sh(`dumpsys meminfo ${PKG}`);
    const pick = (re) => +((re.exec(m) ?? [])[1] ?? NaN);
    return { totalPssKb: pick(/TOTAL PSS:\s+(\d+)/) || pick(/TOTAL\s+(\d+)/), graphicsKb: pick(/Graphics:\s+(\d+)/), eglKb: pick(/EGL mtrack\s+(\d+)/), glKb: pick(/GL mtrack\s+(\d+)/) };
  } catch (e) {
    return { error: String(e.message).slice(0, 80) };
  }
};
await R.check('graphics memory before 6 s of pinching (dumpsys meminfo, KB)', meminfo);
await place(page, 1, 20000, 0);
await page.evaluate(() => {
  const s = (window.__bench = { frames: [] });
  let last = performance.now();
  const tick = (t) => {
    s.frames.push(t - last);
    last = t;
    if (!s.stop) requestAnimationFrame(tick);
  };
  requestAnimationFrame(tick);
});
const g = await gridBox(page);
const c = { x: g.x + g.w / 2, y: g.y + g.h / 2 };
const t0 = Date.now();
while (Date.now() - t0 < 6000) {
  await F.pinch(c, 140, 360, { steps: 14, stepMs: 16 });
  await sleep(150);
  await F.pinch(c, 360, 140, { steps: 14, stepMs: 16 });
  await sleep(150);
}
await settled(page, 300);
await R.check('frame times while pinching 6 s (emulator GPU: compare runs, not with a phone)', () =>
  page.evaluate(() => {
    const s = window.__bench;
    s.stop = true;
    const f = s.frames.slice(2).sort((a, b) => a - b);
    const avg = f.reduce((a, b) => a + b, 0) / f.length;
    return { frames: f.length, fps: +(1000 / avg).toFixed(1), p50: +f[Math.floor(f.length / 2)].toFixed(1), p95: +f[Math.floor(f.length * 0.95)].toFixed(1), worst: Math.round(f[f.length - 1]), view: window.__petakView?.stats };
  }),
);
await R.check('graphics memory after (KB)', meminfo);

// one pinch under the DevTools tracer (open device-out/pinch-trace.json in chrome://tracing or Perfetto)
await R.check('performance trace of one pinch', async () => {
  const cdp = F.cdp;
  const chunks = [];
  cdp.on('Tracing.dataCollected', (e) => chunks.push(...e.value));
  await cdp.send('Tracing.start', { categories: 'devtools.timeline,disabled-by-default-devtools.timeline.frame,blink,cc,gpu,viz', transferMode: 'ReportEvents' });
  await F.pinch(c, 160, 380, { steps: 20, stepMs: 16 });
  await settled(page, 300);
  const done = new Promise((r) => cdp.once('Tracing.tracingComplete', r));
  await cdp.send('Tracing.end');
  await done;
  fs.writeFileSync(path.join(OUT, 'pinch-trace.json'), JSON.stringify({ traceEvents: chunks }));
  const frames = chunks.filter((e) => e.name === 'DrawFrame' || e.name === 'Graphics.Pipeline').length;
  const long = chunks.filter((e) => e.ph === 'X' && e.name === 'RunTask' && e.dur > 16700).length;
  return { events: chunks.length, drawFrames: frames, mainThreadTasksOver16ms: long };
});

const failed = results.filter((r) => !r.ok);
R.note('PINCH SUMMARY', `${results.length - failed.length}/${results.length} pinches: no black on the glass, one grid, right zoom` + (failed.length ? ' — FAILED: ' + failed.map((r) => r.label).join(', ') : ''));
R.save('device-pinch.json');
process.exit(0);
