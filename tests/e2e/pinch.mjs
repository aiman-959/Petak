/**
 * Pinch zoom (Part 1 of the 1.3.0 brief): what the user sees *while the fingers are down*.
 *
 * For an empty sheet scrolled to FN1300, the staff roster (merged cells, fills, large fonts, frozen
 * rows and columns) and a 50,000-row file, at 10, 25, 50, 100, 200 and 400 % zoom, pinch in and out
 * and take screenshots during the gesture:
 *   - no black: no 12 × 12-pixel block may be solid black
 *   - one grid: with the fingers resting, the picture must match what the grid shows after the
 *     fingers lift at the same zoom and place (headers, frozen panes and the scrolling pane are
 *     compared separately: a second, scaled copy of the grid or a header drawn twice differs a lot)
 *   - the zoom after the gesture is the one the fingers asked for (10 % … 400 %)
 * Also fast, slow and repeated pinches, pinch → one-finger scroll → pinch, double tap to zoom on an
 * empty area, and the zoom presets / slider.
 *
 *   BASE_URL=http://localhost:4173/ node tests/e2e/pinch.mjs
 */
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { start, BASE } from './lib.mjs';
import { ROSTER, analyse, closeMenus, fingers, gridBox, makeRows, openFile, place, regionsOf, settled } from './pinchlib.mjs';

const t = await start('pinch');
const { page, step, ev, fresh, tmp } = t;
console.log('Petak e2e pinch zoom —', BASE);
const BIG = await makeRows(50000, 20);
await fresh();
// DOM decorations that are not part of the grid picture (the fill grip, toasts, scroll indicators)
await page.addStyleTag({ content: '.fill-handle, .toast, .scroll-ind, .zoom-pill { visibility: hidden !important; }' });
const F = await fingers(page);
const MIN = 0.1;
const MAX = 4;
const LEVELS = [0.1, 0.25, 0.5, 1, 2, 4];
const SPAN = 470; // largest finger distance (diagonal) that stays on a 412 px wide screen
const report = [];
const shot = async (name) => {
  const b = await page.screenshot();
  if (process.env.SHOTS) fs.writeFileSync(path.join(tmp, name + '.png'), b);
  return b.toString('base64');
};
// thresholds (mean absolute difference 0–255 after 4× downscale). A second grid, a header drawn twice
// or a misplaced pane measures 100–250; tiles of the nearest zoom level, slightly softer than the
// sharp redraw, measure under 15. Below 20 % rows are under 4 px and Univer leaves their text out,
// while a 25 % tile still has it (unreadable grey): headers are allowed more there.
const DIFF = { colHeader: 16, rowHeader: 16, frozenTop: 16, frozenLeft: 16, main: 16, whole: 14 };
const SMALL = 0.2;

/**
 * Pinch from zoom z0 to z1 around the grid's centre, keep the fingers down, look, release, look.
 * Returns the analysis; asserts no black and that the resting picture equals the settled one.
 */
async function pinchCheck(label, z0, z1, { row = 0, col = 0, steps = 16, stepMs = 16, focal } = {}) {
  await closeMenus(page);
  await place(page, z0, row, col);
  const g0 = await gridBox(page);
  const c = focal ?? { x: g0.x + g0.w / 2, y: g0.y + g0.h / 2 };
  const r = z1 / z0;
  const [d0, d1] = r >= 1 ? [Math.max(40, SPAN / r), SPAN] : [SPAN, Math.max(40, SPAN * r)];
  await F.pinch(c, d0, d1, { steps, stepMs, release: false });
  const early = await shot(label + '-early'); // the first frames: content may still be filling in, but never black
  await page.waitForTimeout(400);
  const during = await shot(label + '-during');
  await F.up();
  await settled(page);
  const g1 = await gridBox(page);
  const after = await shot(label + '-after');
  const regions = regionsOf(g1);
  const a = await analyse(page, [during, after], regions);
  const e = await analyse(page, [early], { whole: regions.whole });
  const want = Math.max(MIN, Math.min(MAX, z0 * (d1 / d0)));
  const row1 = { label, z0, want: +want.toFixed(3), zoom: +g1.zoom.toFixed(3), blackEarly: e.whole.black[0], blackDuring: a.whole.black[0], blackAfter: a.whole.black[1] };
  for (const k of Object.keys(regions)) if (a[k]) row1['diff_' + k] = a[k].diff;
  report.push(row1);
  console.log('    ', JSON.stringify(row1));
  assert.ok(Math.abs(g1.zoom - want) / want < 0.03, `zoom after the gesture ${g1.zoom} ≠ ${want}`);
  // solid black 12 × 12 blocks: only as many as the settled grid has (thick black text at 400 %)
  assert.ok(e.whole.black[0] <= a.whole.black[1] + 0.15, `black blocks in the first frames of the gesture: ${e.whole.black[0]} % (settled grid ${a.whole.black[1]} %)`);
  assert.ok(a.whole.black[0] <= a.whole.black[1] + 0.15, `black blocks while the fingers rest: ${a.whole.black[0]} % (settled grid ${a.whole.black[1]} %)`);
  for (const [k, lim] of Object.entries(DIFF)) {
    if (!a[k]) continue;
    const limit = g1.zoom < SMALL && /Header/.test(k) ? 45 : lim;
    assert.ok(a[k].diff <= limit, `${k}: picture during the gesture differs from the settled grid (${a[k].diff} > ${limit})`);
  }
  return a;
}

const files = [
  {
    name: 'empty sheet at FN1300',
    open: async () => {
      await fresh();
      await page.addStyleTag({ content: '.fill-handle, .toast, .scroll-ind, .zoom-pill { visibility: hidden !important; }' });
    },
    row: 1299,
    col: 169,
  },
  { name: 'roster', open: () => openFile(page, ROSTER), row: 0, col: 0 },
  { name: '50k rows', open: () => openFile(page, BIG, 50000), row: 24000, col: 2 },
];

for (const f of files) {
  await step(`${f.name}: opens`, async () => {
    await f.open();
    await settled(page, 200);
  });
  for (const z of LEVELS) {
    await step(`${f.name}: pinch out at ${z * 100} %`, async () => {
      await pinchCheck(`${f.name}-out-${z}`, z, Math.max(MIN, z / 1.8), f);
    });
    await step(`${f.name}: pinch in at ${z * 100} %`, async () => {
      await pinchCheck(`${f.name}-in-${z}`, z, Math.min(MAX, z * 1.8), f);
    });
  }
}

// ------------------------------------------------------------------ gesture shapes (50k file open)
await step('fast pinch (3 moves, no pauses) out to 25 %', async () => {
  await pinchCheck('fast-out', 1, 0.25, { row: 100, col: 0, steps: 3, stepMs: 0 });
});
await step('slow pinch (40 moves over 1.2 s) in to 300 %', async () => {
  await pinchCheck('slow-in', 1, 3, { row: 100, col: 0, steps: 40, stepMs: 30 });
});
await step('ten pinches in a row without waiting: no black, ends at the last zoom', async () => {
  await place(page, 1, 500, 0);
  const g = await gridBox(page);
  const c = { x: g.x + g.w / 2, y: g.y + g.h / 2 };
  let black = 0;
  for (let i = 0; i < 10; i++) {
    const [a, b] = i % 2 ? [160, 420] : [420, 160];
    await F.pinch(c, a, b, { steps: 6, stepMs: 8, release: false });
    if (i % 3 === 0) {
      const s = await shot('repeat-' + i);
      black += (await analyse(page, [s], { whole: regionsOf(g).whole })).whole.black[0];
    }
    await F.up();
  }
  await settled(page);
  const z = await ev(() => window.__view.viewInfo().zoom);
  // five times ×(160/420) and five times ×(420/160) from 100 %: back at 100 %
  assert.ok(Math.abs(z - 1) < 0.03, 'zoom ' + z);
  assert.equal(black, 0, 'black blocks during repeated pinches');
});
await step('pinch, then scroll with the finger left on the glass, then pinch again', async () => {
  await place(page, 1, 300, 0);
  const g = await gridBox(page);
  const c = { x: g.x + g.w / 2, y: g.y + g.h / 2 };
  const pts = await F.pinch(c, 200, 400, { steps: 10, stepMs: 16, release: false }); // → 200 %
  const z1 = await ev(() => window.__petakView?.camera?.z ?? null);
  // lift finger 0; finger 1 keeps going: the sheet must follow it (no jump, no stop)
  await F.up([pts[0]]);
  const f1 = pts[1];
  const y0 = await ev(() => window.__petakView?.camera?.y ?? window.__view.viewInfo().scrollY);
  for (let i = 1; i <= 12; i++) await F.move([{ ...f1, y: f1.y - i * 15 }]) && (await page.waitForTimeout(16));
  await page.waitForTimeout(80);
  const y1 = await ev(() => window.__petakView?.camera?.y ?? window.__view.viewInfo().scrollY);
  const s1 = await shot('handoff-scroll');
  // second finger again: pinch out back to about 100 %
  const f1b = { ...f1, y: f1.y - 180 };
  const other = { x: f1b.x - 280, y: f1b.y, id: 2 };
  await F.down([f1b, other]);
  for (let i = 1; i <= 10; i++) {
    await F.move([{ ...f1b, x: f1b.x - i * 7 }, { ...other, x: other.x + i * 7 }]);
    await page.waitForTimeout(16);
  }
  const s2 = await shot('handoff-pinch2');
  await F.up();
  await settled(page);
  const z2 = await ev(() => window.__view.viewInfo().zoom);
  const a = await analyse(page, [s1, s2], { whole: regionsOf(g).whole });
  console.log('     zoom after first pinch', z1, 'scroll moved', y1 - y0, 'sheet px; zoom at the end', z2, 'black', a.whole.black);
  assert.ok(y1 - y0 > 40, 'one finger left after a pinch scrolls the sheet (moved ' + (y1 - y0) + ')');
  assert.ok(Math.abs(z2 - 2 * (140 / 280)) < 0.06, 'second pinch zooms from where the first ended: ' + z2);
  assert.deepEqual(a.whole.black, [0, 0]);
});

await step('pinch the moment the app is ready, and right after opening a file: zooms, without errors', async () => {
  const pinchNow = async () => {
    const g = await gridBox(page);
    await F.pinch({ x: g.x + g.w / 2, y: g.y + g.h / 2 }, 160, 320, { steps: 6, stepMs: 8 });
    await settled(page, 300);
    return ev(() => window.__view.viewInfo().zoom);
  };
  for (let i = 0; i < 2; i++) {
    await page.goto(BASE + '?fresh');
    // the earliest a finger can reach the grid: the splash screen and Univer's loading placeholder are gone
    await page.waitForSelector('.splash', { state: 'detached', timeout: 60000 });
    await page.waitForFunction(() => {
      const r = document.querySelector('[data-testid=grid]').getBoundingClientRect();
      return document.elementFromPoint(r.left + r.width / 2 - 60, r.top + r.height / 2 - 60)?.tagName === 'CANVAS';
    }, null, { timeout: 30000, polling: 16 });
    const z = await pinchNow();
    assert.ok(Math.abs(z - 2) < 0.05, 'zoom after a pinch at start-up ' + z);
  }
  // a new workbook has a new renderer whose zoom controller arrives a moment later
  const before = await ev(() => window.__doc.wb.getId());
  const [chooser] = await Promise.all([page.waitForEvent('filechooser'), page.click('[data-testid=open]')]);
  await chooser.setFiles(ROSTER);
  await page.waitForFunction((id) => {
    const r = document.querySelector('[data-testid=grid]').getBoundingClientRect();
    return window.__doc.wb.getId() !== id && window.__view.viewInfo() && document.elementFromPoint(r.left + r.width / 2 - 60, r.top + r.height / 2 - 60)?.tagName === 'CANVAS';
  }, before, { timeout: 60000, polling: 16 });
  const z = await pinchNow();
  assert.ok(Math.abs(z - 2) < 0.05, 'zoom after a pinch right after opening a file ' + z);
  // the same, made certain: Univer refuses the first five zoom requests as it does while its zoom
  // controller is not registered ("Cannot find SheetsZoomRenderController"); the hand-off asks again
  await place(page, 1, 0, 0);
  await ev(() => {
    const api = window.__doc.api;
    const orig = api.executeCommand.bind(api);
    let refused = 0;
    api.executeCommand = (id, params, opts) => {
      if (id === 'sheet.operation.set-zoom-ratio' && refused < 5) {
        refused++;
        return Promise.reject(new Error('[redi]: Cannot find "SheetsZoomRenderController" registered by any injector.'));
      }
      return orig(id, params, opts);
    };
    window.__refused = () => refused;
  });
  const z2 = await pinchNow();
  assert.equal(await ev(() => window.__refused()), 5);
  assert.ok(Math.abs(z2 - 2) < 0.05, 'zoom after Univer refused five times ' + z2);
});

// ------------------------------------------------------------------ double tap and zoom controls
await step('double tap on an empty area zooms in, again zooms back', async () => {
  await fresh();
  await page.addStyleTag({ content: '.fill-handle, .toast, .scroll-ind { visibility: hidden !important; }' });
  await place(page, 1, 1299, 169);
  const g = await gridBox(page);
  const p = { x: g.x + g.w * 0.6, y: g.y + g.h * 0.6, id: 0 };
  const dbl = async () => {
    await F.down([p]);
    await F.up([p]);
    await page.waitForTimeout(60);
    await F.down([p]);
    await F.up([p]);
    await settled(page, 500);
  };
  await dbl();
  const z1 = await ev(() => window.__view.viewInfo().zoom);
  const editing = await ev(() => document.activeElement?.getAttribute('data-testid'));
  await dbl();
  const z2 = await ev(() => window.__view.viewInfo().zoom);
  console.log('     zoom', 1, '→', z1, '→', z2, 'focus', editing);
  assert.ok(z1 >= 1.9, 'double tap zooms in: ' + z1);
  assert.ok(Math.abs(z2 - 1) < 0.02, 'second double tap zooms back: ' + z2);
  assert.notEqual(editing, 'formula-input', 'double tap on empty space does not start editing');
});
await step('double tap on a cell with content still edits it', async () => {
  await place(page, 1, 0, 0);
  await ev(() => (window.__doc.wb.getActiveSheet().getRange('B2').setValue('x'), 0));
  await page.waitForTimeout(200);
  const pos = await ev(() => {
    const V = window.__view;
    const v = V.viewInfo();
    const r = document.querySelector('[data-testid=grid]').getBoundingClientRect();
    // the middle of B2 (its corners carry the selection handles)
    const p = V.toScreen(v, (v.colAcc[0] + v.colAcc[1]) / 2, (v.rowAcc[0] + v.rowAcc[1]) / 2);
    return { x: r.left + p.x, y: r.top + p.y, id: 0 };
  });
  await F.down([pos]);
  await F.up([pos]);
  await page.waitForTimeout(60);
  await F.down([pos]);
  await F.up([pos]);
  await page.waitForTimeout(300);
  const focus = await ev(() => document.activeElement?.getAttribute('data-testid'));
  const z = await ev(() => window.__view.viewInfo().zoom);
  await page.keyboard.press('Escape');
  assert.equal(focus, 'formula-input');
  assert.equal(z, 1);
});
await step('zoom presets and slider in the menu; the zoom pill shows the zoom', async () => {
  await page.addStyleTag({ content: '.zoom-pill { visibility: visible !important; }' });
  await place(page, 1, 0, 0);
  await page.click('[data-testid=menu-more]');
  await page.click('[data-testid=menu-zoom]');
  for (const pct of [50, 75, 150, 200, 100]) {
    await page.click(`[data-testid=zoom-preset-${pct}]`);
    await settled(page, 300);
    const z = await ev(() => window.__view.viewInfo().zoom);
    assert.ok(Math.abs(z - pct / 100) < 0.005, `preset ${pct}: ${z}`);
  }
  await page.locator('[data-testid=zoom-slider]').fill('30');
  await settled(page, 300);
  const z = await ev(() => window.__view.viewInfo().zoom);
  assert.ok(Math.abs(z - 0.3) < 0.005, 'slider 30 %: ' + z);
  await page.click('[data-testid=zoom-close]');
  await place(page, 1, 0, 0);
  const g = await gridBox(page);
  const c = { x: g.x + g.w / 2, y: g.y + g.h / 2 };
  await F.pinch(c, 200, 300, { steps: 8, release: false });
  const pill = await page.textContent('[data-testid=zoom-pill]');
  await F.up();
  await settled(page);
  assert.match(pill, /^1[45]\d %$/);
});

await step('Settings → Developer → Performance overlay shows fps, frame time, tiles pending, cache and zoom level', async () => {
  await closeMenus(page);
  await page.click('[data-testid=menu-more]');
  await page.click('[data-testid=menu-settings]');
  await page.check('[data-testid=set-perf-overlay]');
  await closeMenus(page);
  const g = await gridBox(page);
  await F.pinch({ x: g.x + g.w / 2, y: g.y + g.h / 2 }, 200, 300, { steps: 8 });
  await settled(page, 700);
  const text = await page.textContent('[data-testid=perf-overlay]');
  console.log('     ' + text.replace(/\n/g, ' | '));
  for (const re of [/\d+ fps/, /frame [\d.]+ ms/, /tiles pending \d+/, /cache \d+ tiles · \d+\/\d+ MB/, /zoom \d+ % · level [\d.]+ %/]) assert.match(text, re);
  await page.click('[data-testid=menu-more]');
  await page.click('[data-testid=menu-settings]');
  await page.uncheck('[data-testid=set-perf-overlay]');
  await closeMenus(page);
  assert.equal(await page.$('[data-testid=perf-overlay]'), null, 'switched off again');
});

await step('idle: Univer asks for no frames (no copies of its canvas); a change is drawn at once', async () => {
  await closeMenus(page);
  await place(page, 1, 0, 0);
  await settled(page, 800);
  // count the animation frames the page asks for in one idle second (1.2.1: 120, two render loops)
  const frames = await ev(async () => {
    let n = 0;
    const raf = window.requestAnimationFrame;
    window.requestAnimationFrame = (cb) => (n++, raf.call(window, cb));
    await new Promise((r) => setTimeout(r, 1000));
    window.requestAnimationFrame = raf;
    return n;
  });
  console.log('     animation frames asked for in an idle second:', frames);
  assert.ok(frames <= 2, 'idle frames ' + frames);
  // a value set without any touch must appear within 150 ms (the loop wakes up by itself)
  const g = await gridBox(page);
  const cell = { x: g.x + 60, y: g.y + 30, width: 80, height: 22 };
  const before = (await page.screenshot({ clip: cell })).toString('base64');
  await ev(() => window.__doc.api.getActiveWorkbook().getActiveSheet().getRange(0, 0).setValue('WAKE 123'));
  await page.waitForTimeout(150);
  const after = (await page.screenshot({ clip: cell })).toString('base64');
  assert.notEqual(after, before, 'A1 redrawn within 150 ms');
  await ev(() => window.__doc.api.getActiveWorkbook().getActiveSheet().getRange(0, 0).setValue(''));
});

fs.mkdirSync('tests/e2e/out', { recursive: true });
fs.writeFileSync('tests/e2e/out/pinch.json', JSON.stringify(report, null, 1) + '\n');
await t.finish();
