/**
 * Shared helpers for the pinch-zoom tests (tests/e2e/pinch.mjs) and benchmark (tests/perf/pinch.mjs).
 * Only uses what every build exposes (window.__doc, window.__view, real touch events), so the same
 * checks run on the old and the new renderer (before / after numbers).
 */
import ExcelJS from 'exceljs';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';

/** 50,000 rows × 20 columns with a formula per row (same file as tests/perf/phase5.mjs). */
export async function makeRows(rows = 50000, cols = 20) {
  const file = path.join(os.tmpdir(), `perf-${rows}x${cols}.xlsx`);
  if (fs.existsSync(file)) return file;
  const wb = new ExcelJS.stream.xlsx.WorkbookWriter({ filename: file, useStyles: true });
  const ws = wb.addWorksheet('Data');
  ws.addRow(Array.from({ length: cols }, (_, i) => (i === 0 ? 'ID' : i === 1 ? 'Name' : i === 2 ? 'Region' : i === 3 ? 'Date' : i === cols - 1 ? 'Total' : `Metric ${i - 3}`))).commit();
  const regions = ['North', 'South', 'East', 'West'];
  for (let r = 1; r <= rows; r++) {
    const row = [r, `Item ${r}`, regions[r % 4], new Date(Date.UTC(2026, r % 12, (r % 28) + 1))];
    for (let c = 4; c < cols - 1; c++) row.push(Math.round(((r * 7919 + c * 104729) % 100000) / 10) / 10);
    row.push({ formula: `SUM(E${r + 1}:S${r + 1})` });
    ws.addRow(row).commit();
  }
  ws.commit();
  await wb.commit();
  return file;
}

export const ROSTER = path.resolve('tests/fixtures/files/roster.xlsx');

/** Open a file through the Open button (answers the unsaved-changes question). */
export async function openFile(page, file, rows = 0) {
  await closeMenus(page);
  const before = await page.evaluate(() => window.__doc.wb.getId());
  const [chooser] = await Promise.all([
    page.waitForEvent('filechooser'),
    (async () => {
      await page.click('[data-testid=open]');
      const ask = await page.waitForSelector('[data-testid=ask-discard], [data-testid=confirm-ok]', { timeout: 1000 }).catch(() => null);
      if (ask) await ask.click();
    })(),
  ]);
  await chooser.setFiles(file);
  await page.waitForFunction(
    ([n, id, rows]) => window.__doc.snapshot.file.name === n && window.__doc.wb.getId() !== id && window.__doc.wb.getActiveSheet().getLastRow() >= rows - 1,
    [path.basename(file), before, rows],
    { timeout: 600000, polling: 100 },
  );
  await page.evaluate(() => window.__doc.bridge.idle());
  await page.waitForTimeout(500);
}

/** Close bottom sheets, menus and dialogs (Android Back, as the user would). */
export async function closeMenus(page) {
  for (let i = 0; i < 4; i++) {
    const open = await page.$('.sheet-backdrop, .dialog, .bsheet');
    if (!open) return;
    await page.evaluate(() => window.__petakBack?.());
    await page.waitForTimeout(150);
  }
}

/** Wait until no gesture is running and the grid has settled. */
export async function settled(page, ms = 450) {
  await page.waitForFunction(() => !window.__doc.gesturing && !document.querySelector('.grid-wrap.gesturing'), null, { timeout: 15000 });
  await page.waitForTimeout(ms);
}

/** Put the grid at zoom `z` with cell (row, col) (0-based) at the top-left of the scrolling pane. */
export async function place(page, z, row = 0, col = 0) {
  await page.evaluate(
    async ([z, row, col]) => {
      const d = window.__doc;
      const V = window.__view;
      const wb = d.wb;
      const ws = wb.getActiveSheet();
      d.ensureSize(ws.getSheetId(), row + 200, col + 30);
      const frame = () => new Promise((r) => requestAnimationFrame(() => r(0)));
      // right after a page load Univer's zoom controller may not be registered yet: retry
      for (let i = 0; i < 40; i++) {
        const ok = await Promise.resolve(d.api.executeCommand('sheet.operation.set-zoom-ratio', { zoomRatio: z, unitId: wb.getId(), subUnitId: ws.getSheetId() })).catch(() => false);
        if (ok) break;
        await frame();
      }
      for (let i = 0; i < 30; i++) {
        await frame();
        const v = V.viewInfo();
        if (v && Math.abs(v.zoom - z) < 1e-6 && v.colAcc.length > col && v.rowAcc.length > row) break;
      }
      for (let i = 0; i < 6; i++) {
        const v = V.viewInfo();
        const fx = v.frozenW / v.zoom;
        const fy = v.frozenH / v.zoom;
        const x = Math.max(col > 0 ? v.colAcc[col - 1] : 0, fx);
        const y = Math.max(row > 0 ? v.rowAcc[row - 1] : 0, fy);
        V.scrollTo(v, x, y);
        await frame();
        await frame();
      }
      return 0;
    },
    [z, row, col],
  );
  await page.waitForTimeout(400);
}

export async function gridBox(page) {
  return page.evaluate(() => {
    const r = document.querySelector('[data-testid=grid]').getBoundingClientRect();
    const v = window.__view.viewInfo();
    return { x: r.left, y: r.top, w: r.width, h: r.height, zoom: v.zoom, headerW: v.headerW, headerH: v.headerH, frozenW: v.frozenW, frozenH: v.frozenH };
  });
}

/** Two fingers on a diagonal through `c`, `d` px apart. */
const diag = (c, d) => [
  { x: c.x - d / 2 / Math.SQRT2, y: c.y - d / 2 / Math.SQRT2, id: 0 },
  { x: c.x + d / 2 / Math.SQRT2, y: c.y + d / 2 / Math.SQRT2, id: 1 },
];

/** CDP touch driver with explicit finger ids (touchEnd releases only the listed fingers). */
export async function fingers(page) {
  const cdp = await page.context().newCDPSession(page);
  const pt = (p) => ({ x: p.x, y: p.y, id: p.id, radiusX: 4, radiusY: 4, force: 1 });
  return {
    cdp,
    down: (pts) => cdp.send('Input.dispatchTouchEvent', { type: 'touchStart', touchPoints: pts.map(pt) }),
    move: (pts) => cdp.send('Input.dispatchTouchEvent', { type: 'touchMove', touchPoints: pts.map(pt) }),
    up: (pts) => cdp.send('Input.dispatchTouchEvent', { type: 'touchEnd', touchPoints: (pts ?? []).map(pt) }),
    /** pinch from span d0 to d1 around c; the fingers stay down unless `release` */
    async pinch(c, d0, d1, { steps = 16, stepMs = 16, release = true, hold = 0 } = {}) {
      await this.down(diag(c, d0));
      for (let i = 1; i <= steps; i++) {
        await this.move(diag(c, d0 + ((d1 - d0) * i) / steps));
        if (stepMs) await page.waitForTimeout(stepMs);
      }
      if (hold) await page.waitForTimeout(hold);
      if (release) await this.up();
      return diag(c, d1);
    },
  };
}

/**
 * Picture analysis in the page (screenshot → pixels). Regions are in CSS px of the page.
 *  black:  share of 12×12 device-pixel blocks that are entirely near-black (never content: text strokes are thinner)
 *  diff:   mean absolute difference (0–255) of two shots after 4× box downscaling, per region
 */
export async function analyse(page, shots, regions) {
  return page.evaluate(
    async ([shots, regions]) => {
      const load = async (b64) => {
        const img = await createImageBitmap(await (await fetch('data:image/png;base64,' + b64)).blob());
        const c = new OffscreenCanvas(img.width, img.height);
        const ctx = c.getContext('2d', { willReadFrequently: true });
        ctx.drawImage(img, 0, 0);
        return { ctx, w: img.width, h: img.height, k: img.width / innerWidth };
      };
      const imgs = await Promise.all(shots.map(load));
      const out = {};
      for (const [name, r] of Object.entries(regions)) {
        if (r.w < 4 || r.h < 4) continue;
        const res = { black: [], dark: [] };
        const datas = imgs.map((im) => {
          const x = Math.round(r.x * im.k);
          const y = Math.round(r.y * im.k);
          const w = Math.max(1, Math.round(r.w * im.k));
          const h = Math.max(1, Math.round(r.h * im.k));
          return { d: im.ctx.getImageData(x, y, w, h).data, w, h };
        });
        for (const { d, w, h } of datas) {
          const B = 12;
          let blocks = 0;
          let black = 0;
          let dark = 0;
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
          for (let i = 0; i < d.length; i += 16) if (d[i] + d[i + 1] + d[i + 2] < 60) dark++;
          res.black.push(+((black / Math.max(1, blocks)) * 100).toFixed(2));
          res.dark.push(+((dark / (d.length / 16)) * 100).toFixed(2));
        }
        if (datas.length >= 2) {
          // 4× box downscale, then mean absolute difference over RGB
          const small = ({ d, w, h }) => {
            const W = Math.floor(w / 4);
            const H = Math.floor(h / 4);
            const o = new Float32Array(W * H * 3);
            for (let y = 0; y < H * 4; y++)
              for (let x = 0; x < W * 4; x++) {
                const i = (y * w + x) * 4;
                const j = (Math.floor(y / 4) * W + Math.floor(x / 4)) * 3;
                o[j] += d[i] / 16;
                o[j + 1] += d[i + 1] / 16;
                o[j + 2] += d[i + 2] / 16;
              }
            return o;
          };
          const a = small(datas[0]);
          const b = small(datas[1]);
          let s = 0;
          for (let i = 0; i < a.length; i++) s += Math.abs(a[i] - b[i]);
          res.diff = +(s / Math.max(1, a.length)).toFixed(2);
        }
        out[name] = res;
      }
      return out;
    },
    [shots, regions],
  );
}

/** Screen regions of the grid at the current zoom: headers, frozen panes, main pane. */
export function regionsOf(g) {
  const r = {
    colHeader: { x: g.x + g.headerW + 2, y: g.y + 1, w: g.w - g.headerW - 14, h: g.headerH - 2 },
    rowHeader: { x: g.x + 1, y: g.y + g.headerH + 2, w: g.headerW - 2, h: g.h - g.headerH - 14 },
    main: { x: g.x + g.headerW + g.frozenW + 2, y: g.y + g.headerH + g.frozenH + 2, w: g.w - g.headerW - g.frozenW - 14, h: g.h - g.headerH - g.frozenH - 14 },
    whole: { x: g.x, y: g.y, w: g.w, h: g.h },
  };
  if (g.frozenH > 4) r.frozenTop = { x: g.x + g.headerW + g.frozenW + 2, y: g.y + g.headerH + 1, w: g.w - g.headerW - g.frozenW - 14, h: g.frozenH - 2 };
  if (g.frozenW > 4) r.frozenLeft = { x: g.x + g.headerW + 1, y: g.y + g.headerH + g.frozenH + 2, w: g.frozenW - 2, h: g.h - g.headerH - g.frozenH - 14 };
  return r;
}
