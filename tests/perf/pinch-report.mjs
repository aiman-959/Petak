/**
 * tests/perf/PINCH.md from pinch-before.json (Petak 1.2.1) and pinch-after.json (this build).
 *   node tests/perf/pinch-report.mjs
 */
import fs from 'node:fs';
import path from 'node:path';

const HERE = path.dirname(new URL(import.meta.url).pathname);
const before = JSON.parse(fs.readFileSync(path.join(HERE, 'pinch-before.json'), 'utf8'));
const after = JSON.parse(fs.readFileSync(path.join(HERE, 'pinch-after.json'), 'utf8'));
const A = new Map(after.map((r) => [`${r.mode}|${r.tier}|${r.file}`, r]));
const tierName = { high: 'High (÷1)', mid: 'Mid (CPU ÷4)', low: 'Low (CPU ÷6)' };
const arrow = (a, b, unit = '') => (a === undefined || b === undefined || a === null || b === null ? '–' : `${a}${unit} → **${b}${unit}**`);
const rows = (mode, cols) =>
  before
    .filter((r) => r.mode === mode && r.fps !== undefined)
    .map((r) => {
      const n = A.get(`${r.mode}|${r.tier}|${r.file}`) ?? {};
      return `| ${tierName[r.tier]} | ${r.file} | ${cols.map(([k, u]) => arrow(r[k], n[k], u)).join(' | ')} |`;
    })
    .join('\n');

const md = `# Pinch zoom: before and after (1.2.1 → 1.3.0)

Written by \`tests/perf/pinch-report.mjs\` from \`pinch-before.json\` (Petak 1.2.1, commit 7ee8669) and
\`pinch-after.json\` (this build), both measured by \`tests/perf/pinch.mjs\` on the same machine.

**How:** headless Chromium, Pixel 7 screen (412 × 915, DPR 2.625), real two-finger touch events; for each
tier the CPU is slowed (÷1 / ÷4 / ÷6 stands in for High / Mid / Low phones) and the file is pinched in and
out (fingers 140 ↔ 380 px, focal point wandering) for 6 s while every animation frame is timed.
*Dropped* = frames that took longer than 1.5 vsync intervals. *Busy* = main-thread time per second.
*Hand-off* = finger lifted → the sharp grid is on screen (median). *Canvas MB* = all canvases at the peak,
including the tile cache (capped at 96 / 56 / 28 MB for High / Mid / Low).

**Read this first: this server has no GPU.** Chromium here runs "Canvas: Software only, Compositing:
Software only" (chrome://gpu); a phone's WebView rasterises and composites on the GPU. Every canvas
pixel costs main-thread time here, so these numbers are a worst case for the new renderer, which draws
tiles (1.2.1 only scaled one picture, the wrong one). What they show reliably: how much work runs per
frame, and how it compares between the two builds on the same machine.

What the traces of this work found (\`Canvas2DResourceProvider::ProduceCanvasResource\` in a Chrome
trace): **every 2D canvas on screen is copied at every frame the browser commits, changed or not.**
During a pinch that was every visible tile (10–33 copies per frame, 1.5 s of every 4 s at Mid) and, in
1.2.1 too, Univer's own full-screen canvas under the picture. Tiles are now ImageBitmaps in
\`bitmaprenderer\` canvases (no copy) and Univer's canvas is hidden while the opaque gesture surface
covers it: 0–1 copies per frame. The same trace showed the idle app committing 60 frames a second,
each with a copy of Univer's canvas, because Univer's render loops never stop (P-04); they now pause
when nothing is dirty.

1.2.1's numbers are for a broken picture: during every pinch it showed black areas and a second,
misaligned grid (see \`tests/e2e/pinch.mjs\`: 3 / 47 checks passed on 1.2.1, all 50 now).

## With the screen's vsync (60 Hz here)

| Tier | File | fps | dropped % | worst frame ms | busy ms/s | hand-off ms | canvas MB |
|---|---|---|---|---|---|---|---|
${rows('vsync', [['fps'], ['dropped'], ['worst'], ['busyMsPerS'], ['handoffP50'], ['peakCanvasMb']])}

## Uncapped (\`--disable-frame-rate-limit\`): how many frames per second the work allows

| Tier | File | fps | dropped % | busy ms/s | hand-off ms |
|---|---|---|---|---|---|
${rows('uncapped', [['fps'], ['dropped'], ['busyMsPerS'], ['handoffP50']])}
`;
fs.writeFileSync(path.join(HERE, 'PINCH.md'), md);
console.log('wrote', path.join(HERE, 'PINCH.md'));
