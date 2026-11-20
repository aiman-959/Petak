# Pinch zoom: before and after (1.2.1 → 1.3.0)

Written by `tests/perf/pinch-report.mjs` from `pinch-before.json` (Petak 1.2.1, commit 7ee8669) and
`pinch-after.json` (this build), both measured by `tests/perf/pinch.mjs` on the same machine.

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

What the traces of this work found (`Canvas2DResourceProvider::ProduceCanvasResource` in a Chrome
trace): **every 2D canvas on screen is copied at every frame the browser commits, changed or not.**
During a pinch that was every visible tile (10–33 copies per frame, 1.5 s of every 4 s at Mid) and, in
1.2.1 too, Univer's own full-screen canvas under the picture. Tiles are now ImageBitmaps in
`bitmaprenderer` canvases (no copy) and Univer's canvas is hidden while the opaque gesture surface
covers it: 0–1 copies per frame. The same trace showed the idle app committing 60 frames a second,
each with a copy of Univer's canvas, because Univer's render loops never stop (P-04); they now pause
when nothing is dirty.

1.2.1's numbers are for a broken picture: during every pinch it showed black areas and a second,
misaligned grid (see `tests/e2e/pinch.mjs`: 3 / 47 checks passed on 1.2.1, all 50 now).

## With the screen's vsync (60 Hz here)

| Tier | File | fps | dropped % | worst frame ms | busy ms/s | hand-off ms | canvas MB |
|---|---|---|---|---|---|---|---|
| High (÷1) | empty FN1300 | 59.3 → **59.6** | 0.7 → **0.8** | 67 → **33** | 195 → **261** | 47 → **33** | 13.5 → **82.2** |
| High (÷1) | roster | 58.7 → **58.8** | 1.6 → **1.7** | 67 → **50** | 242 → **304** | 182 → **53** | 16.3 → **92.1** |
| High (÷1) | 50k rows | 59.6 → **59.2** | 0.7 → **1.3** | 33 → **33** | 175 → **255** | 50 → **45** | 13.5 → **75.7** |
| Mid (CPU ÷4) | empty FN1300 | 50.8 → **48.9** | 5.5 → **5.9** | 183 → **217** | 551 → **576** | 143 → **94** | 7.9 → **18.3** |
| Mid (CPU ÷4) | roster | 43.2 → **46.1** | 12.2 → **6.9** | 233 → **233** | 608 → **714** | 92 → **194** | 9.5 → **63.5** |
| Mid (CPU ÷4) | 50k rows | 50 → **49.9** | 5.2 → **6.2** | 200 → **183** | 537 → **575** | 131 → **106** | 7.9 → **56.4** |
| Low (CPU ÷6) | empty FN1300 | 47.8 → **46.2** | 7.1 → **7.3** | 217 → **183** | 564 → **704** | 163 → **124** | 4.4 → **12.2** |
| Low (CPU ÷6) | roster | 43.3 → **40.3** | 12.1 → **14.4** | 217 → **300** | 572 → **839** | 104 → **208** | 5.3 → **31.9** |
| Low (CPU ÷6) | 50k rows | 48.2 → **40.7** | 6.4 → **17** | 200 → **267** | 579 → **804** | 149 → **189** | 4.4 → **32.2** |

## Uncapped (`--disable-frame-rate-limit`): how many frames per second the work allows

| Tier | File | fps | dropped % | busy ms/s | hand-off ms |
|---|---|---|---|---|---|
| High (÷1) | empty FN1300 | 860.4 → **398.1** | 49.4 → **30.5** | 1017 → **591** | 21 → **31** |
| High (÷1) | roster | 719.5 → **574.1** | 55.2 → **68.3** | 1019 → **809** | 97 → **53** |
| High (÷1) | 50k rows | 854.8 → **342.7** | 49 → **37.7** | 1016 → **543** | 21 → **41** |
| Mid (CPU ÷4) | empty FN1300 | 174.5 → **117.6** | 50 → **42.4** | 1001 → **1005** | 79 → **107** |
| Mid (CPU ÷4) | roster | 150.5 → **86.5** | 45.5 → **50.3** | 1019 → **1004** | 88 → **167** |
| Mid (CPU ÷4) | 50k rows | 172.3 → **137.6** | 49.6 → **40.6** | 1013 → **971** | 82 → **110** |
| Low (CPU ÷6) | empty FN1300 | 147.8 → **86.9** | 47.4 → **57.1** | 1011 → **1005** | 91 → **110** |
| Low (CPU ÷6) | roster | 123.8 → **61.3** | 52.5 → **75.7** | 1024 → **1017** | 117 → **224** |
| Low (CPU ÷6) | 50k rows | 132.1 → **61.8** | 48.1 → **50.3** | 1014 → **1014** | 157 → **199** |
