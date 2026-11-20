# Petak 1.3.0 — test report

What was tested for this release, how, and what came out. Every bug found has an ID in
[AUDIT.md](AUDIT.md); every number below comes from a script in `tests/` that can be run again
(commands at the end).

## Summary

| Area | Result |
|---|---|
| Unit tests (Vitest, 28 files) | **200 / 200** |
| End-to-end tests on a phone screen (Playwright, 17 suites) | **277 / 277** steps |
| Pinch zoom while the fingers rest, 10–400 %, three files (`pinch.mjs`) | **50 / 50** (1.2.1: 3 / 47) |
| Pinch zoom on the Android emulator's real WebView, judged on screen captures | **16 / 16** |
| A working day with the roster by touch only, timed (`daily.mjs`) | **12 / 12** after fixing F-09, F-10 |
| Controls at least 44 × 44 dp, text at WCAG AA, light and dark (`a11y.mjs`) | **0** below (was 230 controls, 3 texts: F-11, F-12) |
| Saved files against Excel's opening rules (`excel-rules.ts`: XML, parts, relationships, element order, styles, merges, names) | every sample and every Petak save: **0 problems** |
| Excel functions, each with a worked example | **476 / 476** give Excel's result (LibreOffice 24.2.7 re-run today: same reference values) |
| Sample files: open → edit → save → reopen unchanged, LibreOffice recalculates | **9 / 9** (xlsx from Excel and LibreOffice, xlsm, xls, two ods, csv) |
| Long session (`longsession.mjs`, 120 rounds) | LONG_SESSION |
| Findings in 1.3.0 | P-01 (Critical), P-02, P-04, F-09 (High), P-03, F-11, F-12 (Medium), F-10 (Low): **all fixed**, each with its test |
| Findings overall | 55 (47 up to 1.2.1, 8 in 1.3.0); **53 fixed**, each with its test; K-05 (frame rate on the phone: hand check) and K-07 (Android ends the app when Play services restarts: mitigated) as before |

## 1.3.0, Part 1: pinch zoom

| Check | 1.2.1 | 1.3.0 |
|---|---|---|
| `tests/e2e/pinch.mjs`: screenshots while two fingers rest on the glass at 10 / 25 / 50 / 100 / 200 / 400 %, pinching in and out, on an empty sheet at FN1300, the staff roster and a 50,000-row file; fast, slow, ten in a row, pinch → scroll → pinch, double tap, presets, slider, overlay, idle | **3 / 47** (black up to 92.8 % of the grid in 33 of 38 pinches, a second grid in all 38, 10 zooms unreachable) | **50 / 50** |
| Solid-black 12 × 12 blocks during a pinch | up to 92.8 % | none beyond the settled grid's own |
| Difference from the settled grid while the fingers rest (0–255) | 100–250 | under 16 per region |
| Zoom range | 40–400 % | 10–400 % |
| Frames the idle app asks for per second | 120 | 0 |
| `tests/unit/camera.test.ts` (point under the fingers stays put, limits, levels) | – | 13 / 13 |

Frame rate, from `tests/perf/PINCH.md` (CPU slowed ÷1 / ÷4 / ÷6 for High / Mid / Low phones, a server
**without a GPU**, so canvas pixels are drawn by the CPU; a phone draws them on its GPU):

| Tier (60 Hz screen) | fps, 1.2.1 → 1.3.0 | dropped frames % | finger lifted → sharp grid, ms | frames/s the work allows (uncapped) |
|---|---|---|---|---|
| High | 58.7–59.6 → **58.8–59.6** | 0.7–1.6 → **0.8–1.7** | 47–182 → **33–53** | 719.5–860.4 → **342.7–574.1** |
| Mid (CPU ÷4) | 43.2–50.8 → **46.1–49.9** | 5.2–12.2 → **5.9–6.9** | 92–143 → **94–194** | 150.5–174.5 → **86.5–137.6** |
| Low (CPU ÷6) | 43.3–48.2 → **40.3–46.2** | 6.4–12.1 → **7.3–17** | 104–163 → **124–208** | 123.8–147.8 → **61.3–86.9** |

(Ranges over the three files. 1.2.1's figures are for its broken picture: a scaled copy of one
screenshot. 1.3.0 draws the real sheet at every zoom, which costs more per frame on a CPU without a
GPU; at High it still leaves room for 340–570 frames a second, far above the 120 a 120 Hz screen needs.)

The phone itself (Snapdragon 8s Gen 3, 120 Hz) is a hand check below: Settings → Developer →
Performance overlay shows fps, frame time, tiles waiting, the tile cache and the zoom level live.

## 1.3.0, Part 2: re-audit

Everything above was run again on 1.3.0: the 14 earlier suites (206 steps, every earlier finding,
the 45 fixed ones included), the unit tests, the function check against a fresh LibreOffice run, and
the round trips. New for this release:

**A working day on a phone** (`tests/e2e/daily.mjs`): the roster, touches only, each step timed from
the touch to its result on this server (a phone with a GPU is faster on the drawing steps):

| Step | ms |
|---|---:|
| open roster.xlsx (Open → pick → grid shown) | 1711 |
| double tap B5 → keyboard | 273 |
| Enter → saved | 109 |
| drag the fill grip down 3 rows | 781 |
| Format → Bold | 808 |
| Data → Sort A→Z | 875 |
| Data → Filter on | 810 |
| pinch out (fingers 300 → 120 px) | 988 |
| pinch back in | 966 |
| Insert → Picture → pick | 647 |
| press, drag, release a rectangle | 996 |
| File → Save .xlsx | 386 |
| reopen the saved file | 1591 |

What it found: the fill grip did not follow taps after a file was opened, so a quick drag filled from
the cell selected before and overwrote two names (F-09, High; in 1.2.1 too); Escape did not close the
keyboard after Enter (F-10). Both fixed; 1.2.1 fails both checks. Nothing slow (every step under 2 s,
most under 1 s); unlike Excel only in small ways already listed in the README (Petak keeps the
keyboard after Enter for the next cell, as Excel for Android does).

**Touch size and contrast** (`tests/e2e/a11y.mjs`): main screen, six tool panels, File menu, eleven
dialogs, the zoom panel and the long-press menus, light and dark. 230 controls were under 44 dp
(F-11) and three texts under WCAG AA, one of them invisible (F-12). After the fixes: none.

**Saved files and Excel** (`tests/unit/excel-rules.ts`): a strict check of what makes Excel say a file
is damaged (well-formed XML, content types, relationships, the order of a worksheet's elements, style
counts and indexes, merged cells, sheet names, defined names). It is run on every sample and on every
file Petak saves in the unit tests (unchanged and edited, from .xlsx, .xlsm, .xls, .ods and .csv): no
problems. Its self-tests prove it catches each kind of damage.

**Long session** (`tests/perf/longsession.mjs`): LONG_DETAIL

## How it was tested

* **Phone-sized, touch-driven.** The e2e suites run the production build in Chromium with a Pixel 7
  profile (412 × 915, DPR 2.6). Touches are real touch events sent through the DevTools protocol
  (tap, double tap, long press, drag, fling, pinch), not clicks; text is typed key by key.
* **Android emulator.** The CI `device` job installs the debug APK on an Android 14 emulator
  (WebView 113) and drives it with `adb` taps and intents (`tests/device/audit.mjs`): cell editing,
  drawing, "Open with", a second file, an old `.xls`, a cold start, the Back button.
* **Reference results.** Function results are compared with LibreOffice 24.2 run headless on the
  same workbook, and with Microsoft's documented examples where LibreOffice lacks a function. Round
  trips are checked by LibreOffice opening and recalculating the files Petak saves.

## Test suites

| Suite | What it covers | Steps |
|---|---|---:|
| `cells.mjs` | tap / double tap / long press, typing, formula references by tapping, numeric keyboard, scrolling, fling, pinch, header resize and autofit, fill handle | 16 |
| `cells2.mjs` | keyboard, clipboard, long-press menu, status bar, hyperlinks, pictures | 12 |
| `draw-touch.mjs` | drawing mode by touch: shapes, polyline, snapping, angles, layers, locking, Back (K-01…K-04) | 24 |
| `part2.mjs` | comments, notes, validation, sparklines, goal seek, scenarios, data tables, slicers, more charts, URL import, data model | 16 |
| `part3.mjs` | drawing precision tools on a tablet: units, dimensions, guides, align, groups, SVG / DXF / PDF export | 18 |
| `part4.mjs` | device tiers, load governor, thermal / battery levels, fallback without shared memory | 9 |
| `fonts.mjs` | Excel fonts and their substitutes, your own fonts (K-06) | 8 |
| `files.mjs` | unsaved-changes question, damaged / empty files, recent files, Back, restart after Android ends the app (K-07), `.xls`, rich text | 12 |
| `render.mjs` | borders and grid lines at every zoom, frozen panes (K-08, K-11) | 10 |
| `features.mjs` | the original core features: xlsm with VBA, csv, formulas, sheets, freeze, sort, filter, find, formatting, charts, pivot, undo, autosave, macros, transforms | 19 |
| `phase3.mjs` | coloured references, error help, sheet rename / move, CSV notices, grouping, print / PDF (M-03…M-10, L-01, L-02) | 10 |
| `commands.mjs` | every menu command not covered elsewhere, run **without clipboard permission** | 28 |
| `roundtrip.mjs` | open → edit → save → reopen, LibreOffice import and recalculation, well-formed XML in every saved file, `.ods` looks | 11 |
| `taps.mjs` | touches needed for everyday tasks | 13 |
| `pinch.mjs` | pinch zoom at 10–400 % on three files, fast / slow / repeated, pinch → scroll, double tap, zoom panel, overlay, idle frames (1.3.0) | 50 |
| `daily.mjs` | a working day with the roster by touch, timed (1.3.0) | 12 |
| `a11y.mjs` | 44 dp touch targets and WCAG AA contrast everywhere, light and dark (1.3.0) | 9 |
| **Total** | | **277** |

Unit tests (`tests/unit`, 168) cover the formula engine and all 476 functions
(`function-check.test.ts`), the xlsx reader / writer and lossless save, `.xls` and `.ods` formatting, CSV,
rich text, fonts, drawing geometry, the fill handle, PDF layout, the governor, transforms, macros and
the data model.

## Functions

Each of the 476 functions Petak offers gets one example on a shared data block
(`tests/functions/examples.mjs`). Petak's result is compared with LibreOffice's for the same
formula; where LibreOffice 24.2 has no such function, with Microsoft's documented example. Eleven
were wrong (HyperFormula's INT, CEILING.MATH, FLOOR.MATH, INDEX, VALUE, ISREF and the Excel 2007
NORMSDIST, LOGNORMDIST, HYPGEOMDIST, NEGBINOMDIST, BETADIST) and now give Excel's results (F-01).
Seven differ from LibreOffice because LibreOffice differs from Excel; each is explained in
[FUNCTION-CHECK.md](FUNCTION-CHECK.md), which also lists all 476 with both results.

## Round trips

| File | Sheets | Cells | Formulas |
|---|---:|---:|---:|
| roster.xlsx (Excel) | 2 | 165 | 14 |
| roster-lo.xlsx (LibreOffice) | 2 | 165 | 14 |
| qm-report.xlsx (Excel) | 2 | 328 | 19 |
| qm-report-lo.xlsx (LibreOffice) | 2 | 328 | 19 |
| macro.xlsm (with VBA) | 2 | 328 | 19 |
| legacy.xls | 2 | 165 | 14 |
| roster.ods (made by LibreOffice) | 2 | 165 | 14 |
| looks.ods (made by LibreOffice) | 2 | 155 | 1 |
| people.csv | 1 | 20 | 1 |

For each file: (1) the values Petak loads equal LibreOffice's reading of the same file; (2) after
one edit, Save as .xlsx and reopening, every cell's value, formula, font, fill, colour, number
format, alignment and borders, the merges, frozen panes, column widths and row heights are as before
plus the edit; (3) LibreOffice opens Petak's file, recalculates every formula and gets the same
values. The one expected difference: the CSV identifier `0012` stays text in Petak (L-01). Every
part of every saved file must be well-formed XML (this found F-08).

`.ods` files keep their looks (L-03): the roster `.ods` LibreOffice made from the Excel roster looks
as the roster does (fonts, fills, colours, borders, alignment, merges, column widths within LibreOffice's
slightly wider character), and `looks.ods` (`tests/fixtures/make-ods.mjs`) looks as LibreOffice's own
`.xlsx` of it, number formats compared by what they show. LibreOffice's `.xlsx` is wrong in two places,
noted in the tests: four cells get the theme colour instead of their own font colour, and rows it fitted to
their content are written at its 13.8 pt standard height.

## Touches for everyday tasks

Done with real touches on a phone screen, result checked (`taps.mjs`). Closing a tool panel to get
back to the sheet counts as a touch.

| Task | Touches | Keys | How |
|---|---:|---:|---|
| Open a file from the phone | 1 | | Open → pick the file in Android’s picker |
| Type a number into a cell | 2 | 3 | double tap → `42` → Enter |
| Add up a column with a SUM formula | 4 | 8 | double tap → `=SUM(` → tap first cell → `:` → tap last cell → `)` → Enter |
| See the total of a column without a formula | 1 | | tap the column header (status bar shows Sum) |
| Make the header row bold | 4 | | tap row header → Format → Bold → close |
| Colour a cell | 4 | | tap cell → Format → colour → close |
| Insert a row | 2 | | long press row header → Row above |
| Sort a list by a column | 4 | | tap a cell → Data → Sort A→Z → close |
| Keep the header row in view | 3 | | Data → Top row → close |
| Make a bar chart of a table | 5 | | tap → drag the selection handle → Insert → Bar chart → Add |
| Copy a cell to another cell | 4 | | long press → Copy → long press → Paste |
| Share the workbook | 2 | | ⋮ → Share → pick the app |

## Performance and stability

Full tables: [tests/perf/PHASE5.md](tests/perf/PHASE5.md) (fling, pinch, CPU split, 20 files) and
[tests/perf/RESULTS.md](tests/perf/RESULTS.md) (50,000-row load, edit, autosave, restore).

Headless Chromium on a 4-core server **without a GPU**, Pixel 7 screen, touch input; the CPU slowed
÷6 / ÷4 / ÷1 stands in for Low / Mid / High phones.

| Tier | File | Open | Fling: fps · frames > 33 ms | Pinch: fps · frames > 33 ms |
|---|---|---:|---:|---:|
| Low ÷6 | 50 rows of 24–72 pt text | 3.1 s | 42.6 · 8.4 % | 48.7 · 2.9 % |
| Low ÷6 | 50,000 rows × 20 columns, 50,000 formulas | 9.1 s | 38.2 · 11.8 % | 46.9 · 3.5 % |
| Mid ÷4 | 50 rows of 24–72 pt text | 2.4 s | 42.7 · 6.7 % | 52.2 · 1.9 % |
| Mid ÷4 | 50,000 × 20 | 8.3 s | 37.6 · 10.7 % | 48.6 · 3.1 % |
| High | 50 rows of 24–72 pt text | 1.3 s | 58.0 · 0.2 % | 59.4 · 0.2 % |
| High | 50,000 × 20 | 3.9 s | 59.6 · 0 % | 59.7 · 0 % |

Before and after this round, same 50,000-row benchmark (`perf50k.mjs`):

| | High before → now | Mid ÷4 before → now |
|---|---|---|
| Open until all formulas are computed | 5.26 → 3.12 s | 8.65 → 6.64 s |
| Edit → dependent formula updated | 61 → 40 ms | 203 → 255 ms |
| Incremental autosave after an edit | 178 → 25 ms | 1036 → 96 ms |
| Scroll frames > 25 ms | 0.4 → 6.4 % | 74.7 → 37.4 % |

(The scroll figures vary between runs on this server: a second run today gave 5.3 % and 31.3 %.
The earlier autosave figure included the whole-session save that now happens once after opening,
239 ms / 1089 ms.)

* Twenty files opened in a row: JS heap, live canvases and renderer memory return to the same level
  for the same file (#10 and #20: 44.8 → 45.2 MB heap, 10 → 10 canvases); no errors; the app still
  edits and computes.
* While looking for leaks F-07 was found and fixed: Univer kept every cancelled editor command on its
  command stack (≈1.4 per tap for the whole session).
* Where the time goes (Mid, 50k rows, per second of fling): 636 ms main thread busy, of which 269 ms
  script and 349 ms raster / compositing that a phone's GPU does. The slow frames are the moments a
  fling reaches cells not drawn yet.

## What a test server cannot tell: checks by hand on the phone (1.3.0)

The target is a Snapdragon 8s Gen 3 with a 120 Hz screen. Turn on Settings → Developer →
**Performance overlay** first (top right of the grid: fps, frame time, tiles waiting, tile cache, zoom).
Before installing: if the Petak on the phone was not signed with the release key, save open work and
uninstall it once (docs/RELEASE-SIGNING.md, step 5).

1. **Install and version.** Install `Petak-v1.3.0.apk`; About shows **1.3.0**. Later, install the next
   build over it without uninstalling: Android must not say "App not installed".
2. **Pinch on the roster (P-01).** Open the staff roster; pinch out slowly to the smallest zoom, then in
   to the largest, without lifting. Look for: never black, never two sets of row numbers or column
   letters, headers stay at the edges, the frozen rows / columns stay in place, the picture follows the
   fingers. The overlay should show about **120 fps** while the fingers move (anything under 60 is a
   finding), and the grid is sharp within a moment of lifting.
3. **Pinch on an empty sheet far away.** New workbook, scroll (or Find) to about FN1300, pinch out and in
   fast, then slowly. Same checks.
4. **Big file.** Open a file with tens of thousands of rows, fling down, pinch out to 10 %, pinch in to
   400 %: content (not just grid lines) appears within about a second while the fingers rest.
5. **Pinch then scroll.** Pinch, lift one finger and keep moving the other: the sheet scrolls without a
   jump; put the second finger back: it zooms again.
6. **Zoom controls.** Tap the zoom pill (bottom left while zooming): presets 50 / 75 / 100 / 150 / 200 %
   and the slider work; double tap an empty area far from the data: 200 %, again: back to 100 %.
7. **Fill grip (F-09).** Open the roster, double tap a name, type, tap ✓, tap a cell in another column,
   drag the round grip below it down two rows: the *tapped* cell is copied, nothing else changes.
8. **Battery and heat (P-04).** Leave Petak open on a sheet for 10 minutes with the screen on, untouched:
   the phone should stay cool; Settings → Battery should show little use for Petak. Then ten minutes of
   scrolling and pinching on battery: the load governor may step down (About → Live performance).
9. **Clipboard with other apps.** Copy in Petak, paste in another app and back.
10. **Play services restart (K-07).** Open a file while Google Play services updates: Petak restarts with
    that file (seen on the emulator; not reproducible on demand).
11. **The system file picker and share sheet** (Open, Save, Share; not automated on the emulator).

The debug APK (`Petak-v1.3.0-debug.apk`, a separate app "Petak debug") has WebView debugging on:
`chrome://inspect` on a computer → Performance gives exact frame timings if the overlay shows a problem.

## Running the tests

```bash
npm ci
npx tsc -b && npx vitest run                     # unit tests
npx vite build && npx vite preview --port 4173 &
export BASE_URL=http://localhost:4173/
npm run test:e2e                                 # all 17 e2e suites (about 12 minutes)
node tests/e2e/feature-matrix.mjs                # FEATURES.md (after commands.mjs)
node tests/functions/libreoffice.mjs             # refresh LibreOffice's function results (needs soffice)
UPDATE_DOCS=1 npx vitest run tests/unit/function-check.test.ts   # FUNCTION-CHECK.md
node tests/perf/phase5.mjs                       # PHASE5.md (about 15 minutes)
node tests/perf/perf50k.mjs                      # RESULTS.md
LABEL=after node tests/perf/pinch.mjs && node tests/perf/pinch-report.mjs   # PINCH.md (about 16 minutes)
node tests/perf/longsession.mjs                  # long session (ROUNDS=120)
```

The emulator scripts (`tests/device/`) run in the CI `device` job on every push: `audit.mjs` (taps,
drawing, "Open with", cold start) and `e2e-pinch.mjs` (pinch zoom judged on screen captures).
