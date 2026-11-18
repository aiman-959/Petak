# Petak 1.2.0 — test report

What was tested for this release, how, and what came out. Every bug found has an ID in
[AUDIT.md](AUDIT.md); every number below comes from a script in `tests/` that can be run again
(commands at the end).

## Summary

| Area | Result |
|---|---|
| Unit tests (Vitest, 25 files) | **158 / 158** |
| End-to-end tests on a phone screen (Playwright, 14 suites) | **204 / 204** steps |
| Excel functions, each with a worked example | **476 / 476** give Excel's result (437 equal LibreOffice; 26 equal Excel's documented value where LibreOffice lacks the function; 6 volatile ones in range; 7 where LibreOffice differs from Excel) |
| Menu commands with an end-to-end test | **153 / 153** ([FEATURES.md](FEATURES.md)) |
| Sample files: open → edit → save → reopen unchanged, LibreOffice recalculates the saved file to the same values | **8 / 8** (xlsx from Excel and LibreOffice, xlsm, xls, ods, csv) |
| Twenty files opened in a row | no memory growth, no errors, still edits and computes |
| Bugs found in this round (Phases 4–5) | 7 (F-01 … F-07), all fixed with a test |
| Bugs found overall (Phases 1–5) | 7 Critical, 12 Known, 7 High, 10 Medium, 2 Low, F-01 … F-07. All fixed except K-05 (frame rate still to be measured on the phone) and K-07 (mitigated: Android ending the app is outside Petak's control) |

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
| `roundtrip.mjs` | open → edit → save → reopen, LibreOffice import and recalculation | 9 |
| `taps.mjs` | touches needed for everyday tasks | 13 |
| **Total** | | **204** |

Unit tests (`tests/unit`, 158) cover the formula engine and all 476 functions
(`function-check.test.ts`), the xlsx reader / writer and lossless save, `.xls` formatting, CSV,
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
| roster.ods (made by LibreOffice) | 2 | 141 | 14 |
| people.csv | 1 | 20 | 1 |

For each file: (1) the values Petak loads equal LibreOffice's reading of the same file; (2) after
one edit, Save as .xlsx and reopening, every cell's value, formula, font, fill, colour, number
format, alignment and borders, the merges, frozen panes, column widths and row heights are as before
plus the edit; (3) LibreOffice opens Petak's file, recalculates every formula and gets the same
values. The one expected difference: the CSV identifier `0012` stays text in Petak (L-01).
`.ods` files open with values, formulas and merges but without fonts, fills, borders or column widths (hence the
lower cell count above: the roster's 24 blank formatted cells are not read), L-03 in AUDIT.md.

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

## What a test server cannot tell

These need a real phone (the target is a Snapdragon 8s Gen 3 with a 120 Hz screen); the debug APK
has WebView debugging on, so `chrome://inspect` → Performance gives the exact numbers.

1. **Frame rate on the device (K-05).** The test server has no GPU, so canvas raster runs on the
   CPU; under ÷4 / ÷6 slowdown that is more than half of every frame. Check: open the 50,000-row
   file, fling for 10 s, pinch for 10 s, and look for dropped frames at 120 Hz.
2. **Thermal behaviour.** Ten minutes of scrolling on battery: the load governor should step down
   (About → Live performance shows the level).
3. **Clipboard with other apps.** Copy in Petak, paste in another app and back. Inside Petak this
   works even without clipboard access (F-02).
4. **Play services restart (K-07).** Open a file while Google Play services updates: Petak restarts
   with that file (seen on the emulator; not reproducible on demand).
5. **The system file picker and share sheet** (not automated on the emulator).

## Running the tests

```bash
npm ci
npx tsc -b && npx vitest run                     # unit tests
npx vite build && npx vite preview --port 4173 &
export BASE_URL=http://localhost:4173/
npm run test:e2e                                 # all 14 e2e suites (about 8 minutes)
node tests/e2e/feature-matrix.mjs                # FEATURES.md (after commands.mjs)
node tests/functions/libreoffice.mjs             # refresh LibreOffice's function results (needs soffice)
UPDATE_DOCS=1 npx vitest run tests/unit/function-check.test.ts   # FUNCTION-CHECK.md
node tests/perf/phase5.mjs                       # PHASE5.md (about 15 minutes)
node tests/perf/perf50k.mjs                      # RESULTS.md
```
