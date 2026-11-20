# Petak QA: progress and handover

Last updated: 2026-11-20 · Branch `claude/petak-pinch-zoom-audit-ygv8st` (from `claude/practical-davinci-plplss`) · App version `1.3.0`

This file is the handover for the next work session. `AUDIT.md` holds every bug with its ID; this
file says where the work stands.

## Release 1.3.0 brief (current)

| Part | State |
|---|---|
| 1. Pinch zoom (Critical, P-01) | **Done.** New gesture renderer (`src/grid/gestureView.ts`, `tiles.ts`, `camera.ts`); P-01…P-04 fixed with tests (`tests/e2e/pinch.mjs`, `tests/unit/camera.test.ts`); before / after numbers in `tests/perf/PINCH.md`. Frame rate on the real phone: hand check (TEST-REPORT.md). |
| 2. Full re-audit | **Done.** Every suite again (17 suites 277 / 277, unit 200 / 200), LibreOffice function check re-run (476 / 476), new `daily.mjs` (a working day by touch, timed), `a11y.mjs` (44 dp, WCAG AA), Excel opening-rules validator on every save, long session. Found and fixed F-09 (High, fill grip), F-10, F-11 (230 small targets), F-12 (contrast). |
| 3. Release signing, 1.3.0 build | Guide `docs/RELEASE-SIGNING.md`; CI checks the key, signs release *and* debug APKs with it, refuses tag releases without it, prints the certificate (`SIGNING.txt`, job summary); version 1.3.0 (versionCode 5), About shows it. **Waiting for the user:** the four secrets, then "Run workflow". Hand checks: TEST-REPORT.md. |

### Part 1 in short (details: AUDIT.md P-01…P-04)

- **What was wrong (1.2.1):** the pinch picture was a copy of Univer's canvas in *opaque* canvases
  (empty cells black), scaled as one picture together with its headers (a second grid over the real
  one), nothing new while zooming out, minimum 40 %.
- **Now:** one camera; map-style tile levels (12.5 … 400 % + the resting zoom) drawn by Univer's own
  print path; every pane (headers, frozen rows / columns / corner) placed by the same camera; nearest
  cached level, the start picture or plain grid lines while a tile is missing; 10–400 %; zoom pill,
  zoom panel (presets 50–200 %, slider), double tap on empty space, Settings → Performance overlay.
- **Performance findings:** Chrome copies every visible 2D canvas at every frame commit, so tiles are
  ImageBitmaps in `bitmaprenderer` canvases and Univer's canvas is hidden under the surface; Univer's
  render loops never stopped (P-04, 60–120 frames a second while idle), now paused when nothing is dirty.
- **Code map:** `camera.ts` (pure maths, unit-tested), `tiles.ts` (TileStore: render, placeholders,
  bitmaps, eviction), `gestureView.ts` (surface, panes, layers, frame loop, hand-off, prefetch, pill,
  overlay), `input.ts` (gestures → view), `view.ts` (`scrollTo` for inactive panes, `setUniverZoom`,
  `quietRenderLoop`), `ui/ZoomControls.tsx`, `ui/settings.ts`.

## Earlier work (QA plan for 1.2, done)

| Phase | State |
|---|---|
| 1. Full audit | **Done.** `AUDIT.md` written (7 Critical, 12 Known, 7 High, Medium/Low, emulator findings). |
| 2. Known problems | **Done** except device fps (K-05), measured in Phase 5 on the emulator / headless only. |
| 3. Acceptance criteria | **Done.** M-01…M-10, L-01, L-02 fixed with tests (`phase3.mjs`, commit `cdfc0c0`). |
| 4. Every function and feature | **Done.** 476 functions vs LibreOffice / Excel (FUNCTION-CHECK.md, F-01: 11 fixed); 153 menu commands each pressed by a test (FEATURES.md, `commands.mjs`); round trips of every sample file with LibreOffice recalculation (`roundtrip.mjs`); touches for everyday tasks (`taps.mjs`). Found and fixed F-02…F-06. |
| 5. Performance and stability | **Done.** `tests/perf/PHASE5.md`, `tests/perf/RESULTS.md`. Found and fixed F-07. |
| 6. Report and build | **Done.** 1.2.0 / 1.2.1 (L-03 `.ods` formatting, F-08). |

## Fixed and proven by tests

| ID | Problem | Proof |
|---|---|---|
| C-01 | Saving destroyed pictures, links, charts, VBA, etc. | Lossless save: the original xlsx package is patched, not rebuilt. Unit round-trip tests + `features.mjs` |
| C-02 | Typed text thrown away | `cells.mjs` |
| C-03 | Tapping a cell while typing a formula discarded it | `cells.mjs` |
| C-05 | Named ranges not loaded | unit tests |
| C-06 | Table structured references not understood | unit tests |
| C-07 | 1900 dates one day off | unit tests |
| K-05 | Scroll and pinch stutter | 1.3.0: gesture renderer (`gestureView.ts`), `pinch.mjs`, `tests/perf/PINCH.md`. **Device fps: hand check** |
| P-01 | Pinch zoom: second grid, black areas | `pinch.mjs` (50 steps), `camera.test.ts`, emulator `e2e-pinch.mjs` |
| P-02 | Frozen columns wider than the screen stopped scrolling | `pinch.mjs` roster 200 / 400 % |
| P-03 | Zoom right after a workbook appears was lost | `pinch.mjs` (five refusals) |
| P-04 | Idle app drew 60–120 frames a second | `pinch.mjs` "idle" (0 frames; 1.2.1: 120) |
| F-09 | Fill grip stayed on the old cell after opening a file (overwrote cells) | `daily.mjs` (1.2.1 fails) |
| F-10 | Escape did not close the keyboard after Enter | `daily.mjs` (1.2.1 fails) |
| F-11 | 230 controls under 44 dp | `a11y.mjs` |
| F-12 | File menu title invisible; zoom preset contrast | `a11y.mjs` |
| K-09 | Hyperlinks loaded as plain text | `cells2.mjs` (load, Open link, insert, edit, remove, save back) |
| K-10 | Pictures did not load | `cells2.mjs` (load, select, move, resize, delete, full-screen view, insert, save) |
| K-12 | Single tap opened the keyboard | `cells.mjs` (tap selects; double tap or formula bar edits) |
| H-01 | No keyboard shortcuts | `cells.mjs` |
| H-02 | No long-press menu or clipboard | `cells.mjs`, `cells2.mjs` |
| H-04 | Autofilter not loaded | unit + e2e |
| H-05 | No touch resize / autofit / hide / unhide | `cells.mjs` |
| M-02 | No Sum / Average / Count status | `cells2.mjs` |
| K-01 | Shapes can't be selected or deleted after drawing | `draw-touch.mjs`, `part3.mjs`; emulator |
| K-02 | Polyline never finishes | `draw-touch.mjs`; emulator (4-point polyline by taps + double tap) |
| K-03 | Line tool is tap-to-tap | `draw-touch.mjs`; emulator (line by press–drag–release) |
| K-04 | Nothing snaps to cell lines | `draw.test.ts`, `draw-touch.mjs` |
| K-06 | Excel fonts display wrongly | `fonts.mjs`, `fonts.test.ts` |
| K-08 | Borders faded / missing at zoom ≠ 100 % | `render.mjs` (screenshot measured at 60–130 %) |
| K-11 | Frozen panes overlap | no longer reproduces; `render.mjs` |
| C-04 | "Open with Petak" | emulator: running app, second file, `.xls`, cold start |
| K-07 | Opening another file does nothing | cause found (Android ends the app when Play services restarts); session saved 0.4 s after opening; `files.mjs` |
| H-03 | `.xls` without formatting | `biff.test.ts`, `files.mjs` |
| H-06 | Unsaved prompt can't save | `files.mjs` |
| H-07 | Rich text flattened | `richtext.test.ts`, `files.mjs` |
| M-01 | No fill handle | `cells.mjs`, `fill.test.ts` |
| — | Quick tap on the next row started editing | `cells.mjs` (double tap needs the same cell) |
| M-01…M-10, L-01, L-02 | Phase 3 acceptance items | `phase3.mjs`, unit tests |
| F-01 | 11 functions wrong (INT, CEILING.MATH, FLOOR.MATH, INDEX, VALUE, ISREF, 5 Excel 2007 distributions) | `function-check.test.ts`, `functions.test.ts` |
| F-02 | Paste threw without clipboard permission; Univer permission banner | `commands.mjs` (runs without the permission) |
| F-03 | Tap ignored after a row-header menu → long-press menu acted on old cells | `commands.mjs` |
| F-04 | Sort moved a text column's heading | `taps.mjs`, `phase3.test.ts` |
| F-05 | Typing into a Text cell converted `0012` / `=1+1` | `commands.mjs` |
| F-06 | .xls / .ods / csv saved as .xlsx lost default font and column width | `roundtrip.mjs`, `ooxml-roundtrip.test.ts` |
| F-07 | Univer command stack grew with every tap | `commands.mjs` |
| L-03 | `.ods` opened without fonts, fills, borders, number formats, widths | `ods.test.ts`, `roundtrip.mjs` |
| F-08 | New / .xls / .ods / CSV workbook with a currency or date format and a font saved a broken `styles.xml` | `ooxml-roundtrip.test.ts`, `roundtrip.mjs` (well-formed XML check) |

Also new: recent files, clear messages when a file cannot be opened, Format → Font & size with
your own font files, Android Back button handled by the page.

### Last full test run (release 1.3.0)

| Suite | Result |
|---|---|
| Unit (Vitest, 28 files) | 200 / 200 |
| `npm run test:e2e` (17 suites, `tests/e2e/out/summary.json`) | 277 / 277 |
| Emulator `e2e-pinch.mjs` (screen captures) | 16 / 16 |

Emulator (CI `device` job, Android 14 WebView): line, rectangle (fit to cells, tool back to
Select), tap-select, polyline by double tap and Back out of drawing mode all work.
Emulator, files: "Open with" works with the app running (asks about unsaved changes first), for a
second file, for the old `.xls` (with formatting) and from a cold start. When the page vanished on
the second file (about one run in three) Android had ended Petak because the system WebView holds a
connection to Play services' font provider during its first seconds and Play services restarted
(details under K-07 in `AUDIT.md`). Not switchable from the app; mitigated by saving the session
0.4 s after a file opens. The device script re-attaches after such a restart and reports what the
restarted app shows (`afterRestart`).

Device-script lessons: `adb shell uiautomator dump` and Playwright's `device.wait/tap` (driver apk
not installed) both kill the page connection ("Target page, context or browser has been closed").
The script therefore opens files only through intents; the system picker is not automated.

## Still open

| ID | Problem | State |
|---|---|---|
| K-05 | Scroll / pinch speed on the phone | 1.3.0 renderer measured on the server (`tests/perf/PINCH.md`); 120 Hz on the Snapdragon 8s Gen 3 needs the hand check with Settings → Performance overlay (TEST-REPORT.md) |
| K-07 | Android ends the app when Play services restarts | Mitigated (session saved 0.4 s after opening); not fixable in the app |

## Drawing tools rework (K-01 to K-04): done

What was built (details in `AUDIT.md` under each ID):

- **Input.** `GridInput` has `feedDown / feedMove / feedUp / feedCancel` (and `owns`). The draw
  overlay keeps shape gestures and hands pans, flings, pinches and cell taps to the grid; its own
  `scrollBy` pan is gone, the mouse wheel is forwarded to Univer's canvas. Fed touches never open
  menus or the editor. A pointer the WebView cancels mid-shape continues on raw touch events.
- **Select (default).** Handles (22 px finger radius), drag selected shape = move, tap = select
  (16 px), drag elsewhere = scroll, tap empty = select cell and deselect, two fingers = pinch.
  Mouse keeps the marquee. Unfilled shapes are picked by their inside only in drawing mode.
- **Draw.** Press–drag–release with preview; back to Select after one shape unless the tool was
  long-pressed (or double-clicked); second finger within 250 ms cancels and zooms, later it
  turns snapping off. Mode chip `draw-mode-chip` with `chip-magnet`, `chip-fit`,
  `chip-undo-point`, `chip-cancel`, `draw-finish` (Done).
- **Polyline.** Points on release, double tap / first point (`closed`) / Done, Undo point, Back cancels.
- **Snapping.** `SNAP_PX = 12` screen px; scoring and fallbacks as designed; `snaps.enabled`
  (magnet), `snaps.angle15`; cell centres on by default; per-axis edge vs centre uses the same
  priority penalty. Locked-angle arc, "60° ✓" labels, "48.5 mm · 30°" live label.
- **Angle dimension.** Tool `angle`: drag the first arm then tap/drag the second (snaps relative to
  the first arm), or tap two lines (`refs`, follows them; `refreshAngles` on every commit; a
  deleted line leaves a plain angle). Also "Measure the angle" in the floating bar for two lines.
  Rendered, and exported to SVG / DXF / PDF; closed polylines too.
- **Selection UI.** Floating bar (`shape-delete`, `shape-duplicate`, `shape-front`, `shape-back`,
  `shape-lock`, `shape-angle`), only in Select; length / angle fields for lines (`prop-length`,
  `prop-angle`); shapes listed per layer (`layer-shape-<id>`). Locked shapes: selectable, not
  movable / resizable / deletable, fields disabled.
- **Back.** `src/ui/back.ts` (`onBack`, `window.__petakBack`); `MainActivity` asks it first.
  Drawing: gesture → polyline → picked lines → panel → tool → selection → close drawing mode.

## Next, in order

1. The user adds the four signing secrets (`docs/RELEASE-SIGNING.md`) and runs the workflow; the
   `Petak-apks` artifact of that run is the 1.3.0 release (`SIGNING.txt` shows the release key). To
   publish a GitHub Release, push tag `v1.3.0` from a machine that may push tags (the cloud session
   cannot). Until then: CI run 111 (commit `2420352`), artifact `Petak-apks`
   https://github.com/aiman-959/Petak/actions/runs/44793835141/artifacts/14658555883 —
   `Petak-v1.3.0.apk` 6.1 MB (versionName 1.3.0, versionCode 5), per-ABI APKs, debug APK,
   SHA256SUMS.txt, SIGNING.txt, SIZE.md; signed with that runner's throwaway debug key (no secrets yet).
2. Hand checks on the phone (TEST-REPORT.md, "checks by hand"), above all 120 Hz pinching with the
   performance overlay.
3. Possible next work: conditional formats, comments and pictures from `.ods` / `.xls` files.

## How to work in this repo

```bash
npm ci
npx tsc -b                      # type check
npx vitest run                  # unit tests
npx vite build                  # rebuild dist after every source edit
npx vite preview --port 4173 &  # e2e suites run against this
export BASE_URL=http://localhost:4173/
npm run test:e2e                # every e2e suite (tests/e2e/run-all.mjs); one suite: node tests/e2e/<name>.mjs
node tests/e2e/feature-matrix.mjs                    # FEATURES.md, after commands.mjs
node tests/functions/libreoffice.mjs                 # LibreOffice's function results (needs soffice)
UPDATE_DOCS=1 npx vitest run tests/unit/function-check.test.ts   # FUNCTION-CHECK.md
node tests/perf/phase5.mjs && node tests/perf/perf50k.mjs        # PHASE5.md, RESULTS.md
LABEL=after node tests/perf/pinch.mjs && node tests/perf/pinch-report.mjs   # PINCH.md (pinch-before.json: 1.2.1)
```

- Touch gestures in tests: `tests/e2e/touch.mjs` (`tap, doubleTap, longPress, drag, fling, pinch`). On the 412 px phone profile use column D
- Fixtures: `tests/fixtures/make-fixtures.mjs` → `tests/fixtures/files/*` (roster, QM report, xlsm, xls, csv);
  `tests/fixtures/make-ods.mjs` → `looks.ods` and LibreOffice's own `looks-lo.xlsx` of it (needs soffice)
- Emulator: the `device` job in `.github/workflows/android.yml`
- LibreOffice headless (`soffice`) is the reference for formula comparison

## Things that bit us (don't repeat)

- Overlays must be appended to `.grid-wrap`; Univer re-renders its own container
- Do not call `removeScrollBar()` (breaks scrolling); silence its `render` instead
- Univer's hidden in-cell editor must stay blocked on touch (it caused the keyboard pop-up and dead undo)
- With frozen panes, scroll distance is measured from the end of the frozen area
- Pinch end needs the `settle` loop before the sharp redraw
- TypeScript `erasableSyntaxOnly`: no constructor parameter properties
- An open bottom sheet's backdrop covers the header buttons; close the panel before opening a picker
- When a Playwright `waitForEvent` times out, check which call it came from and look at the screenshot first
- `useDraw` / `useSyncExternalStore` selectors must return primitives or stable objects: a selector
  that builds a new object each call loops forever (React error #185, blank page)
- On the phone profile the drawing area shows sheet y ≈ 62–440 (tool bar above, chips and the
  properties card below): taps outside it land on other UI
- With a drawing tool kept on, the floating selection bar must stay hidden or it covers the next shape's start
- Never `await` between disposing the old workbook and creating the new one (UI code reads `doc.wb`)
- Never use `local()` fonts or ask Android for fonts: the WebView's connection to Play services' font provider gets the app killed (K-07); the substitutes are bundled and registered under the Excel names
- The grid canvas is stretched on Low / Mid tiers (`canvasStretch`): measure rendering on the screenshot, not on the canvas
- Device script: never `adb shell uiautomator dump` or Playwright `device.tap/wait` (both cut the page connection); open files through intents
- Device tests: taps must land between the drawing tool bar and the chips (sheet y ≈ 62–440)
- Never draw on one canvas with both a raw 2D context and Univer's context wrapper: the wrapper caches
  state (transform), so Univer's next drawing lands in the wrong place (lost header labels)
- A 2D canvas on screen costs a copy at every frame the browser commits; show finished pictures as
  ImageBitmaps (`bitmaprenderer`), hide what is covered (`visibility: hidden`; opacity 0 does not help)
- `transferFromImageBitmap(null)` allocates a blank full-size bitmap; shrink the element to free one
- A hidden canvas is no touch target: `GridInput.isGridTarget` accepts its container during a gesture
- Univer's main pane goes inactive (negative width) when frozen columns are wider than the screen and
  then ignores scrolling: `scrollTo` sets the panes' scroll itself (P-02)
- `set-zoom-ratio` throws until Univer's zoom controller exists (a moment after a workbook appears):
  use `setUniverZoom`, which reports a refusal
- Perf numbers here come from a GPU-less server (software raster and compositing): compare builds
  on the same machine, never with a phone

## Constraints

- Commit and push only to the session's branch (now `claude/petak-pinch-zoom-audit-ygv8st`); no pull request unless asked
- Tag pushes are refused from the cloud session, so the APK link is the GitHub Actions artifact unless a release is created by hand
- Never commit signing keys
- The repo was renamed to `aiman-959/Petak`; pushes print "This repository moved" and still succeed
- Univer only pops a command from its execution stack when it succeeds: a command cancelled through `BeforeCommandExecute` must be removed by hand (F-07)
- Measure memory after a full garbage collection (a heap snapshot); detached canvases and file inputs seen by DevTools are not necessarily leaks
- Playwright's file chooser keeps the hidden `<input type=file>` alive: DOM node counts in browser tests grow by itself
- LibreOffice's own .xlsx is a good reference for an .ods but not perfect: it writes theme colours instead of
  some cells' font colours, and fitted row heights as 13.8 pt. Check such cells against how the file was made
- SheetJS's .ods values for dates and times are rounded to the second and dates before March 1900 are a day
  short; `ods.ts` reads the exact values
- A big .ods holds ~140 MB of XML for 50 000 rows: never build an element per cell (`children()`); scan once
- Petak's own reader tolerates broken XML that Excel refuses: check saved packages are well-formed (F-08)
