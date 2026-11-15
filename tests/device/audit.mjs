/**
 * Phase 1 audit on a real Android WebView (emulator in CI, or a phone over USB).
 * Observes behaviour only; it never fails the build. Output: device-out/*.png + device-report.json.
 *   node tests/device/audit.mjs <debug.apk> <fixtures dir>
 */
import { adb, sh, sleep, attach, launch, screencap, calibrate, keyboardShown, report, PKG } from './lib.mjs';
import fs from 'node:fs';
import path from 'node:path';

const [apk, fixtures = 'tests/fixtures/files'] = process.argv.slice(2);
const R = report();
if (apk) adb('install', '-r', '-g', apk);
sh('mkdir -p /sdcard/Download');
for (const f of fs.readdirSync(fixtures)) if (fs.statSync(path.join(fixtures, f)).isFile()) adb('push', path.join(fixtures, f), '/sdcard/Download/' + f);
sh(`am force-stop ${PKG}`);
launch();
await sleep(6000);
const { page } = await attach();
await page.waitForFunction(() => window.__doc?.snapshot?.sheets?.length > 0, null, { timeout: 90000 });
await sleep(1500);
screencap('launch');

await R.check('user agent', () => page.evaluate(() => navigator.userAgent));
await R.check('devicePixelRatio / viewport', () => page.evaluate(() => [window.devicePixelRatio, innerWidth, innerHeight]));
await R.check('tier', () => page.evaluate(() => [document.documentElement.dataset.tier, navigator.deviceMemory, navigator.hardwareConcurrency]));
await R.check('crossOriginIsolated / shared mode', () => page.evaluate(() => [crossOriginIsolated, window.__governor?.stats?.shared]));
await R.check('fonts available (Calibri, Carlito, Arial, Liberation Sans, Roboto, Noto Sans CJK SC)', () =>
  page.evaluate(() => ['Calibri', 'Carlito', 'Arial', 'Liberation Sans', 'Roboto', 'Noto Sans CJK SC', 'Times New Roman', 'Courier New'].map((f) => {
    const c = document.createElement('canvas').getContext('2d');
    c.font = '40px monospace';
    const base = c.measureText('mmmmmmmmlli').width;
    c.font = `40px "${f}", monospace`;
    return [f, c.measureText('mmmmmmmmlli').width !== base];
  })),
);

const cal = await R.check('touch calibration', async () => {
  const c = await calibrate(page);
  return c;
});
if (!cal) {
  R.save();
  process.exit(0);
}
const grid = await page.evaluate(() => {
  const r = document.querySelector('[data-testid=grid]').getBoundingClientRect();
  return { x: r.x, y: r.y, w: r.width, h: r.height };
});
const cell = (r, c) => [grid.x + 46 + 88 * (c + 0.5), grid.y + 20 + 24 * (r + 0.5)];
const focus = () => page.evaluate(() => { const a = document.activeElement; return { tag: a?.tagName, id: a?.id, editable: !!a && (a.isContentEditable || a.tagName === 'INPUT' || a.tagName === 'TEXTAREA') }; });
const sel = () => page.evaluate(() => { const r = window.__doc.wb.getActiveSheet().getActiveRange()?.getRange(); return r ? [r.startRow, r.startColumn, r.endRow, r.endColumn] : null; });

// ---- problem 12: single tap and the keyboard
sh('input keyevent 111'); // ESC, in case calibration taps opened something
await sleep(500);
cal.tap(...cell(6, 1));
await sleep(1200);
await R.check('single tap: selection / focus / keyboard shown', async () => [await sel(), await focus(), keyboardShown()]);
screencap('single-tap');
cal.tap(...cell(9, 2));
await sleep(1200);
await R.check('second single tap: keyboard shown', async () => [await sel(), await focus(), keyboardShown()]);
cal.doubleTap(...cell(4, 1));
await sleep(1500);
await R.check('double tap: focus / keyboard shown', async () => [await focus(), keyboardShown()]);
sh('input text hello');
await sleep(500);
screencap('double-tap-typing');
sh('input keyevent 66');
await sleep(1000);
await R.check('typed "hello" + Enter after double tap -> B5', () => page.evaluate(() => window.__doc.wb.getActiveSheet().getRange('B5').getValue()));
sh('input keyevent 4'); // back: hide keyboard
await sleep(800);
cal.tap(...cell(12, 1));
await sleep(1200);
await R.check('single tap after an edit: keyboard shown', async () => [await focus(), keyboardShown()]);
if (keyboardShown()) sh('input keyevent 4');
await sleep(500);

// ---- drawing with real touches (problems 1-4)
const drawShapes = () => page.evaluate(() => (window.__doc.meta.drawings?.[window.__doc.wb.getActiveSheet().getSheetId()]?.shapes ?? []).map((s) => ({ t: s.type, x: Math.round(s.x), y: Math.round(s.y), w: Math.round(s.w), h: Math.round(s.h), n: s.points?.length })));
await R.check('draw: line by press-drag-release', async () => {
  await page.click('[data-testid=tab-draw]');
  await page.click('text=Line');
  await page.waitForSelector('[data-testid=draw-layer]');
  await sleep(500);
  cal.drag(grid.x + 100, grid.y + 240, grid.x + 260, grid.y + 320, 700);
  await sleep(800);
  return drawShapes();
});
screencap('draw-line');
await R.check('draw: rectangle by drag (start mid-cell)', async () => {
  await page.click('[data-testid=draw-tool-rect]');
  cal.drag(grid.x + 130, grid.y + 360, grid.x + 270, grid.y + 450, 700);
  await sleep(800);
  return [await drawShapes(), await page.evaluate(() => document.querySelector('.dbtn.on[data-testid^=draw-tool]')?.dataset.testid)];
});
await R.check('draw: tap the rectangle to select it', async () => {
  cal.tap(grid.x + 131, grid.y + 405);
  await sleep(800);
  return [await page.evaluate(() => !!document.querySelector('[data-testid=draw-props]')), await drawShapes()];
});
await R.check('draw: polyline 3 taps then double tap (expect a 4-point polyline, Done button gone)', async () => {
  // deselect first: the properties card of the selected rectangle covers the lower part of the sheet
  await page.evaluate(() => window.__draw.select([]));
  await page.click('[data-testid=draw-tool-polyline]');
  // inside the free drawing area (below the tool bar, above the chips)
  cal.tap(grid.x + 80, grid.y + 130);
  await sleep(600);
  cal.tap(grid.x + 160, grid.y + 180);
  await sleep(600);
  cal.tap(grid.x + 240, grid.y + 130);
  await sleep(700);
  cal.doubleTap(grid.x + 300, grid.y + 200);
  await sleep(1000);
  return [await drawShapes(), await page.isVisible('[data-testid=draw-finish]'), await page.evaluate(() => document.querySelector('[data-testid=draw-mode-chip]')?.textContent)];
});
await R.check('draw: Back button (keyevent 4) leaves drawing mode step by step', async () => {
  const before = await page.evaluate(() => [window.__draw.get().selected.length, window.__draw.get().active]);
  sh('input keyevent 4');
  await sleep(600);
  const mid = await page.evaluate(() => [window.__draw.get().selected.length, window.__draw.get().active]);
  sh('input keyevent 4');
  await sleep(600);
  return [before, mid, await page.evaluate(() => [window.__draw.get().selected.length, window.__draw.get().active, document.visibilityState])];
});
screencap('draw-shapes');
await R.check('Back closes an open dialog instead of leaving the app', async () => {
  await page.click('[data-testid=menu-more]');
  await sleep(500);
  const open = await page.evaluate(() => !!document.querySelector('.bsheet'));
  sh('input keyevent 4');
  await sleep(700);
  return [open, await page.evaluate(() => [!!document.querySelector('.bsheet'), document.visibilityState])];
});
await R.check('after Back: is the app still in front? (activity, process, crash log)', async () => {
  await sleep(2500);
  const safe = (f) => { try { return String(f()).trim().split('\n').slice(-12).join(' | ').slice(0, 1500); } catch (e) { return 'ERR ' + String(e.message).slice(0, 200); } };
  return {
    resumed: safe(() => sh(`dumpsys activity activities | grep -E "ResumedActivity|mFocusedApp" | head -4`)),
    pid: safe(() => sh(`pidof ${PKG}`)),
    page: await page.evaluate(() => document.visibilityState).catch((e) => 'page: ' + String(e.message).slice(0, 80)),
    log: safe(() => sh(`logcat -d -t 400 | grep -iE "FATAL|AndroidRuntime|petak.*(crash|died|kill)|Force finishing|am_finish|WIN DEATH|chromium.*(crash|renderer)|onBackPressed|finishing" | tail -12`)),
  };
});
if (await page.evaluate(() => window.__draw.get().active).catch(() => false)) await page.click('[data-testid=draw-close]').catch(() => {});

// ---- files: "Open with" from another app (C-04), a second file while one is open (K-07), cold start.
// The system picker is not driven here: `uiautomator dump` and Playwright's device driver both
// knock out the page connection on CI. It was checked by hand-driven taps in run 23.
const mediaId = async (name) => {
  sh(`am broadcast -a android.intent.action.MEDIA_SCANNER_SCAN_FILE -d file:///sdcard/Download/${name} >/dev/null`);
  for (let i = 0; i < 6; i++) {
    await sleep(1200);
    const q = sh(`content query --uri content://media/external/file --projection _id:_display_name --where "_display_name='${name}'"`);
    const id = /_id=(\d+)/.exec(q)?.[1];
    if (id) return id;
  }
  return null;
};
const openWith = async (name, mime) => {
  const id = await mediaId(name);
  if (!id) return 'no MediaStore row for ' + name;
  const out = sh(`am start -W -a android.intent.action.VIEW -d content://media/external/file/${id} -t ${mime} --grant-read-uri-permission -n ${PKG}/com.zetl.petak.MainActivity`);
  await sleep(1500);
  // unsaved changes in the open workbook: Save / Don't save / Cancel
  const asked = await page.$('[data-testid=ask-discard]').catch(() => null);
  const question = asked ? await page.textContent('[data-testid=ask-text]').catch(() => '') : null;
  if (asked) await asked.click();
  await sleep(6000);
  const file = await page.evaluate(() => [window.__doc.snapshot.file.name, window.__doc.snapshot.sheets.map((s) => s.name), document.querySelector('.toast')?.textContent ?? null]).catch((e) => 'page: ' + e.message);
  const res = { start: out.trim().split('\n').slice(-1)[0], asked: question, file };
  if (typeof file === 'string') {
    // the page is gone: restarted process? crashed renderer? killed for memory?
    const safe = (f) => { try { return String(f()).trim().split('\n').slice(-16).join(' | ').slice(0, 3000); } catch (err) { return 'ERR ' + String(err.message).slice(0, 160); } };
    res.pid = safe(() => sh(`pidof ${PKG}`));
    res.mem = safe(() => sh(`dumpsys meminfo ${PKG} | grep -E "TOTAL PSS|TOTAL:" | head -2`));
    res.log = safe(() => sh(`logcat -d -t 4000 | grep -iE "FATAL|AndroidRuntime: |Caused by|chromium.*(crash|kill|fatal|oom)|Renderer|render process|lowmemory|lmkd|am_kill|am_proc_died|Killing |has died|OutOfMemory|WebView.*(crash|gone)|aw_browser|cr_Child" | grep -v "lmkd.*init" | tail -16`));
  }
  return res;
};
const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
await R.check('"Open with" while the app is running (roster.xlsx; the open workbook has unsaved changes)', () => openWith('roster.xlsx', XLSX_MIME));
screencap('roster');
await R.check('"Open with" a second file (qm-report.xlsx): K-07', () => openWith('qm-report.xlsx', XLSX_MIME));
screencap('qm-report');
await R.check('after 2nd open: toast / dialogs / page errors', () => page.evaluate(() => [document.querySelector('.toast')?.textContent, document.querySelector('[data-testid=ask-text]')?.textContent ?? null]));

// scrolling with real swipes on the frozen-pane file (problems 5, 8, 11)
await R.check('swipe scroll: frame times', async () => {
  await page.evaluate(() => { window.__fr = []; let last = performance.now(); window.__m = true; const f = (t) => { window.__fr.push(t - last); last = t; if (window.__m) requestAnimationFrame(f); }; requestAnimationFrame(f); });
  cal.swipe(grid.x + 200, grid.y + 500, grid.x + 200, grid.y + 150, 400);
  await sleep(1500);
  const fr = await page.evaluate(() => { window.__m = false; return window.__fr.slice(1); });
  const avg = fr.reduce((a, b) => a + b, 0) / fr.length;
  return { frames: fr.length, avgMs: +avg.toFixed(1), fps: Math.round(1000 / avg), over33ms: fr.filter((d) => d > 33).length, worstMs: Math.round(Math.max(...fr)), scroll: await page.evaluate(() => window.__doc.wb.getActiveSheet().getScrollState()) };
});
screencap('qm-scrolled');
cal.swipe(grid.x + 200, grid.y + 300, grid.x + 200, grid.y + 330, 300);
await sleep(1200);
screencap('qm-scrolled-2');

// old .xls through "Open with"
await R.check('"Open with" legacy.xls (old format, with formatting)', async () => {
  const r = await openWith('legacy.xls', 'application/vnd.ms-excel');
  const look = await page.evaluate(() => { const ws = window.__doc.wb.getActiveSheet(); const st = ws.getRange('A1').getCellStyleData(); return { bg: st?.bg?.rgb, fs: st?.fs, colB: ws.getColumnWidth(1) }; }).catch((e) => 'page: ' + e.message);
  return { ...(typeof r === 'string' ? { r } : r), look };
});
screencap('open-with-xls');
await R.check('Petak is offered for a content:// .xls (count of matching lines, 0 = not offered)', () => sh('cmd package query-activities -a android.intent.action.VIEW -d content://media/external/file/1 -t application/vnd.ms-excel | grep -c petak || true').trim());
await R.check('cold start by "Open with" (the app is not running): people.csv', async () => {
  const id = await mediaId('people.csv');
  if (!id) return 'no MediaStore row';
  sh(`am force-stop ${PKG}`);
  await sleep(1500);
  sh(`am start -W -a android.intent.action.VIEW -d content://media/external/file/${id} -t text/csv --grant-read-uri-permission -n ${PKG}/com.zetl.petak.MainActivity`);
  await sleep(8000);
  const { page: p2 } = await attach();
  await sleep(2000);
  // the restored session may have unsaved changes
  const asked = await p2.$('[data-testid=ask-discard]').catch(() => null);
  if (asked) await asked.click();
  await sleep(5000);
  return p2.evaluate(() => [window.__doc.snapshot.file.name, window.__doc.wb.getActiveSheet().getRange('A1').getValue()]);
});
screencap('cold-open-with');
await R.check('why Android killed Petak during this run (empty = it never did)', () => sh(`logcat -d | grep -E "Killing [0-9]+:${PKG.replace(/\./g, '\\.')}[/: ]" | tail -4 || true`).trim().slice(0, 1500));
await R.check('provider connections of Petak into other apps', () => sh(`dumpsys activity providers | grep -B12 "${PKG}" | grep -E "ContentProviderRecord|-> .*${PKG}|${PKG} *->|authority=" | head -14 || true`).trim().slice(0, 2500));
R.save();
console.log('device audit finished');
process.exit(0);
