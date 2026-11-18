/**
 * Phase 4: every command in Petak's menus, pressed once on a phone, with its effect checked.
 * Commands whose deeper test lives in another suite are listed in COVERED (suite and step);
 * everything else is pressed here. Writes tests/e2e/out/commands.json, the feature matrix
 * in FEATURES.md is made from it (node tests/e2e/feature-matrix.mjs).
 *   BASE_URL=http://localhost:4173/ node tests/e2e/commands.mjs
 */
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { start, BASE } from './lib.mjs';
import { touch } from './touch.mjs';

// no clipboard permission, as in a WebView that will not let the page read the clipboard:
// Paste then uses what was copied in Petak (cells2.mjs covers the system clipboard)
const t = await start('commands');
const { page, step, ev, idle, val, fresh, select, typeIn, download } = t;
console.log('Petak e2e: every menu command —', BASE);
await fresh();
await page.waitForTimeout(400);
const T = await touch(page);

/** Commands pressed by this suite: "Menu › Command" → step title. */
const pressed = {};
let currentStep = '';
const esc = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
const btn = (scope, label) =>
  page.locator(`${scope} button.tool`).filter({ has: page.locator('span', { hasText: new RegExp(`^${esc(label)}$`) }) }).first();
const closeSheet = async () => {
  if (await page.$('.sheet-backdrop')) {
    await page.mouse.click(200, 120);
    await page.waitForTimeout(150);
  }
};
const MENU = { edit: 'Edit', format: 'Format', insert: 'Insert', draw: 'Draw', data: 'Data', tools: 'Tools' };
/** Press a command in a bottom-bar tab. */
const cmd = async (tab, label, { keepOpen = false } = {}) => {
  await closeSheet();
  await page.click(`[data-testid=tab-${tab}]`);
  await btn('.bsheet', label).click();
  pressed[`${MENU[tab]} › ${label}`] = currentStep;
  await page.waitForTimeout(150);
  if (!keepOpen) await closeSheet();
  await idle();
};
/** Press a command in the File (⋮) menu. */
const fileCmd = async (label) => {
  await closeSheet();
  await page.click('[data-testid=menu-more]');
  await btn('.bsheet', label).click();
  pressed[`File › ${label}`] = currentStep;
  await page.waitForTimeout(200);
};
const at = (ref) =>
  ev((ref) => {
    const V = window.__view;
    const v = V.viewInfo();
    const r = window.__doc.wb.getActiveSheet().getRange(ref).getRange();
    const a = (acc, i) => (i <= 0 ? 0 : acc[i - 1]);
    const p0 = V.toScreen(v, a(v.colAcc, r.startColumn), a(v.rowAcc, r.startRow));
    const p1 = V.toScreen(v, a(v.colAcc, r.endColumn + 1), a(v.rowAcc, r.endRow + 1));
    const h = document.querySelector('[data-testid=grid]').getBoundingClientRect();
    return { x: h.left + (p0.x + p1.x) / 2, y: h.top + (p0.y + p1.y) / 2, x0: h.left + p0.x, y0: h.top + p0.y };
  }, ref);
/** Long press a cell (or its row / column header) and press a command in the menu. */
const ctx = async (ref, label, area = 'cell') => {
  await closeSheet();
  const p = await at(ref);
  const g = await (await page.$('[data-testid=grid]')).boundingBox();
  const x = area === 'row' ? g.x + 14 : p.x;
  const y = area === 'col' ? p.y0 - 10 : p.y;
  await T.longPress(x, y);
  await page.waitForSelector('[data-testid=context-menu]');
  await btn('[data-testid=context-menu]', label).click();
  pressed[`Long press${area === 'cell' ? '' : area === 'row' ? ' (row header)' : ' (column header)'} › ${label}`] = currentStep;
  await page.waitForTimeout(200);
  await idle();
};
const S = (fn, arg) => ev(fn, arg);
const set = (ref, values) => S(([r, v]) => (window.__doc.wb.getActiveSheet().getRange(r).setValues(v), 0), [ref, values]);
const vals = (ref) => S((r) => window.__doc.wb.getActiveSheet().getRange(r).getValues(), ref);
const formula = (ref) => S((r) => window.__doc.wb.getActiveSheet().getRange(r).getFormula(), ref);
const style = (ref) => S((r) => window.__doc.wb.getActiveSheet().getRange(r).getCellStyleData() ?? {}, ref);
const shown = (ref) => S((r) => window.__doc.wb.getActiveSheet().getRange(r).getDisplayValue(), ref);
const clearSheet = async () => {
  await S(() => {
    const ws = window.__doc.wb.getActiveSheet();
    ws.getRange('A1:J30').clear();
    ws.cancelFreeze();
    const m = ws.getSheet();
    for (let i = 0; i < 30; i++) {
      if (!m.getRowVisible(i)) ws.showRows(i, 1);
      if (!m.getColVisible(i)) ws.showColumns(i, 1);
    }
    ws.getRange('A1').activate();
    return 0;
  });
  await idle();
};
const run = (title, fn) =>
  step(title, async () => {
    currentStep = title;
    await clearSheet();
    await fn();
  });

// ------------------------------------------------------------------ Edit
await run('Edit: Undo and Redo', async () => {
  await typeIn('A1', '1');
  await typeIn('A1', '2');
  await cmd('edit', 'Undo');
  assert.equal(await val('A1'), 1);
  await cmd('edit', 'Redo');
  assert.equal(await val('A1'), 2);
});

await run('Edit: Clear, Copy / Paste, Paste values, Cut', async () => {
  await set('A1:B2', [[1, 2], [3, 4]]);
  await S(() => (window.__doc.wb.getActiveSheet().getRange('C1').setValue({ f: '=A1*10' }), 0));
  await idle();
  await select('C1');
  await cmd('edit', 'Copy');
  await select('D2');
  await cmd('edit', 'Paste');
  assert.equal(await formula('D2'), '=B2*10', 'pasted formula moves its reference');
  assert.equal(await val('D2'), 40);
  await select('E1');
  await cmd('edit', 'Paste values');
  assert.equal(await formula('E1'), '', 'value only');
  assert.equal(await val('E1'), 10);
  await select('C1');
  await cmd('edit', 'Cut');
  await select('C5');
  await cmd('edit', 'Paste');
  assert.equal(await val('C1'), null, 'cut leaves the cell empty');
  assert.equal(await formula('C5'), '=A1*10', 'a cut formula keeps its reference');
  await select('A1:B2');
  await cmd('edit', 'Clear');
  assert.deepEqual(await vals('A1:B2'), [[null, null], [null, null]]);
});

await run('Edit: Fill down and Fill right', async () => {
  await set('A1', [[7]]);
  await select('A1:A4');
  await cmd('edit', 'Fill down');
  assert.deepEqual((await vals('A1:A4')).flat(), [7, 7, 7, 7]);
  await select('A1:D1');
  await cmd('edit', 'Fill right');
  assert.deepEqual((await vals('A1:D1')).flat(), [7, 7, 7, 7]);
});

await run('Edit: insert and delete rows and columns', async () => {
  await set('A1:B3', [[1, 2], [3, 4], [5, 6]]);
  await select('A2');
  await cmd('edit', 'Row above');
  assert.deepEqual((await vals('A1:A4')).flat(), [1, null, 3, 5]);
  await cmd('edit', 'Row below');
  assert.deepEqual((await vals('A1:A5')).flat(), [1, null, null, 3, 5]);
  await select('A2:A3');
  await cmd('edit', 'Delete rows');
  assert.deepEqual((await vals('A1:A3')).flat(), [1, 3, 5]);
  await select('B1');
  await cmd('edit', 'Column left');
  assert.deepEqual((await vals('A1:C1')).flat(), [1, null, 2]);
  await cmd('edit', 'Column right');
  assert.deepEqual((await vals('A1:D1')).flat(), [1, null, null, 2]);
  await select('B1:C1');
  await cmd('edit', 'Delete cols');
  assert.deepEqual((await vals('A1:B1')).flat(), [1, 2]);
});

await run('Edit: hide rows, hide columns, unhide all', async () => {
  const rowVis = (i) => S((i) => window.__doc.wb.getActiveSheet().getSheet().getRowVisible(i), i);
  const colVis = (i) => S((i) => window.__doc.wb.getActiveSheet().getSheet().getColVisible(i), i);
  await select('A3:A4');
  await cmd('edit', 'Hide rows');
  assert.deepEqual([await rowVis(2), await rowVis(3), await rowVis(4)], [false, false, true]);
  await select('C1');
  await cmd('edit', 'Hide cols');
  assert.equal(await colVis(2), false);
  await cmd('edit', 'Unhide all');
  assert.deepEqual([await rowVis(2), await rowVis(3), await colVis(2)], [true, true, true]);
});

// ---------------------------------------------------------------- Format
await run('Format: Italic, Underline, Strike, Bold', async () => {
  await set('A1', [['text']]);
  await select('A1');
  await cmd('format', 'Bold');
  await cmd('format', 'Italic');
  await cmd('format', 'Underline');
  await cmd('format', 'Strike');
  const s = await style('A1');
  assert.equal(s.bl, 1, 'bold');
  assert.equal(s.it, 1, 'italic');
  assert.equal(s.ul?.s, 1, 'underline');
  assert.equal(s.st?.s, 1, 'strike');
});

await run('Format: Left, Center, Right, Wrap', async () => {
  await select('A1');
  await cmd('format', 'Center');
  assert.equal((await style('A1')).ht, 2);
  await cmd('format', 'Right');
  assert.equal((await style('A1')).ht, 3);
  await cmd('format', 'Left');
  assert.equal((await style('A1')).ht, 1);
  await cmd('format', 'Wrap');
  assert.equal((await style('A1')).tb, 3, 'wrap text');
});

await run('Format: fill colour, text colour, no fill', async () => {
  await select('A1');
  await closeSheet();
  await page.click('[data-testid=tab-format]');
  await page.click('.bsheet [aria-label="Fill #fde2e4"]');
  await page.click('.bsheet [aria-label="Text #b42318"]');
  pressed['Format › Fill colour'] = currentStep;
  pressed['Format › Text colour'] = currentStep;
  await closeSheet();
  const s = await style('A1');
  assert.equal(s.bg?.rgb?.toLowerCase(), '#fde2e4');
  assert.equal(s.cl?.rgb?.toLowerCase(), '#b42318');
  await page.click('[data-testid=tab-format]');
  await page.click('.bsheet [aria-label="No fill"]');
  pressed['Format › No fill'] = currentStep;
  await closeSheet();
  assert.ok(!(await style('A1')).bg?.rgb, 'fill removed');
});

await run('Format: Merge and Unmerge', async () => {
  const merges = () => S(() => window.__doc.wb.getActiveSheet().getMergedRanges().map((r) => r.getA1Notation()));
  await select('A1:C2');
  await cmd('format', 'Merge');
  assert.deepEqual(await merges(), ['A1:C2']);
  await cmd('format', 'Unmerge');
  assert.deepEqual(await merges(), []);
});

await run('Format: All borders, Outline, Top, Bottom, No border, Clear format', async () => {
  const bd = (ref) => style(ref).then((s) => Object.keys(s.bd ?? {}).filter((k) => s.bd[k]).sort().join(','));
  await select('B2:C3');
  await cmd('format', 'All borders');
  assert.equal(await bd('B2'), 'b,l,r,t');
  await cmd('format', 'No border');
  assert.equal(await bd('B2'), '');
  await cmd('format', 'Outline');
  assert.equal(await bd('B2'), 'l,t', 'the outline gives the top-left cell its top and left edge');
  assert.equal(await bd('C3'), 'b,r');
  await cmd('format', 'No border');
  await select('B2');
  await cmd('format', 'Top');
  assert.equal(await bd('B2'), 't');
  await cmd('format', 'Bottom');
  assert.equal(await bd('B2'), 'b,t');
  await cmd('format', 'Bold');
  await cmd('format', 'Clear format');
  const s = await style('B2');
  assert.ok(!s.bl && !Object.keys(s.bd ?? {}).length, 'format cleared: ' + JSON.stringify(s));
});

await run('Format: every number format shows the expected text', async () => {
  const cases = [
    ['General', 1234.5, '1234.5'],
    ['Number', 1234.5, '1234.50'],
    ['Thousands', 1234.5, '1,234.50'],
    ['Integer', 1234.5, '1,235'],
    ['Currency', 1234.5, '$1,234.50'],
    ['Euro', 1234.5, '1,234.50 €'],
    ['Percent', 0.1234, '12.34%'],
    ['Scientific', 1234.5, '1.23E+03'],
    ['Date', 46339.587, '2026-11-13'],
    ['Date (d/m)', 46339.587, '13/11/2026'],
    ['Date (m/d)', 46339.587, '11/13/2026'],
    ['Time', 46339.587, '14:05'],
    ['Date time', 46339.587, '2026-11-13 14:05'],
  ];
  for (const [label, v, want] of cases) {
    await set('A1', [[v]]);
    await select('A1');
    await cmd('format', 'Number format', { keepOpen: true });
    await page.click(`[data-testid="numfmt-${label}"]`);
    pressed[`Format › Number format › ${label}`] = currentStep;
    await idle();
    assert.equal(await shown('A1'), want, label);
  }
  await set('A2', [['0012']]);
  await select('A2');
  await cmd('format', 'Number format', { keepOpen: true });
  await page.click('[data-testid="numfmt-Text"]');
  pressed['Format › Number format › Text'] = currentStep;
  await idle();
  await typeIn('A2', '0012');
  assert.equal(await val('A2'), '0012', 'a Text cell keeps leading zeros');
  await typeIn('A2', '=1+1');
  assert.equal(await val('A2'), '=1+1', 'and shows a formula as typed');
  assert.equal(await formula('A2'), '');
});

await run('Format: every conditional formatting rule', async () => {
  await set('A1:A6', [[5], [12], [12], [30], ['apple'], ['pear']]);
  const kinds = [
    ['Greater than', '10', 'cellIs'],
    ['Less than', '10', 'cellIs'],
    ['Between', '10', 'cellIs'],
    ['Equal to', '12', 'cellIs'],
    ['Text contains', 'app', 'containsText'],
    ['Duplicates', null, 'duplicateValues'],
    ['Colour scale', null, 'colorScale'],
    ['Data bar', null, 'dataBar'],
  ];
  let n = 0;
  const want = {};
  for (const [label, value, type] of kinds) {
    await select('A1:A6');
    await cmd('format', 'Conditional', { keepOpen: true });
    await page.click(`[data-testid=dlg-cf] button.chip:text-is("${label}")`);
    if (value !== null) await page.fill('[data-testid=cf-value]', value);
    if (label === 'Between') await page.fill('[data-testid=dlg-cf] .row input[inputmode=decimal] >> nth=1', '20');
    await page.click('[data-testid=cf-apply]');
    pressed[`Format › Conditional › ${label}`] = currentStep;
    await idle();
    const rules = await S(() => window.__doc.collectCf().map((r) => r.rule.type));
    n++;
    assert.equal(rules.length, n, label + ' added');
    want[type] = (want[type] ?? 0) + 1;
    const got = {};
    for (const r of rules) got[r] = (got[r] ?? 0) + 1;
    assert.deepEqual(got, want, label);
  }
});

// ---------------------------------------------------------------- Insert
await run('Insert: Function, New sheet', async () => {
  await select('B1');
  await cmd('insert', 'Function', { keepOpen: true });
  await page.click('[role=dialog] .list-item:has-text("SUM")');
  await idle();
  assert.equal(await formula('B1'), '=SUM(A1:A10)');
  const before = await S(() => window.__doc.wb.getSheets().length);
  await cmd('insert', 'New sheet');
  assert.equal(await S(() => window.__doc.wb.getSheets().length), before + 1);
  await S(() => (window.__doc.wb.setActiveSheet(window.__doc.wb.getSheets()[0]), 0));
  await page.waitForTimeout(200);
});

// ------------------------------------------------------------------ Data
await run('Data: Sort Z→A, Clear filters', async () => {
  await set('A1:B4', [['Name', 'Score'], ['b', 2], ['c', 3], ['a', 1]]);
  await select('B2');
  await cmd('data', 'Sort Z→A');
  assert.deepEqual((await vals('B1:B4')).flat(), ['Score', 3, 2, 1], 'header kept, numbers descending');
  await select('A1');
  await cmd('data', 'Filter on/off');
  await S(() => (window.__doc.wb.getActiveSheet().getFilter().setColumnFilterCriteria(0, { colId: 0, filters: { filters: ['a'] } }), 0));
  await idle();
  const hidden = () => S(() => [1, 2, 3].filter((i) => !window.__doc.wb.getActiveSheet().getSheet().getRowVisible(i)).length);
  assert.equal(await hidden(), 2, 'filter hides two rows');
  await cmd('data', 'Clear filters');
  assert.equal(await hidden(), 0, 'all rows back');
  assert.ok(await S(() => !!window.__doc.wb.getActiveSheet().getFilter()), 'the filter itself stays');
  await cmd('data', 'Filter on/off');
  assert.ok(await S(() => !window.__doc.wb.getActiveSheet().getFilter()));
});

await run('Data: freeze top row, first column, at selection, unfreeze', async () => {
  const fz = () => S(() => { const ws = window.__doc.wb.getActiveSheet(); return [ws.getFrozenRows(), ws.getFrozenColumns()]; });
  await cmd('data', 'First column');
  assert.deepEqual(await fz(), [0, 1]);
  await cmd('data', 'Unfreeze');
  assert.deepEqual(await fz(), [0, 0]);
  await select('C4');
  await cmd('data', 'At selection');
  assert.deepEqual(await fz(), [3, 2]);
  await cmd('data', 'Top row');
  assert.deepEqual((await fz())[0], 1);
  await cmd('data', 'Unfreeze');
  assert.deepEqual(await fz(), [0, 0]);
});

// ------------------------------------------------------------ Tools / File
await run('Tools: Charts, Device & perf, About; File: About open their screens', async () => {
  await cmd('tools', 'Charts', { keepOpen: true });
  await page.waitForSelector('[role=dialog][aria-label*="hart"]');
  await page.keyboard.press('Escape');
  await page.evaluate(() => window.__petakBack?.());
  await page.waitForTimeout(200);
  await cmd('tools', 'Device & perf', { keepOpen: true });
  await page.waitForSelector('[role=dialog][aria-label^="About"]');
  await page.evaluate(() => window.__petakBack?.());
  await page.waitForTimeout(200);
  await cmd('tools', 'About', { keepOpen: true });
  await page.waitForSelector('[role=dialog][aria-label^="About"]');
  await page.evaluate(() => window.__petakBack?.());
  await page.waitForTimeout(200);
  await fileCmd('About');
  await page.waitForSelector('[role=dialog][aria-label^="About"]');
  await page.evaluate(() => window.__petakBack?.());
  await page.waitForTimeout(200);
  assert.equal(await page.$('[role=dialog]'), null);
});

await run('File: Dark mode and Light mode', async () => {
  await fileCmd('Dark mode');
  assert.equal(await page.evaluate(() => document.documentElement.dataset.theme), 'dark');
  await fileCmd('Light mode');
  assert.equal(await page.evaluate(() => document.documentElement.dataset.theme), 'light');
});

await run('File: Save .xlsx, Save .xlsm, Share (a download in a browser, in the current format)', async () => {
  await set('A1', [['saved']]);
  const head = (f) => fs.readFileSync(f).subarray(0, 2).toString();
  const x = await download(() => fileCmd('Save .xlsx'));
  assert.match(x, /\.xlsx$/);
  assert.equal(head(x), 'PK');
  const m = await download(() => fileCmd('Save .xlsm'));
  assert.match(m, /\.xlsm$/);
  assert.equal(head(m), 'PK');
  const s = await download(() => fileCmd('Share'));
  assert.match(s, /\.xlsm$/, 'after Save .xlsm the workbook is an .xlsm');
  assert.equal(head(s), 'PK');
});

// ----------------------------------------------------------- Sheet tabs
await run('Sheet tab: Duplicate and Delete', async () => {
  const names = () => S(() => window.__doc.wb.getSheets().map((s) => s.getSheetName()));
  const before = await names();
  const first = before[0];
  const menu = async () => {
    await page.click(`[data-testid="sheet-tab-${first}"]`, { button: 'right' });
    await page.waitForSelector('[data-testid=sheet-rename]');
  };
  await menu();
  await btn('.bsheet', 'Duplicate').click();
  pressed['Sheet tab › Duplicate'] = currentStep;
  await page.waitForTimeout(300);
  const after = await names();
  assert.equal(after.length, before.length + 1);
  const copy = after.find((n) => !before.includes(n));
  assert.ok(copy, 'a new sheet: ' + after);
  await page.click(`[data-testid="sheet-tab-${copy}"]`, { button: 'right' });
  await page.waitForSelector('[data-testid=sheet-rename]');
  await btn('.bsheet', 'Delete').click();
  pressed['Sheet tab › Delete'] = currentStep;
  await page.waitForSelector('[data-testid=confirm-ok], [data-testid=ask-ok]');
  await page.click('[data-testid=confirm-ok], [data-testid=ask-ok]');
  await page.waitForTimeout(300);
  assert.deepEqual(await names(), before);
  await S(() => (window.__doc.wb.setActiveSheet(window.__doc.wb.getSheets()[0]), 0));
  await closeSheet();
});

// --------------------------------------------------------- Long press menu
await run('Long press: Cut, Paste format, Paste formulas', async () => {
  await set('A1:B1', [[2, 3]]);
  await S(() => (window.__doc.wb.getActiveSheet().getRange('C1').setValue({ f: '=A1+B1' }).setFontWeight('bold'), 0));
  await idle();
  await ctx('C1', 'Copy');
  await ctx('C3', 'Paste format');
  assert.equal((await style('C3')).bl, 1, 'format pasted');
  assert.equal(await val('C3'), null, 'no value pasted');
  await ctx('D1', 'Paste formulas');
  assert.equal(await formula('D1'), '=B1+C1');
  assert.ok(!(await style('D1')).bl, 'formula only, no bold');
  await ctx('C1', 'Cut');
  await ctx('C5', 'Paste');
  assert.equal(await formula('C5'), '=A1+B1');
  assert.equal(await val('C1'), null);
});

await run('Long press: insert cells down / right, delete cells up', async () => {
  await set('A1:B3', [[1, 'x'], [2, 'y'], [3, 'z']]);
  await ctx('A2', 'Cells ↓');
  assert.deepEqual((await vals('A1:A4')).flat(), [1, null, 2, 3]);
  assert.deepEqual((await vals('B1:B3')).flat(), ['x', 'y', 'z'], 'the next column did not move');
  await ctx('A2', 'Delete cells ↑');
  assert.deepEqual((await vals('A1:A3')).flat(), [1, 2, 3]);
  await ctx('A1', 'Cells →');
  assert.deepEqual((await vals('A1:C1')).flat(), [null, 1, 'x']);
});

await run('Long press: rows and columns from the cell menu', async () => {
  await set('A1:B2', [[1, 2], [3, 4]]);
  await ctx('A1', 'Row below');
  assert.deepEqual((await vals('A1:A3')).flat(), [1, null, 3]);
  await ctx('A2', 'Delete rows');
  assert.deepEqual((await vals('A1:A2')).flat(), [1, 3]);
  await ctx('B1', 'Column left');
  assert.deepEqual((await vals('A1:C1')).flat(), [1, null, 2]);
  await ctx('A1', 'Column right');
  assert.deepEqual((await vals('A1:D1')).flat(), [1, null, null, 2]);
  await ctx('B1', 'Delete cols');
  assert.deepEqual((await vals('A1:C1')).flat(), [1, null, 2]);
});

await run('Long press: column header hide / unhide, row header unhide', async () => {
  const colVis = (i) => S((i) => window.__doc.wb.getActiveSheet().getSheet().getColVisible(i), i);
  const rowVis = (i) => S((i) => window.__doc.wb.getActiveSheet().getSheet().getRowVisible(i), i);
  await select('B1:C1');
  await ctx('B1', 'Hide columns', 'col');
  assert.equal(await colVis(1), false);
  await select('A1:D1');
  await ctx('A1', 'Unhide columns', 'col');
  assert.equal(await colVis(1), true);
  await ctx('A3', 'Hide rows', 'row');
  assert.equal(await rowVis(2), false);
  await select('A2:A4');
  await ctx('A2', 'Unhide rows', 'row');
  assert.equal(await rowVis(2), true);
});

await run('Long press: Edit, Clear contents, Clear formats, Clear all, Fill right', async () => {
  await set('A1:B1', [[5, null]]);
  await S(() => (window.__doc.wb.getActiveSheet().getRange('A1').setFontWeight('bold'), 0));
  await select('A1:B1');
  await ctx('A1', 'Fill right');
  assert.deepEqual((await vals('A1:B1')).flat(), [5, 5]);
  await select('B1'); // a long press inside the selection keeps it
  await ctx('B1', 'Clear contents');
  assert.equal(await val('B1'), null);
  assert.equal((await style('B1')).bl, 1, 'format kept');
  await ctx('A1', 'Clear formats');
  assert.equal(await val('A1'), 5);
  assert.ok(!(await style('A1')).bl, 'bold gone');
  await ctx('B1', 'Clear all');
  assert.ok(!(await style('B1')).bl);
  await ctx('A1', 'Edit');
  await page.waitForFunction(() => document.activeElement?.matches('input, textarea, [contenteditable=true]'));
  await page.keyboard.press('Control+A');
  await page.keyboard.type('9');
  await page.keyboard.press('Enter');
  await idle();
  assert.equal(await val('A1'), 9);
});

await run('a tap selects the cell even right after a row-header menu and a selection made in code', async () => {
  const g = await (await page.$('[data-testid=grid]')).boundingBox();
  const sel = () => S(() => window.__doc.wb.getActiveSheet().getActiveRange()?.getA1Notation());
  const p3 = await at('A3');
  await T.longPress(g.x + 14, p3.y);
  await page.waitForSelector('[data-testid=context-menu]');
  assert.equal(await sel(), '3:3');
  await page.mouse.click(200, 120);
  await page.waitForTimeout(200);
  await select('B1');
  const p = await at('A1');
  await T.tap(p.x, p.y);
  await page.waitForTimeout(300);
  assert.equal(await sel(), 'A1', 'Univer ignored this tap before the fix');
});

await run('taps leave nothing behind on Univer’s command stack (F-07)', async () => {
  const g = await (await page.$('[data-testid=grid]')).boundingBox();
  for (let i = 0; i < 12; i++) {
    await T.tap(g.x + 80 + (i % 3) * 90, g.y + 60 + (i % 6) * 40);
    await page.waitForTimeout(60);
  }
  assert.equal(await S(() => window.__doc.api._commandService._commandExecutionStack.length), 0);
});

await run('Long press: Format…, Link…, Comment open their panels', async () => {
  await ctx('A1', 'Format…');
  await page.waitForSelector('.bsheet [data-testid=fmt-bold]');
  await closeSheet();
  await ctx('A1', 'Link…');
  await page.waitForSelector('[role=dialog][aria-label*="ink"]');
  await page.evaluate(() => window.__petakBack?.());
  await page.waitForTimeout(200);
  await ctx('A1', 'Comment');
  await page.waitForSelector('[role=dialog][aria-label*="omment"]');
  await page.evaluate(() => window.__petakBack?.());
  await page.waitForTimeout(200);
  assert.equal(await page.$('[role=dialog]'), null);
});

fs.mkdirSync('tests/e2e/out', { recursive: true });
fs.writeFileSync('tests/e2e/out/commands.json', JSON.stringify(pressed, null, 1) + '\n');
await t.finish();
