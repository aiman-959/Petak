/**
 * FEATURES.md: every command in Petak's menus (read from the source) with the test that presses
 * it. Commands pressed by tests/e2e/commands.mjs come from its tests/e2e/out/commands.json; the
 * rest have a deeper test in another suite, listed in ELSEWHERE. Fails when a command has no test.
 *   node tests/e2e/commands.mjs && node tests/e2e/feature-matrix.mjs
 */
import fs from 'node:fs';

const src = (f) => fs.readFileSync(new URL('../../src/' + f, import.meta.url), 'utf8');
const labels = (text) => [...text.matchAll(/<Tool\s+icon=\{[^}]*\}\s+label=(?:"([^"]+)"|\{([^}]+)\})/g)].map((m) => m[1] ?? m[2]);
const nice = (l) => (l.includes("'Edit link'") ? 'Link…' : l.includes('Light mode') ? 'Dark mode' : l);

/** menu → commands, as the app shows them */
const menus = [];
const panels = src('ui/ToolPanels.tsx').split(/case '(\w+)':/).slice(1);
const NAMES = { edit: 'Edit', format: 'Format', insert: 'Insert', draw: 'Draw', data: 'Data', tools: 'Tools' };
for (let i = 0; i < panels.length; i += 2) {
  const extra = panels[i] === 'format' ? ['Fill colour', 'Text colour', 'No fill'] : [];
  menus.push([NAMES[panels[i]], [...labels(panels[i + 1]), ...extra]]);
}
const chrome = src('ui/Chrome.tsx');
const fileMenu = chrome.slice(chrome.indexOf('title="File"'), chrome.indexOf('</Sheet>', chrome.indexOf('title="File"')));
menus.push(['File', labels(fileMenu).map(nice)]);
const tabMenu = chrome.slice(chrome.indexOf('function SheetTabs'), chrome.indexOf('const TABS'));
menus.push(['Sheet tab', labels(tabMenu)]);
const cx = src('ui/CellExtras.tsx');
const ctxMenu = cx.slice(cx.indexOf('export function ContextMenu'), cx.indexOf('export function LinkDialog'));
menus.push(['Long press', [...new Set(labels(ctxMenu).map(nice))]]);
const dialogs = src('ui/Dialogs.tsx');
const numfmts = [...src('ui/actions.ts').matchAll(/\{ label: '([^']+)', pattern:/g)].map((m) => m[1]);
const cfKinds = [...dialogs.matchAll(/\{ id: '\w+', label: '([^']+)' \}/g)].map((m) => m[1]);
menus.push(['Format › Number format', numfmts]);
menus.push(['Format › Conditional', cfKinds]);

const S = (suite, stepTitle) => `\`${suite}\`: ${stepTitle}`;
/** Commands whose test is in another suite. */
const ELSEWHERE = {
  'Edit › Find & replace': S('features.mjs', '3e find & replace'),
  'Format › Font & size': S('fonts.mjs', 'Format → Font & size sets the font and size of the selection'),
  'Insert › Bar chart': S('features.mjs', '5 bar, line, pie charts render and persist') + '; ' + S('taps.mjs', '10. Make a bar chart of a table'),
  'Insert › Line chart': S('features.mjs', '5 bar, line, pie charts render and persist'),
  'Insert › Pie chart': S('features.mjs', '5 bar, line, pie charts render and persist'),
  'Insert › Pivot table': S('features.mjs', '6 pivot table (rows x columns, sum & count)'),
  'Insert › Comment / note': S('part2.mjs', 'comments: thread with reply, and a note'),
  'Insert › Sparklines': S('part2.mjs', 'sparklines: one per row, drawn in the cells'),
  'Insert › Picture': S('cells2.mjs', 'pictures: delete one, view full screen with pinch, insert a new one; all saved'),
  'Insert › Link': S('cells2.mjs', 'insert, edit and remove a link; saved in the file and other links untouched'),
  'Draw › Drawing mode': S('draw-touch.mjs', 'drawing mode opens in Select with a mode chip'),
  'Draw › Rectangle': S('draw-touch.mjs', 'K-03: rectangle is press, drag, release with a live preview'),
  'Draw › Ellipse': S('part3.mjs', 'arrow, ellipse, text box and image'),
  'Draw › Line': S('draw-touch.mjs', 'a line locks to 30° and keeps its length; the magnet turns snapping off'),
  'Draw › Polyline': S('draw-touch.mjs', 'K-02: polyline points are placed on release; a double tap finishes without a stray point'),
  'Draw › Arrow': S('part3.mjs', 'arrow, ellipse, text box and image'),
  'Draw › Text box': S('part3.mjs', 'arrow, ellipse, text box and image'),
  'Draw › Image': S('part3.mjs', 'arrow, ellipse, text box and image'),
  'Draw › Dimension': S('part3.mjs', 'dimension line measures in the chosen unit'),
  'Draw › Cell size (mm)': S('part3.mjs', 'column width / row height in mm'),
  'Data › Sort A→Z': S('features.mjs', '3c sort by column (header kept)') + '; ' + S('taps.mjs', '8. Sort a list by a column'),
  'Data › Validation': S('part2.mjs', 'validation: list with stop alert and input message'),
  'Data › Group': S('phase3.mjs', 'Group rows, collapse, expand; outline levels saved in the file'),
  'Data › Ungroup': S('phase3.mjs', 'Group rows, collapse, expand; outline levels saved in the file'),
  'Data › Collapse': S('phase3.mjs', 'Group rows, collapse, expand; outline levels saved in the file'),
  'Data › Expand': S('phase3.mjs', 'Group rows, collapse, expand; outline levels saved in the file'),
  'Data › Pivot table': S('features.mjs', '6 pivot table (rows x columns, sum & count)'),
  'Data › Transform': S('features.mjs', '9 transform: recorded steps, load, refresh after source change'),
  'Data › What-if': S('part2.mjs', 'goal seek: loan amount for a 500/month payment') + '; ' + S('part2.mjs', 'scenarios: add, show, summary sheet') + '; ' + S('part2.mjs', 'data table: payment for several rates and terms'),
  'Data › Slicers': S('part2.mjs', 'slicers: filter a table and a pivot'),
  'Data › Data model': S('part2.mjs', 'data model: relationship + pivot across tables with DuckDB'),
  'Tools › Macros': S('features.mjs', '8a macro: getCell/setCell/getRange/addSheet in a worker'),
  'Tools › Transform': S('features.mjs', '9 transform: recorded steps, load, refresh after source change'),
  'File › New': S('files.mjs', 'New workbook asks too; without changes nothing is asked'),
  'File › Open': S('files.mjs', 'Save: the workbook is saved first, then the other file opens') + '; ' + S('taps.mjs', '1. Open a file from the phone'),
  'File › Recent': S('files.mjs', 'recent files: listed newest first; one that is gone is explained and removed'),
  'File › Save': S('files.mjs', 'Save: the workbook is saved first, then the other file opens'),
  'File › Export .csv': S('phase3.mjs', 'M-07: CSV export says which sheet was written and that the others are not in it'),
  'File › Print / PDF': S('phase3.mjs', 'M-09: Print / PDF of the roster: pages at print size, frozen header on every page, CJK text'),
  'Sheet tab › Rename': S('phase3.mjs', 'M-10: rename a sheet with Excel’s rules; move it among the tabs'),
  'Sheet tab › Move left': S('phase3.mjs', 'M-10: rename a sheet with Excel’s rules; move it among the tabs'),
  'Sheet tab › Move right': S('phase3.mjs', 'M-10: rename a sheet with Excel’s rules; move it among the tabs'),
  'Long press › Copy': S('cells2.mjs', 'long press opens the menu: copy, paste values, insert row, fill down'),
  'Long press › Paste values': S('cells2.mjs', 'long press opens the menu: copy, paste values, insert row, fill down'),
  'Long press › Row above': S('cells2.mjs', 'long press opens the menu: copy, paste values, insert row, fill down'),
  'Long press › Fill down': S('cells2.mjs', 'long press opens the menu: copy, paste values, insert row, fill down'),
  'Long press › Remove link': S('cells2.mjs', 'insert, edit and remove a link; saved in the file and other links untouched'),
  'Long press › View': S('cells2.mjs', 'pictures: delete one, view full screen with pinch, insert a new one; all saved'),
  'Long press › Delete': S('cells2.mjs', 'pictures: delete one, view full screen with pinch, insert a new one; all saved'),
};

const pressed = JSON.parse(fs.readFileSync(new URL('./out/commands.json', import.meta.url), 'utf8'));
const pressedBy = (menu, cmd) => {
  const keys = Object.keys(pressed).filter((k) => {
    const parts = k.split(' › ');
    if (menu.includes(' › ')) return parts.length === 3 && `${parts[0]} › ${parts[1]}` === menu && parts[2] === cmd;
    return parts.length === 2 && parts[1] === cmd && (parts[0] === menu || (menu === 'Long press' && parts[0].startsWith('Long press')));
  });
  return keys.length ? S('commands.mjs', pressed[keys[0]]) : null;
};

const lines = [
  '# Feature matrix',
  '',
  'Every command in Petak’s menus, read from the source, with the end-to-end test that presses it on a',
  'phone-sized screen and checks the result. Generated by `tests/e2e/feature-matrix.mjs`.',
  '',
];
const missing = [];
let total = 0;
for (const [menu, cmds] of menus) {
  lines.push(`## ${menu}`, '', '| Command | Test |', '|---|---|');
  for (const cmd of [...new Set(cmds)]) {
    const key = `${menu} › ${cmd}`;
    const sub = menu.includes(' › ') ? pressedBy(menu, cmd) : null;
    const test = sub ?? pressedBy(menu, cmd) ?? ELSEWHERE[key];
    total++;
    if (!test) missing.push(key);
    lines.push(`| ${cmd} | ${test ?? '**no test**'} |`);
  }
  lines.push('');
}
lines.push(
  '## Beyond the menus',
  '',
  '| Feature | Test |',
  '|---|---|',
  `| Every one of the 476 Excel functions | \`function-check.test.ts\` (FUNCTION-CHECK.md) |`,
  `| Open → save → reopen keeps cells, styles, merges, panes, sizes; LibreOffice agrees (xlsx, xlsm, xls, ods, csv) | \`roundtrip.mjs\` |`,
  `| Lossless save: parts Petak does not model are kept byte for byte | \`ooxml-roundtrip.test.ts\`, \`features.mjs\` 1b |`,
  `| Touch: tap, double tap, long press, drag-scroll, fling, pinch, resize headers, fill handle | \`cells.mjs\` |`,
  `| Keyboard and clipboard | \`cells2.mjs\` |`,
  `| Formula bar: tap cells to insert references, argument hints, coloured references | \`cells.mjs\`, \`phase3.mjs\` |`,
  `| Status bar Sum / Average / Count | \`cells2.mjs\`, \`taps.mjs\` 4 |`,
  `| Autosave and session restore | \`features.mjs\` 7b, \`part2.mjs\`, \`files.mjs\` (K-07) |`,
  `| Excel fonts and your own fonts | \`fonts.mjs\` |`,
  `| Grid lines and borders at every zoom, frozen panes | \`render.mjs\` |`,
  `| Low / Mid / High device tiers, load governor | \`part4.mjs\` |`,
  `| Android back button | \`files.mjs\`, \`draw-touch.mjs\` |`,
  '',
  `${total} menu commands, ${total - missing.length} with a test.`,
  '',
);
fs.writeFileSync(new URL('../../FEATURES.md', import.meta.url), lines.join('\n'));
console.log(`${total} menu commands, ${missing.length} without a test${missing.length ? ':\n  ' + missing.join('\n  ') : ''}`);
process.exit(missing.length ? 1 : 0);
