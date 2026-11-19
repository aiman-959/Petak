/**
 * Excel fonts (K-06): bundled substitutes under the Excel names, notice on open, original
 * names saved, font and size for the selection, own font files.
 *   BASE_URL=http://localhost:4173/ node tests/e2e/fonts.mjs
 */
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { unzipSync, strFromU8 } from 'fflate';
import { start, BASE } from './lib.mjs';

const t = await start('fonts');
const { page, step, ev, fresh, openFile, select } = t;
console.log('Petak e2e fonts —', BASE);
await fresh();
await page.waitForTimeout(300);
const F = (f) => path.resolve('tests/fixtures/files', f);
const loaded = () => ev(() => [...document.fonts].filter((f) => f.status === 'loaded').map((f) => `${f.family.replace(/"/g, '')} ${f.weight} ${f.style}`));
const width = (font, text = 'Hamburgefonstiv 0123') => ev(([font, text]) => { const c = document.createElement('canvas').getContext('2d'); c.font = '40px ' + font; return c.measureText(text).width; }, [font, text]);
const style = (ref) => ev((ref) => window.__doc.wb.getActiveSheet().getRange(ref).getCellStyleData(), ref);

await step('opening a file loads the substitutes under the Excel names and says so', async () => {
  await openFile(F('roster.xlsx'));
  const toast = await page.textContent('.toast');
  assert.match(toast, /Opened roster\.xlsx\. Fonts: 5 shown with built-in fonts of the same widths/);
  const faces = await loaded();
  for (const f of ['Calibri 400 normal', 'Calibri 700 normal', 'Cambria 400 normal', 'Times New Roman 400 italic', 'Courier New 400 normal', 'Arial 700 normal']) assert.ok(faces.includes(f), f + ' loaded: ' + faces.join(', '));
});

await step('the substitute is what is drawn: Calibri measures as Carlito, serif and mono stay distinct', async () => {
  await ev(() => window.__fonts.ensureFonts(['Carlito', 'Tinos', 'Cousine']));
  assert.equal(await width('Calibri'), await width('Carlito'));
  assert.equal(await width('"Times New Roman"'), await width('Tinos'));
  assert.equal(await width('"Courier New"'), await width('Cousine'));
  assert.notEqual(await width('Calibri'), await width('"Courier New"'));
  // monospace really is: narrow and wide letters take the same space
  assert.equal(await width('"Courier New"', 'iiiiiiiiii'), await width('"Courier New"', 'WWWWWWWWWW'));
  assert.notEqual(await width('Calibri', 'iiiiiiiiii'), await width('Calibri', 'WWWWWWWWWW'));
});

await step('Cyrillic and Greek text loads its own font file', async () => {
  const w0 = await width('Calibri', 'Қазақстан Ελλάδα');
  await ev(() => window.__fonts.ensureFonts(['Calibri', 'Carlito'], 'Қазақстан Ελλάδα'));
  assert.ok((await loaded()).filter((f) => f.startsWith('Calibri 400 normal')).length >= 3, 'Latin, Cyrillic and Greek files');
  const w1 = await width('Calibri', 'Қазақстан Ελλάδα');
  assert.equal(w1, await width('Carlito', 'Қазақстан Ελλάδα'));
  assert.ok(w0 > 0 && w1 > 0);
  assert.equal(await ev(() => document.fonts.check('40px Calibri', 'Қазақстан Ελλάδα')), true);
});

await step('the file is saved with the original font names', async () => {
  const bytes = await ev(async () => Array.from(new Uint8Array((await window.__doc.serialize('xlsx')).data)));
  const xml = strFromU8(unzipSync(new Uint8Array(bytes))['xl/styles.xml']);
  for (const n of ['Calibri', 'Cambria', 'Times New Roman', 'Courier New']) assert.ok(xml.includes(`<name val="${n}"/>`), n);
  assert.ok(!/Carlito|Caladea|Tinos|Cousine|Arimo/.test(xml), 'no substitute names in the file');
});

await step('Format → Font & size sets the font and size of the selection', async () => {
  await fresh();
  await page.waitForTimeout(300);
  await ev(() => window.__doc.wb.getActiveSheet().getRange('B2').setValue('Hello'));
  await select('B2');
  await page.click('[data-testid=tab-format]');
  await page.click('[data-testid=tool-fonts]');
  await page.waitForSelector('[data-testid=dlg-fonts]');
  assert.match(await page.textContent('[data-testid=font-Calibri]'), /shown with Carlito/);
  await page.click('[data-testid=font-Courier-New]');
  await page.click('[data-testid=font-size-14]');
  await page.waitForTimeout(150);
  const st = await style('B2');
  assert.equal(st.ff, 'Courier New');
  assert.equal(st.fs, 14);
  await page.fill('[data-testid=font-size-input]', '13.5');
  await page.press('[data-testid=font-size-input]', 'Enter');
  await page.waitForTimeout(150);
  assert.equal((await style('B2')).fs, 13.5);
});

await step('add your own font file; it survives a restart; remove it', async () => {
  const src = fs.readFileSync('node_modules/@fontsource/caladea/files/caladea-latin-700-normal.woff2');
  await page.setInputFiles('[data-testid=font-file]', { name: 'Test Hand-Bold.woff2', mimeType: 'font/woff2', buffer: src });
  await page.waitForSelector('[data-testid=font-user]');
  assert.match(await page.textContent('[data-testid=font-user]'), /Test Hand Bold/);
  assert.match(await page.textContent('[data-testid=font-Test-Hand]'), /your font/);
  await page.click('[data-testid=font-Test-Hand]');
  await page.waitForTimeout(150);
  assert.equal((await style('B2')).ff, 'Test Hand');
  await ev(() => window.__fonts.ensureFonts(['Caladea']));
  const bw = (f) => ev((f) => { const c = document.createElement('canvas').getContext('2d'); c.font = `bold 40px ${f}`; return c.measureText('Hamburgefonstiv').width; }, f);
  assert.equal(await bw('"Test Hand"'), await bw('Caladea'));
  assert.notEqual(await bw('"Test Hand"'), await bw('sans-serif'));
  await page.goto(BASE);
  await page.waitForFunction(() => window.__doc?.snapshot?.sheets?.length > 0, null, { timeout: 60000 });
  assert.deepEqual(await ev(() => window.__fonts.fonts.user().map((u) => [u.family, u.bold])), [['Test Hand', true]]);
  assert.equal(await ev(() => [...document.fonts].some((f) => f.family.replace(/"/g, '') === 'Test Hand' && f.status === 'loaded')), true);
  await ev(async () => { for (const u of window.__fonts.fonts.user()) await window.__fonts.fonts.remove(u.id); return 0; });
  assert.deepEqual(await ev(() => window.__fonts.fonts.user()), []);
});

await step('a file that is not a font gives a clear message', async () => {
  const msg = await ev(() => window.__fonts.fonts.add('notes.ttf', new TextEncoder().encode('not a font').buffer).then(() => 'added', (e) => e.message));
  assert.match(msg, /is not a font this device can use/);
});

await t.finish();
