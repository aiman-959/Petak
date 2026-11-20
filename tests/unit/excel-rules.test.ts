/**
 * Saved files open in Excel without "We found a problem with some content" (Part 2 of the 1.3.0
 * brief): the rules in excel-rules.ts, checked on the sample files, on what Petak saves for every
 * kind of file it opens (.xlsx, .xlsm, .xls, .ods, .csv, a new workbook), and on broken packages
 * to prove the checker notices each kind of damage. (Every package written by
 * ooxml-roundtrip.test.ts — edits, new styles, pictures added / moved / deleted, a sheet deleted —
 * is checked too.)
 */
import { beforeAll, describe, expect, it } from 'vitest';
import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { strToU8, unzipSync, zipSync, strFromU8 } from 'fflate';
import type { IWorkbookData, IWorksheetData } from '@univerjs/core';
import { excelProblems, parseXml } from './excel-rules';
import { writeOoxml } from '../../src/io/ooxml/write';
import { parseWithSheetJs, streamSheetJs } from '../../src/io/csv';
import { load, ab } from './ooxml-helpers';

const dir = path.resolve('tests/fixtures/files');
const F = (f: string) => path.join(dir, f);
beforeAll(() => {
  if (!fs.existsSync(F('qm-report.xlsx'))) execFileSync('node', ['tests/fixtures/make-fixtures.mjs'], { stdio: 'inherit' });
});

/** Read a file the way the io worker does for non-OOXML formats (SheetJS + Petak's readers). */
function readOther(file: string) {
  const p = parseWithSheetJs(ab(fs.readFileSync(file)), path.basename(file));
  const gen = streamSheetJs(p, 5000);
  const sheets: Record<string, IWorksheetData> = {};
  const order: string[] = [];
  let step = gen.next();
  while (!step.done) {
    const v = step.value as { kind: string; sheet?: IWorksheetData; index?: number; cellData?: object; rowData?: object };
    if (v.kind === 'sheetStart') {
      sheets[v.sheet!.id as string] = { ...(v.sheet as IWorksheetData), cellData: {} };
      order.push(v.sheet!.id as string);
    } else Object.assign(sheets[order[v.index!]].cellData, v.cellData);
    step = gen.next();
  }
  const res = step.value as { styles?: object; defaultStyle?: object };
  return { id: 'wb', name: path.basename(file), sheetOrder: order, sheets, styles: res.styles ?? {}, defaultStyle: res.defaultStyle } as unknown as IWorkbookData;
}

describe('the checker notices damage', () => {
  const base = () => {
    const l = load(F('qm-report.xlsx'));
    return unzipSync(new Uint8Array(l.original));
  };
  const zip = (z: Record<string, Uint8Array>) => zipSync(z);
  const edit = (z: Record<string, Uint8Array>, part: string, fn: (s: string) => string) => ({ ...z, [part]: strToU8(fn(strFromU8(z[part]))) });
  it('a good file has no problems', () => {
    expect(excelProblems(zip(base()))).toEqual([]);
  });
  it('an attribute twice (the F-08 kind of break)', () => {
    expect(() => parseXml('<a x="1" x="2"/>')).toThrow(/twice/);
    expect(excelProblems(zip(edit(base(), 'xl/styles.xml', (s) => s.replace('<fonts ', '<fonts count="9" ')))).join()).toMatch(/twice/);
  });
  it('worksheet elements out of order', () => {
    const z = edit(base(), 'xl/worksheets/sheet1.xml', (s) => s.replace(/(<sheetData[\s\S]*<\/sheetData>)([\s\S]*?)(<sheetViews[\s\S]*?<\/sheetViews>)/, '$3$2$1').replace(/(<sheetViews>[\s\S]*?<\/sheetViews>)([\s\S]*)(<pageMargins[^>]*\/>)/, '$3$2$1'));
    expect(excelProblems(zip(z)).join()).toMatch(/out of order/);
  });
  it('a style index past the end, a missing part, overlapping merges, a bad sheet name', () => {
    let z = edit(base(), 'xl/worksheets/sheet1.xml', (s) => s.replace(/<c r="A1"( s="\d+")?/, '<c r="A1" s="999"'));
    expect(excelProblems(zip(z)).join()).toMatch(/style 999/);
    z = base();
    delete z['xl/styles.xml'];
    expect(excelProblems(zip(z)).join()).toMatch(/missing part xl\/styles\.xml/);
    z = edit(base(), 'xl/worksheets/sheet1.xml', (s) => s.replace('</sheetData>', '</sheetData>').replace(/<mergeCells[^>]*>/, '<mergeCells count="2"><mergeCell ref="A1:B2"/>').replace('</mergeCells>', '<mergeCell ref="B2:C3"/></mergeCells>'));
    expect(excelProblems(zip(z)).join()).toMatch(/overlap|count/);
    z = edit(base(), 'xl/workbook.xml', (s) => s.replace(/(<sheet [^>]*name=")[^"]*"/, '$1Bad/Name"'));
    expect(excelProblems(zip(z)).join()).toMatch(/invalid sheet name/);
  });
});

describe('sample files (made by ExcelJS and LibreOffice) pass the rules', () => {
  for (const f of ['roster.xlsx', 'roster-lo.xlsx', 'qm-report.xlsx', 'qm-report-lo.xlsx', 'macro.xlsm', 'looks-lo.xlsx']) {
    it(f, () => {
      if (!fs.existsSync(F(f))) return;
      expect(excelProblems(fs.readFileSync(F(f)))).toEqual([]);
    });
  }
});

describe('what Petak saves passes the rules', () => {
  for (const f of ['roster.xlsx', 'roster-lo.xlsx', 'qm-report.xlsx', 'qm-report-lo.xlsx', 'macro.xlsm']) {
    it(`${f}: saved unchanged and saved after edits`, () => {
      const l = load(F(f));
      const opts = { cf: l.res.cf, meta: l.res.meta, macroEnabled: f.endsWith('.xlsm'), original: l.original, baseline: l.baseline, names: l.res.names, pictures: l.res.pictures };
      expect(excelProblems(writeOoxml(l.data, opts).buf)).toEqual([]);
      // edits: a value, a formula, a new style, a merge, a new row far down, a deleted cell
      const sid = l.data.sheetOrder[0];
      const s = l.data.sheets[sid] as IWorksheetData;
      const cd = s.cellData as Record<number, Record<number, object>>;
      cd[0] = { ...(cd[0] ?? {}), 0: { v: 'Edited <&> "text"', t: 1 } };
      cd[2] = { ...(cd[2] ?? {}), 5: { v: 3, f: '=SUM(1,2)', t: 2 } };
      cd[500] = { 3: { v: 12.5, t: 2, s: 'petakNew' } };
      (l.data.styles as Record<string, object>).petakNew = { bl: 1, it: 1, bg: { rgb: '#FFE8CC' }, n: { pattern: '#,##0.00 "kg"' }, bd: { b: { s: 1, cl: { rgb: '#000000' } } } };
      s.mergeData = [...(s.mergeData ?? []), { startRow: 600, endRow: 601, startColumn: 0, endColumn: 2 }];
      if (cd[1]) delete cd[1][1];
      expect(excelProblems(writeOoxml(l.data, opts).buf)).toEqual([]);
    });
  }
  for (const f of ['legacy.xls', 'roster.ods', 'looks.ods', 'people.csv']) {
    it(`${f} saved as .xlsx`, () => {
      if (!fs.existsSync(F(f))) return;
      const data = readOther(F(f));
      expect(excelProblems(writeOoxml(data, { cf: [], meta: null, macroEnabled: false }).buf)).toEqual([]);
    });
  }
});
