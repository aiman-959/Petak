import { describe, expect, it } from 'vitest';
import fs from 'node:fs';
import { readBiffFormatting } from '../../src/io/biff';
import { parseWithSheetJs, streamSheetJs } from '../../src/io/csv';
import type { LoadChunk, LoadResult } from '../../src/io/types';

// the staff roster saved as Excel 97–2003 by LibreOffice
const file = fs.readFileSync('tests/fixtures/files/legacy.xls');
const buf = () => file.buffer.slice(file.byteOffset, file.byteOffset + file.byteLength) as ArrayBuffer;

describe('legacy .xls formatting (BIFF8)', () => {
  it('reads fonts, fills, borders, alignment, column widths and row heights', () => {
    const info = readBiffFormatting(buf())!;
    expect(info.defaultFont).toEqual({ name: 'Calibri', size: 11 });
    expect(info.sheets.map((s) => s.name)).toEqual(['排班表 Roster', 'Legend']);
    const sh = info.sheets[0];
    const style = (r: number, c: number) => info.styles[sh.cells.get(r)!.get(c)!]!;
    // title: 22 pt bold white on dark green, centred
    expect(style(0, 0)).toMatchObject({ fs: 22, bl: 1, cl: { rgb: '#ffffff' }, bg: { rgb: '#14532d' }, ht: 2 });
    // header row: amber fill, medium border all round
    const head = style(2, 0);
    expect(head.bg).toEqual({ rgb: '#fbbf24' });
    expect(head.bd).toMatchObject({ l: { s: 8 }, r: { s: 8 }, t: { s: 8 }, b: { s: 8 } });
    // a shift cell: pastel fill and thin borders
    expect(style(3, 2).bg?.rgb).toMatch(/^#[0-9a-f]{6}$/);
    expect(style(3, 2).bd).toMatchObject({ l: { s: 1 }, t: { s: 1 } });
    expect(sh.colWidths.get(0)).toBeCloseTo(6, 1);
    expect(sh.colWidths.get(1)).toBeCloseTo(18, 1);
    expect(sh.rowHeights.get(0)).toBe(42);
  });
  it('the loader applies them: styles on cells, widths and heights on the sheet, default font', () => {
    const p = parseWithSheetJs(buf(), 'legacy.xls');
    const gen = streamSheetJs(p, 1000);
    const chunks: LoadChunk[] = [];
    let res: IteratorResult<LoadChunk, LoadResult>;
    while (!(res = gen.next()).done) chunks.push(res.value);
    const result = res.value as LoadResult & { defaultStyle?: { ff: string; fs: number }; warnings?: string[] };
    const start = chunks.find((c) => c.kind === 'sheetStart' && c.index === 0) as Extract<LoadChunk, { kind: 'sheetStart' }>;
    const sheet = start.sheet as unknown as { columnData: Record<number, { w: number }>; defaultRowHeight: number };
    expect(sheet.columnData[1].w).toBeGreaterThan(sheet.columnData[0].w * 2.5);
    const rows = chunks.filter((c) => c.kind === 'rows' && c.index === 0) as Array<Extract<LoadChunk, { kind: 'rows' }>>;
    const a1 = rows[0].cellData[0][0];
    expect(a1.v).toContain('Staff Roster');
    expect((result.styles as Record<string, { bg?: { rgb: string }; fs?: number }>)[a1.s as string]).toMatchObject({ bg: { rgb: '#14532d' }, fs: 22 });
    expect(rows.at(-1)!.rowData![0]).toEqual({ h: 56, ia: 0 });
    expect(result.defaultStyle).toEqual({ ff: 'Calibri', fs: 11 });
    expect(result.warnings?.[0]).toMatch(/Old \.xls format/);
  });
  it('anything that is not a BIFF8 workbook gives null instead of throwing', () => {
    expect(readBiffFormatting(new Uint8Array([1, 2, 3, 4]).buffer)).toBeNull();
    const xlsx = fs.readFileSync('tests/fixtures/files/roster.xlsx');
    expect(readBiffFormatting(xlsx.buffer.slice(xlsx.byteOffset, xlsx.byteOffset + xlsx.byteLength) as ArrayBuffer)).toBeNull();
  });
});
