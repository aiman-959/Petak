/** Rich text in a cell (H-07): shown as formatted runs, kept on save, replaced when the cell is retyped. */
import { describe, expect, it } from 'vitest';
import fs from 'node:fs';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { writeOoxml } from '../../src/io/ooxml/write';
import { ab, load } from './ooxml-helpers';

/** the roster with its first shared string turned into three runs */
function richBook() {
  const z = unzipSync(new Uint8Array(fs.readFileSync('tests/fixtures/files/roster.xlsx')));
  const xml = strFromU8(z['xl/sharedStrings.xml']);
  const first = /<si><t[^>]*>([^<]*)<\/t><\/si>/.exec(xml)!;
  const runs =
    '<si><r><t xml:space="preserve">Plain </t></r>' +
    '<r><rPr><b/><sz val="14"/><color rgb="FFFF0000"/><rFont val="Arial"/></rPr><t xml:space="preserve">bold red</t></r>' +
    '<r><rPr><i/><u/><vertAlign val="superscript"/></rPr><t xml:space="preserve"> note</t></r></si>';
  z['xl/sharedStrings.xml'] = strToU8(xml.replace(first[0], runs));
  return { buf: ab(zipSync(z)), was: first[1] };
}
const findCell = (l: ReturnType<typeof load>, text: string) => {
  for (const sid of l.data.sheetOrder) {
    const cd = l.data.sheets[sid].cellData as Record<number, Record<number, { v?: unknown; p?: { body: { dataStream: string; textRuns: Array<{ st: number; ed: number; ts: Record<string, unknown> }> } } }>>;
    for (const r in cd) for (const c in cd[r]) if (cd[r][c].v === text) return { sid, r: +r, c: +c, cell: cd[r][c] };
  }
  throw new Error('cell not found: ' + text);
};
const saveOf = (l: ReturnType<typeof load>) => writeOoxml(l.data, { cf: l.res.cf, meta: l.res.meta, macroEnabled: false, original: l.original, baseline: l.baseline, names: l.res.names, pictures: l.res.pictures });

describe('rich text', () => {
  it('is loaded as runs with their own font, size, colour and style', () => {
    const l = load(richBook().buf, 'rich.xlsx');
    const { cell } = findCell(l, 'Plain bold red note');
    expect(cell.p!.body.dataStream).toBe('Plain bold red note\r\n');
    const runs = cell.p!.body.textRuns;
    expect(runs.map((r) => [r.st, r.ed])).toEqual([[0, 6], [6, 14], [14, 19]]);
    // a run without formatting takes the cell's own font (the roster title: bold 22 pt white)
    expect(runs[0].ts).toMatchObject({ fs: 22, bl: 1, cl: { rgb: '#ffffff' } });
    // a run's own formatting wins over the cell's
    expect(runs[1].ts).toMatchObject({ ff: 'Arial', fs: 14, bl: 1, cl: { rgb: '#ff0000' } });
    expect(runs[2].ts).toMatchObject({ it: 1, ul: { s: 1 }, va: 3, fs: 22 });
  });
  it('is saved back unchanged when the cell was not edited', () => {
    const l = load(richBook().buf, 'rich.xlsx');
    const out = strFromU8(unzipSync(new Uint8Array(saveOf(l).buf))['xl/sharedStrings.xml']);
    expect(out).toContain('<r><rPr><b/><sz val="14"/><color rgb="FFFF0000"/><rFont val="Arial"/></rPr><t xml:space="preserve">bold red</t></r>');
  });
  it('typing over the cell saves the new plain text', () => {
    const l = load(richBook().buf, 'rich.xlsx');
    const { cell } = findCell(l, 'Plain bold red note');
    cell.v = 'retyped';
    delete cell.p; // what commitInput does: { v, p: null }
    const out = strFromU8(unzipSync(new Uint8Array(saveOf(l).buf))['xl/sharedStrings.xml']);
    expect(out).toContain('retyped');
    const back = load(saveOf(l).buf, 'rich2.xlsx');
    expect(findCell(back, 'retyped').cell.p).toBeUndefined();
  });
  it('plain strings get no document', () => {
    const l = load('tests/fixtures/files/roster.xlsx');
    const cd = Object.values(l.data.sheets)[0].cellData as Record<number, Record<number, { p?: unknown }>>;
    expect(Object.values(cd).every((row) => Object.values(row).every((c) => !c.p))).toBe(true);
  });
});
