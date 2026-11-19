import { describe, expect, it } from 'vitest';
import fs from 'node:fs';
import { numfmt } from '@univerjs/core';
import type { IStyleData, IWorksheetData } from '@univerjs/core';
import { zipSync, strToU8 } from 'fflate';
import { readOdsFormatting, odsBorder, odsNumberFormat, odsFreeze, lengthPt, odsDate, odsDuration } from '../../src/io/ods';
import { parseWithSheetJs, streamSheetJs } from '../../src/io/csv';
import { children, root } from '../../src/io/ooxml/xml';
import type { LoadChunk, LoadResult } from '../../src/io/types';
import { load, ab } from './ooxml-helpers';

// looks.ods: every look Petak shows, saved by LibreOffice (tests/fixtures/make-ods.mjs);
// looks-lo.xlsx: LibreOffice's own .xlsx of that file, the reference
const ODS = 'tests/fixtures/files/looks.ods';
const REF = 'tests/fixtures/files/looks-lo.xlsx';

type Sheet = IWorksheetData & { columnData: Record<number, { w?: number; hd?: number }>; rowData: Record<number, { h?: number; hd?: number; ia?: number }> };
function loadOds(file: string) {
  const p = parseWithSheetJs(ab(fs.readFileSync(file)), file);
  const gen = streamSheetJs(p, 1000);
  const sheets: Sheet[] = [];
  let res: IteratorResult<LoadChunk, LoadResult>;
  while (!(res = gen.next()).done) {
    const v = res.value;
    if (v.kind === 'sheetStart') sheets[v.index] = { ...(v.sheet as Sheet), cellData: {}, rowData: {} };
    else {
      Object.assign(sheets[v.index].cellData, v.cellData);
      if (v.rowData) sheets[v.index].rowData = v.rowData as never;
    }
  }
  return { sheets, result: res.value as LoadResult & { defaultStyle: { ff: string; fs: number }; warnings: string[] } };
}

const SAMPLES = [1234.5678, -1234.5678, 0, 46345.6283, 1.75];
/**
 * Where LibreOffice's .xlsx is not the reference: LibreOffice 24.2 writes the Default style's theme
 * colour instead of a cell's own font colour into its .xlsx (it shows the right colour itself). These
 * cells are checked against the colour the test file was made with (make-ods.mjs).
 */
const FONT_COLOUR: Record<string, string> = { 'Looks!R6C7': '#dc2626', 'Looks!R6C12': '#2563eb', 'Looks!R8C5': '#ffffff', 'Looks!R8C7': '#ffffff' };
/** What a cell looks like, independent of how the reader spells it. */
function look(st: IStyleData | undefined, def: { ff?: string; fs?: number }, v: unknown) {
  const s = st ?? {};
  const rgb = (x?: { rgb?: string | null }) => (x?.rgb ? String(x.rgb).toLowerCase() : undefined);
  const bd: Record<string, string> = {};
  for (const [k, b] of Object.entries((s.bd ?? {}) as Record<string, { s: number; cl?: { rgb?: string } } | null>).sort()) if (b?.s) bd[k] = `${b.s} ${rgb(b.cl) ?? '#000000'}`;
  const pattern = s.n?.pattern && s.n.pattern !== 'General' ? s.n.pattern : undefined;
  const shown = (x: unknown) => (typeof x === 'number' ? numfmt.format(pattern ?? 'General', x) : String(x));
  const out: Record<string, unknown> = {
    ff: s.ff && s.ff !== def.ff ? s.ff : undefined,
    fs: s.fs && s.fs !== def.fs ? s.fs : undefined,
    bl: s.bl ? 1 : undefined,
    it: s.it ? 1 : undefined,
    ul: s.ul?.s ? 1 : undefined,
    st: s.st?.s ? 1 : undefined,
    cl: rgb(s.cl) === '#000000' ? undefined : rgb(s.cl),
    bg: rgb(s.bg),
    bd: Object.keys(bd).length ? bd : undefined,
    ht: s.ht || undefined,
    vt: s.vt && s.vt !== 3 ? s.vt : undefined,
    tb: s.tb === 3 ? 3 : undefined,
    tr: s.tr ? (s.tr.v ? 'v' : s.tr.a) || undefined : undefined,
    // number formats are compared by what they show, not by how they are written
    n: pattern ? [v, ...SAMPLES].map(shown).join(' | ') : undefined,
  };
  for (const k of Object.keys(out)) if (out[k] === undefined) delete out[k];
  return out;
}

describe('.ods formatting (L-03)', () => {
  const ods = loadOds(ODS);
  const ref = load(REF);
  const refStyles = ref.data.styles as Record<string, IStyleData>;
  const refDef = (ref.res as unknown as { defaultStyle: { ff: string; fs: number } }).defaultStyle;
  const odsStyles = ods.result.styles as Record<string, IStyleData>;
  const odsDef = ods.result.defaultStyle;
  const refSheets = ref.data.sheetOrder.map((id) => ref.data.sheets[id] as Sheet);

  it('the default font is the document’s', () => {
    expect(odsDef).toEqual({ ff: 'Calibri', fs: 11 });
    expect(ods.result.warnings[0]).toMatch(/OpenDocument file/);
  });

  it('every cell looks as in LibreOffice’s own .xlsx of the file: fonts, fills, borders, alignment, wrap, rotation, number formats', () => {
    let compared = 0;
    const diffs: string[] = [];
    refSheets.forEach((rs, i) => {
      const os = ods.sheets[i];
      expect(os.name).toBe(rs.name);
      const rows = new Set([...Object.keys(rs.cellData), ...Object.keys(os.cellData)].map(Number));
      for (const r of rows) {
        const cols = new Set([...Object.keys(rs.cellData[r] ?? {}), ...Object.keys(os.cellData[r] ?? {})].map(Number));
        for (const c of cols) {
          const a = rs.cellData[r]?.[c];
          const b = os.cellData[r]?.[c];
          const la = look(typeof a?.s === 'string' ? refStyles[a.s] : (a?.s as IStyleData), refDef, a?.v);
          const lb = look(typeof b?.s === 'string' ? odsStyles[b.s] : (b?.s as IStyleData), odsDef, b?.v ?? a?.v);
          compared++;
          const at = `${rs.name}!R${r + 1}C${c + 1}`;
          if (FONT_COLOUR[at]) {
            expect(lb.cl, at).toBe(FONT_COLOUR[at]);
            delete lb.cl;
          }
          if (typeof a?.v === 'number') expect(b?.v, `${at} value`).toBeCloseTo(a.v, 9);
          else expect(b?.v, `${at} value`).toEqual(a?.v);
          if (JSON.stringify(la) !== JSON.stringify(lb)) diffs.push(`${at}: xlsx ${JSON.stringify(la)}\n      ods  ${JSON.stringify(lb)}`);
        }
      }
    });
    expect(diffs, diffs.join('\n')).toEqual([]);
    expect(compared).toBeGreaterThan(150);
  });

  it('column widths, row heights, hidden rows and columns, merges', () => {
    refSheets.forEach((rs, i) => {
      const os = ods.sheets[i];
      const w = (s: Sheet, c: number) => s.columnData?.[c]?.w ?? s.defaultColumnWidth!;
      for (let c = 0; c < 17; c++) {
        // LibreOffice measures a character of the default font a little wider than Excel's 7 px,
        // so its own .xlsx widths (in characters) come out up to ~6 % narrower in Petak
        const ratio = w(os, c) / w(rs, c);
        expect(ratio, `${rs.name} column ${c + 1}: ods ${w(os, c)} px, xlsx ${w(rs, c)} px`).toBeGreaterThan(0.97);
        expect(ratio).toBeLessThan(1.08);
        expect(!!os.columnData?.[c]?.hd, `${rs.name} column ${c + 1} hidden`).toBe(!!rs.columnData?.[c]?.hd);
      }
      const h = (s: Sheet, r: number) => s.rowData?.[r]?.h ?? s.defaultRowHeight!;
      for (let r = 0; r < 45; r++) {
        // rows LibreOffice fits to their content (15 pt here) go into its .xlsx at 13.8 pt, its
        // standard height, for Excel to fit again; set heights must agree
        if (os.rowData?.[r] && os.rowData[r].ia === undefined) expect(os.rowData[r].h, `${rs.name} row ${r + 1} fitted height`).toBe(20);
        else expect(Math.abs(h(os, r) - h(rs, r)), `${rs.name} row ${r + 1}: ods ${h(os, r)} px, xlsx ${h(rs, r)} px`).toBeLessThanOrEqual(1);
        expect(!!os.rowData?.[r]?.hd, `${rs.name} row ${r + 1} hidden`).toBe(!!rs.rowData?.[r]?.hd);
      }
      const merges = (s: IWorksheetData) => JSON.stringify([...(s.mergeData ?? [])].map((m) => [m.startRow, m.startColumn, m.endRow, m.endColumn]).sort());
      expect(merges(os)).toBe(merges(rs));
    });
    // spot checks: the hidden column O, the 60 pt alignment row, the wide column P
    const s = ods.sheets[0];
    expect(s.columnData[14].hd).toBe(1);
    expect(s.rowData[3].h).toBe(80);
    expect(s.columnData[15].w).toBeGreaterThan(280);
  });

  it('borders map like LibreOffice’s: width and line style to Excel’s thirteen styles', () => {
    expect(odsBorder('0.74pt solid #000000')).toEqual({ s: 1, cl: { rgb: '#000000' } });
    expect(odsBorder('1.76pt solid #dc2626')?.s).toBe(8);
    expect(odsBorder('2.49pt solid #2563eb')?.s).toBe(13);
    expect(odsBorder('0.06pt solid #000000')?.s).toBe(2);
    expect(odsBorder('0.74pt dotted #000000')?.s).toBe(3);
    expect(odsBorder('0.74pt fine-dashed #000000')?.s).toBe(4);
    expect(odsBorder('1.76pt dashed #000000')?.s).toBe(9);
    expect(odsBorder('1.76pt fine-dashed #000000')?.s).toBe(12);
    expect(odsBorder('1.76pt double-thin #000000')?.s).toBe(7);
    expect(odsBorder('0.05cm solid #ff0000', '0.02cm 0.02cm 0.02cm')?.s).toBe(7);
    // LibreOffice's own presets: 1.5 pt is medium, 2.25 pt thick
    expect(odsBorder('1.5pt solid #000000')?.s).toBe(8);
    expect(odsBorder('2.25pt solid #000000')?.s).toBe(13);
    expect(odsBorder('none')).toBeUndefined();
    expect(odsBorder('0.74pt hidden #000000')).toBeUndefined();
    expect(lengthPt('2.54cm')).toBeCloseTo(72);
    expect(lengthPt('0.5in')).toBe(36);
  });

  it('number styles: conditions, colours, currency, elapsed time and text', () => {
    const styles = (xml: string) =>
      new Map(children(root(`<r xmlns:number="n" xmlns:style="s" xmlns:fo="f" xmlns:loext="l">${xml}</r>`)!.inner).map((e) => [/style:name="([^"]+)"/.exec(e.open)![1], e]));
    const m = styles(
      '<number:number-style style:name="P0" style:volatile="true"><number:number number:decimal-places="2" number:min-decimal-places="2" number:min-integer-digits="1" number:grouping="true"/></number:number-style>' +
        '<number:number-style style:name="N"><style:text-properties fo:color="#ff0000"/><number:text>-</number:text><number:number number:decimal-places="2" number:min-decimal-places="2" number:min-integer-digits="1" number:grouping="true"/><style:map style:condition="value()&gt;=0" style:apply-style-name="P0"/></number:number-style>' +
        '<number:time-style style:name="T" number:truncate-on-overflow="false"><number:hours/><number:text>:</number:text><number:minutes number:style="long"/></number:time-style>' +
        '<number:text-style style:name="X"><number:text>Code </number:text><number:text-content/></number:text-style>' +
        '<number:number-style style:name="C1"><number:number number:decimal-places="0" number:min-integer-digits="1"/></number:number-style>' +
        '<number:number-style style:name="C"><number:text>small </number:text><number:number number:decimal-places="0" number:min-integer-digits="1"/><style:map style:condition="value()&gt;100" style:apply-style-name="C1"/></number:number-style>' +
        '<number:number-style style:name="K"><number:number number:decimal-places="1" number:min-decimal-places="0" number:min-integer-digits="1" number:display-factor="1000"/><number:text>k</number:text></number:number-style>',
    );
    expect(odsNumberFormat('N', m)).toBe('#,##0.00;[Red]\\-#,##0.00');
    expect(numfmt.format(odsNumberFormat('N', m)!, -1234.5)).toBe('-1,234.50');
    expect(odsNumberFormat('T', m)).toBe('[h]:mm');
    expect(numfmt.format('[h]:mm', 1.75)).toBe('42:00');
    expect(odsNumberFormat('X', m)).toBe('"Code "@');
    expect(odsNumberFormat('C', m)).toBe('[>100]0;"small "0');
    expect(numfmt.format(odsNumberFormat('K', m)!, 12345)).toBe('12.3k');
    expect(odsNumberFormat('missing', m)).toBeUndefined();
  });

  it('frozen panes come from the view settings', () => {
    const settings =
      '<config:config-item-map-named config:name="Tables"><config:config-item-map-entry config:name="Data &amp; more">' +
      '<config:config-item config:name="HorizontalSplitMode" config:type="short">2</config:config-item><config:config-item config:name="VerticalSplitMode" config:type="short">2</config:config-item>' +
      '<config:config-item config:name="HorizontalSplitPosition" config:type="int">1</config:config-item><config:config-item config:name="VerticalSplitPosition" config:type="int">3</config:config-item>' +
      '</config:config-item-map-entry><config:config-item-map-entry config:name="Split">' +
      '<config:config-item config:name="HorizontalSplitMode" config:type="short">1</config:config-item><config:config-item config:name="HorizontalSplitPosition" config:type="int">4000</config:config-item>' +
      '</config:config-item-map-entry></config:config-item-map-named>';
    expect([...odsFreeze(settings)]).toEqual([['Data & more', { cols: 1, rows: 3 }]]);
  });

  it('parent styles, column defaults, whole-row and whole-column formats, big gaps, frozen panes', () => {
    const NS = 'xmlns:office="o" xmlns:style="s" xmlns:table="t" xmlns:text="x" xmlns:fo="f" xmlns:config="c"';
    const content = `<office:document-content ${NS}><office:automatic-styles>
      <style:style style:name="co1" style:family="table-column"><style:table-column-properties style:column-width="1in"/></style:style>
      <style:style style:name="co2" style:family="table-column"><style:table-column-properties style:column-width="0.75in"/></style:style>
      <style:style style:name="ro1" style:family="table-row"><style:table-row-properties style:row-height="0.25in" style:use-optimal-row-height="false"/></style:style>
      <style:style style:name="ro2" style:family="table-row"><style:table-row-properties style:row-height="0.5in" style:use-optimal-row-height="false"/></style:style>
      <style:style style:name="ce1" style:family="table-cell" style:parent-style-name="Heading"><style:table-cell-properties fo:background-color="#fde68a"/></style:style>
      <style:style style:name="ce2" style:family="table-cell" style:parent-style-name="Default"><style:text-properties fo:font-style="italic"/></style:style>
      <style:style style:name="yellow" style:family="table-cell" style:parent-style-name="Default"><style:table-cell-properties fo:background-color="#ffff00"/></style:style>
    </office:automatic-styles><office:body><office:spreadsheet><table:table table:name="S">
      <table:table-column table:style-name="co1" table:default-cell-style-name="ce2"/>
      <table:table-column-group><table:table-column table:style-name="co1" table:number-columns-repeated="300" table:visibility="collapse"/></table:table-column-group>
      <table:table-column table:style-name="co2" table:number-columns-repeated="16083"/>
      <table:table-row table:style-name="ro1"><table:table-cell office:value-type="float" office:value="1"><text:p>1</text:p></table:table-cell><table:table-cell table:style-name="ce1" office:value-type="string"><text:p>Title</text:p></table:table-cell><table:table-cell table:number-columns-repeated="16382"/></table:table-row>
      <table:table-row table:style-name="ro1"><table:table-cell table:style-name="yellow" table:number-columns-repeated="16384"/></table:table-row>
      <table:table-row table:style-name="ro1" table:number-rows-repeated="70000"><table:table-cell table:number-columns-repeated="16384"/></table:table-row>
      <table:table-row table:style-name="ro2"><table:table-cell table:style-name="yellow" office:value-type="string"><text:p>far down</text:p></table:table-cell><table:table-cell table:number-columns-repeated="16383"/></table:table-row>
      <table:table-row table:style-name="ro1" table:number-rows-repeated="978000"><table:table-cell table:number-columns-repeated="2"/><table:table-cell table:style-name="yellow"/><table:table-cell table:number-columns-repeated="16381"/></table:table-row>
    </table:table></office:spreadsheet></office:body></office:document-content>`;
    const styles = `<office:document-styles ${NS}><office:styles>
      <style:default-style style:family="table-cell"><style:text-properties style:font-name="Liberation Sans" fo:font-size="10pt"/></style:default-style>
      <style:style style:name="Default" style:family="table-cell"/>
      <style:style style:name="Heading" style:family="table-cell" style:parent-style-name="Default"><style:text-properties fo:font-size="14pt" fo:font-weight="bold"/></style:style>
    </office:styles></office:document-styles>`;
    const settings = `<office:document-settings ${NS}><config:config-item-map-named config:name="Tables"><config:config-item-map-entry config:name="S"><config:config-item config:name="VerticalSplitMode" config:type="short">2</config:config-item><config:config-item config:name="VerticalSplitPosition" config:type="int">1</config:config-item></config:config-item-map-entry></config:config-item-map-named></office:document-settings>`;
    const zip = zipSync({ 'content.xml': strToU8(content), 'styles.xml': strToU8(styles), 'settings.xml': strToU8(settings) });
    const info = readOdsFormatting(ab(zip))!;
    expect(info.defaultFont).toEqual({ name: 'Liberation Sans', size: 10 });
    const sh = info.sheets[0];
    const at = (r: number, c: number) => info.styles[sh.cells.get(r)?.get(c) ?? -1];
    // a cell without a style takes its column's default
    expect(at(0, 0)).toEqual({ it: 1 });
    // a style with a named parent: the parent's bold 14 pt and its own fill
    expect(at(0, 1)).toEqual({ fs: 14, bl: 1, bg: { rgb: '#fde68a' } });
    // a cell repeated to the end of the row is a whole-row format
    expect(info.styles[sh.rowStyles!.get(1)!]).toEqual({ bg: { rgb: '#ffff00' } });
    // a row after a gap of 70 000 rows is an ordinary row with its own height and look
    expect(at(70002, 0)).toEqual({ bg: { rgb: '#ffff00' } });
    expect(sh.rowHeights.get(70002)).toBe(36);
    // the last, repeated row is the rest of the sheet: its height the default, its look a column format
    expect(sh.defaultRowPt).toBe(18);
    expect(info.styles[sh.colStyles!.get(2)!]).toEqual({ bg: { rgb: '#ffff00' } });
    expect(sh.cells.size).toBe(2);
    // widths: the last column element is the default; a hidden group of 300 columns is hidden
    expect(sh.defaultColPx).toBe(72);
    expect(sh.colPx!.get(0)).toBe(96);
    expect(sh.hiddenCols.size).toBe(300);
    expect(sh.hiddenCols.has(300)).toBe(true);
    expect(sh.hiddenCols.has(301)).toBe(false);
    expect(sh.freeze).toEqual({ cols: 0, rows: 1 });
  });

  it('dates and times keep their exact value', () => {
    expect(odsDate('2026-11-19T15:04:45.12')).toBeCloseTo(46345.6283, 9);
    expect(odsDate('1899-12-31T18:00:00')).toBe(1.75);
    expect(odsDate('2026-11-19')).toBe(46345);
    expect(odsDuration('PT42H00M00S')).toBe(1.75);
    expect(odsDuration('-PT12H')).toBe(-0.5);
    expect(odsDuration('P1DT6H')).toBe(1.25);
    expect(odsDuration('PT')).toBeUndefined();
  });

  it('anything that is not an .ods package gives null instead of throwing', () => {
    expect(readOdsFormatting(new Uint8Array([1, 2, 3, 4]).buffer)).toBeNull();
    expect(readOdsFormatting(ab(fs.readFileSync('tests/fixtures/files/roster.xlsx')))).toBeNull();
  });
});
