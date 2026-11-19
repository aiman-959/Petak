/** CSV (and legacy .xls / .ods import) via SheetJS Community Edition. */
import * as XLSX from 'xlsx';
import type { ICellData, IWorksheetData } from '@univerjs/core';
import type { LoadChunk, LoadResult, ParsedMeta } from './types';
import { readBiffFormatting, type BiffInfo } from './biff';
import { readOdsFormatting } from './ods';
import { colWidthPx, maxDigitWidth } from './ooxml/read';

export interface ParsedSheetJs {
  wb: XLSX.WorkBook;
  meta: ParsedMeta;
  /** looks of a legacy .xls or an .ods file (SheetJS reads only their values) */
  biff?: BiffInfo | null;
}

export function parseWithSheetJs(buf: ArrayBuffer, name: string): ParsedSheetJs {
  const isCsv = /\.(csv|tsv|txt)$/i.test(name);
  const format = isCsv ? 'csv' : /\.ods$/i.test(name) ? 'ods' : 'xls';
  // the looks first, so their copy of a big file's XML is gone before SheetJS makes its own
  let biff: BiffInfo | null = null;
  if (format !== 'csv') {
    try {
      biff = format === 'xls' ? readBiffFormatting(buf) : readOdsFormatting(buf);
    } catch (e) {
      console.warn(format + ' formatting not read', e);
    }
  }
  const wb = isCsv
    ? XLSX.read(new TextDecoder().decode(buf), { type: 'string', dense: true, cellDates: false, raw: false })
    : XLSX.read(buf, { type: 'array', dense: true, cellFormula: true, cellNF: true });
  const sheets = wb.SheetNames.map((n, i) => {
    const ws = wb.Sheets[n];
    const ref = ws['!ref'] ? XLSX.utils.decode_range(ws['!ref']) : { e: { r: -1, c: -1 } };
    return { id: 'sheet' + (i + 1), name: isCsv ? name.replace(/\.[^.]+$/, '').slice(0, 31) || 'Sheet1' : n, rows: ref.e.r + 1, cols: ref.e.c + 1 };
  });
  // cells that only have a look (a filled or bordered blank cell) count towards the sheet size
  for (const sh of biff?.sheets ?? []) {
    const m = sheets[wb.SheetNames.indexOf(sh.name)];
    if (m) {
      m.rows = Math.max(m.rows, sh.maxRow + 1);
      m.cols = Math.max(m.cols, sh.maxCol + 1);
    }
  }
  return { wb, meta: { sheets, hasVba: false, format }, biff };
}

/**
 * CSV values that look like numbers but are identifiers (L-01): a leading zero (0012, 0123456789)
 * or more than 15 digits (card, account and phone numbers lose digits as numbers). Kept as text.
 */
export function csvKeepText(w: string): string | null {
  const t = w.trim();
  if (/^[+-]?0\d+$/.test(t)) return t;
  if (/^[+-]?\d{16,}$/.test(t)) return t;
  return null;
}
/** A date number format for CSV text SheetJS read as a date (it gives only the serial number). */
export function csvDateFormat(w: string): string | null {
  const t = w.trim();
  if (/^\d{4}-\d{1,2}-\d{1,2}$/.test(t)) return 'yyyy-mm-dd';
  if (/^\d{4}-\d{1,2}-\d{1,2}[ T]\d{1,2}:\d{2}(:\d{2})?$/.test(t)) return 'yyyy-mm-dd hh:mm' + (t.split(':').length > 2 ? ':ss' : '');
  if (/^\d{1,2}\/\d{1,2}\/\d{2,4}$/.test(t)) return 'm/d/yyyy';
  if (/^\d{1,2}:\d{2}(:\d{2})?$/.test(t)) return t.split(':').length > 2 ? 'h:mm:ss' : 'h:mm';
  return null;
}

export function* streamSheetJs(p: ParsedSheetJs, chunkRows: number): Generator<LoadChunk, LoadResult> {
  const total = p.meta.sheets.reduce((a, s) => a + s.rows, 0) || 1;
  let done = 0;
  const styles: Record<string, Record<string, unknown>> = {};
  const fmtIds = new Map<string, string>();
  const biff = p.biff ?? null;
  const isCsv = p.meta.format === 'csv';
  const isOds = p.meta.format === 'ods';
  const mdw = maxDigitWidth(biff?.defaultFont ?? {});
  /** style id for a cell: its .xls look (XF) combined with its number format */
  const styleId = (xf: number | undefined, z: string | undefined): string | undefined => {
    const look = xf !== undefined ? biff?.styles[xf] : null;
    if (!look && !z) return undefined;
    const key = (look ? 'x' + xf : '') + (z ? '|' + z : '');
    let id = fmtIds.get(key);
    if (!id) {
      id = 's' + fmtIds.size;
      fmtIds.set(key, id);
      // the file's own number format wins over the one SheetJS guessed
      styles[id] = { ...(z ? { n: { pattern: z } } : {}), ...(look ?? {}) };
    }
    return id;
  };
  for (let i = 0; i < p.wb.SheetNames.length; i++) {
    const ws = p.wb.Sheets[p.wb.SheetNames[i]];
    const m = p.meta.sheets[i];
    const bs = biff?.sheets.find((x) => x.name === p.wb.SheetNames[i]);
    const sheet: Omit<IWorksheetData, 'cellData'> = {
      id: m.id,
      name: m.name,
      rowCount: Math.max(1000, m.rows + 100),
      columnCount: Math.max(26, m.cols + 5),
      mergeData: (ws['!merges'] ?? []).map((x) => ({ startRow: x.s.r, endRow: x.e.r, startColumn: x.s.c, endColumn: x.e.c })),
      defaultColumnWidth: 88,
      defaultRowHeight: 24,
    } as never;
    const rowData: Record<number, Record<string, unknown>> = {};
    if (bs) {
      const sx = sheet as unknown as Record<string, unknown>;
      if (bs.defaultColChars) sx.defaultColumnWidth = Math.round(bs.defaultColChars * mdw + 5);
      if (bs.defaultColPx) sx.defaultColumnWidth = bs.defaultColPx;
      if (bs.defaultRowPt) sx.defaultRowHeight = Math.round((bs.defaultRowPt * 4) / 3);
      const columnData: Record<number, { w?: number; hd?: number; s?: string }> = {};
      for (const [c, chars] of bs.colWidths) columnData[c] = { w: colWidthPx(chars, mdw) };
      for (const [c, px] of bs.colPx ?? []) columnData[c] = { w: px };
      for (const c of bs.hiddenCols) columnData[c] = { ...(columnData[c] ?? {}), hd: 1 };
      for (const [c, xf] of bs.colStyles ?? []) {
        const sid = styleId(xf, undefined);
        if (sid) columnData[c] = { ...(columnData[c] ?? {}), s: sid };
      }
      sx.columnData = columnData;
      // a height LibreOffice fitted to the content may be fitted again; one set by hand stays
      for (const [r, pt] of bs.rowHeights) rowData[r] = { h: Math.round(((pt * 4) / 3) * 100) / 100, ...(bs.optimalRows?.has(r) ? {} : { ia: 0 }) };
      for (const r of bs.hiddenRows) rowData[r] = { ...(rowData[r] ?? {}), hd: 1 };
      for (const [r, xf] of bs.rowStyles ?? []) {
        const sid = styleId(xf, undefined);
        if (sid) rowData[r] = { ...(rowData[r] ?? {}), s: sid };
      }
      if (bs.freeze) sx.freeze = { xSplit: bs.freeze.cols, ySplit: bs.freeze.rows, startRow: bs.freeze.rows || -1, startColumn: bs.freeze.cols || -1 };
    }
    yield { kind: 'sheetStart', index: i, sheet };
    const dense = (ws as unknown as { '!data'?: XLSX.CellObject[][] })['!data'] ?? [];
    for (let r0 = 0; r0 < m.rows; r0 += chunkRows) {
      const chunk: Record<number, Record<number, ICellData>> = {};
      const r1 = Math.min(m.rows, r0 + chunkRows);
      for (let r = r0; r < r1; r++) {
        const row = dense[r] ?? [];
        const looks = bs?.cells.get(r);
        if (!row.length && !looks) continue;
        const out: Record<number, ICellData> = {};
        for (let c = 0; c < row.length; c++) {
          const cell = row[c];
          if (!cell) continue;
          const cd: ICellData = {};
          if (cell.f) cd.f = '=' + cell.f;
          const csvText = isCsv && cell.t === 'n' && typeof cell.w === 'string' ? csvKeepText(cell.w) : null;
          const csvDate = isCsv && cell.t === 'n' && typeof cell.w === 'string' && !csvText ? csvDateFormat(cell.w) : null;
          if (csvText !== null) {
            cd.v = csvText;
            cd.t = 1;
          } else if (cell.t === 'n') {
            cd.v = cell.v as number;
            cd.t = 2;
          } else if (cell.t === 'b') {
            cd.v = cell.v ? 1 : 0;
            cd.t = 3;
          } else if (cell.t === 'd') {
            cd.v = ((cell.v as Date).getTime() - Date.UTC(1899, 11, 30)) / 86400000;
            cd.t = 2;
          } else if (cell.t === 'e') {
            cd.v = cell.w ?? '#VALUE!';
            cd.t = 1;
          } else if (cell.v !== undefined) {
            cd.v = String(cell.v);
            cd.t = 1;
          }
          const exact = bs?.dates?.get(r)?.get(c);
          if (exact !== undefined && (cell.t === 'n' || cell.t === 'd')) cd.v = exact;
          // SheetJS's guess at an .ods number format is only used when the looks could not be read
          const z = csvText !== null || (isOds && biff) ? undefined : csvDate ?? (cell.z && cell.z !== 'General' && cell.t === 'n' ? String(cell.z) : undefined);
          const sid = styleId(looks?.get(c), z);
          if (sid) cd.s = sid;
          if (cd.v !== undefined || cd.f || cd.s) out[c] = cd;
        }
        // cells with a look but no value (filled, bordered)
        if (looks)
          for (const [c, xf] of looks) {
            if (out[c]) continue;
            const sid = styleId(xf, undefined);
            if (sid) out[c] = { s: sid };
          }
        chunk[r] = out;
      }
      done += r1 - r0;
      yield { kind: 'rows', index: i, cellData: chunk, done, total, ...(r1 >= m.rows && Object.keys(rowData).length ? { rowData } : {}) };
    }
  }
  const base: LoadResult = { styles: styles as never, sheetOrder: p.meta.sheets.map((s) => s.id), cf: [], meta: null, vba: null, name: '' };
  if (!biff) return base;
  return {
    ...base,
    defaultStyle: isOds ? { ff: biff.defaultFont.name ?? 'Liberation Sans', fs: biff.defaultFont.size ?? 10 } : { ff: biff.defaultFont.name ?? 'Arial', fs: biff.defaultFont.size ?? 10 },
    warnings: [
      isOds
        ? 'OpenDocument file: charts, pictures, comments and conditional formats are not shown. Save writes a new .xlsx.'
        : 'Old .xls format: charts, pictures and conditional formats are not shown. Save writes a new .xlsx.',
    ],
  } as LoadResult;
}

/** Active sheet → CSV text (computed values, not formulas). */
export function sheetToCsv(cellData: Record<number, Record<number, ICellData>>, sep = ','): string {
  const aoa: unknown[][] = [];
  for (const rk in cellData) {
    const row = cellData[rk];
    const out: unknown[] = (aoa[+rk] = []);
    for (const ck in row) {
      const c = row[ck];
      let v = c?.v ?? (c?.p?.body?.dataStream?.replace(/\r?\n$/, '') || null);
      if (c?.t === 3) v = v === 1 || v === true ? 'TRUE' : 'FALSE';
      out[+ck] = v;
    }
  }
  for (let i = 0; i < aoa.length; i++) if (!aoa[i]) aoa[i] = [];
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  return XLSX.utils.sheet_to_csv(ws, { FS: sep });
}
