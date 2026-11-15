/**
 * Formatting of legacy .xls files (BIFF8, Excel 97–2003) — H-03.
 *
 * SheetJS Community Edition reads the values, formulas, merges and number formats of an .xls
 * file but none of its looks. This reads the rest straight from the workbook stream: fonts,
 * fills, borders, alignment (XF records), column widths, row heights, hidden rows / columns
 * and frozen panes. Pure and unit-tested; anything unexpected returns null and the file
 * simply opens unformatted as before.
 */
import * as XLSX from 'xlsx';
import type { IStyleData } from '@univerjs/core';

export interface BiffSheet {
  name: string;
  /** row → column → index into `styles` */
  cells: Map<number, Map<number, number>>;
  /** column → width in characters (1/256 units already divided) */
  colWidths: Map<number, number>;
  hiddenCols: Set<number>;
  /** row → height in points */
  rowHeights: Map<number, number>;
  hiddenRows: Set<number>;
  defaultColChars?: number;
  defaultRowPt?: number;
  freeze?: { cols: number; rows: number };
  maxRow: number;
  maxCol: number;
}
export interface BiffInfo {
  /** one style per cell XF; null = nothing to apply */
  styles: Array<IStyleData | null>;
  defaultFont: { name?: string; size?: number };
  sheets: BiffSheet[];
}

const DEFAULT_PALETTE =
  '000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF 800000 008000 000080 808000 800080 008080 C0C0C0 808080 9999FF 993366 FFFFCC CCFFFF 660066 FF8080 0066CC CCCCFF 000080 FF00FF FFFF00 00FFFF 800080 800000 008080 0000FF 00CCFF CCFFFF CCFFCC FFFF99 99CCFF FF99CC CC99FF FFCC99 3366FF 33CCCC 99CC00 FFCC00 FF9900 FF6600 666699 969696 003366 339966 003300 333300 993300 993366 333399 333333'
    .split(' ')
    .map((c) => '#' + c.toLowerCase());
const FIXED = ['#000000', '#ffffff', '#ff0000', '#00ff00', '#0000ff', '#ffff00', '#ff00ff', '#00ffff'];
/** BIFF border style → Univer BorderStyleTypes */
const BORDER = [0, 1, 8, 4, 3, 13, 7, 2, 9, 5, 10, 6, 11, 12];

interface Font {
  name: string;
  size: number;
  bold: boolean;
  italic: boolean;
  strike: boolean;
  underline: boolean;
  color: number;
}

/** The workbook stream of an .xls compound file, or null. */
function workbookStream(buf: ArrayBuffer): Uint8Array | null {
  const cfb = XLSX.CFB.read(new Uint8Array(buf), { type: 'array' });
  const entry = XLSX.CFB.find(cfb, 'Workbook') ?? XLSX.CFB.find(cfb, 'Book');
  const content = entry?.content as unknown as ArrayLike<number> | undefined;
  return content && content.length ? (content instanceof Uint8Array ? content : new Uint8Array(content)) : null;
}

export function readBiffFormatting(buf: ArrayBuffer): BiffInfo | null {
  let data: Uint8Array | null;
  try {
    data = workbookStream(buf);
  } catch {
    return null;
  }
  if (!data) return null;
  const dv = new DataView(data.buffer, data.byteOffset, data.byteLength);
  const u16 = (o: number) => dv.getUint16(o, true);
  const u32 = (o: number) => dv.getUint32(o, true);
  const str = (o: number, cch: number): string => {
    // BIFF8 string body: option flags byte, then 8-bit or 16-bit characters
    const wide = data![o] & 1;
    let s = '';
    for (let i = 0; i < cch; i++) s += String.fromCharCode(wide ? u16(o + 1 + i * 2) : data![o + 1 + i]);
    return s;
  };

  const fonts: Font[] = [];
  const palette = [...DEFAULT_PALETTE];
  const xfs: Array<{ font: number; style: boolean; align: number; b1: number; b2: number; fill: number }> = [];
  const bounds: Array<{ pos: number; name: string; type: number }> = [];
  let pos = 0;
  // ---- workbook globals
  if (data.length < 8 || u16(0) !== 0x0809 || u16(4) !== 0x0600) return null; // BIFF8 only
  while (pos + 4 <= data.length) {
    const id = u16(pos);
    const len = u16(pos + 2);
    const o = pos + 4;
    if (o + len > data.length) return null;
    if (id === 0x002f) return null; // FILEPASS: encrypted
    if (id === 0x0031 && len >= 16) {
      const grbit = u16(o + 2);
      fonts.push({ size: u16(o) / 20, italic: !!(grbit & 2), strike: !!(grbit & 8), color: u16(o + 4), bold: u16(o + 6) >= 600, underline: data[o + 10] !== 0, name: str(o + 15, data[o + 14]) });
    } else if (id === 0x0092) {
      const n = u16(o);
      for (let i = 0; i < n && i < 56 && o + 2 + i * 4 + 3 <= o + len; i++) {
        const p = o + 2 + i * 4;
        palette[i] = '#' + [data[p], data[p + 1], data[p + 2]].map((v) => v.toString(16).padStart(2, '0')).join('');
      }
    } else if (id === 0x00e0 && len >= 20) {
      xfs.push({ font: u16(o), style: !!(u16(o + 4) & 4), align: data[o + 6], b1: u32(o + 10), b2: u32(o + 14), fill: u16(o + 18) });
    } else if (id === 0x0085 && len >= 8) {
      bounds.push({ pos: u32(o), type: data[o + 5], name: str(o + 7, data[o + 6]) });
    } else if (id === 0x000a) break; // EOF of the globals
    pos = o + len;
  }
  if (!xfs.length || !fonts.length) return null;

  const color = (i: number, auto?: string): string | undefined => (i < 8 ? FIXED[i] : i < 64 ? palette[i - 8] : i === 64 ? '#000000' : i === 65 ? '#ffffff' : auto);
  const fontAt = (i: number): Font | undefined => fonts[i >= 4 ? i - 1 : i]; // index 4 does not exist
  const f0 = fontAt(0)!;
  const styles: Array<IStyleData | null> = xfs.map((x) => {
    if (x.style) return null;
    const s: IStyleData = {};
    const f = fontAt(x.font);
    if (f) {
      if (f.name && f.name !== f0.name) s.ff = f.name;
      if (f.size && f.size !== f0.size) s.fs = f.size;
      if (f.bold) s.bl = 1;
      if (f.italic) s.it = 1;
      if (f.strike) s.st = { s: 1 };
      if (f.underline) s.ul = { s: 1 };
      const c = color(f.color);
      if (c && c !== '#000000') s.cl = { rgb: c };
    }
    const h = x.align & 7;
    if (h === 1 || h === 4) s.ht = 1 as never;
    else if (h === 2 || h === 6) s.ht = 2 as never;
    else if (h === 3) s.ht = 3 as never;
    else if (h === 5 || h === 7) s.ht = 4 as never;
    const v = (x.align >> 4) & 7;
    if (v === 0) s.vt = 1 as never;
    else if (v === 1 || v === 3 || v === 4) s.vt = 2 as never;
    if (x.align & 8) s.tb = 3 as never;
    // fill: pattern in the top 6 bits of the second border word
    const pattern = (x.b2 >>> 26) & 0x3f;
    if (pattern) {
      const fg = color(x.fill & 0x7f);
      const bg = color((x.fill >> 7) & 0x7f);
      const c = pattern === 1 ? fg : (fg ?? bg);
      if (c) s.bg = { rgb: c };
    }
    const bd: Record<string, unknown> = {};
    const side = (key: string, st: number, ci: number) => {
      if (st > 0 && st < BORDER.length) bd[key] = { s: BORDER[st], cl: { rgb: color(ci, '#000000') ?? '#000000' } };
    };
    side('l', x.b1 & 15, (x.b1 >>> 16) & 0x7f);
    side('r', (x.b1 >>> 4) & 15, (x.b1 >>> 23) & 0x7f);
    side('t', (x.b1 >>> 8) & 15, x.b2 & 0x7f);
    side('b', (x.b1 >>> 12) & 15, (x.b2 >>> 7) & 0x7f);
    if (Object.keys(bd).length) s.bd = bd as IStyleData['bd'];
    return Object.keys(s).length ? s : null;
  });

  // ---- worksheets
  const sheets: BiffSheet[] = [];
  for (const b of bounds) {
    if (b.type !== 0) continue; // charts, macro sheets
    const sh: BiffSheet = { name: b.name, cells: new Map(), colWidths: new Map(), hiddenCols: new Set(), rowHeights: new Map(), hiddenRows: new Set(), maxRow: -1, maxCol: -1 };
    const put = (r: number, c: number, xf: number) => {
      if (xf >= styles.length || !styles[xf]) return;
      let row = sh.cells.get(r);
      if (!row) sh.cells.set(r, (row = new Map()));
      row.set(c, xf);
      if (r > sh.maxRow) sh.maxRow = r;
      if (c > sh.maxCol) sh.maxCol = c;
    };
    let frozen = false;
    let pane: { x: number; y: number } | null = null;
    pos = b.pos;
    if (pos + 4 > data.length || u16(pos) !== 0x0809) continue;
    let depth = 0;
    while (pos + 4 <= data.length) {
      const id = u16(pos);
      const len = u16(pos + 2);
      const o = pos + 4;
      if (o + len > data.length) break;
      if (id === 0x0809) depth++;
      else if (id === 0x000a) {
        if (--depth <= 0) break;
      } else if (depth === 1) {
        switch (id) {
          case 0x0201: // BLANK
          case 0x0203: // NUMBER
          case 0x027e: // RK
          case 0x00fd: // LABELSST
          case 0x0204: // LABEL
          case 0x0205: // BOOLERR
          case 0x0006: // FORMULA
            if (len >= 6) put(u16(o), u16(o + 2), u16(o + 4));
            break;
          case 0x00be: {
            // MULBLANK: row, first column, xf…, last column
            const r = u16(o);
            const c0 = u16(o + 2);
            for (let i = 0; o + 4 + i * 2 + 2 <= o + len - 2; i++) put(r, c0 + i, u16(o + 4 + i * 2));
            break;
          }
          case 0x00bd: {
            // MULRK: row, first column, (xf, rk)…, last column
            const r = u16(o);
            const c0 = u16(o + 2);
            for (let i = 0; o + 4 + i * 6 + 6 <= o + len - 2; i++) put(r, c0 + i, u16(o + 4 + i * 6));
            break;
          }
          case 0x007d: {
            // COLINFO
            const first = u16(o);
            const last = Math.min(u16(o + 2), 255);
            const chars = u16(o + 4) / 256;
            const hidden = !!(u16(o + 8) & 1);
            for (let c = first; c <= last; c++) {
              sh.colWidths.set(c, chars);
              if (hidden) sh.hiddenCols.add(c);
            }
            break;
          }
          case 0x0208: {
            // ROW
            const r = u16(o);
            const h = u16(o + 6);
            const flags = len >= 16 ? u32(o + 12) : 0;
            if (!(h & 0x8000)) sh.rowHeights.set(r, (h & 0x7fff) / 20);
            if (flags & 0x20) sh.hiddenRows.add(r);
            break;
          }
          case 0x0055:
            sh.defaultColChars = u16(o);
            break;
          case 0x0225:
            if (len >= 4) sh.defaultRowPt = u16(o + 2) / 20;
            break;
          case 0x023e:
            frozen = !!(u16(o) & 8);
            break;
          case 0x0041:
            pane = { x: u16(o), y: u16(o + 2) };
            break;
        }
      }
      pos = o + len;
    }
    if (frozen && pane && (pane.x || pane.y)) sh.freeze = { cols: pane.x, rows: pane.y };
    sheets.push(sh);
  }
  return { styles, defaultFont: { name: f0.name || undefined, size: f0.size || undefined }, sheets };
}
