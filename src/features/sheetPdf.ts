/**
 * Print / PDF of a sheet (M-09).
 *
 * Each page is drawn on a canvas (cell fills, grid lines, text in the cell's font, alignment,
 * borders, merged cells) and the pages are written as images into a PDF. Drawing through the
 * canvas means every script prints (Chinese, Japanese, Korean, Cyrillic…) in the same fonts as
 * on screen, without embedding font files in the PDF. The frozen rows repeat at the top of every
 * page, like Excel's "rows to repeat at top".
 *
 * The page arithmetic and the PDF writer are pure and unit-tested; the painter needs a canvas.
 */
import type { IStyleData } from '@univerjs/core';

export type Paper = 'a4' | 'letter';
export interface PageSetup {
  paper: Paper;
  landscape: boolean;
  /** shrink so that all used columns fit the page width (never enlarges) */
  fitWidth: boolean;
  gridlines: boolean;
}
export const defaultSetup = (): PageSetup => ({ paper: 'a4', landscape: false, fitWidth: true, gridlines: true });

/** paper size in PDF points (1/72 in) */
export const PAPER: Record<Paper, [number, number]> = { a4: [595.28, 841.89], letter: [612, 792] };
const MARGIN_PT = 36; // 0.5 in
const FOOTER_PT = 18;
const PX_TO_PT = 0.75; // 96 px = 72 pt: 100 % prints at the size Excel prints

export interface Layout {
  /** page size in pt */
  pageW: number;
  pageH: number;
  /** sheet px → pt */
  scale: number;
  /** rows on each page (title rows not included) */
  pages: Array<{ rows: number[]; cols: number[] }>;
  titleRows: number[];
}

/**
 * Split the used range into pages. `colW` / `rowH` are sheet px of the visible columns / rows,
 * `titleRows` (indexes) repeat on every page. Columns that do not fit one page width (when not
 * fitting to width) continue on further pages, Excel's "down, then over" order.
 */
export function layoutPages(rows: number[], cols: number[], rowH: (r: number) => number, colW: (c: number) => number, titleRows: number[], setup: PageSetup): Layout {
  let [pageW, pageH] = PAPER[setup.paper];
  if (setup.landscape) [pageW, pageH] = [pageH, pageW];
  const availW = pageW - 2 * MARGIN_PT;
  const availH = pageH - 2 * MARGIN_PT - FOOTER_PT;
  const totalW = cols.reduce((t, c) => t + colW(c), 0) * PX_TO_PT;
  const scale = PX_TO_PT * (setup.fitWidth && totalW > availW ? availW / totalW : 1);
  // column bands
  const bands: number[][] = [];
  let band: number[] = [];
  let w = 0;
  for (const c of cols) {
    const cw = colW(c) * scale;
    if (band.length && w + cw > availW + 0.01) {
      bands.push(band);
      band = [];
      w = 0;
    }
    band.push(c);
    w += cw;
  }
  if (band.length || !bands.length) bands.push(band);
  // row blocks under the repeated title rows
  const titles = titleRows;
  const titleH = titles.reduce((t, r) => t + rowH(r), 0) * scale;
  const body = rows.filter((r) => !titles.includes(r));
  const blocks: number[][] = [];
  let block: number[] = [];
  let h = titleH;
  for (const r of body) {
    const rh = rowH(r) * scale;
    if (block.length && h + rh > availH + 0.01) {
      blocks.push(block);
      block = [];
      h = titleH;
    }
    block.push(r);
    h += rh;
  }
  if (block.length || !blocks.length) blocks.push(block);
  const pages: Layout['pages'] = [];
  for (const b of bands) for (const rr of blocks) pages.push({ rows: rr, cols: b });
  return { pageW, pageH, scale, pages, titleRows: titles };
}

// ---------------------------------------------------------------- PDF writer
/** A PDF whose pages are JPEG images filling the page (pure: unit-tested). */
export function pdfFromJpegs(pages: Array<{ jpeg: Uint8Array; w: number; h: number; pageW: number; pageH: number }>, title = 'Petak'): Uint8Array {
  const enc = new TextEncoder();
  const parts: Uint8Array[] = [];
  const offsets: number[] = [];
  let size = 0;
  const push = (p: Uint8Array | string) => {
    const b = typeof p === 'string' ? enc.encode(p) : p;
    parts.push(b);
    size += b.length;
  };
  const obj = (n: number, body: () => void) => {
    offsets[n] = size;
    push(`${n} 0 obj\n`);
    body();
    push('\nendobj\n');
  };
  const n2 = (v: number) => +v.toFixed(2);
  push('%PDF-1.4\n%\xE2\xE3\xCF\xD3\n'.replace(/[\x80-\xff]/g, '~'));
  const nPages = pages.length;
  // 1 catalog, 2 pages, 3 info, then per page: page, contents, image
  const pageObj = (i: number) => 4 + i * 3;
  obj(1, () => push('<< /Type /Catalog /Pages 2 0 R >>'));
  obj(2, () => push(`<< /Type /Pages /Count ${nPages} /Kids [${pages.map((_, i) => `${pageObj(i)} 0 R`).join(' ')}] >>`));
  obj(3, () => push(`<< /Producer (Petak) /Title (${title.replace(/[\\()]/g, (m) => '\\' + m).replace(/[^\x20-\x7e]/g, '?')}) >>`));
  pages.forEach((p, i) => {
    const n = pageObj(i);
    obj(n, () => push(`<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${n2(p.pageW)} ${n2(p.pageH)}] /Resources << /XObject << /Im${i} ${n + 2} 0 R >> >> /Contents ${n + 1} 0 R >>`));
    const content = `q ${n2(p.pageW)} 0 0 ${n2(p.pageH)} 0 0 cm /Im${i} Do Q`;
    obj(n + 1, () => push(`<< /Length ${content.length} >>\nstream\n${content}\nendstream`));
    obj(n + 2, () => {
      push(`<< /Type /XObject /Subtype /Image /Width ${p.w} /Height ${p.h} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length ${p.jpeg.length} >>\nstream\n`);
      push(p.jpeg);
      push('\nendstream');
    });
  });
  const xref = size;
  const count = 4 + nPages * 3;
  let x = `xref\n0 ${count}\n0000000000 65535 f \n`;
  for (let i = 1; i < count; i++) x += String(offsets[i]).padStart(10, '0') + ' 00000 n \n';
  push(x);
  push(`trailer\n<< /Size ${count} /Root 1 0 R /Info 3 0 R >>\nstartxref\n${xref}\n%%EOF\n`);
  const out = new Uint8Array(size);
  let at = 0;
  for (const p of parts) {
    out.set(p, at);
    at += p.length;
  }
  return out;
}

// ---------------------------------------------------------------- painter
export interface SheetSource {
  name: string;
  rows: number[];
  cols: number[];
  rowH: (r: number) => number;
  colW: (c: number) => number;
  titleRows: number[];
  text: (r: number, c: number) => string;
  style: (r: number, c: number) => IStyleData | null;
  /** merged ranges (inclusive) */
  merges: Array<{ r0: number; c0: number; r1: number; c1: number }>;
  defaultFont: { ff: string; fs: number };
}

const DPI_SCALE = 150 / 72; // canvas px per pt: 150 dpi, sharp enough for text, small files
const BORDER_W: Record<number, number> = { 1: 1, 2: 0.5, 3: 1, 4: 1, 5: 1, 6: 1, 7: 2.5, 8: 2, 9: 2, 10: 2, 11: 2, 12: 2, 13: 3 };

function cssColor(c: unknown, fallback: string) {
  const rgb = (c as { rgb?: string } | undefined)?.rgb;
  return rgb && /^#?[0-9a-f]{6}$/i.test(rgb) ? (rgb.startsWith('#') ? rgb : '#' + rgb) : fallback;
}

/** Draw every page and write the PDF. */
export async function sheetToPdf(src: SheetSource, setup: PageSetup): Promise<{ pdf: Uint8Array; pages: number }> {
  const L = layoutPages(src.rows, src.cols, src.rowH, src.colW, src.titleRows, setup);
  const k = DPI_SCALE;
  const W = Math.round(L.pageW * k);
  const H = Math.round(L.pageH * k);
  const canvas = document.createElement('canvas');
  canvas.width = W;
  canvas.height = H;
  const ctx = canvas.getContext('2d')!;
  const mergeAt = new Map<string, { r0: number; c0: number; r1: number; c1: number }>();
  const covered = new Set<string>();
  for (const m of src.merges) {
    mergeAt.set(`${m.r0},${m.c0}`, m);
    for (let r = m.r0; r <= m.r1; r++) for (let c = m.c0; c <= m.c1; c++) if (r !== m.r0 || c !== m.c0) covered.add(`${r},${c}`);
  }
  const out: Array<{ jpeg: Uint8Array; w: number; h: number; pageW: number; pageH: number }> = [];
  for (let pi = 0; pi < L.pages.length; pi++) {
    const page = L.pages[pi];
    ctx.setTransform(1, 0, 0, 1, 0, 0);
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, W, H);
    // sheet px → canvas px
    const s = L.scale * k;
    ctx.setTransform(s, 0, 0, s, MARGIN_PT * k, MARGIN_PT * k);
    const rows = [...L.titleRows, ...page.rows];
    const ys = new Map<number, number>();
    let y = 0;
    for (const r of rows) {
      ys.set(r, y);
      y += src.rowH(r);
    }
    const xs = new Map<number, number>();
    let x = 0;
    for (const c of page.cols) {
      xs.set(c, x);
      x += src.colW(c);
    }
    const totalW = x;
    const totalH = y;
    // the extent of a cell, merged cells clipped to this page
    const rect = (r: number, c: number) => {
      const m = mergeAt.get(`${r},${c}`);
      const x0 = xs.get(c)!;
      const y0 = ys.get(r)!;
      if (!m) return { x: x0, y: y0, w: src.colW(c), h: src.rowH(r) };
      let w = 0;
      for (let cc = m.c0; cc <= m.c1; cc++) if (xs.has(cc)) w += src.colW(cc);
      let h = 0;
      for (let rr = m.r0; rr <= m.r1; rr++) if (ys.has(rr)) h += src.rowH(rr);
      return { x: x0, y: y0, w, h };
    };
    // 1. grid lines
    if (setup.gridlines) {
      ctx.strokeStyle = '#d0d7de';
      ctx.lineWidth = (0.4 * k) / s; // 0.4 pt
      ctx.beginPath();
      for (const c of page.cols) {
        const gx = xs.get(c)! + src.colW(c);
        ctx.moveTo(gx, 0);
        ctx.lineTo(gx, totalH);
      }
      for (const r of rows) {
        const gy = ys.get(r)! + src.rowH(r);
        ctx.moveTo(0, gy);
        ctx.lineTo(totalW, gy);
      }
      ctx.moveTo(0, 0);
      ctx.lineTo(totalW, 0);
      ctx.moveTo(0, 0);
      ctx.lineTo(0, totalH);
      ctx.stroke();
    }
    // 2. fills: a fill hides the grid lines under it, a merged cell hides those inside it
    for (const r of rows)
      for (const c of page.cols) {
        if (covered.has(`${r},${c}`)) continue;
        const st = src.style(r, c);
        const merged = mergeAt.has(`${r},${c}`);
        if (!st?.bg && !merged) continue;
        const b = rect(r, c);
        ctx.fillStyle = cssColor(st?.bg, '#ffffff');
        ctx.fillRect(b.x, b.y, b.w, b.h);
      }
    // 3. text
    for (const r of rows)
      for (const c of page.cols) {
        if (covered.has(`${r},${c}`)) continue;
        const t = src.text(r, c);
        if (!t) continue;
        const st = src.style(r, c) ?? {};
        const b = rect(r, c);
        const fs = st.fs ?? src.defaultFont.fs;
        const ff = st.ff ?? src.defaultFont.ff;
        ctx.save();
        ctx.beginPath();
        ctx.rect(b.x, b.y, b.w, b.h);
        ctx.clip();
        ctx.font = `${st.it ? 'italic ' : ''}${st.bl ? 'bold ' : ''}${(fs * 4) / 3}px "${ff}", "Noto Sans CJK SC", sans-serif`;
        ctx.fillStyle = cssColor(st.cl, '#000000');
        const numeric = /^[-+(]?[\d$€£¥.,]/.test(t) && !/[A-Za-zÀ-￿]{2}/.test(t);
        const ht = (st.ht as number | undefined) ?? (numeric ? 3 : 1);
        const vt = (st.vt as number | undefined) ?? 3;
        const pad = 3;
        const lines = st.tb === 3 ? wrap(ctx, t, b.w - 2 * pad) : t.split('\n');
        const lh = ((fs * 4) / 3) * 1.25;
        const blockH = lh * lines.length;
        let ty = vt === 1 ? b.y + pad : vt === 2 ? b.y + (b.h - blockH) / 2 : b.y + b.h - pad - blockH;
        ctx.textBaseline = 'top';
        for (const line of lines) {
          const tw = ctx.measureText(line).width;
          const tx = ht === 2 ? b.x + (b.w - tw) / 2 : ht === 3 ? b.x + b.w - pad - tw : b.x + pad;
          if (st.ul) ctx.fillRect(tx, ty + lh * 0.82, tw, Math.max(0.6, fs / 14));
          if (st.st) ctx.fillRect(tx, ty + lh * 0.45, tw, Math.max(0.6, fs / 14));
          ctx.fillText(line, tx, ty + lh * 0.1);
          ty += lh;
        }
        ctx.restore();
      }
    // 4. borders
    for (const r of rows)
      for (const c of page.cols) {
        if (covered.has(`${r},${c}`)) continue;
        const bd = src.style(r, c)?.bd as Record<string, { s?: number; cl?: { rgb?: string } } | undefined> | undefined;
        if (!bd) continue;
        const b = rect(r, c);
        const side = (key: string, x0: number, y0: number, x1: number, y1: number) => {
          const e = bd[key];
          if (!e || !e.s) return;
          ctx.strokeStyle = cssColor(e.cl, '#000000');
          ctx.lineWidth = ((BORDER_W[e.s] ?? 1) * 0.75 * k) / s; // px of a 96-dpi screen → pt → canvas px → sheet units
          ctx.setLineDash(e.s === 3 ? [1, 1] : e.s === 4 ? [3, 2] : []);
          ctx.beginPath();
          ctx.moveTo(x0, y0);
          ctx.lineTo(x1, y1);
          ctx.stroke();
        };
        side('t', b.x, b.y, b.x + b.w, b.y);
        side('b', b.x, b.y + b.h, b.x + b.w, b.y + b.h);
        side('l', b.x, b.y, b.x, b.y + b.h);
        side('r', b.x + b.w, b.y, b.x + b.w, b.y + b.h);
        ctx.setLineDash([]);
      }
    // footer
    ctx.setTransform(1, 0, 0, 1, 0, 0);
    ctx.fillStyle = '#555555';
    ctx.font = `${Math.round(9 * k)}px sans-serif`;
    ctx.textBaseline = 'alphabetic';
    const foot = `${src.name} — ${pi + 1} / ${L.pages.length}`;
    ctx.fillText(foot, (W - ctx.measureText(foot).width) / 2, H - (MARGIN_PT / 2) * k);
    const blob = await new Promise<Blob>((res, rej) => canvas.toBlob((b) => (b ? res(b) : rej(new Error('Could not draw the page'))), 'image/jpeg', 0.9));
    out.push({ jpeg: new Uint8Array(await blob.arrayBuffer()), w: W, h: H, pageW: L.pageW, pageH: L.pageH });
  }
  return { pdf: pdfFromJpegs(out, src.name), pages: out.length };
}

function wrap(ctx: CanvasRenderingContext2D, text: string, width: number): string[] {
  const out: string[] = [];
  for (const para of text.split('\n')) {
    let line = '';
    for (const word of para.split(/(\s+)/)) {
      const next = line + word;
      if (line && ctx.measureText(next).width > width) {
        out.push(line.trimEnd());
        line = word.trimStart();
      } else line = next;
    }
    out.push(line);
  }
  return out;
}
