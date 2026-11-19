import { describe, expect, it } from 'vitest';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { PAPER, defaultSetup, layoutPages, pdfFromJpegs } from '../../src/features/sheetPdf';

const range = (n: number) => [...Array(n).keys()];

describe('print layout (M-09)', () => {
  it('fits the columns to the page width and splits rows into pages, repeating the title rows', () => {
    // 10 columns of 100 px = 750 pt, wider than A4 portrait (523 pt usable): scaled down
    const L = layoutPages(range(100), range(10), () => 20, () => 100, [0, 1], defaultSetup());
    expect(L.pageW).toBe(PAPER.a4[0]);
    expect(L.scale).toBeCloseTo((595.28 - 72) / 1000, 4);
    expect(L.pages.every((p) => p.cols.length === 10)).toBe(true);
    // all body rows exactly once, none of the title rows
    expect(L.pages.flatMap((p) => p.rows)).toEqual(range(100).slice(2));
    expect(L.titleRows).toEqual([0, 1]);
    // every page holds what fits under the title rows
    const avail = 841.89 - 72 - 18;
    for (const p of L.pages) expect((p.rows.length + 2) * 20 * L.scale).toBeLessThanOrEqual(avail + 0.01);
  });
  it('without fitting, wide sheets continue on further pages (down, then over); landscape is wider', () => {
    const L = layoutPages(range(10), range(10), () => 20, () => 100, [], { ...defaultSetup(), fitWidth: false });
    expect(L.scale).toBe(0.75);
    expect(L.pages.map((p) => p.cols)).toEqual([range(6), range(10).slice(6)]);
    const land = layoutPages(range(10), range(10), () => 20, () => 100, [], { ...defaultSetup(), fitWidth: false, landscape: true });
    expect(land.pageW).toBeGreaterThan(land.pageH);
    expect(land.pages[0].cols.length).toBeGreaterThan(6);
  });
});

describe('PDF writer', () => {
  it('writes a valid PDF with one image page each (checked with pdfinfo)', () => {
    // a 1x1 white JPEG
    const jpeg = Uint8Array.from(Buffer.from('/9j/4AAQSkZJRgABAQEASABIAAD/2wBDAP//////////////////////////////////////////////////////////////////////////////////////wgALCAABAAEBAREA/8QAFBABAAAAAAAAAAAAAAAAAAAAAP/aAAgBAQABPxA=', 'base64'));
    const pdf = pdfFromJpegs([0, 1, 2].map(() => ({ jpeg, w: 1, h: 1, pageW: 595.28, pageH: 841.89 })), 'Roster (draft)');
    const f = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'pdf-')), 'a.pdf');
    fs.writeFileSync(f, pdf);
    let info = '';
    try {
      info = execFileSync('pdfinfo', [f], { encoding: 'utf8' });
    } catch (e) {
      if ((e as NodeJS.ErrnoException).code === 'ENOENT') return; // no poppler here
      throw e;
    }
    expect(info).toMatch(/Pages:\s+3/);
    expect(info).toMatch(/Page size:\s+595\.28 x 841\.89 pts \(A4\)/);
    expect(info).toMatch(/Title:\s+Roster \(draft\)/);
  });
});
