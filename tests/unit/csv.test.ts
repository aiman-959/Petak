import { describe, expect, it } from 'vitest';
import { csvDateFormat, csvKeepText, parseWithSheetJs, streamSheetJs } from '../../src/io/csv';
import type { LoadChunk } from '../../src/io/types';

const load = (text: string) => {
  const p = parseWithSheetJs(new TextEncoder().encode(text).buffer as ArrayBuffer, 'data.csv');
  const gen = streamSheetJs(p, 100);
  const rows: Record<number, Record<number, { v?: unknown; t?: number; s?: string }>> = {};
  let r = gen.next();
  while (!r.done) {
    const c = r.value as LoadChunk;
    if (c.kind === 'rows') Object.assign(rows, c.cellData);
    r = gen.next();
  }
  return { rows, styles: r.value.styles as Record<string, { n?: { pattern: string } }> };
};

describe('CSV import (L-01)', () => {
  it('keeps identifiers with a leading zero or more than 15 digits as text', () => {
    expect(csvKeepText('0012')).toBe('0012');
    expect(csvKeepText('0123456789')).toBe('0123456789');
    expect(csvKeepText('12345678901234567890')).toBe('12345678901234567890');
    expect(csvKeepText('0.5')).toBeNull();
    expect(csvKeepText('0')).toBeNull();
    expect(csvKeepText('120')).toBeNull();
    const { rows } = load('id,phone,card,n\n0012,0123456789,4111111111111111111,1.5\n');
    expect(rows[1][0]).toMatchObject({ v: '0012', t: 1 });
    expect(rows[1][1]).toMatchObject({ v: '0123456789', t: 1 });
    expect(rows[1][2]).toMatchObject({ v: '4111111111111111111', t: 1 });
    expect(rows[1][3]).toMatchObject({ v: 1.5, t: 2 });
  });
  it('dates keep a date format instead of showing the serial number', () => {
    expect(csvDateFormat('2024-01-05')).toBe('yyyy-mm-dd');
    expect(csvDateFormat('2024-01-05 13:45')).toBe('yyyy-mm-dd hh:mm');
    expect(csvDateFormat('1.5')).toBeNull();
    const { rows, styles } = load('d\n2024-01-05\n');
    expect(rows[1][0].v).toBe(45296);
    expect(styles[rows[1][0].s!].n?.pattern).toBe('yyyy-mm-dd');
  });
});
