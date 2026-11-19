import { describe, expect, it } from 'vitest';
import { REF_COLORS, formulaRefs } from '../../src/ui/formulaRefs';

describe('coloured references in a formula (M-03)', () => {
  it('finds cells, ranges, whole columns / rows and sheet references', () => {
    const f = "=SUM(A1:B3)+$C$2*Data!D4-'My sheet'!E5+F:F+2:3";
    const r = formulaRefs(f);
    expect(r.map((x) => x.text)).toEqual(['A1:B3', '$C$2', 'Data!D4', "'My sheet'!E5", 'F:F', '2:3']);
    expect(r[0]).toMatchObject({ r0: 0, c0: 0, r1: 2, c1: 1 });
    expect(r[2].sheet).toBe('Data');
    expect(r[3].sheet).toBe('My sheet');
    expect(f.slice(r[1].start, r[1].end)).toBe('$C$2');
  });
  it('the same reference keeps its colour, different ones get different colours', () => {
    const r = formulaRefs('=A1+B2+a1+B2');
    expect(r.map((x) => x.color)).toEqual([REF_COLORS[0], REF_COLORS[1], REF_COLORS[0], REF_COLORS[1]]);
  });
  it('skips text in quotes, function names, names and plain numbers', () => {
    expect(formulaRefs('="A1"&B2').map((x) => x.text)).toEqual(['B2']);
    expect(formulaRefs('=LOG10(100)+ATAN2(1,2)').length).toBe(0);
    expect(formulaRefs('=Qty*2+Tax_rate').length).toBe(0);
    // two- and three-letter columns are references, not names
    expect(formulaRefs('=AB12+XFD1048576').map((x) => [x.text, x.c0])).toEqual([['AB12', 27], ['XFD1048576', 16383]]);
    expect(formulaRefs('=12+3.5').length).toBe(0);
    expect(formulaRefs('plain A1 text').length).toBe(0);
  });
});
