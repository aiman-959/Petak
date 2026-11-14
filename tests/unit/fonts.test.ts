import { describe, expect, it } from 'vitest';
import { fontFaceCss, fontReport, nameFromFile, readFontName, reportText, substituteFor } from '../../src/fonts/fonts';

describe('font substitutes', () => {
  it('maps the Excel fonts to metric-compatible ones', () => {
    expect(substituteFor('Calibri')?.family).toBe('Carlito');
    expect(substituteFor('cambria')?.family).toBe('Caladea');
    expect(substituteFor('Arial')?.family).toBe('Arimo');
    expect(substituteFor(' Times New Roman ')?.family).toBe('Tinos');
    expect(substituteFor('Courier New')?.family).toBe('Cousine');
    expect(substituteFor('Liberation Sans')?.family).toBe('Arimo');
    expect(substituteFor('Comic Sans MS')).toBeNull();
  });
  it('registers each substitute under the original name, real fonts first', () => {
    const css = fontFaceCss((pkg, subset, w, st) => `/f/${pkg}-${subset}-${w}-${st}.woff2`);
    expect(css).toContain('@font-face{font-family:"Calibri";font-style:normal;font-weight:400;font-display:swap;src:local("Calibri"), url(/f/carlito-latin-400-normal.woff2)');
    expect(css).toContain('font-family:"Times New Roman";font-style:italic;font-weight:700');
    expect(css).toContain('url(/f/tinos-cyrillic-700-italic.woff2)');
    expect(css).toMatch(/font-family:"Carlito";[^}]*src:url\(/); // the substitute's own name has no local()
    expect(css).not.toContain('caladea-cyrillic'); // Caladea has no Cyrillic
    // a missing file is skipped, not written as "undefined"
    expect(fontFaceCss(() => undefined)).toBe('');
  });
  it('reports what a workbook\'s fonts are shown with', () => {
    const r = fontReport(['Calibri', 'Calibri', 'Arial', 'Carlito', '']);
    expect(r.substituted).toEqual([['Calibri', 'Carlito'], ['Arial', 'Arimo']]);
    expect(reportText(r)).toBe('Fonts: Calibri → Carlito, Arial → Arimo (same widths)');
    expect(reportText({ substituted: [['A', 'a'], ['B', 'b'], ['C', 'c']], missing: [] })).toBe('Fonts: 3 shown with built-in fonts of the same widths');
    expect(reportText({ substituted: [], missing: [] })).toBeNull();
    expect(reportText({ substituted: [], missing: ['Papyrus'] })).toContain('not on this device: Papyrus');
  });
});

describe('own font files', () => {
  it('guesses family and style from the file name', () => {
    expect(nameFromFile('Roboto-BoldItalic.ttf')).toEqual({ family: 'Roboto', bold: true, italic: true });
    expect(nameFromFile('Open_Sans-Regular.otf')).toEqual({ family: 'Open Sans', bold: false, italic: false });
    expect(nameFromFile('MyFont.woff2')).toEqual({ family: 'MyFont', bold: false, italic: false });
  });
  it('reads the family from a TrueType name table', () => {
    // minimal sfnt: one table ('name') with family "Test Sans" (id 1) and subfamily "Bold Italic" (id 2), Windows UTF-16BE
    const enc = (s: string) => [...s].flatMap((c) => [0, c.charCodeAt(0)]);
    const fam = enc('Test Sans');
    const sub = enc('Bold Italic');
    const nameTable = [0, 0, 0, 2, 0, 6 + 24, ...[0, 3, 0, 1, 4, 9, 0, 1, 0, fam.length, 0, 0], ...[0, 3, 0, 1, 4, 9, 0, 2, 0, sub.length, 0, fam.length], ...fam, ...sub];
    const head = [0, 1, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, ...[0x6e, 0x61, 0x6d, 0x65], 0, 0, 0, 0, 0, 0, 0, 28, 0, 0, 0, nameTable.length];
    expect(readFontName(new Uint8Array([...head, ...nameTable]).buffer)).toEqual({ family: 'Test Sans', bold: true, italic: true });
    expect(readFontName(new Uint8Array([1, 2, 3, 4, 5, 6, 7, 8]).buffer)).toBeNull();
  });
});
