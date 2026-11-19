/**
 * Cell references in a formula being typed (M-03): where they are in the text and which
 * colour each one gets. The same reference always gets the same colour, as in Excel. Text in
 * quotes and function names are skipped. Pure: unit-tested.
 */
import { parseRange } from '../io/a1';

export const REF_COLORS = ['#1f6feb', '#d1242f', '#8250df', '#1a7f37', '#bf8700', '#e16f24', '#0a7ea4', '#bf3989'];

export interface FormulaRef {
  start: number;
  end: number;
  text: string;
  sheet?: string;
  r0: number;
  c0: number;
  r1: number;
  c1: number;
  color: string;
}

// optional sheet ('My sheet'! or Sheet1!), then A1, $A$1, A1:B2, A:A or 1:3
const REF = /(?:(?:'((?:[^']|'')+)'|([A-Za-z_\u00C0-\uFFFF][\w.\u00C0-\uFFFF]*))!)?(\$?[A-Za-z]{1,3}\$?\d{1,7}(?::\$?[A-Za-z]{1,3}\$?\d{1,7})?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?\d{1,7}:\$?\d{1,7})/y;

export function formulaRefs(text: string): FormulaRef[] {
  if (!text.startsWith('=')) return [];
  const out: FormulaRef[] = [];
  const colorOf = new Map<string, string>();
  let i = 1;
  while (i < text.length) {
    const ch = text[i];
    if (ch === '"') {
      // skip a string literal ("" is an escaped quote)
      i++;
      while (i < text.length && !(text[i] === '"' && text[i + 1] !== '"')) i += text[i] === '"' ? 2 : 1;
      i++;
      continue;
    }
    // a reference starts after an operator, bracket, comma, space or at the beginning
    const prev = text[i - 1];
    if (!/[=+\-*/^&,;(:<> %]/.test(prev) && i > 1) {
      i++;
      continue;
    }
    REF.lastIndex = i;
    const m = REF.exec(text);
    const next = m ? text[REF.lastIndex] : undefined;
    // not a reference when followed by "(" (a function such as LOG10) or more letters / digits (a name)
    if (!m || (next !== undefined && /[\w(.\u00C0-\uFFFF]/.test(next))) {
      i++;
      continue;
    }
    const sheet = m[1] !== undefined ? m[1].replace(/''/g, "'") : m[2];
    try {
      const r = parseRange(m[3]);
      const key = `${(sheet ?? '').toLowerCase()}!${r.r0},${r.c0},${r.r1},${r.c1}`;
      let color = colorOf.get(key);
      if (!color) colorOf.set(key, (color = REF_COLORS[colorOf.size % REF_COLORS.length]));
      out.push({ start: i, end: REF.lastIndex, text: m[0], sheet, r0: r.r0, c0: r.c0, r1: r.r1, c1: r.c1, color });
    } catch {
      /* not a reference after all */
    }
    i = REF.lastIndex;
  }
  return out;
}
