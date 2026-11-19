/**
 * Formatting of OpenDocument spreadsheets (.ods) — L-03.
 *
 * SheetJS reads the values, formulas and merges of an .ods file but none of its looks. This reads
 * the rest from the package: cell styles (fonts, fills, borders, alignment, wrap, rotation and
 * number formats, with their parent styles), column widths, row heights, hidden rows / columns,
 * whole-row / whole-column formats and frozen panes. The result has the same shape as the
 * .xls reader's, so the loader applies both the same way. Pure and unit-tested; anything
 * unexpected returns null and the file opens unformatted as before.
 */
import { unzipSync, strFromU8 } from 'fflate';
import type { IStyleData } from '@univerjs/core';
import { attrs, children, decodeXml, type El } from './ooxml/xml';
import type { BiffInfo, BiffSheet } from './biff';

/** Columns and rows Petak keeps one entry for; further ones only count as "the rest of the sheet". */
const MAX_COLS = 1024;
const MAX_ROWS = 1_048_576;
/** The last row element, repeated this often and reaching this far, is the rest of the sheet: its
 *  look is the sheet's default, not something to copy into a million rows. */
const REST_ROWS = 65_536;
const REST_REPEAT = 1000;
const isRest = (r: number, n: number) => n > REST_REPEAT && r + n >= REST_ROWS;
/** At most this many rows are expanded from one repeated row element. */
const MAX_REPEAT = 10_000;

type Props = Record<string, string>;
interface StyleDef {
  parent?: string;
  /** own properties, normalised (see ownProps); `data` is the number style */
  props: Props;
}

/** "1.5cm", "0.4626in", "12pt" → points (NaN for anything else). */
export function lengthPt(s: string | undefined): number {
  const m = /^\s*(-?[\d.]+)\s*(cm|mm|in|inch|pt|pc|px)?\s*$/.exec(s ?? '');
  if (!m) return NaN;
  const v = +m[1];
  switch (m[2]) {
    case 'cm':
      return (v * 72) / 2.54;
    case 'mm':
      return (v * 72) / 25.4;
    case 'in':
    case 'inch':
      return v * 72;
    case 'pc':
      return v * 12;
    case 'px':
      return v * 0.75;
    default:
      return v;
  }
}
const unquote = (s: string) => s.replace(/^\s*['"]|['"]\s*$/g, '').trim();
const hex = (s: string | undefined) => (s && /^#[0-9a-f]{6}$/i.test(s.trim()) ? s.trim().toLowerCase() : undefined);

const LINE_STYLES = new Set(['none', 'hidden', 'solid', 'dotted', 'dashed', 'fine-dashed', 'dash-dot', 'dash-dot-dot', 'double', 'double-thin', 'groove', 'ridge', 'inset', 'outset', 'wave']);
/**
 * An ODF border ("0.74pt solid #000000") → Univer border, the way LibreOffice maps widths and
 * line styles to Excel's (thin 0.75 pt, medium 1.75 pt, thick 2.5 pt; dashed / dotted / dash-dot
 * variants by width).
 */
export function odsBorder(value: string | undefined, lineWidths?: string): { s: number; cl: { rgb: string } } | undefined {
  if (!value) return undefined;
  let width = NaN;
  let style = 'solid';
  let color = '#000000';
  for (const t of value.trim().split(/\s+/)) {
    if (LINE_STYLES.has(t)) style = t;
    else if (t.startsWith('#')) color = hex(t) ?? color;
    else if (t === 'thin') width = 0.75;
    else if (t === 'medium') width = 1.75;
    else if (t === 'thick') width = 2.5;
    else if (!Number.isNaN(lengthPt(t))) width = lengthPt(t);
  }
  if (style === 'none' || style === 'hidden') return undefined;
  if (Number.isNaN(width)) width = 0.75;
  const twips = Math.round(width * 20);
  if (twips < 1) return undefined;
  const double = style === 'double' || style === 'double-thin' || (lineWidths ?? '').trim().split(/\s+/).length === 3;
  let s: number;
  if (double) s = 7;
  else if (twips >= 45) s = 13; // thick
  else if (twips >= 30) s = { dashed: 9, 'dash-dot': 10, 'dash-dot-dot': 11, 'fine-dashed': 12 }[style] ?? 8; // medium…
  else if (twips >= 15) s = { dashed: 4, 'fine-dashed': 4, 'dash-dot': 5, 'dash-dot-dot': 6, dotted: 3 }[style] ?? 1; // thin…
  else s = 2; // hair
  return { s, cl: { rgb: color } };
}

// ------------------------------------------------------------------ number formats
/** Windows locale ids Excel writes in [$€-407]: language-COUNTRY → hex. */
const LCID: Record<string, string> = {
  'en-US': '409', 'en-GB': '809', 'en-AU': 'C09', 'en-CA': '1009', 'en-IN': '4009', 'en-SG': '4809', 'en-NZ': '1409', 'en-IE': '1809', 'en-ZA': '1C09',
  'de-DE': '407', 'de-AT': 'C07', 'de-CH': '807', 'fr-FR': '40C', 'fr-BE': '80C', 'fr-CH': '100C', 'fr-CA': 'C0C', 'it-IT': '410', 'it-CH': '810',
  'es-ES': 'C0A', 'es-MX': '80A', 'pt-PT': '816', 'pt-BR': '416', 'nl-NL': '413', 'nl-BE': '813', 'sv-SE': '41D', 'da-DK': '406', 'nb-NO': '414',
  'fi-FI': '40B', 'pl-PL': '415', 'cs-CZ': '405', 'sk-SK': '41B', 'hu-HU': '40E', 'ro-RO': '418', 'el-GR': '408', 'tr-TR': '41F', 'ru-RU': '419',
  'uk-UA': '422', 'he-IL': '40D', 'ar-SA': '401', 'hi-IN': '439', 'th-TH': '41E', 'vi-VN': '42A', 'id-ID': '421', 'ms-MY': '43E', 'fil-PH': '464',
  'zh-CN': '804', 'zh-TW': '404', 'zh-HK': 'C04', 'zh-SG': '1004', 'ja-JP': '411', 'ko-KR': '412',
};
const lcid = (a: Props): string | undefined => (a['number:language'] && a['number:country'] ? LCID[`${a['number:language']}-${a['number:country']}`] : undefined);
const COLOR_NAMES: Record<string, string> = {
  '#000000': 'Black', '#ffffff': 'White', '#ff0000': 'Red', '#00ff00': 'Green', '#0000ff': 'Blue',
  '#ffff00': 'Yellow', '#ff00ff': 'Magenta', '#00ffff': 'Cyan',
};
/** Literal text in an Excel number format: one character escaped, more quoted. */
function literal(t: string, dateish: boolean): string {
  if (!t) return '';
  if (t.length === 1) return dateish && (t === ':' || t === '/') ? t : '\\' + t;
  return '"' + t.replace(/"/g, '""') + '"';
}
function digits(a: Props): string {
  const minInt = a['number:min-integer-digits'] === undefined ? 1 : +a['number:min-integer-digits'];
  const dec = +(a['number:decimal-places'] ?? 0);
  const minDec = a['number:min-decimal-places'] === undefined ? dec : Math.min(dec, +a['number:min-decimal-places']);
  let int = '0'.repeat(Math.max(0, minInt)) || '#';
  if (a['number:grouping'] === 'true') int = int.length >= 4 ? int.replace(/\B(?=(\d{3})+(?!\d))/g, ',') : '#,' + '#'.repeat(3 - int.length) + int;
  const frac = dec > 0 ? '.' + '0'.repeat(minDec) + '#'.repeat(dec - minDec) : '';
  // display-factor 1000 → one trailing comma per thousand
  const factor = +(a['number:display-factor'] ?? 1);
  const scale = factor > 1 ? ','.repeat(Math.round(Math.log10(factor) / 3)) : '';
  return int + frac + scale;
}
/** One ODF number style element (without its conditions) → one Excel format section. */
function section(el: El): string {
  const dateish = el.name === 'date-style' || el.name === 'time-style';
  const elapsed = el.name === 'time-style' && attrs(el.open)['number:truncate-on-overflow'] === 'false';
  let out = '';
  let color = '';
  let firstTime = true;
  const long = (a: Props) => a['number:style'] === 'long';
  for (const c of children(el.inner)) {
    const a = attrs(c.open);
    switch (c.name) {
      case 'number':
        out += digits(a);
        break;
      case 'scientific-number': {
        const exp = '0'.repeat(Math.max(1, +(a['number:min-exponent-digits'] ?? 2)));
        out += digits({ ...a, 'number:grouping': 'false' }) + (a['number:forced-exponent-sign'] === 'false' ? 'E-' : 'E+') + exp;
        break;
      }
      case 'fraction': {
        const int = a['number:min-integer-digits'];
        const num = '?'.repeat(Math.max(1, +(a['number:min-numerator-digits'] ?? 1)));
        const den = a['number:denominator-value'] ?? '?'.repeat(Math.max(1, +(a['number:min-denominator-digits'] ?? 1)));
        out += (int === undefined ? '' : (+int > 0 ? '0'.repeat(+int) : '#') + ' ') + num + '/' + den;
        break;
      }
      case 'text': {
        const t = c.inner.replace(/<[^>]*>/g, '').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
        out += el.name === 'percentage-style' ? t.split('%').map((p) => literal(p, false)).join('%') : literal(t, dateish);
        break;
      }
      case 'text-content':
        out += '@';
        break;
      case 'currency-symbol': {
        const sym = c.inner.replace(/<[^>]*>/g, '').trim();
        const id = lcid(a);
        out += sym ? `[$${sym}${id ? '-' + id : ''}]` : '\\$';
        break;
      }
      case 'day':
        out += long(a) ? 'dd' : 'd';
        break;
      case 'month':
        out += a['number:textual'] === 'true' ? (long(a) ? 'mmmm' : 'mmm') : long(a) ? 'mm' : 'm';
        break;
      case 'year':
        out += long(a) ? 'yyyy' : 'yy';
        break;
      case 'day-of-week':
        out += long(a) ? 'dddd' : 'ddd';
        break;
      case 'hours': {
        const h = long(a) ? 'hh' : 'h';
        out += elapsed && firstTime ? `[${h}]` : h;
        firstTime = false;
        break;
      }
      case 'minutes': {
        const m = long(a) ? 'mm' : 'm';
        out += elapsed && firstTime ? `[${m}]` : m;
        firstTime = false;
        break;
      }
      case 'seconds': {
        const dp = +(a['number:decimal-places'] ?? 0);
        const s = (long(a) ? 'ss' : 's') + (dp > 0 ? '.' + '0'.repeat(dp) : '');
        out += elapsed && firstTime ? `[${s}]` : s;
        firstTime = false;
        break;
      }
      case 'am-pm':
        out += 'AM/PM';
        break;
      case 'quarter':
        out += long(a) ? '"Quarter "' : '"Q"';
        break;
      case 'text-properties': {
        const name = COLOR_NAMES[hex(a['fo:color']) ?? ''];
        if (name) color = `[${name}]`;
        break;
      }
    }
  }
  const id = lcid(attrs(el.open));
  return color + (id && id !== '409' && !out.includes('[$') ? `[$-${id}]` : '') + out;
}
/** An ODF data style (with its style:map conditions) → Excel number format code, or undefined for General. */
export function odsNumberFormat(name: string, numStyles: Map<string, El>): string | undefined {
  const el = numStyles.get(name);
  if (!el) return undefined;
  const main = section(el);
  const maps = children(el.inner)
    .filter((c) => c.name === 'map')
    .map((m) => ({ cond: attrs(m.open)['style:condition'] ?? '', apply: attrs(m.open)['style:apply-style-name'] ?? '' }));
  if (!maps.length) return main || undefined;
  const sections: string[] = [];
  maps.slice(0, 3).forEach((m, i) => {
    const sub = numStyles.get(m.apply);
    const body = sub ? section(sub) : '';
    const c = /value\(\)\s*(>=|<=|!=|>|<|=)\s*(-?[\d.]+)/.exec(m.cond);
    // Excel's plain sections: positive (> 0 or >= 0), negative (< 0), zero
    const plain = c && +c[2] === 0 && ((i === 0 && (c[1] === '>' || c[1] === '>=')) || (i === 1 && c[1] === '<') || (i === 2 && c[1] === '='));
    sections.push(plain || !c ? body : `[${c[1] === '!=' ? '<>' : c[1]}${c[2]}]${body}`);
  });
  sections.push(main);
  return sections.join(';');
}

// ------------------------------------------------------------------ values
const EPOCH = Date.UTC(1899, 11, 30);
/** to 1/100 000 of a second: files store times to the microsecond at most */
const round = (days: number) => Math.round(days * 864e7) / 864e7;
/** office:date-value ("2026-11-19T15:04:45.12") → serial day number, as LibreOffice and Excel count. */
export function odsDate(v: string | undefined): number | undefined {
  const m = /^(-?\d{4,})-(\d\d)-(\d\d)(?:T(\d\d):(\d\d)(?::(\d\d)(\.\d+)?)?)?/.exec(v ?? '');
  if (!m) return undefined;
  const d = new Date(0);
  d.setUTCFullYear(+m[1], +m[2] - 1, +m[3]);
  d.setUTCHours(+(m[4] ?? 0), +(m[5] ?? 0), +(m[6] ?? 0), 0);
  // whole seconds first, then the fraction: 15:04:45.12 is 0.6283 exactly as the file meant
  return round((Math.round((d.getTime() - EPOCH) / 1000) + (m[7] ? +m[7] : 0)) / 86400);
}
/** office:time-value ("PT42H00M00S", "-P1DT2H") → days. */
export function odsDuration(v: string | undefined): number | undefined {
  const m = /^(-)?P(?:(\d+)D)?(?:T(?:(\d+)H)?(?:(\d+)M)?(?:([\d.]+)S)?)?$/.exec((v ?? '').trim());
  if (!m || !/\d/.test(v ?? '')) return undefined;
  const days = round(+(m[2] ?? 0) + (+(m[3] ?? 0) * 3600 + +(m[4] ?? 0) * 60 + +(m[5] ?? 0)) / 86400);
  return m[1] ? -days : days;
}

// ------------------------------------------------------------------ styles
interface Ctx {
  defs: Map<string, StyleDef>;
  fonts: Map<string, string>;
  numStyles: Map<string, El>;
  resolved: Map<string, Props>;
  colWidth: Map<string, number>;
  rowStyle: Map<string, { pt: number; optimal: boolean }>;
}
/** Normalised own properties of a style element: later keys win when merged over the parent's. */
function ownProps(el: El, fonts: Map<string, string>): Props {
  const p: Props = {};
  for (const c of children(el.inner)) {
    if (c.name !== 'table-cell-properties' && c.name !== 'paragraph-properties' && c.name !== 'text-properties') continue;
    const a = attrs(c.open);
    if (c.name === 'text-properties') {
      const family = a['fo:font-family'] ?? (a['style:font-name'] ? fonts.get(a['style:font-name']) ?? a['style:font-name'] : undefined);
      if (family) p.font = unquote(family.split(',')[0]);
      if (a['fo:font-size'] && !a['fo:font-size'].endsWith('%')) p.size = String(lengthPt(a['fo:font-size']));
      if (a['fo:font-weight']) p.bold = a['fo:font-weight'] === 'bold' || +a['fo:font-weight'] >= 600 ? '1' : '';
      if (a['fo:font-style']) p.italic = a['fo:font-style'] === 'italic' || a['fo:font-style'] === 'oblique' ? '1' : '';
      if (a['style:text-underline-style']) p.underline = a['style:text-underline-style'] !== 'none' ? '1' : '';
      if (a['style:text-line-through-style']) p.strike = a['style:text-line-through-style'] !== 'none' ? '1' : '';
      if (a['style:use-window-font-color'] === 'true') p.color = '';
      else if (a['fo:color']) p.color = hex(a['fo:color']) ?? '';
    } else if (c.name === 'paragraph-properties') {
      if (a['fo:text-align']) p.align = a['fo:text-align'];
    } else {
      if (a['style:text-align-source'] === 'value-type') p.align = '';
      if (a['fo:background-color']) p.bg = hex(a['fo:background-color']) ?? '';
      if (a['fo:border'] !== undefined) {
        for (const s of ['t', 'b', 'l', 'r']) {
          p['bd' + s] = a['fo:border'];
          p['bw' + s] = a['style:border-line-width'] ?? '';
        }
      }
      for (const [s, side] of [['t', 'top'], ['b', 'bottom'], ['l', 'left'], ['r', 'right']]) {
        if (a['fo:border-' + side] !== undefined) {
          p['bd' + s] = a['fo:border-' + side];
          p['bw' + s] = a['style:border-line-width-' + side] ?? '';
        }
      }
      if (a['style:diagonal-tl-br'] !== undefined) p.tlbr = a['style:diagonal-tl-br'];
      if (a['style:diagonal-bl-tr'] !== undefined) p.bltr = a['style:diagonal-bl-tr'];
      if (a['style:vertical-align']) p.valign = a['style:vertical-align'];
      if (a['fo:wrap-option']) p.wrap = a['fo:wrap-option'] === 'wrap' ? '1' : '';
      if (a['style:rotation-angle']) p.rot = String(parseFloat(a['style:rotation-angle']) || 0);
      if (a['style:direction']) p.ttb = a['style:direction'] === 'ttb' ? '1' : '';
    }
  }
  return p;
}
const withData = (p: Props, a: Props): Props => (a['style:data-style-name'] ? { ...p, data: a['style:data-style-name'] } : p);
function resolve(ctx: Ctx, name: string | undefined, depth = 0): Props {
  const key = name ?? '';
  const hit = ctx.resolved.get(key);
  if (hit) return hit;
  const def = ctx.defs.get(key);
  // the default style of the family is the root of every chain
  const base = key === '' || !def || depth > 20 ? (key === '' ? {} : resolve(ctx, '', depth + 1)) : resolve(ctx, def.parent ?? '', depth + 1);
  const p: Props = { ...base, ...(def?.props ?? {}) };
  ctx.resolved.set(key, p);
  return p;
}
const H: Record<string, number> = { start: 1, left: 1, center: 2, end: 3, right: 3, justify: 4 };
const V: Record<string, number> = { top: 1, middle: 2 };
/** Resolved properties → Univer style, leaving out what equals the workbook default. */
function toStyle(p: Props, def: { name?: string; size?: number }, numStyles: Map<string, El>): IStyleData | null {
  const s: IStyleData = {};
  if (p.font && p.font !== def.name) s.ff = p.font;
  const size = p.size ? Math.round(+p.size * 10) / 10 : NaN;
  if (size > 0 && size !== def.size) s.fs = size;
  if (p.bold) s.bl = 1;
  if (p.italic) s.it = 1;
  if (p.underline) s.ul = { s: 1 };
  if (p.strike) s.st = { s: 1 };
  if (p.color && p.color !== '#000000') s.cl = { rgb: p.color };
  if (p.bg) s.bg = { rgb: p.bg };
  const bd: Record<string, unknown> = {};
  for (const k of ['t', 'b', 'l', 'r']) {
    const b = odsBorder(p['bd' + k], p['bw' + k]);
    if (b) bd[k] = b;
  }
  const tlbr = odsBorder(p.tlbr);
  if (tlbr) bd.tl_br = tlbr;
  const bltr = odsBorder(p.bltr);
  if (bltr) bd.bl_tr = bltr;
  if (Object.keys(bd).length) s.bd = bd as IStyleData['bd'];
  if (p.align && H[p.align]) s.ht = H[p.align] as never;
  if (p.valign && V[p.valign]) s.vt = V[p.valign] as never;
  if (p.wrap) s.tb = 3 as never;
  if (p.ttb) s.tr = { a: 0, v: 1 as never };
  else if (p.rot && +p.rot % 360) {
    const r = ((+p.rot % 360) + 360) % 360;
    // ODF turns counter-clockwise; Univer's angle is Excel's (clockwise positive), within ±90°
    const a = r <= 180 ? -r : 360 - r;
    s.tr = { a: Math.max(-90, Math.min(90, a)) };
  }
  if (p.data) {
    const pattern = odsNumberFormat(p.data, numStyles);
    if (pattern) s.n = { pattern };
  }
  return Object.keys(s).length ? s : null;
}

// ------------------------------------------------------------------ package
/** The elements of one top-level section (office:styles, office:automatic-styles…), found by
 *  name so the sheets of a big content.xml are not scanned for it. */
function stylesOf(xml: string, section: string): El[] {
  const open = new RegExp(`<([\\w.-]+):${section}(?=[\\s/>])[^>]*>`).exec(xml);
  if (!open || open[0].endsWith('/>')) return [];
  const from = open.index + open[0].length;
  const to = xml.indexOf(`</${open[1]}:${section}>`, from);
  return to < 0 ? [] : children(xml.slice(from, to));
}
/** Frozen panes per sheet from settings.xml (split mode 2 = frozen). */
export function odsFreeze(settings: string): Map<string, { cols: number; rows: number }> {
  const out = new Map<string, { cols: number; rows: number }>();
  const at = settings.search(/config:name="Tables"/);
  if (at < 0) return out;
  const re = /<config:config-item-map-entry\s+config:name="([^"]*)"\s*>([\s\S]*?)<\/config:config-item-map-entry>/g;
  re.lastIndex = at;
  let m: RegExpExecArray | null;
  const item = (body: string, name: string) => {
    const x = new RegExp(`config:name="${name}"[^>]*>([^<]*)<`).exec(body);
    return x ? +x[1] : 0;
  };
  while ((m = re.exec(settings))) {
    const body = m[2];
    const cols = item(body, 'HorizontalSplitMode') === 2 ? item(body, 'HorizontalSplitPosition') : 0;
    const rows = item(body, 'VerticalSplitMode') === 2 ? item(body, 'VerticalSplitPosition') : 0;
    if (cols || rows) out.set(m[1].replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'"), { cols, rows });
  }
  return out;
}

export function readOdsFormatting(buf: ArrayBuffer): BiffInfo | null {
  let files: Record<string, Uint8Array>;
  try {
    files = unzipSync(new Uint8Array(buf), { filter: (f) => f.name === 'content.xml' || f.name === 'styles.xml' || f.name === 'settings.xml' });
  } catch {
    return null;
  }
  if (!files['content.xml']) return null;
  const content = strFromU8(files['content.xml']);
  const stylesXml = files['styles.xml'] ? strFromU8(files['styles.xml']) : '';
  const settings = files['settings.xml'] ? strFromU8(files['settings.xml']) : '';

  // fonts, number styles, cell / column / row styles
  const fonts = new Map<string, string>();
  for (const xml of [stylesXml, content])
    for (const f of stylesOf(xml, 'font-face-decls')) {
      const a = attrs(f.open);
      if (a['style:name']) fonts.set(a['style:name'], unquote(a['svg:font-family'] ?? a['style:name']));
    }
  const ctx: Ctx = { defs: new Map(), fonts, numStyles: new Map(), resolved: new Map(), colWidth: new Map(), rowStyle: new Map() };
  const collect = (els: El[]) => {
    for (const el of els) {
      const a = attrs(el.open);
      if (el.name.endsWith('-style') && el.open.startsWith('<number:')) {
        if (a['style:name']) ctx.numStyles.set(a['style:name'], el);
      } else if (el.name === 'default-style' && a['style:family'] === 'table-cell') {
        ctx.defs.set('', { props: withData(ownProps(el, fonts), a) });
      } else if (el.name === 'style' && a['style:name']) {
        const fam = a['style:family'];
        if (fam === 'table-cell') ctx.defs.set(a['style:name'], { parent: a['style:parent-style-name'], props: withData(ownProps(el, fonts), a) });
        else if (fam === 'table-column') {
          const w = lengthPt(attrs(children(el.inner).find((c) => c.name === 'table-column-properties')?.open ?? '')['style:column-width']);
          if (w > 0) ctx.colWidth.set(a['style:name'], (w * 4) / 3);
        } else if (fam === 'table-row') {
          const ra = attrs(children(el.inner).find((c) => c.name === 'table-row-properties')?.open ?? '');
          const h = lengthPt(ra['style:row-height'] ?? ra['style:min-row-height']);
          if (h > 0) ctx.rowStyle.set(a['style:name'], { pt: h, optimal: ra['style:use-optimal-row-height'] === 'true' });
        }
      }
    }
  };
  collect(stylesOf(stylesXml, 'styles'));
  collect(stylesOf(stylesXml, 'automatic-styles'));
  collect(stylesOf(content, 'automatic-styles'));
  // the default-style holds the document's font; "Default" usually overrides it
  const base = resolve(ctx, ctx.defs.has('Default') ? 'Default' : '');
  const defaultFont = { name: base.font || undefined, size: base.size ? Math.round(+base.size * 10) / 10 : undefined };

  const styles: Array<IStyleData | null> = [];
  const index = new Map<string, number>();
  /** style index for a cell style name, or -1 when it adds nothing to the defaults */
  const styleIndex = (name: string | undefined): number => {
    const key = name ?? 'Default';
    let i = index.get(key);
    if (i === undefined) {
      const st = toStyle(resolve(ctx, ctx.defs.has(key) ? key : 'Default'), defaultFont, ctx.numStyles);
      i = st ? styles.push(st) - 1 : -1;
      index.set(key, i);
    }
    return i;
  };

  const freeze = odsFreeze(settings);
  // One pass over the sheets with a single regular expression: a 50 000-row file has a million
  // cells, and building an element per cell would cost seconds.
  const T = prefixOf(content, 'urn:oasis:names:tc:opendocument:xmlns:table:1.0', 'table');
  const O = prefixOf(content, 'urn:oasis:names:tc:opendocument:xmlns:office:1.0', 'office');
  const A = {
    name: T + ':name',
    style: T + ':style-name',
    cellDefault: T + ':default-cell-style-name',
    rowsRepeated: T + ':number-rows-repeated',
    colsRepeated: T + ':number-columns-repeated',
    visibility: T + ':visibility',
    type: O + ':value-type',
    date: O + ':date-value',
    time: O + ':time-value',
  };
  const startAt = content.search(new RegExp(`<${O}:spreadsheet[\\s>]`));
  if (startAt < 0) return null;
  const tagRe = new RegExp(`<(/?)${T}:(table|table-row|table-cell|covered-table-cell|table-column)(?=[\\s/>])([^>]*)>`, 'g');
  tagRe.lastIndex = startAt;

  const sheets: BiffSheet[] = [];
  let m: RegExpExecArray | null;
  while ((m = tagRe.exec(content))) {
    if (m[2] !== 'table' || m[1]) continue;
    // ---- one sheet: from its start tag to its end tag
    const name = attrOf(m[3], A.name) ?? 'Sheet' + (sheets.length + 1);
    const sh: BiffSheet & { colPx: Map<number, number> } = {
      name,
      cells: new Map(),
      colWidths: new Map(),
      colPx: new Map(),
      defaultColPx: 0,
      hiddenCols: new Set(),
      rowHeights: new Map(),
      hiddenRows: new Set(),
      maxRow: -1,
      maxCol: -1,
    };
    const colDefault: Array<string | undefined> = [];
    const colSpecs: Array<{ c: number; n: number; px: number; hidden: boolean }> = [];
    const rowSpecs: Array<{ r: number; n: number; pt: number; optimal: boolean; hidden: boolean; rest: boolean }> = [];
    const colStyles = new Map<number, number>();
    const rowStyles = new Map<number, number>();
    const optimalRows = new Set<number>();
    const dates = new Map<number, Map<number, number>>();
    const tableStart = tagRe.lastIndex;
    // the last row element of the sheet may be "the rest of the sheet"
    const tableEnd = content.indexOf(`</${T}:table>`, tableStart);
    let lastRowAt = content.lastIndexOf(`<${T}:table-row`, tableEnd < 0 ? content.length : tableEnd);
    // not <table:table-rows>, the container
    while (lastRowAt > 0 && !/[\s/>]/.test(content[lastRowAt + T.length + 11])) lastRowAt = content.lastIndexOf(`<${T}:table-row`, lastRowAt - 1);
    let col = 0;
    let row = 0;
    let depth = 1;
    // the row being read
    let inRow = false;
    let n = 1;
    let rest = false;
    let expand = 1;
    let rowDefault: string | undefined;
    let c = 0;
    const put = (rr: number, cc: number, i: number) => {
      let cm = sh.cells.get(rr);
      if (!cm) sh.cells.set(rr, (cm = new Map()));
      cm.set(cc, i);
      if (rr > sh.maxRow) sh.maxRow = rr;
      if (cc > sh.maxCol) sh.maxCol = cc;
    };
    const endRow = () => {
      if (!inRow) return;
      inRow = false;
      row = Math.min(MAX_ROWS, row + n);
    };
    while ((m = tagRe.exec(content))) {
      const [, close, kind, at] = m;
      if (kind === 'table') {
        if (!close && !at.endsWith('/')) depth++;
        else if (close && --depth === 0) break;
        continue;
      }
      if (depth > 1) continue; // a table inside a cell
      if (kind === 'table-row') {
        if (close) {
          endRow();
          continue;
        }
        endRow();
        n = Math.max(1, +(attrOf(at, A.rowsRepeated) ?? 1) || 1);
        const rs = ctx.rowStyle.get(attrOf(at, A.style) ?? '');
        const vis = attrOf(at, A.visibility);
        rest = m.index === lastRowAt && isRest(row, n);
        rowSpecs.push({ r: row, n, pt: rs?.pt ?? NaN, optimal: rs?.optimal ?? true, hidden: vis === 'collapse' || vis === 'filter', rest });
        expand = rest ? 0 : Math.min(n, MAX_REPEAT);
        rowDefault = attrOf(at, A.cellDefault);
        c = 0;
        inRow = true;
        if (at.endsWith('/')) endRow();
        continue;
      }
      if (close) continue;
      if (kind === 'table-column') {
        const cn = Math.max(1, +(attrOf(at, A.colsRepeated) ?? 1) || 1);
        const vis = attrOf(at, A.visibility);
        colSpecs.push({ c: col, n: cn, px: ctx.colWidth.get(attrOf(at, A.style) ?? '') ?? NaN, hidden: vis === 'collapse' || vis === 'filter' });
        const def = attrOf(at, A.cellDefault);
        for (let k = 0; k < cn && col + k < MAX_COLS; k++) colDefault[col + k] = def;
        col += cn;
        continue;
      }
      if (!inRow) continue;
      // a cell or a cell covered by a merge
      const cn = Math.max(1, +(attrOf(at, A.colsRepeated) ?? 1) || 1);
      const own = attrOf(at, A.style);
      if (c + cn >= MAX_COLS && cn > 1) {
        // a cell repeated to the end of the row: a whole-row format
        const i = styleIndex(own ?? rowDefault ?? colDefault[c]);
        if (i >= 0) for (let k = 0; k < expand; k++) rowStyles.set(row + k, i);
        c += cn;
        continue;
      }
      const type = attrOf(at, A.type);
      if (type === 'date' || type === 'time') {
        const exact = type === 'date' ? odsDate(attrOf(at, A.date)) : odsDuration(attrOf(at, A.time));
        if (exact !== undefined)
          for (let j = 0; j < expand; j++) {
            let dm = dates.get(row + j);
            if (!dm) dates.set(row + j, (dm = new Map()));
            for (let k = 0; k < cn && c + k < MAX_COLS; k++) dm.set(c + k, exact);
          }
      }
      for (let k = 0; k < cn && c + k < MAX_COLS; k++) {
        const i = styleIndex(own ?? rowDefault ?? colDefault[c + k]);
        if (i < 0) continue;
        // the same look down the rest of the sheet: a whole-column format
        if (rest) colStyles.set(c + k, i);
        else for (let j = 0; j < expand; j++) put(row + j, c + k, i);
      }
      c += cn;
    }
    endRow();

    // column widths: the last column element that runs to the end of the sheet is the default
    const lastCol = colSpecs.at(-1);
    sh.defaultColPx = lastCol && lastCol.c + lastCol.n >= MAX_COLS && lastCol.px > 0 ? Math.round(lastCol.px) : 85;
    for (const s of colSpecs)
      for (let k = 0; k < s.n && s.c + k < MAX_COLS; k++) {
        if (s.px > 0 && Math.round(s.px) !== sh.defaultColPx) sh.colPx.set(s.c + k, Math.round(s.px));
        if (s.hidden && !(s === lastCol && s.c + s.n >= MAX_COLS)) sh.hiddenCols.add(s.c + k);
      }
    // row heights: likewise the last row element when it is the rest of the sheet; else LibreOffice's 0.178 in
    const lastRow = rowSpecs.at(-1);
    const fs = defaultFont.size ?? 10;
    sh.defaultRowPt = lastRow?.rest && lastRow.pt > 0 ? lastRow.pt : fs <= 11 ? 12.8 : Math.round(fs * 1.17 * 10) / 10;
    for (const s of rowSpecs) {
      if (s.rest) continue;
      for (let k = 0; k < Math.min(s.n, MAX_REPEAT); k++) {
        const rr = s.r + k;
        if (s.pt > 0 && Math.abs(s.pt - sh.defaultRowPt) > 0.05) {
          sh.rowHeights.set(rr, Math.round(s.pt * 20) / 20); // to the twip: 0.8335 in is 60 pt
          if (s.optimal) optimalRows.add(rr);
        }
        if (s.hidden) sh.hiddenRows.add(rr);
      }
    }
    const fr = freeze.get(name);
    if (fr) sh.freeze = fr;
    if (colStyles.size) sh.colStyles = colStyles;
    if (rowStyles.size) sh.rowStyles = rowStyles;
    if (optimalRows.size) sh.optimalRows = optimalRows;
    if (dates.size) sh.dates = dates;
    sheets.push(sh);
  }
  return sheets.length ? { styles, defaultFont, sheets } : null;
}

/** The prefix a document binds to a namespace (LibreOffice always uses the usual ones). */
function prefixOf(xml: string, ns: string, fallback: string): string {
  const head = xml.slice(0, 4000);
  const m = new RegExp(`xmlns:([\\w.-]+)="${ns.replace(/[.:]/g, '\\$&')}"`).exec(head);
  return m ? m[1] : fallback;
}
/** One attribute of a start tag's attribute text, decoded; fast path for the million cells of a big sheet. */
function attrOf(tag: string, name: string): string | undefined {
  let i = tag.indexOf(name + '=');
  while (i > 0 && !/\s/.test(tag[i - 1])) i = tag.indexOf(name + '=', i + 1);
  if (i < 0) return undefined;
  const q = tag[i + name.length + 1];
  const end = tag.indexOf(q, i + name.length + 2);
  const v = tag.slice(i + name.length + 2, end);
  return v.indexOf('&') < 0 ? v : decodeXml(v);
}
