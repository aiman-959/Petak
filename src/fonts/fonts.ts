/**
 * Excel fonts on a device that does not have them (K-06).
 *
 * Android has no Calibri, Cambria, Arial, Times New Roman or Courier New. Petak ships fonts
 * with the same character widths and registers them under the Excel names:
 *
 *   Calibri → Carlito     Cambria → Caladea     Arial / Helvetica → Arimo
 *   Times New Roman → Tinos     Courier New → Cousine
 *   (Arimo, Tinos and Cousine are the Liberation Sans / Serif / Mono designs.)
 *
 * Because the substitute answers to the original name, nothing in the workbook changes: text
 * keeps its width, and the file is saved with the original font names. The files are in the app (no network), split by script and
 * loaded only when a workbook uses them. Chinese, Japanese and Korean text falls back to the
 * system's Noto Sans CJK, which every Android has.
 *
 * Other fonts can be added from a .ttf / .otf / .woff / .woff2 file; they are kept in the app's
 * storage and registered at start.
 */
import { kv } from '../state/storage';

export interface Substitute {
  /** substitute family (what is really drawn) */
  family: string;
  /** package folder name */
  pkg: string;
  subsets: string[];
}
const FULL = ['latin', 'latin-ext', 'cyrillic', 'cyrillic-ext', 'greek'];
const SUBS: Record<string, Substitute> = {
  carlito: { family: 'Carlito', pkg: 'carlito', subsets: FULL },
  caladea: { family: 'Caladea', pkg: 'caladea', subsets: ['latin', 'latin-ext'] },
  arimo: { family: 'Arimo', pkg: 'arimo', subsets: FULL },
  tinos: { family: 'Tinos', pkg: 'tinos', subsets: FULL },
  cousine: { family: 'Cousine', pkg: 'cousine', subsets: FULL },
};
/** Excel / Office font name (lower case) → substitute */
const ALIASES: Record<string, keyof typeof SUBS> = {
  calibri: 'carlito',
  'calibri light': 'carlito',
  carlito: 'carlito',
  cambria: 'caladea',
  caladea: 'caladea',
  arial: 'arimo',
  helvetica: 'arimo',
  'liberation sans': 'arimo',
  arimo: 'arimo',
  'times new roman': 'tinos',
  times: 'tinos',
  'liberation serif': 'tinos',
  tinos: 'tinos',
  'courier new': 'cousine',
  courier: 'cousine',
  'liberation mono': 'cousine',
  cousine: 'cousine',
};
/** Names offered in the font list (the Excel names people know). */
export const BUILT_IN = ['Calibri', 'Arial', 'Times New Roman', 'Cambria', 'Courier New'];
const RANGES: Record<string, string> = {
  latin: 'U+0000-00FF,U+0131,U+0152-0153,U+02BB-02BC,U+02C6,U+02DA,U+02DC,U+0304,U+0308,U+0329,U+2000-206F,U+20AC,U+2122,U+2191,U+2193,U+2212,U+2215,U+FEFF,U+FFFD',
  'latin-ext': 'U+0100-02BA,U+02BD-02C5,U+02C7-02CC,U+02CE-02D7,U+02DD-02FF,U+1D00-1DBF,U+1E00-1E9F,U+1EF2-1EFF,U+2020,U+20A0-20AB,U+20AD-20C0,U+2113,U+2C60-2C7F,U+A720-A7FF',
  cyrillic: 'U+0301,U+0400-045F,U+0490-0491,U+04B0-04B1,U+2116',
  'cyrillic-ext': 'U+0460-052F,U+1C80-1C8A,U+20B4,U+2DE0-2DFF,U+A640-A69F,U+FE2E-FE2F',
  greek: 'U+0370-0377,U+037A-037F,U+0384-038A,U+038C,U+038E-03A1,U+03A3-03FF',
};

export const substituteFor = (name: string): Substitute | null => {
  const k = ALIASES[name.trim().toLowerCase()];
  return k ? SUBS[k] : null;
};

/** @font-face rules for every alias (pure: unit-tested). `url` resolves a font file. */
export function fontFaceCss(url: (pkg: string, subset: string, weight: number, style: string) => string | undefined): string {
  const out: string[] = [];
  for (const [alias, key] of Object.entries(ALIASES)) {
    const sub = SUBS[key];
    const name = alias.replace(/\b\w/g, (c) => c.toUpperCase());
    for (const subset of sub.subsets)
      for (const weight of [400, 700])
        for (const style of ['normal', 'italic']) {
          const u = url(sub.pkg, subset, weight, style);
          if (!u) continue;
          // no local(): see hasLocalFont
          out.push(`@font-face{font-family:"${name}";font-style:${style};font-weight:${weight};font-display:swap;src:url(${u}) format("woff2");unicode-range:${RANGES[subset]}}`);
        }
  }
  return out.join('\n');
}

// ---------------------------------------------------------------- runtime
const browser = typeof document !== 'undefined';
const installed = new Set<string>();
let registered = false;
const listeners = new Set<() => void>();
let onFontsLoaded: () => void = () => {};

/**
 * Is a font of this name on the device? Compared by text width against the generic fallbacks.
 * `local()` is deliberately not used anywhere in this file: on Android the WebView answers it by
 * connecting to Google Play services' font provider, and Android then kills every app holding
 * that connection whenever Play services restarts ("depends on provider …FontsProvider in dying
 * proc"). The app came back with its old session and a file just opened seemed to vanish (K-07).
 * The price: an Excel font that really is installed (desktop browsers) is not preferred over the
 * bundled look-alike, and a system alias (Android answers "Arial" with Roboto) counts as missing.
 */
function hasLocalFont(name: string): boolean {
  try {
    const ctx = document.createElement('canvas').getContext('2d')!;
    const text = 'mmmmmmmmmmlliWW@#1';
    const w = (f: string) => ((ctx.font = `72px ${f}`), ctx.measureText(text).width);
    return ['monospace', 'serif', 'sans-serif'].every((base) => w(`"${name.replace(/["\\]/g, '')}", ${base}`) !== w(base));
  } catch {
    return false;
  }
}
const localCache = new Map<string, boolean>();
/** Remember whether these fonts are on the device (only asked for names that have no bundled substitute). */
export async function probeFonts(names: Iterable<string>) {
  if (!browser) return;
  for (const n of new Set(names)) {
    const k = n.toLowerCase();
    if (!n || localCache.has(k) || substituteFor(n)) continue;
    const ok = hasLocalFont(n);
    localCache.set(k, ok);
    if (ok) installed.add(k);
  }
}

/** Register the bundled substitutes and the user's own fonts. Call once, before the first workbook. */
export async function initFonts(onLoaded: () => void = () => {}) {
  if (!browser || registered) return;
  registered = true;
  onFontsLoaded = onLoaded;
  const files = import.meta.glob('/node_modules/@fontsource/{carlito,caladea,arimo,tinos,cousine}/files/*-{latin,latin-ext,cyrillic,cyrillic-ext,greek}-{400,700}-{normal,italic}.woff2', {
    eager: true,
    query: '?url',
    import: 'default',
  }) as Record<string, string>;
  const byName = new Map(Object.entries(files).map(([k, v]) => [k.split('/').pop()!, v]));
  const style = document.createElement('style');
  style.dataset.petak = 'fonts';
  style.textContent = fontFaceCss((pkg, subset, weight, st) => byName.get(`${pkg}-${subset}-${weight}-${st}.woff2`));
  document.head.appendChild(style);
  // a face that finishes loading later (another script, bold…): measure and draw again
  document.fonts.addEventListener('loadingdone', () => onFontsLoaded());
  await loadUserFonts().catch((e) => console.warn('user fonts', e));
}

/** Load the faces a workbook needs before it is measured and drawn. Never throws; gives up after 4 s. */
export async function ensureFonts(names: Iterable<string>, sample = ''): Promise<void> {
  if (!browser) return;
  await probeFonts(names);
  const text = 'Ag' + sample.slice(0, 200);
  const jobs: Promise<unknown>[] = [];
  for (const n of new Set(names)) {
    if (!n || (!substituteFor(n) && !user.some((u) => u.family.toLowerCase() === n.toLowerCase()))) continue;
    for (const v of ['', 'bold ', 'italic ', 'italic bold ']) jobs.push(document.fonts.load(`${v}16px "${n}"`, text).catch(() => null));
  }
  if (!jobs.length) return;
  await Promise.race([Promise.all(jobs), new Promise((r) => setTimeout(r, 4000))]);
}

export interface FontReport {
  /** shown with a bundled font of the same widths: [name in the file, what is drawn] */
  substituted: Array<[string, string]>;
  /** neither on the device nor bundled: shown with the default font */
  missing: string[];
}
/** What happens to each font a workbook uses. */
export function fontReport(names: Iterable<string>): FontReport {
  const substituted: Array<[string, string]> = [];
  const missing: string[] = [];
  for (const n of new Set([...names].map((x) => x.trim()).filter(Boolean))) {
    const k = n.toLowerCase();
    if (installed.has(k) || user.some((u) => u.family.toLowerCase() === k)) continue;
    const sub = substituteFor(n);
    if (sub) {
      if (sub.family.toLowerCase() !== k) substituted.push([n, sub.family]);
    } else if (localCache.get(k) === false) missing.push(n); // unknown until probed
  }
  return { substituted, missing };
}
export function reportText(r: FontReport): string | null {
  const parts: string[] = [];
  if (r.substituted.length > 2) parts.push(`${r.substituted.length} shown with built-in fonts of the same widths`);
  else if (r.substituted.length) parts.push(`${r.substituted.map(([a, b]) => `${a} → ${b}`).join(', ')} (same widths)`);
  if (r.missing.length) parts.push(`not on this device: ${r.missing.join(', ')} (add them under Format → Fonts)`);
  return parts.length ? 'Fonts: ' + parts.join('; ') : null;
}

// ---------------------------------------------------------------- the user's own fonts
export interface UserFont {
  id: string;
  family: string;
  bold: boolean;
  italic: boolean;
  file: string;
  size: number;
}
const USER_KEY = 'fonts:list';
let user: UserFont[] = [];
const faces = new Map<string, FontFace>();

/** Family and style from a TrueType / OpenType file's name table (null for other formats). */
export function readFontName(buf: ArrayBuffer): { family: string; bold: boolean; italic: boolean } | null {
  try {
    const d = new DataView(buf);
    const tag = d.getUint32(0);
    if (tag !== 0x00010000 && tag !== 0x4f54544f && tag !== 0x74727565) return null; // TTF, OTTO, true
    const n = d.getUint16(4);
    let off = 0;
    for (let i = 0; i < n; i++) {
      const p = 12 + i * 16;
      if (d.getUint32(p) === 0x6e616d65) off = d.getUint32(p + 8); // 'name'
    }
    if (!off) return null;
    const count = d.getUint16(off + 2);
    const strings = off + d.getUint16(off + 4);
    const names: Record<number, string> = {};
    for (let i = 0; i < count; i++) {
      const r = off + 6 + i * 12;
      const platform = d.getUint16(r);
      const id = d.getUint16(r + 6);
      const len = d.getUint16(r + 8);
      const at = strings + d.getUint16(r + 10);
      if (![1, 2, 16, 17].includes(id)) continue;
      let s = '';
      if (platform === 3 || platform === 0) for (let k = 0; k + 1 < len; k += 2) s += String.fromCharCode(d.getUint16(at + k));
      else if (!names[id]) for (let k = 0; k < len; k++) s += String.fromCharCode(d.getUint8(at + k));
      if (s && (platform === 3 || !names[id])) names[id] = s;
    }
    const family = names[1] ?? names[16];
    if (!family) return null;
    const sub = (names[2] ?? names[17] ?? '').toLowerCase();
    return { family, bold: /bold|black|heavy/.test(sub), italic: /italic|oblique/.test(sub) };
  } catch {
    return null;
  }
}
/** Family and style guessed from a file name: "Roboto-BoldItalic.ttf". */
export function nameFromFile(file: string): { family: string; bold: boolean; italic: boolean } {
  const base = file.replace(/\.[^.]+$/, '');
  const m = /^(.*?)[-_ ]?(bold ?italic|bold ?oblique|bold|italic|oblique|regular)?$/i.exec(base)!;
  const style = (m[2] ?? '').toLowerCase();
  return { family: (m[1] || base).replace(/[-_]+/g, ' ').trim() || base, bold: style.includes('bold'), italic: /italic|oblique/.test(style) };
}

async function register(f: UserFont, data: ArrayBuffer) {
  const face = new FontFace(f.family, data, { weight: f.bold ? '700' : '400', style: f.italic ? 'italic' : 'normal' });
  await face.load();
  document.fonts.add(face);
  faces.set(f.id, face);
}
async function loadUserFonts() {
  user = (await kv.get<UserFont[]>(USER_KEY)) ?? [];
  for (const f of user) {
    const data = await kv.get<ArrayBuffer>('fonts:data:' + f.id);
    if (data) await register(f, data).catch((e) => console.warn('font', f.file, e));
  }
  emit();
}
const emit = () => listeners.forEach((l) => l());

export const fonts = {
  subscribe(l: () => void) {
    listeners.add(l);
    return () => listeners.delete(l);
  },
  user: () => user,
  /** every family that can be chosen: bundled first, then the user's */
  families(): string[] {
    return [...BUILT_IN, ...[...new Set(user.map((u) => u.family))].filter((f) => !BUILT_IN.some((b) => b.toLowerCase() === f.toLowerCase()))];
  },
  /** Add a font file. Throws a readable error when the file is not a usable font. */
  async add(file: string, data: ArrayBuffer): Promise<UserFont> {
    if (data.byteLength > 30 * 1024 * 1024) throw new Error('This font file is larger than 30 MB');
    const meta = readFontName(data) ?? nameFromFile(file);
    const f: UserFont = { id: Date.now().toString(36) + Math.random().toString(36).slice(2, 6), ...meta, file, size: data.byteLength };
    try {
      await register(f, data);
    } catch {
      throw new Error(`“${file}” is not a font this device can use (TTF, OTF, WOFF or WOFF2)`);
    }
    // the same family and style again replaces the earlier file
    const old = user.filter((u) => u.family === f.family && u.bold === f.bold && u.italic === f.italic);
    for (const o of old) await this.remove(o.id, false);
    user = [...user, f];
    await kv.set('fonts:data:' + f.id, data);
    await kv.set(USER_KEY, user);
    emit();
    onFontsLoaded();
    return f;
  },
  async remove(id: string, notify = true) {
    const face = faces.get(id);
    if (face) document.fonts.delete(face);
    faces.delete(id);
    user = user.filter((u) => u.id !== id);
    await kv.del('fonts:data:' + id);
    await kv.set(USER_KEY, user);
    if (notify) {
      emit();
      onFontsLoaded();
    }
  },
};
if (browser) (globalThis as unknown as { __fonts: unknown }).__fonts = { fonts, fontReport, ensureFonts, substituteFor, installed: () => [...installed] };
