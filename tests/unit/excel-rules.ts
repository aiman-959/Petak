/**
 * Checks an .xlsx / .xlsm package against the rules Excel enforces when it opens a file. Breaking
 * one of these gives "We found a problem with some content … Do you want us to try to recover?":
 *
 *  - every part is well-formed XML (strict parser below: nesting, one root, attributes once,
 *    known entities), and every part has a content type
 *  - every relationship of the package points at a part that exists; relationship ids are unique
 *  - workbook: sheet names unique (case-insensitive), at most 31 characters, none of : \ / ? * [ ];
 *    sheet ids unique; every sheet's r:id is a worksheet relationship; defined names valid
 *  - worksheet: child elements in the order of the schema (CT_Worksheet); rows ascending and
 *    unique; cells ascending inside a row and in that row; style index inside cellXfs; shared
 *    string index inside the table; booleans 0/1; numbers are numbers; column ranges ordered and
 *    not overlapping; merged ranges not overlapping and their count right; hyperlink r:ids exist
 *  - styles: every cellXfs entry's font / fill / border / number format exists; counts match
 * Independent of Petak's own reader and writer (it does not use src/io).
 */
import { strFromU8, unzipSync } from 'fflate';

export interface XNode {
  name: string;
  attrs: Record<string, string>;
  children: XNode[];
  text: string;
}

/** Strict XML parser (enough for OOXML): throws on anything Excel's parser would reject. */
export function parseXml(xml: string, part = ''): XNode {
  let i = 0;
  const err = (m: string): never => {
    throw new Error(`${part}: ${m} at ${i} (…${xml.slice(Math.max(0, i - 40), i + 40).replace(/\n/g, ' ')}…)`);
  };
  const entity = (s: string) => {
    const bad = /&(?!(?:lt|gt|amp|quot|apos|#\d+|#x[0-9a-fA-F]+);)/.exec(s);
    if (bad) err('bad entity ' + s.slice(bad.index, bad.index + 10));
    return s;
  };
  if (xml.charCodeAt(0) === 0xfeff) i = 1;
  const skipMisc = () => {
    for (;;) {
      while (i < xml.length && /\s/.test(xml[i])) i++;
      if (xml.startsWith('<?', i)) {
        const e = xml.indexOf('?>', i);
        if (e < 0) err('unclosed declaration');
        i = e + 2;
      } else if (xml.startsWith('<!--', i)) {
        const e = xml.indexOf('-->', i);
        if (e < 0) err('unclosed comment');
        i = e + 3;
      } else return;
    }
  };
  const nameRe = /[A-Za-z_][\w.:-]*/y;
  const readName = () => {
    nameRe.lastIndex = i;
    const m = nameRe.exec(xml);
    if (!m) err('name expected');
    i += m![0].length;
    return m![0];
  };
  const element = (): XNode => {
    if (xml[i] !== '<') err('< expected');
    i++;
    const name = readName();
    const attrs: Record<string, string> = {};
    for (;;) {
      while (/\s/.test(xml[i])) i++;
      if (xml.startsWith('/>', i)) {
        i += 2;
        return { name, attrs, children: [], text: '' };
      }
      if (xml[i] === '>') {
        i++;
        break;
      }
      const a = readName();
      while (/\s/.test(xml[i])) i++;
      if (xml[i] !== '=') err(`= expected after ${a}`);
      i++;
      while (/\s/.test(xml[i])) i++;
      const q = xml[i];
      if (q !== '"' && q !== "'") err('quote expected');
      const e = xml.indexOf(q, i + 1);
      if (e < 0) err('unclosed attribute');
      const v = xml.slice(i + 1, e);
      if (v.includes('<')) err('< in attribute');
      if (a in attrs) err(`attribute ${a} twice on <${name}>`);
      attrs[a] = entity(v);
      i = e + 1;
    }
    const node: XNode = { name, attrs, children: [], text: '' };
    for (;;) {
      const lt = xml.indexOf('<', i);
      if (lt < 0) err(`<${name}> not closed`);
      if (lt > i) node.text += entity(xml.slice(i, lt));
      i = lt;
      if (xml.startsWith('</', i)) {
        i += 2;
        const n = readName();
        if (n !== name) err(`</${n}> closes <${name}>`);
        while (/\s/.test(xml[i])) i++;
        if (xml[i] !== '>') err('> expected');
        i++;
        return node;
      }
      if (xml.startsWith('<!--', i)) {
        const e = xml.indexOf('-->', i);
        if (e < 0) err('unclosed comment');
        i = e + 3;
        continue;
      }
      if (xml.startsWith('<![CDATA[', i)) {
        const e = xml.indexOf(']]>', i);
        if (e < 0) err('unclosed CDATA');
        node.text += xml.slice(i + 9, e);
        i = e + 3;
        continue;
      }
      node.children.push(element());
    }
  };
  skipMisc();
  const root = element();
  skipMisc();
  if (i < xml.length) err('content after the root element');
  return root;
}

const local = (n: string) => n.slice(n.indexOf(':') + 1);
const kids = (n: XNode, name: string) => n.children.filter((c) => local(c.name) === name);
const kid = (n: XNode, name: string) => n.children.find((c) => local(c.name) === name);

/** CT_Worksheet's sequence (ECMA-376 Part 1, 18.3.1.99). Excel rejects other orders. */
const SHEET_ORDER = [
  'sheetPr', 'dimension', 'sheetViews', 'sheetFormatPr', 'cols', 'sheetData', 'sheetCalcPr', 'sheetProtection', 'protectedRanges', 'scenarios', 'autoFilter', 'sortState',
  'dataConsolidate', 'customSheetViews', 'mergeCells', 'phoneticPr', 'conditionalFormatting', 'dataValidations', 'hyperlinks', 'printOptions', 'pageMargins', 'pageSetup',
  'headerFooter', 'rowBreaks', 'colBreaks', 'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags', 'drawing', 'legacyDrawing', 'legacyDrawingHF', 'drawingHF',
  'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst',
];
/** CT_Stylesheet's sequence */
const STYLE_ORDER = ['numFmts', 'fonts', 'fills', 'borders', 'cellStyleXfs', 'cellXfs', 'cellStyles', 'dxfs', 'tableStyles', 'colors', 'extLst'];
/** number format ids below 164 are Excel's own (built in); 164 and up must be defined in styles.xml */
const isBuiltinNumFmt = (id: number) => id >= 0 && id < 164;

const colNum = (s: string) => [...s].reduce((a, ch) => a * 26 + ch.charCodeAt(0) - 64, 0);
function ref(r: string) {
  const m = /^\$?([A-Z]{1,3})\$?(\d+)$/.exec(r);
  return m ? { c: colNum(m[1]), r: +m[2] } : null;
}
function range(s: string) {
  const [a, b = a] = s.split(':');
  const p = ref(a);
  const q = ref(b);
  return p && q ? { r0: Math.min(p.r, q.r), c0: Math.min(p.c, q.c), r1: Math.max(p.r, q.r), c1: Math.max(p.c, q.c) } : null;
}
const resolve = (base: string, target: string) => {
  if (target.startsWith('/')) return target.slice(1);
  const parts = base.split('/').slice(0, -1);
  for (const seg of target.split('/')) {
    if (seg === '..') parts.pop();
    else if (seg !== '.') parts.push(seg);
  }
  return parts.join('/');
};

/** All problems Excel would report for this package (empty: it opens cleanly). */
export function excelProblems(buf: ArrayBuffer | Uint8Array): string[] {
  const out: string[] = [];
  const zip = unzipSync(buf instanceof Uint8Array ? buf : new Uint8Array(buf));
  const names = Object.keys(zip).filter((k) => !k.endsWith('/'));
  const xml: Record<string, XNode> = {};
  for (const n of names) {
    if (!/\.(xml|rels|vml)$/i.test(n)) continue;
    try {
      xml[n] = parseXml(strFromU8(zip[n]), n);
    } catch (e) {
      out.push(String((e as Error).message).slice(0, 300));
    }
  }
  const ct = xml['[Content_Types].xml'];
  if (!ct) return [...out, 'no [Content_Types].xml'];
  const defaults = new Map(kids(ct, 'Default').map((d) => [d.attrs.Extension?.toLowerCase(), d.attrs.ContentType]));
  const overrides = new Map(kids(ct, 'Override').map((o) => [o.attrs.PartName?.replace(/^\//, ''), o.attrs.ContentType]));
  for (const n of names) if (n !== '[Content_Types].xml' && !overrides.has(n) && !defaults.has(n.split('.').pop()!.toLowerCase())) out.push(`${n}: no content type`);
  for (const p of overrides.keys()) if (!names.includes(p!)) out.push(`[Content_Types].xml: override for missing part ${p}`);
  // relationships
  const rels = new Map<string, Map<string, { type: string; target: string; external: boolean }>>();
  for (const n of names.filter((x) => x.endsWith('.rels'))) {
    const r = xml[n];
    if (!r) continue;
    const owner = n.replace(/_rels\/([^/]+)\.rels$/, '$1').replace(/^_rels\/\.rels$/, '');
    const map = new Map<string, { type: string; target: string; external: boolean }>();
    for (const rel of kids(r, 'Relationship')) {
      const { Id, Type = '', Target = '', TargetMode } = rel.attrs;
      if (map.has(Id)) out.push(`${n}: relationship id ${Id} twice`);
      const external = TargetMode === 'External';
      const target = external ? Target : resolve(owner === '.rels' || n === '_rels/.rels' ? '' : owner, Target);
      map.set(Id, { type: Type.split('/').pop()!, target, external });
      if (!external && !names.includes(target)) out.push(`${n}: ${Id} points at missing part ${target}`);
    }
    rels.set(n === '_rels/.rels' ? '' : owner, map);
  }
  const relsOf = (part: string) => rels.get(part) ?? new Map();
  // workbook
  const wbRel = [...relsOf('').values()].find((r) => r.type === 'officeDocument');
  if (!wbRel) return [...out, 'no officeDocument relationship'];
  const wbPart = wbRel.target;
  const wb = xml[wbPart];
  if (!wb) return [...out, `${wbPart} missing`];
  const sheets = kids(kid(wb, 'sheets') ?? { children: [] } as never, 'sheet');
  const seenNames = new Set<string>();
  const seenIds = new Set<string>();
  const sheetParts: string[] = [];
  for (const s of sheets) {
    const nm = s.attrs.name ?? '';
    if (!nm || nm.length > 31 || /[:\\/?*[\]]/.test(nm) || /^'|'$/.test(nm)) out.push(`${wbPart}: invalid sheet name "${nm}"`);
    if (seenNames.has(nm.toLowerCase())) out.push(`${wbPart}: sheet name "${nm}" twice`);
    seenNames.add(nm.toLowerCase());
    if (seenIds.has(s.attrs.sheetId)) out.push(`${wbPart}: sheetId ${s.attrs.sheetId} twice`);
    seenIds.add(s.attrs.sheetId);
    const rid = s.attrs['r:id'];
    const rel = relsOf(wbPart).get(rid);
    if (!rel) out.push(`${wbPart}: sheet "${nm}" r:id ${rid} has no relationship`);
    else if (rel.type === 'worksheet') sheetParts.push(rel.target);
  }
  for (const dn of kids(kid(wb, 'definedNames') ?? ({ children: [] } as never), 'definedName')) {
    const n = dn.attrs.name ?? '';
    if (!/^(_xlnm\.)?[A-Za-z_\\][\w.\\]*$/.test(n) || /^[A-Za-z]{1,3}\d+$/.test(n) || /^[rRcC]\d*$/.test(n)) out.push(`${wbPart}: invalid defined name "${n}"`);
    if (dn.attrs.localSheetId !== undefined && +dn.attrs.localSheetId >= sheets.length) out.push(`${wbPart}: defined name ${n} for sheet ${dn.attrs.localSheetId} of ${sheets.length}`);
    if (!dn.text.trim()) out.push(`${wbPart}: defined name ${n} is empty`);
  }
  // styles and shared strings
  const wbRels = [...relsOf(wbPart).values()];
  const stylesPart = wbRels.find((r) => r.type === 'styles')?.target;
  const st = stylesPart ? xml[stylesPart] : undefined;
  let xfCount = 1;
  if (st) {
    const order = st.children.map((c) => local(c.name)).filter((n) => STYLE_ORDER.includes(n));
    const idx = order.map((n) => STYLE_ORDER.indexOf(n));
    if (idx.some((v, k) => k > 0 && v <= idx[k - 1])) out.push(`${stylesPart}: elements out of order (${order.join(', ')})`);
    const coll = (n: string, item: string) => {
      const c = kid(st, n);
      const list = c ? kids(c, item) : [];
      if (c && c.attrs.count !== undefined && +c.attrs.count !== list.length) out.push(`${stylesPart}: <${n} count="${c.attrs.count}"> has ${list.length}`);
      return list;
    };
    const numFmts = new Set(coll('numFmts', 'numFmt').map((f) => +f.attrs.numFmtId));
    const fonts = coll('fonts', 'font').length;
    const fills = coll('fills', 'fill').length;
    const borders = coll('borders', 'border').length;
    const styleXfs = coll('cellStyleXfs', 'xf').length;
    const xfs = coll('cellXfs', 'xf');
    coll('dxfs', 'dxf');
    xfCount = xfs.length;
    if (!xfs.length) out.push(`${stylesPart}: no cellXfs`);
    xfs.forEach((x, k) => {
      const a = x.attrs;
      if (+(a.fontId ?? 0) >= fonts) out.push(`${stylesPart}: xf ${k} fontId ${a.fontId} of ${fonts}`);
      if (+(a.fillId ?? 0) >= fills) out.push(`${stylesPart}: xf ${k} fillId ${a.fillId} of ${fills}`);
      if (+(a.borderId ?? 0) >= borders) out.push(`${stylesPart}: xf ${k} borderId ${a.borderId} of ${borders}`);
      if (a.xfId !== undefined && +a.xfId >= styleXfs) out.push(`${stylesPart}: xf ${k} xfId ${a.xfId} of ${styleXfs}`);
      const nf = +(a.numFmtId ?? 0);
      if (!numFmts.has(nf) && !isBuiltinNumFmt(nf)) out.push(`${stylesPart}: xf ${k} numFmtId ${nf} not defined`);
    });
  } else out.push('no styles part');
  const ssPart = wbRels.find((r) => r.type === 'sharedStrings')?.target;
  const ssCount = ssPart && xml[ssPart] ? kids(xml[ssPart], 'si').length : 0;
  if (ssPart && xml[ssPart]?.attrs.uniqueCount !== undefined && +xml[ssPart].attrs.uniqueCount !== ssCount) out.push(`${ssPart}: uniqueCount ${xml[ssPart].attrs.uniqueCount} but ${ssCount} strings`);
  // worksheets
  for (const sp of sheetParts) {
    const ws = xml[sp];
    if (!ws) continue;
    const order = ws.children.map((c) => local(c.name));
    let last = -1;
    for (const n of order) {
      const k = SHEET_ORDER.indexOf(n);
      if (k < 0) {
        out.push(`${sp}: unknown element <${n}>`);
        continue;
      }
      // conditionalFormatting may repeat; everything else appears once and in order
      if (k < last || (k === last && n !== 'conditionalFormatting')) out.push(`${sp}: <${n}> out of order (${order.join(', ')})`);
      last = k;
    }
    const data = kid(ws, 'sheetData');
    if (!data) out.push(`${sp}: no sheetData`);
    let prevRow = 0;
    for (const row of data ? kids(data, 'row') : []) {
      const r = +row.attrs.r;
      if (!(r > prevRow)) out.push(`${sp}: row ${r} after row ${prevRow}`);
      if (r > 1048576) out.push(`${sp}: row ${r} beyond the sheet`);
      prevRow = r;
      if (row.attrs.s !== undefined && +row.attrs.s >= xfCount) out.push(`${sp}: row ${r} style ${row.attrs.s} of ${xfCount}`);
      let prevCol = 0;
      for (const c of kids(row, 'c')) {
        const p = c.attrs.r ? ref(c.attrs.r) : { r, c: prevCol + 1 };
        if (!p) {
          out.push(`${sp}: bad cell reference ${c.attrs.r}`);
          continue;
        }
        if (p.r !== r) out.push(`${sp}: cell ${c.attrs.r} inside row ${r}`);
        if (!(p.c > prevCol)) out.push(`${sp}: cell ${c.attrs.r} after column ${prevCol}`);
        if (p.c > 16384) out.push(`${sp}: cell ${c.attrs.r} beyond the sheet`);
        prevCol = p.c;
        if (c.attrs.s !== undefined && +c.attrs.s >= xfCount) out.push(`${sp}: ${c.attrs.r} style ${c.attrs.s} of ${xfCount}`);
        const v = kid(c, 'v')?.text;
        const t = c.attrs.t ?? 'n';
        if (v !== undefined) {
          if (t === 's' && !(+v >= 0 && +v < ssCount && Number.isInteger(+v))) out.push(`${sp}: ${c.attrs.r} shared string ${v} of ${ssCount}`);
          if (t === 'b' && v !== '0' && v !== '1') out.push(`${sp}: ${c.attrs.r} boolean "${v}"`);
          if (t === 'n' && v.trim() !== '' && !Number.isFinite(+v)) out.push(`${sp}: ${c.attrs.r} number "${v}"`);
        }
        if (!['n', 's', 'str', 'b', 'e', 'inlineStr', 'd'].includes(t)) out.push(`${sp}: ${c.attrs.r} type "${t}"`);
        const f = kid(c, 'f');
        if (f && f.attrs.t === 'shared' && f.attrs.si === undefined) out.push(`${sp}: ${c.attrs.r} shared formula without si`);
        if (f && f.attrs.ref && !range(f.attrs.ref)) out.push(`${sp}: ${c.attrs.r} formula ref ${f.attrs.ref}`);
      }
    }
    const cols = kid(ws, 'cols');
    let prevMax = 0;
    for (const col of cols ? kids(cols, 'col') : []) {
      const mn = +col.attrs.min;
      const mx = +col.attrs.max;
      if (!(mn >= 1 && mx >= mn && mx <= 16384)) out.push(`${sp}: <col min="${mn}" max="${mx}">`);
      if (mn <= prevMax) out.push(`${sp}: columns ${mn}–${mx} overlap the ones before`);
      prevMax = mx;
      if (col.attrs.style !== undefined && +col.attrs.style >= xfCount) out.push(`${sp}: column style ${col.attrs.style} of ${xfCount}`);
    }
    const mc = kid(ws, 'mergeCells');
    if (mc) {
      const list = kids(mc, 'mergeCell').map((m) => ({ s: m.attrs.ref, r: range(m.attrs.ref ?? '') }));
      if (mc.attrs.count !== undefined && +mc.attrs.count !== list.length) out.push(`${sp}: <mergeCells count="${mc.attrs.count}"> has ${list.length}`);
      for (let a = 0; a < list.length; a++) {
        if (!list[a].r) {
          out.push(`${sp}: merge ${list[a].s}`);
          continue;
        }
        for (let b = a + 1; b < list.length; b++) {
          const p = list[a].r!;
          const q = list[b].r;
          if (q && p.r0 <= q.r1 && q.r0 <= p.r1 && p.c0 <= q.c1 && q.c0 <= p.c1) out.push(`${sp}: merges ${list[a].s} and ${list[b].s} overlap`);
        }
      }
    }
    for (const h of kids(kid(ws, 'hyperlinks') ?? ({ children: [] } as never), 'hyperlink')) {
      if (!range(h.attrs.ref ?? '')) out.push(`${sp}: hyperlink ref ${h.attrs.ref}`);
      const rid = h.attrs['r:id'];
      if (rid && !relsOf(sp).has(rid)) out.push(`${sp}: hyperlink ${h.attrs.ref} r:id ${rid} has no relationship`);
    }
    for (const d of ['drawing', 'legacyDrawing']) {
      const e = kid(ws, d);
      if (e && !relsOf(sp).has(e.attrs['r:id'])) out.push(`${sp}: <${d}> r:id ${e.attrs['r:id']} has no relationship`);
    }
    for (const tp of kids(kid(ws, 'tableParts') ?? ({ children: [] } as never), 'tablePart')) if (!relsOf(sp).has(tp.attrs['r:id'])) out.push(`${sp}: tablePart r:id ${tp.attrs['r:id']} has no relationship`);
  }
  return out;
}
