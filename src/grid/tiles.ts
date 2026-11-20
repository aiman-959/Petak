/**
 * Sheet tiles for the gesture renderer (see gestureView.ts).
 *
 * A tile is a small canvas (512 device pixels a side, plus one pixel of bleed all round) showing a
 * rectangle of the sheet at one zoom *level*, drawn by Univer's own sheet renderer
 * (`Spreadsheet.render`, the path Univer uses for printing), so it shows exactly what the grid
 * shows: fills, borders, fonts, merged cells, overflowing text, conditional formats, sparklines,
 * pictures and drawing shapes (the last three are sheet extensions too). Row and column headers
 * have their own strip tiles, drawn by Univer's header components.
 *
 * Levels work like a map app's zoom levels: 12.5 %, 25 %, 50 %, 100 %, 200 %, 400 %, plus the zoom
 * the sheet is resting at (so a pan at, say, 85 % is as sharp as the grid itself). Tiles are
 * opaque: the sheet background goes *under* what Univer draws (Univer clears the grid lines under
 * merged cells to transparent), so nothing behind a tile shows through and nothing is ever black.
 *
 * Tiles are shown as ImageBitmaps in 'bitmaprenderer' canvases, which the gesture renderer puts in
 * layers that the compositor moves and scales. Not as 2D canvases: Chrome copies the pixels of
 * every 2D canvas on screen at every frame it commits, changed or not (measured: 16 tiles at 60 fps,
 * 960 copies a second, 180 ms of main-thread time; a bitmap canvas costs nothing per frame). Univer
 * draws into a scratch canvas that is never in the page; the result is moved into an ImageBitmap.
 *
 * Memory: every tile is the same small bitmap (about 1 MB), far below any WebView canvas limit;
 * the cache is capped in bytes per device tier and evicts the tiles farthest from the view first.
 * Elements of evicted tiles are reused.
 */
import { Canvas as UCanvas } from '@univerjs/engine-render';
import { renderUnit } from './view';
export { LEVELS, TILE, span } from './camera';

/** Device pixels drawn beyond each edge, so neighbouring tiles overlap and no seam shows when scaled. */
export const BLEED = 1;

/**
 * c / ct / cl / cc  cells: scrolling pane, frozen rows (top), frozen columns (left), frozen corner
 * h / hf            column headers: scrolling part, frozen part
 * v / vf            row headers: scrolling part, frozen part
 * Frozen panes have their own tiles (cut at the freeze line), so a canvas is only ever in one pane.
 */
export type TileKind = 'c' | 'ct' | 'cl' | 'cc' | 'h' | 'hf' | 'v' | 'vf';
export const isCells = (k: TileKind) => k[0] === 'c';
export const isColHeader = (k: TileKind) => k[0] === 'h';
export const isRowHeader = (k: TileKind) => k[0] === 'v';

interface Bound {
  left: number;
  top: number;
  right: number;
  bottom: number;
}
interface USkeleton {
  rowHeaderWidthAndMarginLeft: number;
  columnHeaderHeightAndMarginTop: number;
  rowHeightAccumulation: number[];
  columnWidthAccumulation: number[];
  worksheet: { getSheetId(): string };
  _drawingRange?: unknown;
  _visibleRangeMap?: Map<string, unknown>;
  _cacheRangeMap?: Map<string, unknown>;
  gridlinesColor?: string;
  defaultGridlinesColor?: string;
  showGridlines?: number;
}
interface UComponent {
  getSkeleton(): USkeleton | undefined;
  getParentScale(): { scaleX: number; scaleY: number };
  render?(ctx: unknown, vp: unknown): unknown;
  draw?(ctx: unknown, b: unknown): unknown;
}
interface UCanvasLike {
  getCanvasEle(): HTMLCanvasElement;
  getContext(): CanvasRenderingContext2D & { save(): void; restore(): void };
  setSize(w: number, h: number, ratio: number): void;
}

/** The sheet rectangle (sheet px at 100 %) a tile shows. */
export interface TileRect {
  sx: number;
  sy: number;
  sw: number;
  sh: number;
}
export interface Tile extends TileRect {
  key: string;
  kind: TileKind;
  level: number;
  /** the element on screen: a 'bitmaprenderer' canvas (a 2D canvas where bitmaps are missing) */
  el: HTMLCanvasElement;
  /** device pixels of the content (without the bleed) */
  pw: number;
  ph: number;
  ver: number;
  /** only background and grid lines (a placeholder until the real tile is drawn) */
  plain: boolean;
  /** the sheet's extent (sheet px) when it was drawn: beyond it the tile has no cells */
  ew: number;
  eh: number;
  /** last frame that showed it */
  seen: number;
  bytes: number;
  /** set by the renderer: where the element was last placed */
  placed?: string;
}

export interface SheetParts {
  sheet: UComponent;
  rowHeader: UComponent | null;
  colHeader: UComponent | null;
  sk: USkeleton;
  /** sheet px: row header width, column header height (at 100 %) */
  rh: number;
  ch: number;
  colorService: unknown;
}

let bitmapOk: boolean | undefined;
/** OffscreenCanvas.transferToImageBitmap and 'bitmaprenderer' canvases (Chrome 69+, every WebView Petak runs on). */
export function bitmapsWork(): boolean {
  if (bitmapOk === undefined) {
    try {
      bitmapOk =
        typeof OffscreenCanvas !== 'undefined' &&
        typeof OffscreenCanvas.prototype.transferToImageBitmap === 'function' &&
        !!document.createElement('canvas').getContext('bitmaprenderer');
    } catch {
      bitmapOk = false;
    }
  }
  return bitmapOk;
}
const offscreens = new Map<string, OffscreenCanvas>();
type Ctx2D = CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D;

/**
 * Put a picture in a display element: `paint` draws it (w × h device px) and the result is shown
 * by `el` without a 2D canvas on screen (see the top of this file). Synchronous: the bitmap is made
 * on a reused OffscreenCanvas and handed to the element without a copy.
 */
export function present(el: HTMLCanvasElement, w: number, h: number, paint: (ctx: Ctx2D) => void) {
  if (el.width !== w) el.width = w;
  if (el.height !== h) el.height = h;
  if (!bitmapsWork()) {
    const ctx = el.getContext('2d')!;
    ctx.setTransform(1, 0, 0, 1, 0, 0);
    ctx.globalCompositeOperation = 'source-over';
    paint(ctx);
    return;
  }
  const key = w + 'x' + h;
  let off = offscreens.get(key);
  if (!off) {
    if (offscreens.size > 6) offscreens.clear();
    off = new OffscreenCanvas(w, h);
    offscreens.set(key, off);
  }
  const ctx = off.getContext('2d')!;
  ctx.setTransform(1, 0, 0, 1, 0, 0);
  ctx.globalCompositeOperation = 'source-over';
  ctx.globalAlpha = 1;
  paint(ctx);
  (el.getContext('bitmaprenderer') as ImageBitmapRenderingContext).transferFromImageBitmap(off.transferToImageBitmap());
}
/**
 * Let go of an element's picture (its memory) while the element waits for reuse. Shrinking the
 * element does it: \`transferFromImageBitmap(null)\` would allocate a blank bitmap of the full size
 * instead (measured: 40–60 ms per eviction at Mid).
 */
export function release(el: HTMLCanvasElement) {
  el.width = 1;
  el.height = 1;
}

/** Univer's render objects for the active sheet (null while there is no grid). */
export function sheetParts(): SheetParts | null {
  const unit = renderUnit() as unknown as { scene: { getObject(k: string): UComponent | undefined }; engine: { canvasColorService?: unknown } } | null;
  if (!unit) return null;
  const sheet = unit.scene.getObject('__SpreadsheetRender__');
  const sk = sheet?.getSkeleton();
  if (!sheet || !sk) return null;
  return {
    sheet,
    rowHeader: unit.scene.getObject('__SpreadsheetRowHeader__') ?? null,
    colHeader: unit.scene.getObject('__SpreadsheetColumnHeader__') ?? null,
    sk,
    rh: sk.rowHeaderWidthAndMarginLeft,
    ch: sk.columnHeaderHeightAndMarginTop,
    colorService: unit.engine.canvasColorService,
  };
}

export class TileStore {
  private tiles = new Map<string, Tile>();
  /** display elements of evicted tiles, for reuse */
  private pool: HTMLCanvasElement[] = [];
  /** Univer draws into these (one per tile size; never in the page) */
  private scratch = new Map<string, UCanvasLike>();
  /** content version: tiles of an older version are not drawn */
  ver = 1;
  bytes = 0;
  capBytes: number;
  rendered = 0;
  /** running average render cost per kind and level (ms), used to decide what fits in a frame */
  private costs = new Map<string, number>();
  cost(kind: TileKind, level: number): number {
    return this.costs.get(kind[0] + level) ?? (isCells(kind) ? (level < 0.4 ? 14 : 5) : 1);
  }
  private sheetKey = '';
  private sk: unknown = null;
  dpr = 1;
  bg = '#ffffff';
  headerBg = '#f6f8f7';
  lineColor = '#e0e3e1';
  /** called when a tile's canvas leaves the cache (the renderer takes it out of the DOM) */
  onEvict: (t: Tile) => void = () => {};

  constructor(capBytes: number) {
    this.capBytes = capBytes;
  }

  get size() {
    return this.tiles.size;
  }
  all() {
    return this.tiles.values();
  }
  get(key: string): Tile | undefined {
    const t = this.tiles.get(key);
    return t && t.ver === this.ver ? t : undefined;
  }

  /** Drop everything when the sheet (or its layout object) changes. */
  checkSheet(parts: SheetParts, dpr: number) {
    const key = parts.sk.worksheet.getSheetId() + '|' + dpr + '|' + parts.rh + '|' + parts.ch;
    if (key !== this.sheetKey || this.sk !== parts.sk) {
      this.sheetKey = key;
      this.sk = parts.sk;
      this.dpr = dpr;
      this.clear();
    }
  }
  /** Something on the sheet changed: every tile is out of date. */
  invalidate() {
    this.ver++;
  }
  clear() {
    for (const t of this.tiles.values()) this.drop(t);
    this.tiles.clear();
    this.bytes = 0;
    this.ver++;
  }
  private drop(t: Tile) {
    this.onEvict(t);
    release(t.el);
    if (this.pool.length < 24) this.pool.push(t.el);
  }

  private element(): HTMLCanvasElement {
    const el = this.pool.pop() ?? document.createElement('canvas');
    return el;
  }

  /**
   * Univer's canvas for drawing a tile of this size. Univer draws through its own context wrapper,
   * which caches state (the transform, colours): nothing else ever draws on these.
   */
  private scratchFor(w: number, h: number, parts: SheetParts): UCanvasLike {
    const key = w + 'x' + h;
    let u = this.scratch.get(key);
    if (!u) {
      if (this.scratch.size > 6) this.scratch.clear();
      // its context has the precision helpers the renderer needs; the colour service is the
      // engine's, so dark mode turns colours the same way as on the grid
      u = new (UCanvas as unknown as new (p: object) => UCanvasLike)({ width: w, height: h, pixelRatio: 1, colorService: parts.colorService });
      this.scratch.set(key, u);
    }
    return u;
  }

  /** Take the tile's slot (a placeholder or an old version) and reuse its element. */
  private slot(key: string) {
    const old = this.tiles.get(key);
    if (old) {
      this.tiles.delete(key);
      this.bytes -= old.bytes;
      return { el: old.el, old };
    }
    return { el: this.element(), old: undefined };
  }

  private store(key: string, kind: TileKind, level: number, r: TileRect, el: HTMLCanvasElement, pw: number, ph: number, plain: boolean, ew: number, eh: number, frame: number, old?: Tile) {
    const bytes = el.width * el.height * 4;
    // a new size or a new element: the renderer places it again
    const placed = old && old.el === el && old.pw === pw && old.ph === ph ? old.placed : undefined;
    const tile: Tile = { key, kind, level, ...r, el, pw, ph, ver: this.ver, plain, ew, eh, seen: frame, bytes, placed };
    this.tiles.set(key, tile);
    this.bytes += bytes;
    return tile;
  }

  private dims(level: number, r: TileRect) {
    const s = level * this.dpr;
    return { pw: Math.max(1, Math.ceil(r.sw * s)), ph: Math.max(1, Math.ceil(r.sh * s)) };
  }

  /**
   * Draw the tile with Univer now (the caller decides whether it fits in the frame). Returns null when
   * Univer's objects are not available.
   */
  render(parts: SheetParts, key: string, kind: TileKind, level: number, r: TileRect, frame: number): Tile | null {
    const t0 = performance.now();
    const { rh, ch, sk } = parts;
    const { pw, ph } = this.dims(level, r);
    const comp = isCells(kind) ? parts.sheet : isColHeader(kind) ? parts.colHeader : parts.rowHeader;
    if (!comp) return null;
    const W = pw + 2 * BLEED;
    const H = ph + 2 * BLEED;
    const canvas = this.scratchFor(W, H, parts);
    const src = canvas.getCanvasEle();
    const ctx = canvas.getContext() as unknown as CanvasRenderingContext2D;
    const s = level * this.dpr;
    // the renderer moves Univer's visible-range bookkeeping to the tile: put it back afterwards
    const drawing = sk._drawingRange;
    const vpKey = 'petakTile';
    const scale = comp.getParentScale;
    const B = BLEED;
    const bx = B / s; // the bleed in sheet px
    ctx.save();
    try {
      ctx.setTransform(1, 0, 0, 1, 0, 0);
      ctx.globalCompositeOperation = 'source-over';
      ctx.clearRect(0, 0, W, H);
      comp.getParentScale = () => ({ scaleX: level, scaleY: level });
      if (isCells(kind)) {
        const b: Bound = { left: r.sx - bx + rh, top: r.sy - bx + ch, right: r.sx + r.sw + bx + rh, bottom: r.sy + r.sh + bx + ch };
        ctx.setTransform(s, 0, 0, s, -(r.sx + rh) * s + B, -(r.sy + ch) * s + B);
        comp.render!(ctx, {
          viewBound: b,
          cacheBound: b,
          diffBounds: [],
          diffCacheBounds: [],
          diffX: 0,
          diffY: 0,
          viewPortPosition: { left: 0, top: 0, right: r.sw, bottom: r.sh },
          cacheViewPortPosition: b,
          viewportKey: vpKey,
          isDirty: 0,
          isForceDirty: true,
          shouldCacheUpdate: 0,
          leftOrigin: 0,
          topOrigin: 0,
          bufferEdgeX: 0,
          bufferEdgeY: 0,
        });
      } else if (isColHeader(kind)) {
        const b: Bound = { left: r.sx - bx + rh, top: 0, right: r.sx + r.sw + bx + rh, bottom: ch };
        ctx.setTransform(s, 0, 0, s, -(r.sx + rh) * s + B, B);
        comp.draw!(ctx, { viewBound: b, cacheBound: b, viewportKey: vpKey + 'H', diffBounds: [] });
      } else {
        const b: Bound = { left: 0, top: r.sy - bx + ch, right: rh, bottom: r.sy + r.sh + bx + ch };
        ctx.setTransform(s, 0, 0, s, B, -(r.sy + ch) * s + B);
        comp.draw!(ctx, { viewBound: b, cacheBound: b, viewportKey: vpKey + 'V', diffBounds: [] });
      }
    } catch (e) {
      console.warn('tile render failed', e);
    } finally {
      comp.getParentScale = scale;
      ctx.restore();
      sk._drawingRange = drawing;
      for (const m of [sk._visibleRangeMap, sk._cacheRangeMap]) for (const k of [vpKey, vpKey + 'H', vpKey + 'V']) m?.delete(k);
    }
    // the background goes *under* what was drawn (Univer leaves cleared, transparent areas)
    const { el, old } = this.slot(key);
    const bg = isCells(kind) ? this.bg : this.headerBg;
    present(el, W, H, (c) => {
      c.fillStyle = bg;
      c.fillRect(0, 0, W, H);
      c.drawImage(src, 0, 0);
    });
    const ew = sk.columnWidthAccumulation[sk.columnWidthAccumulation.length - 1] ?? 0;
    const eh = sk.rowHeightAccumulation[sk.rowHeightAccumulation.length - 1] ?? 0;
    const tile = this.store(key, kind, level, r, el, pw, ph, false, ew, eh, frame, old);
    this.rendered++;
    const ms = performance.now() - t0;
    const ck = kind[0] + level;
    const prev = this.costs.get(ck);
    this.costs.set(ck, prev === undefined ? ms : prev * 0.7 + ms * 0.3);
    return tile;
  }

  /**
   * A placeholder for a tile that is not drawn yet: background and plain grid lines (cells) or the
   * header background with its separators. A fraction of a millisecond.
   */
  renderPlain(key: string, kind: TileKind, level: number, r: TileRect, lines: { cols: number[]; rows: number[]; grid: boolean }, frame: number): Tile {
    const { pw, ph } = this.dims(level, r);
    const W = pw + 2 * BLEED;
    const H = ph + 2 * BLEED;
    const { el, old } = this.slot(key);
    const s = level * this.dpr;
    present(el, W, H, (ctx) => {
      ctx.fillStyle = isCells(kind) ? this.bg : this.headerBg;
      ctx.fillRect(0, 0, W, H);
      if (lines.grid || !isCells(kind)) {
        ctx.fillStyle = this.lineColor;
        const w = Math.max(1, Math.round(this.dpr / 2));
        if (!isRowHeader(kind)) for (const x of lines.cols) ctx.fillRect(Math.round((x - r.sx) * s) + BLEED - (w >> 1), 0, w, H);
        if (!isColHeader(kind)) for (const y of lines.rows) ctx.fillRect(0, Math.round((y - r.sy) * s) + BLEED - (w >> 1), W, w);
      }
    });
    return this.store(key, kind, level, r, el, pw, ph, true, 0, 0, frame, old);
  }

  /** Mark a cached tile as wanted in this frame (protects it from eviction). */
  touch(t: Tile, frame: number) {
    t.seen = frame;
  }

  /**
   * Keep the cache under its byte cap: evict tiles not shown in this frame, the ones farthest from
   * `center` (sheet px) and from the current zoom level first. Out-of-date tiles and placeholders go
   * before all others.
   */
  evict(frame: number, center: { x: number; y: number }, zoom: number) {
    if (this.bytes <= this.capBytes) return;
    const score = (t: Tile) => {
      if (t.ver !== this.ver) return -1e12;
      if (t.seen === frame) return 1e12;
      if (t.plain) return -1e11;
      const cx = t.sx + t.sw / 2;
      const cy = t.sy + t.sh / 2;
      const dist = Math.hypot(cx - center.x, cy - center.y) * zoom; // screen px away
      const levelGap = Math.abs(Math.log2(t.level / zoom));
      return -(dist + levelGap * 600 + (frame - t.seen) * 2);
    };
    const all = [...this.tiles.values()].sort((a, b) => score(a) - score(b));
    for (const t of all) {
      if (this.bytes <= this.capBytes * 0.85) break;
      if (t.seen === frame && t.ver === this.ver) break;
      this.tiles.delete(t.key);
      this.bytes -= t.bytes;
      this.drop(t);
    }
  }
}
