/**
 * Where the sheet is on screen. Maps between CSS pixels of the grid host and sheet
 * pixels (origin = top-left of A1 at 100 % zoom), for every pane: with frozen rows /
 * columns the grid has up to four panes that scroll differently.
 */
import { IRenderManagerService } from '@univerjs/engine-render';
import { SheetScrollManagerService } from '@univerjs/sheets-ui';
import { doc } from '../state/doc';

interface UViewport {
  viewportKey: string;
  left: number;
  top: number;
  width?: number;
  height?: number;
  viewportScrollX: number;
  viewportScrollY: number;
  isActive: boolean;
}
interface UScene {
  scaleX: number;
  scaleY: number;
  getViewport(key: string): UViewport | undefined;
  getViewports(): UViewport[];
  render(): void;
  makeDirty(state?: boolean): void;
}
export interface RenderUnit {
  scene: UScene;
  engine: { getCanvasElement(): HTMLCanvasElement; getPixelRatio?(): number; width: number; height: number };
}

export function renderUnit(): RenderUnit | null {
  try {
    const inj = (doc.api as unknown as { _injector: { get<X>(t: unknown): X } })._injector;
    const rm = inj.get<{ getRenderUnitById(id: string): RenderUnit | null }>(IRenderManagerService);
    return rm.getRenderUnitById(doc.wb.getId());
  } catch {
    return null;
  }
}

export interface Pane {
  key: 'main' | 'left' | 'top' | 'corner';
  /** screen rectangle (CSS px relative to the grid host) */
  x: number;
  y: number;
  w: number;
  h: number;
  /** sheet px shown at the pane's top-left corner */
  sx: number;
  sy: number;
}
export interface ViewInfo {
  zoom: number;
  /** screen = origin + sheet * zoom, for the main (scrolling) pane */
  ox: number;
  oy: number;
  /** header sizes on screen */
  headerW: number;
  headerH: number;
  /** total size of the grid canvas on screen */
  width: number;
  height: number;
  canvas: HTMLCanvasElement | null;
  rowAcc: number[];
  colAcc: number[];
  panes: Pane[];
  /** frozen rows / columns */
  frozenRows: number;
  frozenCols: number;
  /** frozen area on screen (after the headers) */
  frozenW: number;
  frozenH: number;
  /** current scroll of the main pane in sheet px */
  scrollX: number;
  scrollY: number;
}

const KEY: Record<string, Pane['key']> = { viewMain: 'main', viewMainLeft: 'left', viewMainTop: 'top', viewMainLeftTop: 'corner' };

export function viewInfo(): ViewInfo | null {
  try {
    const r = renderUnit();
    if (!r) return null;
    const main = r.scene.getViewport('viewMain');
    const ws = doc.wb.getActiveSheet();
    const sk = ws.getSkeleton();
    if (!main || !sk) return null;
    const zoom = r.scene.scaleX || 1;
    const canvas = r.engine.getCanvasElement();
    const width = canvas.clientWidth;
    const height = canvas.clientHeight;
    const headerW = sk.rowHeaderWidth * zoom;
    const headerH = sk.columnHeaderHeight * zoom;
    const fz = ws.getFreeze();
    const frozenCols = fz && fz.xSplit > 0 ? fz.xSplit : 0;
    const frozenRows = fz && fz.ySplit > 0 ? fz.ySplit : 0;
    const panes: Pane[] = [];
    for (const key of ['viewMain', 'viewMainLeft', 'viewMainTop', 'viewMainLeftTop']) {
      const v = r.scene.getViewport(key);
      if (!v || !v.isActive) continue;
      const w = v.width ?? width - v.left;
      const h = v.height ?? height - v.top;
      if (w <= 0 || h <= 0) continue;
      // Univer's scroll values are relative to the pane: the sheet position at the pane's corner is
      // its distance from the headers plus the scroll
      panes.push({ key: KEY[key], x: v.left, y: v.top, w, h, sx: (v.left - headerW) / zoom + v.viewportScrollX, sy: (v.top - headerH) / zoom + v.viewportScrollY });
    }
    return {
      zoom,
      ox: headerW - main.viewportScrollX * zoom,
      oy: headerH - main.viewportScrollY * zoom,
      headerW,
      headerH,
      width,
      height,
      canvas,
      rowAcc: sk.rowHeightAccumulation,
      colAcc: sk.columnWidthAccumulation,
      panes,
      frozenRows,
      frozenCols,
      frozenW: Math.max(0, main.left - headerW),
      frozenH: Math.max(0, main.top - headerH),
      scrollX: main.viewportScrollX,
      scrollY: main.viewportScrollY,
    };
  } catch {
    return null;
  }
}

/** Pane under a screen point (headers are not panes). */
export function paneAt(v: ViewInfo, x: number, y: number): Pane | null {
  // frozen panes sit on top of the main pane
  const order: Pane['key'][] = ['corner', 'top', 'left', 'main'];
  for (const k of order) {
    const p = v.panes.find((q) => q.key === k);
    if (p && x >= p.x && x < p.x + p.w && y >= p.y && y < p.y + p.h) return p;
  }
  return null;
}
/** Screen point -> sheet px (uses the pane under the point; the main pane outside of all). */
export function toSheet(v: ViewInfo, x: number, y: number): { x: number; y: number; pane: Pane | null } {
  const p = paneAt(v, x, y);
  if (!p) return { x: (x - v.ox) / v.zoom, y: (y - v.oy) / v.zoom, pane: null };
  return { x: (x - p.x) / v.zoom + p.sx, y: (y - p.y) / v.zoom + p.sy, pane: p };
}
/** Sheet px -> screen point, in the pane that shows that part of the sheet. */
export function toScreen(v: ViewInfo, sx: number, sy: number): { x: number; y: number } {
  // frozen columns / rows do not scroll: their screen position only depends on the zoom
  const fx = v.frozenW / v.zoom;
  const fy = v.frozenH / v.zoom;
  const left = v.panes.find((p) => p.key === 'left' || p.key === 'corner');
  const top = v.panes.find((p) => p.key === 'top' || p.key === 'corner');
  const x = left && sx < left.sx + fx ? v.headerW + (sx - left.sx) * v.zoom : v.ox + sx * v.zoom;
  const y = top && sy < top.sy + fy ? v.headerH + (sy - top.sy) * v.zoom : v.oy + sy * v.zoom;
  return { x, y };
}

export function indexAt(acc: number[], v: number): number {
  let lo = 0;
  let hi = acc.length - 1;
  while (lo < hi) {
    const m = (lo + hi) >> 1;
    if (acc[m] <= v) lo = m + 1;
    else hi = m;
  }
  return lo;
}
/** Cell under a screen point, or the header part that was hit. */
export function hitGrid(v: ViewInfo, x: number, y: number): { area: 'cell' | 'colHeader' | 'rowHeader' | 'corner'; row: number; col: number; sheetX: number; sheetY: number } {
  const inColHeader = y < v.headerH;
  const inRowHeader = x < v.headerW;
  // a header scrolls with the panes below / beside it
  const sxp = inRowHeader ? 0 : toSheet(v, x, Math.max(y, v.headerH + 0.5)).x;
  const syp = inColHeader ? 0 : toSheet(v, Math.max(x, v.headerW + 0.5), y).y;
  return {
    area: inColHeader && inRowHeader ? 'corner' : inColHeader ? 'colHeader' : inRowHeader ? 'rowHeader' : 'cell',
    row: inColHeader ? -1 : indexAt(v.rowAcc, syp),
    col: inRowHeader ? -1 : indexAt(v.colAcc, sxp),
    sheetX: sxp,
    sheetY: syp,
  };
}

/**
 * Scroll so that the main pane's top-left corner shows sheet position (sx, sy).
 * Univer takes the scroll as "row / column whose start is the scroll distance, plus an
 * offset"; with frozen panes the distance is measured from the end of the frozen area.
 */
export function scrollTo(v: ViewInfo, sx: number, sy: number) {
  const wb = doc.wb;
  const at = (acc: number[], i: number) => (i <= 0 ? 0 : acc[Math.min(i, acc.length) - 1] ?? 0);
  const x = Math.max(0, sx - v.frozenW / v.zoom);
  const y = Math.max(0, sy - v.frozenH / v.zoom);
  const c = Math.min(indexAt(v.colAcc, x), v.colAcc.length - 1);
  const r = Math.min(indexAt(v.rowAcc, y), v.rowAcc.length - 1);
  const offsetX = Math.max(0, x - at(v.colAcc, c));
  const offsetY = Math.max(0, y - at(v.rowAcc, r));
  doc.api.executeCommand('sheet.operation.set-scroll', {
    unitId: wb.getId(),
    sheetId: wb.getActiveSheet().getSheetId(),
    sheetViewStartRow: r,
    sheetViewStartColumn: c,
    offsetX,
    offsetY,
  });
  // Frozen columns wider than the screen (e.g. the roster at 200 % in portrait) leave Univer's
  // scrolling pane with no width: Univer then ignores every scroll, also the vertical one, and the
  // sheet cannot be scrolled at all (P-02). Excel still scrolls the rows: set the panes' scroll
  // the way Univer does it itself after a layout change, and keep it as Univer's scroll state.
  const unit = renderUnit() as unknown as { scene: { getViewport(k: string): (UViewport & { viewportScrollX: number; viewportScrollY: number }) | undefined }; with?: <T>(t: unknown) => T } | null;
  const main = unit?.scene.getViewport('viewMain');
  if (unit && main && !main.isActive) {
    const vx = at(v.colAcc, c) + offsetX;
    const vy = at(v.rowAcc, r) + offsetY;
    main.viewportScrollX = vx;
    main.viewportScrollY = vy;
    const set = (key: string, axis: 'x' | 'y') => {
      const vp = unit.scene.getViewport(key);
      if (vp) vp[axis === 'x' ? 'viewportScrollX' : 'viewportScrollY'] = axis === 'x' ? vx : vy;
    };
    set('viewRowBottom', 'y');
    set('viewMainLeft', 'y');
    set('viewColumnRight', 'x');
    set('viewMainTop', 'x');
    try {
      unit.with?.<{ setValidScrollStateToCurrSheet(s: object): void }>(SheetScrollManagerService).setValidScrollStateToCurrSheet({
        sheetViewStartRow: r,
        sheetViewStartColumn: c,
        offsetX,
        offsetY,
        viewportScrollX: vx,
        viewportScrollY: vy,
        scrollX: 0,
        scrollY: 0,
      });
    } catch (e) {
      console.warn('scroll state', e);
    }
    (unit.scene as unknown as UScene).makeDirty(true);
  }
}
/**
 * Thin lines on a stretched canvas (K-08). On Low / Mid tiers the grid canvas has fewer pixels
 * than the screen and is stretched by a non-whole factor (e.g. 2.625 / 2). A border that is one
 * canvas pixel wide is then smeared to a faint grey line, or seems to vanish between coloured
 * cells. Borders and grid lines are therefore drawn one *screen* pixel wide: the canvas line
 * width is multiplied by the stretch. Idempotent; call when a render unit exists.
 */
export function fixThinLines(stretch: number) {
  if (stretch <= 1.02) return;
  try {
    const ctx = (renderUnit() as unknown as { engine: { getCanvas(): { getContext(): object } } } | null)?.engine.getCanvas().getContext();
    const proto = ctx && (Object.getPrototypeOf(ctx) as { setLineWidthByPrecision?: (w: number) => void; __petakLines?: boolean; __petakOrig?: (w: number) => void });
    if (!proto?.setLineWidthByPrecision || proto.__petakLines) return;
    const orig = proto.setLineWidthByPrecision;
    proto.setLineWidthByPrecision = function (this: unknown, w: number) {
      orig.call(this, w * stretch);
    };
    proto.__petakLines = true;
    proto.__petakOrig = orig; // tests switch the fix off to prove they see the difference
  } catch (e) {
    console.warn('line fix not applied', e);
  }
}
interface ULoopScene {
  isDirty(): boolean;
  getLayers(): object[];
}
interface ULoopEngine {
  _renderFunction: (t: number) => void;
  _renderingQueueLaunched: boolean;
  _renderFrameTasks: unknown[];
  _scenes?: Record<string, ULoopScene>;
  startRenderLoop(): void;
  __petakQuiet?: { paused: boolean; idle: number };
}
let quietTimer = 0;
const watching = new WeakSet<object>();
const quietEngines = new Set<ULoopEngine>();
function wake(e: ULoopEngine | null | undefined) {
  const q = e?.__petakQuiet;
  if (!q?.paused) return;
  q.paused = false;
  q.idle = 0;
  e!.startRenderLoop();
}

/**
 * Univer draws the grid from a requestAnimationFrame loop that never stops: 60 (on a 120 Hz screen
 * 120) callbacks a second while nothing changes. Each one makes the browser commit a frame, and every
 * commit copies the grid's canvas (measured: 60 commits and 60 canvas copies a second with the app
 * idle, 90 ms of main-thread time per second here): CPU time and battery for nothing. This stops
 * an engine's loop after a few frames with nothing to draw and starts it again the moment anything
 * on its scenes is marked dirty (Univer marks its layers dirty for every change it wants drawn).
 */
export function quietRenderLoop() {
  try {
    // every render unit: the workbook's, and the documents of Univer's cell editors
    const inj = (doc.api as unknown as { _injector: { get<X>(t: unknown): X } })._injector;
    const rm = inj.get<{ getRenderAll(): Map<string, unknown>; created$?: { subscribe(f: () => void): unknown } }>(IRenderManagerService);
    for (const unit of rm.getRenderAll().values()) quiet(unit as { engine: ULoopEngine; scene: ULoopScene });
    // render units made later (a new workbook's grid appears a moment after the workbook)
    if (!watching.has(rm) && rm.created$) {
      watching.add(rm);
      rm.created$.subscribe(() => setTimeout(quietRenderLoop, 0));
    }
  } catch (err) {
    console.warn('render loop not quietened', err);
  }
}
function quiet(unit: { engine: ULoopEngine; scene: ULoopScene } | null) {
  try {
    const e = unit?.engine;
    if (!unit || !e || e.__petakQuiet || typeof e._renderFunction !== 'function') return;
    const layer = unit.scene.getLayers()[0] as { makeDirty?: (s?: boolean) => unknown } | undefined;
    const proto = layer && (Object.getPrototypeOf(layer) as { makeDirty?: (s?: boolean) => unknown; __petakWake?: boolean });
    if (!proto?.makeDirty) return;
    if (!proto.__petakWake) {
      const md = proto.makeDirty;
      proto.makeDirty = function (this: { scene?: { getEngine?(): ULoopEngine | null } }, state = true) {
        const r = md.call(this, state);
        if (state !== false) wake(this.scene?.getEngine?.());
        return r;
      };
      proto.__petakWake = true;
    }
    const q: { paused: boolean; idle: number } = { paused: false, idle: 0 };
    e.__petakQuiet = q;
    const run = e._renderFunction;
    const dirty = () => Object.values(e._scenes ?? {}).some((sc) => sc.isDirty());
    e._renderFunction = (t: number) => {
      if (dirty()) q.idle = 0;
      else if (++q.idle > 12 && e._renderFrameTasks.length) {
        e._renderingQueueLaunched = false; // Univer's own "loop not running" state
        q.paused = true;
        return;
      }
      run(t);
    };
    quietEngines.add(e);
    // a safety net for anything that makes a scene dirty another way: look 4 times a second
    if (!quietTimer)
      quietTimer = window.setInterval(() => {
        for (const en of quietEngines) {
          if (!en._scenes) quietEngines.delete(en);
          else if (en.__petakQuiet?.paused && Object.values(en._scenes).some((sc) => sc.isDirty())) wake(en);
        }
      }, 250);
  } catch (err) {
    console.warn('render loop not quietened', err);
  }
}

/**
 * Set Univer's zoom. Resolves false instead of throwing when Univer is not ready for it yet: its
 * zoom controller is registered a moment after a workbook appears, and a zoom sent before that
 * (a pinch right after a file opened) failed with an uncaught error (P-03).
 */
export function setUniverZoom(z: number): Promise<boolean> {
  try {
    const wb = doc.wb;
    return Promise.resolve(doc.api.executeCommand('sheet.operation.set-zoom-ratio', { zoomRatio: z, unitId: wb.getId(), subUnitId: wb.getActiveSheet().getSheetId() }))
      .then((ok) => !!ok)
      .catch(() => false);
  } catch {
    return Promise.resolve(false);
  }
}
/** Sheet position at the main pane's top-left corner. */
export const mainOrigin = (v: ViewInfo) => ({ x: v.frozenW / v.zoom + v.scrollX, y: v.frozenH / v.zoom + v.scrollY });

(globalThis as unknown as { __view: unknown }).__view = { viewInfo, renderUnit, hitGrid, toSheet, toScreen, scrollTo, mainOrigin };
