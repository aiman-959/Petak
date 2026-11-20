/**
 * The grid while a finger moves it: pan, fling, pinch zoom, double-tap zoom and animated zoom.
 *
 * Why this exists (AUDIT.md, P-01). Univer redraws the whole sheet for every scroll or zoom step
 * (a zoom also lays the sheet out again), far too slow for 60–120 frames per second on a phone.
 * Petak 1.1–1.2 therefore moved *copies* of the grid with CSS transforms. That design broke on
 * pinch: the copies were opaque canvases that start black, so every empty (transparent) cell
 * turned black; the whole copy, headers included, was scaled as one picture, so on zoom-out the
 * live grid showed around it (a second grid with its own headers) and on zoom-in it grew over the
 * app's own bars; and areas the copies did not cover could not be filled until the fingers lifted.
 *
 * Now, for the length of a gesture, a surface (`.gesture-surface`) covers the grid area. It is
 * built like a map app: every pane of the grid (the scrolling cells, the frozen rows, columns and
 * corner, the two parts of the column headers and of the row headers) is a clipping box, and in it
 * one layer per zoom level holds that level's tiles (tiles.ts: small canvases drawn by Univer's own
 * renderer). One camera (zoom + scroll) places every pane and layer each frame, so headers, frozen
 * panes, cells, shapes and pictures move and scale together around the fingers and only one grid
 * is ever visible. Moving and scaling is done by the GPU compositor (CSS transforms): Petak's work
 * per frame is a few style writes, whatever the zoom.
 *
 * Where a tile of the best level is missing, the nearest cached level shows instead (map style),
 * or the grid's own picture from the start of the gesture, or a placeholder with plain grid lines
 * (drawn in a fraction of a millisecond); missing tiles are drawn in the time left over in each
 * frame, and all of them after the gesture. Nothing is opaque-black: panes have the sheet colour.
 * When the fingers lift, Univer is put at the final zoom and position, renders once sharply, and
 * only then is the surface hidden (the hand-off). No React state, no layout of the app and no text
 * measuring happen during a gesture.
 */
import { doc } from '../state/doc';
import { getTier } from '../device/tier';
import { mainOrigin, renderUnit, scrollTo, setUniverZoom, viewInfo, type ViewInfo } from './view';
import { BLEED, LEVELS, TileStore, isCells, present, release, sheetParts, span, type SheetParts, type Tile, type TileKind, type TileRect } from './tiles';
import { anchorAt, badness, cameraFor, clampCamera, fallbackLevels, targetLevel, type Camera, type Frame } from './camera';

export { MIN_ZOOM, MAX_ZOOM, type Camera } from './camera';

interface Geo extends Frame {
  /** sheet start of the frozen columns / rows */
  fx0: number;
  fy0: number;
  colAcc: number[];
  rowAcc: number[];
  gridlines: boolean;
  lineColor: string;
}
type PaneId = 'main' | 'top' | 'left' | 'corner' | 'hmain' | 'hfrozen' | 'vmain' | 'vfrozen';
const KIND: Record<PaneId, TileKind> = { main: 'c', top: 'ct', left: 'cl', corner: 'cc', hmain: 'h', hfrozen: 'hf', vmain: 'v', vfrozen: 'vf' };
/** which of a kind's axes are cut to the frozen band */
const BAND: Record<TileKind, { x: 'band' | 'grid' | 'header'; y: 'band' | 'grid' | 'header' }> = {
  c: { x: 'grid', y: 'grid' },
  ct: { x: 'grid', y: 'band' },
  cl: { x: 'band', y: 'grid' },
  cc: { x: 'band', y: 'band' },
  h: { x: 'grid', y: 'header' },
  hf: { x: 'band', y: 'header' },
  v: { x: 'header', y: 'grid' },
  vf: { x: 'header', y: 'band' },
};
interface Pane {
  id: PaneId;
  kind: TileKind;
  /** css px on the surface */
  x: number;
  y: number;
  w: number;
  h: number;
  /** sheet px at the pane's top-left corner */
  ox: number;
  oy: number;
}
interface PaneEl {
  div: HTMLDivElement;
  layers: Map<string, { div: HTMLDivElement; bx: number; by: number; level: number; css: string }>;
  shown: Set<HTMLCanvasElement>;
  snap: HTMLCanvasElement | null;
  sel: HTMLDivElement;
  css: string;
}
export interface ViewStats {
  frames: number;
  /** ms of Petak's own work in the last frame (placing layers + drawing tiles) */
  frameMs: number;
  frameMsMax: number;
  fps: number;
  pending: number;
  tiles: number;
  tileBytes: number;
  tileCap: number;
  tilesRendered: number;
  level: number;
  zoom: number;
  handoffMs: number;
  /** Univer redraws during the last gesture (should be 0) */
  renders: number;
  /** tile canvases in the DOM at the moment, and the largest canvas (device px) */
  layersShown: number;
}

const ease = (t: number) => 1 - Math.pow(1 - t, 3);
const px = (v: number) => Math.round(v * 100) / 100;

export class GestureView {
  private wrap: HTMLElement;
  private root: HTMLDivElement;
  private panesEl = new Map<PaneId, PaneEl>();
  private chromeEls: Record<'corner' | 'hline' | 'vline' | 'freezeV' | 'freezeH', HTMLDivElement>;
  private snapCam: Camera | null = null;
  private snapK = 1;
  private pill: HTMLDivElement;
  private overlay: HTMLPreElement | null = null;
  private barV: HTMLDivElement;
  private barH: HTMLDivElement;
  private tiles: TileStore;
  private parts: SheetParts | null = null;
  private g: Geo | null = null;
  private cam: Camera = { z: 1, x: 0, y: 0 };
  private restZoom = 1;
  private dpr = 1;
  private raf = 0;
  private frameNo = 0;
  private queue = new Map<string, { kind: TileKind; level: number; r: TileRect; pri: number }>();
  private anims: Array<(t: number) => boolean> = [];
  private moving = 0; // frames since the camera last moved
  private lastCam = '';
  private lastT = 0;
  /** ms per frame for drawing tiles while the picture moves (adapts to late frames) */
  private jobBudget = 4;
  private handing: { t0: number; z: number; x: number; y: number; stable: number; tries: number } | null = null;
  private idleHandle = 0;
  private pillTimer: ReturnType<typeof setTimeout> | undefined;
  private fpsT = performance.now();
  private fpsN = 0;
  private overlayRaf = 0;
  /** set by the input layer: called with true when a session starts, false after the hand-off */
  onActive: (active: boolean) => void = () => {};
  /** Opens the zoom controls (set by the app). */
  onPillTap: () => void = () => {};
  /** the surface is shown and Petak draws the grid */
  active = false;
  /** Univer's canvas while the surface covers it */
  private hidden: HTMLCanvasElement | null = null;
  stats: ViewStats = { frames: 0, frameMs: 0, frameMsMax: 0, fps: 0, pending: 0, tiles: 0, tileBytes: 0, tileCap: 0, tilesRendered: 0, level: 1, zoom: 1, handoffMs: 0, renders: 0, layersShown: 0 };

  constructor(host: HTMLElement) {
    this.wrap = host.parentElement ?? host;
    this.root = document.createElement('div');
    this.root.className = 'gesture-surface';
    this.root.setAttribute('aria-hidden', 'true');
    for (const id of ['corner', 'left', 'top', 'main', 'vfrozen', 'vmain', 'hfrozen', 'hmain'] as PaneId[]) {
      const div = document.createElement('div');
      div.className = 'gs-pane ' + (isCells(KIND[id]) ? 'cells' : 'header');
      div.dataset.pane = id;
      const sel = document.createElement('div');
      sel.className = isCells(KIND[id]) ? 'gs-sel' : 'gs-hsel';
      div.appendChild(sel);
      this.root.appendChild(div);
      this.panesEl.set(id, { div, layers: new Map(), shown: new Set(), snap: null, sel, css: '' });
    }
    const mk = (cls: string) => {
      const d = document.createElement('div');
      d.className = cls;
      this.root.appendChild(d);
      return d;
    };
    this.chromeEls = { corner: mk('gs-corner'), hline: mk('gs-line h'), vline: mk('gs-line v'), freezeV: mk('gs-freeze v'), freezeH: mk('gs-freeze h') };
    this.pill = document.createElement('div');
    this.pill.className = 'zoom-pill';
    this.pill.dataset.testid = 'zoom-pill';
    this.pill.setAttribute('role', 'button');
    this.pill.setAttribute('aria-label', 'Zoom');
    this.pill.addEventListener('click', () => this.onPillTap());
    this.barV = document.createElement('div');
    this.barH = document.createElement('div');
    this.barV.className = 'scroll-ind v';
    this.barH.className = 'scroll-ind h';
    this.wrap.append(this.root, this.barV, this.barH, this.pill);
    const tier = (() => {
      try {
        return getTier().tier;
      } catch {
        return 'mid';
      }
    })();
    // the tile cache's memory cap (each tile is about 1 MB of canvas)
    this.tiles = new TileStore((tier === 'high' ? 96 : tier === 'mid' ? 56 : 28) * 1048576);
    this.tiles.onEvict = (t) => {
      t.el.remove();
      for (const pe of this.panesEl.values()) pe.shown.delete(t.el);
      t.placed = undefined;
    };
    this.stats.tileCap = this.tiles.capBytes;
    new ResizeObserver(() => {
      if (this.active && !this.handing) this.end();
    }).observe(host);
    document.fonts?.addEventListener?.('loadingdone', () => this.invalidate());
    (globalThis as unknown as { __petakView: unknown }).__petakView = this;
  }

  /** Univer's canvas while the surface covers it (hidden: touches land on its container) */
  get hiddenCanvas(): HTMLCanvasElement | null {
    return this.hidden;
  }

  get camera(): Camera {
    return this.cam;
  }
  get zoom() {
    return this.active ? this.cam.z : (viewInfo()?.zoom ?? 1);
  }

  // ------------------------------------------------------------------ session
  /** Start drawing the grid from the cache. Returns false while there is no grid. */
  begin(): boolean {
    if (this.active) {
      // a new finger during the hand-off: keep going from where the picture is
      this.handing = null;
      return true;
    }
    const v = viewInfo();
    const parts = sheetParts();
    if (!v || !v.canvas || !parts || v.width < 4 || v.height < 4) return false;
    this.parts = parts;
    this.dpr = window.devicePixelRatio || 1;
    this.tiles.checkSheet(parts, this.dpr);
    this.colours(parts);
    this.quietScrollBars();
    this.g = this.geometry(v, parts);
    const o = mainOrigin(v);
    this.cam = { z: v.zoom, x: o.x, y: o.y };
    this.restZoom = v.zoom;
    this.root.style.width = v.width + 'px';
    this.root.style.height = v.height + 'px';
    this.takeSnapshot(v);
    this.stats = { ...this.stats, frames: 0, frameMsMax: 0, renders: 0, tilesRendered: this.tiles.rendered };
    this.queue.clear();
    this.anims = [];
    this.moving = 0;
    this.selCache = null;
    this.draw(); // the first frame is complete before the surface is shown
    this.root.style.display = 'block';
    // the surface is opaque and covers Univer's canvas: hidden, the browser does not copy its
    // pixels at every frame (see tiles.ts); shown again in the same frame the surface goes
    this.hidden = v.canvas;
    v.canvas.style.visibility = 'hidden';
    this.wrap.classList.add('gesturing');
    this.active = true;
    this.onActive(true);
    this.showPill();
    this.request();
    return true;
  }

  private colours(parts: SheetParts) {
    this.tiles.bg = getComputedStyle(this.wrap).backgroundColor || '#fff';
    const dark = document.documentElement.dataset.theme === 'dark';
    this.tiles.headerBg = dark ? '#1f2a24' : '#f6f8f7';
    this.tiles.lineColor = parts.sk.gridlinesColor ?? parts.sk.defaultGridlinesColor ?? '#e0e3e1';
    this.root.style.setProperty('--gs-bg', this.tiles.bg);
    this.root.style.setProperty('--gs-hbg', this.tiles.headerBg);
    this.root.style.setProperty('--gs-line', this.tiles.lineColor);
  }

  private geometry(v: ViewInfo, parts: SheetParts): Geo {
    const z = v.zoom;
    const left = v.panes.find((p) => p.key === 'left' || p.key === 'corner');
    const top = v.panes.find((p) => p.key === 'top' || p.key === 'corner');
    const o = mainOrigin(v);
    const fw = v.frozenW / z;
    const fh = v.frozenH / z;
    const totalW = v.colAcc[v.colAcc.length - 1] ?? 0;
    const totalH = v.rowAcc[v.rowAcc.length - 1] ?? 0;
    const sk = parts.sk;
    return {
      rh: parts.rh,
      ch: parts.ch,
      fw,
      fh,
      fx0: left ? left.sx : 0,
      fy0: top ? top.sy : 0,
      colAcc: v.colAcc,
      rowAcc: v.rowAcc,
      minX: o.x - v.scrollX,
      minY: o.y - v.scrollY,
      maxX: Math.max(o.x - v.scrollX, totalW),
      maxY: Math.max(o.y - v.scrollY, totalH),
      W: v.width,
      H: v.height,
      gridlines: sk.showGridlines !== 0,
      lineColor: sk.gridlinesColor ?? sk.defaultGridlinesColor ?? '#e0e3e1',
    };
  }

  /** Univer paints its scroll bars into the canvas; Petak shows thin indicators instead. */
  private quietScrollBars() {
    const unit = renderUnit();
    if (!unit) return;
    // count Univer's real redraws during a gesture (its render loop calls render() every frame;
    // render() returns at once unless something is dirty)
    const scene = unit.scene as unknown as { render(): void; isDirty(): boolean; __petakCount?: boolean };
    if (!scene.__petakCount) {
      scene.__petakCount = true;
      const render = scene.render.bind(scene);
      scene.render = () => {
        if (this.active && !this.handing && scene.isDirty()) this.stats.renders++;
        render();
      };
    }
    let repaint = false;
    for (const key of ['viewMain', 'viewMainLeft', 'viewMainTop']) {
      const sb = (unit.scene.getViewport(key) as unknown as { getScrollBar?: () => { render: unknown; __quiet?: boolean } | null } | undefined)?.getScrollBar?.();
      if (sb && !sb.__quiet) {
        sb.__quiet = true;
        sb.render = () => {};
        repaint = true;
      }
    }
    if (repaint) {
      unit.scene.makeDirty(true);
      unit.scene.render();
    }
  }

  /**
   * The grid as it is now, cut into its panes: exact at the starting zoom (it also shows Univer's
   * selection handles and filter buttons), used like one more tile level.
   */
  private takeSnapshot(v: ViewInfo) {
    const src = v.canvas!;
    this.snapK = src.width / Math.max(1, src.clientWidth || v.width);
    this.snapCam = { ...this.cam };
    const k = this.snapK;
    for (const p of this.panes()) {
      const pe = this.panesEl.get(p.id)!;
      const w = Math.max(1, Math.round(p.w * k));
      const h = Math.max(1, Math.round(p.h * k));
      let c = pe.snap;
      if (!c) {
        c = document.createElement('canvas');
        c.className = 'gs-snap';
        pe.snap = c;
      }
      // a bitmap, not a 2D canvas: see tiles.ts (a 2D canvas on screen is copied at every frame)
      const bg = isCells(p.kind) ? this.tiles.bg : this.tiles.headerBg;
      const sx = Math.round(p.x * k);
      const sy = Math.round(p.y * k);
      present(c, w, h, (ctx) => {
        ctx.fillStyle = bg;
        ctx.fillRect(0, 0, w, h);
        ctx.drawImage(src, sx, sy, w, h, 0, 0, w, h);
      });
      (c as HTMLCanvasElement & { __pane?: Pane }).__pane = p;
    }
  }

  // ------------------------------------------------------------------ camera
  /** The sheet point under a surface point (css px), in the scrolling pane's coordinates (also over frozen panes and headers). */
  anchorAt(px0: number, py0: number, c: Camera = this.cam) {
    return anchorAt(this.g!, c, px0, py0);
  }
  /** Camera that puts sheet point `a` under surface point (px, py) at zoom z (clamped to the sheet). */
  cameraFor(a: { x: number; y: number }, px0: number, py0: number, z: number): Camera {
    return cameraFor(this.g!, a, px0, py0, z);
  }
  setCamera(c: Camera) {
    if (!this.active) return;
    this.cam = clampCamera(this.g!, c);
    this.request();
  }
  /** at a limit on that axis (a fling stops there) */
  atLimit(axis: 'x' | 'y') {
    const c = this.cam;
    const d = clampCamera(this.g!, { z: c.z, x: axis === 'x' ? -1e12 : c.x, y: axis === 'y' ? -1e12 : c.y });
    const e = clampCamera(this.g!, { z: c.z, x: axis === 'x' ? 1e12 : c.x, y: axis === 'y' ? 1e12 : c.y });
    return c[axis] <= d[axis] + 0.01 || c[axis] >= e[axis] - 0.01;
  }
  /** Run `step` every frame until it returns false (fling, zoom animation). */
  animate(step: (t: number) => boolean) {
    this.anims.push(step);
    this.request();
  }
  stopAnimations() {
    this.anims = [];
  }
  get animating() {
    return this.anims.length > 0;
  }

  /** Zoom to z keeping surface point (fx, fy) still, animated; ends with the hand-off. */
  zoomTo(z: number, fx?: number, fy?: number, ms = 240): boolean {
    if (!this.begin()) return false;
    const g = this.g!;
    const from = { ...this.cam };
    const px0 = fx ?? (g.rh + g.fw) * from.z + (g.W - (g.rh + g.fw) * from.z) / 2;
    const py0 = fy ?? (g.ch + g.fh) * from.z + (g.H - (g.ch + g.fh) * from.z) / 2;
    const a = this.anchorAt(px0, py0, from);
    const t0 = performance.now();
    const motion = reducedMotion() ? 0 : ms;
    this.animate((t) => {
      const f = motion ? Math.min(1, (t - t0) / motion) : 1;
      // zoom changes geometrically so the speed looks even
      const zz = from.z * Math.pow(z / from.z, ease(f));
      this.setCamera(this.cameraFor(a, px0, py0, zz));
      if (f >= 1) {
        this.end();
        return false;
      }
      return true;
    });
    return true;
  }

  /** Zoom slider: show zoom z around the middle of the sheet while the slider moves. */
  private sliderAnchor: { a: { x: number; y: number }; px: number; py: number } | null = null;
  previewZoom(z: number) {
    if (!this.begin()) return;
    const g = this.g!;
    if (!this.sliderAnchor) {
      const c = this.cam;
      const px0 = (g.rh + g.fw) * c.z + (g.W - (g.rh + g.fw) * c.z) / 2;
      const py0 = (g.ch + g.fh) * c.z + (g.H - (g.ch + g.fh) * c.z) / 2;
      this.sliderAnchor = { a: this.anchorAt(px0, py0), px: px0, py: py0 };
    }
    const s = this.sliderAnchor;
    this.setCamera(this.cameraFor(s.a, s.px, s.py, z));
  }
  /** The slider was released. */
  commitZoom() {
    this.sliderAnchor = null;
    this.end();
  }

  // ------------------------------------------------------------------ frames
  private request() {
    if (!this.raf) this.raf = requestAnimationFrame(this.frame);
  }
  private frame = (t: number) => {
    this.raf = 0;
    if (!this.active) return;
    try {
      this.frameBody(t);
    } catch (e) {
      // never leave the grid covered (or Univer's canvas hidden) after an error
      console.warn('gesture frame failed', e);
      this.finish();
    }
  };
  private frameBody(t: number) {
    const t0 = performance.now();
    if (this.anims.length) this.anims = this.anims.filter((a) => a(t));
    if (!this.active) return; // an animation ended the session
    const key = `${this.cam.z}|${this.cam.x}|${this.cam.y}`;
    this.moving = key === this.lastCam ? this.moving + 1 : 0;
    this.lastCam = key;
    const hz = Math.max(60, Math.min(144, getTierRefresh()));
    const vsync = 1000 / hz;
    // late frames (the device is slower than the estimate) halve the time given to tiles while the
    // picture moves; on time, it grows back
    const dt = this.lastT ? t - this.lastT : vsync;
    this.lastT = t;
    if (dt > vsync * 1.45) this.jobBudget = Math.max(0, this.jobBudget * 0.5 - 0.5);
    else this.jobBudget = Math.min(vsync * 0.5, this.jobBudget + 0.25);
    this.draw();
    this.handoffStep();
    if (!this.active) return;
    // draw missing tiles with the time left in this frame; when the picture rests (or the fingers
    // are up) more: a long frame is not seen then
    // during the hand-off Univer is about to show the sharp grid: no more tiles then
    const still = this.moving >= 2;
    const rendered = this.handing ? 0 : this.renderJobs(t0, still ? 2 * vsync : this.jobBudget);
    const ms = performance.now() - t0;
    this.stats.frameMs = ms;
    this.stats.frameMsMax = Math.max(this.stats.frameMsMax, ms);
    this.stats.frames++;
    this.fpsN++;
    if (t - this.fpsT >= 500) {
      this.stats.fps = Math.round((this.fpsN * 1000) / (t - this.fpsT));
      this.fpsT = t;
      this.fpsN = 0;
    }
    this.updateStats();
    if (this.anims.length || this.queue.size || this.handing || this.moving < 2 || rendered) this.request();
  }

  private renderJobs(t0: number, budget: number): number {
    if (!this.queue.size || !this.parts) return 0;
    const jobs = [...this.queue.entries()].sort((a, b) => a[1].pri - b[1].pri);
    let n = 0;
    for (const [key, j] of jobs) {
      const est = this.tiles.cost(j.kind, j.level);
      // header strips are cheap and keep the labels in place: one per frame is always allowed
      if (performance.now() - t0 + est > budget && (n > 0 || isCells(j.kind))) break;
      this.queue.delete(key);
      const t = this.tiles.render(this.parts, key, j.kind, j.level, j.r, this.frameNo);
      if (t) this.request();
      n++;
    }
    this.tiles.evict(this.frameNo, this.center(), this.cam.z);
    return n;
  }

  private center() {
    const g = this.g!;
    const c = this.cam;
    return { x: c.x + (g.W / c.z - g.rh - g.fw) / 2, y: c.y + (g.H / c.z - g.ch - g.fh) / 2 };
  }

  private targetLevel(z: number) {
    return targetLevel(z, this.restZoom);
  }

  private panes(c: Camera = this.cam): Pane[] {
    const g = this.g!;
    const { z, x, y } = c;
    const hw = g.rh * z;
    const hh = g.ch * z;
    const fwz = g.fw * z;
    const fhz = g.fh * z;
    const out: Pane[] = [];
    const add = (id: PaneId, px0: number, py0: number, w: number, h: number, ox: number, oy: number) => out.push({ id, kind: KIND[id], x: px0, y: py0, w: Math.max(0, w), h: Math.max(0, h), ox, oy });
    add('main', hw + fwz, hh + fhz, g.W - hw - fwz, g.H - hh - fhz, x, y);
    add('top', hw + fwz, hh, g.W - hw - fwz, fhz, x, g.fy0);
    add('left', hw, hh + fhz, fwz, g.H - hh - fhz, g.fx0, y);
    add('corner', hw, hh, fwz, fhz, g.fx0, g.fy0);
    add('hmain', hw + fwz, 0, g.W - hw - fwz, hh, x, 0);
    add('hfrozen', hw, 0, fwz, hh, g.fx0, 0);
    add('vmain', 0, hh + fhz, hw, g.H - hh - fhz, 0, y);
    add('vfrozen', 0, hh, hw, fhz, 0, g.fy0);
    return out;
  }

  // ----------------------------------------------------------- tile geometry
  /** Index range along one axis of a kind's tile grid for the sheet span [a, b). */
  private axis(mode: 'band' | 'grid' | 'header', S: number, a: number, b: number, band0: number, bandLen: number, ext: number) {
    if (mode === 'header') return { i0: 0, i1: 0 };
    if (mode === 'band') {
      const lo = Math.max(a, band0);
      const hi = Math.min(b, band0 + bandLen, ext);
      if (hi <= lo) return { i0: 0, i1: -1 };
      return { i0: Math.floor((lo - band0) / S), i1: Math.floor((hi - band0 - 1e-6) / S) };
    }
    const hi = Math.min(b, ext);
    if (hi <= a) return { i0: 0, i1: -1 };
    return { i0: Math.max(0, Math.floor(a / S)), i1: Math.floor((hi - 1e-6) / S) };
  }
  private rectOf(kind: TileKind, L: number, i: number, j: number): TileRect {
    const g = this.g!;
    const S = span(L, this.dpr);
    const b = BAND[kind];
    const one = (mode: 'band' | 'grid' | 'header', k: number, band0: number, bandLen: number, header: number) => {
      if (mode === 'header') return [0, header];
      if (mode === 'band') {
        const s0 = band0 + k * S;
        return [s0, Math.max(0, Math.min(S, band0 + bandLen - s0))];
      }
      return [k * S, S];
    };
    const [sx, sw] = one(b.x, i, g.fx0, g.fw, g.rh);
    const [sy, sh] = one(b.y, j, g.fy0, g.fh, g.ch);
    return { sx, sy, sw, sh };
  }
  private key(kind: TileKind, L: number, i: number, j: number) {
    const g = this.g!;
    const band = kind === 'c' || kind === 'h' || kind === 'v' ? '' : `|${g.fx0},${g.fw},${g.fy0},${g.fh}`;
    return `${kind}|${L}|${i}|${j}${band}`;
  }
  /** Tiles of level L of a kind that cover the sheet span of pane p (index ranges). */
  private range(p: Pane, L: number) {
    const g = this.g!;
    const S = span(L, this.dpr);
    const b = BAND[p.kind];
    const z = this.cam.z;
    const extW = g.colAcc[g.colAcc.length - 1] ?? 0;
    const extH = g.rowAcc[g.rowAcc.length - 1] ?? 0;
    const x = this.axis(b.x, S, p.ox, p.ox + p.w / z, g.fx0, g.fw, extW);
    const y = this.axis(b.y, S, p.oy, p.oy + p.h / z, g.fy0, g.fh, extH);
    return { x, y };
  }

  // ------------------------------------------------------------------ drawing
  private draw() {
    const g = this.g;
    if (!g) return;
    this.frameNo++;
    const z = this.cam.z;
    const level = this.targetLevel(z);
    this.stats.level = level;
    this.stats.zoom = z;
    // where a tile of the best level is missing: the best cached source that covers it
    const cands: Array<number | 'snap'> = fallbackLevels(z, level, [this.restZoom]).reverse();
    if (this.snapCam) {
      const b = badness(this.snapCam.z, z);
      const i = cands.findIndex((c) => c !== 'snap' && badness(c, z) > b);
      cands.splice(i < 0 ? cands.length : i, 0, 'snap');
    }
    let plainBudget = this.handing ? 0 : 2.5; // ms of placeholders per frame
    let shownTotal = 0;
    for (const p of this.panes()) {
      const pe = this.panesEl.get(p.id)!;
      const css = `transform:translate3d(${px(p.x)}px,${px(p.y)}px,0);width:${px(p.w)}px;height:${px(p.h)}px;display:${p.w >= 0.5 && p.h >= 0.5 ? 'block' : 'none'}`;
      if (css !== pe.css) {
        pe.div.style.cssText = css;
        pe.css = css;
      }
      const shown = new Set<HTMLCanvasElement>();
      const used = new Map<string, number>(); // layer key → rank (higher on top)
      if (p.w >= 0.5 && p.h >= 0.5) {
        const exactSnap = !!this.snapCam && Math.abs(this.snapCam.z - z) < 1e-6;
        const { x, y } = this.range(p, level);
        const cx = p.ox + p.w / z / 2;
        const cy = p.oy + p.h / z / 2;
        for (let j = y.i0; j <= y.i1; j++) {
          for (let i = x.i0; i <= x.i1; i++) {
            const r = this.rectOf(p.kind, level, i, j);
            if (r.sw <= 0 || r.sh <= 0) continue;
            const key = this.key(p.kind, level, i, j);
            const t = this.tiles.get(key);
            const real = t && !t.plain;
            const stale = real && (t!.ew < Math.min(r.sx + r.sw, g.colAcc[g.colAcc.length - 1] ?? 0) - 0.5 || t!.eh < Math.min(r.sy + r.sh, g.rowAcc[g.rowAcc.length - 1] ?? 0) - 0.5);
            if (!real || stale) {
              const pri = (isCells(p.kind) ? 1e6 : 0) + (stale ? 2e6 : 0) + Math.hypot(r.sx + r.sw / 2 - cx, r.sy + r.sh / 2 - cy) * z;
              const q = this.queue.get(key);
              if (!q || q.pri > pri) this.queue.set(key, { kind: p.kind, level, r, pri });
            }
            if (real) {
              this.show(p, t!, shown, used, 1000);
              continue;
            }
            // missing: the snapshot (exact at the start), a cached level that covers it, else a placeholder
            let covered = false;
            for (const c of cands) {
              if (c === 'snap') {
                if (this.snapCovers(p, r)) {
                  used.set('snap', exactSnap ? 2000 : 500 - badness(this.snapCam!.z, z));
                  covered = true;
                  break;
                }
                continue;
              }
              const ts = this.coverTiles(p, c, r);
              if (ts) {
                for (const ct of ts) this.show(p, ct, shown, used, 500 - badness(c, z));
                covered = true;
                break;
              }
            }
            if (covered) continue;
            if (t && t.plain) this.show(p, t, shown, used, 1000);
            else if (plainBudget > 0) {
              const t1 = performance.now();
              const pt = this.tiles.renderPlain(key, p.kind, level, r, this.linesIn(p.kind, r), this.frameNo);
              plainBudget -= performance.now() - t1;
              this.show(p, pt, shown, used, 1000);
            }
          }
        }
        if (exactSnap && this.snapCam) used.set('snap', 2000);
      }
      // what was shown last frame and not now
      for (const el of pe.shown) if (!shown.has(el)) el.style.display = 'none';
      pe.shown = shown;
      shownTotal += shown.size;
      this.placeLayers(p, pe, used);
      // at the starting zoom the grid's own picture is on top and has Univer's selection already
      if (this.snapCam && Math.abs(this.snapCam.z - z) < 1e-6) pe.sel.style.display = 'none';
      else if (isCells(p.kind)) this.selection(p, pe);
      else this.headerSelection(p, pe);
    }
    this.stats.layersShown = shownTotal;
    this.chrome();
    this.indicators();
  }

  /** Put tile t in its level's layer of pane p (once), mark it shown. */
  private show(p: Pane, t: Tile, shown: Set<HTMLCanvasElement>, used: Map<string, number>, rank: number) {
    if (shown.has(t.el)) return;
    this.tiles.touch(t, this.frameNo);
    const pe = this.panesEl.get(p.id)!;
    const lk = String(t.level);
    let layer = pe.layers.get(lk);
    const s = t.level * this.dpr;
    if (!layer) {
      const div = document.createElement('div');
      div.className = 'gs-layer';
      pe.div.insertBefore(div, pe.sel);
      // the layer's origin: a tile corner near the view (positions stay small numbers at any row)
      layer = { div, bx: t.sx, by: t.sy, level: t.level, css: '' };
      pe.layers.set(lk, layer);
    }
    // far from the layer's origin: move the origin (and every tile in the layer)
    if (Math.abs(t.sx - layer.bx) * s > 2e6 || Math.abs(t.sy - layer.by) * s > 2e6) {
      layer.bx = t.sx;
      layer.by = t.sy;
      for (const c of layer.div.children) (c as HTMLElement & { __tile?: Tile }).__tile && ((c as HTMLElement & { __tile?: Tile }).__tile!.placed = undefined);
    }
    const pos = `${Math.round((t.sx - layer.bx) * s) - BLEED},${Math.round((t.sy - layer.by) * s) - BLEED}`;
    if (t.el.parentElement !== layer.div) layer.div.appendChild(t.el);
    if (t.placed !== pos) {
      const [l, tp] = pos.split(',');
      t.el.style.cssText = `position:absolute;left:${l}px;top:${tp}px;width:${t.el.width}px;height:${t.el.height}px;pointer-events:none`;
      t.placed = pos;
      (t.el as HTMLElement & { __tile?: Tile }).__tile = t;
    } else if (t.el.style.display === 'none') t.el.style.display = '';
    shown.add(t.el);
    used.set(lk, Math.max(used.get(lk) ?? -1e9, rank));
  }

  /** One transform per layer: level css px → pane css px at the camera's zoom. */
  private placeLayers(p: Pane, pe: PaneEl, used: Map<string, number>) {
    const z = this.cam.z;
    const order = [...used.entries()].sort((a, b) => a[1] - b[1]);
    const zi = new Map(order.map(([k], i) => [k, i + 1]));
    for (const [lk, layer] of pe.layers) {
      const on = zi.has(lk);
      if (!on) {
        if (layer.css !== 'none') {
          layer.div.style.display = 'none';
          layer.css = 'none';
        }
        continue;
      }
      const sc = z / (layer.level * this.dpr);
      const tx = (layer.bx - p.ox) * z;
      const ty = (layer.by - p.oy) * z;
      const css = `transform:translate3d(${tx}px,${ty}px,0) scale(${sc});z-index:${zi.get(lk)}`;
      if (css !== layer.css) {
        layer.div.style.cssText = css;
        layer.css = css;
      }
    }
    const snap = pe.snap;
    if (snap) {
      const p0 = (snap as HTMLCanvasElement & { __pane?: Pane }).__pane;
      if (zi.has('snap') && p0 && this.snapCam) {
        if (snap.parentElement !== pe.div) pe.div.insertBefore(snap, pe.sel);
        const z0 = this.snapCam.z;
        const sc = z / (z0 * this.snapK);
        snap.style.cssText = `width:${snap.width}px;height:${snap.height}px;transform:translate3d(${(p0.ox - p.ox) * z}px,${(p0.oy - p.oy) * z}px,0) scale(${sc});z-index:${zi.get('snap')}`;
      } else if (snap.parentElement) snap.remove();
    }
  }

  /** Real tiles of level L that together cover r (null if one is missing). */
  private coverTiles(p: Pane, L: number, r: TileRect): Tile[] | null {
    const S = span(L, this.dpr);
    const g = this.g!;
    const b = BAND[p.kind];
    const x = this.axis(b.x, S, r.sx, r.sx + r.sw, g.fx0, g.fw, Infinity);
    const y = this.axis(b.y, S, r.sy, r.sy + r.sh, g.fy0, g.fh, Infinity);
    if ((x.i1 - x.i0 + 1) * (y.i1 - y.i0 + 1) > 16 || x.i1 < x.i0 || y.i1 < y.i0) return null;
    const out: Tile[] = [];
    for (let j = y.i0; j <= y.i1; j++)
      for (let i = x.i0; i <= x.i1; i++) {
        const t = this.tiles.get(this.key(p.kind, L, i, j));
        if (!t || t.plain) return null;
        out.push(t);
      }
    return out;
  }
  private snapCovers(p: Pane, r: TileRect) {
    const snap = this.panesEl.get(p.id)!.snap as (HTMLCanvasElement & { __pane?: Pane }) | null;
    const p0 = snap?.__pane;
    if (!p0 || !this.snapCam) return false;
    const z0 = this.snapCam.z;
    // the visible part of r must lie inside the picture the pane had at the start
    const z = this.cam.z;
    const ax = Math.max(r.sx, p.ox);
    const ay = Math.max(r.sy, p.oy);
    const bx = Math.min(r.sx + r.sw, p.ox + p.w / z);
    const by = Math.min(r.sy + r.sh, p.oy + p.h / z);
    return ax >= p0.ox - 0.01 && ay >= p0.oy - 0.01 && bx <= p0.ox + p0.w / z0 + 0.01 && by <= p0.oy + p0.h / z0 + 0.01;
  }

  /** Grid line positions (sheet px) inside r, for a placeholder. */
  private linesIn(kind: TileKind, r: TileRect) {
    const cols: number[] = [];
    const rows: number[] = [];
    for (let c = this.colAt(r.sx); c <= this.colAt(r.sx + r.sw) + 1 && cols.length < 600; c++) cols.push(this.colPos(c));
    for (let rr = this.rowAt(r.sy); rr <= this.rowAt(r.sy + r.sh) + 1 && rows.length < 2000; rr++) rows.push(this.rowPos(rr));
    return { cols, rows, grid: this.g!.gridlines || !isCells(kind) };
  }

  // -------------------------------------------------------------- plain parts
  private colPos(i: number) {
    const a = this.g!.colAcc;
    if (i <= 0) return 0;
    if (i <= a.length) return a[i - 1];
    return a[a.length - 1] + (i - a.length) * 88;
  }
  private rowPos(i: number) {
    const a = this.g!.rowAcc;
    if (i <= 0) return 0;
    if (i <= a.length) return a[i - 1];
    return a[a.length - 1] + (i - a.length) * 24;
  }
  private colAt(x: number) {
    const a = this.g!.colAcc;
    const end = a[a.length - 1] ?? 0;
    if (x >= end) return a.length + Math.floor((x - end) / 88);
    let lo = 0;
    let hi = a.length - 1;
    while (lo < hi) {
      const m = (lo + hi) >> 1;
      if (a[m] <= x) lo = m + 1;
      else hi = m;
    }
    return lo;
  }
  private rowAt(y: number) {
    const a = this.g!.rowAcc;
    const end = a[a.length - 1] ?? 0;
    if (y >= end) return a.length + Math.floor((y - end) / 24);
    let lo = 0;
    let hi = a.length - 1;
    while (lo < hi) {
      const m = (lo + hi) >> 1;
      if (a[m] <= y) lo = m + 1;
      else hi = m;
    }
    return lo;
  }

  /** Selection rectangle (the snapshot has Univer's own at the starting zoom; this one follows any zoom). */
  private selCache: Array<{ r0: number; c0: number; r1: number; c1: number }> | null = null;
  private sel() {
    // read once per gesture: the selection does not change while the sheet moves
    if (!this.selCache) {
      try {
        const r = doc.wb.getActiveSheet().getActiveRange()?.getRange();
        this.selCache = r ? [{ r0: r.startRow, c0: r.startColumn, r1: r.endRow, c1: r.endColumn }] : [];
      } catch {
        this.selCache = [];
      }
    }
    return this.selCache;
  }
  private selection(p: Pane, pe: PaneEl) {
    const s = this.sel()[0];
    const z = this.cam.z;
    if (!s) {
      pe.sel.style.display = 'none';
      return;
    }
    const x0 = (this.colPos(s.c0) - p.ox) * z;
    const x1 = (this.colPos(s.c1 + 1) - p.ox) * z;
    const y0 = (this.rowPos(s.r0) - p.oy) * z;
    const y1 = (this.rowPos(s.r1 + 1) - p.oy) * z;
    pe.sel.style.cssText = `display:${x1 < 0 || y1 < 0 || x0 > p.w || y0 > p.h ? 'none' : 'block'};transform:translate3d(${px(x0)}px,${px(y0)}px,0);width:${px(x1 - x0)}px;height:${px(y1 - y0)}px`;
  }
  /** The selected columns / rows are tinted in the headers, as Univer does. */
  private headerSelection(p: Pane, pe: PaneEl) {
    const s = this.sel()[0];
    const z = this.cam.z;
    if (!s) {
      pe.sel.style.display = 'none';
      return;
    }
    if (p.kind[0] === 'h') {
      const x0 = (this.colPos(s.c0) - p.ox) * z;
      const x1 = (this.colPos(s.c1 + 1) - p.ox) * z;
      pe.sel.style.cssText = `display:${x1 < 0 || x0 > p.w ? 'none' : 'block'};transform:translate3d(${px(x0)}px,0,0);width:${px(x1 - x0)}px;height:${px(p.h)}px`;
    } else {
      const y0 = (this.rowPos(s.r0) - p.oy) * z;
      const y1 = (this.rowPos(s.r1 + 1) - p.oy) * z;
      pe.sel.style.cssText = `display:${y1 < 0 || y0 > p.h ? 'none' : 'block'};transform:translate3d(0,${px(y0)}px,0);width:${px(p.w)}px;height:${px(y1 - y0)}px`;
    }
  }

  /** Corner box, header edges and the freeze dividers. */
  private chrome() {
    const g = this.g!;
    const z = this.cam.z;
    const hw = g.rh * z;
    const hh = g.ch * z;
    const e = this.chromeEls;
    e.corner.style.cssText = `width:${px(hw)}px;height:${px(hh)}px`;
    e.vline.style.cssText = `transform:translate3d(${px(hw - 1 / this.dpr)}px,0,0)`;
    e.hline.style.cssText = `transform:translate3d(0,${px(hh - 1 / this.dpr)}px,0)`;
    e.freezeV.style.cssText = g.fw > 0 ? `display:block;transform:translate3d(${px(hw + g.fw * z - 1)}px,0,0)` : 'display:none';
    e.freezeH.style.cssText = g.fh > 0 ? `display:block;transform:translate3d(0,${px(hh + g.fh * z - 1)}px,0)` : 'display:none';
  }

  private indicators() {
    const g = this.g!;
    const c = this.cam;
    const paneW = (g.W - (g.rh + g.fw) * c.z) / c.z;
    const paneH = (g.H - (g.ch + g.fh) * c.z) / c.z;
    const trackH = g.H - (g.ch + g.fh) * c.z - 8;
    const trackW = g.W - (g.rh + g.fw) * c.z - 8;
    const spanY = Math.max(paneH, g.maxY - g.minY + paneH * 0.5);
    const spanX = Math.max(paneW, g.maxX - g.minX + paneW * 0.5);
    const hV = Math.max(28, (paneH / spanY) * trackH);
    const wH = Math.max(28, (paneW / spanX) * trackW);
    const fy = Math.min(1, Math.max(0, (c.y - g.minY) / Math.max(1, spanY - paneH)));
    const fx = Math.min(1, Math.max(0, (c.x - g.minX) / Math.max(1, spanX - paneW)));
    this.barV.style.cssText = `opacity:${spanY > paneH + 1 && trackH > 40 ? 1 : 0};height:${hV}px;transform:translate3d(0,${(g.ch + g.fh) * c.z + 4 + fy * (trackH - hV)}px,0)`;
    this.barH.style.cssText = `opacity:${spanX > paneW + 1 && trackW > 40 ? 1 : 0};width:${wH}px;transform:translate3d(${(g.rh + g.fw) * c.z + 4 + fx * (trackW - wH)}px,0,0)`;
  }

  // ------------------------------------------------------------------ hand-off
  /** Fingers are up and nothing animates: put Univer at the camera, then hide the surface. */
  end() {
    if (!this.active || this.handing) return;
    this.anims = [];
    const { z, x, y } = this.cam;
    this.handing = { t0: performance.now(), z, x, y, stable: 0, tries: 0 };
    try {
      const g = this.g!;
      const ws = doc.wb.getActiveSheet();
      // the sheet grows to what is in view (plus room to pan further next time)
      const lastCol = this.colAt(x + g.W / z);
      const lastRow = this.rowAt(y + g.H / z);
      doc.ensureSize(ws.getSheetId(), lastRow + 1, lastCol + 1, true);
      const v = viewInfo();
      if (v && Math.abs(v.zoom - z) > 1e-6) {
        void setUniverZoom(z);
      } else if (v) {
        scrollTo(v, x, y);
      }
    } catch (e) {
      console.warn('hand-off', e);
    }
    this.request();
  }

  /** One frame of the hand-off: keep putting the scroll where the camera is until Univer stays there. */
  private handoffStep() {
    const h = this.handing;
    if (!h) return;
    h.tries++;
    const v = viewInfo();
    const unit = renderUnit();
    // Univer not at the zoom yet (it may not have been ready for it): ask again now and then
    if (v && Math.abs(v.zoom - h.z) >= 1e-6 && h.tries % 8 === 0) void setUniverZoom(h.z);
    if (v && unit && Math.abs(v.zoom - h.z) < 1e-6) {
      const o = mainOrigin(v);
      if (Math.abs(o.x - h.x) < 1 && Math.abs(o.y - h.y) < 1) h.stable++;
      else {
        h.stable = 0;
        scrollTo(v, h.x, h.y);
        unit.scene.makeDirty(true);
        unit.scene.render();
      }
    }
    if (h.stable >= 2 || h.tries > 90) this.finish();
  }

  private finish() {
    const h = this.handing;
    this.handing = null;
    this.active = false;
    if (this.hidden) this.hidden.style.visibility = '';
    this.hidden = null;
    this.root.style.display = 'none';
    this.wrap.classList.remove('gesturing');
    this.selCache = null;
    this.snapCam = null;
    for (const pe of this.panesEl.values()) {
      pe.snap?.remove();
      if (pe.snap) release(pe.snap);
    }
    this.restZoom = this.cam.z;
    this.sliderAnchor = null;
    if (h) this.stats.handoffMs = Math.round(performance.now() - h.t0);
    this.updateStats(true);
    this.showPill();
    this.onActive(false);
    this.schedulePrefetch();
  }

  /** Something on the sheet changed (an edit, a recalculation, fonts, pictures): tiles are redrawn. */
  invalidate() {
    this.tiles.invalidate();
    if (!this.active) this.schedulePrefetch();
  }
  /** Another sheet, file or theme: forget everything. */
  reset() {
    this.tiles.clear();
  }

  // -------------------------------------------------------------- prefetch
  /**
   * While the app is idle, draw the tiles around what is on screen: the resting zoom (for panning),
   * the next level down and up (for pinching), so a gesture starts with the cache filled.
   */
  schedulePrefetch() {
    if (this.idleHandle) return;
    const ric = (window as unknown as { requestIdleCallback?: (cb: (d: { timeRemaining(): number }) => void, o?: object) => number }).requestIdleCallback;
    const run = (deadline: { timeRemaining(): number; didTimeout?: boolean }) => {
      this.idleHandle = 0;
      if (this.active) return;
      const more = this.prefetch(deadline);
      if (more) this.schedulePrefetch();
    };
    // a short pause first: right after a gesture another one often follows
    this.idleHandle = setTimeout(() => {
      this.idleHandle = ric ? ric(run, { timeout: 2000 }) : (setTimeout(() => run({ timeRemaining: () => 8 }), 120) as unknown as number);
    }, 250) as unknown as number;
  }
  private prefetchList: Array<{ key: string; kind: TileKind; level: number; r: TileRect }> = [];
  private prefetchKey = '';
  private prefetch(deadline: { timeRemaining(): number; didTimeout?: boolean }): boolean {
    const v = viewInfo();
    const parts = sheetParts();
    if (!v || !parts || doc.gesturing) return false;
    this.dpr = window.devicePixelRatio || 1;
    this.tiles.checkSheet(parts, this.dpr);
    this.colours(parts);
    this.parts = parts;
    const o = mainOrigin(v);
    const key = `${this.tiles.ver}|${v.zoom}|${Math.round(o.x)}|${Math.round(o.y)}|${v.width}x${v.height}`;
    if (key !== this.prefetchKey) {
      this.prefetchKey = key;
      this.g = this.geometry(v, parts);
      this.restZoom = v.zoom;
      const saved = this.cam;
      this.cam = { z: v.zoom, x: o.x, y: o.y };
      const list: Array<{ key: string; kind: TileKind; level: number; r: TileRect; pri: number }> = [];
      const z = v.zoom;
      const below = [...LEVELS].reverse().find((L) => L < z * 0.99) ?? LEVELS[0];
      const above = LEVELS.find((L) => L > z * 1.01) ?? LEVELS[LEVELS.length - 1];
      for (const [L, rank] of [[z, 0], [below, 1], [above, 2]] as Array<[number, number]>) {
        const S = span(L, this.dpr);
        for (const p of this.panes()) {
          if (p.w < 0.5 || p.h < 0.5) continue;
          // the resting level gets a margin of one tile around the view, for panning
          const m = rank === 0 ? S : 0;
          const grow = { ...p, ox: p.ox - m, oy: p.oy - m, w: p.w + 2 * m * z, h: p.h + 2 * m * z };
          const { x, y } = this.range(grow, L);
          for (let j = y.i0; j <= y.i1; j++)
            for (let i = x.i0; i <= x.i1; i++) {
              const r = this.rectOf(p.kind, L, i, j);
              if (r.sw <= 0 || r.sh <= 0) continue;
              const inView = r.sx + r.sw > p.ox && r.sx < p.ox + p.w / z && r.sy + r.sh > p.oy && r.sy < p.oy + p.h / z;
              list.push({ key: this.key(p.kind, L, i, j), kind: p.kind, level: L, r, pri: rank * 10 + (inView ? 0 : 5) + (isCells(p.kind) ? 1 : 0) });
            }
        }
      }
      this.cam = saved;
      list.sort((a, b) => a.pri - b.pri);
      // never prefetch more than 60 % of the cache: the rest is for the gesture itself
      let bytes = 0;
      this.prefetchList = [];
      for (const j of list) {
        bytes += 1060000;
        if (bytes > this.tiles.capBytes * 0.6) break;
        this.prefetchList.push(j);
      }
    }
    let n = 0;
    while (this.prefetchList.length) {
      const j = this.prefetchList[0];
      const have = this.tiles.get(j.key);
      if (have && !have.plain) {
        this.prefetchList.shift();
        continue;
      }
      // only what fits in the idle time (one tile anyway when the browser was never idle for 2 s)
      if ((n > 0 || !deadline.didTimeout) && deadline.timeRemaining() < this.tiles.cost(j.kind, j.level) + 1) break;
      this.prefetchList.shift();
      this.tiles.render(parts, j.key, j.kind, j.level, j.r, this.frameNo);
      n++;
    }
    this.tiles.evict(this.frameNo, { x: o.x, y: o.y }, v.zoom);
    this.updateStats();
    return this.prefetchList.length > 0;
  }

  // -------------------------------------------------------------- zoom pill, debug overlay
  showPill() {
    const z = this.active ? this.cam.z : (viewInfo()?.zoom ?? 1);
    this.pill.textContent = `${Math.round(z * 100)} %`;
    this.pill.classList.add('on');
    clearTimeout(this.pillTimer);
    if (this.active) return;
    // stays while the sheet is not at 100 % (tap it to change the zoom); otherwise fades out
    if (Math.abs(z - 1) > 0.004) return;
    this.pillTimer = setTimeout(() => this.pill.classList.remove('on'), 1500);
  }

  setOverlay(on: boolean) {
    if (on && !this.overlay) {
      this.overlay = document.createElement('pre');
      this.overlay.className = 'perf-overlay';
      this.overlay.dataset.testid = 'perf-overlay';
      this.wrap.appendChild(this.overlay);
      let last = performance.now();
      let n = 0;
      let t0 = last;
      let worst = 0;
      const tick = (t: number) => {
        worst = Math.max(worst, t - last);
        last = t;
        n++;
        if (t - t0 >= 500) {
          if (!this.active) this.stats.fps = Math.round((n * 1000) / (t - t0));
          this.worst = worst;
          n = 0;
          t0 = t;
          worst = 0;
          this.updateStats(true);
        }
        this.overlayRaf = requestAnimationFrame(tick);
      };
      this.overlayRaf = requestAnimationFrame(tick);
    } else if (!on && this.overlay) {
      cancelAnimationFrame(this.overlayRaf);
      this.overlay.remove();
      this.overlay = null;
    }
  }

  private worst = 0;
  private statsT = 0;
  private updateStats(force = false) {
    const s = this.stats;
    s.pending = this.queue.size + (this.active ? 0 : this.prefetchList.length);
    s.tiles = this.tiles.size;
    s.tileBytes = this.tiles.bytes;
    s.tilesRendered = this.tiles.rendered;
    if (this.active) this.pill.textContent = `${Math.round(this.cam.z * 100)} %`;
    if (!this.overlay) return;
    const now = performance.now();
    if (!force && now - this.statsT < 250) return;
    this.statsT = now;
    this.overlay.textContent =
      `${s.fps} fps  worst ${this.worst.toFixed(1)} ms\n` +
      `frame ${s.frameMs.toFixed(1)} ms (max ${s.frameMsMax.toFixed(1)})\n` +
      `tiles pending ${s.pending}\n` +
      `cache ${s.tiles} tiles · ${(s.tileBytes / 1048576).toFixed(0)}/${(s.tileCap / 1048576).toFixed(0)} MB\n` +
      `zoom ${Math.round(s.zoom * 100)} % · level ${Math.round(s.level * 1000) / 10} %\n` +
      `hand-off ${s.handoffMs} ms · ${this.active ? 'gesture' : 'idle'}`;
  }
}

function reducedMotion() {
  try {
    return getTier().animations === 'reduced' || matchMedia('(prefers-reduced-motion: reduce)').matches;
  } catch {
    return false;
  }
}
function getTierRefresh(): number {
  try {
    return getTier().refreshRate || 60;
  } catch {
    return 60;
  }
}
