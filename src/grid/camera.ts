/**
 * Camera maths of the gesture renderer, kept free of the DOM so it can be unit-tested
 * (tests/unit/camera.test.ts).
 *
 * The camera is a zoom z and the sheet position (x, y) at the top-left corner of the scrolling pane.
 * On screen, left to right: the row headers (rh sheet px wide), the frozen columns (fw), then the
 * scrolling pane; top to bottom likewise with ch and fh. Headers and frozen panes scale with z, so
 * a sheet point sx in the scrolling pane is at screen x = z · (rh + fw + sx − x).
 */
/** a tile is TILE × TILE device pixels (1 MB of canvas) */
export const TILE = 512;
/** fixed zoom levels (powers of two): any zoom from 10 % to 400 % is within 2× of one of them */
export const LEVELS = [0.125, 0.25, 0.5, 1, 2, 4];
/** sheet px covered by one tile side at `level` */
export const span = (level: number, dpr: number) => TILE / (level * dpr);

export const MIN_ZOOM = 0.1;
export const MAX_ZOOM = 4;

export interface Camera {
  z: number;
  x: number;
  y: number;
}
export interface Frame {
  /** header sizes and frozen sizes, sheet px */
  rh: number;
  ch: number;
  fw: number;
  fh: number;
  /** the surface, css px */
  W: number;
  H: number;
  /** scroll limits of the pane's top-left corner */
  minX: number;
  minY: number;
  /** the sheet's current extent (sheet px) */
  maxX: number;
  maxY: number;
}

/** The sheet point under surface point (px, py), in the scrolling pane's coordinates (also over frozen panes and headers). */
export function anchorAt(f: Frame, c: Camera, px: number, py: number) {
  return { x: c.x + px / c.z - f.rh - f.fw, y: c.y + py / c.z - f.ch - f.fh };
}

/** Keep the camera on the sheet: the last row / column may come up to the middle of the scrolling pane. */
export function clampCamera(f: Frame, c: Camera): Camera {
  const z = Math.max(MIN_ZOOM, Math.min(MAX_ZOOM, c.z));
  const paneW = (f.W - (f.rh + f.fw) * z) / z;
  const paneH = (f.H - (f.ch + f.fh) * z) / z;
  const maxX = Math.max(f.minX, f.maxX - paneW * 0.5);
  const maxY = Math.max(f.minY, f.maxY - paneH * 0.5);
  return { z, x: Math.max(f.minX, Math.min(maxX, c.x)), y: Math.max(f.minY, Math.min(maxY, c.y)) };
}

/** Camera that puts sheet point `a` under surface point (px, py) at zoom z (clamped). */
export function cameraFor(f: Frame, a: { x: number; y: number }, px: number, py: number, z: number): Camera {
  const zz = Math.max(MIN_ZOOM, Math.min(MAX_ZOOM, z));
  return clampCamera(f, { z: zz, x: a.x + f.rh + f.fw - px / zz, y: a.y + f.ch + f.fh - py / zz });
}

/**
 * Tile level to draw for zoom z: the zoom the sheet rests at (exactly as sharp as the grid), else
 * the next fixed level up; a little upscaling (up to 15 %) is allowed to avoid 4× the tiles.
 */
export function targetLevel(z: number, rest: number): number {
  if (Math.abs(z - rest) < 1e-4) return rest;
  for (const L of LEVELS) if (L >= z / 1.15) return L;
  return LEVELS[LEVELS.length - 1];
}

/** How bad tiles of level L look at zoom z: 0 is exact; upscaled (blurry) is worse than downscaled. */
export function badness(L: number, z: number) {
  return L >= z ? Math.log2(L / z) : 1.6 * Math.log2(z / L);
}

/** Cached levels to draw under the best one, worst first; levels needing far too many tiles are skipped. */
export function fallbackLevels(z: number, best: number, extra: number[]): number[] {
  const set = new Set<number>([...LEVELS, ...extra]);
  set.delete(best);
  return [...set].filter((L) => L <= z * 4 && L >= z / 16).sort((a, b) => badness(b, z) - badness(a, z));
}

/** Double tap: from 100 % (or anything up to 150 %) to 200 %, otherwise back to 100 %. */
export function doubleTapZoom(z: number) {
  return z < 0.95 || z >= 1.5 ? 1 : 2;
}
