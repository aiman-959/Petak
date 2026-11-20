import { describe, expect, it } from 'vitest';
import { LEVELS, MAX_ZOOM, MIN_ZOOM, TILE, anchorAt, badness, cameraFor, clampCamera, doubleTapZoom, fallbackLevels, span, targetLevel, type Frame } from '../../src/grid/camera';

// Pixel 7 grid area; roster-like frozen panes (2 columns = 168 px, 3 rows = 96 px)
const frozen: Frame = { rh: 46, ch: 20, fw: 168, fh: 96, W: 412, H: 626, minX: 168, minY: 96, maxX: 88 * 1000, maxY: 24 * 50000 };
const plain: Frame = { ...frozen, fw: 0, fh: 0, minX: 0, minY: 0 };
const screenX = (f: Frame, c: { z: number; x: number }, sx: number) => c.z * (f.rh + f.fw + sx - c.x);
const screenY = (f: Frame, c: { z: number; y: number }, sy: number) => c.z * (f.ch + f.fh + sy - c.y);

describe('pinch around the fingers', () => {
  for (const f of [plain, frozen]) {
    it(`keeps the sheet point under the fingers at every zoom (${f.fw ? 'frozen panes' : 'no freeze'})`, () => {
      const c0 = { z: 1, x: f.minX + 5000, y: f.minY + 30000 }; // far from A1: a 10 % view is 10× wider
      const focal = { x: 250, y: 400 };
      const a = anchorAt(f, c0, focal.x, focal.y);
      for (const z of [0.1, 0.25, 0.5, 1, 2, 4]) {
        const c = cameraFor(f, a, focal.x, focal.y, z);
        expect(screenX(f, c, a.x)).toBeCloseTo(focal.x, 6);
        expect(screenY(f, c, a.y)).toBeCloseTo(focal.y, 6);
      }
    });
  }
  it('fingers moving while pinching also pan', () => {
    const c0 = { z: 1, x: 1000, y: 1000 };
    const a = anchorAt(plain, c0, 200, 300);
    const c = cameraFor(plain, a, 260, 340, 1.5);
    expect(screenX(plain, c, a.x)).toBeCloseTo(260, 6);
    expect(screenY(plain, c, a.y)).toBeCloseTo(340, 6);
  });
  it('zoom stays within 10 % … 400 %', () => {
    const a = { x: 500, y: 500 };
    expect(cameraFor(plain, a, 200, 300, 0.01).z).toBe(MIN_ZOOM);
    expect(cameraFor(plain, a, 200, 300, 40).z).toBe(MAX_ZOOM);
    expect(MIN_ZOOM).toBe(0.1);
    expect(MAX_ZOOM).toBe(4);
  });
});

describe('camera limits', () => {
  it('cannot scroll into the frozen panes or before A1', () => {
    const c = clampCamera(frozen, { z: 1, x: -500, y: -500 });
    expect(c.x).toBe(frozen.minX);
    expect(c.y).toBe(frozen.minY);
  });
  it('the sheet end may come to the middle of the pane, not further', () => {
    const c = clampCamera(plain, { z: 1, x: 1e9, y: 1e9 });
    expect(c.x).toBeCloseTo(plain.maxX - (412 - 46) / 2);
    expect(c.y).toBeCloseTo(plain.maxY - (626 - 20) / 2);
  });
  it('a small sheet zoomed out stays at its start', () => {
    const tiny: Frame = { ...plain, maxX: 300, maxY: 200 };
    const c = clampCamera(tiny, { z: 0.1, x: 50, y: 50 });
    expect(c).toEqual({ z: 0.1, x: 0, y: 0 });
  });
});

describe('zoom levels (like a map app)', () => {
  it('fixed levels from 12.5 % to 400 %, each a power of two apart', () => {
    expect(LEVELS).toEqual([0.125, 0.25, 0.5, 1, 2, 4]);
  });
  it('tiles of neighbouring levels nest (one tile = 2 × 2 tiles one level up)', () => {
    for (const dpr of [1, 1.5, 2, 2.625, 3]) for (let i = 1; i < LEVELS.length; i++) expect(span(LEVELS[i - 1], dpr)).toBeCloseTo(2 * span(LEVELS[i], dpr));
  });
  it('a tile is 512 device pixels: far inside every WebView canvas limit', () => {
    expect(TILE).toBe(512);
    expect(TILE * TILE * 4).toBe(1048576);
  });
  it('the resting zoom is drawn exactly; any other zoom uses the next level up (≤ 15 % upscale)', () => {
    expect(targetLevel(0.85, 0.85)).toBe(0.85);
    expect(targetLevel(0.85, 1)).toBe(1);
    expect(targetLevel(1.1, 1)).toBe(1);
    expect(targetLevel(1.2, 1)).toBe(2);
    expect(targetLevel(0.1, 1)).toBe(0.125);
    expect(targetLevel(4, 1)).toBe(4);
    for (let z = MIN_ZOOM; z <= MAX_ZOOM; z += 0.01) {
      const L = targetLevel(z, -1);
      expect(L).toBeGreaterThanOrEqual(z / 1.15 - 1e-9);
      expect(L / z).toBeLessThanOrEqual(2.3); // never more than ~2.3× the pixels needed
    }
  });
  it('fallback levels go worst first and upscaling counts as worse', () => {
    expect(badness(1, 1)).toBe(0);
    expect(badness(0.5, 1)).toBeGreaterThan(badness(2, 1));
    const lv = fallbackLevels(1, 1, [0.85]);
    expect(lv[lv.length - 1]).toBe(0.85); // the nearest is drawn last, on top
    expect(lv).not.toContain(1);
    expect(fallbackLevels(0.1, 0.125, [])).not.toContain(4); // 40× the tiles: skipped
  });
});

describe('double tap', () => {
  it('zooms in to 200 % and back to 100 %', () => {
    expect(doubleTapZoom(1)).toBe(2);
    expect(doubleTapZoom(2)).toBe(1);
    expect(doubleTapZoom(0.5)).toBe(1);
    expect(doubleTapZoom(3.2)).toBe(1);
  });
});
