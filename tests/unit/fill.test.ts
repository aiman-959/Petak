import { describe, expect, it, vi } from 'vitest';

vi.mock('../../src/state/doc', () => ({ doc: {} }));
vi.mock('../../src/ui/store', () => ({ guard: () => {} }));
vi.mock('../../src/ui/editSession', () => ({ useEdit: () => false }));
vi.mock('../../src/ui/cellActions', () => ({ autoFill: () => {} }));
vi.mock('../../src/grid/view', () => ({ hitGrid: () => ({}), toScreen: () => ({}), viewInfo: () => null }));
const { fillTarget } = await import('../../src/ui/FillHandle');

describe('fill handle target', () => {
  const sel = { startRow: 2, endRow: 3, startColumn: 1, endColumn: 2 };
  it('extends along the axis the finger went further', () => {
    expect(fillTarget(sel, 7, 2)).toEqual({ ...sel, endRow: 7 });
    expect(fillTarget(sel, 0, 1)).toEqual({ ...sel, startRow: 0 });
    expect(fillTarget(sel, 3, 6)).toEqual({ ...sel, endColumn: 6 });
    expect(fillTarget(sel, 2, 0)).toEqual({ ...sel, startColumn: 0 });
    // diagonal: 4 rows down against 1 column right is a fill down
    expect(fillTarget(sel, 7, 3)).toEqual({ ...sel, endRow: 7 });
    expect(fillTarget(sel, 4, 8)).toEqual({ ...sel, endColumn: 8 });
  });
  it('inside the selection there is nothing to fill', () => {
    expect(fillTarget(sel, 2, 1)).toBeNull();
    expect(fillTarget(sel, 3, 2)).toBeNull();
  });
});
