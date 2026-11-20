/**
 * Fill handle for touch (M-01): a small grip under the bottom-right corner of the selection.
 * Drag it down / up / right / left over more cells and release: the selection is copied, or
 * continued as a series (1, 2, 3 → 4, 5; Mon, Tue → Wed; dates; formulas adjust), exactly as
 * Univer's auto-fill does for a mouse.
 */
import { useEffect, useRef, useState } from 'react';
import { ChevronsDown } from 'lucide-react';
import { doc } from '../state/doc';
import { guard } from './store';
import { useEdit } from './editSession';
import { autoFill } from './cellActions';
import { hitGrid, toScreen, viewInfo } from '../grid/view';

interface Rect {
  startRow: number;
  endRow: number;
  startColumn: number;
  endColumn: number;
}
/** The range a drag from `sel` to cell (row, col) fills: along the axis the finger went further. Pure. */
export function fillTarget(sel: Rect, row: number, col: number): Rect | null {
  const dr = row > sel.endRow ? row - sel.endRow : row < sel.startRow ? row - sel.startRow : 0;
  const dc = col > sel.endColumn ? col - sel.endColumn : col < sel.startColumn ? col - sel.startColumn : 0;
  if (!dr && !dc) return null;
  if (Math.abs(dr) >= Math.abs(dc)) return dr > 0 ? { ...sel, endRow: row } : { ...sel, startRow: row };
  return dc > 0 ? { ...sel, endColumn: col } : { ...sel, startColumn: col };
}
const at = (acc: number[], i: number) => (i <= 0 ? 0 : acc[Math.min(i, acc.length) - 1] ?? 0);

export function FillHandle() {
  const editing = useEdit((s) => s.editing);
  const [, tick] = useState(0);
  const [target, setTarget] = useState<Rect | null>(null);
  const drag = useRef<{ sel: Rect } | null>(null);
  useEffect(() => {
    let raf = 0;
    const again = () => {
      cancelAnimationFrame(raf);
      raf = requestAnimationFrame(() => tick((x) => x + 1));
    };
    // all workbooks: this grip stays mounted while files are opened (F-09)
    const d1 = doc.onCommand(again);
    const d2 = doc.subscribe(again);
    return () => {
      d1.dispose();
      d2();
      cancelAnimationFrame(raf);
    };
  }, []);
  const v = viewInfo();
  const sel = doc.wb.getActiveSheet().getActiveRange()?.getRange() as Rect | undefined;
  if (!v || !sel || editing || (doc.gesturing && !drag.current)) return null;
  const box = (r: Rect) => {
    const a = toScreen(v, at(v.colAcc, r.startColumn), at(v.rowAcc, r.startRow));
    const b = toScreen(v, at(v.colAcc, r.endColumn + 1), at(v.rowAcc, r.endRow + 1));
    return { left: a.x, top: a.y, width: b.x - a.x, height: b.y - a.y, right: b.x, bottom: b.y };
  };
  const s = box(sel);
  const host = (document.querySelector('[data-testid=grid]') as HTMLElement | null)?.getBoundingClientRect();
  // the corner must be on screen, below the headers
  if (!host || s.right < v.headerW + 8 || s.bottom < v.headerH + 8 || s.right > host.width - 4 || s.bottom > host.height - 4) return null;
  const move = (e: React.PointerEvent) => {
    if (!drag.current) return;
    const cur = viewInfo();
    if (!cur) return;
    const g = hitGrid(cur, e.clientX - host.left, e.clientY - host.top);
    if (g.area !== 'cell') return;
    setTarget(fillTarget(drag.current.sel, g.row, g.col));
  };
  const end = (commit: boolean) => {
    const d = drag.current;
    drag.current = null;
    const t = target;
    setTarget(null);
    if (commit && d && t)
      guard(async () => {
        await autoFill(d.sel, t);
        doc.wb.getActiveSheet().getRange(t.startRow, t.startColumn, t.endRow - t.startRow + 1, t.endColumn - t.startColumn + 1).activate();
      });
  };
  const tb = target && box(target);
  return (
    <>
      {tb && <div className="fill-preview" data-testid="fill-preview" style={{ left: tb.left, top: tb.top, width: tb.width, height: tb.height }} />}
      <button
        className="fill-handle"
        aria-label="Fill: drag over more cells"
        title="Drag to fill"
        data-testid="fill-handle"
        style={{ left: s.right - 14, top: s.bottom + 10 }}
        onPointerDown={(e) => {
          e.preventDefault();
          e.stopPropagation();
          e.currentTarget.setPointerCapture(e.pointerId);
          drag.current = { sel: { startRow: sel.startRow, endRow: sel.endRow, startColumn: sel.startColumn, endColumn: sel.endColumn } };
        }}
        onPointerMove={move}
        onPointerUp={() => end(true)}
        onPointerCancel={() => end(false)}
      >
        <ChevronsDown size={16} />
      </button>
    </>
  );
}
