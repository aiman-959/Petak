/** While a formula is typed, each reference is outlined on the grid in its colour (M-03). */
import { useEffect, useState } from 'react';
import { doc } from '../state/doc';
import { useEdit } from './editSession';
import { formulaRefs } from './formulaRefs';
import { toScreen, viewInfo } from '../grid/view';

const at = (acc: number[], i: number) => (i <= 0 ? 0 : acc[Math.min(i, acc.length) - 1] ?? 0);

export function RefHighlights() {
  const editing = useEdit((s) => s.editing);
  const text = useEdit((s) => s.text);
  const [, tick] = useState(0);
  useEffect(() => {
    if (!editing) return;
    let raf = 0;
    const again = () => {
      cancelAnimationFrame(raf);
      raf = requestAnimationFrame(() => tick((x) => x + 1));
    };
    const d1 = doc.wb.onCommandExecuted(again);
    const d2 = doc.subscribe(again);
    return () => {
      d1.dispose();
      d2();
      cancelAnimationFrame(raf);
    };
  }, [editing]);
  if (!editing || !text.startsWith('=') || doc.gesturing) return null;
  const v = viewInfo();
  const host = document.querySelector('[data-testid=grid]') as HTMLElement | null;
  if (!v || !host) return null;
  const W = host.clientWidth;
  const H = host.clientHeight;
  const sheetName = doc.wb.getActiveSheet().getSheetName().toLowerCase();
  const boxes = formulaRefs(text)
    .filter((r) => !r.sheet || r.sheet.toLowerCase() === sheetName)
    .map((r, i) => {
      const a = toScreen(v, at(v.colAcc, r.c0), at(v.rowAcc, r.r0));
      const b = toScreen(v, at(v.colAcc, r.c1 + 1), at(v.rowAcc, r.r1 + 1));
      // whole columns / rows reach far beyond the screen: clip to the grid
      const x0 = Math.max(v.headerW, a.x);
      const y0 = Math.max(v.headerH, a.y);
      const x1 = Math.min(W, b.x);
      const y1 = Math.min(H, b.y);
      if (x1 <= x0 || y1 <= y0) return null;
      return <div key={i} className="ref-box" data-testid="ref-box" data-color={r.color} style={{ left: x0, top: y0, width: x1 - x0, height: y1 - y0, borderColor: r.color, background: r.color + '14' }} />;
    });
  return <>{boxes}</>;
}
