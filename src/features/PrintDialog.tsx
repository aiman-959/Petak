/** File → Print / PDF: page setup for the active sheet, then a PDF to save or share (M-09). */
import { useState } from 'react';
import type { ICellData, IStyleData, IWorkbookData } from '@univerjs/core';
import { Dialog } from '../ui/Chrome';
import { guard, ui } from '../ui/store';
import { doc } from '../state/doc';
import { saveFile } from '../io/files';
import { defaultSetup, sheetToPdf, type PageSetup, type SheetSource } from './sheetPdf';

/** Everything the painter needs about the active sheet's used range. */
export function activeSheetSource(): SheetSource {
  const fws = doc.wb.getActiveSheet();
  const ws = fws.getSheet();
  const snap = doc.wb.save() as IWorkbookData;
  const sd = snap.sheets[fws.getSheetId()];
  const cells = (sd.cellData ?? {}) as Record<number, Record<number, ICellData>>;
  const styleOf = (s: ICellData['s']): IStyleData | null => (typeof s === 'string' ? ((snap.styles?.[s] as IStyleData) ?? null) : (s as IStyleData) ?? null);
  // used range: values, formulas, fills and borders
  let maxR = -1;
  let maxC = -1;
  for (const r in cells)
    for (const c in cells[r]) {
      const cd = cells[r][c];
      const st = styleOf(cd?.s);
      if ((cd?.v !== undefined && cd.v !== null && cd.v !== '') || cd?.f || cd?.p || st?.bg || st?.bd) {
        maxR = Math.max(maxR, +r);
        maxC = Math.max(maxC, +c);
      }
    }
  const merges = (sd.mergeData ?? []).map((m) => ({ r0: m.startRow, c0: m.startColumn, r1: m.endRow, c1: m.endColumn }));
  for (const m of merges) {
    if (m.r0 <= maxR && m.c0 <= maxC) {
      maxR = Math.max(maxR, m.r1);
      maxC = Math.max(maxC, m.c1);
    }
  }
  if (maxR < 0) throw new Error('This sheet is empty: nothing to print.');
  const rows = [...Array(maxR + 1).keys()].filter((r) => ws.getRowVisible(r));
  const cols = [...Array(maxC + 1).keys()].filter((c) => ws.getColVisible(c));
  const display = fws.getRange(0, 0, maxR + 1, maxC + 1).getDisplayValues();
  const freeze = fws.getFreeze();
  const titleRows = freeze && freeze.ySplit > 0 ? rows.filter((r) => r < freeze.ySplit) : [];
  const def = (snap as { defaultStyle?: { ff?: string; fs?: number } }).defaultStyle;
  return {
    name: fws.getSheetName(),
    rows,
    cols,
    rowH: (r) => fws.getRowHeight(r),
    colW: (c) => fws.getColumnWidth(c),
    titleRows,
    text: (r, c) => display[r]?.[c] ?? '',
    style: (r, c) => styleOf(cells[r]?.[c]?.s),
    merges,
    defaultFont: { ff: def?.ff ?? 'Arial', fs: def?.fs ?? 11 },
  };
}

export default function PrintDialog() {
  const [setup, setSetup] = useState<PageSetup>(defaultSetup);
  const set = (p: Partial<PageSetup>) => setSetup((s) => ({ ...s, ...p }));
  const make = () =>
    guard(async () => {
      ui.progress(0, 1, 'Drawing pages');
      const src = activeSheetSource();
      const { pdf, pages } = await sheetToPdf(src, setup);
      const base = doc.snapshot.file.name.replace(/\.[^.]+$/, '') || 'Sheet';
      await saveFile(pdf.buffer.slice(pdf.byteOffset, pdf.byteOffset + pdf.byteLength) as ArrayBuffer, `${base} - ${src.name}.pdf`.replace(/[\\/:*?"<>|]+/g, '_'), 'application/pdf');
      ui.close();
      ui.toast(`PDF saved: ${pages} page${pages === 1 ? '' : 's'}`);
    });
  return (
    <Dialog title="Print / PDF" onClose={() => ui.close()} testId="dlg-print">
      <p className="muted">The active sheet’s used range, at the size Excel prints it. Frozen rows repeat on every page.</p>
      <div className="seg" role="radiogroup" aria-label="Paper">
        {(['a4', 'letter'] as const).map((p) => (
          <button key={p} className={'chip' + (setup.paper === p ? ' on' : '')} aria-pressed={setup.paper === p} data-testid={'print-paper-' + p} onClick={() => set({ paper: p })}>
            {p === 'a4' ? 'A4' : 'Letter'}
          </button>
        ))}
        <button className={'chip' + (!setup.landscape ? ' on' : '')} data-testid="print-portrait" onClick={() => set({ landscape: false })}>
          Portrait
        </button>
        <button className={'chip' + (setup.landscape ? ' on' : '')} data-testid="print-landscape" onClick={() => set({ landscape: true })}>
          Landscape
        </button>
      </div>
      <label className="check">
        <input type="checkbox" checked={setup.fitWidth} data-testid="print-fit" onChange={(e) => set({ fitWidth: e.target.checked })} /> Fit all columns on the page width
      </label>
      <label className="check">
        <input type="checkbox" checked={setup.gridlines} data-testid="print-grid" onChange={(e) => set({ gridlines: e.target.checked })} /> Print grid lines
      </label>
      <div className="row" style={{ justifyContent: 'flex-end' }}>
        <button className="primary" data-testid="print-save" onClick={make}>
          Save PDF
        </button>
      </div>
    </Dialog>
  );
}
