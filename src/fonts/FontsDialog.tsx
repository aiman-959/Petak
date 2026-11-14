/** Format → Fonts: font and size for the selection, what each Excel font is shown with, own font files. */
import { useEffect, useRef, useState, useSyncExternalStore } from 'react';
import { Plus, Trash2 } from 'lucide-react';
import { Dialog } from '../ui/Chrome';
import { guard, ui } from '../ui/store';
import { activeRange } from '../ui/actions';
import { doc } from '../state/doc';
import { BUILT_IN, ensureFonts, fontReport, fonts, probeFonts, substituteFor } from './fonts';

const SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48, 72];
let version = 0;
fonts.subscribe(() => version++);

export default function FontsDialog() {
  useSyncExternalStore(fonts.subscribe, () => version);
  const file = useRef<HTMLInputElement>(null);
  const [cur, setCur] = useState<{ ff: string; fs: number }>({ ff: '', fs: 11 });
  const read = () => {
    const st = activeRange().getCellStyleData() as { ff?: string; fs?: number } | null;
    const def = (doc.wb.getWorkbook().getSnapshot() as { defaultStyle?: { ff?: string; fs?: number } }).defaultStyle;
    setCur({ ff: st?.ff ?? def?.ff ?? 'Arial', fs: st?.fs ?? def?.fs ?? 11 });
  };
  const [, probed] = useState(0);
  useEffect(read, []);
  const setFamily = (ff: string) =>
    guard(async () => {
      await ensureFonts([ff]);
      activeRange().setFontFamily(ff);
      read();
    });
  const setSize = (fs: number) =>
    guard(() => {
      if (!(fs >= 1 && fs <= 409)) throw new Error('Font size must be between 1 and 409');
      activeRange().setFontSize(fs);
      read();
    });
  const add = (list: FileList | null) =>
    guard(async () => {
      for (const f of Array.from(list ?? [])) {
        const added = await fonts.add(f.name, await f.arrayBuffer());
        ui.toast(`Added ${added.family}${added.bold ? ' Bold' : ''}${added.italic ? ' Italic' : ''}`);
      }
      if (file.current) file.current.value = '';
    });
  const families = fonts.families();
  const shown = cur.ff && !families.some((f) => f.toLowerCase() === cur.ff.toLowerCase()) ? [cur.ff, ...families] : families;
  useEffect(() => void probeFonts(shown).then(() => probed((x) => x + 1)), [shown.join('|')]);
  const report = fontReport(shown);
  return (
    <Dialog title="Fonts" onClose={() => ui.close()} testId="dlg-fonts">
      <h4>Font of the selected cells</h4>
      <div className="font-list" role="listbox" aria-label="Font">
        {shown.map((f) => {
          const sub = report.substituted.find(([n]) => n === f)?.[1];
          const missing = report.missing.includes(f);
          return (
            <button key={f} role="option" aria-selected={cur.ff.toLowerCase() === f.toLowerCase()} className={'font-row' + (cur.ff.toLowerCase() === f.toLowerCase() ? ' on' : '')} data-testid={'font-' + f.replace(/\s+/g, '-')} onClick={() => setFamily(f)}>
              <span style={{ fontFamily: `"${f}"` }}>{f}</span>
              <small>{missing ? 'not on this device' : sub ? `shown with ${sub} (same widths)` : !BUILT_IN.includes(f) && fonts.user().some((u) => u.family === f) ? 'your font' : 'on this device'}</small>
            </button>
          );
        })}
      </div>
      <h4>Size</h4>
      <div className="size-row">
        {SIZES.map((s) => (
          <button key={s} className={'chip' + (cur.fs === s ? ' on' : '')} data-testid={'font-size-' + s} onClick={() => setSize(s)}>
            {s}
          </button>
        ))}
        <input className="input" inputMode="decimal" aria-label="Font size" data-testid="font-size-input" defaultValue={cur.fs} key={cur.fs} onKeyDown={(e) => e.key === 'Enter' && setSize(parseFloat((e.target as HTMLInputElement).value.replace(',', '.')))} onBlur={(e) => parseFloat(e.target.value) !== cur.fs && e.target.value.trim() && setSize(parseFloat(e.target.value.replace(',', '.')))} />
      </div>
      <h4>Your fonts</h4>
      {fonts.user().length === 0 && <p className="muted">None yet. Add a .ttf, .otf, .woff or .woff2 file; it stays on this device and is used whenever a workbook asks for that font.</p>}
      {fonts.user().map((u) => (
        <div key={u.id} className="font-user" data-testid="font-user">
          <span>
            {u.family}
            {u.bold ? ' Bold' : ''}
            {u.italic ? ' Italic' : ''} <small>{(u.size / 1024).toFixed(0)} KB</small>
          </span>
          <button className="icon-btn" aria-label={`Remove ${u.family}`} onClick={() => guard(() => fonts.remove(u.id))}>
            <Trash2 size={18} />
          </button>
        </div>
      ))}
      <button className="secondary" data-testid="font-add" onClick={() => file.current?.click()}>
        <Plus size={16} /> Add a font file
      </button>
      <input ref={file} type="file" hidden multiple accept=".ttf,.otf,.woff,.woff2,font/*" data-testid="font-file" onChange={(e) => add(e.target.files)} />
      <p className="muted">
        Calibri, Cambria, Arial, Times New Roman and Courier New are drawn with built-in fonts of the same widths ({['Calibri', 'Cambria', 'Arial', 'Times New Roman', 'Courier New'].map((n) => substituteFor(n)!.family).join(', ')}), so text takes the same space as in Excel. Files are always saved with the original font names.
      </p>
    </Dialog>
  );
}
