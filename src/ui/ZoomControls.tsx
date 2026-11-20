/**
 * Zoom panel (File menu → Zoom, or tap the zoom pill on the grid): presets and a slider; the
 * grid stays visible while it is open so the zoom can be judged. Settings dialog (File menu →
 * Settings): double tap to zoom, and the developer performance overlay.
 */
import { useEffect, useRef, useState } from 'react';
import { X, ZoomIn, ZoomOut } from 'lucide-react';
import { doc } from '../state/doc';
import { onBack } from './back';
import { ui } from './store';
import { Dialog } from './Chrome';
import { settings, useSettings } from './settings';
import { MAX_ZOOM, MIN_ZOOM } from '../grid/gestureView';
import { setUniverZoom } from '../grid/view';

const ZOOM_PRESETS = [50, 75, 100, 150, 200];

/** Set the sheet's zoom (animated on touch devices, around the middle of the sheet). */
function setZoom(z: number) {
  const zz = Math.max(MIN_ZOOM, Math.min(MAX_ZOOM, z));
  if (doc.view?.zoomTo(zz)) return;
  // desktop (no gesture renderer): Univer zooms directly
  void setUniverZoom(zz);
}
const currentZoom = () => {
  try {
    return doc.view?.zoom ?? (doc.wb.getActiveSheet() as unknown as { getZoom(): number }).getZoom() ?? 1;
  } catch {
    return 1;
  }
};

export function ZoomPanel() {
  const [pct, setPct] = useState(() => Math.round(currentZoom() * 100));
  const slider = useRef<HTMLInputElement>(null);
  const close = () => ui.close();
  useEffect(() => onBack(() => (close(), true)), []);
  // the slider previews on the gesture surface while it moves and commits when released
  useEffect(() => {
    const el = slider.current;
    if (!el) return;
    const commit = () => {
      const v = +el.value;
      if (doc.view) doc.view.commitZoom();
      else setZoom(v / 100);
    };
    el.addEventListener('change', commit);
    return () => el.removeEventListener('change', commit);
  }, []);
  useEffect(() => {
    // follow pinches made while the panel is open
    const t = setInterval(() => !doc.gesturing && setPct(Math.round(currentZoom() * 100)), 300);
    return () => clearInterval(t);
  }, []);
  const preset = (p: number) => {
    setPct(p);
    setZoom(p / 100);
  };
  return (
    <div className="zoom-panel" role="dialog" aria-label="Zoom" data-testid="zoom-panel">
      <div className="zoom-row">
        <button className="icon-btn" aria-label="Zoom out" data-testid="zoom-out" onClick={() => preset(Math.max(MIN_ZOOM * 100, Math.round(pct / 1.25 / 5) * 5))}>
          <ZoomOut size={22} />
        </button>
        <input
          ref={slider}
          type="range"
          min={MIN_ZOOM * 100}
          max={MAX_ZOOM * 100}
          step={5}
          value={pct}
          aria-label="Zoom"
          data-testid="zoom-slider"
          onChange={(e) => {
            const v = +e.target.value;
            setPct(v);
            if (doc.view) doc.view.previewZoom(v / 100);
          }}
        />
        <button className="icon-btn" aria-label="Zoom in" data-testid="zoom-in" onClick={() => preset(Math.min(MAX_ZOOM * 100, Math.round((pct * 1.25) / 5) * 5))}>
          <ZoomIn size={22} />
        </button>
        <b className="zoom-value" data-testid="zoom-value">{pct} %</b>
        <button className="icon-btn" aria-label="Close" data-testid="zoom-close" onClick={close}>
          <X size={22} />
        </button>
      </div>
      <div className="zoom-row presets">
        {ZOOM_PRESETS.map((p) => (
          <button key={p} className={'chip' + (p === pct ? ' on' : '')} data-testid={`zoom-preset-${p}`} onClick={() => preset(p)}>
            {p} %
          </button>
        ))}
      </div>
    </div>
  );
}

export function SettingsDialog() {
  const s = useSettings();
  return (
    <Dialog title="Settings" onClose={() => ui.close()} testId="dlg-settings">
      <h4>Zoom</h4>
      <label className="switch-row">
        <input type="checkbox" checked={s.doubleTapZoom} data-testid="set-doubletap-zoom" onChange={(e) => settings.set({ doubleTapZoom: e.target.checked })} />
        <span>
          Double tap on empty space to zoom
          <small className="muted">Away from your data a double tap zooms in, or back to 100 %. On or next to data it still edits the cell.</small>
        </span>
      </label>
      <button className="secondary" onClick={() => ui.open('zoom')}>
        Zoom…
      </button>
      <h4>Developer</h4>
      <label className="switch-row">
        <input type="checkbox" checked={s.perfOverlay} data-testid="set-perf-overlay" onChange={(e) => settings.set({ perfOverlay: e.target.checked })} />
        <span>
          Performance overlay
          <small className="muted">Frame rate, frame time, tiles waiting to be drawn, tile cache size and zoom level, over the grid.</small>
        </span>
      </label>
    </Dialog>
  );
}
