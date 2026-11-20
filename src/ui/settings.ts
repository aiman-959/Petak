/** App settings (File menu → Settings), kept in local preferences. */
import { useSyncExternalStore } from 'react';
import { prefs } from '../state/storage';

export interface Settings {
  /** double tap on empty space away from the data zooms in / back to 100 % */
  doubleTapZoom: boolean;
  /** Developer: frame rate, frame time, tiles pending, tile cache and zoom level over the grid */
  perfOverlay: boolean;
}

const DEFAULTS: Settings = { doubleTapZoom: true, perfOverlay: false };

function load(): Settings {
  try {
    return { ...DEFAULTS, ...JSON.parse(prefs.get('settings', '{}')) };
  } catch {
    return { ...DEFAULTS };
  }
}

let state = load();
const listeners = new Set<() => void>();

export const settings = {
  get: () => state,
  set(patch: Partial<Settings>) {
    state = { ...state, ...patch };
    prefs.set('settings', JSON.stringify(state));
    listeners.forEach((l) => l());
  },
  subscribe(l: () => void) {
    listeners.add(l);
    return () => listeners.delete(l);
  },
};

export const useSettings = () => useSyncExternalStore(settings.subscribe, settings.get);
