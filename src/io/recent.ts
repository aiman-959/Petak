/** Recently opened / saved files (the last 12 that can be reached again). */
import { kv } from '../state/storage';

export interface RecentFile {
  name: string;
  /** Android content:// uri with a persistable grant */
  uri?: string;
  /** browser File System Access handle */
  handle?: FileSystemFileHandle;
  at: number;
}
const KEY = 'recent:files';
const MAX = 12;
const same = (a: { name: string; uri?: string }, b: { name: string; uri?: string }) => (a.uri || b.uri ? a.uri === b.uri : a.name === b.name);

/** Newest first, one entry per file (pure: unit-tested). */
export function withRecent(list: RecentFile[], f: RecentFile, max = MAX): RecentFile[] {
  return [f, ...list.filter((x) => !same(x, f))].slice(0, max);
}
export const recent = {
  list: async (): Promise<RecentFile[]> => (await kv.get<RecentFile[]>(KEY).catch(() => undefined)) ?? [],
  /** Only files that can be opened again are kept (a plain browser upload cannot). */
  async add(f: { name: string; uri?: string; handle?: FileSystemFileHandle }) {
    if (!f.uri && !f.handle) return;
    await kv.set(KEY, withRecent(await this.list(), { ...f, at: Date.now() })).catch(() => {});
  },
  async remove(f: { name: string; uri?: string }) {
    await kv.set(KEY, (await this.list()).filter((x) => !same(x, f))).catch(() => {});
  },
  clear: () => kv.del(KEY),
};
(globalThis as unknown as { __recent: typeof recent }).__recent = recent;
