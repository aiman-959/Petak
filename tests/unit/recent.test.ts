import { describe, expect, it } from 'vitest';
import { withRecent } from '../../src/io/recent';

describe('recent files', () => {
  it('newest first, one entry per file, limited', () => {
    let l = withRecent([], { name: 'a.xlsx', uri: 'u:a', at: 1 });
    l = withRecent(l, { name: 'b.xlsx', uri: 'u:b', at: 2 });
    l = withRecent(l, { name: 'a.xlsx', uri: 'u:a', at: 3 });
    expect(l.map((x) => [x.name, x.at])).toEqual([['a.xlsx', 3], ['b.xlsx', 2]]);
    // same name in another place is another file
    expect(withRecent(l, { name: 'a.xlsx', uri: 'u:other', at: 4 })).toHaveLength(3);
    for (let i = 0; i < 20; i++) l = withRecent(l, { name: `f${i}`, uri: `u:${i}`, at: 10 + i });
    expect(l).toHaveLength(12);
    expect(l[0].name).toBe('f19');
  });
});
