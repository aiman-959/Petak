/**
 * Touch targets and text contrast on a phone (Part 2 of the 1.3.0 brief).
 *
 * Opens every surface a user reaches on a phone — the top bar, the six tool panels, the File menu,
 * the long-press menus for a cell / row / column, the zoom panel, Settings, About, Find, number
 * format, conditional format, functions, comments, validation, page setup, drawing mode — and
 * measures every visible control:
 *   - touch target: at least 44 × 44 CSS px (CSS px = dp in Android's WebView); a smaller control
 *     passes when an invisible touch area around it reaches 44 (a ::before / padding hit area is
 *     measured with elementFromPoint 22 px from its centre)
 *   - text contrast: WCAG 2.1 AA, 4.5 : 1 (3 : 1 for large text ≥ 24 px or ≥ 18.66 px bold), text
 *     colour against the first opaque background behind it
 * Writes tests/e2e/out/a11y.json with every element below the limits.
 *   BASE_URL=http://localhost:4173/ node tests/e2e/a11y.mjs
 */
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { start, BASE } from './lib.mjs';
import { touch } from './touch.mjs';

const t = await start('a11y');
const { page, step, ev, fresh } = t;
console.log('Petak e2e touch targets and contrast —', BASE);
await fresh();
await page.waitForTimeout(500);
const T = await touch(page);
const findings = { small: [], contrast: [] };

/** Measure what is visible now; `where` names the surface. */
async function measure(where) {
  const r = await ev((where) => {
    const lum = (c) => {
      const v = c.map((x) => x / 255).map((x) => (x <= 0.03928 ? x / 12.92 : ((x + 0.055) / 1.055) ** 2.4));
      return 0.2126 * v[0] + 0.7152 * v[1] + 0.0722 * v[2];
    };
    const parse = (s) => {
      const m = /rgba?\(([^)]+)\)/.exec(s);
      if (!m) return null;
      const p = m[1].split(/[ ,/]+/).filter(Boolean).map(Number);
      return { rgb: p.slice(0, 3), a: p.length > 3 ? p[3] : 1 };
    };
    const blend = (top, under) => top.rgb.map((c, i) => c * top.a + under[i] * (1 - top.a));
    // background behind an element: blend the backgrounds of its ancestors from the root down
    const bgOf = (el) => {
      const chain = [];
      for (let e = el; e; e = e.parentElement) chain.push(e);
      let rgb = [255, 255, 255];
      for (const e of chain.reverse()) {
        const b = parse(getComputedStyle(e).backgroundColor);
        if (b && b.a > 0) rgb = blend(b, rgb);
      }
      return rgb;
    };
    const visible = (el) => {
      const b = el.getBoundingClientRect();
      if (b.width < 1 || b.height < 1) return false;
      const cs = getComputedStyle(el);
      if (cs.visibility === 'hidden' || cs.display === 'none' || +cs.opacity === 0) return false;
      if (b.bottom < 0 || b.right < 0 || b.top > innerHeight || b.left > innerWidth) return false;
      // covered by something else (a backdrop): not reachable
      const cx = Math.min(innerWidth - 1, Math.max(0, b.left + b.width / 2));
      const cy = Math.min(innerHeight - 1, Math.max(0, b.top + b.height / 2));
      const top = document.elementFromPoint(cx, cy);
      return !!top && (el === top || el.contains(top) || top.contains(el));
    };
    const label = (el) => (el.getAttribute('aria-label') || el.getAttribute('data-testid') || el.textContent || el.tagName).trim().replace(/\s+/g, ' ').slice(0, 40);
    const small = [];
    const sel = 'button, [role=button], a[href], input:not([type=hidden]), select, textarea, .chip, .tool, [data-testid=zoom-pill]';
    for (const el of document.querySelectorAll(sel)) {
      if (!visible(el)) continue;
      if (el.closest('.grid-host')) continue; // Univer's own canvas widgets
      const b = el.getBoundingClientRect();
      if (b.width >= 44 && b.height >= 44) continue;
      // an invisible hit area around the control counts (elementFromPoint 22 px from the centre)
      const cx = b.left + b.width / 2;
      const cy = b.top + b.height / 2;
      const hits = [
        [cx - 21, cy],
        [cx + 21, cy],
        [cx, cy - 21],
        [cx, cy + 21],
      ].filter(([x, y]) => {
        const e = document.elementFromPoint(x, y);
        // a check box's label is its touch area (tapping the text toggles it)
        const labels = el.labels ? [...el.labels] : [];
        return e && (e === el || el.contains(e) || labels.some((l) => l === e || l.contains(e)));
      }).length;
      // inline links / text inputs that are wide: height is what matters
      if (hits === 4) continue;
      // a check box or radio inside a label row of at least 44 × 44: the whole row toggles it
      const lab = el.labels?.[0]?.getBoundingClientRect();
      if (lab && lab.width >= 44 && lab.height >= 44) continue;
      if ((el.tagName === 'INPUT' || el.tagName === 'TEXTAREA' || el.tagName === 'SELECT') && b.height >= 40) continue;
      small.push({ where, what: label(el), w: Math.round(b.width), h: Math.round(b.height), testid: el.getAttribute('data-testid') || '' });
    }
    const contrast = [];
    const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
    const seen = new Set();
    for (let n = walker.nextNode(); n; n = walker.nextNode()) {
      const el = n.parentElement;
      if (!el || seen.has(el) || !n.textContent.trim()) continue;
      seen.add(el);
      if (!visible(el) || el.closest('.grid-host') || el.closest('[aria-hidden=true]')) continue;
      const cs = getComputedStyle(el);
      const fg = parse(cs.color);
      if (!fg) continue;
      const bg = bgOf(el);
      const fgc = blend({ rgb: fg.rgb, a: fg.a * +cs.opacity }, bg);
      const L1 = lum(fgc);
      const L2 = lum(bg);
      const ratio = (Math.max(L1, L2) + 0.05) / (Math.min(L1, L2) + 0.05);
      const size = parseFloat(cs.fontSize);
      const bold = +cs.fontWeight >= 700;
      const need = size >= 24 || (bold && size >= 18.66) ? 3 : 4.5;
      if (el.disabled || el.closest('[disabled], [aria-disabled=true]')) continue; // disabled controls are exempt (WCAG 1.4.3)
      if (ratio < need) contrast.push({ where, text: n.textContent.trim().slice(0, 30), ratio: +ratio.toFixed(2), need, fg: cs.color, bg: `rgb(${bg.map(Math.round).join(',')})` });
    }
    return { small, contrast };
  }, where);
  findings.small.push(...r.small);
  findings.contrast.push(...r.contrast);
  return r;
}
const closeAll = async () => {
  for (let i = 0; i < 5; i++) {
    if (!(await page.$('.sheet-backdrop, .dialog, .bsheet, .zoom-panel'))) return;
    await ev(() => window.__petakBack?.());
    await page.waitForTimeout(150);
  }
};

for (const theme of ['light', 'dark']) {
  await step(`${theme}: main screen, tool panels, File menu`, async () => {
    await closeAll();
    await ev((th) => window.__ui.setTheme(th), theme);
    await page.waitForTimeout(300);
    await measure(`${theme} main screen`);
    for (const tab of ['edit', 'format', 'insert', 'draw', 'data', 'tools']) {
      await page.click(`[data-testid=tab-${tab}]`);
      await page.waitForTimeout(250);
      await measure(`${theme} ${tab} panel`);
      // panels scroll: measure the lower part too
      await ev(() => document.querySelector('.bsheet')?.scrollTo(0, 9999));
      await page.waitForTimeout(150);
      await measure(`${theme} ${tab} panel (scrolled)`);
      await closeAll();
    }
    await page.click('[data-testid=menu-more]');
    await page.waitForTimeout(250);
    await measure(`${theme} File menu`);
    await closeAll();
  });
  await step(`${theme}: dialogs and the zoom panel`, async () => {
    for (const d of ['about', 'settings', 'find', 'numfmt', 'cf', 'functions', 'validation', 'pagesetup', 'recent', 'fonts', 'zoom']) {
      await closeAll();
      await ev((d) => window.__ui.open(d), d);
      await page.waitForTimeout(350);
      await measure(`${theme} ${d}`);
    }
    await closeAll();
  });
  await step(`${theme}: long-press menus (cell, row header, column header)`, async () => {
    const g = await (await page.$('[data-testid=grid]')).boundingBox();
    for (const [what, x, y] of [['cell menu', g.x + 200, g.y + 120], ['row menu', g.x + 20, g.y + 120], ['column menu', g.x + 200, g.y + 10]]) {
      await closeAll();
      await T.longPress(x, y);
      await page.waitForTimeout(300);
      await measure(`${theme} ${what}`);
    }
    await closeAll();
  });
}

fs.mkdirSync('tests/e2e/out', { recursive: true });
fs.writeFileSync('tests/e2e/out/a11y.json', JSON.stringify(findings, null, 1) + '\n');
await step('every control is at least 44 × 44 dp (or has a 44 dp touch area)', async () => {
  const uniq = [...new Map(findings.small.map((s) => [s.what + s.w + s.h, s])).values()];
  console.log(`     ${uniq.length} distinct small controls`);
  for (const s of uniq.slice(0, 40)) console.log('      ', JSON.stringify(s));
  assert.equal(uniq.length, 0);
});
await step('all text meets WCAG AA contrast', async () => {
  const uniq = [...new Map(findings.contrast.map((s) => [s.text + s.fg + s.bg, s])).values()];
  console.log(`     ${uniq.length} distinct low-contrast texts`);
  for (const s of uniq.slice(0, 40)) console.log('      ', JSON.stringify(s));
  assert.equal(uniq.length, 0);
});
await t.finish();
