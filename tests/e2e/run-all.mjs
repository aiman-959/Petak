/**
 * Runs every e2e suite one after another against BASE_URL and prints one line per suite.
 *   npx vite build && npx vite preview --port 4173 &
 *   BASE_URL=http://localhost:4173/ node tests/e2e/run-all.mjs
 * Writes tests/e2e/out/summary.json. Exit code 1 when a suite failed.
 */
import { spawnSync } from 'node:child_process';
import fs from 'node:fs';

const SUITES = ['cells', 'cells2', 'draw-touch', 'part2', 'part3', 'part4', 'fonts', 'files', 'render', 'features', 'phase3', 'commands', 'roundtrip', 'taps'];
const only = process.argv.slice(2);
const summary = [];
for (const s of SUITES.filter((x) => !only.length || only.includes(x))) {
  const t0 = Date.now();
  const r = spawnSync('node', [`tests/e2e/${s}.mjs`], { encoding: 'utf8', env: process.env, timeout: 30 * 60 * 1000 });
  const out = (r.stdout ?? '') + (r.stderr ?? '');
  const m = /(\d+)\/(\d+) passed/.exec(out);
  const failed = [...out.matchAll(/✖ (.*)/g)].map((x) => x[1]);
  const line = { suite: s, passed: m ? +m[1] : 0, total: m ? +m[2] : 0, seconds: Math.round((Date.now() - t0) / 1000), failed, ok: r.status === 0 && !!m && m[1] === m[2] };
  summary.push(line);
  console.log(`${line.ok ? '✔' : '✖'} ${s}: ${line.passed}/${line.total} (${line.seconds} s)${failed.length ? '\n    ' + failed.join('\n    ') : ''}${!m ? '\n' + out.slice(-1500) : ''}`);
}
fs.mkdirSync('tests/e2e/out', { recursive: true });
fs.writeFileSync('tests/e2e/out/summary.json', JSON.stringify(summary, null, 1) + '\n');
process.exit(summary.every((x) => x.ok) ? 0 : 1);
