/**
 * Excel's opening rules (tests/unit/excel-rules.ts) for the e2e suites: problems Excel would report
 * for a saved file (empty = it opens without "We found a problem with some content").
 *   node tests/e2e/excelRulesCli.mjs saved.xlsx
 */
import fs from 'node:fs';
import { excelProblems } from '../unit/excel-rules.ts';

export const excelProblemsOf = (file) => excelProblems(fs.readFileSync(file));

if (process.argv[1] && import.meta.url.endsWith(process.argv[1].split('/').pop())) {
  const p = excelProblemsOf(process.argv[2]);
  console.log(p.length ? p.join('\n') : 'no problems');
  process.exit(p.length ? 1 : 0);
}
