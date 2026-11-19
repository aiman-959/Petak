/**
 * The .ods test files (L-03), made by LibreOffice so they are what users really get:
 *   looks.ods       every cell look Petak shows: fonts, fills, each border style, alignment, wrap,
 *                   rotation, number formats, column widths, row heights, hidden rows / columns, merges
 *   looks-lo.xlsx   LibreOffice's own .xlsx of looks.ods: the reference Petak's .ods reading must match
 *   node tests/fixtures/make-ods.mjs [outDir]          (needs soffice)
 */
import ExcelJS from 'exceljs';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { execFileSync } from 'node:child_process';

const out = path.resolve(process.argv[2] ?? 'tests/fixtures/files');
const black = { argb: 'FF000000' };

export function looks() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Looks');
  ws.getColumn(1).width = 22;
  for (let c = 2; c <= 14; c++) ws.getColumn(c).width = 12 + (c % 4) * 2;
  ws.getColumn(15).hidden = true;
  ws.getColumn(16).width = 40;

  // row 1–2: every border style, all four sides
  const borders = ['thin', 'medium', 'thick', 'hair', 'dotted', 'dashed', 'double', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot'];
  ws.getCell('A1').value = 'Borders';
  borders.forEach((s, i) => {
    const c = ws.getCell(2, 2 + i);
    c.value = s;
    const b = { style: s, color: i % 3 === 0 ? black : { argb: i % 3 === 1 ? 'FFDC2626' : 'FF2563EB' } };
    c.border = { top: b, bottom: b, left: b, right: b };
  });
  // a cell with different sides
  ws.getCell('A2').border = { top: { style: 'thin', color: black }, bottom: { style: 'double', color: black }, left: { style: 'medium', color: { argb: 'FF16A34A' } } };
  ws.getCell('A2').value = 'mixed sides';

  // row 4: alignment, wrap, rotation
  ws.getCell('A4').value = 'Alignment';
  const aligns = [
    { horizontal: 'left' },
    { horizontal: 'center' },
    { horizontal: 'right' },
    { horizontal: 'justify', wrapText: true },
    { vertical: 'top' },
    { vertical: 'middle' },
    { vertical: 'bottom' },
    { horizontal: 'center', vertical: 'middle', wrapText: true },
    { textRotation: 45 },
    { textRotation: 90 },
    { textRotation: -30 },
    { horizontal: 'left', indent: 2 },
  ];
  aligns.forEach((a, i) => {
    const c = ws.getCell(4, 2 + i);
    c.value = a.wrapText ? 'a longer text that wraps in the cell' : JSON.stringify(a).replace(/[{}"]/g, '');
    c.alignment = a;
  });
  ws.getRow(4).height = 60;

  // row 6: fonts
  ws.getCell('A6').value = 'Fonts';
  const fonts = [
    { bold: true },
    { italic: true },
    { underline: true },
    { underline: 'double' },
    { strike: true },
    { color: { argb: 'FFDC2626' } },
    { size: 18 },
    { size: 8 },
    { name: 'Times New Roman' },
    { name: 'Courier New', size: 12 },
    { name: 'Arial', bold: true, italic: true, color: { argb: 'FF2563EB' } },
    { name: 'Calibri', size: 14, underline: true, strike: true },
  ];
  fonts.forEach((f, i) => {
    const c = ws.getCell(6, 2 + i);
    c.value = 'Font ' + (i + 1);
    c.font = f;
  });
  ws.getRow(6).height = 30;

  // row 8: fills
  ws.getCell('A8').value = 'Fills';
  ['FFFDE68A', 'FFBBF7D0', 'FFBFDBFE', 'FF14532D', 'FFFFFFFF', 'FF000000'].forEach((argb, i) => {
    const c = ws.getCell(8, 2 + i);
    c.value = argb.slice(2);
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb } };
    if (i >= 3) c.font = { color: { argb: i === 4 ? 'FF000000' : 'FFFFFFFF' } };
  });
  // a filled, bordered cell without a value
  ws.getCell('H8').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFECACA' } };
  ws.getCell('H8').border = { top: { style: 'thin', color: black }, bottom: { style: 'thin', color: black } };

  // rows 10…: number formats, each with a positive, a negative and a zero
  ws.getCell('A10').value = 'Number formats';
  const numfmts = [
    '0',
    '0.00',
    '#,##0',
    '#,##0.00',
    '0%',
    '0.00%',
    '0.00E+00',
    '# ?/?',
    '$#,##0.00',
    '#,##0.00 [$€-407]',
    '#,##0;[Red]-#,##0',
    '#,##0.00;(#,##0.00);"-"',
    '0.0" kg"',
    '"Total: "0',
    '00000',
    '@',
  ];
  numfmts.forEach((f, i) => {
    const r = 11 + i;
    ws.getCell(r, 1).value = f;
    [1234.5678, -1234.5678, 0].forEach((v, k) => {
      const c = ws.getCell(r, 2 + k);
      c.value = f === '@' ? String(v) : v;
      c.numFmt = f;
    });
  });
  const dates = ['yyyy-mm-dd', 'd-mmm-yy', 'mmmm d, yyyy', 'dddd', 'dd/mm/yyyy hh:mm', 'h:mm AM/PM', 'h:mm:ss', '[h]:mm', 'mm:ss.0', 'mmm-yy'];
  dates.forEach((f, i) => {
    const r = 11 + numfmts.length + i;
    ws.getCell(r, 1).value = f;
    const c = ws.getCell(r, 2);
    c.value = 46345.6283; // 2026-11-19 15:04:45
    c.numFmt = f;
    const d = ws.getCell(r, 3);
    d.value = 1.75; // 42 hours
    d.numFmt = f;
  });
  const last = 11 + numfmts.length + dates.length;
  // row heights, a hidden row, a merge
  ws.getRow(last + 1).height = 45;
  ws.getRow(last + 2).hidden = true;
  ws.getCell(last + 2, 1).value = 'hidden row';
  ws.mergeCells(last + 3, 1, last + 4, 3);
  const m = ws.getCell(last + 3, 1);
  m.value = 'merged and centred';
  m.alignment = { horizontal: 'center', vertical: 'middle' };
  m.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE0E7FF' } };
  m.font = { bold: true, size: 13 };
  ws.getCell('P1').value = 'wide column';

  // a second sheet in another default look: one cell, the rest left as the defaults
  const ws2 = wb.addWorksheet('Plain');
  ws2.getCell('A1').value = 'only defaults';
  ws2.getCell('B2').value = 42;
  return wb;
}

if (process.argv[1] && path.resolve(process.argv[1]) === path.resolve(new URL(import.meta.url).pathname)) {
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'petak-ods-'));
  const xlsx = path.join(tmp, 'looks.xlsx');
  await looks().xlsx.writeFile(xlsx);
  const run = (fmt, f, dir) => execFileSync('soffice', [`-env:UserInstallation=file://${tmp}/profile`, '--headless', '--convert-to', fmt, '--outdir', dir, f], { stdio: 'pipe' });
  run('ods', xlsx, tmp);
  fs.mkdirSync(path.join(tmp, 'back'));
  run('xlsx:Calc MS Excel 2007 XML', path.join(tmp, 'looks.ods'), path.join(tmp, 'back'));
  fs.copyFileSync(path.join(tmp, 'looks.ods'), path.join(out, 'looks.ods'));
  fs.copyFileSync(path.join(tmp, 'back', 'looks.xlsx'), path.join(out, 'looks-lo.xlsx'));
  fs.rmSync(tmp, { recursive: true, force: true });
  console.log('looks.ods, looks-lo.xlsx →', out);
}
