/** Phase 3 helpers: circular reference path (M-06), sheet name rules (M-10), error help. */
import { describe, expect, it, vi } from 'vitest';
import { FormulaEngine } from '../../src/engine/formulaCore';
import { ERROR_HELP, errorOf } from '../../src/ui/errorHelp';

vi.mock('../../src/state/doc', () => ({ doc: {} }));
const { sheetNameError, hasHeaderRow } = await import('../../src/ui/actions');

describe('circular references (M-06)', () => {
  it('finds the cells of the loop, in order, across sheets', () => {
    const e = new FormulaEngine();
    e.addSheet('s1', 'Sheet1');
    e.addSheet('s2', 'Data');
    e.suspend();
    // A1 = B1 + 1, B1 = Data!A1 * 2, Data!A1 = Sheet1!A1 ; C1 is outside the loop but uses it
    e.loadRows('s1', 0, [['=B1+1', '=Data!A1*2', '=A1+5']]);
    e.loadRows('s2', 0, [['=Sheet1!A1']]);
    e.resumeAll();
    expect(e.cyclePath('s1', 0, 0)).toEqual([['s1', 0, 0], ['s1', 0, 1], ['s2', 0, 0]]);
    expect(e.cyclePath('s1', 0, 2)).toEqual([]);
  });
  it('a formula that refers to itself through a range', () => {
    const e = new FormulaEngine();
    e.addSheet('s1', 'Sheet1');
    e.suspend();
    e.loadRows('s1', 0, [[1], [2], ['=SUM(A1:A3)']]);
    e.resumeAll();
    expect(e.cyclePath('s1', 2, 0)).toEqual([['s1', 2, 0]]);
  });
});

describe('sheet names (M-10)', () => {
  it('follows Excel’s rules', () => {
    expect(sheetNameError('Budget 2026', ['Sheet1'])).toBeNull();
    expect(sheetNameError('  ', [])).toMatch(/empty/);
    expect(sheetNameError('x'.repeat(32), [])).toMatch(/31 characters/);
    for (const c of [':', '\\', '/', '?', '*', '[', ']']) expect(sheetNameError(`a${c}b`, []), c).toMatch(/cannot contain/);
    expect(sheetNameError("'quoted", [])).toMatch(/apostrophe/);
    expect(sheetNameError('history', [])).toMatch(/reserved/);
    expect(sheetNameError('SHEET1', ['Sheet1'])).toMatch(/already/);
  });
});

describe('error help', () => {
  it('knows every error value Petak can show', () => {
    for (const e of ['#CYCLE!', '#DIV/0!', '#NAME?', '#REF!', '#VALUE!', '#N/A', '#NUM!', '#NULL!', '#SPILL!', '#ERROR!']) expect(ERROR_HELP[e]?.text.length, e).toBeGreaterThan(30);
    expect(errorOf('#div/0!')).toBe('#DIV/0!');
    expect(errorOf('hello')).toBeNull();
    expect(errorOf(5)).toBeNull();
  });
});

describe('theme colours come from the file’s own theme (M-08)', async () => {
  const { colorOf, parseTheme } = await import('../../src/io/ooxml/styles');
  const xml = `<a:theme xmlns:a="x"><a:themeElements><a:clrScheme name="Custom">
    <a:dk1><a:sysClr val="windowText" lastClr="111111"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FEFEFE"/></a:lt1>
    <a:dk2><a:srgbClr val="222222"/></a:dk2><a:lt2><a:srgbClr val="EEEEEE"/></a:lt2>
    <a:accent1><a:srgbClr val="8E44AD"/></a:accent1><a:accent2><a:srgbClr val="16A085"/></a:accent2>
    <a:accent3><a:srgbClr val="C0392B"/></a:accent3><a:accent4><a:srgbClr val="F39C12"/></a:accent4>
    <a:accent5><a:srgbClr val="2C3E50"/></a:accent5><a:accent6><a:srgbClr val="27AE60"/></a:accent6>
    <a:hlink><a:srgbClr val="0000EE"/></a:hlink><a:folHlink><a:srgbClr val="551A8B"/></a:folHlink>
  </a:clrScheme></a:themeElements></a:theme>`;
  it('accent and text colours, with tints', () => {
    const t = parseTheme(xml);
    expect(colorOf('<fgColor theme="4"/>', t)).toBe('#8e44ad'); // accent1 of this theme, not Office blue 4472C4
    expect(colorOf('<color theme="1"/>', t)).toBe('#111111');
    expect(colorOf('<color theme="5" tint="0.5"/>', t)).not.toBe('#16a085');
    expect(colorOf('<fgColor theme="4"/>', parseTheme(undefined))).toBe('#4472c4'); // no theme part: Office default
  });
});

describe('sort keeps a header row (found by the Phase 4 tap test)', () => {
  it('guesses headings like Excel', () => {
    const list = [['Item', 'Qty'], ['pear', 3], ['apple', 5]];
    expect(hasHeaderRow(list)).toBe(true); // Qty over numbers
    expect(hasHeaderRow([['Name'], ['pear'], ['apple']])).toBe(false); // all text, nothing to tell
    expect(hasHeaderRow([['Name'], ['pear'], ['apple']], true)).toBe(true); // a bold first row
    expect(hasHeaderRow([[3, 'x'], [1, 'y']])).toBe(false); // a number on top is data
    expect(hasHeaderRow([['Date', 'Note'], [46000, 'a'], [46001, null]])).toBe(true);
    expect(hasHeaderRow([['only']])).toBe(false);
  });
});

