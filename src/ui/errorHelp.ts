/** What an error value in a cell means and what usually fixes it (M-06). Pure: unit-tested. */
export const ERROR_HELP: Record<string, { title: string; text: string }> = {
  '#CYCLE!': {
    title: 'Circular reference',
    text: 'This formula depends on its own result, directly or through other cells, so it can never finish calculating. Change one of the formulas in the loop so that it no longer refers back.',
  },
  '#DIV/0!': { title: 'Division by zero', text: 'The formula divides by zero or by an empty cell. Check the divisor, or wrap the formula: =IFERROR(A1/B1, 0).' },
  '#NAME?': { title: 'Unknown name', text: 'The formula uses a function or name that does not exist (often a typing mistake, or a name that was deleted). Check the spelling.' },
  '#REF!': { title: 'Invalid reference', text: 'A cell the formula referred to was deleted, or the reference points outside the sheet. Edit the formula to point at existing cells.' },
  '#VALUE!': { title: 'Wrong kind of value', text: 'An argument has the wrong type, for example text where a number is needed. Check the cells the formula uses.' },
  '#N/A': { title: 'Not available', text: 'A lookup (VLOOKUP, XLOOKUP, MATCH…) did not find the value it was looking for, or a value is missing.' },
  '#NUM!': { title: 'Invalid number', text: 'A calculation gave a number that is too large or impossible (for example the square root of a negative number), or an iteration did not converge.' },
  '#NULL!': { title: 'Empty intersection', text: 'Two ranges separated by a space do not overlap. Use a comma or a colon between ranges.' },
  '#SPILL!': { title: 'Cannot spill', text: 'The formula returns several values but the cells it needs are not empty. Clear the cells below / to the right.' },
  '#ERROR!': { title: 'Formula error', text: 'The formula could not be read. Check brackets, commas and quotes.' },
};
export const errorOf = (v: unknown): string | null => (typeof v === 'string' && ERROR_HELP[v.trim().toUpperCase()] ? v.trim().toUpperCase() : null);
