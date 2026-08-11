/**
 * Sequence Completion Module
 *
 * Ensures Column A forms an unbroken numeric sequence that
 * ALWAYS STARTS AT 1.
 *
 * CONFIRMED BUSINESS RULES (from real output comparison):
 *
 * 1. The sequence always begins at 1. If the smallest code in the
 *    source is e.g. 1001, then codes 1 through 1000 are inserted
 *    first, and only then the known values continue (1001, 1002, …).
 *
 * 2. If the source skips a code — e.g. 37 missing between 36 and 38,
 *    or 59-60 missing between 58 and 61 — those numbers are inserted
 *    so the final sequence has no gaps.
 *
 * Inserted rows carry the code number in BOTH Column A and Column B
 * (the description falls back to the code number itself).
 *
 * Real code numbers from the source are preserved exactly as-is;
 * nothing is ever renumbered or overwritten.
 */

import { RowData } from '../types';
import { debugLog } from '../utils/debug';

export interface GapFillResult {
  data: RowData[];
  insertedCodes: number[];
}

/**
 * Fill missing integers from 1 up to the highest numeric
 * Column A value in the dataset.
 *
 * The sequence always starts at 1 — leading codes are backfilled
 * before the first known value (e.g. smallest code 1001 → insert
 * 1…1000 first), and interior gaps are backfilled as encountered.
 *
 * Existing rows are left completely untouched (values, order).
 * Missing codes get a brand-new row where both Column A and
 * Column B hold the missing number.
 * Non-numeric Column A rows (if any slipped through) are preserved
 * and appended after the numeric run, unchanged.
 *
 * @param data        Rows already sorted ascending by columnAKey
 * @param columnAKey  Key holding the code number
 * @param columnBKey  Key holding the description (code number on inserted rows)
 */
export function fillSequenceGaps(
  data: RowData[],
  columnAKey: string,
  columnBKey: string
): GapFillResult {
  const numericRows = data.filter(row => typeof row[columnAKey] === 'number');
  const nonNumericRows = data.filter(row => typeof row[columnAKey] !== 'number');

  debugLog('SequenceFiller', `Checking ${numericRows.length} numeric rows for gaps`);

  if (numericRows.length === 0) {
    debugLog('SequenceFiller', 'No numeric Column A values found — nothing to fill');
    return { data, insertedCodes: [] };
  }

  const result: RowData[] = [];
  const insertedCodes: number[] = [];

  // The sequence always starts at 1 — if the first known code is
  // higher (e.g. 1001), every number before it (1…1000) is inserted
  // before the known values are emitted.
  let expected = 1;

  for (const row of numericRows) {
    const current = row[columnAKey] as number;

    // Insert a placeholder for every missing number before `current`
    while (expected < current) {
      // Prompt requirement: if Column B is missing, populate it with Column A's value
      result.push({ [columnAKey]: expected, [columnBKey]: expected });
      insertedCodes.push(expected);
      expected++;
    }

    result.push(row);
    expected = current + 1;
  }

  if (insertedCodes.length > 0) {
    debugLog('SequenceFiller', `Inserted ${insertedCodes.length} missing code(s): ${insertedCodes.join(', ')}`);
  } else {
    debugLog('SequenceFiller', 'No gaps found — sequence already complete');
  }

  return { data: [...result, ...nonNumericRows], insertedCodes };
}

/**
 * Verify the numeric Column A sequence starts at 1 and has no gaps
 * between consecutive values. Non-numeric rows (if any) are ignored.
 */
export function verifyNoGaps(data: RowData[], columnAKey: string): boolean {
  const numeric = data
    .map(row => row[columnAKey])
    .filter((v): v is number => typeof v === 'number');

  if (numeric.length === 0) return true;

  if (numeric[0] !== 1) {
    debugLog('SequenceFiller', `Sequence does not start at 1 (starts at ${numeric[0]})`);
    return false;
  }

  for (let i = 1; i < numeric.length; i++) {
    if (numeric[i] !== numeric[i - 1] + 1) {
      debugLog('SequenceFiller', `Gap detected between ${numeric[i - 1]} and ${numeric[i]}`);
      return false;
    }
  }
  return true;
}
