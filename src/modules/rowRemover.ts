/**
 * Row Removal Module
 * 
 * Removes hierarchy marker rows from the dataset.
 * 
 * VBA Equivalent Logic:
 * - Checks Column B for specific hierarchy markers
 * - Removes entire row (not just clearing cells)
 * - Works even when Column A is blank
 * - Case-insensitive, trimmed comparison
 * 
 * Excluded patterns:
 *   (NET)
 *   (SUBNET)
 *   (SUB-SUBNET)
 *   (SUB-SUB-SUBNET)
 * 
 * The comparison is: trim → uppercase → exact match against the list.
 * If the VBA behaviour was different (e.g. "contains" match), 
 * update MATCH_MODE below.
 */

import { RowData } from '../types';
import { debugLog } from '../utils/debug';

/**
 * Hierarchy marker patterns to exclude.
 * Stored in uppercase for case-insensitive comparison.
 */
const EXCLUDED_PATTERNS: readonly string[] = [
  '(NET)',
  '(SUBNET)',
  '(SUB-SUBNET)',
  '(SUB-SUB-SUBNET)',
];

// 'exact'    = value must equal a pattern exactly (after trim/uppercase)
// 'contains' = value contains a pattern anywhere
//
// 'contains' is used because the markers are highly distinctive
// (parenthesised, all-caps) and in real files they frequently appear
// appended to a label, e.g. "TOTAL AWARENESS (NET)". Those rows are
// still hierarchy headings and must not reach the export.
const MATCH_MODE: 'exact' | 'contains' = 'contains';

/**
 * Check if a single cell value matches any excluded pattern.
 */
function matchesExcluded(value: string | number | null | undefined): boolean {
  if (value === null || value === undefined) return false;

  const normalized = String(value).trim().toUpperCase();
  if (normalized === '') return false;

  if (MATCH_MODE === 'exact') {
    return EXCLUDED_PATTERNS.includes(normalized);
  }
  // contains mode
  return EXCLUDED_PATTERNS.some(p => normalized.includes(p));
}

/**
 * Check if a row should be excluded based on Column B value.
 * 
 * @param row - The row data to check
 * @param columnBKey - The key for Column B in the data structure
 * @returns true if the row should be removed
 */
export function isExcludedRow(row: RowData, columnBKey: string): boolean {
  return matchesExcluded(row[columnBKey]);
}

/**
 * Remove excluded hierarchy marker rows from dataset.
 * 
 * Returns a NEW array (does not mutate the original).
 * Also returns the count of removed rows for statistics.
 */
export function removeExcludedRows(
  data: RowData[],
  columnBKey: string = 'Column B'
): { filteredData: RowData[]; removedCount: number } {
  debugLog('RowRemover', `Checking ${data.length} rows against ${EXCLUDED_PATTERNS.length} exclusion patterns`);
  debugLog('RowRemover', `Column B key: "${columnBKey}"`);
  debugLog('RowRemover', `Match mode: ${MATCH_MODE}`);
  
  const filteredData: RowData[] = [];
  let removedCount = 0;
  
  for (let i = 0; i < data.length; i++) {
    if (isExcludedRow(data[i], columnBKey)) {
      removedCount++;
    } else {
      filteredData.push(data[i]);
    }
  }
  
  debugLog('RowRemover', `Result: removed ${removedCount}, kept ${filteredData.length}`);
  return { filteredData, removedCount };
}

/**
 * Create a summary of how many rows match each exclusion pattern.
 * Useful for the processing log.
 */
export function getRemovalSummary(
  data: RowData[],
  columnBKey: string = 'Column B'
): { pattern: string; count: number }[] {
  const counts: Record<string, number> = {};
  for (const p of EXCLUDED_PATTERNS) counts[p] = 0;

  for (const row of data) {
    const val = row[columnBKey];
    if (val === null || val === undefined) continue;
    const normalized = String(val).trim().toUpperCase();

    // Attribute the row to the longest matching pattern so that
    // "(SUB-SUB-SUBNET)" is not miscounted as "(SUBNET)".
    let matched = '';
    for (const pattern of EXCLUDED_PATTERNS) {
      const hit = MATCH_MODE === 'exact'
        ? normalized === pattern
        : normalized.includes(pattern);
      if (hit && pattern.length > matched.length) {
        matched = pattern;
      }
    }

    if (matched) counts[matched]++;
  }

  return Object.entries(counts)
    .filter(([, c]) => c > 0)
    .map(([pattern, count]) => ({ pattern, count }));
}
