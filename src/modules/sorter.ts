/**
 * Sorting Module
 * 
 * Sorts data based on Column A values using numeric comparison.
 * 
 * VBA Equivalent Logic:
 * - Sorts by Column A in ascending order
 * - Uses numeric comparison (not string comparison)
 * - Preserves original order for equal values (stable sort)
 * 
 * Important: Numbers like 1, 2, 10, 20 sort as 1, 2, 10, 20
 * NOT as 1, 10, 2, 20 (string sort)
 */

import { RowData } from '../types';
import { debugLog } from '../utils/debug';

/**
 * Check if a value is numeric
 */
export function isNumeric(value: unknown): boolean {
  if (typeof value === 'number') return true;
  if (typeof value === 'string') {
    const trimmed = value.trim();
    return trimmed !== '' && !isNaN(Number(trimmed));
  }
  return false;
}

/**
 * Convert a value to number for comparison
 * Returns the numeric value or NaN if not convertible
 */
export function toNumber(value: unknown): number {
  if (typeof value === 'number') return value;
  if (typeof value === 'string') {
    const trimmed = value.trim();
    const num = Number(trimmed);
    return isNaN(num) ? NaN : num;
  }
  return NaN;
}

/**
 * Compare two values for sorting
 * - Numeric values sort numerically
 * - Non-numeric values sort as strings
 * - null/undefined sort to the end
 */
export function compareValues(a: unknown, b: unknown): number {
  const numA = toNumber(a);
  const numB = toNumber(b);
  
  // Both are numeric - compare as numbers
  if (!isNaN(numA) && !isNaN(numB)) {
    return numA - numB;
  }
  
  // Only A is numeric - numeric values come first
  if (!isNaN(numA)) return -1;
  
  // Only B is numeric - numeric values come first
  if (!isNaN(numB)) return 1;
  
  // Neither is numeric - compare as strings (case-insensitive)
  const strA = String(a ?? '').toLowerCase();
  const strB = String(b ?? '').toLowerCase();
  return strA.localeCompare(strB);
}

/**
 * Sort data by Column A using numeric comparison
 * 
 * @param data - The data array to sort
 * @param columnAKey - The key for Column A (default: 'Column A')
 * @returns New sorted array (does not mutate original)
 */
export function sortData(
  data: RowData[],
  columnAKey: string = 'Column A'
): RowData[] {
  debugLog('Sorter', `Sorting ${data.length} rows by Column A: "${columnAKey}"`);
  
  const sorted = [...data].sort((rowA, rowB) => {
    const valueA = rowA[columnAKey];
    const valueB = rowB[columnAKey];
    
    return compareValues(valueA, valueB);
  });
  
  // Log first few sorted values for debugging
  const sampleSize = Math.min(5, sorted.length);
  const sample = sorted.slice(0, sampleSize).map(row => row[columnAKey]);
  debugLog('Sorter', `First ${sampleSize} sorted values: ${sample.join(', ')}`);
  
  return sorted;
}

/**
 * Sort data by multiple columns (for tie-breaking)
 * 
 * @param data - The data array to sort
 * @param configs - Array of { column, direction } configurations
 */
export function sortByMultipleColumns(
  data: RowData[],
  configs: { column: string; direction: 'asc' | 'desc' }[]
): RowData[] {
  debugLog('Sorter', `Sorting by ${configs.length} columns`);
  
  return [...data].sort((rowA, rowB) => {
    for (const config of configs) {
      const valueA = rowA[config.column];
      const valueB = rowB[config.column];
      
      let result = compareValues(valueA, valueB);
      
      // Reverse for descending
      if (config.direction === 'desc') {
        result = -result;
      }
      
      // If not equal, return this result
      if (result !== 0) {
        return result;
      }
    }
    
    return 0; // All columns equal
  });
}
