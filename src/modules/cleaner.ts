/**
 * Data Cleaning Module
 * 
 * Cleans imported Excel data by:
 * - Trimming whitespace from all string values
 * - Converting numbers stored as strings to native numbers (only for pure numeric values)
 * - Preserving empty cells as null
 * - Preserving text descriptions as strings (not converting to numbers)
 */

import { RowData } from '../types';
import { debugLog } from '../utils/debug';

/**
 * Clean a single cell value
 * - Trims whitespace from strings
 * - Converts purely numeric strings to numbers (e.g. "123" → 123)
 * - Does NOT convert strings that merely start with digits (e.g. "10 apples" stays as string)
 * - Handles null/undefined → null
 * - Empty string → null
 */
export function cleanCellValue(value: string | number | null | undefined): string | number | null {
  if (value === null || value === undefined) {
    return null;
  }
  
  if (typeof value === 'number') {
    return value;
  }
  
  const trimmed = String(value).trim();
  
  if (trimmed === '') {
    return null;
  }
  
  // Only convert to number if the ENTIRE string is a valid number
  // Match: optional negative, digits, optional decimal, optional digits
  if (/^-?\d+(\.\d+)?$/.test(trimmed)) {
    const num = Number(trimmed);
    if (!isNaN(num) && isFinite(num)) {
      return num;
    }
  }
  
  return trimmed;
}

/**
 * Clean all values in a row
 */
export function cleanRow(row: RowData): RowData {
  const cleaned: RowData = {};
  
  for (const [key, value] of Object.entries(row)) {
    cleaned[key] = cleanCellValue(value);
  }
  
  return cleaned;
}

/**
 * Clean entire dataset.
 * Creates a new array — does NOT mutate original.
 */
export function cleanData(data: RowData[]): RowData[] {
  debugLog('Cleaner', `Cleaning ${data.length} rows`);
  
  const cleaned = data.map(row => cleanRow(row));
  
  debugLog('Cleaner', 'Data cleaning complete');
  return cleaned;
}
