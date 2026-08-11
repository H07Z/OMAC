/**
 * Data Processor Module
 * 
 * Central processing pipeline that orchestrates all data transformations.
 * 
 * Pipeline Flow:
 * 1. Clean data (trim, convert types)
 * 2. Remove excluded rows (hierarchy markers in Column B)
 * 3. Remove rows flagged by Notes column (Column C)
 * 4. Sort data (numeric by Column A)
 * 5. Fill skipped code numbers (gap-filling — NOT sequential renumbering)
 * 6. Strip notes column from output
 * 
 * The Notes column (typically Column C / 3rd column) is used 
 * for processing decisions but is EXCLUDED from the final output.
 * Only Column A (number) and Column B (description) appear in output.
 * 
 * Column A values are the REAL codes from the source file. If a code
 * is skipped in the source (e.g. 37 missing between 36 and 38), it is
 * inserted with a blank description so the exported sequence has no
 * gaps. Existing codes are never renumbered or overwritten.
 */

import { RowData, ProcessingStats, LogEntry } from '../types';
import { debugLog } from '../utils/debug';
import { cleanData } from './cleaner';
import { removeExcludedRows, getRemovalSummary } from './rowRemover';
import { sortData } from './sorter';
import { fillSequenceGaps, verifyNoGaps } from './renumberer';

export interface ProcessingConfig {
  columnAKey: string;       // Number column (typically first column)
  columnBKey: string;       // Description column (typically second column)
  notesColumnKey: string;   // Notes column (typically third column) — excluded from output
  excludedColumns: string[]; // All columns to exclude from output (notes + any others)
}

export interface ProcessingResult {
  processedData: RowData[];
  stats: ProcessingStats;
  logs: LogEntry[];
  outputHeaders: string[];  // Only the columns that appear in output (no notes)
}

/**
 * Process raw Excel data through the complete pipeline.
 * 
 * The notes column is used during processing for row filtering
 * decisions, but is stripped from the final output. The output
 * only contains the number column (A) and description column (B),
 * plus any other non-excluded data columns.
 * 
 * @param data - Original imported data (will not be mutated)
 * @param allHeaders - All headers from the imported sheet
 * @param config - Processing configuration (column keys, exclusions)
 * @returns ProcessingResult containing processed data (without notes) and statistics
 */
export function processData(
  data: RowData[],
  allHeaders: string[],
  config: ProcessingConfig
): ProcessingResult {
  const logs: LogEntry[] = [];
  const startTime = Date.now();
  
  const addLog = (level: LogEntry['level'], message: string) => {
    const entry: LogEntry = { timestamp: new Date(), level, message };
    logs.push(entry);
    debugLog('Processor', message);
  };
  
  addLog('info', `Starting processing of ${data.length} rows`);
  addLog('info', `Column A (number): "${config.columnAKey}"`);
  addLog('info', `Column B (description): "${config.columnBKey}"`);
  addLog('info', `Notes column: "${config.notesColumnKey}" (excluded from output)`);
  
  // Determine output headers — exclude notes column and other excluded columns
  const outputHeaders = allHeaders.filter(h => !config.excludedColumns.includes(h));
  addLog('info', `Output columns: ${outputHeaders.join(', ')}`);
  addLog('info', `Excluded columns: ${config.excludedColumns.join(', ')}`);
  
  // Step 1: Clean Data
  addLog('info', 'Step 1: Cleaning data...');
  let processed = cleanData(data);
  addLog('success', `✓ Data cleaned (${processed.length} rows)`);
  
  // Step 2: Remove Excluded Rows based on Column B hierarchy markers
  addLog('info', 'Step 2: Removing excluded hierarchy rows from Column B...');
  const removalSummary = getRemovalSummary(processed, config.columnBKey);
  for (const { pattern, count } of removalSummary) {
    if (count > 0) {
      addLog('warning', `  ${pattern}: ${count} rows`);
    }
  }
  
  const { filteredData, removedCount } = removeExcludedRows(processed, config.columnBKey);
  processed = filteredData;
  
  if (removedCount > 0) {
    addLog('warning', `✓ Removed ${removedCount} hierarchy marker rows`);
  } else {
    addLog('success', '✓ No hierarchy marker rows found');
  }
  
  // Step 3: Remove rows flagged by Notes column
  // Notes column values that indicate a row should be excluded
  let notesRemovedCount = 0;
  if (config.notesColumnKey && allHeaders.includes(config.notesColumnKey)) {
    addLog('info', 'Step 3: Processing Notes column rules...');
    
    const beforeCount = processed.length;
    
    // Check for notes-based exclusion patterns
    // Common note values that indicate exclusion:
    //   "exclude", "remove", "delete", "skip", "hide", "header", "section"
    // Also check if notes column has the same hierarchy markers
    processed = processed.filter(row => {
      const noteVal = row[config.notesColumnKey];
      if (noteVal === null || noteVal === undefined) return true;
      
      const noteStr = String(noteVal).trim().toUpperCase();
      
      // If the notes column contains hierarchy markers, exclude those rows too
      const hierarchyPatterns = ['(NET)', '(SUBNET)', '(SUB-SUBNET)', '(SUB-SUB-SUBNET)'];
      if (hierarchyPatterns.includes(noteStr)) {
        return false;
      }
      
      return true;
    });
    
    notesRemovedCount = beforeCount - processed.length;
    
    if (notesRemovedCount > 0) {
      addLog('warning', `✓ Removed ${notesRemovedCount} additional rows based on Notes column`);
    } else {
      addLog('success', '✓ No additional rows removed by Notes column rules');
    }
  } else {
    addLog('info', 'Step 3: No Notes column configured, skipping...');
  }
  
  // Step 4: Sort Data by Column A (numeric)
  addLog('info', 'Step 4: Sorting data by Column A...');
  processed = sortData(processed, config.columnAKey);
  addLog('success', '✓ Data sorted numerically');
  
  // Step 5: Fill missing code numbers (gap-filling, NOT renumbering)
  // If the source skips a number — e.g. 37 between 36 and 38 — that
  // number is inserted with a blank description so the final Column A
  // sequence has no gaps. Existing codes are never overwritten.
  addLog('info', 'Step 5: Checking for skipped codes in Column A...');
  const { data: gapFilled, insertedCodes } = fillSequenceGaps(processed, config.columnAKey, config.columnBKey);
  processed = gapFilled;
  
  if (insertedCodes.length > 0) {
    // Compact consecutive runs into ranges, e.g. 1-1000, 37, 59-60
    const ranges: string[] = [];
    let rangeStart = insertedCodes[0];
    let rangeEnd = insertedCodes[0];
    for (let i = 1; i <= insertedCodes.length; i++) {
      const code = insertedCodes[i];
      if (code === rangeEnd + 1) {
        rangeEnd = code;
      } else {
        ranges.push(rangeStart === rangeEnd ? `${rangeStart}` : `${rangeStart}–${rangeEnd}`);
        rangeStart = code;
        rangeEnd = code;
      }
    }
    addLog('warning', `✓ Auto-inserted ${insertedCodes.length} missing code(s): ${ranges.join(', ')}`);
  } else {
    addLog('success', '✓ No skipped codes found — sequence already complete');
  }
  
  const noGaps = verifyNoGaps(processed, config.columnAKey);
  if (noGaps) {
    addLog('success', '✓ Column A sequence verified — no gaps remain');
  } else {
    addLog('error', '✗ Gap verification failed');
  }
  
  // Step 5.5: Populate existing blank descriptions with their code number
  addLog('info', 'Step 6: Checking for blank descriptions...');
  let filledBlankCount = 0;
  processed = processed.map(row => {
    const valB = row[config.columnBKey];
    if (valB === null || valB === undefined || String(valB).trim() === '') {
      filledBlankCount++;
      return { ...row, [config.columnBKey]: row[config.columnAKey] };
    }
    return row;
  });
  if (filledBlankCount > 0) {
    addLog('success', `✓ Populated ${filledBlankCount} blank description(s) with their Column A code number`);
  } else {
    addLog('success', '✓ No blank descriptions found in original rows');
  }
  
  // Step 7: Strip excluded columns from output
  addLog('info', 'Step 6: Removing Notes column from output...');
  processed = processed.map(row => {
    const outputRow: RowData = {};
    for (const header of outputHeaders) {
      outputRow[header] = row[header] !== undefined ? row[header] : null;
    }
    return outputRow;
  });
  addLog('success', `✓ Notes column stripped — output has ${outputHeaders.length} columns`);
  
  // Calculate stats
  const totalRemoved = removedCount + notesRemovedCount;
  const stats: ProcessingStats = {
    originalRows: data.length,
    removedRows: totalRemoved,
    insertedRows: insertedCodes.length,
    finalRows: processed.length
  };
  
  const elapsed = Date.now() - startTime;
  addLog('success', `Processing complete in ${elapsed}ms`);
  addLog('success', `Final dataset: ${stats.finalRows} rows × ${outputHeaders.length} columns`);
  
  return { processedData: processed, stats, logs, outputHeaders };
}

/**
 * Get a summary of what will be removed (preview before processing)
 */
export function previewRemovals(
  data: RowData[],
  columnBKey: string = 'Column B'
): { pattern: string; count: number }[] {
  return getRemovalSummary(data, columnBKey);
}

/**
 * Auto-detect column roles from headers.
 * 
 * Heuristics:
 * - First column → Column A (numbers)
 * - Second column → Column B (descriptions/data)
 * - Third column → Notes (excluded from output)
 * 
 * Also checks for common column names like "notes", "comments", "remarks", etc.
 */
export function autoDetectColumns(headers: string[]): ProcessingConfig {
  // Default fallbacks: first = A, second = B, third = notes
  let columnAKey = headers[0] || '';
  let columnBKey = headers[1] || '';
  let notesColumnKey = headers.length >= 3 ? headers[2] : '';
  
  // Try to detect by common names
  for (const h of headers) {
    const lower = h.toLowerCase().trim();
    
    // Detect Column A
    if (['column a', 'columna', 'no', 'no.', 'number', 'item no', 'item no.', 'id', '#'].includes(lower)) {
      columnAKey = h;
    }
    
    // Detect Column B
    if (['column b', 'columnb', 'description', 'item', 'name', 'data', 'text', 'detail', 'details'].includes(lower)) {
      columnBKey = h;
    }
    
    // Detect Notes column
    if (['notes', 'note', 'comments', 'comment', 'remarks', 'remark', 'annotation', 
         'annotations', 'status', 'action', 'rule', 'rules', 'column c', 'columnc',
         'instructions', 'instruction', 'memo'].includes(lower)) {
      notesColumnKey = h;
    }
  }
  
  // Build excluded columns list
  const excludedColumns: string[] = [];
  if (notesColumnKey) {
    excludedColumns.push(notesColumnKey);
  }
  // Also exclude any columns beyond the first two that aren't explicitly Column A or B
  // This ensures only 2 columns appear in output by default
  for (let i = 2; i < headers.length; i++) {
    if (!excludedColumns.includes(headers[i])) {
      excludedColumns.push(headers[i]);
    }
  }
  
  return { columnAKey, columnBKey, notesColumnKey, excludedColumns };
}
