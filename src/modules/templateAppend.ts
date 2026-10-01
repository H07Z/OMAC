/**
 * Template Append Processor
 *
 * Browser equivalent of VBA CopyAndAppend_ByTemplate_Columns.
 *
 * Template format:
 *   Column A = output column header
 *   Columns B..N = source headers to find in the uploaded source file
 *
 * For each template row, the processor creates one output column. It finds
 * every mapped source header in source row 1, copies that source column from
 * row 2 down to the last used row in that source column, and appends those
 * values under the output column header.
 *
 * Auto-generated / Special Columns:
 * 1. SERVED (Column 1):
 *    - rows copied from the 1st mapped source header get SERVED = 1
 *    - rows copied from the 2nd mapped source header get SERVED = 2
 *    - and so on.
 * 2. Serial (Column 2 or mapped position):
 *    - Sequential number per row from 1 up to total rows of data (1..N).
 *    - If Serial is on the mapping, it is populated with 1..N.
 *    - If there was no Serial on the map, it is automatically added.
 * 3. PARITY and comparison source columns (appended):
 *    - For the selected mapping row, compares source mapping cells B and C
 *      (the first two source headers), e.g. I_1_Q18 vs I_2_Q18.
 *    - If values match (TRUE), coded as 1.
 *    - If values do not match (FALSE), coded as 2.
 *    - Both compared source fields are the final two output columns, using
 *      their original headers and the same respondent alignment as PARITY.
 */

import * as XLSX from 'xlsx';

export type CellValue = string | number | null;

export interface TemplateMappingRow {
  targetHeader: string;
  sourceHeaders: string[];
}

export interface AppendOptions {
  /** Output mapping row whose first two mapped source fields are compared. */
  parityMappingHeader?: string;
}

export interface AppendResult {
  headers: string[];
  rows: CellValue[][];
  log: string[];
  missingHeaders: string[];
  outputColumnLengths: number[];
  servedSequence: number[];
  parityMappingHeader?: string;
  parityCol1?: string;
  parityCol2?: string;
  parityStats?: {
    matches: number;
    mismatches: number;
  };
}

function normalizeCell(value: unknown): CellValue {
  if (value === null || value === undefined) return null;
  if (typeof value === 'number') return value;
  const str = String(value).trim();
  return str === '' ? null : str;
}

function toSearchKey(value: unknown): string {
  return String(value ?? '').trim();
}

/** Check if two cell values match for PARITY comparison. */
export function checkParityMatch(val1: CellValue, val2: CellValue): boolean {
  if (val1 === null || val1 === undefined || val2 === null || val2 === undefined) {
    return false;
  }
  const s1 = String(val1).trim();
  const s2 = String(val2).trim();
  if (s1 === '' || s2 === '') {
    return false;
  }
  const n1 = Number(s1);
  const n2 = Number(s2);
  if (!isNaN(n1) && !isNaN(n2)) {
    return n1 === n2;
  }
  return s1.toLowerCase() === s2.toLowerCase();
}

/** Parse a pasted Excel range (tab/newline separated) into template rows. */
export function parseTemplateText(text: string): TemplateMappingRow[] {
  return text
    .split(/\r?\n/)
    .map(line => line.split('\t').map(cell => cell.trim()))
    .filter(cells => cells.some(Boolean))
    .map(cells => ({
      targetHeader: cells[0] ?? '',
      sourceHeaders: cells.slice(1).filter(Boolean),
    }))
    .filter(row => row.targetHeader !== '');
}

/** Convert an XLSX worksheet to a 2D array preserving empty cells. */
export function worksheetToMatrix(worksheet: XLSX.WorkSheet): CellValue[][] {
  const rows = XLSX.utils.sheet_to_json<(string | number | null)[]>(worksheet, {
    header: 1,
    defval: null,
    raw: false,
  });

  return rows.map(row => row.map(normalizeCell));
}

/** Parse a template sheet where Column A is target and B..N are source headers. */
export function parseTemplateMatrix(matrix: CellValue[][]): TemplateMappingRow[] {
  return matrix
    .filter(row => row.some(cell => cell !== null && String(cell).trim() !== ''))
    .map(row => ({
      targetHeader: toSearchKey(row[0]),
      sourceHeaders: row.slice(1).map(toSearchKey).filter(Boolean),
    }))
    .filter(row => row.targetHeader !== '');
}

/**
 * Execute the append-by-template operation.
 * Supports:
 * - SERVED column as column 1 (1 = 1st amend, 2 = 2nd amend, ...)
 * - Serial column (1..N sequence; auto-added if not in mapping)
 * - PARITY column (compares 2 selected columns: 1 = TRUE, 2 = FALSE)
 * - Both PARITY source fields appended as the final two output columns
 */
export function appendByTemplate(
  sourceMatrix: CellValue[][],
  mappings: TemplateMappingRow[],
  options: AppendOptions = {},
): AppendResult {
  const log: string[] = [];
  const missing = new Set<string>();

  if (sourceMatrix.length === 0) {
    return {
      headers: [],
      rows: [],
      log: ['Source worksheet is empty.'],
      missingHeaders: [],
      outputColumnLengths: [],
      servedSequence: [],
    };
  }

  const sourceHeaders = (sourceMatrix[0] ?? []).map(toSearchKey);
  const sourceRows = sourceMatrix.slice(1);

  // SERVED is generated, so ignore an accidental user-supplied SERVED mapping
  const effectiveMappings = mappings.filter(
    mapping => mapping.targetHeader.trim().toUpperCase() !== 'SERVED'
  );
  if (effectiveMappings.length !== mappings.length) {
    log.push('Ignored template row "SERVED" because SERVED is generated automatically as column 1.');
  }

  // Check if Serial was specified in the mapping
  const serialMappingIndex = effectiveMappings.findIndex(
    m => m.targetHeader.trim().toUpperCase() === 'SERIAL'
  );
  const hasSerialInMapping = serialMappingIndex !== -1;

  const outputColumns: CellValue[][] = [];
  const servedByColumn: Array<Array<number | null>> = [];
  const outputHeaders: string[] = [];

  for (let mIdx = 0; mIdx < effectiveMappings.length; mIdx++) {
    const mapping = effectiveMappings[mIdx];
    const isSerialCol = mapping.targetHeader.trim().toUpperCase() === 'SERIAL';

    outputHeaders.push(isSerialCol ? 'Serial' : mapping.targetHeader);

    // If this is Serial, we will populate it with 1..N after maxLength is computed
    if (isSerialCol) {
      log.push(`Output column "Serial" (will be populated with sequence 1..N)`);
      outputColumns.push([]); // placeholder, filled below
      servedByColumn.push([]);
      continue;
    }

    const outputValues: CellValue[] = [];
    const outputServed: Array<number | null> = [];

    log.push(`Output column "${mapping.targetHeader}"`);

    for (let amendmentIndex = 0; amendmentIndex < mapping.sourceHeaders.length; amendmentIndex++) {
      const sourceHeader = mapping.sourceHeaders[amendmentIndex];
      const servedValue = amendmentIndex + 1;
      const sourceColIndex = sourceHeaders.findIndex(h => h === sourceHeader);

      if (sourceColIndex === -1) {
        missing.add(sourceHeader);
        log.push(`  skipped missing source header: ${sourceHeader}`);
        continue;
      }

      let lastUsedRow = -1;
      for (let r = sourceRows.length - 1; r >= 0; r--) {
        const value = sourceRows[r]?.[sourceColIndex];
        if (value !== null && value !== undefined && String(value).trim() !== '') {
          lastUsedRow = r;
          break;
        }
      }

      if (lastUsedRow === -1) {
        log.push(`  source header found but column is empty: ${sourceHeader}`);
        continue;
      }

      for (let r = 0; r <= lastUsedRow; r++) {
        outputValues.push(sourceRows[r]?.[sourceColIndex] ?? null);
        outputServed.push(servedValue);
      }

      log.push(`  appended ${lastUsedRow + 1} value(s) from ${sourceHeader} (SERVED ${servedValue})`);
    }

    outputColumns.push(outputValues);
    servedByColumn.push(outputServed);
  }

  // Compute total output rows (maxLength)
  const maxLength = outputColumns.reduce((max, col) => Math.max(max, col.length), 0);

  // Fill Serial if it was in the mapping
  if (hasSerialInMapping) {
    outputColumns[serialMappingIndex] = Array.from({ length: maxLength }, (_, i) => i + 1);
    servedByColumn[serialMappingIndex] = Array.from({ length: maxLength }, () => 1);
  }

  // Compute SERVED sequence
  const servedSequence: number[] = Array.from({ length: maxLength }, (_, rowIndex) => {
    for (const servedColumn of servedByColumn) {
      const value = servedColumn[rowIndex];
      if (typeof value === 'number') return value;
    }
    return 1;
  });

  // Prepare final column list and headers:
  // Column 1 is ALWAYS 'SERVED'
  // Column 2 is 'Serial' (if not already mapped)
  let intermediateHeaders: string[];
  let intermediateColumns: CellValue[][];

  if (!hasSerialInMapping) {
    log.push('Auto-added "Serial" column with sequential numbering 1..' + maxLength);
    const autoSerialColumn: CellValue[] = Array.from({ length: maxLength }, (_, i) => i + 1);
    intermediateHeaders = ['SERVED', 'Serial', ...outputHeaders];
    intermediateColumns = [
      servedSequence,
      autoSerialColumn,
      ...outputColumns,
    ];
  } else {
    log.push('Populated mapped "Serial" column with sequential numbering 1..' + maxLength);
    intermediateHeaders = ['SERVED', ...outputHeaders];
    intermediateColumns = [
      servedSequence,
      ...outputColumns,
    ];
  }

  // Build intermediate 2D rows (before PARITY)
  let finalRows: CellValue[][] = Array.from({ length: maxLength }, (_, rowIndex) =>
    intermediateColumns.map(col => col[rowIndex] ?? null)
  );
  let finalHeaders = [...intermediateHeaders];

  // PARITY compares source mapping columns 2 and 3 (B and C) for one selected
  // output mapping row. Example: Q18 -> I_1_Q18 vs I_2_Q18.
  const parityMapping = effectiveMappings.find(mapping =>
    mapping.targetHeader.trim().toUpperCase() === options.parityMappingHeader?.trim().toUpperCase()
  ) ?? effectiveMappings.find(mapping =>
    mapping.targetHeader.trim().toUpperCase() !== 'SERIAL' &&
    mapping.sourceHeaders.length >= 2 &&
    mapping.sourceHeaders[0].toUpperCase() !== mapping.sourceHeaders[1].toUpperCase()
  );
  const parityMappingHeader = parityMapping?.targetHeader;
  const parityCol1 = parityMapping?.sourceHeaders[0];
  const parityCol2 = parityMapping?.sourceHeaders[1];

  // Compare paired source values on the same source row, then repeat that
  // comparison for each amendment segment in the output's append order.
  let parityStats: AppendResult['parityStats'] = undefined;

  if (parityMapping && parityCol1 && parityCol2) {
    const sourceIndex1 = sourceHeaders.findIndex(header => header === parityCol1);
    const sourceIndex2 = sourceHeaders.findIndex(header => header === parityCol2);

    if (sourceIndex1 !== -1 && sourceIndex2 !== -1) {
      const lastUsedFor = (sourceIndex: number) => {
        for (let r = sourceRows.length - 1; r >= 0; r--) {
          const value = sourceRows[r]?.[sourceIndex];
          if (value !== null && value !== undefined && String(value).trim() !== '') return r;
        }
        return -1;
      };
      const segmentLengths = [lastUsedFor(sourceIndex1) + 1, lastUsedFor(sourceIndex2) + 1];
      const perRespondent: number[] = [];
      const comparisonLength = Math.max(...segmentLengths);

      for (let r = 0; r < comparisonLength; r++) {
        perRespondent.push(checkParityMatch(
          sourceRows[r]?.[sourceIndex1] ?? null,
          sourceRows[r]?.[sourceIndex2] ?? null,
        ) ? 1 : 2);
      }

      let matches = 0;
      let mismatches = 0;
      for (const value of perRespondent) {
        if (value === 1) matches++;
        else mismatches++;
      }

      const parityValues: number[] = [];
      const comparisonPairs: Array<[CellValue, CellValue]> = [];
      segmentLengths.forEach(length => {
        for (let r = 0; r < length; r++) {
          parityValues.push(perRespondent[r] ?? 2);
          // Repeat both original values alongside their comparison in each
          // amendment block, without stacking the two reference fields.
          comparisonPairs.push([
            sourceRows[r]?.[sourceIndex1] ?? null,
            sourceRows[r]?.[sourceIndex2] ?? null,
          ]);
        }
      });
      while (parityValues.length < maxLength) {
        parityValues.push(2);
        comparisonPairs.push([null, null]);
      }

      finalHeaders.push('PARITY', parityCol1, parityCol2);
      finalRows = finalRows.map((row, r) => [
        ...row,
        parityValues[r],
        ...comparisonPairs[r],
      ]);
      parityStats = { matches, mismatches };
      log.push(
        `Auto-added "PARITY" for mapping "${parityMappingHeader}" by comparing mapped source fields "${parityCol1}" and "${parityCol2}" row-by-row: ` +
        `${matches} matches (code 1), ${mismatches} mismatches (code 2)`
      );
      log.push(`Appended comparison source columns "${parityCol1}" and "${parityCol2}" at the end of the output.`);
    }
  }

  return {
    headers: finalHeaders,
    rows: finalRows,
    log,
    missingHeaders: Array.from(missing),
    outputColumnLengths: [
      ...intermediateColumns.map(col => col.length),
      ...(parityStats ? [maxLength, maxLength, maxLength] : []),
    ],
    servedSequence,
    parityMappingHeader,
    parityCol1,
    parityCol2,
    parityStats,
  };
}

export function exportAppendResult(result: AppendResult, originalFileName: string): void {
  const workbook = XLSX.utils.book_new();
  const matrix = [result.headers, ...result.rows];
  const worksheet = XLSX.utils.aoa_to_sheet(matrix);

  worksheet['!cols'] = result.headers.map((header, colIndex) => {
    let max = String(header ?? '').length;
    for (const row of result.rows) {
      const value = row[colIndex];
      if (value !== null && value !== undefined) max = Math.max(max, String(value).length);
    }
    return { wch: Math.min(Math.max(max + 2, 10), 60) };
  });

  XLSX.utils.book_append_sheet(workbook, worksheet, 'OUTPUT');

  const baseName = originalFileName.replace(/\.(xlsx|xls|csv)$/i, '') || 'source';
  XLSX.writeFile(workbook, `${baseName}_OUTPUT.xlsx`, { bookType: 'xlsx' });
}

export function exportAppendResultCsv(result: AppendResult, originalFileName: string): void {
  const workbook = XLSX.utils.book_new();
  const matrix = [result.headers, ...result.rows];
  const worksheet = XLSX.utils.aoa_to_sheet(matrix);

  XLSX.utils.book_append_sheet(workbook, worksheet, 'OUTPUT');

  const baseName = originalFileName.replace(/\.(xlsx|xls|csv)$/i, '') || 'source';
  XLSX.writeFile(workbook, `${baseName}_OUTPUT.csv`, { bookType: 'csv' });
}
