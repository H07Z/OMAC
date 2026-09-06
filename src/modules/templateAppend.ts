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
 * A generated SERVED column is always inserted first:
 * - rows copied from the first mapped/amend source header get SERVED = 1
 * - rows copied from the second mapped/amend source header get SERVED = 2
 * - and so on
 */

import * as XLSX from 'xlsx';

export type CellValue = string | number | null;

export interface TemplateMappingRow {
  targetHeader: string;
  sourceHeaders: string[];
}

export interface AppendResult {
  headers: string[];
  rows: CellValue[][];
  log: string[];
  missingHeaders: string[];
  outputColumnLengths: number[];
  servedSequence: number[];
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
 * This intentionally mimics Excel Range.Find(... LookAt:=xlWhole): exact match,
 * first match wins if duplicate source headers exist.
 */
export function appendByTemplate(
  sourceMatrix: CellValue[][],
  mappings: TemplateMappingRow[]
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

  const outputColumns: CellValue[][] = [];
  const servedByColumn: Array<Array<number | null>> = [];
  const outputHeaders: string[] = [];

  // SERVED is generated, so ignore an accidentally supplied SERVED mapping.
  const effectiveMappings = mappings.filter(mapping => mapping.targetHeader.trim().toUpperCase() !== 'SERVED');
  if (effectiveMappings.length !== mappings.length) {
    log.push('Ignored template row "SERVED" because SERVED is generated automatically.');
  }

  for (const mapping of effectiveMappings) {
    outputHeaders.push(mapping.targetHeader);
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

  const maxLength = outputColumns.reduce((max, col) => Math.max(max, col.length), 0);
  const servedSequence: number[] = Array.from({ length: maxLength }, (_, rowIndex) => {
    // Columns should align by amendment. Use the first output column that has
    // a SERVED marker at this row as the canonical generated value.
    for (const servedColumn of servedByColumn) {
      const value = servedColumn[rowIndex];
      if (typeof value === 'number') return value;
    }
    return 1;
  });

  const outputRows: CellValue[][] = Array.from({ length: maxLength }, (_, rowIndex) => [
    servedSequence[rowIndex],
    ...outputColumns.map(col => col[rowIndex] ?? null),
  ]);

  return {
    headers: ['SERVED', ...outputHeaders],
    rows: outputRows,
    log,
    missingHeaders: Array.from(missing),
    outputColumnLengths: outputColumns.map(col => col.length),
    servedSequence,
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
    return { wch: Math.min(Math.max(max + 2, 12), 60) };
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
