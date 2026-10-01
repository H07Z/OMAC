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
 * 3. PARITY (Last column):
 *    - For the selected mapping row, compares source mapping cells B and C
 *      (the first two source headers), e.g. I_1_Q18 vs I_2_Q18.
 *    - If values match (TRUE), coded as 1.
 *    - If values do not match (FALSE), coded as 2.
 * 4. ROUND override (optional):
 *    - A user-entered ROUND number can populate the whole output ROUND column.
 *    - If ROUND is absent from the mapping, the column is auto-added.
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
  /** Optional fixed ROUND number/text applied to every output data row. */
  roundValue?: string | number | null;
}

export type StackDataCheckStatus = 'pass' | 'warn' | 'fail';

export interface StackDataCheck {
  id: string;
  label: string;
  status: StackDataCheckStatus;
  summary: string;
  detail?: string;
}

export interface StackDataMissingColumn {
  header: string;
  index: number;
  missing: number;
  total: number;
  missingRate: number;
}

export interface StackDataMissingExample {
  outputRowNumber: number;
  serial: CellValue;
  served: CellValue;
  missingHeaders: string[];
}

export interface StackDataQuality {
  totalRows: number;
  totalAnswerCells: number;
  missingCells: number;
  rowsWithMissing: number;
  missingByColumn: StackDataMissingColumn[];
  missingExamples: StackDataMissingExample[];
  serialComplete: boolean;
  serialDuplicates: number;
  servedCounts: Record<string, number>;
  servedMonotonic: boolean;
  roundMode: 'override' | 'auto-added' | 'mapped' | 'none';
  roundValue?: CellValue;
  roundFilled: number;
  checks: StackDataCheck[];
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
  roundValue?: CellValue;
  roundMode?: 'override' | 'auto-added' | 'mapped' | 'none';
  quality?: StackDataQuality;
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

function isMissingCell(value: CellValue): boolean {
  return value === null || value === undefined || String(value).trim() === '';
}

function normalizeRoundValue(value: string | number | null | undefined): CellValue {
  if (value === null || value === undefined) return null;
  if (typeof value === 'number' && Number.isFinite(value)) return value;
  const text = String(value).trim();
  if (text === '') return null;
  if (/^-?\d+(\.\d+)?$/.test(text)) return Number(text);
  return text;
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

function buildStackDataQuality(args: {
  headers: string[];
  rows: CellValue[][];
  missingHeaders: string[];
  roundMode: 'override' | 'auto-added' | 'mapped' | 'none';
  roundValue?: CellValue;
  parityStats?: AppendResult['parityStats'];
  parityCol1?: string;
  parityCol2?: string;
}): StackDataQuality {
  const { headers, rows, missingHeaders } = args;
  const totalRows = rows.length;
  const headerIndex = (name: string) => headers.findIndex(h => h.trim().toUpperCase() === name);
  const servedIndex = headerIndex('SERVED');
  const serialIndex = headerIndex('SERIAL');
  const parityIndex = headerIndex('PARITY');
  const roundIndex = headerIndex('ROUND');

  const answerIndexes = headers
    .map((header, index) => ({ header, index }))
    .filter(({ index }) => index !== servedIndex && index !== serialIndex && index !== parityIndex)
    .map(({ index }) => index);

  const missingByColumn: StackDataMissingColumn[] = answerIndexes.map(index => ({
    header: headers[index],
    index,
    missing: 0,
    total: totalRows,
    missingRate: 0,
  }));
  const missingByIndex = new Map<number, StackDataMissingColumn>();
  missingByColumn.forEach(item => missingByIndex.set(item.index, item));

  const missingExamples: StackDataMissingExample[] = [];
  let missingCells = 0;
  let rowsWithMissing = 0;

  rows.forEach((row, rowIndex) => {
    const missingHeadersInRow: string[] = [];
    for (const index of answerIndexes) {
      if (isMissingCell(row[index] ?? null)) {
        missingCells += 1;
        missingHeadersInRow.push(headers[index]);
        const column = missingByIndex.get(index);
        if (column) column.missing += 1;
      }
    }
    if (missingHeadersInRow.length > 0) {
      rowsWithMissing += 1;
      if (missingExamples.length < 8) {
        missingExamples.push({
          outputRowNumber: rowIndex + 2,
          serial: serialIndex >= 0 ? rows[rowIndex][serialIndex] ?? null : null,
          served: servedIndex >= 0 ? rows[rowIndex][servedIndex] ?? null : null,
          missingHeaders: missingHeadersInRow.slice(0, 6),
        });
      }
    }
  });

  missingByColumn.forEach(column => {
    column.missingRate = column.total > 0 ? column.missing / column.total : 0;
  });
  const columnsWithMissing = missingByColumn
    .filter(column => column.missing > 0)
    .sort((a, b) => b.missing - a.missing);

  let serialComplete = totalRows > 0;
  let serialDuplicates = 0;
  if (serialIndex >= 0) {
    const seen = new Set<string>();
    let duplicates = 0;
    for (let rowIndex = 0; rowIndex < totalRows; rowIndex++) {
      const raw = rows[rowIndex][serialIndex];
      const key = raw === null || raw === undefined ? '' : String(raw).trim();
      if (key === '' || Number(key) !== rowIndex + 1) serialComplete = false;
      if (key !== '') {
        if (seen.has(key)) duplicates += 1;
        else seen.add(key);
      }
    }
    serialDuplicates = duplicates;
    if (duplicates > 0) serialComplete = false;
  } else {
    serialComplete = false;
  }

  const servedCounts: Record<string, number> = {};
  let servedMonotonic = true;
  let previousServed = Number.NEGATIVE_INFINITY;
  let servedMissing = 0;
  if (servedIndex >= 0) {
    for (let rowIndex = 0; rowIndex < totalRows; rowIndex++) {
      const raw = rows[rowIndex][servedIndex];
      if (isMissingCell(raw)) {
        servedMissing += 1;
        servedMonotonic = false;
        continue;
      }
      const numeric = Number(String(raw).trim());
      const key = Number.isFinite(numeric) ? String(numeric) : String(raw).trim();
      servedCounts[key] = (servedCounts[key] ?? 0) + 1;
      if (Number.isFinite(numeric)) {
        if (numeric < previousServed) servedMonotonic = false;
        previousServed = numeric;
      }
    }
  } else {
    servedMonotonic = false;
  }

  let roundFilled = 0;
  if (roundIndex >= 0) {
    for (let rowIndex = 0; rowIndex < totalRows; rowIndex++) {
      if (!isMissingCell(rows[rowIndex][roundIndex] ?? null)) roundFilled += 1;
    }
  }

  const checks: StackDataCheck[] = [];
  if (missingCells === 0) {
    checks.push({
      id: 'missing-answers',
      label: 'No missing answers',
      status: 'pass',
      summary: `Checked ${answerIndexes.length.toLocaleString()} answer columns across ${totalRows.toLocaleString()} rows.`,
    });
  } else {
    checks.push({
      id: 'missing-answers',
      label: 'Missing answers found',
      status: 'fail',
      summary: `${missingCells.toLocaleString()} missing cells in ${rowsWithMissing.toLocaleString()} rows.`,
      detail: columnsWithMissing.slice(0, 5).map(column => `${column.header}: ${column.missing}`).join(' • '),
    });
  }

  if (missingHeaders.length === 0) {
    checks.push({
      id: 'mapping-coverage',
      label: 'Mapped source headers found',
      status: 'pass',
      summary: 'Every mapped source header was found in the source worksheet.',
    });
  } else {
    checks.push({
      id: 'mapping-coverage',
      label: 'Mapped source headers missing',
      status: 'fail',
      summary: `${missingHeaders.length.toLocaleString()} mapped source header(s) were not found.`,
      detail: missingHeaders.slice(0, 8).join(', '),
    });
  }

  if (serialComplete && serialDuplicates === 0) {
    checks.push({
      id: 'serial-sequence',
      label: 'Serial sequence complete',
      status: 'pass',
      summary: `Serial runs 1 to ${totalRows.toLocaleString()} with no gaps or duplicates.`,
    });
  } else {
    checks.push({
      id: 'serial-sequence',
      label: 'Serial sequence needs attention',
      status: 'fail',
      summary: serialDuplicates > 0
        ? `${serialDuplicates.toLocaleString()} duplicate Serial value(s) found.`
        : 'Serial is not a complete 1-to-N sequence.',
    });
  }

  if (servedMissing === 0 && servedMonotonic) {
    const segments = Object.entries(servedCounts)
      .sort((a, b) => Number(a[0]) - Number(b[0]))
      .map(([value, count]) => `${value}: ${count.toLocaleString()}`)
      .join(' • ');
    checks.push({
      id: 'served-sequence',
      label: 'SERVED stacking looks consistent',
      status: 'pass',
      summary: segments ? `Counts by SERVED — ${segments}.` : 'SERVED values are present.',
    });
  } else {
    checks.push({
      id: 'served-sequence',
      label: 'SERVED stacking needs attention',
      status: 'warn',
      summary: servedMissing > 0
        ? `${servedMissing.toLocaleString()} row(s) have no SERVED value.`
        : 'SERVED values are not in contiguous amendment order.',
    });
  }

  if (args.roundMode === 'override' || args.roundMode === 'auto-added') {
    if (roundFilled === totalRows && totalRows > 0) {
      checks.push({
        id: 'round-filled',
        label: 'ROUND populated',
        status: 'pass',
        summary: `ROUND is "${String(args.roundValue ?? '')}" on all ${totalRows.toLocaleString()} rows.`,
      });
    } else {
      checks.push({
        id: 'round-filled',
        label: 'ROUND incomplete',
        status: 'fail',
        summary: `ROUND is filled on ${roundFilled.toLocaleString()} of ${totalRows.toLocaleString()} rows.`,
      });
    }
  } else if (roundIndex >= 0) {
    const missingRound = totalRows - roundFilled;
    checks.push({
      id: 'round-filled',
      label: missingRound === 0 ? 'ROUND populated' : 'ROUND has gaps',
      status: missingRound === 0 ? 'pass' : 'warn',
      summary: missingRound === 0
        ? `Mapped ROUND has values on all ${totalRows.toLocaleString()} rows.`
        : `${missingRound.toLocaleString()} ROUND cell(s) are blank.`,
    });
  } else {
    checks.push({
      id: 'round-filled',
      label: 'No ROUND override supplied',
      status: 'pass',
      summary: 'No fixed ROUND number was requested.',
    });
  }

  if (args.parityStats) {
    checks.push({
      id: 'parity-available',
      label: 'PARITY calculated',
      status: 'pass',
      summary: `${args.parityCol1} vs ${args.parityCol2}: ${args.parityStats.matches.toLocaleString()} equal, ${args.parityStats.mismatches.toLocaleString()} different.`,
    });
  } else {
    checks.push({
      id: 'parity-available',
      label: 'PARITY not calculated',
      status: 'warn',
      summary: 'No eligible mapping row with two distinct source headers was available.',
    });
  }

  return {
    totalRows,
    totalAnswerCells: totalRows * answerIndexes.length,
    missingCells,
    rowsWithMissing,
    missingByColumn: columnsWithMissing,
    missingExamples,
    serialComplete,
    serialDuplicates,
    servedCounts,
    servedMonotonic,
    roundMode: args.roundMode,
    roundValue: args.roundValue,
    roundFilled,
    checks,
  };
}

/**
 * Execute the append-by-template operation.
 * Supports:
 * - SERVED column as column 1 (1 = 1st amend, 2 = 2nd amend, ...)
 * - Serial column (1..N sequence; auto-added if not in mapping)
 * - ROUND override (optional fixed ROUND number for every output row)
 * - PARITY column (compares 2 selected columns: 1 = TRUE, 2 = FALSE)
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
  let effectiveMappings = mappings.filter(
    mapping => mapping.targetHeader.trim().toUpperCase() !== 'SERVED'
  );
  if (effectiveMappings.length !== mappings.length) {
    log.push('Ignored template row "SERVED" because SERVED is generated automatically as column 1.');
  }

  const roundOverride = normalizeRoundValue(options.roundValue);
  const hasRoundMapping = effectiveMappings.some(
    mapping => mapping.targetHeader.trim().toUpperCase() === 'ROUND'
  );
  let autoAddedRound = false;
  if (roundOverride !== null && !hasRoundMapping) {
    const serialPosition = effectiveMappings.findIndex(
      mapping => mapping.targetHeader.trim().toUpperCase() === 'SERIAL'
    );
    const roundMapping: TemplateMappingRow = { targetHeader: 'ROUND', sourceHeaders: [] };
    effectiveMappings = serialPosition === -1
      ? [roundMapping, ...effectiveMappings]
      : [
          ...effectiveMappings.slice(0, serialPosition + 1),
          roundMapping,
          ...effectiveMappings.slice(serialPosition + 1),
        ];
    autoAddedRound = true;
    log.push(`Auto-added "ROUND" column and filled every row with "${String(roundOverride)}".`);
  }

  // Check if Serial was specified in the mapping
  const serialMappingIndex = effectiveMappings.findIndex(
    m => m.targetHeader.trim().toUpperCase() === 'SERIAL'
  );
  const hasSerialInMapping = serialMappingIndex !== -1;
  const roundMappingIndex = effectiveMappings.findIndex(
    m => m.targetHeader.trim().toUpperCase() === 'ROUND'
  );

  const outputColumns: CellValue[][] = [];
  const servedByColumn: Array<Array<number | null>> = [];
  const outputHeaders: string[] = [];

  for (let mIdx = 0; mIdx < effectiveMappings.length; mIdx++) {
    const mapping = effectiveMappings[mIdx];
    const normalizedTarget = mapping.targetHeader.trim().toUpperCase();
    const isSerialCol = normalizedTarget === 'SERIAL';
    const isRoundOverrideCol = normalizedTarget === 'ROUND' && roundOverride !== null;

    outputHeaders.push(isSerialCol ? 'Serial' : mapping.targetHeader);

    // If this is Serial, we will populate it with 1..N after maxLength is computed
    if (isSerialCol) {
      log.push(`Output column "Serial" (will be populated with sequence 1..N)`);
      outputColumns.push([]); // placeholder, filled below
      servedByColumn.push([]);
      continue;
    }

    // A fixed ROUND number replaces source lookup and is applied to every row.
    if (isRoundOverrideCol) {
      log.push(`Output column "ROUND" (will be populated with fixed value "${String(roundOverride)}")`);
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

  // Fill ROUND with the requested fixed value. ROUND is a control column and
  // must not influence the SERVED amendment sequence.
  let roundMode: AppendResult['roundMode'] = 'none';
  if (roundOverride !== null && roundMappingIndex !== -1) {
    outputColumns[roundMappingIndex] = Array.from({ length: maxLength }, () => roundOverride);
    servedByColumn[roundMappingIndex] = Array.from({ length: maxLength }, () => null);
    roundMode = autoAddedRound ? 'auto-added' : 'override';
  } else if (roundMappingIndex !== -1) {
    roundMode = 'mapped';
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
      // Determine length of each amendment segment from servedSequence
      const segmentLengths: number[] = [];
      let currentServed = 1;
      let currentLen = 0;
      for (let i = 0; i < maxLength; i++) {
        const s = servedSequence[i];
        if (s !== currentServed) {
          segmentLengths.push(currentLen);
          currentServed = s;
          currentLen = 1;
        } else {
          currentLen++;
        }
      }
      if (currentLen > 0) {
        segmentLengths.push(currentLen);
      }
      if (segmentLengths.length === 0) {
        segmentLengths.push(maxLength);
      }

      const col1Values: CellValue[] = [];
      const col2Values: CellValue[] = [];
      const parityValues: number[] = [];

      let matches = 0;
      let mismatches = 0;

      // Count matches and mismatches across source respondents
      const maxSourceLen = Math.max(...segmentLengths);
      for (let r = 0; r < maxSourceLen; r++) {
        const v1 = sourceRows[r]?.[sourceIndex1] ?? null;
        const v2 = sourceRows[r]?.[sourceIndex2] ?? null;
        if (checkParityMatch(v1, v2)) {
          matches++;
        } else {
          mismatches++;
        }
      }

      // Populate values for every row in the output across each amendment segment
      segmentLengths.forEach(length => {
        for (let r = 0; r < length; r++) {
          const v1 = sourceRows[r]?.[sourceIndex1] ?? null;
          const v2 = sourceRows[r]?.[sourceIndex2] ?? null;
          col1Values.push(v1);
          col2Values.push(v2);
          parityValues.push(checkParityMatch(v1, v2) ? 1 : 2);
        }
      });

      while (col1Values.length < maxLength) {
        col1Values.push(null);
        col2Values.push(null);
        parityValues.push(2);
      }

      // Add the 2 columns that parity used, along with PARITY at the end of the output table
      finalHeaders.push('PARITY', parityCol1, parityCol2);
      finalRows = finalRows.map((row, r) => [
        ...row,
        parityValues[r],
        col1Values[r],
        col2Values[r],
      ]);
      parityStats = { matches, mismatches };
      log.push(
        `Auto-added "PARITY" (1=TRUE, 2=FALSE) and the 2 comparison source columns "${parityCol1}" and "${parityCol2}" at the end: ` +
        `${matches} matches (code 1), ${mismatches} mismatches (code 2)`
      );
    }
  }

  const quality = buildStackDataQuality({
    headers: finalHeaders,
    rows: finalRows,
    missingHeaders: Array.from(missing),
    roundMode,
    roundValue: roundOverride === null ? undefined : roundOverride,
    parityStats,
    parityCol1,
    parityCol2,
  });
  const missingCheck = quality.checks.find(check => check.id === 'missing-answers');
  if (missingCheck) log.push(`Data check — ${missingCheck.label}: ${missingCheck.summary}`);
  if (roundMode === 'override' || roundMode === 'auto-added') {
    log.push(`ROUND mode: ${roundMode}; applied value "${String(roundOverride)}" to ${quality.roundFilled.toLocaleString()} row(s).`);
  }

  return {
    headers: finalHeaders,
    rows: finalRows,
    log,
    missingHeaders: Array.from(missing),
    outputColumnLengths: intermediateColumns.map(col => col.length),
    servedSequence,
    parityMappingHeader,
    parityCol1,
    parityCol2,
    parityStats,
    roundValue: roundOverride === null ? undefined : roundOverride,
    roundMode,
    quality,
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
