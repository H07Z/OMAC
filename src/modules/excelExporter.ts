/**
 * Excel + TXT Exporter (single column, multi-tab)
 *
 * The output is a SINGLE COLUMN of text lines — column A of the source
 * (the code number) is NOT written. Each data row contributes one line:
 * "R " + description. Because blank descriptions are backfilled with the
 * code number upstream, a gap-filled row still renders as e.g. "R 37".
 *
 * XLSX: one worksheet per source tab, each a complete standalone block.
 * TXT : one file, all tabs merged:
 *       - only the first tab keeps <<< Project Entries >>>
 *       - intermediate Z terminators are dropped
 *       - only the last tab ends with Z
 *
 *   Line 1       <<< Project Entries >>>     (first tab only in TXT)
 *   Line 2       Q B                          (every tab block)
 *   Line 3       Q [<input1>] <input2>        (per-tab)
 *   Line 4       T M
 *   Line 5       L <in1>L<in2>R<in3>          (per-tab)
 *   Line 6 … N   R <description>              (data)
 *   Line N+1     Q E
 *   Line N+2     Z                            (last tab only in TXT)
 */

import * as XLSX from 'xlsx';
import { RowData } from '../types';
import { debugLog } from '../utils/debug';

/** Line number (1-indexed) where the first data record is written */
export const DATA_START_ROW = 6;

export const TITLE_LINE = '<<< Project Entries >>>';
export const BEGIN_LINE = 'Q B';
export const TABLE_LINE = 'T M';
export const END_LINE = 'Q E';
export const TERMINATOR_LINE = 'Z';

/** Prefix written in front of every description on data lines */
export const CODE_PREFIX = 'R ';

const HIERARCHY_MARKERS: readonly string[] = [
  '(NET)',
  '(SUBNET)',
  '(SUB-SUBNET)',
  '(SUB-SUB-SUBNET)',
];

export interface HeaderLines {
  qLine: string;
  lLine: string;
}

export interface SheetExportSpec {
  name: string;
  data: RowData[];
  outputHeaders: string[];
  headerLines: HeaderLines;
}

export type RowKind =
  | 'title' | 'begin' | 'qline' | 'table' | 'lline'
  | 'data' | 'end' | 'terminator';

/** One output line. `rowNumber` is its real position in the file/sheet. */
export interface PreviewRow {
  rowNumber: number;
  text: string;
  kind: RowKind;
}

/* ─── Guards ─────────────────────────────── */

function stripHierarchyRows(data: RowData[], descriptionKey: string): RowData[] {
  if (!descriptionKey) return data;
  return data.filter(record => {
    const value = record[descriptionKey];
    if (value === null || value === undefined) return true;
    const normalized = String(value).trim().toUpperCase();
    return !HIERARCHY_MARKERS.some(marker => normalized.includes(marker));
  });
}

/** Resolve one record to its output text: "R " + description (code fallback) */
function dataLineText(record: RowData, outputHeaders: string[]): string {
  const codeKey = outputHeaders[0] ?? '';
  const descKey = outputHeaders[1] ?? outputHeaders[0] ?? '';

  const rawDesc = record[descKey];
  const descStr = rawDesc === null || rawDesc === undefined ? '' : String(rawDesc);

  if (descStr.trim() !== '') return `${CODE_PREFIX}${descStr}`;

  // Blank description → fall back to the code number
  const rawCode = record[codeKey];
  const codeStr = rawCode === null || rawCode === undefined ? '' : String(rawCode);
  return codeStr === '' ? CODE_PREFIX.trimEnd() : `${CODE_PREFIX}${codeStr}`;
}

/* ─── Line building ──────────────────────── */

/**
 * Build one complete tab block as numbered lines.
 * Shared by the XLSX writer, the TXT writer and the preview so the
 * three can never drift apart.
 */
function buildLines(
  data: RowData[],
  outputHeaders: string[],
  headerLines: HeaderLines,
  options: { includeTitle: boolean; includeTerminator: boolean; startNumber: number }
): PreviewRow[] {
  const lines: PreviewRow[] = [];
  let n = options.startNumber;
  const push = (text: string, kind: RowKind) => {
    lines.push({ rowNumber: n++, text, kind });
  };

  if (options.includeTitle) push(TITLE_LINE, 'title');

  push(BEGIN_LINE, 'begin');
  push(headerLines.qLine, 'qline');
  push(TABLE_LINE, 'table');
  push(headerLines.lLine, 'lline');

  for (const record of data) {
    push(dataLineText(record, outputHeaders), 'data');
  }

  push(END_LINE, 'end');
  if (options.includeTerminator) push(TERMINATOR_LINE, 'terminator');

  return lines;
}

function sanitizeSheetName(name: string, used: Set<string>): string {
  let clean = (name || 'Sheet').replace(/[:\\/?*[\]]/g, ' ').trim() || 'Sheet';
  clean = clean.slice(0, 31);
  let candidate = clean;
  let suffix = 2;
  while (used.has(candidate.toLowerCase())) {
    const tail = ` (${suffix++})`;
    candidate = clean.slice(0, 31 - tail.length) + tail;
  }
  used.add(candidate.toLowerCase());
  return candidate;
}

/* ─── XLSX Export (one sheet per tab) ────── */

export function exportWorkbook(
  specs: SheetExportSpec[],
  originalFileName: string
): void {
  debugLog('Exporter', `Exporting XLSX workbook with ${specs.length} sheet(s)`);

  const workbook = XLSX.utils.book_new();
  const usedNames = new Set<string>();

  for (const spec of specs) {
    const sheetName = sanitizeSheetName(spec.name, usedNames);
    const descriptionKey = spec.outputHeaders[1] ?? spec.outputHeaders[0] ?? '';
    const cleanData = stripHierarchyRows(spec.data, descriptionKey);
    const droppedCount = spec.data.length - cleanData.length;

    // Every sheet is a complete standalone block: title … Z
    const lines = buildLines(cleanData, spec.outputHeaders, spec.headerLines, {
      includeTitle: true,
      includeTerminator: true,
      startNumber: 1,
    });

    // Single column
    const matrix = lines.map(l => [l.text]);
    const worksheet = XLSX.utils.aoa_to_sheet(matrix);

    const widest = lines.reduce((max, l) => Math.max(max, l.text.length), 10);
    worksheet['!cols'] = [{ wch: Math.min(widest + 2, 120) }];

    XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);

    debugLog('Exporter', `  Sheet "${sheetName}": ${cleanData.length} records`
      + (droppedCount > 0 ? ` (guard removed ${droppedCount})` : ''));
  }

  const baseName = originalFileName.replace(/\.(xlsx|xls|csv)$/i, '') || 'output';
  const outputFileName = `${baseName}_processed.xlsx`;

  XLSX.writeFile(workbook, outputFileName, { bookType: 'xlsx' });
  debugLog('Exporter', `Download triggered: ${outputFileName}`);
}

/* ─── TXT Merged Export ──────────────────── */

/**
 * Build the combined TXT lines (all tabs merged).
 * Only the first tab carries the title; only the last carries Z.
 */
export function buildCombinedTxtLines(specs: SheetExportSpec[]): PreviewRow[] {
  const all: PreviewRow[] = [];
  let counter = 1;

  specs.forEach((spec, idx) => {
    const descriptionKey = spec.outputHeaders[1] ?? spec.outputHeaders[0] ?? '';
    const cleanData = stripHierarchyRows(spec.data, descriptionKey);

    const block = buildLines(cleanData, spec.outputHeaders, spec.headerLines, {
      includeTitle: idx === 0,
      includeTerminator: idx === specs.length - 1,
      startNumber: counter,
    });

    all.push(...block);
    counter += block.length;
  });

  return all;
}

export function exportCombinedTxt(
  specs: SheetExportSpec[],
  originalFileName: string
): void {
  const lines = buildCombinedTxtLines(specs);
  const content = lines.map(l => l.text).join('\r\n');

  const baseName = originalFileName.replace(/\.(xlsx|xls|csv)$/i, '') || 'output';
  const outputFileName = `${baseName}_combined.txt`;

  const blob = new Blob([content], { type: 'text/plain;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = outputFileName;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);

  debugLog('Exporter', `Combined TXT download triggered: ${outputFileName} (${lines.length} lines)`);
}

/* ─── Previews ───────────────────────────── */

/** Preview of ONE sheet: full header block, first data lines, then the closing lines. */
export function buildExportPreview(
  data: RowData[],
  outputHeaders: string[],
  headerLines: HeaderLines,
  maxDataRows = 5
): { rows: PreviewRow[]; hiddenCount: number } {
  const descriptionKey = outputHeaders[1] ?? outputHeaders[0] ?? '';
  const cleanData = stripHierarchyRows(data, descriptionKey);

  const all = buildLines(cleanData, outputHeaders, headerLines, {
    includeTitle: true,
    includeTerminator: true,
    startNumber: 1,
  });

  const head = all.filter(r => r.kind !== 'data' && r.rowNumber < DATA_START_ROW);
  const dataRows = all.filter(r => r.kind === 'data');
  const tail = all.filter(r => r.kind === 'end' || r.kind === 'terminator');

  const shown = dataRows.slice(0, maxDataRows);

  return {
    rows: [...head, ...shown, ...tail],
    hiddenCount: Math.max(dataRows.length - shown.length, 0),
  };
}

/** Preview of the merged TXT: first and last slices with a gap in between. */
export function buildCombinedTxtPreview(
  specs: SheetExportSpec[],
  maxTotalLines = 60
): { rows: PreviewRow[]; hiddenCount: number } {
  const all = buildCombinedTxtLines(specs);
  if (all.length <= maxTotalLines) return { rows: all, hiddenCount: 0 };

  const half = Math.floor(maxTotalLines / 2);
  const first = all.slice(0, half);
  const last = all.slice(-half);

  return { rows: [...first, ...last], hiddenCount: all.length - first.length - last.length };
}
