/**
 * QPS Script Generator
 *
 * Faithful port of the VBA GenerateQPSScriptNettings macro.
 *
 * Source columns (on upload):
 *   Column A / J = Code
 *   Column B / K = Label   (includes hierarchy markers)
 *
 * Per-tab inputs (VBA cells C3/C4/C6/C10):
 *   variableName  (C3)  e.g. "VP2"
 *   questions     (C4)  e.g. "P2"  (comma-separated: "B1,B2,B3")
 *   questionLabel (C6)  e.g. "P2 REASONS FOR PRODUCT PREFERENCE"
 *   tableType     (C10) e.g. "M"  → rendered as "T M"
 *
 * Output (single column of text lines):
 *   V B
 *   V [<variableName>] <questionLabel>
 *   T <tableType>
 *   for each data row (empty rows removed, markers kept):
 *     marker row:  R <b><label></b>   then  Y <netting script>
 *     code row:    R <label>          then  Y $<q>/<code>[+$<q2>/<code>]
 *
 * Netting rules (per marker level):
 *   (NET)              collects codes until the next (NET)                       or END
 *   (SUBNET)           collects until next (SUBNET)/(NET)                         or END
 *   (SUB-SUBNET)       collects until next (SUB-SUBNET)/(SUBNET)/(NET)           or END
 *   (SUB-SUB-SUBNET)   collects until next (SUB-SUB-SUBNET)/(SUB-SUBNET)/…/(NET) or END
 *
 * Codes are grouped in ROW ORDER into consecutive-value runs:
 *   3001,3002,…,3016 → "3001..3016";  a gap or out-of-order value closes the run.
 */

import * as XLSX from 'xlsx';
import { RowData } from '../types';
import { debugLog } from '../utils/debug';

/* ─── Config & types ─────────────────────── */

export interface QpsConfig {
  variableName: string;
  questions: string;     // comma-separated
  questionLabel: string;
  tableType: string;     // default "M"
}

export interface QpsSheetSpec {
  name: string;
  data: RowData[];       // ORIGINAL rows (markers kept, unsorted)
  codeKey: string;       // Column A key
  labelKey: string;      // Column B key
  config: QpsConfig;
}

export type QpsLineKind =
  | 'title' | 'begin' | 'variable' | 'table'
  | 'net-label' | 'net-script'
  | 'code-label' | 'code-script'
  | 'option-open' | 'option-end'
  | 'plain' | 'end' | 'terminator';

export interface QpsLine {
  text: string;
  kind: QpsLineKind;
}

/* ─── Hierarchy levels ───────────────────── */

/** Marker → stop set (from the VBA per-level loops) */
const STOPS: Record<number, string[]> = {
  1: ['(NET)'],
  2: ['(SUBNET)', '(NET)'],
  3: ['(SUB-SUBNET)', '(SUBNET)', '(NET)'],
  4: ['(SUB-SUB-SUBNET)', '(SUB-SUBNET)', '(SUBNET)', '(NET)'],
};

/** Detect the hierarchy level of a label (0 = not a marker) */
function detectLevel(labelUpper: string): number {
  if (labelUpper.includes('(SUB-SUB-SUBNET)')) return 4;
  if (labelUpper.includes('(SUB-SUBNET)')) return 3;
  if (labelUpper.includes('(SUBNET)')) return 2;
  if (labelUpper.includes('(NET)')) return 1;
  return 0;
}

/** VBA Val(): leading numeric part, blank/non-numeric → 0 */
function val(v: unknown): number {
  const n = parseFloat(String(v ?? '').trim());
  return isNaN(n) ? 0 : n;
}

/* ─── Netting: build consecutive-run ranges ─ */

/**
 * Replicates the VBA inner range-collection loop for one marker.
 * Scans forward from `startIdx`, forming consecutive-value runs in
 * row order, and stops at the first same-or-higher-level marker
 * (per `stops`) or an "END" row.
 */
function buildRangesForMarker(
  rows: RowData[],
  startIdx: number,
  codeKey: string,
  labelKey: string,
  stops: string[]
): string[] {
  const ranges: string[] = [];
  let strCodes = '';
  let strHold = '';
  const n = rows.length;

  for (let x = startIdx; x < n; x++) {
    const codeVal = rows[x][codeKey];
    const codeStr = codeVal === null || codeVal === undefined ? '' : String(codeVal).trim();
    const labelUpper = String(rows[x][labelKey] ?? '').toUpperCase();

    // Extend / start the current run
    if (codeStr !== '') {
      if (strCodes === '') {
        strCodes = codeStr;
        strHold = codeStr;
      } else if (val(codeStr) - val(strHold) === 1) {
        strHold = codeStr;
      }
    }

    const diff = val(codeStr) - val(strHold);
    const isLast = x === n - 1;

    if ((diff !== 1 && diff !== 0) || isLast) {
      // Close the current run
      let range = strCodes;
      if (strCodes !== strHold) range = `${strCodes}..${strHold}`;
      if (range !== '') ranges.push(range);

      // Reset run to the current row's code
      strCodes = codeStr;
      strHold = codeStr;

      // Stop when we reach a same-or-higher-level marker (or END)
      if (labelUpper === 'END' || stops.some(s => labelUpper.includes(s))) {
        break;
      }
    }
  }

  return ranges;
}

interface OptionBoundary {
  rowIndex: number;
  level: number;
}

/**
 * Find the first and last code beneath every hierarchy marker.
 * A hierarchy block ends at END or at the next marker of the same or a
 * higher hierarchy level. The QPS output adds O R after the first code and
 * O E after the last code, matching the supplied expected output.
 */
function buildOptionBoundaries(
  rows: RowData[],
  codeKey: string,
  labelKey: string,
): { opens: Map<number, OptionBoundary[]>; closes: Map<number, OptionBoundary[]> } {
  const opens = new Map<number, OptionBoundary[]>();
  const closes = new Map<number, OptionBoundary[]>();

  const add = (map: Map<number, OptionBoundary[]>, boundary: OptionBoundary) => {
    map.set(boundary.rowIndex, [...(map.get(boundary.rowIndex) ?? []), boundary]);
  };

  for (let i = 0; i < rows.length; i++) {
    const level = detectLevel(String(rows[i][labelKey] ?? '').toUpperCase());
    if (level === 0) continue;

    let firstCode = -1;
    let lastCode = -1;

    for (let x = i + 1; x < rows.length; x++) {
      const labelUpper = String(rows[x][labelKey] ?? '').trim().toUpperCase();
      if (labelUpper === 'END') break;

      const nextLevel = detectLevel(labelUpper);
      if (nextLevel > 0 && nextLevel <= level) break;

      const code = String(rows[x][codeKey] ?? '').trim();
      if (code !== '') {
        if (firstCode === -1) firstCode = x;
        lastCode = x;
      }
    }

    if (firstCode !== -1 && lastCode !== -1) {
      add(opens, { rowIndex: firstCode, level });
      add(closes, { rowIndex: lastCode, level });
    }
  }

  // At a shared first row, open outer scopes before inner scopes. At a shared
  // last row, close inner scopes before outer scopes.
  for (const values of opens.values()) values.sort((a, b) => a.level - b.level);
  for (const values of closes.values()) values.sort((a, b) => b.level - a.level);

  return { opens, closes };
}

/* ─── Line generation ────────────────────── */

/** Split & clean the comma-separated question list */
function parseQuestions(questions: string): string[] {
  return questions.split(',').map(s => s.trim()).filter(s => s !== '');
}

/**
 * Generate all QPS script lines for one tab.
 */
export function generateQpsLines(
  rows: RowData[],
  codeKey: string,
  labelKey: string,
  config: QpsConfig
): QpsLine[] {
  // Remove fully-empty rows (both code and label blank) — VBA step 1
  const data = rows.filter(r => {
    const c = String(r[codeKey] ?? '').trim();
    const l = String(r[labelKey] ?? '').trim();
    return !(c === '' && l === '');
  });

  const questions = parseQuestions(config.questions);
  const lines: QpsLine[] = [];
  const optionBoundaries = buildOptionBoundaries(data, codeKey, labelKey);

  // Header block
  lines.push({ text: 'V B', kind: 'begin' });
  lines.push({ text: `V [${config.variableName}] ${config.questionLabel}`.trimEnd(), kind: 'variable' });
  lines.push({ text: `T ${config.tableType}`.trimEnd(), kind: 'table' });

  // Body — one R (+ optional Y) per row, in original order
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const codeStr = String(row[codeKey] ?? '').trim();
    const label = String(row[labelKey] ?? '').trim();
    const level = detectLevel(label.toUpperCase());

    if (level > 0) {
      // Marker row → bold label + aggregated netting script
      lines.push({ text: `R <b>${label}</b>`, kind: 'net-label' });
      const ranges = buildRangesForMarker(data, i, codeKey, labelKey, STOPS[level]);
      const parts: string[] = [];
      for (const range of ranges) {
        for (const q of questions) parts.push(`$${q}/${range}`);
      }
      lines.push({ text: `Y ${parts.join('+')}`.trimEnd(), kind: 'net-script' });
    } else if (codeStr !== '') {
      // Code row → label + single-code script
      lines.push({ text: `R ${label}`, kind: 'code-label' });
      const parts = questions.map(q => `$${q}/${codeStr}`);
      lines.push({ text: `Y ${parts.join('+')}`.trimEnd(), kind: 'code-script' });

      // Add the option boundaries directly below the first/last code scripts.
      for (const _boundary of optionBoundaries.opens.get(i) ?? []) {
        lines.push({ text: 'O R', kind: 'option-open' });
      }
      for (const _boundary of optionBoundaries.closes.get(i) ?? []) {
        lines.push({ text: 'O E', kind: 'option-end' });
      }
    } else {
      // Label-only row (no code, not a marker) → R only
      lines.push({ text: `R ${label}`, kind: 'plain' });
    }
  }

  // V E closes each question block
  lines.push({ text: 'V E', kind: 'end' });

  return lines;
}

/* ─── Excel export (one sheet per tab) ────── */

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

export function exportQpsWorkbook(specs: QpsSheetSpec[], originalFileName: string): void {
  debugLog('QPS', `Exporting QPS workbook with ${specs.length} sheet(s)`);

  const workbook = XLSX.utils.book_new();
  const used = new Set<string>();

  for (const spec of specs) {
    const tabLines = generateQpsLines(spec.data, spec.codeKey, spec.labelKey, spec.config);
    // Wrap each sheet: <<< Project Entries >>> at top, Z at bottom
    const lines: QpsLine[] = [
      { text: '<<< Project Entries >>>', kind: 'title' },
      ...tabLines,
      { text: 'Z', kind: 'terminator' },
    ];
    // Single column A, one line per row
    const matrix = lines.map(l => [l.text]);
    const worksheet = XLSX.utils.aoa_to_sheet(matrix);
    worksheet['!cols'] = [{ wch: 100 }];

    const sheetName = sanitizeSheetName(`${spec.name}_QPS`, used);
    XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);

    debugLog('QPS', `  Sheet "${sheetName}": ${lines.length} lines`);
  }

  const baseName = originalFileName.replace(/\.(xlsx|xls|csv)$/i, '') || 'output';
  XLSX.writeFile(workbook, `${baseName}_QPS.xlsx`, { bookType: 'xlsx' });

  debugLog('QPS', 'QPS XLSX download triggered');
}

/* ─── TXT export (all tabs combined) ──────── */

/**
 * Build the combined QPS text (all tabs concatenated).
 *
 * Structure:
 *   <<< Project Entries >>>        ← first line, once
 *   V B                            ← tab 1 block
 *   V [var] label
 *   T M
 *   R/Y …
 *   V E                            ← end of tab 1
 *   V B                            ← tab 2 block
 *   …
 *   V E                            ← end of last tab
 *   Z                              ← terminator, once at end
 */
export function buildCombinedQpsLines(specs: QpsSheetSpec[]): QpsLine[] {
  const all: QpsLine[] = [];

  // Leading title line
  all.push({ text: '<<< Project Entries >>>', kind: 'title' });

  // Each tab contributes its full V B … V E block
  specs.forEach(spec => {
    all.push(...generateQpsLines(spec.data, spec.codeKey, spec.labelKey, spec.config));
  });

  // Final terminator
  all.push({ text: 'Z', kind: 'terminator' });

  return all;
}

export function exportQpsCombinedTxt(specs: QpsSheetSpec[], originalFileName: string): void {
  const lines = buildCombinedQpsLines(specs);
  const content = lines.map(l => l.text).join('\r\n');

  const baseName = originalFileName.replace(/\.(xlsx|xls|csv)$/i, '') || 'output';
  const outputFileName = `${baseName}_QPS_combined.txt`;

  const blob = new Blob([content], { type: 'text/plain;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = outputFileName;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);

  debugLog('QPS', `QPS combined TXT download triggered: ${outputFileName} (${lines.length} lines)`);
}

/* ─── Preview helper ─────────────────────── */

export function buildQpsPreview(
  rows: RowData[],
  codeKey: string,
  labelKey: string,
  config: QpsConfig,
  maxLines = 40
): { lines: QpsLine[]; hiddenCount: number } {
  const tabLines = generateQpsLines(rows, codeKey, labelKey, config);
  // Wrap with title and terminator to match actual output
  const all: QpsLine[] = [
    { text: '<<< Project Entries >>>', kind: 'title' },
    ...tabLines,
    { text: 'Z', kind: 'terminator' },
  ];
  if (all.length <= maxLines) return { lines: all, hiddenCount: 0 };
  // Show first portion + last 3 (V E, Z)
  const headCount = maxLines - 3;
  const head = all.slice(0, headCount);
  const tail = all.slice(-3);
  return { lines: [...head, ...tail], hiddenCount: all.length - headCount - 3 };
}
