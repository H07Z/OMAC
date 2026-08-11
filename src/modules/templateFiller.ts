/**
 * Summary Template Filler
 *
 * Browser equivalent of the Hedonics / Just Right transfer part of the VBA.
 *
 * Source workbook shape expected from your screenshots:
 * - Wide source summary tables containing question headers across columns.
 * - Product labels under each question (PRODUCT X, PRODUCT Y, etc.).
 * - Base/sample row above the % row.
 * - Metric rows such as Top 2 Box (NET), Like extremely, Mean score,
 *   Bottom 2 Box (NET), and a "just right" option row.
 *
 * Template workbook shape:
 * - Question labels in the left label column.
 * - Header bands identifying PANEL 1 / PANEL 2, PRODUCT X / PRODUCT Y,
 *   and the metric column (TOP BOX SCORE, TOP-2 BOX SCORE, MEAN SCORE,
 *   Much/Some too strong, Just Right, Much/Some too weak).
 *
 * Transfer behavior:
 * - Normal question rows: copy source value row and significance marker row.
 * - Blue/base rows: copy source base row based on the question label beneath.
 * - All respondents row: copy the base from the first matched question.
 */

import * as XLSX from 'xlsx';
import { debugLog } from '../utils/debug';

type Mode = 'HD' | 'JR';
type MetricKey = 'topbox' | 'top2box' | 'mean' | 'toostrong' | 'justright' | 'tooweak';

export interface SourceCellValue {
  value: string | number;
  marker?: string;
}

interface SourceQuestion {
  baseByProduct: Record<string, SourceCellValue>;
  metricsByProduct: Record<string, Partial<Record<MetricKey, SourceCellValue>>>;
}

export interface ParsedSourceSummary {
  data: Record<Mode, Record<number, Record<string, SourceQuestion>>>;
  log: string[];
}

export interface FilledCell {
  row: number;
  col: number;
  value: string | number;
  displayValue?: string;
  marker?: string;
  questionCode: string;
  panel: number;
  product: string;
  metric: string;
}

export interface MatchStats {
  filledCells: FilledCell[];
  unmatchedRows: string[];
  log: string[];
}

interface TargetColumn {
  col: number;
  panel: number;
  product: string;
  metric: MetricKey;
  mode: Mode;
}

interface TargetRow {
  row: number;
  label: string;
  questionCode: string;
}

function normalizeText(value: unknown): string {
  return String(value ?? '').trim();
}

function normalizeKey(value: unknown): string {
  return normalizeText(value).toLowerCase().replace(/[^a-z0-9]+/g, '');
}

export function extractQuestionCode(text: string): string | null {
  const match = normalizeText(text).match(/Q\d+(\.\d+)?/i);
  return match ? match[0].toUpperCase() : null;
}

function isNumericLike(value: unknown): boolean {
  const text = normalizeText(value).replace(/%$/, '');
  return text !== '' && !isNaN(Number(text));
}

function getCell(ws: XLSX.WorkSheet, row: number, col: number): XLSX.CellObject | undefined {
  const direct = ws[XLSX.utils.encode_cell({ r: row, c: col })] as XLSX.CellObject | undefined;
  if (direct && direct.v !== undefined && direct.v !== null && String(direct.v).trim() !== '') return direct;

  const merges = ws['!merges'] || [];
  for (const merge of merges) {
    if (row >= merge.s.r && row <= merge.e.r && col >= merge.s.c && col <= merge.e.c) {
      return ws[XLSX.utils.encode_cell({ r: merge.s.r, c: merge.s.c })] as XLSX.CellObject | undefined;
    }
  }
  return direct;
}

function getText(ws: XLSX.WorkSheet, row: number, col: number): string {
  const cell = getCell(ws, row, col);
  return normalizeText(cell?.w ?? cell?.v ?? '');
}

function getRawValue(ws: XLSX.WorkSheet, row: number, col: number): string | number | null {
  const cell = getCell(ws, row, col);
  if (!cell || cell.v === undefined || cell.v === null || String(cell.v).trim() === '') return null;
  return typeof cell.v === 'number' ? cell.v : normalizeText(cell.w ?? cell.v);
}

function rowLabel(ws: XLSX.WorkSheet, row: number, maxCol = 6): string {
  for (let c = 0; c < maxCol; c++) {
    const text = getText(ws, row, c);
    if (text) return text;
  }
  return '';
}

function isSignificanceMarker(value: string): boolean {
  return /^[A-Za-z]{1,2}$/.test(value.trim());
}

/**
 * Source marker rows have blank label cells but may contain significance
 * letters in early data columns. Ignore standalone marker-looking cells
 * when deciding if a row has a metric label. This is required for Just Right
 * markers to be copied the same way as Hedonics markers.
 */
function sourceDescriptorIgnoringMarkers(ws: XLSX.WorkSheet, row: number, maxCol = 8): string {
  for (let c = 0; c < maxCol; c++) {
    const text = getText(ws, row, c);
    if (text && !isSignificanceMarker(text)) return text;
  }
  return '';
}

function categorizeSourceMetric(label: string, mode: Mode): MetricKey | null {
  const normalizedLabel = label.toLowerCase().trim();
  const key = normalizeKey(label);

  // IMPORTANT FOR HEDONICS:
  // TOP BOX SCORE must come from the source row containing "Like extremely".
  // Do not map it from a generic header row named "Top Box Score".
  const isLikeExtremely = /(^|\b)like\s+extremely\b/.test(normalizedLabel)
    && !/(^|\b)dislike\s+extremely\b/.test(normalizedLabel);
  const isLikeExtremelyKey = (key === 'likeextremely' || key.startsWith('likeextremely'))
    && !key.startsWith('dislikeextremely');

  if (mode === 'HD' && (isLikeExtremely || isLikeExtremelyKey)) {
    return 'topbox';
  }

  if (key.includes('top2box') || key.includes('top2boxnet') || normalizedLabel.includes('top 2 box')) return mode === 'JR' ? 'toostrong' : 'top2box';
  if (key.includes('bottom2box') || key.includes('bottom2boxnet') || normalizedLabel.includes('bottom 2 box')) return 'tooweak';
  if (key.includes('meanscore') || normalizedLabel.includes('mean score')) return 'mean';
  if (key.includes('justright') || key.includes('noneedtochange') || normalizedLabel.includes('just right')) return 'justright';

  return null;
}

/**
 * Split a source value into the numeric value and optional significance marker.
 *
 * Some workbooks store the marker in the cell below; some store it in the same
 * cell as a line break (e.g. "90.0\nB"). In both cases the marker's exact case
 * must be preserved: "b" stays lower-case, "B" stays upper-case.
 */
function splitValueAndMarker(rawValue: string | number, rawMarker: string | number | null): SourceCellValue {
  if (typeof rawValue === 'number') {
    const marker = rawMarker === null ? '' : normalizeText(rawMarker);
    return marker ? { value: rawValue, marker } : { value: rawValue };
  }

  const parts = String(rawValue).split(/\r?\n/).map(part => part.trim()).filter(Boolean);
  const valueText = parts[0] ?? '';
  const inlineMarker = parts.slice(1).join('\n').trim();
  const belowMarker = rawMarker === null ? '' : normalizeText(rawMarker);
  const marker = inlineMarker || belowMarker;

  const numeric = Number(valueText);
  const value = valueText !== '' && !isNaN(numeric) && /^-?\d+(\.\d+)?$/.test(valueText)
    ? numeric
    : valueText;

  return marker ? { value, marker } : { value };
}

function detectModeAround(ws: XLSX.WorkSheet, row: number): Mode | null {
  for (let r = row; r >= Math.max(0, row - 20); r--) {
    const line = Array.from({ length: 12 }, (_, c) => getText(ws, r, c).toUpperCase()).join(' ');
    if (line.includes('JUST RIGHT')) return 'JR';
    if (line.includes('HEDONICS') || line.includes('LIKING')) return 'HD';
  }
  return null;
}

function detectPanelAround(ws: XLSX.WorkSheet, row: number, colStart: number): number {
  for (let r = Math.max(0, row - 8); r <= row + 3; r++) {
    for (let c = Math.max(0, colStart - 2); c <= colStart + 3; c++) {
      const text = getText(ws, r, c).toUpperCase();
      if (text.includes('PANEL 2') || text.includes('TABLE 5') || text.includes('TABLE 6')) return 2;
      if (text.includes('PANEL 1') || text.includes('TABLE 3') || text.includes('TABLE 4')) return 1;
    }
  }
  return 1;
}

function collectQuestionHeaderColumns(ws: XLSX.WorkSheet, row: number, lastCol: number): Array<{ col: number; code: string }> {
  const out: Array<{ col: number; code: string }> = [];
  for (let c = 0; c <= lastCol; c++) {
    // Use only physically populated header cells here. Propagating a merged
    // question header across every cell in its merge would create duplicate
    // question starts and break the two-brand column block boundaries.
    const direct = ws[XLSX.utils.encode_cell({ r: row, c })] as XLSX.CellObject | undefined;
    const code = extractQuestionCode(normalizeText(direct?.w ?? direct?.v ?? ''));
    if (code) out.push({ col: c, code });
  }
  return out;
}

function productFallback(index: number): string {
  if (index === 0) return 'PRODUCT X';
  if (index === 1) return 'PRODUCT Y';
  return `PRODUCT ${index + 1}`;
}

function findBaseRow(ws: XLSX.WorkSheet, startRow: number, dataCols: number[]): number {
  for (let r = startRow; r <= startRow + 8; r++) {
    const numericCount = dataCols.filter(c => isNumericLike(getText(ws, r, c))).length;
    const nextPercentCount = dataCols.filter(c => getText(ws, r + 1, c).includes('%')).length;
    const label = rowLabel(ws, r).toUpperCase();
    if (numericCount >= Math.min(2, dataCols.length) && (nextPercentCount > 0 || label.includes('PANEL') || label.includes('ALL'))) {
      return r;
    }
  }
  return startRow + 4;
}

export function parseSourceSummary(ws: XLSX.WorkSheet): ParsedSourceSummary {
  const ref = XLSX.utils.decode_range(ws['!ref'] || 'A1:A1');
  const parsed: ParsedSourceSummary = { data: { HD: {}, JR: {} }, log: [] };

  for (let r = ref.s.r; r <= ref.e.r; r++) {
    const qHeaders = collectQuestionHeaderColumns(ws, r, ref.e.c);
    if (qHeaders.length < 2) continue;

    const mode = detectModeAround(ws, r);
    if (!mode) continue;

    const panel = detectPanelAround(ws, r, qHeaders[0].col);
    const productRow = r + 1;
    const allDataCols: number[] = [];

    parsed.log.push(`Detected ${mode} panel ${panel} source header row ${r + 1} with ${qHeaders.length} questions.`);

    for (let h = 0; h < qHeaders.length; h++) {
      const startCol = qHeaders[h].col;
      const endCol = h + 1 < qHeaders.length ? qHeaders[h + 1].col - 1 : ref.e.c;
      for (let c = startCol; c <= endCol; c++) {
        const product = (getText(ws, productRow, c).toUpperCase() || productFallback(c - startCol));
        if (product.includes('PRODUCT')) allDataCols.push(c);
      }
    }

    const baseRow = findBaseRow(ws, r + 1, allDataCols);
    const percentRow = baseRow + 1;
    let metricStart = percentRow + 1;

    while (metricStart <= ref.e.r && rowLabel(ws, metricStart) === '') metricStart++;

    const nextHeaderRow = (() => {
      for (let nr = r + 1; nr <= ref.e.r; nr++) {
        if (collectQuestionHeaderColumns(ws, nr, ref.e.c).length >= 2) return nr;
        const label = rowLabel(ws, nr).toUpperCase();
        if (nr > metricStart && (label.includes('RETURN TO CONTENTS') || label.includes('HEDONICS /') || label.includes('JUST RIGHT SUMMARY'))) return nr;
      }
      return ref.e.r + 1;
    })();

    if (!parsed.data[mode][panel]) parsed.data[mode][panel] = {};

    for (let h = 0; h < qHeaders.length; h++) {
      const { col: startCol, code } = qHeaders[h];
      const endCol = h + 1 < qHeaders.length ? qHeaders[h + 1].col - 1 : ref.e.c;
      if (!parsed.data[mode][panel][code]) {
        parsed.data[mode][panel][code] = { baseByProduct: {}, metricsByProduct: {} };
      }

      for (let c = startCol; c <= endCol; c++) {
        const product = (getText(ws, productRow, c).toUpperCase() || productFallback(c - startCol));
        if (!product.includes('PRODUCT')) continue;

        const base = getRawValue(ws, baseRow, c);
        if (base !== null) parsed.data[mode][panel][code].baseByProduct[product] = { value: base };
        if (!parsed.data[mode][panel][code].metricsByProduct[product]) parsed.data[mode][panel][code].metricsByProduct[product] = {};

        for (let mr = metricStart; mr < nextHeaderRow; mr++) {
          const label = rowLabel(ws, mr);
          if (!label) continue;
          const metric = categorizeSourceMetric(label, mode);
          if (!metric) continue;

          const rawValue = getRawValue(ws, mr, c);
          if (rawValue === null) continue;

          const nextLabel = sourceDescriptorIgnoringMarkers(ws, mr + 1);
          const rawMarker = nextLabel === '' ? getRawValue(ws, mr + 1, c) : null;
          parsed.data[mode][panel][code].metricsByProduct[product][metric] = splitValueAndMarker(rawValue, rawMarker);
        }
      }
    }
  }

  return parsed;
}

export function detectSummaryTemplateMode(ws: XLSX.WorkSheet): Mode | null {
  const ref = XLSX.utils.decode_range(ws['!ref'] || 'A1:A1');
  for (let r = ref.s.r; r <= Math.min(ref.e.r, 30); r++) {
    const row = Array.from({ length: Math.min(ref.e.c + 1, 15) }, (_, c) => getText(ws, r, c).toUpperCase()).join(' ');
    if (row.includes('JUST RIGHT')) return 'JR';
    if (row.includes('HEDONICS') || row.includes('LIKING')) return 'HD';
  }
  return null;
}

function parseTargetColumns(ws: XLSX.WorkSheet, mode: Mode): TargetColumn[] {
  const ref = XLSX.utils.decode_range(ws['!ref'] || 'A1:A1');
  const columns: TargetColumn[] = [];
  const scanRows = Math.min(ref.e.r, 20);

  for (let c = ref.s.c; c <= ref.e.c; c++) {
    let panel = 1;
    let product = '';
    let metric: MetricKey | null = null;

    for (let r = ref.s.r; r <= scanRows; r++) {
      const text = getText(ws, r, c).toUpperCase();
      if (text.includes('PANEL 2')) panel = 2;
      if (text.includes('PANEL 1')) panel = 1;
      if (text.includes('PRODUCT X')) product = 'PRODUCT X';
      if (text.includes('PRODUCT Y')) product = 'PRODUCT Y';

      if (mode === 'HD') {
        if (text.includes('TOP BOX SCORE') && !text.includes('TOP-2')) metric = 'topbox';
        if (text.includes('TOP-2 BOX SCORE')) metric = 'top2box';
        if (text.includes('MEAN SCORE')) metric = 'mean';
      } else {
        if (text.includes('TOO STRONG') || text.includes('(4-5)')) metric = 'toostrong';
        if (text.includes('JUST RIGHT') || text.includes('(3)')) metric = 'justright';
        if (text.includes('TOO WEAK') || text.includes('(1-2)')) metric = 'tooweak';
      }
    }

    if (product && metric) columns.push({ col: c, panel, product, metric, mode });
  }

  return columns;
}

function parseTargetRows(ws: XLSX.WorkSheet): { questionRows: TargetRow[]; allRespondentsRows: number[]; baseRows: Array<{ row: number; questionCode: string; label: string }> } {
  const ref = XLSX.utils.decode_range(ws['!ref'] || 'A1:A1');
  const questionRows: TargetRow[] = [];
  const allRespondentsRows: number[] = [];
  const baseRows: Array<{ row: number; questionCode: string; label: string }> = [];

  for (let r = ref.s.r; r <= ref.e.r; r++) {
    let labelCol = -1;
    let label = '';
    for (let c = ref.s.c; c <= Math.min(ref.e.c, 5); c++) {
      const text = getText(ws, r, c);
      if (text) { label = text; labelCol = c; break; }
    }
    if (labelCol === -1) continue;

    const qCode = extractQuestionCode(label);
    if (qCode) questionRows.push({ row: r, questionCode: qCode, label });
    if (normalizeKey(label).includes('allrespondents')) allRespondentsRows.push(r);

    const nextLabel = rowLabel(ws, r + 1);
    const nextQ = extractQuestionCode(nextLabel);
    if (!qCode && nextQ && !normalizeKey(label).includes('allrespondents')) {
      baseRows.push({ row: r, questionCode: nextQ, label });
    }
  }

  return { questionRows, allRespondentsRows, baseRows };
}

type CellStyle = {
  fill?: { fgColor: { rgb: string } };
  font?: { bold?: boolean; color?: { rgb: string }; name?: string; sz?: number };
  alignment?: { horizontal?: string; vertical?: string; wrapText?: boolean };
  border?: {
    top?: { style: string; color: { rgb: string } };
    bottom?: { style: string; color: { rgb: string } };
    left?: { style: string; color: { rgb: string } };
    right?: { style: string; color: { rgb: string } };
  };
  numFmt?: string;
};

function thinBorder(): CellStyle['border'] {
  const edge = { style: 'thin', color: { rgb: 'B7B7B7' } };
  return { top: edge, bottom: edge, left: edge, right: edge };
}

function styleCell(ws: XLSX.WorkSheet, row: number, col: number, style: CellStyle): void {
  const ref = XLSX.utils.encode_cell({ r: row, c: col });
  const existing = (ws[ref] || { t: 's', v: '' }) as XLSX.CellObject & { s?: CellStyle };
  ws[ref] = {
    ...existing,
    s: {
      ...(existing.s || {}),
      ...style,
      font: { ...(existing.s?.font || {}), ...(style.font || {}) },
      fill: style.fill || existing.s?.fill,
      alignment: { ...(existing.s?.alignment || {}), ...(style.alignment || {}) },
      border: style.border || existing.s?.border,
      numFmt: style.numFmt || existing.s?.numFmt,
    },
  };
}

function formatMetricDisplay(
  value: string | number,
  metric: MetricKey | 'base' | string
): { rawNum: number | null; displayStr: string; numFmt?: string } {
  if (value === null || value === undefined || value === '') {
    return { rawNum: null, displayStr: '' };
  }

  const strVal = String(value).trim();
  const num = Number(strVal);

  if (!isNaN(num) && /^-?\d+(\.\d+)?$/.test(strVal)) {
    if (metric === 'mean') {
      return {
        rawNum: num,
        displayStr: num.toFixed(2),
        numFmt: '0.00',
      };
    } else if (
      metric === 'topbox' ||
      metric === 'top2box' ||
      metric === 'toostrong' ||
      metric === 'justright' ||
      metric === 'tooweak'
    ) {
      return {
        rawNum: num,
        displayStr: num.toFixed(1),
        numFmt: '0.0',
      };
    }
    return {
      rawNum: num,
      displayStr: strVal,
    };
  }

  return { rawNum: null, displayStr: strVal };
}

function writeCell(
  ws: XLSX.WorkSheet,
  row: number,
  col: number,
  value: string | number,
  style?: CellStyle,
  metric?: MetricKey | 'base' | string
): string {
  const ref = XLSX.utils.encode_cell({ r: row, c: col });
  const existing = (ws[ref] || {}) as XLSX.CellObject & { s?: CellStyle };

  const { rawNum, displayStr, numFmt } = formatMetricDisplay(value, metric || '');

  let next: XLSX.CellObject & { s?: CellStyle };

  if (rawNum !== null) {
    next = {
      ...existing,
      t: 'n',
      v: rawNum, // keeps full float value intact!
      w: displayStr, // formatted string for display (1 or 2 decimals)
      z: numFmt || style?.numFmt,
    };
  } else {
    next = {
      ...existing,
      t: 's',
      v: displayStr,
      w: displayStr,
    };
  }

  const mergedStyle = {
    ...(style || {}),
    numFmt: numFmt || style?.numFmt || existing.s?.numFmt,
  };

  if (mergedStyle) {
    next.s = {
      ...(existing.s || {}),
      ...mergedStyle,
      font: { ...(existing.s?.font || {}), ...(mergedStyle.font || {}) },
      fill: mergedStyle.fill || existing.s?.fill,
      alignment: { ...(existing.s?.alignment || {}), ...(mergedStyle.alignment || {}) },
      border: mergedStyle.border || existing.s?.border,
      numFmt: mergedStyle.numFmt,
    };
  }

  ws[ref] = next;
  return displayStr;
}

function applySummaryFormatting(
  ws: XLSX.WorkSheet,
  mode: Mode,
  columns: TargetColumn[],
  rows: { questionRows: TargetRow[]; allRespondentsRows: number[]; baseRows: Array<{ row: number; questionCode: string; label: string }> },
): void {
  const ref = XLSX.utils.decode_range(ws['!ref'] || 'A1:A1');
  const dataCols = columns.map(c => c.col);
  const minDataCol = dataCols.length ? Math.min(...dataCols) : 1;
  const maxDataCol = dataCols.length ? Math.max(...dataCols) : ref.e.c;
  const labelCol = 0;
  const paleYellow = 'FFFDE9A8';
  const softYellow = 'FFFFF2CC';
  const segmentBlue = 'FF00B0F0';
  const white = 'FFFFFFFF';
  const black = 'FF000000';
  const red = 'FFFF0000';

  // Header band formatting
  for (let r = 0; r <= 6; r++) {
    for (let c = labelCol; c <= maxDataCol; c++) {
      const text = getText(ws, r, c).toUpperCase();
      const isHeaderBand =
        text.includes('PANEL') ||
        text.includes('TOP BOX') ||
        text.includes('TOP-2') ||
        text.includes('MEAN SCORE') ||
        text.includes('PRODUCT') ||
        text.includes('TOO STRONG') ||
        text.includes('JUST RIGHT') ||
        text.includes('TOO WEAK') ||
        text.includes('(A)') ||
        text.includes('(B)');

      if (isHeaderBand || c >= minDataCol) {
        styleCell(ws, r, c, {
          fill: { fgColor: { rgb: softYellow } },
          font: { bold: true, color: { rgb: black }, name: 'Calibri', sz: 10 },
          alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
          border: thinBorder(),
        });
      }
    }
  }

  // All respondents and percent row
  for (const rowIndex of rows.allRespondentsRows) {
    for (let c = labelCol; c <= maxDataCol; c++) {
      styleCell(ws, rowIndex, c, {
        fill: { fgColor: { rgb: paleYellow } },
        font: { bold: true, color: { rgb: black }, name: 'Calibri', sz: 10 },
        alignment: { horizontal: c === labelCol ? 'left' : 'center', vertical: 'center' },
        border: thinBorder(),
      });
      styleCell(ws, rowIndex + 1, c, {
        fill: { fgColor: { rgb: softYellow } },
        font: { bold: false, color: { rgb: black }, name: 'Calibri', sz: 10 },
        alignment: { horizontal: c === labelCol ? 'left' : 'center', vertical: 'center' },
        border: thinBorder(),
      });
    }
  }

  // Question rows + marker rows
  for (const qRow of rows.questionRows) {
    const isOverall = /Q16/i.test(qRow.questionCode) || /overall/i.test(qRow.label);
    for (let c = labelCol; c <= maxDataCol; c++) {
      styleCell(ws, qRow.row, c, {
        fill: { fgColor: { rgb: paleYellow } },
        font: {
          bold: c === labelCol ? true : isOverall,
          color: { rgb: black },
          name: 'Calibri',
          sz: isOverall && c > labelCol ? 11 : 10,
        },
        alignment: { horizontal: c === labelCol ? 'left' : 'center', vertical: 'center' },
        border: thinBorder(),
      });
      styleCell(ws, qRow.row + 1, c, {
        fill: { fgColor: { rgb: white } },
        font: { bold: true, color: { rgb: red }, name: 'Calibri', sz: 10 },
        alignment: { horizontal: c === labelCol ? 'left' : 'center', vertical: 'center' },
        border: thinBorder(),
      });
    }
  }

  // Blue segment/base rows
  for (const baseRow of rows.baseRows) {
    for (let c = labelCol; c <= maxDataCol; c++) {
      styleCell(ws, baseRow.row, c, {
        fill: { fgColor: { rgb: segmentBlue } },
        font: { bold: true, color: { rgb: white }, name: 'Calibri', sz: 10 },
        alignment: { horizontal: c === labelCol ? 'left' : 'center', vertical: 'center' },
        border: thinBorder(),
      });
    }
  }

  // Column widths and row heights
  ws['!cols'] = Array.from({ length: maxDataCol + 1 }, (_, index) => ({
    wch: index === labelCol ? 42 : 10,
  }));
  ws['!rows'] = Array.from({ length: ref.e.r + 1 }, (_, index) => {
    if (index <= 5) return { hpt: 18 };
    if (rows.questionRows.some(q => q.row + 1 === index)) return { hpt: 14 };
    return { hpt: 16 };
  });

  // Freeze panes just under the header band when possible
  const freezeAt = rows.allRespondentsRows[0] ?? 7;
  ws['!freeze'] = { xSplit: 1, ySplit: freezeAt, topLeftCell: XLSX.utils.encode_cell({ r: freezeAt, c: 1 }), activePane: 'bottomRight', state: 'frozen' };

  // Keep mode available for future mode-specific polish
  void mode;
}

function lookupQuestion(parsed: ParsedSourceSummary, mode: Mode, panel: number, questionCode: string): SourceQuestion | null {
  return parsed.data[mode][panel]?.[questionCode] || null;
}

function getMetricValue(q: SourceQuestion, product: string, metric: MetricKey): SourceCellValue | undefined {
  return q.metricsByProduct[product]?.[metric];
}

export function fillSummaryTemplateWorkbook(
  sourceSheet: XLSX.WorkSheet,
  templateWorkbook: XLSX.WorkBook,
  templateSheetName: string
): { workbook: XLSX.WorkBook; stats: MatchStats; parsed: ParsedSourceSummary } {
  const parsed = parseSourceSummary(sourceSheet);
  const workbook = JSON.parse(JSON.stringify(templateWorkbook)) as XLSX.WorkBook;
  const ws = workbook.Sheets[templateSheetName];
  if (!ws) throw new Error(`Template sheet "${templateSheetName}" not found.`);

  const mode = detectSummaryTemplateMode(ws);
  if (!mode) throw new Error('Unable to detect template type (HEDONICS/LIKING or JUST RIGHT).');

  const columns = parseTargetColumns(ws, mode);
  const rows = parseTargetRows(ws);
  const stats: MatchStats = { filledCells: [], unmatchedRows: [], log: [...parsed.log] };

  stats.log.push(`Template mode: ${mode}`);
  stats.log.push(`Mapped ${columns.length} target data columns and ${rows.questionRows.length} target question rows.`);

  const firstKnownQuestion = rows.questionRows.find(row => parsed.data[mode][1]?.[row.questionCode] || parsed.data[mode][2]?.[row.questionCode]);

  // All respondents / base row
  for (const rowIndex of rows.allRespondentsRows) {
    // The row directly below All respondents must ALWAYS be only "%"
    // for both Hedonics and Just Right templates.
    for (const col of columns) {
      writeCell(ws, rowIndex + 1, col.col, '%', {
        fill: { fgColor: { rgb: 'FFFFF2CC' } },
        font: { bold: false, color: { rgb: 'FF000000' }, name: 'Calibri', sz: 10 },
        alignment: { horizontal: 'center', vertical: 'center' },
        border: thinBorder(),
      });
    }

    for (const col of columns) {
      const q = firstKnownQuestion ? lookupQuestion(parsed, mode, col.panel, firstKnownQuestion.questionCode) : null;
      const base = q?.baseByProduct[col.product];
      if (base) {
        writeCell(ws, rowIndex, col.col, base.value, {
          fill: { fgColor: { rgb: 'FFFDE9A8' } },
          font: { bold: true, color: { rgb: 'FF000000' }, name: 'Calibri', sz: 10 },
          alignment: { horizontal: 'center', vertical: 'center' },
          border: thinBorder(),
        });
        stats.filledCells.push({ row: rowIndex, col: col.col, value: base.value, questionCode: 'ALL', panel: col.panel, product: col.product, metric: 'base' });
      }
    }
  }

  // Blue/base rows driven by the question below the base row
  for (const baseRow of rows.baseRows) {
    for (const col of columns) {
      const q = lookupQuestion(parsed, mode, col.panel, baseRow.questionCode);
      const base = q?.baseByProduct[col.product];
      if (base) {
        writeCell(ws, baseRow.row, col.col, base.value, {
          fill: { fgColor: { rgb: 'FF00B0F0' } },
          font: { bold: true, color: { rgb: 'FFFFFFFF' }, name: 'Calibri', sz: 10 },
          alignment: { horizontal: 'center', vertical: 'center' },
          border: thinBorder(),
        });
        stats.filledCells.push({ row: baseRow.row, col: col.col, value: base.value, questionCode: baseRow.questionCode, panel: col.panel, product: col.product, metric: 'base' });
      }
    }
  }

  // Normal question rows
  for (const row of rows.questionRows) {
    let any = false;
    for (const col of columns) {
      const q = lookupQuestion(parsed, mode, col.panel, row.questionCode);
      const found = q ? getMetricValue(q, col.product, col.metric) : undefined;
      if (!found) continue;

      const displayVal = writeCell(
        ws,
        row.row,
        col.col,
        found.value,
        {
          fill: { fgColor: { rgb: 'FFFDE9A8' } },
          font: { bold: /Q16/i.test(row.questionCode), color: { rgb: 'FF000000' }, name: 'Calibri', sz: 10 },
          alignment: { horizontal: 'center', vertical: 'center' },
          border: thinBorder(),
        },
        col.metric
      );

      // For BOTH Hedonics and Just Right, if the source has a significance
      // letter below the value, place that exact letter (case preserved) in
      // the row directly below the populated value.
      if (found.marker) {
        writeCell(ws, row.row + 1, col.col, found.marker, {
          fill: { fgColor: { rgb: 'FFFFFFFF' } },
          font: { bold: true, color: { rgb: 'FFFF0000' }, name: 'Calibri', sz: 10 },
          alignment: { horizontal: 'center', vertical: 'center' },
          border: thinBorder(),
        });
      }
      stats.filledCells.push({
        row: row.row,
        col: col.col,
        value: found.value,
        displayValue: displayVal,
        marker: found.marker,
        questionCode: row.questionCode,
        panel: col.panel,
        product: col.product,
        metric: col.metric,
      });
      any = true;
    }
    if (!any) stats.unmatchedRows.push(row.questionCode);
  }

  applySummaryFormatting(ws, mode, columns, rows);

  // Final safeguard: both Hedonics and Just Right must have only "%" in
  // the row immediately below every "All respondents" row. This runs after
  // all population/formatting so it cannot be overwritten by base values.
  for (const rowIndex of rows.allRespondentsRows) {
    for (const col of columns) {
      writeCell(ws, rowIndex + 1, col.col, '%', {
        fill: { fgColor: { rgb: 'FFFFF2CC' } },
        font: { bold: false, color: { rgb: 'FF000000' }, name: 'Calibri', sz: 10 },
        alignment: { horizontal: 'center', vertical: 'center' },
        border: thinBorder(),
      });
    }
  }

  stats.log.push(`Filled ${stats.filledCells.length} cells. Unmatched question rows: ${stats.unmatchedRows.length}.`);
  stats.log.push('Applied Hedonics / Just Right summary table formatting.');
  debugLog('SummaryFiller', stats.log[stats.log.length - 1]);

  return { workbook, stats, parsed };
}
