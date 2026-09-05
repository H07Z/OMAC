import { useState, useCallback, useMemo, Fragment } from 'react';
import { RowData, LogEntry, FilterConfig, ProcessingStats, SortConfig } from './types';
import { debugLog, getDebugLogs } from './utils/debug';
import { readWorkbook, readCSV, loadOeSheet, getFileType } from './modules/excelLoader';
import { processData, autoDetectColumns, ProcessingConfig } from './modules/dataProcessor';
import { FORMAT_TEMPLATES, buildFormattedString } from './modules/formatTemplates';
import { validateFile } from './modules/validator';
import { applyFilters, getUniqueValues } from './modules/filterSearch';
import {
  exportWorkbook,
  exportCombinedTxt,
  buildExportPreview,
  buildCombinedTxtPreview,
  HeaderLines,
} from './modules/excelExporter';
import {
  QpsConfig,
  exportQpsWorkbook,
  exportQpsCombinedTxt,
  buildQpsPreview,
} from './modules/qpsScript';
import TemplateAppendPage from './TemplateAppendPage';
import SummaryTableFillerPage from './SummaryTableFillerPage';
import HeaderNav, { PageId } from './components/HeaderNav';
import TopExportBar from './components/TopExportBar';
import * as XLSX from 'xlsx';

/* ─── SVG ICON COMPONENTS ───────────────────────────────── */

const UploadIcon = () => (
  <svg className="w-10 h-10 mx-auto text-slate-400 dark:text-slate-500" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={1.5}>
    <path strokeLinecap="round" strokeLinejoin="round" d="M19.5 14.25v-2.625a3.375 3.375 0 00-3.375-3.375h-1.5A1.125 1.125 0 0113.5 7.125v-1.5a3.375 3.375 0 00-3.375-3.375H8.25m6.75 12l-3-3m0 0l-3 3m3-3v6m-1.5-15H5.625c-.621 0-1.125.504-1.125 1.125v17.25c0 .621.504 1.125 1.125 1.125h12.75c.621 0 1.125-.504 1.125-1.125V11.25a9 9 0 00-9-9z" />
  </svg>
);

const SpinnerIcon = () => (
  <svg className="w-5 h-5 animate-spin" fill="none" viewBox="0 0 24 24">
    <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" />
    <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z" />
  </svg>
);

const ChevronIcon = ({ direction }: { direction: 'asc' | 'desc' | 'none' }) => (
  <svg className="w-4 h-4 shrink-0" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
    {direction === 'asc' && <path strokeLinecap="round" strokeLinejoin="round" d="M5 15l7-7 7 7" />}
    {direction === 'desc' && <path strokeLinecap="round" strokeLinejoin="round" d="M19 9l-7 7-7-7" />}
    {direction === 'none' && <path strokeLinecap="round" strokeLinejoin="round" d="M8 9l4-4 4 4m0 6l-4 4-4-4" />}
  </svg>
);

/* ─── APPLICATION STATE ─────────────────────────────────── */

interface SheetState {
  name: string;
  questionCode: string;
  questionLabel: string;
  allHeaders: string[];
  outputHeaders: string[];
  originalData: RowData[];
  processedData: RowData[];
  filters: FilterConfig;
  stats: ProcessingStats;
  processingLog: LogEntry[];
  columnConfig: ProcessingConfig;
  processed: boolean;
}

interface AppState {
  workbook: XLSX.WorkBook | null;
  sheets: SheetState[];
  activeSheet: string;
  fileName: string;
  fileType: 'xlsx' | 'xls' | 'csv' | null;
  isProcessing: boolean;
  searchTerm: string;
  sortConfig: SortConfig;
  errors: string[];
}

const emptyStats: ProcessingStats = { originalRows: 0, removedRows: 0, insertedRows: 0, finalRows: 0 };

const initialState: AppState = {
  workbook: null,
  sheets: [],
  activeSheet: '',
  fileName: '',
  fileType: null,
  isProcessing: false,
  searchTerm: '',
  sortConfig: { column: '', direction: null },
  errors: [],
};

type FormatInputMap = Record<string, Record<string, string[]>>;

function buildSheetState(workbook: XLSX.WorkBook, name: string): SheetState {
  const sd = loadOeSheet(workbook, name);
  const config = autoDetectColumns(sd.headers);
  const outHeaders = sd.headers.filter(h => !config.excludedColumns.includes(h));
  const initFilters: FilterConfig = {};
  outHeaders.forEach(h => { initFilters[h] = []; });

  return {
    name,
    questionCode: sd.questionCode,
    questionLabel: sd.questionLabel,
    allHeaders: sd.headers,
    outputHeaders: outHeaders,
    originalData: sd.data,
    processedData: [],
    filters: initFilters,
    stats: { ...emptyStats, originalRows: sd.data.length },
    processingLog: [],
    columnConfig: config,
    processed: false,
  };
}

function buildQuestionLabel(questionCode: string, questionLabel: string): string {
  const code = questionCode.trim().replace(/[.\s]+$/, '');
  const label = questionLabel.trim();
  if (!code) return label;
  if (!label) return code;
  if (label.toUpperCase().startsWith(code.toUpperCase())) return label;
  return `${code}. ${label}`;
}

function initialFormatInputs(
  sheetName: string,
  questionCode = '',
  questionLabel = '',
  questionIndex = 1,
  highestCodeLength = 1,
): Record<string, string[]> {
  const qMatch = (questionCode || sheetName).match(/Q\d+(\.\d+)?/i);
  const code = qMatch ? qMatch[0].toUpperCase() : questionCode.trim();
  const safeCodeLength = Math.max(1, highestCodeLength);
  return {
    'q-pattern': [code, buildQuestionLabel(code, questionLabel)],
    // Workbook tab 1 -> L 1L{code length}R{matching count of 9s};
    // tab 2 -> L 2L...; and so on.
    'l-pattern': [
      String(Math.max(1, questionIndex)),
      String(safeCodeLength),
      '9'.repeat(safeCodeLength),
    ],
  };
}

/** Return the digit length of the highest numeric code in Column A. */
function getHighestCodeLength(sheet: SheetState): number {
  let highestNumber = Number.NEGATIVE_INFINITY;
  let highestRaw = '';

  for (const row of sheet.originalData) {
    const raw = String(row[sheet.columnConfig.columnAKey] ?? '').trim();
    if (!/^-?\d+(?:\.\d+)?$/.test(raw)) continue;

    const numeric = Number(raw);
    if (numeric > highestNumber) {
      highestNumber = numeric;
      highestRaw = raw;
    }
  }

  if (!highestRaw) return 1;
  const integerDigits = highestRaw.replace(/^-/, '').split('.')[0].length;
  return Math.max(1, integerDigits);
}

function initialQpsConfig(
  sheetName: string,
  questionCode = '',
  questionLabel = '',
): QpsConfig {
  const question = (questionCode || sheetName).trim().replace(/^V/i, '');
  const variableName = `V${question}`;
  // Questions must use Excel row 1, Column A (for example Q17).
  const questions = (questionCode || question).trim();
  return {
    variableName,
    questions,
    questionLabel: buildQuestionLabel(questionCode || question, questionLabel),
    tableType: 'M',
  };
}

const PAGE_SIZE = 100;

/* ─── MAIN COMPONENT ───────────────────────────────────── */

export default function App() {
  const [page, setPage] = useState<PageId>('processor');
  const [state, setState] = useState<AppState>(initialState);
  const [showDebug, setShowDebug] = useState(false);
  const [showLog, setShowLog] = useState(true);
  const [showColumnConfig, setShowColumnConfig] = useState(false);
  const [showExportPreview] = useState(true);
  const [previewMode, setPreviewMode] = useState<'xlsx' | 'txt'>('xlsx');
  const [dragActive, setDragActive] = useState(false);
  const [currentPage, setCurrentPage] = useState(0);

  const [formatInputs, setFormatInputs] = useState<FormatInputMap>({});
  const [openFormats, setOpenFormats] = useState<Record<string, boolean>>({});
  const [qpsConfigs, setQpsConfigs] = useState<Record<string, QpsConfig>>({});
  const [openQps, setOpenQps] = useState<Record<string, boolean>>({});
  const [showQpsPreview, setShowQpsPreview] = useState(true);

  const activeSheet = state.sheets.find(s => s.name === state.activeSheet) ?? state.sheets[0] ?? null;

  const updateSheet = useCallback((name: string, updater: (s: SheetState) => SheetState) => {
    setState(prev => ({
      ...prev,
      sheets: prev.sheets.map(s => (s.name === name ? updater(s) : s)),
    }));
  }, []);

  const getHeaderLines = useCallback((sheetName: string): HeaderLines => {
    const qTemplate = FORMAT_TEMPLATES.find(t => t.id === 'q-pattern');
    const lTemplate = FORMAT_TEMPLATES.find(t => t.id === 'l-pattern');
    const inputs = formatInputs[sheetName] || {};
    const qInputs = inputs['q-pattern'] || [];
    const lInputs = inputs['l-pattern'] || [];

    return {
      qLine: qTemplate && qInputs.some(v => v && v.trim() !== '') ? buildFormattedString(qTemplate, qInputs) : '',
      lLine: lTemplate && lInputs.some(v => v && v.trim() !== '') ? buildFormattedString(lTemplate, lInputs) : '',
    };
  }, [formatInputs]);

  const setFormatInput = useCallback((sheetName: string, templateId: string, index: number, value: string) => {
    setFormatInputs(prev => {
      const next = { ...prev };
      const sheetInputs = { ...(next[sheetName] || initialFormatInputs(sheetName)) };
      const templateInputs = [...(sheetInputs[templateId] || [])];
      templateInputs[index] = value;
      sheetInputs[templateId] = templateInputs;
      next[sheetName] = sheetInputs;
      return next;
    });
  }, []);

  const transformSheet = useCallback((sheet: SheetState, isActive: boolean): RowData[] => {
    let data = applyFilters(sheet.processedData, isActive ? state.searchTerm : '', sheet.filters, sheet.outputHeaders);

    if (isActive && state.sortConfig.column && state.sortConfig.direction) {
      const col = state.sortConfig.column;
      const dir = state.sortConfig.direction === 'asc' ? 1 : -1;
      data = [...data].sort((a, b) => {
        const valA = a[col]; const valB = b[col];
        const numA = typeof valA === 'number' ? valA : parseFloat(String(valA));
        const numB = typeof valB === 'number' ? valB : parseFloat(String(valB));
        if (!isNaN(numA) && !isNaN(numB)) return (numA - numB) * dir;
        return String(valA ?? '').localeCompare(String(valB ?? '')) * dir;
      });
    }
    return data;
  }, [state.searchTerm, state.sortConfig]);

  const activeData = useMemo(() => (activeSheet ? transformSheet(activeSheet, true) : []), [activeSheet, transformSheet]);

  const uniqueValues = useMemo(() => {
    const values: Record<string, string[]> = {};
    if (!activeSheet) return values;
    for (const header of activeSheet.outputHeaders) {
      values[header] = getUniqueValues(activeSheet.processedData, header);
    }
    return values;
  }, [activeSheet]);

  const totalPages = Math.max(1, Math.ceil(activeData.length / PAGE_SIZE));
  const safePage = Math.min(currentPage, totalPages - 1);
  const pagedData = activeData.slice(safePage * PAGE_SIZE, (safePage + 1) * PAGE_SIZE);

  const allProcessed = state.sheets.length > 0 && state.sheets.every(s => s.processed);

  const exportPreview = useMemo(() => {
    if (!activeSheet || !activeSheet.processed || activeData.length === 0) return { rows: [], hiddenCount: 0 };
    return buildExportPreview(activeData, activeSheet.outputHeaders, getHeaderLines(activeSheet.name), 5);
  }, [activeSheet, activeData, getHeaderLines]);

  const combinedSpecs = useMemo(() => {
    if (!allProcessed) return [];
    return state.sheets.map(sheet => {
      const isActive = sheet.name === activeSheet?.name;
      return {
        name: sheet.name,
        data: transformSheet(sheet, isActive),
        outputHeaders: sheet.outputHeaders,
        headerLines: getHeaderLines(sheet.name),
      };
    });
  }, [allProcessed, state.sheets, activeSheet, transformSheet, getHeaderLines]);

  const combinedTxtPreview = useMemo(() => {
    if (combinedSpecs.length === 0) return { rows: [], hiddenCount: 0 };
    return buildCombinedTxtPreview(combinedSpecs, 60);
  }, [combinedSpecs]);

  const totalRecords = useMemo(() => state.sheets.reduce((sum, s) => sum + s.stats.finalRows, 0), [state.sheets]);

  const qpsPreview = useMemo(() => {
    if (!activeSheet) return { lines: [], hiddenCount: 0 };
    const cfg = qpsConfigs[activeSheet.name]
      ?? initialQpsConfig(activeSheet.name, activeSheet.questionCode, activeSheet.questionLabel);
    return buildQpsPreview(activeSheet.originalData, activeSheet.columnConfig.columnAKey, activeSheet.columnConfig.columnBKey, cfg, 40);
  }, [activeSheet, qpsConfigs]);

  /* ═══════════════════════════════════════════════════════ */
  /* ── File handling                                     ── */
  /* ═══════════════════════════════════════════════════════ */

  const handleFile = useCallback(async (file: File) => {
    const fileValidation = validateFile(file);
    if (!fileValidation.isValid) {
      setState(prev => ({ ...prev, errors: fileValidation.errors.map(e => e.message) }));
      return;
    }

    try {
      const fileType = getFileType(file.name);
      const result = fileType === 'csv' ? await readCSV(file) : await readWorkbook(file);

      const sheets = result.sheetNames.map(name => buildSheetState(result.workbook, name));

      const newFormatInputs: FormatInputMap = {};
      const newOpenFormats: Record<string, boolean> = {};
      const newQpsConfigs: Record<string, QpsConfig> = {};
      const newOpenQps: Record<string, boolean> = {};
      sheets.forEach((s, i) => {
        newFormatInputs[s.name] = initialFormatInputs(
          s.name,
          s.questionCode,
          s.questionLabel,
          i + 1,
          getHighestCodeLength(s),
        );
        newOpenFormats[s.name] = i === 0;
        newQpsConfigs[s.name] = initialQpsConfig(s.name, s.questionCode, s.questionLabel);
        newOpenQps[s.name] = i === 0;
      });
      setFormatInputs(newFormatInputs);
      setOpenFormats(newOpenFormats);
      setQpsConfigs(newQpsConfigs);
      setOpenQps(newOpenQps);

      setState(prev => ({
        ...prev,
        workbook: result.workbook,
        sheets,
        activeSheet: sheets[0]?.name ?? '',
        fileName: file.name,
        fileType,
        searchTerm: '',
        sortConfig: { column: '', direction: null },
        errors: [],
        isProcessing: false,
      }));
      setCurrentPage(0);
      debugLog('App', `File loaded: ${file.name} — ${sheets.length} tab(s): ${sheets.map(s => s.name).join(', ')}`);
    } catch (error) {
      setState(prev => ({ ...prev, errors: [`Failed to read file: ${error}`] }));
    }
  }, []);

  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault(); setDragActive(false);
    const file = e.dataTransfer.files[0];
    if (file) handleFile(file);
  }, [handleFile]);

  const handleFileInput = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) handleFile(file);
  }, [handleFile]);

  const handleSheetChange = useCallback((sheetName: string) => {
    setState(prev => ({ ...prev, activeSheet: sheetName, searchTerm: '', sortConfig: { column: '', direction: null } }));
    setCurrentPage(0);
    setOpenFormats(prev => ({ ...prev, [sheetName]: true }));
  }, []);

  const updateColumnConfig = useCallback((field: keyof ProcessingConfig, value: string) => {
    if (!activeSheet) return;
    updateSheet(activeSheet.name, sheet => {
      const newConfig = { ...sheet.columnConfig, [field]: value };

      const excluded: string[] = [];
      if (newConfig.notesColumnKey) excluded.push(newConfig.notesColumnKey);
      for (const h of sheet.allHeaders) {
        if (h === newConfig.columnAKey || h === newConfig.columnBKey) continue;
        if (!excluded.includes(h)) excluded.push(h);
      }
      newConfig.excludedColumns = excluded;

      const outHeaders = sheet.allHeaders.filter(h => !excluded.includes(h));
      const initFilters: FilterConfig = {};
      outHeaders.forEach(h => { initFilters[h] = []; });

      return {
        ...sheet,
        columnConfig: newConfig,
        outputHeaders: outHeaders,
        filters: initFilters,
        processedData: [],
        processed: false,
        stats: { ...emptyStats, originalRows: sheet.originalData.length },
        processingLog: [],
      };
    });
  }, [activeSheet, updateSheet]);

  const handleProcess = useCallback(() => {
    if (state.sheets.length === 0) return;
    setState(prev => ({ ...prev, isProcessing: true }));

    setTimeout(() => {
      setState(prev => ({
        ...prev,
        isProcessing: false,
        sheets: prev.sheets.map(sheet => {
          const result = processData(sheet.originalData, sheet.allHeaders, sheet.columnConfig);
          const initFilters: FilterConfig = {};
          result.outputHeaders.forEach(h => { initFilters[h] = []; });

          debugLog('App', `Processed tab "${sheet.name}": ${result.stats.finalRows} rows (removed ${result.stats.removedRows}, inserted ${result.stats.insertedRows})`);

          return {
            ...sheet,
            processedData: result.processedData,
            outputHeaders: result.outputHeaders,
            stats: result.stats,
            processingLog: result.logs,
            filters: initFilters,
            processed: true,
          };
        }),
      }));
      setCurrentPage(0);
    }, 50);
  }, [state.sheets.length]);

  const handleReset = useCallback(() => {
    setState(prev => ({
      ...prev,
      sheets: prev.sheets.map(sheet => ({
        ...sheet,
        processedData: [],
        processed: false,
        stats: { ...emptyStats, originalRows: sheet.originalData.length },
        processingLog: [],
        filters: Object.fromEntries(sheet.outputHeaders.map(h => [h, []])),
      })),
      searchTerm: '',
      sortConfig: { column: '', direction: null },
    }));
    setCurrentPage(0);
  }, []);

  const handleExport = useCallback(() => {
    if (!allProcessed) return;
    const specs = state.sheets.map(sheet => {
      const isActive = sheet.name === activeSheet?.name;
      return {
        name: sheet.name,
        data: transformSheet(sheet, isActive),
        outputHeaders: sheet.outputHeaders,
        headerLines: getHeaderLines(sheet.name),
      };
    });
    exportWorkbook(specs, state.fileName);
  }, [allProcessed, state.sheets, state.fileName, activeSheet, transformSheet, getHeaderLines]);

  const handleExportTxt = useCallback(() => {
    if (!allProcessed) return;
    const specs = state.sheets.map(sheet => {
      const isActive = sheet.name === activeSheet?.name;
      return {
        name: sheet.name,
        data: transformSheet(sheet, isActive),
        outputHeaders: sheet.outputHeaders,
        headerLines: getHeaderLines(sheet.name),
      };
    });
    exportCombinedTxt(specs, state.fileName);
  }, [allProcessed, state.sheets, state.fileName, activeSheet, transformSheet, getHeaderLines]);

  const handleSort = useCallback((column: string) => {
    setState(prev => {
      let dir: 'asc' | 'desc' | null;
      if (prev.sortConfig.column !== column) dir = 'asc';
      else if (prev.sortConfig.direction === 'asc') dir = 'desc';
      else if (prev.sortConfig.direction === 'desc') dir = null;
      else dir = 'asc';
      return { ...prev, sortConfig: { column, direction: dir } };
    });
  }, []);

  const updateActiveFilters = useCallback((updater: (f: FilterConfig) => FilterConfig) => {
    if (!activeSheet) return;
    updateSheet(activeSheet.name, sheet => ({ ...sheet, filters: updater(sheet.filters) }));
    setCurrentPage(0);
  }, [activeSheet, updateSheet]);

  const handleNewFile = useCallback(() => {
    setState(initialState);
    setFormatInputs({});
    setOpenFormats({});
    setQpsConfigs({});
    setOpenQps({});
    setCurrentPage(0);
  }, []);

  const setQpsField = useCallback((sheetName: string, field: keyof QpsConfig, value: string) => {
    const sheet = state.sheets.find(item => item.name === sheetName);
    const normalizedValue = field === 'variableName'
      ? `V${value.trim().replace(/^V/i, '')}`
      : value;
    setQpsConfigs(prev => ({
      ...prev,
      [sheetName]: {
        ...(prev[sheetName] ?? initialQpsConfig(
          sheetName,
          sheet?.questionCode,
          sheet?.questionLabel,
        )),
        [field]: normalizedValue,
      },
    }));
  }, [state.sheets]);

  const buildQpsSpecs = useCallback(() => {
    return state.sheets.map(sheet => ({
      name: sheet.name,
      data: sheet.originalData,
      codeKey: sheet.columnConfig.columnAKey,
      labelKey: sheet.columnConfig.columnBKey,
      config: qpsConfigs[sheet.name]
        ?? initialQpsConfig(sheet.name, sheet.questionCode, sheet.questionLabel),
    }));
  }, [state.sheets, qpsConfigs]);

  const handleExportQpsXlsx = useCallback(() => {
    if (state.sheets.length === 0) return;
    exportQpsWorkbook(buildQpsSpecs(), state.fileName);
  }, [state.sheets.length, buildQpsSpecs, state.fileName]);

  const handleExportQpsTxt = useCallback(() => {
    if (state.sheets.length === 0) return;
    exportQpsCombinedTxt(buildQpsSpecs(), state.fileName);
  }, [state.sheets.length, buildQpsSpecs, state.fileName]);

  /* ═══════════════════════════════════════════════════════ */
  /* ── RENDER                                            ── */
  /* ═══════════════════════════════════════════════════════ */

  return (
    <div className="theme-page min-h-screen text-slate-800 dark:text-slate-100 flex flex-col transition-colors w-full max-w-full overflow-x-hidden">
      <HeaderNav
        currentPage={page}
        onNavigate={(p) => setPage(p)}
        showDebug={showDebug}
        onToggleDebug={() => setShowDebug(d => !d)}
      />

      {/* ─── TAB STRIP (Processor Only) ─────────────────── */}
      {page === 'processor' && state.sheets.length > 1 && (
        <div className="border-b border-slate-200/80 dark:border-slate-800 bg-white dark:bg-slate-900 shadow-sm">
          <div className="max-w-[1400px] mx-auto px-4 sm:px-6 lg:px-8 flex gap-1 overflow-x-auto py-1.5 scrollbar-thin">
            {state.sheets.map(sheet => {
              const isActive = sheet.name === activeSheet?.name;
              return (
                <button
                  type="button"
                  key={sheet.name}
                  onClick={() => handleSheetChange(sheet.name)}
                  className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-medium whitespace-nowrap transition-colors ${
                    isActive
                      ? 'btn-brand focus-ring-brand shadow-sm'
                      : 'text-slate-600 dark:text-slate-400 hover:bg-slate-100 dark:hover:bg-slate-800'
                  }`}
                >
                  {sheet.processed && <span className={`w-1.5 h-1.5 rounded-full ${isActive ? 'bg-emerald-300' : 'bg-emerald-500'}`} />}
                  {sheet.name}
                  <span className={`text-[10px] ${isActive ? 'text-indigo-200' : 'text-slate-400'}`}>
                    {sheet.processed ? sheet.stats.finalRows : sheet.originalData.length}
                  </span>
                </button>
              );
            })}
          </div>
        </div>
      )}

      {/* ─── OE Entries Page Content ─── */}
      <div className={page === 'processor' ? 'block w-full' : 'hidden'}>
        <main className="flex-1 max-w-[1400px] w-full mx-auto px-4 sm:px-6 lg:px-8 py-6">
          <TopExportBar title="OE Entries Exports" badge={allProcessed ? `${state.sheets.length} sheets` : undefined}>
            <button
              type="button"
              onClick={handleExport}
              disabled={!allProcessed}
              className={`flex-1 min-w-[155px] h-9 px-3 text-xs font-semibold rounded-xl shadow-sm transition-all flex items-center justify-center gap-1.5 ${
                allProcessed
                  ? 'btn-brand focus-ring-brand'
                  : 'bg-slate-200 dark:bg-slate-800 text-slate-400 dark:text-slate-600 cursor-not-allowed'
              }`}
            >
              <svg className="w-3.5 h-3.5 shrink-0" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}><path strokeLinecap="round" strokeLinejoin="round" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4" /></svg>
              Workbook (.xlsx)
            </button>

            <button
              type="button"
              onClick={handleExportTxt}
              disabled={!allProcessed}
              className={`flex-1 min-w-[155px] h-9 px-3 text-xs font-semibold rounded-xl shadow-sm transition-all border flex items-center justify-center gap-1.5 ${
                allProcessed
                  ? 'bg-white dark:bg-slate-800 text-slate-700 dark:text-slate-200 border-slate-300 dark:border-slate-700 hover:bg-slate-50 dark:hover:bg-slate-700'
                  : 'bg-slate-100 dark:bg-slate-800 text-slate-400 dark:text-slate-600 border-slate-200 dark:border-slate-800 cursor-not-allowed'
              }`}
            >
              <svg className="w-3.5 h-3.5 shrink-0" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}><path strokeLinecap="round" strokeLinejoin="round" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" /></svg>
              Combined (.txt)
            </button>

            <button
              type="button"
              onClick={handleExportQpsXlsx}
              disabled={state.sheets.length === 0}
              className={`flex-1 min-w-[155px] h-9 px-3 text-xs font-semibold rounded-xl shadow-sm transition-all flex items-center justify-center gap-1.5 ${
                state.sheets.length > 0
                  ? 'bg-teal-600 text-white hover:bg-teal-700 shadow-teal-500/20'
                  : 'bg-slate-200 dark:bg-slate-800 text-slate-400 dark:text-slate-600 cursor-not-allowed'
              }`}
            >
              QPS (.xlsx)
            </button>

            <button
              type="button"
              onClick={handleExportQpsTxt}
              disabled={state.sheets.length === 0}
              className={`flex-1 min-w-[155px] h-9 px-3 text-xs font-semibold rounded-xl shadow-sm transition-all border flex items-center justify-center gap-1.5 ${
                state.sheets.length > 0
                  ? 'bg-white dark:bg-slate-800 text-teal-700 dark:text-teal-400 border-teal-300 dark:border-teal-800 hover:bg-teal-50 dark:hover:bg-teal-950/40'
                  : 'bg-slate-100 dark:bg-slate-800 text-slate-400 dark:text-slate-600 border-slate-200 dark:border-slate-800 cursor-not-allowed'
              }`}
            >
              QPS Combined (.txt)
            </button>
          </TopExportBar>

          <div className="flex flex-col lg:flex-row gap-6">
            <div className="lg:w-80 shrink-0 space-y-5">
              {/* Import File */}
              <section className="bg-white dark:bg-slate-900 rounded-2xl border border-slate-200 dark:border-slate-800 shadow-sm overflow-hidden">
                <div className="px-4 py-3 border-b border-slate-100 dark:border-slate-800 flex items-center justify-between">
                  <h2 className="text-sm font-semibold text-slate-800 dark:text-slate-200">Import File</h2>
                  {state.fileName && (
                    <button onClick={handleNewFile} className="text-xs text-indigo-600 dark:text-indigo-400 hover:text-indigo-700 font-medium">Change file</button>
                  )}
                </div>
                <div className="p-4">
                  {!state.fileName ? (
                    <div
                      onDragOver={e => { e.preventDefault(); setDragActive(true); }}
                      onDragLeave={() => setDragActive(false)}
                      onDrop={handleDrop}
                      className={`relative border-2 border-dashed rounded-xl p-6 text-center transition-all cursor-pointer ${
                        dragActive
                          ? 'border-indigo-400 bg-indigo-50 dark:bg-indigo-950/40'
                          : 'border-slate-300 dark:border-slate-700 hover:border-slate-400 dark:hover:border-slate-600'
                      }`}
                    >
                      <input type="file" accept=".xlsx,.xls,.csv" onChange={handleFileInput} className="absolute inset-0 w-full h-full opacity-0 cursor-pointer" />
                      <UploadIcon />
                      <p className="mt-2 text-sm font-medium text-slate-600 dark:text-slate-300">Drop Excel file here</p>
                      <p className="text-xs text-slate-400 dark:text-slate-500 mt-0.5">or click to browse · .xlsx .xls .csv</p>
                      <p className="text-[10px] text-slate-400 dark:text-slate-500 mt-1">All workbook tabs are loaded automatically</p>
                    </div>
                  ) : (
                    <div className="space-y-4">
                      <div className="flex items-center gap-3 p-3 bg-emerald-50 dark:bg-emerald-950/40 rounded-lg border border-emerald-100 dark:border-emerald-900/60">
                        <div className="w-10 h-10 bg-emerald-100 dark:bg-emerald-900/60 rounded-lg flex items-center justify-center shrink-0">
                          <svg className="w-5 h-5 text-emerald-600 dark:text-emerald-400" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2}><path strokeLinecap="round" strokeLinejoin="round" d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" /></svg>
                        </div>
                        <div className="flex-1 min-w-0">
                          <p className="text-sm font-medium text-slate-900 dark:text-white truncate">{state.fileName}</p>
                          <p className="text-xs text-slate-500 dark:text-slate-400">
                            {state.sheets.length} tab{state.sheets.length === 1 ? '' : 's'} · {state.sheets.reduce((n, s) => n + s.originalData.length, 0).toLocaleString()} rows total
                          </p>
                        </div>
                      </div>

                      {activeSheet && (
                        <div className="rounded-lg border border-brand-200 bg-brand-50 px-3 py-2">
                          <p className="text-[10px] font-bold uppercase tracking-wide text-brand">Auto-derived from Excel row 1</p>
                          <p className="mt-1 text-xs font-mono text-slate-700 dark:text-slate-200">
                            A1: {activeSheet.questionCode || '(blank)'}
                          </p>
                          <p className="text-xs font-mono text-slate-700 dark:text-slate-200">
                            B1: {activeSheet.questionLabel || '(blank)'}
                          </p>
                          <p className="mt-1 text-[10px] text-slate-500 dark:text-slate-400">Code/label records begin at Excel row 3.</p>
                        </div>
                      )}

                      {/* Column Config */}
                      {activeSheet && (
                        <div>
                          <button
                            onClick={() => setShowColumnConfig(s => !s)}
                            className="w-full text-xs flex items-center justify-between px-3 py-2 bg-slate-50 dark:bg-slate-800/60 rounded-lg hover:bg-slate-100 dark:hover:bg-slate-800 transition-colors text-slate-600 dark:text-slate-400 font-semibold"
                          >
                            <span>Column Configuration <span className="text-indigo-500 font-bold">({activeSheet.name})</span></span>
                            <svg className={`w-4 h-4 transition-transform ${showColumnConfig ? 'rotate-180' : ''}`} fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}><path strokeLinecap="round" strokeLinejoin="round" d="M19 9l-7 7-7-7" /></svg>
                          </button>

                          {showColumnConfig && (
                            <div className="mt-2 space-y-3 p-3 bg-slate-50 dark:bg-slate-800/60 rounded-lg border border-slate-200 dark:border-slate-700">
                              <div>
                                <label className="block text-xs font-medium text-slate-500 dark:text-slate-400 mb-1">Column A (Number)</label>
                                <select value={activeSheet.columnConfig.columnAKey} onChange={e => updateColumnConfig('columnAKey', e.target.value)} className="w-full text-xs px-2 py-1.5 border border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-800 rounded-md text-slate-800 dark:text-slate-200">
                                  {activeSheet.allHeaders.map(h => <option key={h} value={h}>{h}</option>)}
                                </select>
                              </div>
                              <div>
                                <label className="block text-xs font-medium text-slate-500 dark:text-slate-400 mb-1">Column B (Description)</label>
                                <select value={activeSheet.columnConfig.columnBKey} onChange={e => updateColumnConfig('columnBKey', e.target.value)} className="w-full text-xs px-2 py-1.5 border border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-800 rounded-md text-slate-800 dark:text-slate-200">
                                  {activeSheet.allHeaders.map(h => <option key={h} value={h}>{h}</option>)}
                                </select>
                              </div>
                              <div>
                                <label className="block text-xs font-medium text-slate-500 dark:text-slate-400 mb-1">Notes Column (excluded)</label>
                                <select value={activeSheet.columnConfig.notesColumnKey} onChange={e => updateColumnConfig('notesColumnKey', e.target.value)} className="w-full text-xs px-2 py-1.5 border border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-800 rounded-md text-slate-800 dark:text-slate-200">
                                  <option value="">— None —</option>
                                  {activeSheet.allHeaders.map(h => <option key={h} value={h}>{h}</option>)}
                                </select>
                              </div>
                            </div>
                          )}
                        </div>
                      )}

                      {/* Process button */}
                      <button
                        onClick={handleProcess}
                        disabled={state.sheets.length === 0 || state.isProcessing}
                        className={`w-full py-3 rounded-xl font-bold transition-all flex items-center justify-center gap-2 text-sm shadow-sm ${
                          state.isProcessing
                            ? 'bg-slate-200 dark:bg-slate-800 text-slate-500 dark:text-slate-400 cursor-wait'
                            : allProcessed
                            ? 'bg-emerald-600 text-white hover:bg-emerald-700 shadow-emerald-500/20'
                            : 'btn-brand focus-ring-brand shadow-sm'
                        }`}
                      >
                        {state.isProcessing
                          ? <><SpinnerIcon /> Processing...</>
                          : allProcessed
                          ? '↻ Re-Process All Tabs'
                          : '▶ Process All Tabs'}
                      </button>
                    </div>
                  )}

                  {state.errors.length > 0 && (
                    <div className="mt-3 p-3 bg-red-50 dark:bg-red-950/40 border border-red-200 dark:border-red-900/60 rounded-lg">
                      {state.errors.map((err, i) => <p key={i} className="text-xs text-red-600 dark:text-red-400">{err}</p>)}
                    </div>
                  )}
                </div>
              </section>

              {/* Per-tab Format Builders */}
              {state.sheets.length > 0 && (
                <section className="bg-gradient-to-br from-indigo-50/70 to-violet-50/70 dark:from-indigo-950/30 dark:to-violet-950/30 rounded-2xl border border-indigo-100 dark:border-indigo-900/50 shadow-sm overflow-hidden">
                  <div className="px-4 py-3 border-b border-indigo-100 dark:border-indigo-900/50">
                    <h2 className="text-sm font-bold text-indigo-900 dark:text-indigo-300">Format Builders</h2>
                    <p className="text-[10px] text-indigo-500 dark:text-indigo-400">One builder per worksheet tab — rows 3 &amp; 5 of the exported sheet</p>
                  </div>
                  <div className="p-3 space-y-2">
                    {state.sheets.map(sheet => {
                      const isOpen = !!openFormats[sheet.name];
                      const lines = getHeaderLines(sheet.name);
                      return (
                        <div key={sheet.name} className="bg-white dark:bg-slate-900 rounded-xl border border-indigo-100 dark:border-indigo-900/40 shadow-sm overflow-hidden">
                          <button
                            onClick={() => setOpenFormats(prev => ({ ...prev, [sheet.name]: !isOpen }))}
                            className="w-full px-3 py-2 flex items-center justify-between text-left hover:bg-indigo-50/50 dark:hover:bg-indigo-950/30 transition-colors"
                          >
                            <div className="flex items-center gap-2 min-w-0">
                              <span className={`w-5 h-5 rounded flex items-center justify-center text-[9px] font-bold text-white ${sheet.name === activeSheet?.name ? 'bg-indigo-500' : 'bg-slate-400'}`}>
                                {sheet.name.slice(0, 2).toUpperCase()}
                              </span>
                              <span className="text-xs font-semibold text-slate-800 dark:text-slate-200 truncate">{sheet.name}</span>
                            </div>
                            <svg className={`w-3.5 h-3.5 text-slate-400 transition-transform ${isOpen ? 'rotate-180' : ''}`} fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                              <path strokeLinecap="round" strokeLinejoin="round" d="M19 9l-7 7-7-7" />
                            </svg>
                          </button>

                          {isOpen && (
                            <div className="px-3 pb-3 space-y-2 border-t border-indigo-50 dark:border-indigo-900/30 pt-2">
                              {FORMAT_TEMPLATES.map(template => (
                                <div key={template.id} className="space-y-1">
                                  <p className="text-[9px] font-bold text-slate-500 dark:text-slate-400 uppercase tracking-wide">
                                    {template.id === 'q-pattern' ? 'Row 3 (Q-Pattern)' : 'Row 5 (L-Pattern)'}
                                  </p>
                                  <div className="text-[9px] font-mono bg-slate-900 dark:bg-black/60 text-emerald-400 dark:text-emerald-300 px-1.5 py-1 rounded-md overflow-x-auto whitespace-nowrap">
                                    {template.templateString}
                                  </div>
                                  <div className={`grid gap-1.5 ${template.inputCount === 3 ? 'grid-cols-3' : 'grid-cols-2'}`}>
                                    {Array.from({ length: template.inputCount }).map((_, idx) => (
                                      <input
                                        key={idx}
                                        type="text"
                                        value={formatInputs[sheet.name]?.[template.id]?.[idx] ?? ''}
                                        onChange={e => setFormatInput(sheet.name, template.id, idx, e.target.value)}
                                        placeholder={template.exampleInputs[idx] || ''}
                                        className="w-full text-[11px] px-1.5 py-1 border border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-800 text-slate-800 dark:text-slate-200 rounded-md focus:ring-1 focus:ring-indigo-500 focus:border-indigo-500 outline-none"
                                      />
                                    ))}
                                  </div>
                                  <p className="text-[10px] font-mono font-semibold text-indigo-700 dark:text-indigo-300 truncate">
                                    {(template.id === 'q-pattern' ? lines.qLine : lines.lLine) || <span className="text-slate-300 font-normal">(empty)</span>}
                                  </p>
                                </div>
                              ))}
                            </div>
                          )}
                        </div>
                      );
                    })}
                  </div>
                </section>
              )}

              {/* QPS Script Builder */}
              {state.sheets.length > 0 && (
                <section className="bg-gradient-to-br from-teal-50/70 to-cyan-50/70 dark:from-teal-950/30 dark:to-cyan-950/30 rounded-2xl border border-teal-100 dark:border-teal-900/50 shadow-sm overflow-hidden">
                  <div className="px-4 py-3 border-b border-teal-100 dark:border-teal-900/50">
                    <h2 className="text-sm font-bold text-teal-900 dark:text-teal-300">QPS Script Builder</h2>
                    <p className="text-[10px] text-teal-600 dark:text-teal-400">Netting output (V B / V […] / T M + R/Y + O R/O E). Uses original codes.</p>
                  </div>
                  <div className="p-3 space-y-2">
                    {state.sheets.map(sheet => {
                      const isOpen = !!openQps[sheet.name];
                      const cfg = qpsConfigs[sheet.name]
                        ?? initialQpsConfig(sheet.name, sheet.questionCode, sheet.questionLabel);
                      return (
                        <div key={sheet.name} className="bg-white dark:bg-slate-900 rounded-xl border border-teal-100 dark:border-teal-900/40 shadow-sm overflow-hidden">
                          <button
                            onClick={() => setOpenQps(prev => ({ ...prev, [sheet.name]: !isOpen }))}
                            className="w-full px-3 py-2 flex items-center justify-between text-left hover:bg-teal-50/50 dark:hover:bg-teal-950/30 transition-colors"
                          >
                            <div className="flex items-center gap-2 min-w-0">
                              <span className={`w-5 h-5 rounded flex items-center justify-center text-[9px] font-bold text-white ${sheet.name === activeSheet?.name ? 'bg-teal-500' : 'bg-slate-400'}`}>
                                {sheet.name.slice(0, 2).toUpperCase()}
                              </span>
                              <span className="text-xs font-semibold text-slate-800 dark:text-slate-200 truncate">{sheet.name}</span>
                            </div>
                            <span className="text-[9px] font-mono text-slate-400 truncate max-w-24">V [{cfg.variableName}]</span>
                          </button>

                          {isOpen && (
                            <div className="px-3 pb-3 pt-2 space-y-2 border-t border-teal-50 dark:border-teal-900/30">
                              <div>
                                <label className="block text-[9px] text-slate-500 dark:text-slate-400 font-medium">Variable Name (C3)</label>
                                <input type="text" value={cfg.variableName} onChange={e => setQpsField(sheet.name, 'variableName', e.target.value)} placeholder="VP2"
                                  className="w-full text-[11px] px-1.5 py-1 border border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-800 text-slate-800 dark:text-slate-200 rounded focus:ring-1 focus:ring-teal-500 focus:border-teal-500 outline-none" />
                              </div>
                              <div>
                                <label className="block text-[9px] text-slate-500 dark:text-slate-400 font-medium">Questions (C4) — comma separated</label>
                                <input type="text" value={cfg.questions} onChange={e => setQpsField(sheet.name, 'questions', e.target.value)} placeholder="P2 or B1,B2"
                                  className="w-full text-[11px] px-1.5 py-1 border border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-800 text-slate-800 dark:text-slate-200 rounded focus:ring-1 focus:ring-teal-500 focus:border-teal-500 outline-none" />
                              </div>
                              <div>
                                <label className="block text-[9px] text-slate-500 dark:text-slate-400 font-medium">Question Label (C6)</label>
                                <input type="text" value={cfg.questionLabel} onChange={e => setQpsField(sheet.name, 'questionLabel', e.target.value)} placeholder="P2 REASONS FOR..."
                                  className="w-full text-[11px] px-1.5 py-1 border border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-800 text-slate-800 dark:text-slate-200 rounded focus:ring-1 focus:ring-teal-500 focus:border-teal-500 outline-none" />
                              </div>
                              <div>
                                <label className="block text-[9px] text-slate-500 dark:text-slate-400 font-medium">Table Type (→ "T _")</label>
                                <input type="text" value={cfg.tableType} onChange={e => setQpsField(sheet.name, 'tableType', e.target.value)} placeholder="M"
                                  className="w-20 text-[11px] px-1.5 py-1 border border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-800 text-slate-800 dark:text-slate-200 rounded focus:ring-1 focus:ring-teal-500 focus:border-teal-500 outline-none" />
                              </div>
                            </div>
                          )}
                        </div>
                      );
                    })}

                    <div className="grid grid-cols-2 gap-2 pt-1">
                      <button onClick={handleExportQpsXlsx} className="py-2 rounded-xl text-xs font-bold bg-teal-600 text-white hover:bg-teal-700 transition-colors flex items-center justify-center gap-1 shadow-teal-500/20">
                        QPS .XLSX
                      </button>
                      <button onClick={handleExportQpsTxt} className="py-2 rounded-xl text-xs font-bold bg-white dark:bg-slate-800 text-teal-700 dark:text-teal-300 border border-teal-300 dark:border-teal-800 hover:bg-teal-50 dark:hover:bg-teal-950/40 transition-colors flex items-center justify-center gap-1">
                        QPS .TXT
                      </button>
                    </div>
                  </div>
                </section>
              )}

              {/* Statistics (active tab) */}
              {activeSheet && activeSheet.stats.finalRows > 0 && (
                <section className="bg-white dark:bg-slate-900 rounded-2xl border border-slate-200 dark:border-slate-800 shadow-sm overflow-hidden">
                  <div className="px-4 py-3 border-b border-slate-100 dark:border-slate-800 flex items-center justify-between">
                    <h2 className="text-sm font-semibold text-slate-800 dark:text-slate-200">Processing Results</h2>
                    <span className="text-[10px] text-indigo-500 dark:text-indigo-400 font-semibold">{activeSheet.name}</span>
                  </div>
                  <div className="p-4 grid grid-cols-4 gap-2">
                    <div className="text-center p-2 bg-slate-50 dark:bg-slate-800/60 rounded-lg">
                      <p className="text-lg font-bold text-slate-700 dark:text-slate-200">{activeSheet.stats.originalRows.toLocaleString()}</p>
                      <p className="text-[9px] text-slate-500 mt-0.5 dark:text-slate-400">Original</p>
                    </div>
                    <div className="text-center p-2 bg-amber-50 dark:bg-amber-950/40 rounded-lg">
                      <p className="text-lg font-bold text-amber-600 dark:text-amber-400">{activeSheet.stats.removedRows.toLocaleString()}</p>
                      <p className="text-[9px] text-amber-600 dark:text-amber-400 mt-0.5">Removed</p>
                    </div>
                    <div className="text-center p-2 bg-sky-50 dark:bg-sky-950/40 rounded-lg">
                      <p className="text-lg font-bold text-sky-600 dark:text-sky-400">{activeSheet.stats.insertedRows.toLocaleString()}</p>
                      <p className="text-[9px] text-sky-600 dark:text-sky-400 mt-0.5">Auto-added</p>
                    </div>
                    <div className="text-center p-2 bg-emerald-50 dark:bg-emerald-950/40 rounded-lg">
                      <p className="text-lg font-bold text-emerald-600 dark:text-emerald-400">{activeSheet.stats.finalRows.toLocaleString()}</p>
                      <p className="text-[9px] text-emerald-600 dark:text-emerald-400 mt-0.5">Final</p>
                    </div>
                  </div>
                </section>
              )}

              {/* Export / Reset Panel */}
              {state.sheets.some(s => s.processed) && (
                <section className="bg-white dark:bg-slate-900 rounded-2xl border border-slate-200 dark:border-slate-800 shadow-sm overflow-hidden">
                  <div className="px-4 py-3 border-b border-slate-100 dark:border-slate-800">
                    <h2 className="text-sm font-semibold text-slate-800 dark:text-slate-200">Quick Controls</h2>
                  </div>
                  <div className="p-4 space-y-3">
                    <button onClick={handleReset} className="w-full py-2 text-slate-500 dark:text-slate-400 rounded-xl text-xs font-semibold hover:bg-slate-100 dark:hover:bg-slate-800 transition-colors">
                      ↩ Reset Results
                    </button>
                    <div className="bg-amber-50 dark:bg-amber-950/40 border border-amber-200 dark:border-amber-900/60 rounded-lg p-2.5">
                      <p className="text-[10px] text-amber-800 dark:text-amber-400 font-semibold">TXT Merge Rules</p>
                      <ul className="text-[10px] text-amber-700 dark:text-amber-400 mt-1 space-y-0.5 list-disc pl-3">
                        <li>Title only on 1st tab</li>
                        <li>Intermediate Z removed</li>
                        <li>Only last tab ends with Z</li>
                      </ul>
                    </div>
                  </div>
                </section>
              )}

              {/* Processing Log */}
              {state.sheets.some(s => s.processingLog.length > 0) && (
                <section className="bg-white dark:bg-slate-900 rounded-2xl border border-slate-200 dark:border-slate-800 shadow-sm overflow-hidden">
                  <button onClick={() => setShowLog(s => !s)} className="w-full px-4 py-3 border-b border-slate-100 dark:border-slate-800 flex items-center justify-between text-left">
                    <h2 className="text-sm font-semibold text-slate-800 dark:text-slate-200">Processing Log</h2>
                    <svg className={`w-4 h-4 text-slate-400 transition-transform ${showLog ? 'rotate-180' : ''}`} fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}><path strokeLinecap="round" strokeLinejoin="round" d="M19 9l-7 7-7-7" /></svg>
                  </button>
                  {showLog && (
                    <div className="p-3 max-h-56 overflow-y-auto space-y-3">
                      {state.sheets.filter(s => s.processingLog.length > 0).map(sheet => (
                        <div key={sheet.name}>
                          <p className="text-[10px] font-bold text-slate-700 dark:text-slate-300 uppercase tracking-wide mb-0.5">{sheet.name}</p>
                          <div className="space-y-0.5 font-mono text-[10px] leading-relaxed">
                            {sheet.processingLog.map((entry, i) => (
                              <div key={i} className={`${entry.level === 'error' ? 'text-red-500' : entry.level === 'warning' ? 'text-amber-600' : entry.level === 'success' ? 'text-emerald-600' : 'text-slate-500'}`}>
                                {entry.message}
                              </div>
                            ))}
                          </div>
                        </div>
                      ))}
                    </div>
                  )}
                </section>
              )}
            </div>

            {/* Main Content */}
            <div className="flex-1 min-w-0 space-y-4">
              {/* Search & filters */}
              {activeSheet && activeSheet.processed && (
                <div className="bg-white dark:bg-slate-900 rounded-2xl border border-slate-200 dark:border-slate-800 shadow-sm overflow-hidden">
                  <div className="p-3 flex flex-col sm:flex-row gap-3 items-stretch sm:items-center">
                    <div className="flex-1 relative">
                      <input
                        type="text"
                        placeholder={`Search ${activeSheet.name}…`}
                        value={state.searchTerm}
                        onChange={e => { setState(prev => ({ ...prev, searchTerm: e.target.value })); setCurrentPage(0); }}
                        className="w-full pl-9 pr-8 py-2 text-sm border border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-800 text-slate-800 dark:text-slate-200 rounded-xl focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none"
                      />
                      <svg className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400 dark:text-slate-500" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}><path strokeLinecap="round" strokeLinejoin="round" d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z" /></svg>
                    </div>
                    <div className="flex flex-wrap gap-2">
                      {activeSheet.outputHeaders.map(header => (
                        <select
                          key={header}
                          value={activeSheet.filters[header]?.[0] || ''}
                          onChange={e => updateActiveFilters(f => ({ ...f, [header]: e.target.value ? [e.target.value] : [] }))}
                          className="text-xs border border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-800 text-slate-800 dark:text-slate-200 rounded-xl px-2 py-1.5 hover:border-slate-300 focus:ring-1 focus:ring-indigo-500 focus:border-indigo-500 max-w-40 truncate"
                        >
                          <option value="">All {header}</option>
                          {(uniqueValues[header] || []).slice(0, 50).map(val => <option key={val} value={val}>{val}</option>)}
                        </select>
                      ))}
                    </div>
                  </div>
                </div>
              )}

              {/* Export Preview */}
              {(exportPreview.rows.length > 0 || combinedTxtPreview.rows.length > 0) && activeSheet && (
                <div className="bg-white dark:bg-slate-900 rounded-2xl border border-slate-200 dark:border-slate-800 shadow-sm overflow-hidden">
                  <div className="px-4 py-2.5 border-b border-slate-100 dark:border-slate-800 flex items-center justify-between gap-3">
                    <div className="flex items-center gap-2 min-w-0">
                      <h2 className="text-sm font-semibold text-slate-800 dark:text-slate-200 shrink-0">Export Preview</h2>
                      <span className="text-[10px] text-slate-400 dark:text-slate-500 truncate">
                        {previewMode === 'txt' ? `all tabs merged` : `single sheet columns`}
                      </span>
                    </div>
                    <div className="flex items-center gap-2 shrink-0 bg-slate-100 dark:bg-slate-800/90 p-0.5 rounded-lg">
                      <button type="button" onClick={() => setPreviewMode('xlsx')} className={`px-2.5 py-1 text-[11px] font-bold rounded-md transition-all ${previewMode === 'xlsx' ? 'bg-white dark:bg-slate-950 text-slate-900 dark:text-white shadow-sm' : 'text-slate-500'}`}>
                        .xlsx sheet
                      </button>
                      <button type="button" onClick={() => setPreviewMode('txt')} disabled={combinedTxtPreview.rows.length === 0} className={`px-2.5 py-1 text-[11px] font-bold rounded-md transition-all ${previewMode === 'txt' ? 'bg-white dark:bg-slate-950 text-slate-900 dark:text-white shadow-sm' : 'text-slate-500'}`}>
                        .txt merged
                      </button>
                    </div>
                  </div>
                  {showExportPreview && (
                    <div className={`overflow-auto max-h-[24rem] ${previewMode === 'txt' ? 'bg-slate-950' : ''}`}>
                      <table className="w-full text-xs font-mono border-collapse">
                        <tbody>
                          {(previewMode === 'txt' ? combinedTxtPreview.rows : exportPreview.rows).map((r, idx) => {
                            const isTitle = r.kind === 'title';
                            return (
                              <Fragment key={idx}>
                                {previewMode === 'txt' && r.kind === 'begin' && idx > 0 && (
                                  <tr><td className="w-10" /><td className="py-1 px-2 text-[9px] text-slate-500 italic">— next tab —</td></tr>
                                )}
                                <tr className={isTitle ? 'text-violet-700 dark:text-violet-400 font-bold' : 'text-slate-600 dark:text-slate-400'}>
                                  <td className="w-10 px-2 py-1 text-right text-[10px] text-slate-400 dark:text-slate-600 select-none">{(r as any).rowNumber ?? idx + 1}</td>
                                  <td className="px-2 py-1 whitespace-nowrap">{(r as any).text || (r as any).cells?.join(' ') || ''}</td>
                                </tr>
                              </Fragment>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>
                  )}
                </div>
              )}

              {/* QPS Preview */}
              {activeSheet && qpsPreview.lines.length > 0 && (
                <div className="bg-white dark:bg-slate-900 rounded-2xl border border-teal-200 dark:border-teal-900 shadow-sm overflow-hidden">
                  <button onClick={() => setShowQpsPreview(s => !s)} className="w-full px-4 py-2.5 border-b border-teal-100 dark:border-teal-900 flex items-center justify-between text-left">
                    <h2 className="text-sm font-semibold text-teal-900 dark:text-teal-300">QPS Script Preview <span className="text-[10px] text-teal-500 font-medium">({activeSheet.name})</span></h2>
                    <svg className={`w-4 h-4 text-slate-400 transition-transform ${showQpsPreview ? 'rotate-180' : ''}`} fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}><path strokeLinecap="round" strokeLinejoin="round" d="M19 9l-7 7-7-7" /></svg>
                  </button>
                  {showQpsPreview && (
                    <div className="max-h-64 overflow-y-auto p-3 bg-slate-900 dark:bg-black text-[11px] font-mono text-cyan-300 space-y-0.5">
                      {qpsPreview.lines.map((line, i) => <div key={i}>{line.text}</div>)}
                      {qpsPreview.hiddenCount > 0 && <div className="text-slate-500 italic pt-1">… {qpsPreview.hiddenCount} more lines</div>}
                    </div>
                  )}
                </div>
              )}

              {/* Data Table */}
              <div className="bg-white dark:bg-slate-900 rounded-2xl border border-slate-200 dark:border-slate-800 shadow-sm overflow-hidden">
                {pagedData.length > 0 && activeSheet ? (
                  <>
                    <div className="overflow-x-auto">
                      <table className="w-full text-sm">
                        <thead className="bg-slate-50 dark:bg-slate-800/80 border-b border-slate-200 dark:border-slate-700 sticky top-0 z-10">
                          <tr>
                            {activeSheet.outputHeaders.map(header => (
                              <th key={header} onClick={() => handleSort(header)} className="px-4 py-3 text-left text-xs font-bold text-slate-600 dark:text-slate-300 cursor-pointer hover:bg-slate-100 transition-colors select-none">
                                <div className="flex items-center gap-1">{header} <ChevronIcon direction={state.sortConfig.column === header ? state.sortConfig.direction || 'none' : 'none'} /></div>
                              </th>
                            ))}
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-100 dark:divide-slate-800">
                          {pagedData.map((row, i) => (
                            <tr key={safePage * PAGE_SIZE + i} className="hover:bg-slate-50/50 dark:hover:bg-slate-800/40 transition-colors">
                              {activeSheet.outputHeaders.map(header => {
                                const val = row[header];
                                return (
                                  <td key={header} className={`px-4 py-2 text-slate-700 dark:text-slate-300 ${typeof val === 'number' ? 'font-mono text-center' : ''}`}>
                                    {val !== null && val !== undefined ? String(val) : <span className="text-slate-300 dark:text-slate-600">—</span>}
                                  </td>
                                );
                              })}
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                    {totalPages > 1 && (
                      <div className="px-4 py-3 border-t border-slate-100 dark:border-slate-800 flex items-center justify-between bg-slate-50 dark:bg-slate-800/40">
                        <span className="text-xs text-slate-500">Page {safePage + 1} of {totalPages}</span>
                        <div className="flex gap-1">
                          <button disabled={safePage === 0} onClick={() => setCurrentPage(0)} className="px-2 py-1 text-xs rounded-lg border border-slate-300 disabled:opacity-40 hover:bg-white dark:border-slate-700 dark:hover:bg-slate-800">⟨⟨</button>
                          <button disabled={safePage === 0} onClick={() => setCurrentPage(p => Math.max(0, p - 1))} className="px-2 py-1 text-xs rounded-lg border border-slate-300 disabled:opacity-40 hover:bg-white dark:border-slate-700 dark:hover:bg-slate-800">⟨ Prev</button>
                          <button disabled={safePage >= totalPages - 1} onClick={() => setCurrentPage(p => p + 1)} className="px-2 py-1 text-xs rounded-lg border border-slate-300 disabled:opacity-40 hover:bg-white dark:border-slate-700 dark:hover:bg-slate-800">Next ⟩</button>
                        </div>
                      </div>
                    )}
                  </>
                ) : (
                  <div className="p-16 text-center">
                    <UploadIcon />
                    <h3 className="text-base font-semibold text-slate-800 dark:text-slate-200 mt-2">No data to display</h3>
                    <p className="text-sm text-slate-500 mt-1">Upload an Excel file &amp; click Process All Tabs to clean, sort &amp; gap-fill entries.</p>
                  </div>
                )}
              </div>
            </div>
          </div>
        </main>
      </div>

      {/* ─── Stack Data Page (Mounted) ─── */}
      <div className={page === 'templateAppend' ? 'block w-full' : 'hidden'}>
        <TemplateAppendPage
          onNavigate={(p) => setPage(p)}
          showDebug={showDebug}
          onToggleDebug={() => setShowDebug(d => !d)}
          hideHeader={true}
        />
      </div>

      {/* ─── Summary Topline Page (Mounted) ─── */}
      <div className={page === 'summaryFiller' ? 'block w-full' : 'hidden'}>
        <SummaryTableFillerPage
          onNavigate={(p) => setPage(p)}
          showDebug={showDebug}
          onToggleDebug={() => setShowDebug(d => !d)}
          hideHeader={true}
        />
      </div>

      {/* ─── FOOTER ──────────────────────────────────────── */}
      <footer className="border-t border-slate-200 dark:border-slate-800 bg-white dark:bg-slate-900 mt-auto transition-colors z-20">
        <div className="max-w-[1400px] mx-auto px-4 sm:px-6 lg:px-8 py-4">
          <p className="text-[11px] text-slate-400 dark:text-slate-500 text-center font-medium tracking-wide">
            EZ Web Toolkit &middot; A professional-grade web application that replaces Excel/VBA workflows for data processing
          </p>
        </div>
      </footer>

      {/* Debug panel */}
      {showDebug && (
        <div className="fixed bottom-4 right-4 bg-slate-900 dark:bg-black border border-slate-700 p-3 rounded-xl shadow-2xl max-w-sm w-[350px] z-50">
          <div className="flex items-center justify-between border-b border-slate-700 pb-1 mb-2">
            <span className="text-xs font-bold text-white font-mono">Debug Panel</span>
            <button onClick={() => setShowDebug(false)} className="text-slate-400 hover:text-white text-xs">✕</button>
          </div>
          <div className="text-[10px] font-mono text-slate-400 max-h-60 overflow-auto space-y-1">
            <p>Page: {page}</p>
            <p>Records: {totalRecords}</p>
            <div className="border-t border-slate-700 pt-1 mt-1">
              {getDebugLogs().slice(-15).map((log, i) => (
                <div key={i} className="truncate">[{log.category}] {log.message}</div>
              ))}
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
