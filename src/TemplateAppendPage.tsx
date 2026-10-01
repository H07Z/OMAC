import { useMemo, useState, useCallback } from 'react';
import * as XLSX from 'xlsx';
import { readCSV, readWorkbook, getFileType } from './modules/excelLoader';
import {
  AppendResult,
  TemplateMappingRow,
  appendByTemplate,
  exportAppendResult,
  exportAppendResultCsv,
  parseTemplateMatrix,
  parseTemplateText,
  worksheetToMatrix,
  CellValue,
} from './modules/templateAppend';
import HeaderNav, { PageId } from './components/HeaderNav';
import TopExportBar from './components/TopExportBar';
import { MappingNote, ParityBadge, ParitySummary } from './components/StackDataIndicators';

interface TemplateAppendPageProps {
  onNavigate: (page: PageId) => void;
  showDebug?: boolean;
  onToggleDebug?: () => void;
  hideHeader?: boolean;
}

const SAMPLE_TEMPLATE = [
  ['Serial', 'SERIAL', 'SERIAL'],
  ['S18_BANNER', 'S18_BANNER', 'S18_BANNER'],
  ['PRODUCT', 'I_1_PRODUCT_TRIED', 'I_2_PRODUCT_TRIED'],
  ['Q1', 'I_1_Q1', 'I_2_Q1'],
  ['Q2', 'I_1_Q2', 'I_2_Q2'],
  ['Q3', 'I_1_Q3', 'I_2_Q3'],
  ['Q6', 'I_1_Q6', 'I_2_Q6'],
  ['Q7', 'I_1_Q7', 'I_2_Q7'],
  ['Q8', 'I_1_Q8', 'I_2_Q8'],
  ['Q9', 'I_1_Q9', 'I_2_Q9'],
  ['Q13', 'I_1_Q13', 'I_2_Q13'],
].map(row => row.join('\t')).join('\n');

function UploadBox({
  title,
  subtitle,
  onFile,
  fileName,
}: {
  title: string;
  subtitle: string;
  onFile: (file: File) => void;
  fileName?: string;
}) {
  const [drag, setDrag] = useState(false);
  return (
    <div
      onDragOver={e => { e.preventDefault(); setDrag(true); }}
      onDragLeave={() => setDrag(false)}
      onDrop={e => {
        e.preventDefault();
        setDrag(false);
        const file = e.dataTransfer.files[0];
        if (file) onFile(file);
      }}
      className={`relative rounded-2xl border-2 border-dashed p-5 text-center transition-all ${
        drag ? 'border-teal-400 bg-teal-50/40 dark:bg-teal-950/40' : 'border-slate-300 dark:border-slate-700 hover:border-slate-400 dark:hover:border-slate-600'
      }`}
    >
      <input
        type="file"
        accept=".xlsx,.xls,.csv"
        onChange={e => {
          const file = e.target.files?.[0];
          if (file) onFile(file);
        }}
        className="absolute inset-0 h-full w-full cursor-pointer opacity-0"
      />
      <div className="mx-auto mb-2 flex h-10 w-10 items-center justify-center rounded-xl bg-slate-100 dark:bg-slate-800 text-slate-400 dark:text-slate-500">
        <svg className="h-5 w-5" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
          <path strokeLinecap="round" strokeLinejoin="round" d="M12 16V4m0 0l-4 4m4-4l4 4M4 20h16" />
        </svg>
      </div>
      <p className="text-sm font-semibold text-slate-800 dark:text-slate-200">{fileName || title}</p>
      <p className="mt-0.5 text-xs text-slate-400 dark:text-slate-500">{fileName ? subtitle : `${subtitle} · drop or browse`}</p>
    </div>
  );
}

export default function TemplateAppendPage({ onNavigate, showDebug, onToggleDebug, hideHeader }: TemplateAppendPageProps) {
  const [sourceWorkbook, setSourceWorkbook] = useState<XLSX.WorkBook | null>(null);
  const [sourceFileName, setSourceFileName] = useState('');
  const [sourceSheetName, setSourceSheetName] = useState('');
  const [templateText, setTemplateText] = useState('');
  const [templateRows, setTemplateRows] = useState<TemplateMappingRow[]>([]);
  const [parityMappingHeader, setParityMappingHeader] = useState('');
  const [roundInput, setRoundInput] = useState('');
  const [result, setResult] = useState<AppendResult | null>(null);
  const [errors, setErrors] = useState<string[]>([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [showMappingTemplate, setShowMappingTemplate] = useState(true);
  const [showMappingHelper, setShowMappingHelper] = useState(true);
  const [showDataChecks, setShowDataChecks] = useState(true);

  const sourceMatrix = useMemo<CellValue[][]>(() => {
    if (!sourceWorkbook || !sourceSheetName) return [];
    const ws = sourceWorkbook.Sheets[sourceSheetName];
    return ws ? worksheetToMatrix(ws) : [];
  }, [sourceWorkbook, sourceSheetName]);

  const sourceHeaders = useMemo(() => (sourceMatrix[0] ?? []).map(v => String(v ?? '').trim()), [sourceMatrix]);

  const templatePreview = useMemo(() => {
    if (templateRows.length > 0) return templateRows;
    return templateText.trim() ? parseTemplateText(templateText) : [];
  }, [templateRows, templateText]);

  // Select the output mapping row; PARITY compares that row's first two
  // mapped source fields (for example Q18: I_1_Q18 vs I_2_Q18).
  const parityMappings = useMemo(() => templatePreview.filter(mapping =>
    !['SERVED', 'SERIAL', 'PARITY'].includes(mapping.targetHeader.trim().toUpperCase()) &&
    mapping.sourceHeaders.length >= 2 &&
    mapping.sourceHeaders[0].toUpperCase() !== mapping.sourceHeaders[1].toUpperCase()
  ), [templatePreview]);
  const selectedParityMapping = parityMappings.find(mapping => mapping.targetHeader === parityMappingHeader)
    ?? parityMappings[0];

  const handleSourceFile = useCallback(async (file: File) => {
    setErrors([]);
    setResult(null);
    try {
      const type = getFileType(file.name);
      const loaded = type === 'csv' ? await readCSV(file) : await readWorkbook(file);
      setSourceWorkbook(loaded.workbook);
      setSourceFileName(file.name);
      setSourceSheetName(loaded.sheetNames[0] ?? '');
    } catch (err) {
      setErrors([`Unable to read source file: ${err}`]);
    }
  }, []);

  const handleTemplateFile = useCallback(async (file: File) => {
    setErrors([]);
    setResult(null);
    try {
      const type = getFileType(file.name);
      const loaded = type === 'csv' ? await readCSV(file) : await readWorkbook(file);
      const firstSheet = loaded.workbook.SheetNames[0];
      const ws = loaded.workbook.Sheets[firstSheet];
      const matrix = worksheetToMatrix(ws);
      const rows = parseTemplateMatrix(matrix);
      setTemplateRows(rows);
      setTemplateText(matrix.map(row => row.map(v => v ?? '').join('\t')).join('\n'));
    } catch (err) {
      setErrors([`Unable to read template file: ${err}`]);
    }
  }, []);

  const handleProcess = useCallback(() => {
    setErrors([]);
    setResult(null);

    const mappings = templateText.trim() ? parseTemplateText(templateText) : templateRows;
    if (sourceMatrix.length === 0) {
      setErrors(['Upload a source file first.']);
      return;
    }
    if (mappings.length === 0) {
      setErrors(['Paste or upload a template mapping first.']);
      return;
    }

    setIsProcessing(true);
    setTimeout(() => {
      const output = appendByTemplate(sourceMatrix, mappings, {
        parityMappingHeader: selectedParityMapping?.targetHeader,
        roundValue: roundInput.trim() === '' ? undefined : roundInput.trim(),
      });
      setTemplateRows(mappings);
      setResult(output);
      if (selectedParityMapping) setParityMappingHeader(selectedParityMapping.targetHeader);
      setShowDataChecks(true);
      setIsProcessing(false);
    }, 50);
  }, [sourceMatrix, templateRows, templateText, selectedParityMapping, roundInput]);

  const resultPreviewRows = result?.rows.slice(0, 50) ?? [];

  return (
    <div className="theme-page min-h-screen text-slate-800 dark:text-slate-100 flex flex-col transition-colors w-full max-w-full overflow-x-hidden">
      {!hideHeader && (
        <HeaderNav
          currentPage="templateAppend"
          onNavigate={onNavigate}
          showDebug={showDebug}
          onToggleDebug={onToggleDebug}
        />
      )}

      <main className="flex-1 w-full max-w-[1400px] mx-auto px-4 sm:px-6 lg:px-8 py-6">
        {/* Top-Left Export Action Bar */}
        <TopExportBar
          title="Stack Data Exports"
          badge={result ? `${result.headers.length} columns · ${result.rows.length.toLocaleString()} rows` : undefined}
        >
          <button
            type="button"
            onClick={() => result && exportAppendResult(result, sourceFileName)}
            disabled={!result}
            className={`flex-1 min-w-[155px] h-9 px-3 text-xs font-semibold rounded-xl shadow-sm transition-all flex items-center justify-center gap-1.5 ${
              result
                ? 'btn-brand focus-ring-brand'
                : 'bg-slate-200 dark:bg-slate-800 text-slate-400 dark:text-slate-600 cursor-not-allowed'
            }`}
          >
            <svg className="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
              <path strokeLinecap="round" strokeLinejoin="round" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4" />
            </svg>
            Export OUTPUT.xlsx
          </button>

          <button
            type="button"
            onClick={() => result && exportAppendResultCsv(result, sourceFileName)}
            disabled={!result}
            className={`flex-1 min-w-[155px] h-9 px-3 text-xs font-semibold rounded-xl shadow-sm transition-all border flex items-center justify-center gap-1.5 ${
              result
                ? 'bg-white dark:bg-slate-800 text-brand border-brand-200 hover:bg-slate-50 dark:hover:bg-slate-700'
                : 'bg-slate-100 dark:bg-slate-800 text-slate-400 dark:text-slate-600 border-slate-200 dark:border-slate-800 cursor-not-allowed'
            }`}
          >
            <svg className="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
              <path strokeLinecap="round" strokeLinejoin="round" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
            </svg>
            Export OUTPUT.csv
          </button>
        </TopExportBar>

        <div className="grid gap-6 lg:grid-cols-[360px_minmax(0,1fr)]">
          <aside className="space-y-5">
            {/* Quick Mapping Helper */}
            <section className="rounded-2xl border border-slate-200/80 dark:border-slate-800 bg-white dark:bg-slate-900 shadow-sm overflow-hidden">
              <button
                type="button"
                onClick={() => setShowMappingHelper(v => !v)}
                aria-expanded={showMappingHelper}
                className="w-full border-b border-slate-100 dark:border-slate-800 px-4 py-3 flex items-center justify-between gap-2 text-left hover:bg-slate-50 dark:hover:bg-slate-800/60 transition-colors"
              >
                <span className="flex items-center gap-2 min-w-0">
                  <span className="text-sm font-semibold text-slate-800 dark:text-slate-200">Mapping Template</span>
                  <span className="text-[10px] px-2 py-0.5 rounded-full bg-indigo-50 dark:bg-indigo-950/40 text-indigo-700 dark:text-indigo-300 font-semibold border border-indigo-200/60 dark:border-indigo-800">
                    Auto-Serial
                  </span>
                </span>
                <span className="flex items-center gap-2 shrink-0">
                  <span className="text-[11px] font-semibold text-slate-500 dark:text-slate-400">
                    {showMappingHelper ? 'Hide' : 'Show'}
                  </span>
                  <svg className={`w-4 h-4 text-slate-400 transition-transform ${showMappingHelper ? 'rotate-180' : ''}`} fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                    <path strokeLinecap="round" strokeLinejoin="round" d="M19 9l-7 7-7-7" />
                  </svg>
                </span>
              </button>
              {showMappingHelper && (
                <div className="space-y-3 p-4 text-sm text-slate-600 dark:text-slate-400">
                  <p className="text-xs">
                    Copy a mapping range from Excel and paste it below. Column A is the target header; Columns B+ are mapped source fields.
                  </p>
                <div className="rounded-xl bg-slate-50 dark:bg-slate-800/60 p-2.5 text-[11px] text-slate-500 dark:text-slate-400 font-mono space-y-0.5">
                  <div className="text-teal-600 dark:text-teal-400 font-bold">• SERVED: auto-generated (1, 2, ...)</div>
                  <div className="text-indigo-600 dark:text-indigo-400 font-bold">• Serial: auto-numbered 1..total rows</div>
                  <div className="text-sky-600 dark:text-sky-400 font-bold">• ROUND: optional fixed value for every row</div>
                  <div className="text-purple-600 dark:text-purple-400 font-bold">• PARITY: auto-coded 1 (TRUE), 2 (FALSE)</div>
                </div>
                  <button
                    type="button"
                    onClick={() => {
                      setTemplateText(SAMPLE_TEMPLATE);
                      setTemplateRows(parseTemplateText(SAMPLE_TEMPLATE));
                      setResult(null);
                    }}
                    className="w-full rounded-xl border border-slate-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-800/80 px-3 py-2 text-xs font-semibold text-slate-700 dark:text-slate-300 hover:bg-slate-100 dark:hover:bg-slate-700 transition-colors"
                  >
                    Load Sample Mapping (with Serial)
                  </button>
                </div>
              )}
            </section>

            {/* 1. Upload Source File */}
            <section className="rounded-2xl border border-slate-200/80 dark:border-slate-800 bg-white dark:bg-slate-900 p-4 shadow-sm space-y-3">
              <h2 className="text-xs font-bold uppercase tracking-wider text-slate-400">1. Upload Source File</h2>
              <UploadBox title="Upload source file" subtitle="File containing row-1 headers" fileName={sourceFileName} onFile={handleSourceFile} />
              {sourceWorkbook && sourceWorkbook.SheetNames.length > 1 && (
                <div>
                  <label className="mb-1 block text-xs font-medium text-slate-500 dark:text-slate-400">Source sheet</label>
                  <select
                    value={sourceSheetName}
                    onChange={e => { setSourceSheetName(e.target.value); setResult(null); }}
                    className="w-full rounded-xl border border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-800 px-3 py-2 text-sm focus:border-teal-500 focus:ring-2 focus:ring-teal-500 outline-none"
                  >
                    {sourceWorkbook.SheetNames.map(name => <option key={name} value={name}>{name}</option>)}
                  </select>
                </div>
              )}
              {sourceHeaders.length > 0 && (
                <p className="text-[11px] text-slate-400 dark:text-slate-500">{sourceHeaders.length} source headers detected.</p>
              )}
            </section>

            {/* 2. Mapping Template input */}
            <section className="rounded-2xl border border-slate-200/80 dark:border-slate-800 bg-white dark:bg-slate-900 shadow-sm overflow-hidden">
              <button
                type="button"
                onClick={() => setShowMappingTemplate(v => !v)}
                aria-expanded={showMappingTemplate}
                className="w-full px-4 py-3 flex items-center justify-between text-left hover:bg-slate-50 dark:hover:bg-slate-800/60 transition-colors"
              >
                <span className="flex items-center gap-2 min-w-0">
                  <span className="text-xs font-bold uppercase tracking-wider text-slate-400">2. Mapping Template</span>
                  {templatePreview.length > 0 && (
                    <span className="text-[10px] px-2 py-0.5 rounded-full bg-slate-100 dark:bg-slate-800 text-slate-500 dark:text-slate-400 font-semibold border border-slate-200/60 dark:border-slate-700">
                      {templatePreview.length} rows
                    </span>
                  )}
                </span>
                <span className="flex items-center gap-2 shrink-0">
                  <span className="text-[11px] font-semibold text-slate-500 dark:text-slate-400">
                    {showMappingTemplate ? 'Hide' : 'Show'}
                  </span>
                  <svg className={`w-4 h-4 text-slate-400 transition-transform ${showMappingTemplate ? 'rotate-180' : ''}`} fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                    <path strokeLinecap="round" strokeLinejoin="round" d="M19 9l-7 7-7-7" />
                  </svg>
                </span>
              </button>
              {showMappingTemplate && (
                <div className="p-4 pt-0 space-y-3">
                  <UploadBox title="Upload mapping template" subtitle="Optional .xlsx / .csv file" onFile={handleTemplateFile} />
                  <div>
                    <label className="mb-1 block text-xs font-medium text-slate-500 dark:text-slate-400">Or paste tab-delimited mapping</label>
                    <textarea
                      value={templateText}
                      onChange={e => { setTemplateText(e.target.value); setTemplateRows([]); setResult(null); }}
                      placeholder={'Serial\tSERIAL\tSERIAL\nS18_BANNER\tS18_BANNER\tS18_BANNER\nPRODUCT\tI_1_PRODUCT_TRIED\tI_2_PRODUCT_TRIED\nQ1\tI_1_Q1\tI_2_Q1'}
                      className="h-32 w-full rounded-xl border border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-800 p-3 font-mono text-xs focus:border-teal-500 focus:ring-2 focus:ring-teal-500 outline-none text-slate-800 dark:text-slate-200"
                    />
                  </div>
                </div>
              )}
            </section>

            {/* 3. ROUND fixed value */}
            <section className="rounded-2xl border border-slate-200/80 dark:border-slate-800 bg-white dark:bg-slate-900 p-4 shadow-sm space-y-3">
              <div className="flex items-center justify-between gap-2">
                <h2 className="text-xs font-bold uppercase tracking-wider text-slate-400">3. ROUND Number</h2>
                <span className="text-[10px] px-2 py-0.5 rounded-full bg-sky-50 dark:bg-sky-950/40 text-sky-700 dark:text-sky-300 font-semibold border border-sky-200/60 dark:border-sky-800">
                  Fills ROUND column
                </span>
              </div>
              <p className="text-xs text-slate-500 dark:text-slate-400">
                Enter the round number/text to write into every data row of the output <span className="font-mono font-semibold">ROUND</span> column.
              </p>
              <div className="flex gap-2">
                <input
                  type="text"
                  inputMode="numeric"
                  value={roundInput}
                  onChange={e => { setRoundInput(e.target.value); setResult(null); }}
                  placeholder="e.g. 1"
                  aria-label="ROUND number to populate"
                  className="w-full rounded-xl border border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-800 px-3 py-2 text-sm font-semibold text-slate-800 dark:text-slate-200 focus-ring-brand"
                />
                {roundInput.trim() !== '' && (
                  <button
                    type="button"
                    onClick={() => { setRoundInput(''); setResult(null); }}
                    className="shrink-0 rounded-xl border border-slate-200 dark:border-slate-700 px-3 py-2 text-xs font-semibold text-slate-500 dark:text-slate-400 hover:bg-slate-50 dark:hover:bg-slate-800"
                  >
                    Clear
                  </button>
                )}
              </div>
              <p className="text-[11px] text-slate-400 dark:text-slate-500">
                {roundInput.trim() === ''
                  ? 'Leave blank to keep mapped/source ROUND values. If ROUND is absent and this is filled, the column is auto-added.'
                  : `On generate, every output row gets ROUND = "${roundInput.trim()}".`}
              </p>
            </section>

            {/* 4. PARITY compares mapped source columns B and C */}
            <section className="rounded-2xl border border-slate-200/80 dark:border-slate-800 bg-white dark:bg-slate-900 p-4 shadow-sm space-y-3">
              <div className="flex items-center justify-between">
                <h2 className="text-xs font-bold uppercase tracking-wider text-slate-400">4. PARITY Mapping Pair</h2>
                <span className="text-[10px] px-2 py-0.5 rounded-full bg-purple-50 dark:bg-purple-950/40 text-purple-700 dark:text-purple-300 font-semibold border border-purple-200/60 dark:border-purple-800">
                  Source columns 2 &amp; 3
                </span>
              </div>
              <p className="text-xs text-slate-500 dark:text-slate-400">
                Choose the mapping row to check. PARITY compares that row's second and third mapping cells (the two source headers), e.g. <span className="font-mono font-semibold">I_1_Q18</span> against <span className="font-mono font-semibold">I_2_Q18</span> for the same source respondent.
              </p>
              {parityMappings.length > 0 && selectedParityMapping ? (
                <div className="rounded-xl bg-slate-50 dark:bg-slate-800/60 p-3 text-[11px] text-slate-600 dark:text-slate-400 font-mono space-y-1">
                  <label className="block text-[10px] font-semibold text-slate-500 dark:text-slate-400">Mapping row used for PARITY</label>
                  <select
                    value={selectedParityMapping.targetHeader}
                    onChange={e => { setParityMappingHeader(e.target.value); setResult(null); }}
                    className="w-full rounded-lg border border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-800 px-2.5 py-2 text-xs font-mono text-slate-800 dark:text-slate-200 focus-ring-brand"
                  >
                    {parityMappings.map(mapping => (
                      <option key={mapping.targetHeader} value={mapping.targetHeader}>
                        {mapping.targetHeader}: {mapping.sourceHeaders[0]} vs {mapping.sourceHeaders[1]}
                      </option>
                    ))}
                  </select>
                  <div>Compare: <span className="font-bold text-slate-800 dark:text-slate-200">{selectedParityMapping.sourceHeaders[0]}</span> == <span className="font-bold text-slate-800 dark:text-slate-200">{selectedParityMapping.sourceHeaders[1]}</span></div>
                  <div className="pt-1 border-t border-slate-200 dark:border-slate-700 flex flex-wrap gap-x-4">
                    <span><span className="font-bold text-emerald-600 dark:text-emerald-400">1</span> = TRUE / equal</span>
                    <span><span className="font-bold text-amber-600 dark:text-amber-400">2</span> = FALSE / different</span>
                  </div>
                  {result?.parityStats && (
                    <div className="pt-1 border-t border-slate-200 dark:border-slate-700 flex flex-wrap gap-x-4">
                      <span className="text-emerald-600 dark:text-emerald-400 font-bold">Matches: {result.parityStats.matches.toLocaleString()}</span>
                      <span className="text-amber-600 dark:text-amber-400 font-bold">Mismatches: {result.parityStats.mismatches.toLocaleString()}</span>
                    </div>
                  )}
                </div>
              ) : (
                <p className="rounded-xl bg-amber-50 dark:bg-amber-950/30 p-3 text-xs text-amber-700 dark:text-amber-300">
                  Add a mapping row with two distinct source headers, such as <span className="font-mono">Q18 / I_1_Q18 / I_2_Q18</span>, to enable PARITY.
                </p>
              )}

              <button
                type="button"
                onClick={handleProcess}
                disabled={isProcessing}
                className="w-full rounded-xl px-4 py-3 text-sm font-bold text-white shadow-sm transition-colors btn-brand focus-ring-brand disabled:bg-slate-200 dark:disabled:bg-slate-800 disabled:text-slate-400"
              >
                {isProcessing ? 'Processing...' : '▶ Generate OUTPUT'}
              </button>
            </section>

            {errors.length > 0 && (
              <section className="rounded-2xl border border-red-200 dark:border-red-900/50 bg-red-50 dark:bg-red-950/30 p-4">
                {errors.map((e, i) => <p key={i} className="text-xs text-red-600 dark:text-red-400">{e}</p>)}
              </section>
            )}
          </aside>

          <section className="space-y-5 min-w-0">
            {/* Top Stat Cards */}
            <div className="grid gap-3 grid-cols-2 sm:grid-cols-3 xl:grid-cols-6">
              <div className="rounded-2xl border border-slate-200/80 dark:border-slate-800 bg-white dark:bg-slate-900 p-4 shadow-sm">
                <p className="text-xs text-slate-400">Source rows</p>
                <p className="mt-1 text-xl font-bold text-slate-800 dark:text-white">{Math.max(sourceMatrix.length - 1, 0).toLocaleString()}</p>
              </div>
              <div className="rounded-2xl border border-slate-200/80 dark:border-slate-800 bg-white dark:bg-slate-900 p-4 shadow-sm">
                <p className="text-xs text-slate-400">Output rows</p>
                <p className="mt-1 text-xl font-bold text-slate-800 dark:text-white">{(result?.rows.length ?? 0).toLocaleString()}</p>
              </div>
              <div className="rounded-2xl border border-slate-200/80 dark:border-slate-800 bg-white dark:bg-slate-900 p-4 shadow-sm">
                <p className="text-xs text-slate-400">Total Columns</p>
                <p className="mt-1 text-xl font-bold text-teal-600 dark:text-teal-400">{(result?.headers.length ?? 0).toLocaleString()}</p>
                <p className="text-[10px] text-slate-400 mt-0.5">+SERVED, Serial, PARITY</p>
              </div>
              <div className="rounded-2xl border border-slate-200/80 dark:border-slate-800 bg-white dark:bg-slate-900 p-4 shadow-sm">
                <p className="text-xs text-slate-400">Missing answers</p>
                <p className={`mt-1 text-xl font-bold ${result ? (result.quality && result.quality.missingCells > 0 ? 'text-red-600 dark:text-red-400' : 'text-emerald-600 dark:text-emerald-400') : 'text-slate-800 dark:text-white'}`}>
                  {result?.quality ? result.quality.missingCells.toLocaleString() : '—'}
                </p>
                <p className="text-[10px] text-slate-400 mt-0.5">
                  {result?.quality
                    ? `${result.quality.rowsWithMissing.toLocaleString()} rows affected`
                    : 'Generate OUTPUT'}
                </p>
              </div>
              <div className="rounded-2xl border border-slate-200/80 dark:border-slate-800 bg-white dark:bg-slate-900 p-4 shadow-sm">
                <p className="text-xs text-slate-400">ROUND</p>
                <p className="mt-1 text-xl font-bold text-sky-600 dark:text-sky-400">
                  {result?.roundValue !== undefined && result?.roundValue !== null && String(result.roundValue).trim() !== ''
                    ? String(result.roundValue)
                    : roundInput.trim() !== ''
                      ? roundInput.trim()
                      : '—'}
                </p>
                <p className="text-[10px] text-slate-400 mt-0.5">
                  {result?.roundMode === 'override'
                    ? 'Fixed override applied'
                    : result?.roundMode === 'auto-added'
                      ? 'Auto-added column'
                      : result?.roundMode === 'mapped'
                        ? 'From mapped source'
                        : 'No override'}
                </p>
              </div>
              <div className="rounded-2xl border border-slate-200/80 dark:border-slate-800 bg-white dark:bg-slate-900 p-4 shadow-sm">
                <p className="text-xs text-slate-400">PARITY Rate</p>
                <p className="mt-1 text-xl font-bold text-purple-600 dark:text-purple-400">
                  {result?.parityStats
                    ? `${Math.round((result.parityStats.matches / (result.rows.length || 1)) * 100)}%`
                    : '—'}
                </p>
                <p className="text-[10px] text-slate-400 mt-0.5">
                  {result?.parityStats
                    ? `${result.parityStats.matches} true / ${result.parityStats.mismatches} false`
                    : 'Select 2 columns'}
                </p>
              </div>
            </div>

            {/* Data Quality Checks */}
            {result?.quality && (
              <div className="rounded-2xl border border-slate-200/80 dark:border-slate-800 bg-white dark:bg-slate-900 shadow-sm overflow-hidden">
                <button
                  type="button"
                  onClick={() => setShowDataChecks(v => !v)}
                  aria-expanded={showDataChecks}
                  className="w-full flex items-center justify-between gap-3 px-4 py-3 text-left hover:bg-slate-50 dark:hover:bg-slate-800/60 transition-colors"
                >
                  <span>
                    <span className="block text-sm font-semibold text-slate-800 dark:text-slate-200">Data Quality Checks</span>
                    <span className="block text-xs text-slate-400">Missing answers, Serial/SERVED consistency, ROUND coverage, and PARITY availability.</span>
                  </span>
                  <span className="flex items-center gap-2 shrink-0">
                    {(() => {
                      const failed = result.quality.checks.filter(check => check.status === 'fail').length;
                      const warned = result.quality.checks.filter(check => check.status === 'warn').length;
                      if (failed > 0) {
                        return (
                          <span className="text-[10px] px-2 py-0.5 rounded-full bg-red-50 dark:bg-red-950/40 text-red-700 dark:text-red-300 font-bold border border-red-200/60 dark:border-red-800">
                            {failed} issue{failed === 1 ? '' : 's'} need attention
                          </span>
                        );
                      }
                      if (warned > 0) {
                        return (
                          <span className="text-[10px] px-2 py-0.5 rounded-full bg-amber-50 dark:bg-amber-950/40 text-amber-700 dark:text-amber-300 font-bold border border-amber-200/60 dark:border-amber-800">
                            {warned} warning{warned === 1 ? '' : 's'}
                          </span>
                        );
                      }
                      return (
                        <span className="text-[10px] px-2 py-0.5 rounded-full bg-emerald-50 dark:bg-emerald-950/40 text-emerald-700 dark:text-emerald-300 font-bold border border-emerald-200/60 dark:border-emerald-800">
                          All checks passed
                        </span>
                      );
                    })()}
                    <span className="text-[11px] font-semibold text-slate-500 dark:text-slate-400">
                      {showDataChecks ? 'Hide' : 'Show'}
                    </span>
                    <svg className={`w-4 h-4 text-slate-400 transition-transform ${showDataChecks ? 'rotate-180' : ''}`} fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                      <path strokeLinecap="round" strokeLinejoin="round" d="M19 9l-7 7-7-7" />
                    </svg>
                  </span>
                </button>

                {showDataChecks && (
                  <div className="border-t border-slate-100 dark:border-slate-800 px-4 py-4 space-y-4">
                    <div className="grid gap-2 md:grid-cols-2">
                      {result.quality.checks.map(check => (
                        <div
                          key={check.id}
                          className={`rounded-xl border p-3 ${
                            check.status === 'pass'
                              ? 'border-emerald-200/70 dark:border-emerald-900/60 bg-emerald-50/50 dark:bg-emerald-950/20'
                              : check.status === 'warn'
                                ? 'border-amber-200/70 dark:border-amber-900/60 bg-amber-50/50 dark:bg-amber-950/20'
                                : 'border-red-200/70 dark:border-red-900/60 bg-red-50/50 dark:bg-red-950/20'
                          }`}
                        >
                          <div className="flex items-center justify-between gap-2">
                            <p className="text-xs font-bold text-slate-800 dark:text-slate-200">{check.label}</p>
                            <span className={`text-[10px] px-2 py-0.5 rounded-full font-bold uppercase tracking-wide ${
                              check.status === 'pass'
                                ? 'bg-emerald-100 dark:bg-emerald-900/50 text-emerald-700 dark:text-emerald-300'
                                : check.status === 'warn'
                                  ? 'bg-amber-100 dark:bg-amber-900/50 text-amber-700 dark:text-amber-300'
                                  : 'bg-red-100 dark:bg-red-900/50 text-red-700 dark:text-red-300'
                            }`}>
                              {check.status}
                            </span>
                          </div>
                          <p className="mt-1 text-xs text-slate-600 dark:text-slate-400">{check.summary}</p>
                          {check.detail && (
                            <p className="mt-1 text-[11px] font-mono text-slate-500 dark:text-slate-400 break-words">{check.detail}</p>
                          )}
                        </div>
                      ))}
                    </div>

                    {result.quality.missingByColumn.length > 0 && (
                      <div className="overflow-auto rounded-xl border border-slate-200/70 dark:border-slate-800">
                        <table className="w-full text-xs">
                          <thead className="bg-slate-50 dark:bg-slate-800">
                            <tr>
                              <th className="px-3 py-2 text-left font-semibold text-slate-500 dark:text-slate-400">Column with missing answers</th>
                              <th className="px-3 py-2 text-right font-semibold text-slate-500 dark:text-slate-400">Missing</th>
                              <th className="px-3 py-2 text-right font-semibold text-slate-500 dark:text-slate-400">Checked rows</th>
                              <th className="px-3 py-2 text-right font-semibold text-slate-500 dark:text-slate-400">Missing %</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-slate-100 dark:divide-slate-800">
                            {result.quality.missingByColumn.slice(0, 12).map(column => (
                              <tr key={`${column.header}-${column.index}`}>
                                <td className="px-3 py-2 font-semibold text-slate-800 dark:text-slate-200">{column.header}</td>
                                <td className="px-3 py-2 text-right font-mono font-bold text-red-600 dark:text-red-400">{column.missing.toLocaleString()}</td>
                                <td className="px-3 py-2 text-right font-mono text-slate-600 dark:text-slate-400">{column.total.toLocaleString()}</td>
                                <td className="px-3 py-2 text-right font-mono text-slate-600 dark:text-slate-400">{(column.missingRate * 100).toFixed(1)}%</td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    )}

                    {result.quality.missingExamples.length > 0 && (
                      <div className="overflow-auto rounded-xl border border-slate-200/70 dark:border-slate-800">
                        <table className="w-full text-xs">
                          <thead className="bg-slate-50 dark:bg-slate-800">
                            <tr>
                              <th className="px-3 py-2 text-left font-semibold text-slate-500 dark:text-slate-400">Output row</th>
                              <th className="px-3 py-2 text-left font-semibold text-slate-500 dark:text-slate-400">Serial</th>
                              <th className="px-3 py-2 text-left font-semibold text-slate-500 dark:text-slate-400">SERVED</th>
                              <th className="px-3 py-2 text-left font-semibold text-slate-500 dark:text-slate-400">Missing columns</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-slate-100 dark:divide-slate-800">
                            {result.quality.missingExamples.map(example => (
                              <tr key={example.outputRowNumber}>
                                <td className="px-3 py-2 font-mono font-bold text-slate-800 dark:text-slate-200">{example.outputRowNumber}</td>
                                <td className="px-3 py-2 font-mono text-slate-600 dark:text-slate-400">{example.serial ?? '—'}</td>
                                <td className="px-3 py-2 font-mono text-slate-600 dark:text-slate-400">{example.served ?? '—'}</td>
                                <td className="px-3 py-2 font-mono text-red-600 dark:text-red-400">{example.missingHeaders.join(', ')}</td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    )}
                  </div>
                )}
              </div>
            )}

            {/* Template Mapping Preview */}
            <div className="rounded-2xl border border-slate-200/80 dark:border-slate-800 bg-white dark:bg-slate-900 shadow-sm overflow-hidden">
              <div className="flex items-center justify-between border-b border-slate-100 dark:border-slate-800 px-4 py-3">
                <div>
                  <h2 className="text-sm font-semibold text-slate-800 dark:text-slate-200">Template Mapping Preview</h2>
                  <p className="text-xs text-slate-400">Target output headers and their mapped source fields from row 1.</p>
                </div>
              </div>
              <div className="max-h-52 overflow-auto">
                {templatePreview.length > 0 ? (
                  <table className="w-full text-sm">
                    <thead className="sticky top-0 bg-slate-50 dark:bg-slate-800">
                      <tr>
                        <th className="w-48 min-w-[180px] px-3 py-2 text-left text-xs font-semibold text-slate-500 dark:text-slate-400">Output Header</th>
                        <th className="px-3 py-2 text-left text-xs font-semibold text-slate-500 dark:text-slate-400">Mapped Source Headers</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100 dark:divide-slate-800">
                      {templatePreview.map((row, i) => (
                        <tr key={`${row.targetHeader}-${i}`}>
                          <td className="px-3 py-2.5 font-medium text-slate-800 dark:text-slate-200">
                            <div className="flex min-h-6 items-center gap-2.5 whitespace-nowrap">
                              <span className="shrink-0">{row.targetHeader}</span>
                              {row.targetHeader.toUpperCase() === 'SERIAL' && (
                                <MappingNote><span className="mr-1 font-mono font-semibold">1..N</span> sequence</MappingNote>
                              )}
                              {row.targetHeader.toUpperCase() === 'ROUND' && roundInput.trim() !== '' && (
                                <MappingNote><span className="font-mono">= {roundInput.trim()}</span></MappingNote>
                              )}
                            </div>
                          </td>
                          <td className="px-3 py-2 font-mono text-xs text-slate-500 dark:text-slate-400">
                            {row.sourceHeaders.join(', ') || <span className="text-slate-400 italic">(auto-populated)</span>}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                ) : (
                  <div className="p-8 text-center text-sm text-slate-400">Paste or upload a mapping template.</div>
                )}
              </div>
            </div>

            {/* OUTPUT Preview */}
            <div className="rounded-2xl border border-slate-200/80 dark:border-slate-800 bg-white dark:bg-slate-900 shadow-sm overflow-hidden">
              <div className="flex flex-wrap items-center justify-between gap-x-6 gap-y-3 border-b border-slate-100 px-4 py-4 dark:border-slate-800">
                <div className="min-w-0 flex-1 basis-80">
                  <h2 className="text-sm font-semibold text-slate-800 dark:text-slate-200">OUTPUT Preview</h2>
                  <p className="mt-1 text-xs leading-relaxed text-slate-500 dark:text-slate-400">
                    {result?.parityCol1 && result?.parityCol2 ? (
                      <>PARITY compares <span className="break-all font-mono text-[11px] text-slate-700 dark:text-slate-300">{result.parityCol1}</span> with <span className="break-all font-mono text-[11px] text-slate-700 dark:text-slate-300">{result.parityCol2}</span>. Both source fields are included at the end.</>
                    ) : 'Your stacked data, with generated SERVED, Serial and ROUND fields.'}
                  </p>
                </div>
                {result?.parityStats && (
                  <ParitySummary matches={result.parityStats.matches} mismatches={result.parityStats.mismatches} />
                )}
              </div>

              <div className="overflow-auto max-h-[30rem]">
                {result ? (
                  <table className="min-w-full text-sm">
                    <thead className="sticky top-0 bg-slate-50 dark:bg-slate-800 z-10">
                      <tr>
                        {result.headers.map((h, i) => {
                          const isServed = h === 'SERVED';
                          const isSerial = h === 'Serial';
                          const isParity = h === 'PARITY';
                          const isRound = h.trim().toUpperCase() === 'ROUND';
                          const isParitySource = h === result.parityCol1 || h === result.parityCol2;
                          return (
                            <th
                              key={`${h}-${i}`}
                              className={`border-b border-slate-200 dark:border-slate-700 px-3 py-2 text-xs font-semibold whitespace-nowrap ${isParity ? 'min-w-[112px] text-center' : 'text-left'} ${
                                isServed
                                  ? 'text-teal-700 dark:text-teal-300 bg-teal-50/50 dark:bg-teal-950/30'
                                  : isSerial
                                  ? 'text-indigo-700 dark:text-indigo-300 bg-indigo-50/50 dark:bg-indigo-950/30'
                                  : isParity
                                  ? 'text-purple-700 dark:text-purple-300 bg-purple-50/50 dark:bg-purple-950/30'
                                  : isRound
                                  ? 'text-sky-700 dark:text-sky-300 bg-sky-50/50 dark:bg-sky-950/30 font-bold'
                                  : isParitySource
                                  ? 'text-amber-700 dark:text-amber-300 bg-amber-50/50 dark:bg-amber-950/30 font-bold'
                                  : 'text-slate-600 dark:text-slate-300'
                              }`}
                            >
                              {h}
                            </th>
                          );
                        })}
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100 dark:divide-slate-800">
                      {resultPreviewRows.map((row, r) => (
                        <tr key={r} className="hover:bg-slate-50/50 dark:hover:bg-slate-800/50">
                          {result.headers.map((h, c) => {
                            const val = row[c];
                            const isServed = h === 'SERVED';
                            const isSerial = h === 'Serial';
                            const isParity = h === 'PARITY';
                            const isRound = h.trim().toUpperCase() === 'ROUND';

                            if (isServed) {
                              return (
                                <td key={c} className="px-3 py-2 font-mono font-bold text-teal-600 dark:text-teal-400 bg-teal-50/20 dark:bg-teal-950/10">
                                  {val ?? '—'}
                                </td>
                              );
                            }

                            if (isSerial) {
                              return (
                                <td key={c} className="px-3 py-2 font-mono font-bold text-indigo-600 dark:text-indigo-400 bg-indigo-50/20 dark:bg-indigo-950/10">
                                  {val ?? '—'}
                                </td>
                              );
                            }

                            if (isRound) {
                              return (
                                <td key={c} className="px-3 py-2 font-mono font-bold text-sky-700 dark:text-sky-300 bg-sky-50/20 dark:bg-sky-950/10">
                                  {val ?? '—'}
                                </td>
                              );
                            }

                            if (isParity) {
                              return (
                                <td key={c} className="px-3 py-2 text-center whitespace-nowrap">
                                  <ParityBadge value={val} />
                                </td>
                              );
                            }

                            const isParitySource = h === result.parityCol1 || h === result.parityCol2;
                            if (isParitySource) {
                              return (
                                <td key={c} className="px-3 py-2 font-mono font-semibold text-amber-700 dark:text-amber-300 bg-amber-50/20 dark:bg-amber-950/10 whitespace-nowrap">
                                  {val !== null && val !== undefined && val !== '' ? (
                                    String(val)
                                  ) : (
                                    <span className="text-slate-300 dark:text-slate-600">—</span>
                                  )}
                                </td>
                              );
                            }

                            return (
                              <td key={c} className="px-3 py-2 text-slate-700 dark:text-slate-300 whitespace-nowrap">
                                {val !== null && val !== undefined && val !== '' ? (
                                  String(val)
                                ) : (
                                  <span className="text-slate-300 dark:text-slate-600">—</span>
                                )}
                              </td>
                            );
                          })}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                ) : (
                  <div className="p-12 text-center text-sm text-slate-400">Generate OUTPUT to preview results.</div>
                )}
              </div>
              {result && result.rows.length > 50 && (
                <div className="p-3 text-center text-xs text-slate-400 border-t border-slate-100 dark:border-slate-800 bg-slate-50 dark:bg-slate-900">
                  Showing first 50 rows of {result.rows.length.toLocaleString()}. All rows are included in the downloaded file.
                </div>
              )}
            </div>

            {/* Processing Log */}
            {result && result.log.length > 0 && (
              <details className="rounded-2xl border border-slate-200/80 dark:border-slate-800 bg-white dark:bg-slate-900 p-4 shadow-sm">
                <summary className="cursor-pointer text-sm font-semibold text-slate-800 dark:text-slate-200">Processing log</summary>
                <div className="mt-3 max-h-56 overflow-auto font-mono text-xs text-slate-500 dark:text-slate-400 space-y-0.5">
                  {result.log.map((line, i) => <div key={i}>{line}</div>)}
                </div>
              </details>
            )}
          </section>
        </div>
      </main>
    </div>
  );
}
