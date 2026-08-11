import { useMemo, useState, useCallback } from 'react';
import * as XLSX from 'xlsx';
import { readCSV, readWorkbook, getFileType } from './modules/excelLoader';
import {
  AppendResult,
  TemplateMappingRow,
  appendByTemplate,
  exportAppendResult,
  parseTemplateMatrix,
  parseTemplateText,
  worksheetToMatrix,
  CellValue,
} from './modules/templateAppend';
import HeaderNav, { PageId } from './components/HeaderNav';
import TopExportBar from './components/TopExportBar';

interface TemplateAppendPageProps {
  onNavigate: (page: PageId) => void;
  showDebug?: boolean;
  onToggleDebug?: () => void;
  hideHeader?: boolean;
}

const SAMPLE_TEMPLATE = [
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
  const [result, setResult] = useState<AppendResult | null>(null);
  const [errors, setErrors] = useState<string[]>([]);
  const [isProcessing, setIsProcessing] = useState(false);

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
      const output = appendByTemplate(sourceMatrix, mappings);
      setTemplateRows(mappings);
      setResult(output);
      setIsProcessing(false);
    }, 50);
  }, [sourceMatrix, templateRows, templateText]);

  const resultPreviewRows = result?.rows.slice(0, 30) ?? [];

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-50 via-white to-zinc-100 dark:from-slate-950 dark:via-slate-900 dark:to-slate-950 text-slate-800 dark:text-slate-100 flex flex-col transition-colors w-full max-w-full overflow-x-hidden">
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
        <TopExportBar title="Stack Data Exports" badge={result ? `${result.headers.length} columns` : undefined}>
          <button
            type="button"
            onClick={() => result && exportAppendResult(result, sourceFileName)}
            disabled={!result}
            className={`px-4 py-2 text-xs font-semibold rounded-xl shadow-sm transition-all flex items-center gap-1.5 ${
              result
                ? 'bg-teal-600 text-white hover:bg-teal-700 shadow-teal-500/20'
                : 'bg-slate-200 dark:bg-slate-800 text-slate-400 dark:text-slate-600 cursor-not-allowed'
            }`}
          >
            <svg className="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
              <path strokeLinecap="round" strokeLinejoin="round" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4" />
            </svg>
            Export OUTPUT.xlsx
          </button>
        </TopExportBar>

        <div className="grid gap-6 lg:grid-cols-[360px_1fr]">
          <aside className="space-y-5">
            <section className="rounded-2xl border border-slate-200/80 dark:border-slate-800 bg-white dark:bg-slate-900 shadow-sm overflow-hidden">
              <div className="border-b border-slate-100 dark:border-slate-800 px-4 py-3">
                <h2 className="text-sm font-semibold text-slate-800 dark:text-slate-200">Fastest Mapping Setup</h2>
              </div>
              <div className="space-y-3 p-4 text-sm text-slate-600 dark:text-slate-400">
                <p className="text-xs">
                  Copy a mapping range from Excel and paste it below. Column A is the target header; Columns B+ are mapped source fields.
                </p>
                <div className="rounded-xl bg-slate-50 dark:bg-slate-800/60 p-3 text-[11px] text-slate-500 dark:text-slate-400 font-mono">
                  S18_BANNER [tab] S18_BANNER [tab] S18_BANNER
                  <br />
                  PRODUCT [tab] I_1_PRODUCT_TRIED ...
                </div>
                <button
                  type="button"
                  onClick={() => { setTemplateText(SAMPLE_TEMPLATE); setTemplateRows(parseTemplateText(SAMPLE_TEMPLATE)); setResult(null); }}
                  className="w-full rounded-xl border border-slate-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-800/80 px-3 py-2 text-xs font-semibold text-slate-700 dark:text-slate-300 hover:bg-slate-100 dark:hover:bg-slate-700 transition-colors"
                >
                  Load Sample Mapping
                </button>
              </div>
            </section>

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

            <section className="rounded-2xl border border-slate-200/80 dark:border-slate-800 bg-white dark:bg-slate-900 p-4 shadow-sm space-y-3">
              <h2 className="text-xs font-bold uppercase tracking-wider text-slate-400">2. Mapping Template</h2>
              <UploadBox title="Upload mapping template" subtitle="Optional .xlsx / .csv file" onFile={handleTemplateFile} />
              <div>
                <label className="mb-1 block text-xs font-medium text-slate-500 dark:text-slate-400">Or paste tab-delimited mapping</label>
                <textarea
                  value={templateText}
                  onChange={e => { setTemplateText(e.target.value); setTemplateRows([]); setResult(null); }}
                  placeholder={'PRODUCT\tI_1_PRODUCT_TRIED\tI_2_PRODUCT_TRIED\nQ1\tI_1_Q1\tI_2_Q1'}
                  className="h-36 w-full rounded-xl border border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-800 p-3 font-mono text-xs focus:border-teal-500 focus:ring-2 focus:ring-teal-500 outline-none text-slate-800 dark:text-slate-200"
                />
              </div>
              <button
                type="button"
                onClick={handleProcess}
                disabled={isProcessing}
                className="w-full rounded-xl bg-teal-600 px-4 py-3 text-sm font-bold text-white shadow-sm hover:bg-teal-700 disabled:bg-slate-200 dark:disabled:bg-slate-800 disabled:text-slate-400 transition-colors"
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
            <div className="grid gap-3 grid-cols-2 sm:grid-cols-4">
              <div className="rounded-2xl border border-slate-200/80 dark:border-slate-800 bg-white dark:bg-slate-900 p-4 shadow-sm">
                <p className="text-xs text-slate-400">Source rows</p>
                <p className="mt-1 text-xl font-bold text-slate-800 dark:text-white">{Math.max(sourceMatrix.length - 1, 0).toLocaleString()}</p>
              </div>
              <div className="rounded-2xl border border-slate-200/80 dark:border-slate-800 bg-white dark:bg-slate-900 p-4 shadow-sm">
                <p className="text-xs text-slate-400">Mappings</p>
                <p className="mt-1 text-xl font-bold text-slate-800 dark:text-white">{templatePreview.length.toLocaleString()}</p>
              </div>
              <div className="rounded-2xl border border-slate-200/80 dark:border-slate-800 bg-white dark:bg-slate-900 p-4 shadow-sm">
                <p className="text-xs text-slate-400">Output columns</p>
                <p className="mt-1 text-xl font-bold text-teal-600 dark:text-teal-400">{(result?.headers.length ?? 0).toLocaleString()}</p>
              </div>
              <div className="rounded-2xl border border-slate-200/80 dark:border-slate-800 bg-white dark:bg-slate-900 p-4 shadow-sm">
                <p className="text-xs text-slate-400">Missing headers</p>
                <p className="mt-1 text-xl font-bold text-amber-600 dark:text-amber-400">{(result?.missingHeaders.length ?? 0).toLocaleString()}</p>
              </div>
            </div>

            <div className="rounded-2xl border border-slate-200/80 dark:border-slate-800 bg-white dark:bg-slate-900 shadow-sm overflow-hidden">
              <div className="flex items-center justify-between border-b border-slate-100 dark:border-slate-800 px-4 py-3">
                <div>
                  <h2 className="text-sm font-semibold text-slate-800 dark:text-slate-200">Template Mapping Preview</h2>
                  <p className="text-xs text-slate-400">Rows become output columns; B onward are appended source fields.</p>
                </div>
              </div>
              <div className="max-h-56 overflow-auto">
                {templatePreview.length > 0 ? (
                  <table className="w-full text-sm">
                    <thead className="sticky top-0 bg-slate-50 dark:bg-slate-800">
                      <tr>
                        <th className="px-3 py-2 text-left text-xs font-semibold text-slate-500 dark:text-slate-400">Output header</th>
                        <th className="px-3 py-2 text-left text-xs font-semibold text-slate-500 dark:text-slate-400">Mapped source headers</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100 dark:divide-slate-800">
                      {templatePreview.map((row, i) => (
                        <tr key={`${row.targetHeader}-${i}`}>
                          <td className="px-3 py-2 font-medium text-slate-800 dark:text-slate-200">{row.targetHeader}</td>
                          <td className="px-3 py-2 font-mono text-xs text-slate-500 dark:text-slate-400">{row.sourceHeaders.join(', ') || '—'}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                ) : (
                  <div className="p-8 text-center text-sm text-slate-400">Paste or upload a mapping template.</div>
                )}
              </div>
            </div>

            <div className="rounded-2xl border border-slate-200/80 dark:border-slate-800 bg-white dark:bg-slate-900 shadow-sm overflow-hidden">
              <div className="flex items-center justify-between border-b border-slate-100 dark:border-slate-800 px-4 py-3">
                <div>
                  <h2 className="text-sm font-semibold text-slate-800 dark:text-slate-200">OUTPUT Preview</h2>
                  <p className="text-xs text-slate-400">SERVED column inserted first (1 = 1st amend, 2 = 2nd amend), followed by mapped output columns.</p>
                </div>
              </div>
              <div className="overflow-auto max-h-[28rem]">
                {result ? (
                  <table className="min-w-full text-sm">
                    <thead className="sticky top-0 bg-slate-50 dark:bg-slate-800">
                      <tr>
                        {result.headers.map((h, i) => (
                          <th key={`${h}-${i}`} className="border-b border-slate-200 dark:border-slate-700 px-3 py-2 text-left text-xs font-semibold text-slate-600 dark:text-slate-300">
                            {h}
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100 dark:divide-slate-800">
                      {resultPreviewRows.map((row, r) => (
                        <tr key={r} className="hover:bg-slate-50/50 dark:hover:bg-slate-800/50">
                          {result.headers.map((_, c) => (
                            <td key={c} className={`px-3 py-2 text-slate-700 dark:text-slate-300 ${c === 0 ? 'font-mono font-bold text-teal-600 dark:text-teal-400' : ''}`}>
                              {row[c] ?? <span className="text-slate-300 dark:text-slate-600">—</span>}
                            </td>
                          ))}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                ) : (
                  <div className="p-12 text-center text-sm text-slate-400">Generate OUTPUT to preview results.</div>
                )}
              </div>
            </div>

            {result && result.log.length > 0 && (
              <details className="rounded-2xl border border-slate-200/80 dark:border-slate-800 bg-white dark:bg-slate-900 p-4 shadow-sm">
                <summary className="cursor-pointer text-sm font-semibold text-slate-800 dark:text-slate-200">Processing log</summary>
                <div className="mt-3 max-h-56 overflow-auto font-mono text-xs text-slate-500 dark:text-slate-400">
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
