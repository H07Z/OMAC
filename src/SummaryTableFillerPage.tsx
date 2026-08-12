import { useState, useCallback } from 'react';
import * as XLSX from 'xlsx';
import XLSXStyle from 'xlsx-js-style';
import { readCSV, readWorkbook, getFileType } from './modules/excelLoader';
import {
  detectSummaryTemplateMode,
  fillSummaryTemplateWorkbook,
  MatchStats,
} from './modules/templateFiller';
import { debugLog } from './utils/debug';
import HeaderNav, { PageId } from './components/HeaderNav';
import TopExportBar from './components/TopExportBar';

interface SummaryTableFillerPageProps {
  onNavigate: (page: PageId) => void;
  showDebug?: boolean;
  onToggleDebug?: () => void;
  hideHeader?: boolean;
}

interface SummaryRun {
  sheetName: string;
  sourceSheetName: string;
  mode: 'HD' | 'JR';
  stats: MatchStats;
}

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
        accept=".xlsx,.xls"
        onChange={e => {
          const file = e.target.files?.[0];
          if (file) onFile(file);
        }}
        className="absolute inset-0 h-full w-full cursor-pointer opacity-0"
      />
      <div className="mx-auto mb-2 flex h-10 w-10 items-center justify-center rounded-xl bg-teal-50 dark:bg-teal-950/60 text-teal-600 dark:text-teal-400">
        <svg className="h-5 w-5" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
          <path strokeLinecap="round" strokeLinejoin="round" d="M12 16V4m0 0l-4 4m4-4l4 4M4 20h16" />
        </svg>
      </div>
      <p className="text-sm font-semibold text-slate-800 dark:text-slate-200">{fileName || title}</p>
      <p className="mt-0.5 text-xs text-slate-400 dark:text-slate-500">{fileName ? subtitle : `${subtitle} · drop or browse`}</p>
    </div>
  );
}

export default function SummaryTableFillerPage({ onNavigate, showDebug, onToggleDebug, hideHeader }: SummaryTableFillerPageProps) {
  const [sourceFile, setSourceFile] = useState<File | null>(null);
  const [sourceWb, setSourceWb] = useState<XLSX.WorkBook | null>(null);
  const [sourceSheet, setSourceSheet] = useState('');

  const [templateFile, setTemplateFile] = useState<File | null>(null);
  const [templateWb, setTemplateWb] = useState<XLSX.WorkBook | null>(null);
  const [templateSheet, setTemplateSheet] = useState('');

  const [filledWb, setFilledWb] = useState<XLSX.WorkBook | null>(null);
  const [stats, setStats] = useState<MatchStats | null>(null);
  const [summaryRuns, setSummaryRuns] = useState<SummaryRun[]>([]);
  const [log, setLog] = useState<string[]>([]);
  const [isProcessing, setIsProcessing] = useState(false);

  const handleSourceFile = useCallback(async (file: File) => {
    setLog([]);
    setFilledWb(null);
    setStats(null);
    setSummaryRuns([]);
    try {
      const type = getFileType(file.name);
      const loaded = type === 'csv' ? await readCSV(file) : await readWorkbook(file);
      setSourceFile(file);
      setSourceWb(loaded.workbook);
      setSourceSheet(loaded.sheetNames[0] || '');
    } catch (err) {
      setLog([`Failed to read source file: ${err}`]);
    }
  }, []);

  const handleTemplateFile = useCallback(async (file: File) => {
    setLog([]);
    setFilledWb(null);
    setStats(null);
    setSummaryRuns([]);
    try {
      const type = getFileType(file.name);
      const loaded = type === 'csv' ? await readCSV(file) : await readWorkbook(file);
      setTemplateFile(file);
      setTemplateWb(loaded.workbook);
      setTemplateSheet(loaded.sheetNames[0] || '');
    } catch (err) {
      setLog([`Failed to read template file: ${err}`]);
    }
  }, []);

  const processAndFill = useCallback(async () => {
    if (!sourceWb || !templateWb) {
      setLog(['Please upload both source data and template files.']);
      return;
    }

    setIsProcessing(true);
    setLog([]);

    try {
      const targetSheets = templateWb.SheetNames
        .map(name => ({ name, mode: detectSummaryTemplateMode(templateWb.Sheets[name]) }))
        .filter((item): item is { name: string; mode: 'HD' | 'JR' } => item.mode !== null);

      if (targetSheets.length === 0) {
        setLog(['No HEDONICS / LIKING or JUST RIGHT template sheet was detected.']);
        return;
      }

      const chooseSourceSheet = (mode: 'HD' | 'JR'): string => {
        const preferred = sourceWb.SheetNames.find(name => {
          const upper = name.toUpperCase();
          return mode === 'HD'
            ? upper === 'HD' || upper.includes('HEDONIC') || upper.includes('LIKING')
            : upper === 'JR' || upper.includes('JUST RIGHT') || upper.includes('JUSTRIGHT');
        });
        return preferred || sourceSheet || sourceWb.SheetNames[0];
      };

      let currentWorkbook = templateWb;
      const runs: SummaryRun[] = [];
      const combinedStats: MatchStats = { filledCells: [], unmatchedRows: [], log: [] };

      for (const target of targetSheets) {
        const sourceSheetName = chooseSourceSheet(target.mode);
        const sourceWs = sourceWb.Sheets[sourceSheetName];
        if (!sourceWs) continue;

        const result = fillSummaryTemplateWorkbook(
          sourceWs,
          currentWorkbook,
          target.name,
        );
        currentWorkbook = result.workbook;
        runs.push({
          sheetName: target.name,
          sourceSheetName,
          mode: target.mode,
          stats: result.stats,
        });
        combinedStats.filledCells.push(...result.stats.filledCells);
        combinedStats.unmatchedRows.push(...result.stats.unmatchedRows.map(code => `${target.name}: ${code}`));
        combinedStats.log.push(
          `[${target.mode === 'HD' ? 'Hedonics' : 'Just Right'}] ${target.name} ← ${sourceSheetName}`,
          ...result.stats.log,
        );
      }

      setFilledWb(currentWorkbook);
      setStats(combinedStats);
      setSummaryRuns(runs);
      setLog([
        ...combinedStats.log,
        `Filled ${runs.length} summary sheet(s) simultaneously.`,
        `Total filled cells: ${combinedStats.filledCells.length}`,
        `Total unmatched question rows: ${combinedStats.unmatchedRows.length}`,
      ]);
      debugLog('SummaryFiller', `Auto-fill completed for ${runs.length} sheets with ${combinedStats.filledCells.length} cells.`);
    } catch (err) {
      setLog([`Error during processing: ${err}`]);
    } finally {
      setIsProcessing(false);
    }
  }, [sourceWb, sourceSheet, templateWb, templateSheet]);

  const handleDownload = useCallback(() => {
    if (!filledWb) return;
    const baseName = (templateFile?.name || 'template').replace(/\.(xlsx|xls)$/i, '');
    XLSXStyle.writeFile(filledWb as any, `${baseName}_FILLED.xlsx`, { bookType: 'xlsx' });
  }, [filledWb, templateFile]);

  return (
    <div className="theme-page min-h-screen text-slate-800 dark:text-slate-100 flex flex-col transition-colors w-full max-w-full overflow-x-hidden">
      {!hideHeader && (
        <HeaderNav
          currentPage="summaryFiller"
          onNavigate={onNavigate}
          showDebug={showDebug}
          onToggleDebug={onToggleDebug}
        />
      )}

      <main className="flex-1 w-full max-w-[1400px] mx-auto px-4 sm:px-6 lg:px-8 py-6">
        {/* Top-Left Export Action Bar */}
        <TopExportBar title="Summary Topline Exports" badge={filledWb ? 'Ready' : undefined}>
          <button
            type="button"
            onClick={handleDownload}
            disabled={!filledWb}
            className={`flex-1 min-w-[155px] h-9 px-3 text-xs font-semibold rounded-xl shadow-sm transition-all flex items-center justify-center gap-1.5 ${
              filledWb
                ? 'btn-brand focus-ring-brand'
                : 'bg-slate-200 dark:bg-slate-800 text-slate-400 dark:text-slate-600 cursor-not-allowed'
            }`}
          >
            <svg className="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
              <path strokeLinecap="round" strokeLinejoin="round" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4" />
            </svg>
            Download Filled Template (.xlsx)
          </button>
        </TopExportBar>

        <div className="grid gap-6 lg:grid-cols-[360px_1fr]">
          <aside className="space-y-5">
            <section className="rounded-2xl border border-slate-200/80 dark:border-slate-800 bg-white dark:bg-slate-900 p-4 shadow-sm space-y-3">
              <h2 className="text-xs font-bold uppercase tracking-wider text-slate-400">1. Upload Source Data</h2>
              <UploadBox title="Upload Source Data" subtitle="Contains Tables 3, 4, 5, 6" fileName={sourceFile?.name} onFile={handleSourceFile} />
              {sourceWb && sourceWb.SheetNames.length > 1 && (
                <div>
                  <label className="mb-1 block text-xs font-medium text-slate-500 dark:text-slate-400">Source worksheet</label>
                  <select
                    value={sourceSheet}
                    onChange={e => { setSourceSheet(e.target.value); setFilledWb(null); setStats(null); setSummaryRuns([]); setLog([]); }}
                    className="w-full rounded-xl border border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-800 px-3 py-2 text-sm focus:border-teal-500 focus:ring-2 focus:ring-teal-500 outline-none"
                  >
                    {sourceWb.SheetNames.map(name => <option key={name} value={name}>{name}</option>)}
                  </select>
                </div>
              )}
            </section>

            <section className="rounded-2xl border border-slate-200/80 dark:border-slate-800 bg-white dark:bg-slate-900 p-4 shadow-sm space-y-3">
              <h2 className="text-xs font-bold uppercase tracking-wider text-slate-400">2. Upload Template</h2>
              <UploadBox title="Upload Template" subtitle="Hedonics &amp; Just Right grid" fileName={templateFile?.name} onFile={handleTemplateFile} />
              {templateWb && templateWb.SheetNames.length > 1 && (
                <div>
                  <label className="mb-1 block text-xs font-medium text-slate-500 dark:text-slate-400">Template worksheet</label>
                  <select
                    value={templateSheet}
                    onChange={e => { setTemplateSheet(e.target.value); setFilledWb(null); setStats(null); setSummaryRuns([]); setLog([]); }}
                    className="w-full rounded-xl border border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-800 px-3 py-2 text-sm focus:border-teal-500 focus:ring-2 focus:ring-teal-500 outline-none"
                  >
                    {templateWb.SheetNames.map(name => <option key={name} value={name}>{name}</option>)}
                  </select>
                </div>
              )}
            </section>

            <button
              type="button"
              onClick={processAndFill}
              disabled={isProcessing || !sourceWb || !templateWb}
              className="w-full py-3 rounded-xl font-bold text-sm btn-brand focus-ring-brand disabled:bg-slate-200 dark:disabled:bg-slate-800 disabled:text-slate-400 transition-colors shadow-sm flex items-center justify-center gap-2"
            >
              {isProcessing ? (
                <>
                  <svg className="w-5 h-5 animate-spin" fill="none" viewBox="0 0 24 24">
                    <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" />
                    <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z" />
                  </svg>
                  Auto-filling cells...
                </>
              ) : (
                '▶ Auto-Fill Summary Tables'
              )}
            </button>
          </aside>

          <section className="space-y-5 min-w-0">
            {/* Visual Review Sections per Summary Run */}
            {summaryRuns.length > 0 && (
              <div className="grid gap-5 xl:grid-cols-2">
                {summaryRuns.map(run => {
                  const metricOrder = run.mode === 'HD'
                    ? ['topbox', 'top2box', 'mean']
                    : ['toostrong', 'justright', 'tooweak'];
                  const metricLabel: Record<string, string> = {
                    topbox: 'TOP BOX SCORE',
                    top2box: 'TOP-2 BOX SCORE',
                    mean: 'MEAN SCORE',
                    toostrong: 'MUCH/SOME TOO STRONG',
                    justright: 'JUST RIGHT',
                    tooweak: 'MUCH/SOME TOO WEAK',
                  };
                  const questionCodes = Array.from(new Set(
                    run.stats.filledCells
                      .filter(cell => cell.questionCode !== 'ALL' && cell.metric !== 'base')
                      .map(cell => cell.questionCode),
                  ));

                  return (
                    <div key={run.sheetName} className="rounded-2xl border border-slate-200/80 dark:border-slate-800 bg-white dark:bg-slate-900 shadow-sm overflow-hidden">
                      <div className={`px-4 py-3 border-b ${run.mode === 'HD' ? 'bg-indigo-50 dark:bg-indigo-950/40 border-indigo-100 dark:border-indigo-900' : 'bg-teal-50 dark:bg-teal-950/40 border-teal-100 dark:border-teal-900'}`}>
                        <div className="flex items-center justify-between gap-3">
                          <div>
                            <h3 className="text-sm font-bold text-slate-900 dark:text-white">
                              {run.mode === 'HD' ? 'HEDONICS / LIKING SUMMARY TABLE' : 'JUST RIGHT SUMMARY TABLE'}
                            </h3>
                            <p className="text-[11px] text-slate-500 dark:text-slate-400">{run.sheetName} ← source sheet {run.sourceSheetName}</p>
                          </div>
                          <span className={`rounded-full px-2.5 py-1 text-[10px] font-semibold ${run.mode === 'HD' ? 'bg-indigo-600 text-white' : 'bg-teal-600 text-white'}`}>
                            {run.stats.filledCells.length} filled
                          </span>
                        </div>
                      </div>

                      <div className="grid grid-cols-3 gap-2 p-3">
                        {metricOrder.map(metric => (
                          <div key={metric} className="rounded-xl border border-slate-100 dark:border-slate-800 bg-slate-50 dark:bg-slate-800/60 p-2.5">
                            <p className="text-[9px] font-semibold text-slate-500 dark:text-slate-400 leading-tight">{metricLabel[metric]}</p>
                            <p className="mt-1 text-xl font-bold text-slate-800 dark:text-white">
                              {run.stats.filledCells.filter(cell => cell.metric === metric).length}
                            </p>
                            {run.mode === 'HD' && metric === 'topbox' && (
                              <p className="mt-1 text-[9px] text-indigo-600 dark:text-indigo-400">Source: Like extremely</p>
                            )}
                          </div>
                        ))}
                      </div>

                      <div className="max-h-80 overflow-auto border-t border-slate-100 dark:border-slate-800">
                        <table className="w-full text-xs">
                          <thead className="sticky top-0 bg-white dark:bg-slate-900">
                            <tr className="border-b border-slate-100 dark:border-slate-800">
                              <th className="px-3 py-2 text-left font-semibold text-slate-500 dark:text-slate-400">Question</th>
                              <th className="px-3 py-2 text-left font-semibold text-slate-500 dark:text-slate-400">Metric</th>
                              <th className="px-3 py-2 text-right font-semibold text-slate-500 dark:text-slate-400">P1 X</th>
                              <th className="px-3 py-2 text-right font-semibold text-slate-500 dark:text-slate-400">P1 Y</th>
                              <th className="px-3 py-2 text-right font-semibold text-slate-500 dark:text-slate-400">P2 X</th>
                              <th className="px-3 py-2 text-right font-semibold text-slate-500 dark:text-slate-400">P2 Y</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-slate-50 dark:divide-slate-800">
                            {questionCodes.slice(0, 30).flatMap(code => metricOrder.map(metric => {
                              const find = (panel: number, product: string) => run.stats.filledCells.find(cell =>
                                cell.questionCode === code && cell.metric === metric && cell.panel === panel && cell.product === product,
                              );
                              const renderValue = (cell: ReturnType<typeof find>) => {
                                if (!cell) return '—';
                                const formattedVal = cell.displayValue || (
                                  typeof cell.value === 'number'
                                    ? (cell.metric === 'mean' ? cell.value.toFixed(2) : cell.value.toFixed(1))
                                    : String(cell.value)
                                );
                                return `${formattedVal}${cell.marker ? ` ${cell.marker}` : ''}`;
                              };
                              return (
                                <tr key={`${code}-${metric}`} className={run.mode === 'HD' && metric === 'topbox' ? 'bg-indigo-50/30 dark:bg-indigo-950/20' : ''}>
                                  <td className="px-3 py-1.5 font-semibold text-slate-800 dark:text-slate-200">{code}</td>
                                  <td className="px-3 py-1.5 text-[10px] text-slate-500 dark:text-slate-400">{metricLabel[metric]}</td>
                                  <td className="px-3 py-1.5 text-right font-mono">{renderValue(find(1, 'PRODUCT X'))}</td>
                                  <td className="px-3 py-1.5 text-right font-mono">{renderValue(find(1, 'PRODUCT Y'))}</td>
                                  <td className="px-3 py-1.5 text-right font-mono">{renderValue(find(2, 'PRODUCT X'))}</td>
                                  <td className="px-3 py-1.5 text-right font-mono">{renderValue(find(2, 'PRODUCT Y'))}</td>
                                </tr>
                              );
                            }))}
                          </tbody>
                        </table>
                        {questionCodes.length > 30 && (
                          <p className="p-2 text-center text-[10px] text-slate-400">Showing first 30 questions. All questions are populated in the workbook.</p>
                        )}
                      </div>
                    </div>
                  );
                })}
              </div>
            )}

            {stats && (
              <div className="rounded-2xl border border-slate-200/80 dark:border-slate-800 bg-white dark:bg-slate-900 shadow-sm overflow-hidden">
                <div className="px-4 py-3 border-b border-slate-100 dark:border-slate-800 flex items-center justify-between">
                  <div>
                    <h3 className="text-sm font-semibold text-slate-800 dark:text-slate-200">Review Mapped Values Log</h3>
                    <p className="text-xs text-slate-400">Detailed row-by-row and cell-by-cell mapping verification.</p>
                  </div>
                  <div className="text-xs text-slate-500">
                    {stats.filledCells.length.toLocaleString()} filled · {stats.unmatchedRows.length.toLocaleString()} unmatched
                  </div>
                </div>

                <div className="max-h-80 overflow-auto border-t border-slate-100 dark:border-slate-800">
                  <table className="w-full text-sm">
                    <thead className="sticky top-0 bg-slate-50 dark:bg-slate-800">
                      <tr className="border-b border-slate-200 dark:border-slate-700">
                        <th className="px-3 py-2 text-left text-xs font-semibold text-slate-500 dark:text-slate-400">Question</th>
                        <th className="px-3 py-2 text-left text-xs font-semibold text-slate-500 dark:text-slate-400">Panel</th>
                        <th className="px-3 py-2 text-left text-xs font-semibold text-slate-500 dark:text-slate-400">Product</th>
                        <th className="px-3 py-2 text-left text-xs font-semibold text-slate-500 dark:text-slate-400">Metric</th>
                        <th className="px-3 py-2 text-right text-xs font-semibold text-slate-500 dark:text-slate-400">Value</th>
                        <th className="px-3 py-2 text-right text-xs font-semibold text-slate-500 dark:text-slate-400">Marker below</th>
                        <th className="px-3 py-2 text-right text-xs font-semibold text-slate-500 dark:text-slate-400">Template cell</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100 dark:divide-slate-800">
                      {stats.filledCells.slice(0, 300).map((cell, index) => (
                        <tr key={`${cell.row}-${cell.col}-${index}`} className={cell.metric === 'topbox' ? 'bg-teal-50/30 dark:bg-teal-950/20' : ''}>
                          <td className="px-3 py-2 font-semibold text-slate-800 dark:text-slate-200">{cell.questionCode}</td>
                          <td className="px-3 py-2 text-slate-600 dark:text-slate-400">Panel {cell.panel}</td>
                          <td className="px-3 py-2 text-slate-600 dark:text-slate-400">{cell.product}</td>
                          <td className="px-3 py-2 font-mono text-xs text-slate-500 dark:text-slate-400">{cell.metric}</td>
                          <td className="px-3 py-2 text-right font-mono font-semibold text-teal-700 dark:text-teal-400">
                            {cell.displayValue || (typeof cell.value === 'number' ? (cell.metric === 'mean' ? cell.value.toFixed(2) : cell.value.toFixed(1)) : String(cell.value))}
                          </td>
                          <td className="px-3 py-2 text-right font-mono font-bold text-red-600 dark:text-red-400">{cell.marker || '—'}</td>
                          <td className="px-3 py-2 text-right font-mono text-xs text-slate-400">
                            {XLSX.utils.encode_cell({ r: cell.row, c: cell.col })}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                  {stats.filledCells.length > 300 && (
                    <p className="p-3 text-center text-xs text-slate-400">Showing first 300 filled cells. Full output is included in the download.</p>
                  )}
                </div>
              </div>
            )}

            {log.length > 0 && (
              <div className="rounded-2xl border border-slate-200/80 dark:border-slate-800 bg-white dark:bg-slate-900 p-4 shadow-sm">
                <h3 className="text-sm font-semibold text-slate-800 dark:text-slate-200 mb-2">Processing Log</h3>
                <div className="max-h-56 overflow-auto text-xs font-mono bg-slate-900 text-emerald-400 p-3 rounded-xl">
                  {log.map((l, i) => <div key={i}>{l}</div>)}
                </div>
              </div>
            )}
          </section>
        </div>
      </main>
    </div>
  );
}
