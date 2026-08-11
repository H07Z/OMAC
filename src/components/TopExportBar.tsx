import { ReactNode } from 'react';

interface TopExportBarProps {
  children: ReactNode;
  title?: string;
  badge?: string;
}

export default function TopExportBar({ children, title, badge }: TopExportBarProps) {
  return (
    <div className="w-full mb-5 p-3.5 bg-white dark:bg-slate-900 rounded-2xl border border-slate-200/80 dark:border-slate-800 shadow-sm transition-colors">
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
        <div className="flex items-center gap-2">
          <div className="w-7 h-7 rounded-xl bg-indigo-50 dark:bg-indigo-950/60 text-indigo-600 dark:text-indigo-400 border border-indigo-100 dark:border-indigo-900 flex items-center justify-center shrink-0">
            <svg className="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
              <path strokeLinecap="round" strokeLinejoin="round" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4" />
            </svg>
          </div>
          <div>
            <span className="text-xs font-bold text-slate-800 dark:text-white uppercase tracking-wider">
              {title || 'Quick Exports'}
            </span>
            {badge && (
              <span className="ml-2 px-2 py-0.5 text-[10px] font-semibold bg-indigo-50 text-indigo-700 dark:bg-indigo-950 dark:text-indigo-300 rounded-full border border-indigo-200/60 dark:border-indigo-800">
                {badge}
              </span>
            )}
          </div>
        </div>

        {/* Top-left Export Actions Container */}
        <div className="flex flex-wrap items-center gap-2">
          {children}
        </div>
      </div>
    </div>
  );
}
