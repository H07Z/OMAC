import type { ReactNode } from 'react';
import type { CellValue } from '../modules/templateAppend';

export function MappingNote({ children }: { children: ReactNode }) {
  return (
    <span className="stack-mapping-note inline-flex shrink-0 items-center whitespace-nowrap rounded-md bg-brand-50 px-2 py-1 text-[10px] font-medium leading-none text-brand">
      {children}
    </span>
  );
}

export function ParitySummary({ matches, mismatches }: { matches: number; mismatches: number }) {
  return (
    <div role="group" aria-label="PARITY comparison totals" className="flex max-w-full shrink-0 flex-wrap items-center gap-x-5 gap-y-2">
      <span className="inline-flex items-center gap-2 whitespace-nowrap text-xs font-medium text-emerald-700 dark:text-emerald-300">
        <span aria-hidden="true" className="h-1.5 w-1.5 shrink-0 rounded-full bg-emerald-500" />
        <span>TRUE <span className="text-[10px] opacity-70">(1)</span></span>
        <span className="font-mono text-sm font-semibold tabular-nums text-slate-800 dark:text-slate-100">{matches.toLocaleString()}</span>
      </span>
      <span className="inline-flex items-center gap-2 whitespace-nowrap text-xs font-medium text-amber-700 dark:text-amber-300">
        <span aria-hidden="true" className="h-1.5 w-1.5 shrink-0 rounded-full bg-amber-500" />
        <span>FALSE <span className="text-[10px] opacity-70">(2)</span></span>
        <span className="font-mono text-sm font-semibold tabular-nums text-slate-800 dark:text-slate-100">{mismatches.toLocaleString()}</span>
      </span>
    </div>
  );
}

export function ParityBadge({ value }: { value: CellValue }) {
  if (value !== 1 && value !== 2) {
    return <span className="text-xs text-slate-400 dark:text-slate-500">{value ?? 'Not calculated'}</span>;
  }

  const isMatch = value === 1;
  return (
    <span
      title={isMatch ? 'Compared values match' : 'Compared values differ'}
      aria-label={`PARITY ${value}: ${isMatch ? 'TRUE, values match' : 'FALSE, values differ'}`}
      className={`inline-flex h-6 min-w-[84px] shrink-0 items-center justify-center gap-2 whitespace-nowrap rounded-md px-2.5 text-[11px] font-medium leading-none ${
        isMatch
          ? 'bg-emerald-50 text-emerald-700 dark:bg-emerald-950/50 dark:text-emerald-300'
          : 'bg-amber-50 text-amber-700 dark:bg-amber-950/50 dark:text-amber-300'
      }`}
    >
      <span className="font-mono font-semibold tabular-nums">{value}</span>
      <span aria-hidden="true" className="h-2.5 w-px bg-current opacity-25" />
      <span className="text-[10px] tracking-wide">{isMatch ? 'TRUE' : 'FALSE'}</span>
    </span>
  );
}