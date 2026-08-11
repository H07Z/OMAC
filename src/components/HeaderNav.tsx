import { useState } from 'react';
import { useTheme, COLOR_PALETTES } from '../context/ThemeContext';

export type PageId = 'processor' | 'templateAppend' | 'summaryFiller';

interface HeaderNavProps {
  currentPage: PageId;
  onNavigate: (page: PageId) => void;
  showDebug?: boolean;
  onToggleDebug?: () => void;
}

export default function HeaderNav({
  currentPage,
  onNavigate,
  showDebug,
  onToggleDebug,
}: HeaderNavProps) {
  const { mode, toggleMode, palette, setPaletteId } = useTheme();
  const [showThemePicker, setShowThemePicker] = useState(false);

  return (
    <header className="sticky top-0 z-30 border-b border-slate-200/80 dark:border-slate-800 bg-white/95 dark:bg-slate-900/95 backdrop-blur-md shadow-xs transition-colors">
      <div className="mx-auto flex max-w-[1400px] w-full items-center justify-between px-4 py-2.5 sm:px-6 lg:px-8">
        {/* Left: Brand with E# logo */}
        <div className="flex items-center gap-3 shrink-0">
          <div className="w-10 h-10 rounded-2xl bg-gradient-to-br from-indigo-600 via-indigo-500 to-violet-600 dark:from-indigo-500 dark:to-violet-500 flex items-center justify-center text-white font-extrabold text-base tracking-tighter shadow-md shadow-indigo-500/20 select-none">
            E#
          </div>
          <div>
            <h1 className="text-base font-bold text-slate-900 dark:text-white leading-tight tracking-tight">
              EZ Web Toolkit
            </h1>
            <p className="text-[11px] font-medium text-slate-400 dark:text-slate-500">
              VBA Migration Suite
            </p>
          </div>
        </div>

        {/* Center: Uiverse thin-owl-11 style segmented pill navigation */}
        <div className="hidden sm:flex items-center bg-slate-100/90 dark:bg-slate-800/90 p-1 rounded-full border border-slate-200/80 dark:border-slate-700/80 shadow-inner gap-1">
          {/* Page 1: OE Entries */}
          <button
            type="button"
            onClick={() => onNavigate('processor')}
            className={`px-4 py-1.5 rounded-full text-xs font-semibold flex items-center gap-2 transition-all duration-200 ${
              currentPage === 'processor'
                ? 'bg-white dark:bg-slate-950 text-slate-900 dark:text-white shadow-md shadow-slate-200/50 dark:shadow-none'
                : 'text-slate-600 dark:text-slate-400 hover:text-slate-900 dark:hover:text-slate-200'
            }`}
          >
            <span
              className={`w-5 h-5 rounded-full flex items-center justify-center text-[10px] font-bold ${
                currentPage === 'processor'
                  ? 'bg-indigo-600 text-white'
                  : 'bg-indigo-100 text-indigo-700 dark:bg-indigo-950 dark:text-indigo-300'
              }`}
            >
              O
            </span>
            OE Entries
          </button>

          {/* Page 2: Stack Data */}
          <button
            type="button"
            onClick={() => onNavigate('templateAppend')}
            className={`px-4 py-1.5 rounded-full text-xs font-semibold flex items-center gap-2 transition-all duration-200 ${
              currentPage === 'templateAppend'
                ? 'bg-white dark:bg-slate-950 text-slate-900 dark:text-white shadow-md shadow-slate-200/50 dark:shadow-none'
                : 'text-slate-600 dark:text-slate-400 hover:text-slate-900 dark:hover:text-slate-200'
            }`}
          >
            <span
              className={`w-5 h-5 rounded-full flex items-center justify-center text-[10px] font-bold ${
                currentPage === 'templateAppend'
                  ? 'bg-teal-600 text-white'
                  : 'bg-teal-100 text-teal-700 dark:bg-teal-950 dark:text-teal-300'
              }`}
            >
              D
            </span>
            Stack Data
          </button>

          {/* Page 3: Summary Topline */}
          <button
            type="button"
            onClick={() => onNavigate('summaryFiller')}
            className={`px-4 py-1.5 rounded-full text-xs font-semibold flex items-center gap-2 transition-all duration-200 ${
              currentPage === 'summaryFiller'
                ? 'bg-white dark:bg-slate-950 text-slate-900 dark:text-white shadow-md shadow-slate-200/50 dark:shadow-none'
                : 'text-slate-600 dark:text-slate-400 hover:text-slate-900 dark:hover:text-slate-200'
            }`}
          >
            <span
              className={`w-5 h-5 rounded-full flex items-center justify-center text-[10px] font-bold ${
                currentPage === 'summaryFiller'
                  ? 'bg-emerald-600 text-white'
                  : 'bg-emerald-100 text-emerald-700 dark:bg-emerald-950 dark:text-emerald-300'
              }`}
            >
              S
            </span>
            Summary Topline
          </button>
        </div>

        {/* Right: Theme picker, Dark mode toggle, Debug button */}
        <div className="flex items-center gap-2 shrink-0">
          {/* Color Hunt Theme Picker Popover Toggle */}
          <div className="relative">
            <button
              type="button"
              onClick={() => setShowThemePicker(!showThemePicker)}
              title="Color Hunt Themes"
              className="p-2 rounded-xl bg-slate-100/80 dark:bg-slate-800/80 hover:bg-slate-200/80 dark:hover:bg-slate-700/80 text-slate-600 dark:text-slate-300 border border-slate-200/60 dark:border-slate-700/60 transition-colors flex items-center gap-1.5 text-xs font-medium shadow-sm"
            >
              <div className="flex -space-x-1">
                {palette.swatches.slice(0, 3).map((hex, i) => (
                  <span
                    key={i}
                    className="w-3 h-3 rounded-full border border-white dark:border-slate-900"
                    style={{ backgroundColor: hex }}
                  />
                ))}
              </div>
              <span className="hidden md:inline">Theme</span>
              <svg className="w-3.5 h-3.5 text-slate-400" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                <path strokeLinecap="round" strokeLinejoin="round" d="M19 9l-7 7-7-7" />
              </svg>
            </button>

            {/* Theme Picker Dropdown */}
            {showThemePicker && (
              <div
                className="absolute right-0 mt-2 w-64 p-3 bg-white dark:bg-slate-800 rounded-2xl border border-slate-200 dark:border-slate-700 shadow-xl z-50 transition-colors"
                role="dialog"
                aria-label="Theme picker"
                onMouseLeave={() => setShowThemePicker(false)}
              >
                <div className="flex items-center justify-between mb-2.5 pb-2 border-b border-slate-100 dark:border-slate-700">
                  <span className="text-xs font-bold text-slate-800 dark:text-white flex items-center gap-1.5">
                    <svg className="w-4 h-4 text-amber-500" fill="currentColor" viewBox="0 0 24 24">
                      <path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm1 17.93c-3.95-.49-7-3.85-7-7.93 0-.62.08-1.21.21-1.79L9 15v1c0 1.1.9 2 2 2v1.93zm6.9-2.54c-.26-.81-1-1.39-1.9-1.39h-1v-3c0-.55-.45-1-1-1H8v-2h2c.55 0 1-.45 1-1V7h2c1.1 0 2-.9 2-2v-.41c2.93 1.19 5 4.06 5 7.41 0 2.08-.8 3.97-2.1 5.4zm-7.9-13.3" />
                    </svg>
                    Color Hunt Palettes
                  </span>
                  <button
                    type="button"
                    onClick={() => setShowThemePicker(false)}
                    className="text-slate-400 hover:text-slate-600 dark:hover:text-slate-300 text-xs"
                    aria-label="Close theme picker"
                  >
                    ✕
                  </button>
                </div>
                <div className="space-y-1.5 max-h-60 overflow-y-auto">
                  {COLOR_PALETTES.map(p => (
                    <button
                      key={p.id}
                      type="button"
                      onClick={() => {
                        setPaletteId(p.id);
                        setShowThemePicker(false);
                      }}
                      className={`w-full flex items-center justify-between p-2 rounded-xl text-xs transition-colors ${
                        palette.id === p.id
                          ? 'bg-slate-100 dark:bg-slate-700 font-bold text-slate-900 dark:text-white'
                          : 'hover:bg-slate-50 dark:hover:bg-slate-700/50 text-slate-600 dark:text-slate-300'
                      }`}
                    >
                      <span>{p.name}</span>
                      <div className="flex -space-x-1">
                        {p.swatches.map((hex, i) => (
                          <span
                            key={i}
                            className="w-3.5 h-3.5 rounded-full border border-white dark:border-slate-800 shadow-2xs"
                            style={{ backgroundColor: hex }}
                          />
                        ))}
                      </div>
                    </button>
                  ))}
                </div>
              </div>
            )}
          </div>

          {/* Light / Dark Mode Toggle */}
          <button
            type="button"
            onClick={toggleMode}
            title={mode === 'light' ? 'Switch to Dark Mode' : 'Switch to Light Mode'}
            className="p-2 rounded-xl bg-slate-100/80 dark:bg-slate-800/80 hover:bg-slate-200/80 dark:hover:bg-slate-700/80 text-slate-600 dark:text-slate-300 border border-slate-200/60 dark:border-slate-700/60 transition-colors shadow-sm"
          >
            {mode === 'light' ? (
              <svg className="w-4 h-4 text-slate-700" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                <path strokeLinecap="round" strokeLinejoin="round" d="M20.354 15.354A9 9 0 018.646 3.646 9.003 9.003 0 0012 21a9.003 9.003 0 008.354-5.646z" />
              </svg>
            ) : (
              <svg className="w-4 h-4 text-amber-400" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                <path strokeLinecap="round" strokeLinejoin="round" d="M12 3v1m0 16v1m9-9h-1M4 12H3m15.364 6.364l-.707-.707M6.343 6.343l-.707-.707m12.728 0l-.707.707M6.343 17.657l-.707.707M16 12a6 6 0 11-12 0 6 6 0 0112 0z" />
              </svg>
            )}
          </button>

          {/* Debug Toggle */}
          {onToggleDebug && (
            <button
              type="button"
              onClick={onToggleDebug}
              title="Toggle Debug Panel"
              className="p-2 rounded-xl bg-slate-100/80 dark:bg-slate-800/80 hover:bg-slate-200/80 dark:hover:bg-slate-700/80 text-slate-600 dark:text-slate-300 border border-slate-200/60 dark:border-slate-700/60 transition-colors text-xs font-mono shadow-sm"
            >
              {showDebug ? '✕' : '⚙'}
            </button>
          )}
        </div>
      </div>

      {/* Mobile navigation bar */}
      <div className="sm:hidden border-t border-slate-200 dark:border-slate-800 bg-slate-50 dark:bg-slate-900 px-4 py-2 flex justify-around">
        <button
          type="button"
          onClick={() => onNavigate('processor')}
          className={`px-3 py-1 rounded-full text-xs font-semibold transition-colors ${
            currentPage === 'processor'
              ? 'bg-indigo-600 text-white shadow-sm'
              : 'text-slate-600 dark:text-slate-400'
          }`}
        >
          OE Entries
        </button>
        <button
          type="button"
          onClick={() => onNavigate('templateAppend')}
          className={`px-3 py-1 rounded-full text-xs font-semibold transition-colors ${
            currentPage === 'templateAppend'
              ? 'bg-teal-600 text-white shadow-sm'
              : 'text-slate-600 dark:text-slate-400'
          }`}
        >
          Stack Data
        </button>
        <button
          type="button"
          onClick={() => onNavigate('summaryFiller')}
          className={`px-3 py-1 rounded-full text-xs font-semibold transition-colors ${
            currentPage === 'summaryFiller'
              ? 'bg-emerald-600 text-white shadow-sm'
              : 'text-slate-600 dark:text-slate-400'
          }`}
        >
          Summary Topline
        </button>
      </div>
    </header>
  );
}
