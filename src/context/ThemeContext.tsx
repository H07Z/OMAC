import { createContext, useContext, useState, useEffect, ReactNode } from 'react';

export type ThemeMode = 'light' | 'dark';

export interface ColorPalette {
  id: string;
  name: string;
  swatches: [string, string, string, string];
  primaryHex: string;
  primaryHoverHex: string;
  bgHexLight: string;
  bgHexDark: string;
  accentClass: string;
  btnPrimaryClass: string;
  btnHoverClass: string;
  borderAccentClass: string;
  badgeBgClass: string;
}

export const COLOR_PALETTES: ColorPalette[] = [
  {
    id: 'indigo',
    name: 'Indigo Velvet',
    swatches: ['#4F46E5', '#818CF8', '#C7D2FE', '#1E1B4B'],
    primaryHex: '#4F46E5',
    primaryHoverHex: '#4338CA',
    bgHexLight: '#F8FAFC',
    bgHexDark: '#0F172A',
    accentClass: 'text-indigo-600 dark:text-indigo-400',
    btnPrimaryClass: 'bg-indigo-600 text-white dark:bg-indigo-500',
    btnHoverClass: 'hover:bg-indigo-700 dark:hover:bg-indigo-600',
    borderAccentClass: 'border-indigo-200 dark:border-indigo-800',
    badgeBgClass: 'bg-indigo-500 text-white',
  },
  {
    id: 'emerald',
    name: 'Emerald Mint',
    swatches: ['#059669', '#34D399', '#A7F3D0', '#064E3B'],
    primaryHex: '#059669',
    primaryHoverHex: '#047857',
    bgHexLight: '#F0FDF4',
    bgHexDark: '#022C22',
    accentClass: 'text-emerald-600 dark:text-emerald-400',
    btnPrimaryClass: 'bg-emerald-600 text-white dark:bg-emerald-500',
    btnHoverClass: 'hover:bg-emerald-700 dark:hover:bg-emerald-600',
    borderAccentClass: 'border-emerald-200 dark:border-emerald-800',
    badgeBgClass: 'bg-emerald-500 text-white',
  },
  {
    id: 'violet',
    name: 'Violet Dream',
    swatches: ['#7C3AED', '#A78BFA', '#DDD6FE', '#2E1065'],
    primaryHex: '#7C3AED',
    primaryHoverHex: '#6D28D9',
    bgHexLight: '#F5F3FF',
    bgHexDark: '#1E1B4B',
    accentClass: 'text-violet-600 dark:text-violet-400',
    btnPrimaryClass: 'bg-violet-600 text-white dark:bg-violet-500',
    btnHoverClass: 'hover:bg-violet-700 dark:hover:bg-violet-600',
    borderAccentClass: 'border-violet-200 dark:border-violet-800',
    badgeBgClass: 'bg-violet-500 text-white',
  },
  {
    id: 'rose',
    name: 'Rose Crimson',
    swatches: ['#E11D48', '#FB7185', '#FECDD3', '#4C0519'],
    primaryHex: '#E11D48',
    primaryHoverHex: '#BE123C',
    bgHexLight: '#FFF1F2',
    bgHexDark: '#1C0510',
    accentClass: 'text-rose-600 dark:text-rose-400',
    btnPrimaryClass: 'bg-rose-600 text-white dark:bg-rose-500',
    btnHoverClass: 'hover:bg-rose-700 dark:hover:bg-rose-600',
    borderAccentClass: 'border-rose-200 dark:border-rose-800',
    badgeBgClass: 'bg-rose-500 text-white',
  },
  {
    id: 'cyan',
    name: 'Ocean Cyan',
    swatches: ['#0284C7', '#38BDF8', '#BAE6FD', '#082F49'],
    primaryHex: '#0284C7',
    primaryHoverHex: '#0369A1',
    bgHexLight: '#F0F9FF',
    bgHexDark: '#0B192C',
    accentClass: 'text-sky-600 dark:text-sky-400',
    btnPrimaryClass: 'bg-sky-600 text-white dark:bg-sky-500',
    btnHoverClass: 'hover:bg-sky-700 dark:hover:bg-sky-600',
    borderAccentClass: 'border-sky-200 dark:border-sky-800',
    badgeBgClass: 'bg-sky-500 text-white',
  },
  {
    id: 'amber',
    name: 'Amber Autumn',
    swatches: ['#D97706', '#FBBF24', '#FEF3C7', '#451A03'],
    primaryHex: '#D97706',
    primaryHoverHex: '#B45309',
    bgHexLight: '#FFFBEB',
    bgHexDark: '#1F1202',
    accentClass: 'text-amber-600 dark:text-amber-400',
    btnPrimaryClass: 'bg-amber-600 text-white dark:bg-amber-500',
    btnHoverClass: 'hover:bg-amber-700 dark:hover:bg-amber-600',
    borderAccentClass: 'border-amber-200 dark:border-amber-800',
    badgeBgClass: 'bg-amber-500 text-white',
  },
];

interface ThemeContextType {
  mode: ThemeMode;
  toggleMode: () => void;
  palette: ColorPalette;
  setPaletteId: (id: string) => void;
}

const ThemeContext = createContext<ThemeContextType | undefined>(undefined);

export function ThemeProvider({ children }: { children: ReactNode }) {
  const getInitialMode = (): ThemeMode => {
    if (typeof window === 'undefined') return 'light';
    try {
      const saved = localStorage.getItem('ez_toolkit_theme_mode');
      if (saved === 'light' || saved === 'dark') return saved;
    } catch {}
    return window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
  };

  const getInitialPalette = (): string => {
    if (typeof window === 'undefined') return 'indigo';
    try {
      const saved = localStorage.getItem('ez_toolkit_palette');
      return COLOR_PALETTES.some(p => p.id === saved) ? (saved as string) : 'indigo';
    } catch {}
    return 'indigo';
  };

  const [mode, setMode] = useState<ThemeMode>(getInitialMode);
  const [paletteId, setPaletteIdState] = useState<string>(getInitialPalette);

  const palette = COLOR_PALETTES.find(p => p.id === paletteId) || COLOR_PALETTES[0];

  useEffect(() => {
    try { localStorage.setItem('ez_toolkit_theme_mode', mode); } catch {}
    const root = document.documentElement;
    if (mode === 'dark') root.classList.add('dark');
    else root.classList.remove('dark');
  }, [mode]);

  useEffect(() => {
    try { localStorage.setItem('ez_toolkit_palette', paletteId); } catch {}
    const root = document.documentElement;
    root.style.setProperty('--primary-color', palette.primaryHex);
    root.style.setProperty('--brand-500', palette.primaryHex);
    root.style.setProperty('--brand-600', palette.primaryHoverHex || palette.primaryHex);
    root.style.setProperty('--brand-700', palette.primaryHoverHex);
    root.setAttribute('data-theme', palette.id);
  }, [paletteId, palette]);

  const toggleMode = () => {
    setMode(prev => (prev === 'light' ? 'dark' : 'light'));
  };

  const setPaletteId = (id: string) => {
    if (COLOR_PALETTES.some(p => p.id === id)) setPaletteIdState(id);
  };

  return (
    <ThemeContext.Provider value={{ mode, toggleMode, palette, setPaletteId }}>
      {children}
    </ThemeContext.Provider>
  );
}

export function useTheme() {
  const context = useContext(ThemeContext);
  if (!context) {
    throw new Error('useTheme must be used within a ThemeProvider');
  }
  return context;
}
