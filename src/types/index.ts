/**
 * Type definitions for Excel VBA Migration Application
 */

// Represents a single row of data (column name -> value)
export interface RowData {
  [columnName: string]: string | number | null;
}

// Column definition for table rendering
export interface ColumnDef {
  key: string;
  label: string;
  type: 'string' | 'number';
}

// Log entry for processing history
export interface LogEntry {
  timestamp: Date;
  level: 'info' | 'success' | 'warning' | 'error';
  message: string;
}

// Sort configuration
export interface SortConfig {
  column: string;
  direction: 'asc' | 'desc' | null;
}

// Filter configuration
export interface FilterConfig {
  [column: string]: string[];
}

// Processing statistics
export interface ProcessingStats {
  originalRows: number;
  removedRows: number;
  insertedRows: number;
  finalRows: number;
}

// Application state
export interface AppState {
  // File state
  workbook: XLSX.WorkBook | null;
  selectedSheet: string;
  fileName: string;
  fileType: 'xlsx' | 'xls' | 'csv' | null;
  
  // Data state
  originalData: RowData[];
  processedData: RowData[];
  filteredData: RowData[];
  headers: string[];
  
  // UI state
  isProcessing: boolean;
  isLoading: boolean;
  searchTerm: string;
  filters: FilterConfig;
  sortConfig: SortConfig;
  
  // Stats
  stats: ProcessingStats;
  
  // Log
  processingLog: LogEntry[];
  
  // Debug
  debugMode: boolean;
}

// SheetJS type augmentation
declare namespace XLSX {
  interface WorkBook {
    SheetNames: string[];
    Sheets: { [sheet: string]: WorkSheet };
  }
  
  interface WorkSheet {
    [cell: string]: CellObject | Range | undefined;
  }
  
  interface CellObject {
    t: string;
    v: string | number | Date;
    w?: string;
    f?: string;
  }
  
  interface Range {
    s: { c: number; r: number };
    e: { c: number; r: number };
  }
}
