/**
 * Excel Loader Module
 * 
 * Handles reading Excel files (.xlsx, .xls) and CSV files using SheetJS.
 * Converts worksheet data into JavaScript RowData arrays.
 */

import * as XLSX from 'xlsx';
import { RowData } from '../types';
import { debugLog } from '../utils/debug';

export interface LoadResult {
  workbook: XLSX.WorkBook;
  sheetNames: string[];
}

export interface SheetData {
  headers: string[];
  data: RowData[];
}

/**
 * Read a File object as an Excel workbook
 */
export function readWorkbook(file: File): Promise<LoadResult> {
  return new Promise((resolve, reject) => {
    debugLog('ExcelLoader', `Reading file: ${file.name} (${file.size} bytes)`);
    
    const reader = new FileReader();
    
    reader.onload = (e) => {
      try {
        const target = e.target;
        if (!target || !target.result) {
          throw new Error('No data received from file reader');
        }
        
        const data = new Uint8Array(target.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array', cellDates: true });
        
        debugLog('ExcelLoader', `Workbook loaded: ${workbook.SheetNames.length} sheet(s)`);
        
        resolve({
          workbook,
          sheetNames: workbook.SheetNames
        });
      } catch (error) {
        debugLog('ExcelLoader', `Error reading workbook: ${error}`);
        reject(new Error(`Failed to read Excel file: ${error}`));
      }
    };
    
    reader.onerror = () => {
      reject(new Error('File reader error'));
    };
    
    reader.readAsArrayBuffer(file);
  });
}

/**
 * Read a CSV file
 */
export function readCSV(file: File): Promise<LoadResult> {
  return new Promise((resolve, reject) => {
    debugLog('CSVLoader', `Reading CSV file: ${file.name}`);
    
    const reader = new FileReader();
    
    reader.onload = (e) => {
      try {
        const target = e.target;
        if (!target || !target.result) {
          throw new Error('No data received from file reader');
        }
        
        const csvText = target.result as string;
        const workbook = XLSX.read(csvText, { type: 'string', cellDates: true });
        
        debugLog('CSVLoader', `CSV loaded: ${workbook.SheetNames.length} sheet(s)`);
        
        resolve({
          workbook,
          sheetNames: workbook.SheetNames
        });
      } catch (error) {
        debugLog('CSVLoader', `Error reading CSV: ${error}`);
        reject(new Error(`Failed to read CSV file: ${error}`));
      }
    };
    
    reader.onerror = () => {
      reject(new Error('File reader error'));
    };
    
    reader.readAsText(file);
  });
}

/**
 * Load a specific worksheet from a workbook
 */
export function loadSheet(workbook: XLSX.WorkBook, sheetName: string): SheetData {
  debugLog('ExcelLoader', `Loading sheet: "${sheetName}"`);
  
  const worksheet = workbook.Sheets[sheetName];
  
  if (!worksheet) {
    throw new Error(`Worksheet "${sheetName}" not found`);
  }
  
  // Convert to JSON with header row
  const jsonData = XLSX.utils.sheet_to_json<RowData>(worksheet, { 
    defval: null,  // Default value for empty cells
    raw: false     // Convert to strings where possible
  });
  
  // Extract headers from the first row of data
  const headers = jsonData.length > 0 ? Object.keys(jsonData[0]) : [];
  
  debugLog('ExcelLoader', `Sheet loaded: ${jsonData.length} rows, ${headers.length} columns`);
  
  if (headers.length === 0) {
    debugLog('ExcelLoader', 'Warning: No headers found in worksheet');
  } else {
    debugLog('ExcelLoader', `Headers: ${headers.join(', ')}`);
  }
  
  return {
    headers,
    data: jsonData
  };
}

/**
 * Validate file type
 */
export function isValidFileType(file: File): boolean {
  const validExtensions = ['.xlsx', '.xls', '.csv'];
  const fileName = file.name.toLowerCase();
  
  return validExtensions.some(ext => fileName.endsWith(ext));
}

/**
 * Get file type from extension
 */
export function getFileType(fileName: string): 'xlsx' | 'xls' | 'csv' | null {
  const lower = fileName.toLowerCase();
  
  if (lower.endsWith('.xlsx')) return 'xlsx';
  if (lower.endsWith('.xls')) return 'xls';
  if (lower.endsWith('.csv')) return 'csv';
  
  return null;
}

/**
 * Get Excel file icon based on type
 */
export function getFileIcon(fileType: 'xlsx' | 'xls' | 'csv' | null): string {
  switch (fileType) {
    case 'xlsx':
      return 'excel';
    case 'xls':
      return 'excel';
    case 'csv':
      return 'csv';
    default:
      return 'file';
  }
}
