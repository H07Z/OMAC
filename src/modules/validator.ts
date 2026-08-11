/**
 * Validator Module
 * 
 * Validates imported data and file structures.
 * Ensures data integrity and provides helpful error messages.
 */

import { RowData } from '../types';
import { debugLog } from '../utils/debug';

export interface ValidationResult {
  isValid: boolean;
  errors: ValidationError[];
  warnings: ValidationWarning[];
}

export interface ValidationError {
  field: string;
  message: string;
}

export interface ValidationWarning {
  field: string;
  message: string;
}

/**
 * Validate imported worksheet data
 */
export function validateData(
  data: RowData[],
  headers: string[]
): ValidationResult {
  const errors: ValidationError[] = [];
  const warnings: ValidationWarning[] = [];
  
  debugLog('Validator', `Validating ${data.length} rows with ${headers.length} columns`);
  
  // Check for empty data
  if (data.length === 0) {
    errors.push({
      field: 'data',
      message: 'No data found in the worksheet'
    });
    return { isValid: false, errors, warnings };
  }
  
  // Check for missing headers
  if (headers.length === 0) {
    errors.push({
      field: 'headers',
      message: 'No column headers found'
    });
  }
  
  // Check for Column A
  const columnA = headers.find(h => 
    h.toLowerCase() === 'column a' || 
    h.toLowerCase() === 'columna' ||
    h === 'A' ||
    h === 'A1'
  );
  
  if (!columnA) {
    warnings.push({
      field: 'Column A',
      message: 'Column A not found. Sorting may not work correctly.'
    });
  }
  
  // Check for Column B
  const columnB = headers.find(h => 
    h.toLowerCase() === 'column b' ||
    h.toLowerCase() === 'columnb' ||
    h === 'B' ||
    h === 'B1'
  );
  
  if (!columnB) {
    warnings.push({
      field: 'Column B',
      message: 'Column B not found. Hierarchy marker removal may not work.'
    });
  }
  
  // Check for potential issues with Column A values
  if (columnA) {
    const nonNumericCount = data.filter(row => {
      const value = row[columnA];
      if (value === null || value === undefined) return false;
      return isNaN(Number(value));
    }).length;
    
    if (nonNumericCount > data.length * 0.5) {
      warnings.push({
        field: 'Column A',
        message: 'Most Column A values are non-numeric. Numeric sorting will be applied to valid numbers.'
      });
    }
  }
  
  // Check for completely empty rows
  const emptyRowCount = data.filter(row => 
    Object.values(row).every(v => v === null || v === undefined || String(v).trim() === '')
  ).length;
  
  if (emptyRowCount > 0) {
    warnings.push({
      field: 'data',
      message: `${emptyRowCount} empty row(s) found in the data`
    });
  }
  
  // Check for duplicate headers
  const headerCounts: Record<string, number> = {};
  for (const header of headers) {
    headerCounts[header] = (headerCounts[header] || 0) + 1;
  }
  
  const duplicates = Object.entries(headerCounts)
    .filter(([, count]) => count > 1)
    .map(([header]) => header);
  
  if (duplicates.length > 0) {
    warnings.push({
      field: 'headers',
      message: `Duplicate column headers found: ${duplicates.join(', ')}`
    });
  }
  
  const isValid = errors.length === 0;
  
  debugLog('Validator', `Validation complete: ${isValid ? 'VALID' : 'INVALID'}`);
  if (warnings.length > 0) {
    debugLog('Validator', `Warnings: ${warnings.length}`);
  }
  
  return { isValid, errors, warnings };
}

/**
 * Validate file before import
 */
export function validateFile(file: File): ValidationResult {
  const errors: ValidationError[] = [];
  const warnings: ValidationWarning[] = [];
  
  // Check file exists
  if (!file) {
    errors.push({
      field: 'file',
      message: 'No file selected'
    });
    return { isValid: false, errors, warnings };
  }
  
  // Check file size (max 50MB)
  const maxSize = 50 * 1024 * 1024;
  if (file.size > maxSize) {
    errors.push({
      field: 'file',
      message: `File too large (${formatBytes(file.size)}). Maximum size is ${formatBytes(maxSize)}.`
    });
  }
  
  // Check file extension
  const validExtensions = ['.xlsx', '.xls', '.csv'];
  const fileName = file.name.toLowerCase();
  const hasValidExtension = validExtensions.some(ext => fileName.endsWith(ext));
  
  if (!hasValidExtension) {
    errors.push({
      field: 'file',
      message: `Invalid file type. Supported formats: ${validExtensions.join(', ')}`
    });
  }
  
  const isValid = errors.length === 0;
  return { isValid, errors, warnings };
}

/**
 * Format bytes to human readable string
 */
function formatBytes(bytes: number): string {
  if (bytes === 0) return '0 Bytes';
  
  const k = 1024;
  const sizes = ['Bytes', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}
