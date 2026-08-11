/**
 * Filter and Search Module
 * 
 * Handles filtering and searching of processed data.
 * Does not modify the underlying data, only creates filtered views.
 */

import { RowData, FilterConfig } from '../types';
import { debugLog } from '../utils/debug';

/**
 * Filter data based on search term (searches all columns)
 */
export function searchData(
  data: RowData[],
  searchTerm: string,
  headers: string[]
): RowData[] {
  if (!searchTerm || searchTerm.trim() === '') {
    return data;
  }
  
  const term = searchTerm.toLowerCase().trim();
  
  debugLog('FilterSearch', `Searching "${term}" in ${data.length} rows`);
  
  const filtered = data.filter(row => {
    // Search all columns
    for (const header of headers) {
      const value = row[header];
      if (value !== null && value !== undefined) {
        const strValue = String(value).toLowerCase();
        if (strValue.includes(term)) {
          return true;
        }
      }
    }
    return false;
  });
  
  debugLog('FilterSearch', `Search found ${filtered.length} matching rows`);
  
  return filtered;
}

/**
 * Filter data based on column-specific filters
 */
export function filterData(
  data: RowData[],
  filters: FilterConfig
): RowData[] {
  const activeFilters = Object.entries(filters).filter(
    ([, values]) => values && values.length > 0
  );
  
  if (activeFilters.length === 0) {
    return data;
  }
  
  debugLog('FilterSearch', `Applying ${activeFilters.length} filter(s)`);
  
  const filtered = data.filter(row => {
    // All filters must match (AND logic)
    return activeFilters.every(([column, values]) => {
      // No values selected means no filtering for this column
      if (!values || values.length === 0) {
        return true;
      }
      
      const rowValue = row[column];
      const rowValueStr = rowValue !== null && rowValue !== undefined 
        ? String(rowValue) 
        : '';
      
      // Check if row value matches any of the selected filter values
      return values.includes(rowValueStr);
    });
  });
  
  debugLog('FilterSearch', `Filters reduced to ${filtered.length} rows`);
  
  return filtered;
}

/**
 * Apply both search and filters
 */
export function applyFilters(
  data: RowData[],
  searchTerm: string,
  filters: FilterConfig,
  headers: string[]
): RowData[] {
  let result = data;
  
  // First apply column filters
  result = filterData(result, filters);
  
  // Then apply search
  result = searchData(result, searchTerm, headers);
  
  return result;
}

/**
 * Get unique values for a column (for filter dropdowns)
 */
export function getUniqueValues(
  data: RowData[],
  column: string
): string[] {
  const values = new Set<string>();
  
  for (const row of data) {
    const value = row[column];
    if (value !== null && value !== undefined) {
      values.add(String(value));
    }
  }
  
  return Array.from(values).sort((a, b) => {
    // Try numeric sort for numeric-looking values
    const numA = Number(a);
    const numB = Number(b);
    
    if (!isNaN(numA) && !isNaN(numB)) {
      return numA - numB;
    }
    
    return a.localeCompare(b);
  });
}

/**
 * Get active filter count
 */
export function getActiveFilterCount(filters: FilterConfig): number {
  return Object.values(filters).reduce((count, values) => {
    return count + (values && values.length > 0 ? values.length : 0);
  }, 0);
}

/**
 * Clear all filters
 */
export function clearFilters(filters: FilterConfig): FilterConfig {
  const cleared: FilterConfig = {};
  
  for (const key of Object.keys(filters)) {
    cleared[key] = [];
  }
  
  return cleared;
}
