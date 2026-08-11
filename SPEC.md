# Excel VBA Migration - Web Application Specification

## 1. Concept & Vision

A professional-grade web application that replaces Excel/VBA workflows for data processing. The application provides a clean, modern interface for uploading Excel files, processing data according to VBA-derived business rules, and exporting clean results. It feels like a dedicated desktop application embedded in a browser—fast, reliable, and trustworthy for mission-critical data work.

## 2. Design Language

### Aesthetic Direction
Industrial-modern with a focus on data clarity. Inspired by Notion's clean utility and Linear's precision. Every element serves a purpose; no decorative excess.

### Color Palette
- **Primary**: `#4F46E5` (Indigo 600) - Actions, buttons, active states
- **Primary Hover**: `#4338CA` (Indigo 700)
- **Secondary**: `#0F172A` (Slate 900) - Headers, important text
- **Accent Success**: `#10B981` (Emerald 500) - Success states, completed operations
- **Accent Warning**: `#F59E0B` (Amber 500) - Warnings, removed items
- **Accent Error**: `#EF4444` (Red 500) - Errors, destructive actions
- **Background**: `#F8FAFC` (Slate 50) - Page background
- **Surface**: `#FFFFFF` - Cards, panels
- **Border**: `#E2E8F0` (Slate 200) - Dividers, borders
- **Text Primary**: `#1E293B` (Slate 800)
- **Text Secondary**: `#64748B` (Slate 500)
- **Text Muted**: `#94A3B8` (Slate 400)

### Typography
- **Headings**: Inter (700 weight) - Clean, professional
- **Body**: Inter (400, 500 weight)
- **Monospace/Data**: JetBrains Mono - For numbers, IDs, code-like content
- **Scale**: 12px (small), 14px (body), 16px (large), 20px (h3), 24px (h2), 32px (h1)

### Spatial System
- Base unit: 4px
- Spacing scale: 4, 8, 12, 16, 20, 24, 32, 40, 48, 64px
- Card padding: 24px
- Section gaps: 32px
- Border radius: 8px (cards), 6px (buttons), 4px (inputs)

### Motion Philosophy
- **Transitions**: 150ms ease-out for micro-interactions, 300ms for state changes
- **Processing**: Subtle pulse animation during data processing
- **Success feedback**: Brief scale + color transition on completion
- **No gratuitous animation**: Motion serves function, not decoration

### Visual Assets
- **Icons**: Lucide React - consistent 24px stroke icons
- **File types**: Custom SVG icons for Excel, CSV, unknown files
- **Status indicators**: Colored dots with labels

## 3. Layout & Structure

### Page Architecture
```
┌─────────────────────────────────────────────────────────┐
│  HEADER: Logo + Title                                   │
├─────────────────────────────────────────────────────────┤
│  PROCESSING LOG (collapsible) - Shows operation history │
├──────────────────────┬──────────────────────────────────┤
│                      │                                  │
│  LEFT PANEL:         │  MAIN CONTENT:                   │
│  - Upload Zone       │  - File Info (after upload)      │
│  - File Info         │  - Processing Stats             │
│  - Sheet Selector    │  - Data Table                    │
│  - Process Button    │  - Search/Filter/Sort Controls   │
│  - Processing Stats   │  - Pagination                   │
│  - Export Button      │                                 │
│  - Reset Button       │                                 │
│                      │                                  │
├──────────────────────┴──────────────────────────────────┤
│  FOOTER: Processing notes, debug info (toggle)          │
└─────────────────────────────────────────────────────────┘
```

### Responsive Strategy
- Desktop (>1024px): Side-by-side layout as shown
- Tablet (768-1024px): Stacked layout, full-width panels
- Mobile (<768px): Single column, simplified table view

## 4. Features & Interactions

### 4.1 File Upload
- **Drag & drop zone**: Visual highlight on dragover (border color change, background tint)
- **Click to browse**: Opens native file picker
- **Validation**: Only accepts .xlsx, .xls, .csv
- **Invalid file**: Red border flash, error message displayed
- **Success**: Green checkmark animation, immediate file info display

### 4.2 Worksheet Selection
- **Dropdown**: Lists all worksheets in the workbook
- **Default**: First worksheet selected automatically
- **Single sheet**: Dropdown hidden if only one sheet

### 4.3 Data Import
- **Progress**: Brief loading indicator for large files
- **Headers**: First row treated as column headers
- **Empty cells**: Represented as empty strings, not undefined

### 4.4 Data Processing Pipeline
```
originalData → cleanData → removeExcludedRows → sortData → renumberData → processedData
```

#### Step 4.4.1: Clean Data
- Trim whitespace from all cell values
- Convert numbers stored as strings to native numbers where safe
- Preserve empty cells as empty strings

#### Step 4.4.2: Remove Excluded Rows
- **Rule**: Remove entire row when Column B (index 1) matches any of:
  - `(NET)`
  - `(SUBNET)`
  - `(SUB-SUBNET)`
  - `(SUB-SUB-SUBNET)`
- **Comparison**: Case-insensitive, exact match, trimmed
- **Applies**: Even when Column A is blank

#### Step 4.4.3: Sort Data
- **Column**: Sort by Column A (index 0) numerically
- **Order**: Ascending (1, 2, 3... not 1, 10, 2)
- **Tie-breaking**: Maintain original order for equal values

#### Step 4.4.4: Renumber Data
- **Start**: Always at 1
- **Increment**: Sequential integers (1, 2, 3...)
- **Column**: Write to Column A
- **Overwrite**: Any existing numbers in Column A

### 4.5 Table Display
- **Columns**: Dynamic based on imported data
- **Headers**: Clickable for sorting (visual indicator)
- **Rows**: Virtual scrolling for performance (>100 rows)
- **Selection**: Row highlighting on hover
- **Empty state**: "No data to display" with icon

### 4.6 Search
- **Input**: Real-time filtering as user types
- **Scope**: Searches all visible columns
- **Debounce**: 150ms delay to prevent excessive filtering
- **Clear**: X button to clear search

### 4.7 Filtering
- **Column filters**: Dropdown per column with unique values
- **Multi-select**: Can select multiple values per column
- **Active filters**: Displayed as pills with remove option

### 4.8 Sorting (UI)
- **Click header**: Toggle asc → desc → none
- **Visual**: Arrow icon indicating direction
- **Multi-sort**: Shift+click for secondary sort (optional enhancement)

### 4.9 Export
- **Button**: "Export to Excel"
- **Output**: .xlsx file with processed data
- **Location**: Data starts at A10 (rows 1-9 reserved)
- **Headers**: Included in export at row 10
- **Filename**: Original name + "_processed" suffix

### 4.10 Reset
- **Action**: Clears processed data, returns to original import state
- **Preserves**: Uploaded file reference (no re-upload needed)
- **Confirmation**: None required (fast iteration)

## 5. Component Inventory

### Upload Zone
- **Default**: Dashed border, icon, "Drop Excel file here" text
- **Hover**: Border darkens, subtle background tint
- **Dragover**: Border solid indigo, stronger background tint
- **Error**: Red border, error message below
- **Success**: Green border briefly, then transitions to file info

### File Info Card
- **Content**: Filename, sheet count, row count, column count
- **Icon**: File type indicator (Excel green / CSV blue)
- **Actions**: Change file button (small, secondary)

### Sheet Selector
- **Type**: Native select dropdown with custom styling
- **Label**: "Worksheet" above
- **Hidden**: When only one sheet available

### Process Button
- **Default**: Large, full-width in panel, indigo background
- **Disabled**: Grayed out before file upload
- **Loading**: Spinner icon, "Processing..." text
- **Success**: Brief green flash, then returns to default

### Statistics Cards
- **Layout**: 3-column grid (Original / Removed / Final)
- **Original**: Total imported rows
- **Removed**: Rows deleted by filter rules
- **Final**: Rows in processed dataset
- **Animation**: Numbers count up on first display

### Data Table
- **Header**: Sticky, gray background, sort indicators
- **Rows**: Alternating subtle gray/white (optional)
- **Cells**: Left-aligned text, right-aligned numbers
- **Overflow**: Horizontal scroll with shadow indicators
- **Loading**: Skeleton rows during initial render
- **Empty**: Centered message with upload prompt

### Search Input
- **Icon**: Magnifying glass left side
- **Placeholder**: "Search data..."
- **Clear**: X button appears when text present
- **Focus**: Indigo ring

### Filter Pills
- **Style**: Rounded, colored background, X to remove
- **Column label**: Small text above group
- **Clear all**: Link below filters

### Export Button
- **Style**: Secondary button, download icon
- **Disabled**: When no processed data
- **Success**: Brief checkmark, "Downloaded!" text

### Reset Button
- **Style**: Ghost button, subtle
- **Position**: Below export button
- **Hover**: Light background

### Processing Log
- **Collapsible**: Click header to toggle
- **Entries**: Timestamped, color-coded by type
- **Types**: Info (gray), Success (green), Warning (amber), Error (red)
- **Max height**: Scrollable when many entries

### Debug Panel
- **Toggle**: Hidden by default, enable in settings
- **Content**: Raw data inspection, processing times
- **Style**: Monospace font, dark background

## 6. Technical Approach

### Architecture
```
/src
  /components
    UploadZone.tsx
    FileInfo.tsx
    SheetSelector.tsx
    ProcessButton.tsx
    StatsCards.tsx
    DataTable.tsx
    SearchBar.tsx
    FilterControls.tsx
    ExportButton.tsx
    ResetButton.tsx
    ProcessingLog.tsx
    DebugPanel.tsx
  /modules
    excelLoader.ts      - SheetJS integration for reading
    csvLoader.ts        - CSV parsing
    dataProcessor.ts    - Main processing pipeline
    cleaner.ts          - Data cleaning functions
    rowRemover.ts       - Excluded row removal
    sorter.ts           - Numeric sorting
    renumberer.ts       - Sequential numbering
    exporter.ts         - SheetJS Excel export
    validator.ts        - Input validation
  /types
    index.ts            - TypeScript interfaces
  /utils
    debug.ts            - Debug logging utility
  App.tsx
  main.tsx
  index.css
```

### State Management
```typescript
interface AppState {
  // File state
  workbook: XLSX.WorkBook | null;
  selectedSheet: string;
  fileName: string;
  
  // Data state
  originalData: RowData[];
  processedData: RowData[];
  filteredData: RowData[];
  
  // UI state
  isProcessing: boolean;
  searchTerm: string;
  filters: Record<string, string[]>;
  sortConfig: { column: string; direction: 'asc' | 'desc' | null };
  
  // Stats
  stats: {
    originalRows: number;
    removedRows: number;
    finalRows: number;
  };
  
  // Log
  processingLog: LogEntry[];
  
  // Debug
  debugMode: boolean;
  debugLogs: string[];
}
```

### Key Dependencies
- **react**: UI framework
- **xlsx** (SheetJS): Excel reading and writing
- **typescript**: Type safety
- **tailwindcss**: Styling
- **lucide-react**: Icons

### Data Model
```typescript
interface RowData {
  [columnName: string]: string | number | null;
}

interface ColumnDef {
  key: string;
  label: string;
  type: 'string' | 'number';
}
```

### VBA Function Mapping
| VBA Concept | JavaScript Implementation |
|-------------|---------------------------|
| Worksheet | `workbook.SheetNames[index]` |
| Range/Cells | `worksheet[row][col]` |
| LastRow | `data.length` |
| For Each Loop | `for...of` iterator |
| If Condition | `if` statement |
| Delete Row | `Array.splice()` or filter |
| Sort | `Array.sort()` with numeric comparator |
| Find | `Array.find()` |
| CountIf | `Array.filter().length` |
| Trim() | `String.prototype.trim()` |
| UCase() | `String.prototype.toUpperCase()` |

### Excluded Row Patterns
```typescript
const EXCLUDED_PATTERNS = [
  '(NET)',
  '(SUBNET)',
  '(SUB-SUBNET)',
  '(SUB-SUB-SUBNET)'
];

function isExcludedRow(row: RowData): boolean {
  const colB = String(row['Column B'] || '').trim().toUpperCase();
  return EXCLUDED_PATTERNS.includes(colB);
}
```

### Export Format
- Output starts at row 10 (A10)
- Rows 1-9: Reserved for future headers/metadata
- First row of data (row 10): Column headers
- Subsequent rows: Data records

## 7. Test Cases

### Test 1: Remove Hierarchy Markers
**Input:**
```
Column A | Column B
--------|--------
        | (NET)
123     | Normal data
        | (SUBNET)
456     | Normal data
```
**Expected Output:**
```
1 | Apple
2 | Banana
3 | Orange
```

### Test 2: Numeric Sorting
**Input:**
```
Column A
--------
5
9
20
1
15
```
**Expected Output:**
```
1
2
3
4
5
```

### Test 3: Renumbering After Processing
**Input:**
```
Column A
--------
100
200
300
```
**Expected Output:**
```
1
2
3
```

### Test 4: Blank Column A with Exclusion
**Input:**
```
Column A | Column B
--------|--------
        | (NET)
        | (SUBNET)
        | Apple
```
**Expected Output:**
```
1 | Apple
```

### Test 5: Case Insensitive Matching
**Input:**
```
Column A | Column B
--------|--------
        | (net)
        | (Subnet)
        | (sub-sub-net)
        | Valid Data
```
**Expected Output:**
```
1 | Valid Data
```
