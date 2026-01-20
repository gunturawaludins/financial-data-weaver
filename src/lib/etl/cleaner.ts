// Data Cleaner - Removes junk rows, empty columns, and footer text

export interface CleaningResult {
  cleanedData: Record<string, unknown>[];
  removedRows: number;
  removedColumns: string[];
  warnings: string[];
}

// Keywords that indicate junk rows (footer, notes, etc.)
const JUNK_KEYWORDS = [
  'apabila diperlukan',
  'baris baru dapat ditambahkan',
  'catatan:',
  'note:',
  '*)',
  '**)',
  '***)',
  '****)',
  '*****)',
  'peringatan:',
  'keterangan:',
  'halaman',
  'page',
  'dicetak pada',
  'printed on',
];

// Keywords that indicate header/title rows to skip
const HEADER_KEYWORDS = [
  'perusahaan efek',
  'tanggal',
  'direktur',
  'formulir',
];

export function cleanData(
  data: Record<string, unknown>[],
  headers: string[]
): CleaningResult {
  const warnings: string[] = [];
  let removedRows = 0;

  // Step 1: Remove mostly empty columns
  const { filteredHeaders, columnsToRemove } = removeEmptyColumns(data, headers);
  
  // Step 2: Remove junk and empty rows
  const cleanedData = data.filter((row, index) => {
    // Check if row is mostly empty
    const nonEmptyValues = Object.values(row).filter(
      (v) => v !== null && v !== undefined && String(v).trim() !== ''
    );
    
    if (nonEmptyValues.length === 0) {
      removedRows++;
      return false;
    }

    // Check if row is less than 10% filled (likely a junk row)
    const fillRatio = nonEmptyValues.length / filteredHeaders.length;
    if (fillRatio < 0.1 && filteredHeaders.length > 5) {
      removedRows++;
      return false;
    }

    // Check for junk keywords
    const rowText = Object.values(row)
      .filter((v) => typeof v === 'string')
      .join(' ')
      .toLowerCase();

    for (const keyword of JUNK_KEYWORDS) {
      if (rowText.includes(keyword.toLowerCase())) {
        removedRows++;
        warnings.push(`Baris ${index + 1} dihapus (mengandung: "${keyword}")`);
        return false;
      }
    }

    return true;
  });

  // Step 3: Remove columns that were identified as empty
  const finalData = cleanedData.map((row) => {
    const newRow: Record<string, unknown> = {};
    for (const header of filteredHeaders) {
      if (!columnsToRemove.includes(header)) {
        newRow[header] = row[header];
      }
    }
    return newRow;
  });

  return {
    cleanedData: finalData,
    removedRows,
    removedColumns: columnsToRemove,
    warnings,
  };
}

function removeEmptyColumns(
  data: Record<string, unknown>[],
  headers: string[]
): { filteredHeaders: string[]; columnsToRemove: string[] } {
  const columnsToRemove: string[] = [];
  const filteredHeaders: string[] = [];

  for (const header of headers) {
    // Skip columns with "Unnamed" prefix if they're mostly empty
    const isUnnamed = header.startsWith('Unnamed_');
    
    // Count non-empty values in this column
    const nonEmptyCount = data.filter((row) => {
      const value = row[header];
      return value !== null && value !== undefined && String(value).trim() !== '';
    }).length;

    const fillRatio = data.length > 0 ? nonEmptyCount / data.length : 0;

    // Remove columns that are:
    // 1. Unnamed AND less than 20% filled, OR
    // 2. Less than 5% filled for any column
    if ((isUnnamed && fillRatio < 0.2) || fillRatio < 0.05) {
      columnsToRemove.push(header);
    } else {
      filteredHeaders.push(header);
    }
  }

  return { filteredHeaders, columnsToRemove };
}

export function isHeaderRow(row: unknown[]): boolean {
  const rowText = row
    .filter((v) => typeof v === 'string')
    .join(' ')
    .toLowerCase();

  for (const keyword of HEADER_KEYWORDS) {
    if (rowText.includes(keyword)) {
      return true;
    }
  }

  return false;
}

export function findDataStartRow(rows: unknown[][], maxRowsToCheck = 15): number {
  // Skip letterhead/header rows and find where actual data starts
  for (let i = 0; i < Math.min(rows.length, maxRowsToCheck); i++) {
    const row = rows[i];
    
    // Check if this row contains typical header patterns (A, B, C... or numbers like 6, 7, 8...)
    const firstFewCells = row.slice(0, 5).map((c) => String(c ?? '').trim());
    
    // Look for column indicator rows (typically row with A, B, C or single digits)
    const hasColumnIndicators = firstFewCells.some(
      (c) => /^[A-Za-z]$/.test(c) || /^[0-9]$/.test(c)
    );
    
    if (hasColumnIndicators && i > 3) {
      // The actual header is likely the next row
      return i + 1;
    }
  }

  // Default: skip first 6 rows (letterhead area)
  return 6;
}
