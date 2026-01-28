/**
 * Excel Exporter - Generate Excel file dengan format asli tapi angka hasil re-kalkulasi
 * 
 * File ini adalah fungsi mandiri untuk export hasil audit MKBD ke Excel.
 * Format Excel mengikuti template asli, hanya angka yang di-update sesuai kalkulasi web.
 */

import * as XLSX from 'xlsx';
import { ProcessedSheet, MKBDCalculationResult, VD59Update } from './types';

export interface ExportOptions {
  fileName?: string;
  includeMetadata?: boolean;
}

/**
 * Generate Excel file dari data ProcessedSheet yang sudah di-koreksi
 * Format mengikuti struktur asli, angka menggunakan hasil kalkulasi web
 */
export function generateCorrectedExcel(
  sheets: ProcessedSheet[],
  mkbdResult: MKBDCalculationResult | null,
  options: ExportOptions = {}
): Blob {
  const workbook = XLSX.utils.book_new();

  for (const sheet of sheets) {
    // Convert data array ke worksheet
    const wsData: unknown[][] = [];

    // Add header row
    wsData.push(sheet.headers);

    // Add data rows
    for (const row of sheet.data) {
      const rowData: unknown[] = [];
      for (const header of sheet.headers) {
        let value = row[header];
        
        // Format numeric values
        if (typeof value === 'number') {
          rowData.push(value);
        } else if (value === null || value === undefined) {
          rowData.push('');
        } else {
          rowData.push(value);
        }
      }
      wsData.push(rowData);
    }

    const worksheet = XLSX.utils.aoa_to_sheet(wsData);

    // Apply column widths
    const colWidths = sheet.headers.map(header => ({
      wch: Math.max(header.length, 15)
    }));
    worksheet['!cols'] = colWidths;

    // Add sheet to workbook
    XLSX.utils.book_append_sheet(workbook, worksheet, sheet.sheetName.substring(0, 31));
  }

  // Add summary sheet if MKBD result exists
  if (mkbdResult) {
    const summarySheet = createSummarySheet(mkbdResult);
    XLSX.utils.book_append_sheet(workbook, summarySheet, 'MKBD_Summary');
  }

  // Generate binary
  const wbout = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
  
  return new Blob([wbout], { 
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
  });
}

/**
 * Create summary sheet dengan hasil kalkulasi MKBD
 */
function createSummarySheet(mkbdResult: MKBDCalculationResult): XLSX.WorkSheet {
  const data: unknown[][] = [
    ['RINGKASAN HASIL AUDIT MKBD'],
    [''],
    ['Tanggal Proses', new Date().toLocaleDateString('id-ID')],
    [''],
    ['== SUMBER DATA =='],
    ['Total Ekuitas (VD52)', mkbdResult.totalEkuitas],
    ['Total Aset Lancar (VD59)', mkbdResult.totalAsetLancar],
    ['Total Liabilitas (VD59)', mkbdResult.totalLiabilitas],
    [''],
    ['== HASIL KALKULASI VD510 =='],
    ['Total Ranking Liabilities', mkbdResult.totalRankingLiabilities],
    [''],
    ['== HASIL KALKULASI VD59 =='],
  ];

  // Add VD59 updates
  if (mkbdResult.vd59Updates) {
    for (const update of mkbdResult.vd59Updates) {
      data.push([update.rowLabel, update.newValue]);
    }
  }

  data.push(['']);
  data.push(['== FORMULA YANG DIGUNAKAN ==']);
  data.push(['Nilai_Rangking_Liabilities', 'GRUP_NILAI_PASAR_WAJAR - (20% × TOTAL EKUITAS)']);
  data.push(['Total Modal Kerja', 'TOTAL ASET LANCAR - TOTAL LIABILITAS - TOTAL RANKING LIABILITIES']);
  data.push(['MKBD Disesuaikan', 'TOTAL MODAL KERJA BERSIH (Baris 18) - SUM(Baris 33-92)']);
  data.push(['Lebih/Kurang MKBD', 'MKBD Disesuaikan - NILAI MKBD YANG DIWAJIBKAN']);

  const worksheet = XLSX.utils.aoa_to_sheet(data);
  
  // Set column widths
  worksheet['!cols'] = [
    { wch: 50 },
    { wch: 25 }
  ];

  return worksheet;
}

/**
 * Download Excel file ke browser
 */
export function downloadExcel(
  sheets: ProcessedSheet[],
  mkbdResult: MKBDCalculationResult | null,
  fileName?: string
): void {
  const blob = generateCorrectedExcel(sheets, mkbdResult);
  
  const defaultFileName = `MKBD_Corrected_${new Date().toISOString().split('T')[0]}.xlsx`;
  const finalFileName = fileName || defaultFileName;

  // Create download link
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = finalFileName;
  
  // Trigger download
  document.body.appendChild(link);
  link.click();
  
  // Cleanup
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

/**
 * Generate Excel dengan format yang persis sama dengan template asli
 * Ini membaca workbook asli dan hanya mengganti nilai yang di-koreksi
 */
export async function generateFromTemplate(
  originalFile: File,
  correctedSheets: ProcessedSheet[],
  mkbdResult: MKBDCalculationResult | null
): Promise<Blob> {
  // Read original workbook
  const arrayBuffer = await originalFile.arrayBuffer();
  const workbook = XLSX.read(arrayBuffer, { type: 'array' });

  // Create map of corrected data by sheet name
  const correctedMap = new Map<string, ProcessedSheet>();
  for (const sheet of correctedSheets) {
    // Match by original sheet name pattern
    correctedMap.set(sheet.sheetName, sheet);
    // Also try without suffix
    const baseName = sheet.sheetName.replace(/_TABEL_10C$/, '');
    correctedMap.set(baseName, sheet);
  }

  // Update each sheet with corrected values
  for (const sheetName of workbook.SheetNames) {
    const worksheet = workbook.Sheets[sheetName];
    const correctedSheet = findMatchingSheet(sheetName, correctedSheets);

    if (correctedSheet) {
      applyCorrectionsToWorksheet(worksheet, correctedSheet, mkbdResult);
    }
  }

  // Generate output
  const wbout = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
  
  return new Blob([wbout], { 
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
  });
}

/**
 * Find matching ProcessedSheet for a given sheet name
 */
function findMatchingSheet(
  sheetName: string, 
  sheets: ProcessedSheet[]
): ProcessedSheet | null {
  const normalizedName = sheetName.toLowerCase().replace(/[\s\-_]/g, '');
  
  for (const sheet of sheets) {
    const normalizedSheetName = sheet.sheetName.toLowerCase().replace(/[\s\-_]/g, '');
    if (normalizedSheetName.includes(normalizedName) || normalizedName.includes(normalizedSheetName)) {
      return sheet;
    }
  }
  
  return null;
}

/**
 * Apply corrections to worksheet cells
 */
function applyCorrectionsToWorksheet(
  worksheet: XLSX.WorkSheet,
  correctedSheet: ProcessedSheet,
  mkbdResult: MKBDCalculationResult | null
): void {
  // Get worksheet range
  const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
  
  // Find header row in worksheet
  let headerRowIdx = -1;
  let headerColMap: Map<string, number> = new Map();

  for (let r = range.s.r; r <= Math.min(range.e.r, 20); r++) {
    const rowHeaders: string[] = [];
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cellAddr = XLSX.utils.encode_cell({ r, c });
      const cell = worksheet[cellAddr];
      if (cell && cell.v) {
        rowHeaders.push(String(cell.v));
      }
    }
    
    // Check if this row matches our headers
    const matchCount = correctedSheet.headers.filter(h => 
      rowHeaders.some(rh => normalizeHeader(rh) === normalizeHeader(h))
    ).length;
    
    if (matchCount >= 3) {
      headerRowIdx = r;
      // Build column map
      for (let c = range.s.c; c <= range.e.c; c++) {
        const cellAddr = XLSX.utils.encode_cell({ r, c });
        const cell = worksheet[cellAddr];
        if (cell && cell.v) {
          headerColMap.set(normalizeHeader(String(cell.v)), c);
        }
      }
      break;
    }
  }

  if (headerRowIdx === -1) return;

  // Apply corrected values
  for (let i = 0; i < correctedSheet.data.length; i++) {
    const row = correctedSheet.data[i];
    const excelRowIdx = headerRowIdx + 1 + i;

    for (const [header, value] of Object.entries(row)) {
      const colIdx = headerColMap.get(normalizeHeader(header));
      if (colIdx !== undefined && value !== null && value !== undefined) {
        const cellAddr = XLSX.utils.encode_cell({ r: excelRowIdx, c: colIdx });
        
        if (typeof value === 'number') {
          worksheet[cellAddr] = { t: 'n', v: value };
        } else if (typeof value === 'string' && !isNaN(parseFloat(value.replace(/,/g, '')))) {
          worksheet[cellAddr] = { t: 'n', v: parseFloat(value.replace(/,/g, '')) };
        } else {
          worksheet[cellAddr] = { t: 's', v: String(value) };
        }
      }
    }
  }

  // Apply specific VD59 updates if this is VD59 sheet
  if (mkbdResult?.vd59Updates && correctedSheet.sheetName.toLowerCase().includes('vd59')) {
    applyVD59SpecificUpdates(worksheet, mkbdResult.vd59Updates);
  }
}

/**
 * Apply VD59 specific row updates by matching row labels
 */
function applyVD59SpecificUpdates(
  worksheet: XLSX.WorkSheet,
  updates: VD59Update[]
): void {
  const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
  
  for (const update of updates) {
    const rowLabel = update.rowLabel || update.rowDescription || '';
    
    // Find row by label
    for (let r = range.s.r; r <= range.e.r; r++) {
      let rowText = '';
      let targetCol = -1;
      
      for (let c = range.s.c; c <= range.e.c; c++) {
        const cellAddr = XLSX.utils.encode_cell({ r, c });
        const cell = worksheet[cellAddr];
        if (cell && cell.v) {
          const cellValue = String(cell.v);
          rowText += ' ' + cellValue;
          
          // Find target column (Jumlah or Total)
          if (normalizeHeader(cellValue).includes(normalizeHeader(update.column))) {
            targetCol = c;
          }
        }
      }

      // Check if row label matches
      if (rowText.toUpperCase().includes(rowLabel.toUpperCase().substring(0, 30))) {
        // Find the numeric column (usually the last few columns)
        if (targetCol === -1) {
          // Default to a numeric column
          for (let c = range.e.c; c >= range.s.c; c--) {
            const cellAddr = XLSX.utils.encode_cell({ r, c });
            const cell = worksheet[cellAddr];
            if (cell && (cell.t === 'n' || (cell.v && !isNaN(parseFloat(String(cell.v).replace(/,/g, '')))))) {
              targetCol = c;
              break;
            }
          }
        }

        if (targetCol !== -1) {
          const cellAddr = XLSX.utils.encode_cell({ r, c: targetCol });
          worksheet[cellAddr] = { t: 'n', v: update.newValue };
        }
        break;
      }
    }
  }
}

/**
 * Normalize header string for comparison
 */
function normalizeHeader(header: string): string {
  return header
    .toLowerCase()
    .replace(/[\s\-_]/g, '')
    .replace(/[^a-z0-9]/g, '');
}

/**
 * Download Excel dari template asli dengan nilai terkoreksi
 */
export async function downloadFromTemplate(
  originalFile: File,
  correctedSheets: ProcessedSheet[],
  mkbdResult: MKBDCalculationResult | null,
  fileName?: string
): Promise<void> {
  const blob = await generateFromTemplate(originalFile, correctedSheets, mkbdResult);
  
  const baseName = originalFile.name.replace(/\.xlsx?$/i, '');
  const defaultFileName = `${baseName}_CORRECTED.xlsx`;
  const finalFileName = fileName || defaultFileName;

  // Create download link
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = finalFileName;
  
  // Trigger download
  document.body.appendChild(link);
  link.click();
  
  // Cleanup
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}
