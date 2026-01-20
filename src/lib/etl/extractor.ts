// Excel Extractor - Handles different extraction methods for various sheet types
import * as XLSX from 'xlsx';
import { SheetData, ProcessedSheet, ETLResult } from './types';
import { sanitizeColumnNames, sanitizeTableName } from './sanitizer';
import { cleanData, findDataStartRow } from './cleaner';
import { processEnrichmentPipeline } from './enrichment';
import { isMasterLoaded } from './masterData';
// Sheet patterns that need special handling
const VD510_SHEET_PATTERN = /vd510|vd5[\-_]?10|formulir\s*10/i;
const STANDARD_SHEET_PATTERN = /vd5\d{1,2}|formulir\s*\d+/i;

// Table markers for VD510 special extraction
const TABLE_10C_START = 'TABEL 10C';
const TABLE_10C_STOP_MARKERS = ['TABEL 10D', 'TABEL 10E', 'Apabila diperlukan'];

export async function extractFromExcel(file: File): Promise<ETLResult> {
  const result: ETLResult = {
    success: true,
    sheets: [],
    errors: [],
    warnings: [],
  };

  try {
    const arrayBuffer = await file.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: 'array' });

    for (const sheetName of workbook.SheetNames) {
      try {
        const worksheet = workbook.Sheets[sheetName];
        
        let sheetData: SheetData;
        
        // Determine extraction method based on sheet name
        if (VD510_SHEET_PATTERN.test(sheetName)) {
          sheetData = extractVD510Special(worksheet, sheetName);
          result.warnings.push(`Sheet "${sheetName}" diproses dengan metode khusus (Tabel 10C)`);
        } else {
          sheetData = extractStandard(worksheet, sheetName);
        }

        if (sheetData.rows.length === 0) {
          result.warnings.push(`Sheet "${sheetName}" tidak memiliki data setelah pembersihan`);
          continue;
        }

        // Process and clean the data
        const processedSheet = processSheetData(sheetData, file.name);
        result.sheets.push(processedSheet);
        
      } catch (sheetError) {
        result.warnings.push(`Gagal memproses sheet "${sheetName}": ${sheetError}`);
      }
    }

    if (result.sheets.length === 0) {
      result.success = false;
      result.errors.push('Tidak ada sheet yang berhasil diproses');
    }

  } catch (error) {
    result.success = false;
    result.errors.push(`Gagal membaca file Excel: ${error}`);
  }

  return result;
}

function extractStandard(worksheet: XLSX.WorkSheet, sheetName: string): SheetData {
  // Convert to JSON with all options
  const rawData = XLSX.utils.sheet_to_json<unknown[]>(worksheet, {
    header: 1,
    defval: null,
    blankrows: true,
  });

  if (rawData.length === 0) {
    return { sheetName, headers: [], rows: [], originalHeaders: [], rowCount: 0 };
  }

  // Find where data actually starts (skip letterhead)
  const dataStartRow = findDataStartRow(rawData);
  
  // The header row is typically at dataStartRow, data starts at dataStartRow + 1
  const headerRowIndex = Math.min(dataStartRow, rawData.length - 1);
  const headerRow = rawData[headerRowIndex] as (string | number | null)[];
  
  // Sanitize headers
  const { sanitizedHeaders, warnings } = sanitizeColumnNames(headerRow);
  
  // Extract data rows
  const dataRows = rawData.slice(headerRowIndex + 1);
  
  // Convert to objects
  const rows = dataRows.map((row) => {
    const rowArray = row as unknown[];
    const obj: Record<string, unknown> = {};
    sanitizedHeaders.forEach((header, idx) => {
      obj[header] = rowArray[idx] ?? null;
    });
    return obj;
  });

  return {
    sheetName,
    headers: sanitizedHeaders,
    originalHeaders: headerRow.map(h => String(h ?? '')),
    rows,
    rowCount: rows.length,
  };
}

function extractVD510Special(worksheet: XLSX.WorkSheet, sheetName: string): SheetData {
  // Convert entire sheet to array
  const rawData = XLSX.utils.sheet_to_json<unknown[]>(worksheet, {
    header: 1,
    defval: null,
    blankrows: true,
  });

  // Find TABEL 10C start
  let table10CStart = -1;
  let table10CEnd = rawData.length;

  for (let i = 0; i < rawData.length; i++) {
    const rowText = (rawData[i] as unknown[])
      .filter(c => c !== null && c !== undefined)
      .map(c => String(c))
      .join(' ');

    // Find start marker
    if (table10CStart === -1 && rowText.toUpperCase().includes(TABLE_10C_START)) {
      table10CStart = i;
      continue;
    }

    // Find end marker (after we found start)
    if (table10CStart !== -1) {
      for (const stopMarker of TABLE_10C_STOP_MARKERS) {
        if (rowText.toUpperCase().includes(stopMarker.toUpperCase())) {
          table10CEnd = i;
          break;
        }
      }
      if (table10CEnd !== rawData.length) break;
    }
  }

  if (table10CStart === -1) {
    // Fallback to standard extraction if TABEL 10C not found
    return extractStandard(worksheet, sheetName);
  }

  // Extract only TABEL 10C data
  // Header is typically the row after the title
  const headerRowIndex = table10CStart + 1;
  const headerRow = rawData[headerRowIndex] as (string | number | null)[];
  
  // Sanitize headers
  const { sanitizedHeaders } = sanitizeColumnNames(headerRow);
  
  // Extract data rows (from after header to before end marker)
  const dataRows = rawData.slice(headerRowIndex + 1, table10CEnd);
  
  // Convert to objects
  const rows = dataRows.map((row) => {
    const rowArray = row as unknown[];
    const obj: Record<string, unknown> = {};
    sanitizedHeaders.forEach((header, idx) => {
      obj[header] = rowArray[idx] ?? null;
    });
    return obj;
  });

  return {
    sheetName: `${sheetName}_TABEL_10C`,
    headers: sanitizedHeaders,
    originalHeaders: headerRow.map(h => String(h ?? '')),
    rows,
    rowCount: rows.length,
  };
}

function processSheetData(sheetData: SheetData, fileName: string): ProcessedSheet {
  // Clean the data
  const { cleanedData, removedColumns } = cleanData(sheetData.rows, sheetData.headers);
  
  // Filter out removed columns from headers
  let finalHeaders = sheetData.headers.filter(h => !removedColumns.includes(h));
  let processedData = cleanedData;
  
  // Apply enrichment if master data is loaded
  let enrichmentStats = null;
  if (isMasterLoaded()) {
    const enrichmentResult = processEnrichmentPipeline(cleanedData, finalHeaders);
    processedData = enrichmentResult.processedData;
    finalHeaders = enrichmentResult.newHeaders;
    enrichmentStats = enrichmentResult.stats;
  }

  // Add metadata columns
  const uploadDate = new Date().toISOString();
  const dataWithMetadata = processedData.map((row) => ({
    ...row,
    _fileName: fileName,
    _uploadDate: uploadDate,
  }));

  return {
    sheetName: sheetData.sheetName,
    tableName: sanitizeTableName(sheetData.sheetName),
    headers: [...finalHeaders, '_fileName', '_uploadDate'],
    data: dataWithMetadata,
    metadata: {
      fileName,
      uploadDate,
      originalRowCount: sheetData.rowCount,
      cleanedRowCount: cleanedData.length,
      enrichmentStats,
    },
  };
}

export function getSheetNames(file: File): Promise<string[]> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        resolve(workbook.SheetNames);
      } catch (error) {
        reject(error);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}
