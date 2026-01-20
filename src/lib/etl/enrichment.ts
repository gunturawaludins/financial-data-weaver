// Data Enrichment & Aggregation Module
// Handles VLOOKUP-like enrichment and group value aggregation

import { lookupEmiten, isMasterLoaded } from './masterData';

// Common column names for stock code and market value
const KODE_EFEK_COLUMNS = [
  'Kode Efek',
  'KODE EFEK',
  'Kode_Efek',
  'kode_efek',
  'Kode Saham',
  'KODE SAHAM',
  'Kode',
  'KODE',
  'Symbol',
  'Ticker',
];

const NILAI_PASAR_COLUMNS = [
  'Nilai Pasar Wajar',
  'NILAI PASAR WAJAR',
  'Nilai_Pasar_Wajar',
  'nilai_pasar_wajar',
  'Nilai Pasar',
  'NILAI PASAR',
  'Market Value',
  'Fair Value',
  'Nilai Wajar',
  'NPW',
];

/**
 * Find the column name that matches a pattern
 */
function findMatchingColumn(headers: string[], patterns: string[]): string | null {
  for (const header of headers) {
    for (const pattern of patterns) {
      if (header.toLowerCase().includes(pattern.toLowerCase())) {
        return header;
      }
    }
  }
  return null;
}

/**
 * Parse a numeric value from various formats
 * Handles: "Rp 1.000.000", "1,000,000", "1000000", etc.
 */
export function parseNumericValue(value: unknown): number {
  if (value === null || value === undefined) return 0;
  
  // If already a number, return it
  if (typeof value === 'number') {
    return isNaN(value) ? 0 : value;
  }
  
  // Convert to string and clean
  let strValue = String(value).trim();
  
  // Remove currency symbols and whitespace
  strValue = strValue.replace(/^[Rp\s$€¥£]+/gi, '');
  
  // Handle Indonesian format (1.234.567,89) vs US format (1,234,567.89)
  // Check if last separator is comma (Indonesian/European format)
  const lastComma = strValue.lastIndexOf(',');
  const lastDot = strValue.lastIndexOf('.');
  
  if (lastComma > lastDot && lastComma > strValue.length - 4) {
    // Indonesian/European format: 1.234.567,89
    strValue = strValue.replace(/\./g, '').replace(',', '.');
  } else {
    // US format or no decimals: 1,234,567.89 or 1.234.567
    strValue = strValue.replace(/,/g, '');
  }
  
  // Remove remaining non-numeric characters except decimal point and minus
  strValue = strValue.replace(/[^\d.\-]/g, '');
  
  const parsed = parseFloat(strValue);
  return isNaN(parsed) ? 0 : parsed;
}

/**
 * Enrich data with group information (VLOOKUP-like operation)
 */
export function enrichWithGroupData(
  data: Record<string, unknown>[],
  headers: string[]
): {
  enrichedData: Record<string, unknown>[];
  kodeEfekColumn: string | null;
  matchedCount: number;
  unmatchedCount: number;
} {
  if (!isMasterLoaded()) {
    console.warn('Master data not loaded. Skipping enrichment.');
    return {
      enrichedData: data,
      kodeEfekColumn: null,
      matchedCount: 0,
      unmatchedCount: data.length,
    };
  }
  
  // Find the stock code column
  const kodeEfekColumn = findMatchingColumn(headers, KODE_EFEK_COLUMNS);
  
  if (!kodeEfekColumn) {
    console.warn('No stock code column found. Skipping enrichment.');
    return {
      enrichedData: data,
      kodeEfekColumn: null,
      matchedCount: 0,
      unmatchedCount: data.length,
    };
  }
  
  let matchedCount = 0;
  let unmatchedCount = 0;
  
  const enrichedData = data.map((row) => {
    const kode = String(row[kodeEfekColumn] || '').trim().toUpperCase();
    const emiten = lookupEmiten(kode);
    
    if (emiten) {
      matchedCount++;
      return {
        ...row,
        GRUP_EMITEN: emiten.afiliasiUtama,
        NAMA_EMITEN_MASTER: emiten.namaEmiten,
        KATEGORI_EMITEN: emiten.kategori,
      };
    } else {
      unmatchedCount++;
      return {
        ...row,
        GRUP_EMITEN: 'Non-Grup',
        NAMA_EMITEN_MASTER: '',
        KATEGORI_EMITEN: '',
      };
    }
  });
  
  return {
    enrichedData,
    kodeEfekColumn,
    matchedCount,
    unmatchedCount,
  };
}

/**
 * Calculate group aggregates for market value
 * Each row gets the total value of all members in its group
 */
export function calculateGroupAggregates(
  data: Record<string, unknown>[],
  headers: string[]
): {
  aggregatedData: Record<string, unknown>[];
  nilaiPasarColumn: string | null;
  groupTotals: Map<string, number>;
} {
  // Find the market value column
  const nilaiPasarColumn = findMatchingColumn(headers, NILAI_PASAR_COLUMNS);
  
  if (!nilaiPasarColumn) {
    console.warn('No market value column found. Skipping aggregation.');
    return {
      aggregatedData: data,
      nilaiPasarColumn: null,
      groupTotals: new Map(),
    };
  }
  
  // Step 1: Calculate totals for each group
  const groupTotals = new Map<string, number>();
  
  for (const row of data) {
    const groupName = String(row['GRUP_EMITEN'] || 'Non-Grup');
    const nilaiPasar = parseNumericValue(row[nilaiPasarColumn]);
    
    const currentTotal = groupTotals.get(groupName) || 0;
    groupTotals.set(groupName, currentTotal + nilaiPasar);
  }
  
  // Step 2: Apply group totals to each row
  const aggregatedData = data.map((row) => {
    const groupName = String(row['GRUP_EMITEN'] || 'Non-Grup');
    const nilaiPasarAsli = parseNumericValue(row[nilaiPasarColumn]);
    
    // Get group total
    let grupNilaiPasarWajar: number;
    
    if (groupName === 'Non-Grup' || !groupName) {
      // Exception handling: Non-Grup uses its own value
      grupNilaiPasarWajar = nilaiPasarAsli;
    } else {
      grupNilaiPasarWajar = groupTotals.get(groupName) || nilaiPasarAsli;
    }
    
    return {
      ...row,
      NILAI_PASAR_WAJAR_CLEAN: nilaiPasarAsli,
      GRUP_NILAI_PASAR_WAJAR: grupNilaiPasarWajar,
    };
  });
  
  return {
    aggregatedData,
    nilaiPasarColumn,
    groupTotals,
  };
}

/**
 * Full enrichment pipeline: enrich + aggregate
 */
export function processEnrichmentPipeline(
  data: Record<string, unknown>[],
  headers: string[]
): {
  processedData: Record<string, unknown>[];
  newHeaders: string[];
  stats: {
    kodeEfekColumn: string | null;
    nilaiPasarColumn: string | null;
    matchedCount: number;
    unmatchedCount: number;
    groupCount: number;
    totalGroupValue: number;
  };
} {
  // Step 1: Enrich with group data
  const { enrichedData, kodeEfekColumn, matchedCount, unmatchedCount } = 
    enrichWithGroupData(data, headers);
  
  // Step 2: Calculate group aggregates
  const enrichedHeaders = [...headers, 'GRUP_EMITEN', 'NAMA_EMITEN_MASTER', 'KATEGORI_EMITEN'];
  const { aggregatedData, nilaiPasarColumn, groupTotals } = 
    calculateGroupAggregates(enrichedData, enrichedHeaders);
  
  // Calculate total group value
  let totalGroupValue = 0;
  groupTotals.forEach((value) => {
    totalGroupValue += value;
  });
  
  // Build final headers
  const newHeaders = [
    ...headers,
    'GRUP_EMITEN',
    'NAMA_EMITEN_MASTER', 
    'KATEGORI_EMITEN',
    'NILAI_PASAR_WAJAR_CLEAN',
    'GRUP_NILAI_PASAR_WAJAR',
  ];
  
  return {
    processedData: aggregatedData,
    newHeaders,
    stats: {
      kodeEfekColumn,
      nilaiPasarColumn,
      matchedCount,
      unmatchedCount,
      groupCount: groupTotals.size,
      totalGroupValue,
    },
  };
}
