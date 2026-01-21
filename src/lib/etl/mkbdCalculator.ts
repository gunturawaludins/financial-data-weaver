// MKBD Calculator - Multi-Pass Calculation Engine
// Handles VD59 and VD510 calculations with Ranking Liabilities and Modal Kerja Bersih

import { ProcessedSheet } from './types';
import { parseNumericValue } from './enrichment';

export interface MKBDCalculationResult {
  totalAsetLancar: number;
  totalLiabilitas: number;
  totalRankingLiabilities: number;
  modalKerja: number;
  modalKerjaBersih: number;
  mkbdDisesuaikan: number;
  mkbdDiwajibkan: number;
  lebihKurangMKBD: number;
  calculationSteps: CalculationStep[];
  vd510Details: VD510CalculationDetail[];
}

export interface CalculationStep {
  id: string;
  name: string;
  formula: string;
  inputValues: Record<string, number>;
  result: number;
  source: string;
  editable: boolean;
}

export interface VD510CalculationDetail {
  rowIndex: number;
  kodeEfek: string;
  namaEfek: string;
  nilaiPasarWajar: number;
  grupEmiten: string;
  persentaseTerhadapModal: number;
  batas20Persen: number;
  nilaiRankingLiabilities: number;
  formula: string;
}

export interface FormulaDefinition {
  id: string;
  name: string;
  description: string;
  formula: string;
  inputs: string[];
  calculate: (inputs: Record<string, number>) => number;
}

// Default formulas - can be replaced by user
export const DEFAULT_FORMULAS: Record<string, FormulaDefinition> = {
  modalKerja: {
    id: 'modalKerja',
    name: 'Modal Kerja',
    description: 'Total Aset Lancar - Total Liabilitas - Total Ranking Liabilities',
    formula: 'TOTAL_ASET_LANCAR - TOTAL_LIABILITAS - TOTAL_RANKING_LIABILITIES',
    inputs: ['TOTAL_ASET_LANCAR', 'TOTAL_LIABILITAS', 'TOTAL_RANKING_LIABILITIES'],
    calculate: (inputs) => 
      inputs.TOTAL_ASET_LANCAR - inputs.TOTAL_LIABILITAS - inputs.TOTAL_RANKING_LIABILITIES,
  },
  rankingLiabilitiesPerItem: {
    id: 'rankingLiabilitiesPerItem',
    name: 'Ranking Liabilities per Item',
    description: 'Nilai Grup - (20% x Total Aset Lancar), minimum 0',
    formula: 'MAX(0, NILAI_GRUP - (0.20 * TOTAL_MODAL_SENDIRI))',
    inputs: ['NILAI_GRUP', 'TOTAL_MODAL_SENDIRI'],
    calculate: (inputs) => 
      Math.max(0, inputs.NILAI_GRUP - (0.20 * inputs.TOTAL_MODAL_SENDIRI)),
  },
  mkbdDisesuaikan: {
    id: 'mkbdDisesuaikan',
    name: 'MKBD Disesuaikan',
    description: 'Modal Kerja Bersih dikurangi penyesuaian risiko',
    formula: 'MODAL_KERJA_BERSIH - TOTAL_PENYESUAIAN_RISIKO',
    inputs: ['MODAL_KERJA_BERSIH', 'TOTAL_PENYESUAIAN_RISIKO'],
    calculate: (inputs) => inputs.MODAL_KERJA_BERSIH - inputs.TOTAL_PENYESUAIAN_RISIKO,
  },
};

// Store custom formulas (in-memory, could be persisted)
let customFormulas: Record<string, FormulaDefinition> = { ...DEFAULT_FORMULAS };

export function getFormulas(): Record<string, FormulaDefinition> {
  return { ...customFormulas };
}

export function updateFormula(id: string, formula: Partial<FormulaDefinition>): void {
  if (customFormulas[id]) {
    customFormulas[id] = { ...customFormulas[id], ...formula };
  }
}

export function resetFormulas(): void {
  customFormulas = { ...DEFAULT_FORMULAS };
}

// Extract key values from VD51 sheet
export function extractVD51Values(sheet: ProcessedSheet): Record<string, number> {
  const values: Record<string, number> = {};
  
  // Search patterns for VD51 key rows
  const patterns = {
    TOTAL_ASET_LANCAR: /total\s*aset\s*lancar/i,
    KAS_SETARA_KAS: /kas\s*(dan|&)?\s*setara\s*kas/i,
    DEPOSITO_BANK: /deposito\s*bank/i,
    PIUTANG_NASABAH: /piutang\s*nasabah/i,
    PORTOFOLIO_EFEK: /portofolio\s*efek/i,
  };

  for (const row of sheet.data) {
    const rowText = Object.values(row)
      .filter(v => typeof v === 'string')
      .join(' ')
      .toLowerCase();

    for (const [key, pattern] of Object.entries(patterns)) {
      if (pattern.test(rowText) && !values[key]) {
        // Find numeric value in the row
        const numericValue = findNumericValueInRow(row);
        if (numericValue !== null) {
          values[key] = numericValue;
        }
      }
    }
  }

  return values;
}

// Extract key values from VD52 sheet (Liabilitas)
export function extractVD52Values(sheet: ProcessedSheet): Record<string, number> {
  const values: Record<string, number> = {};
  
  const patterns = {
    TOTAL_LIABILITAS: /total\s*liabilitas/i,
    UTANG_SUB_ORDINASI: /utang\s*sub[\-\s]?ordinasi/i,
    UTANG_JANGKA_PENDEK: /utang\s*jangka\s*pendek/i,
  };

  for (const row of sheet.data) {
    const rowText = Object.values(row)
      .filter(v => typeof v === 'string')
      .join(' ')
      .toLowerCase();

    for (const [key, pattern] of Object.entries(patterns)) {
      if (pattern.test(rowText) && !values[key]) {
        const numericValue = findNumericValueInRow(row);
        if (numericValue !== null) {
          values[key] = numericValue;
        }
      }
    }
  }

  return values;
}

// Extract VD53 values (Ranking Liabilities summary)
export function extractVD53Values(sheet: ProcessedSheet): Record<string, number> {
  const values: Record<string, number> = {};
  
  const patterns = {
    TOTAL_RANKING_LIABILITIES: /total\s*ranking\s*liabilit/i,
  };

  for (const row of sheet.data) {
    const rowText = Object.values(row)
      .filter(v => typeof v === 'string')
      .join(' ')
      .toLowerCase();

    for (const [key, pattern] of Object.entries(patterns)) {
      if (pattern.test(rowText) && !values[key]) {
        const numericValue = findNumericValueInRow(row);
        if (numericValue !== null) {
          values[key] = numericValue;
        }
      }
    }
  }

  return values;
}

// Process VD510 and calculate Ranking Liabilities for each item
export function calculateVD510RankingLiabilities(
  sheet: ProcessedSheet,
  totalModalSendiri: number
): VD510CalculationDetail[] {
  const details: VD510CalculationDetail[] = [];
  const batas20Persen = totalModalSendiri * 0.20;

  // Find relevant columns
  const kodeEfekCol = findColumnByPattern(sheet.headers, /kode\s*efek/i);
  const nilaiPasarCol = findColumnByPattern(sheet.headers, /nilai\s*pasar\s*wajar/i);
  const grupEmitenCol = findColumnByPattern(sheet.headers, /grup\s*emiten/i);
  const rankingLiabCol = findColumnByPattern(sheet.headers, /nilai\s*rank/i);
  
  // Group values by emiten/grup for aggregation
  const grupAggregates: Record<string, number> = {};
  
  // First pass: aggregate by grup
  for (const row of sheet.data) {
    const kodeEfek = String(row[kodeEfekCol] ?? '');
    const grupEmiten = String(row[grupEmitenCol] ?? kodeEfek);
    const nilaiPasar = parseNumericValue(row[nilaiPasarCol]);
    
    if (nilaiPasar > 0 && kodeEfek && !kodeEfek.toLowerCase().includes('other')) {
      const grupKey = grupEmiten || kodeEfek;
      grupAggregates[grupKey] = (grupAggregates[grupKey] || 0) + nilaiPasar;
    }
  }

  // Second pass: calculate ranking liabilities
  let rowIndex = 0;
  for (const row of sheet.data) {
    rowIndex++;
    const kodeEfek = String(row[kodeEfekCol] ?? '');
    const namaEfek = String(row['Jenis_Efek'] ?? row['Nama_Akun'] ?? kodeEfek);
    const nilaiPasarWajar = parseNumericValue(row[nilaiPasarCol]);
    const grupEmiten = String(row[grupEmitenCol] ?? kodeEfek);
    
    if (nilaiPasarWajar > 0 && kodeEfek && !kodeEfek.toLowerCase().includes('other')) {
      const grupKey = grupEmiten || kodeEfek;
      const nilaiGrup = grupAggregates[grupKey] || nilaiPasarWajar;
      
      // Calculate percentage of total modal
      const persentase = (nilaiGrup / totalModalSendiri) * 100;
      
      // Calculate ranking liabilities (excess over 20%)
      const excessValue = nilaiGrup - batas20Persen;
      const nilaiRankingLiabilities = Math.max(0, excessValue);
      
      // Only add if there's an actual excess
      const formula = `MAX(0, ${formatNumber(nilaiGrup)} - (20% × ${formatNumber(totalModalSendiri)}))`;
      
      details.push({
        rowIndex,
        kodeEfek,
        namaEfek,
        nilaiPasarWajar,
        grupEmiten: grupKey,
        persentaseTerhadapModal: persentase,
        batas20Persen,
        nilaiRankingLiabilities,
        formula,
      });
    }
  }

  return details;
}

// Main calculation function - Multi-Pass
export function calculateMKBD(sheets: ProcessedSheet[]): MKBDCalculationResult {
  const calculationSteps: CalculationStep[] = [];
  
  // Find relevant sheets
  const vd51Sheet = sheets.find(s => /vd5[\-_]?1\b|formulir[\-_\s]*1\b/i.test(s.sheetName));
  const vd52Sheet = sheets.find(s => /vd5[\-_]?2\b|formulir[\-_\s]*2\b/i.test(s.sheetName));
  const vd53Sheet = sheets.find(s => /vd5[\-_]?3\b|formulir[\-_\s]*3\b/i.test(s.sheetName));
  const vd59Sheet = sheets.find(s => /vd5[\-_]?9\b|formulir[\-_\s]*9\b/i.test(s.sheetName));
  const vd510Sheet = sheets.find(s => /vd5[\-_]?10\b|formulir[\-_\s]*10\b/i.test(s.sheetName));

  // === PASS 1: Extract Base Values ===
  let totalAsetLancar = 0;
  let totalLiabilitas = 0;
  let utangSubOrdinasi = 0;

  if (vd51Sheet) {
    const vd51Values = extractVD51Values(vd51Sheet);
    totalAsetLancar = vd51Values.TOTAL_ASET_LANCAR || 0;
    
    calculationSteps.push({
      id: 'pass1_aset_lancar',
      name: 'Total Aset Lancar (VD51)',
      formula: 'Extract from VD51 Baris 100',
      inputValues: { value: totalAsetLancar },
      result: totalAsetLancar,
      source: 'VD51',
      editable: false,
    });
  }

  if (vd52Sheet) {
    const vd52Values = extractVD52Values(vd52Sheet);
    totalLiabilitas = vd52Values.TOTAL_LIABILITAS || 0;
    utangSubOrdinasi = vd52Values.UTANG_SUB_ORDINASI || 0;
    
    calculationSteps.push({
      id: 'pass1_liabilitas',
      name: 'Total Liabilitas (VD52)',
      formula: 'Extract from VD52 Baris 164',
      inputValues: { value: totalLiabilitas },
      result: totalLiabilitas,
      source: 'VD52',
      editable: false,
    });
  }

  // === PASS 2: Calculate VD510 Ranking Liabilities ===
  let vd510Details: VD510CalculationDetail[] = [];
  let totalRankingLiabilities = 0;

  // Get Total Modal Sendiri (Ekuitas) - approximation
  const totalModalSendiri = totalAsetLancar - totalLiabilitas;

  if (vd510Sheet) {
    vd510Details = calculateVD510RankingLiabilities(vd510Sheet, totalModalSendiri);
    
    // Sum all ranking liabilities
    totalRankingLiabilities = vd510Details.reduce(
      (sum, item) => sum + item.nilaiRankingLiabilities, 
      0
    );

    calculationSteps.push({
      id: 'pass2_ranking_liabilities',
      name: 'Total Ranking Liabilities (VD510)',
      formula: 'SUM(MAX(0, Nilai_Grup - 20% × Total_Modal_Sendiri))',
      inputValues: { 
        totalModalSendiri,
        batas20Persen: totalModalSendiri * 0.20,
        itemCount: vd510Details.length,
      },
      result: totalRankingLiabilities,
      source: 'VD510',
      editable: true,
    });
  }

  // === PASS 3: Calculate MKBD ===
  const modalKerja = totalAsetLancar - totalLiabilitas - totalRankingLiabilities;
  const modalKerjaBersih = modalKerja + utangSubOrdinasi;
  
  // For simplicity, assume minimal risk adjustments (can be expanded)
  const penyesuaianRisiko = 0; // Would need VD59 haircut calculations
  const mkbdDisesuaikan = modalKerjaBersih - penyesuaianRisiko;
  
  // Minimum MKBD requirement
  const mkbdDiwajibkan = Math.max(25_000_000_000, totalLiabilitas * 0.0625);
  const lebihKurangMKBD = mkbdDisesuaikan - mkbdDiwajibkan;

  calculationSteps.push({
    id: 'pass3_modal_kerja',
    name: 'Modal Kerja',
    formula: 'TOTAL_ASET_LANCAR - TOTAL_LIABILITAS - TOTAL_RANKING_LIABILITIES',
    inputValues: { totalAsetLancar, totalLiabilitas, totalRankingLiabilities },
    result: modalKerja,
    source: 'Calculated',
    editable: true,
  });

  calculationSteps.push({
    id: 'pass3_modal_kerja_bersih',
    name: 'Modal Kerja Bersih',
    formula: 'MODAL_KERJA + UTANG_SUB_ORDINASI',
    inputValues: { modalKerja, utangSubOrdinasi },
    result: modalKerjaBersih,
    source: 'Calculated',
    editable: true,
  });

  calculationSteps.push({
    id: 'pass3_mkbd_disesuaikan',
    name: 'MKBD Disesuaikan',
    formula: 'MODAL_KERJA_BERSIH - PENYESUAIAN_RISIKO',
    inputValues: { modalKerjaBersih, penyesuaianRisiko },
    result: mkbdDisesuaikan,
    source: 'Calculated',
    editable: true,
  });

  calculationSteps.push({
    id: 'pass3_lebih_kurang',
    name: 'Lebih/(Kurang) MKBD',
    formula: 'MKBD_DISESUAIKAN - MKBD_DIWAJIBKAN',
    inputValues: { mkbdDisesuaikan, mkbdDiwajibkan },
    result: lebihKurangMKBD,
    source: 'Calculated',
    editable: false,
  });

  return {
    totalAsetLancar,
    totalLiabilitas,
    totalRankingLiabilities,
    modalKerja,
    modalKerjaBersih,
    mkbdDisesuaikan,
    mkbdDiwajibkan,
    lebihKurangMKBD,
    calculationSteps,
    vd510Details,
  };
}

// Helper functions
function findNumericValueInRow(row: Record<string, unknown>): number | null {
  const values = Object.values(row);
  
  for (let i = values.length - 1; i >= 0; i--) {
    const val = values[i];
    if (val !== null && val !== undefined) {
      const num = parseNumericValue(val);
      if (num !== 0 || String(val).includes('0')) {
        return num;
      }
    }
  }
  
  return null;
}

function findColumnByPattern(headers: string[], pattern: RegExp): string {
  return headers.find(h => pattern.test(h)) || headers[0];
}

function formatNumber(num: number): string {
  if (num >= 1e12) return `${(num / 1e12).toFixed(2)}T`;
  if (num >= 1e9) return `${(num / 1e9).toFixed(2)}B`;
  if (num >= 1e6) return `${(num / 1e6).toFixed(2)}M`;
  return num.toLocaleString('id-ID');
}
