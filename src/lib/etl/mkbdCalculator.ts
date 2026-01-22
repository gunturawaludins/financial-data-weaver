// MKBD Calculator - Multi-Pass Calculation Engine
// Handles VD59 and VD510 calculations with Ranking Liabilities and Modal Kerja Bersih

import { ProcessedSheet } from './types';
import { parseNumericValue } from './enrichment';

export interface MKBDCalculationResult {
  totalAsetLancar: number;
  totalEkuitas: number;
  totalLiabilitas: number;
  totalRankingLiabilities: number;
  modalKerja: number;
  modalKerjaBersih: number;
  mkbdDisesuaikan: number;
  mkbdDiwajibkan: number;
  lebihKurangMKBD: number;
  calculationSteps: CalculationStep[];
  vd510Details: VD510CalculationDetail[];
  vd59Updates: VD59Update[];
  haircutSum: number;
}

export interface VD59Update {
  rowIndex: number;
  rowDescription: string;
  column: string;
  oldValue: number;
  newValue: number;
  formula: string;
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
    description: 'Nilai Grup - (20% x Total Ekuitas)',
    formula: 'NILAI_GRUP - (0.20 * TOTAL_EKUITAS)',
    inputs: ['NILAI_GRUP', 'TOTAL_EKUITAS'],
    calculate: (inputs) => inputs.NILAI_GRUP - (0.20 * inputs.TOTAL_EKUITAS),
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

// Extract TOTAL LIABILITAS from VD52, row with "TOTAL LIABILITAS" and column "Saldo"
export function extractVD52TotalLiabilitas(sheet: ProcessedSheet): number {
  const saldoCol = findColumnOptional(sheet.headers, /saldo/i);

  for (const row of sheet.data) {
    const rowText = Object.values(row)
      .filter(v => typeof v === 'string')
      .join(' ');

    if (/total\s*liabilitas/i.test(rowText)) {
      if (saldoCol && row[saldoCol] !== null && row[saldoCol] !== undefined) {
        return parseNumericValue(row[saldoCol]);
      }
      const numericValue = findNumericValueInRow(row);
      if (numericValue !== null) return numericValue;
    }
  }

  return 0;
}

// Extract TOTAL EKUITAS from VD52, row with "TOTAL EKUITAS" and column "Saldo"
export function extractVD52TotalEkuitas(sheet: ProcessedSheet): number {
  const saldoCol = findColumnOptional(sheet.headers, /saldo/i);

  for (const row of sheet.data) {
    const rowText = Object.values(row)
      .filter(v => typeof v === 'string')
      .join(' ');

    if (/total\s*ekuitas/i.test(rowText)) {
      if (saldoCol && row[saldoCol] !== null && row[saldoCol] !== undefined) {
        return parseNumericValue(row[saldoCol]);
      }
      // fallback to any numeric in the row
      const numericValue = findNumericValueInRow(row);
      if (numericValue !== null) return numericValue;
    }
  }

  return 0;
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

// Extract TOTAL ASET LANCAR from VD51, row with "TOTAL ASET LANCAR" and column "Saldo"
export function extractVD51TotalAsetLancar(sheet: ProcessedSheet): number {
  const saldoCol = findColumnOptional(sheet.headers, /saldo/i);

  for (const row of sheet.data) {
    const rowText = Object.values(row)
      .filter(v => typeof v === 'string')
      .join(' ');

    if (/total\s*aset\s*lancar/i.test(rowText)) {
      if (saldoCol && row[saldoCol] !== null && row[saldoCol] !== undefined) {
        return parseNumericValue(row[saldoCol]);
      }
      const numericValue = findNumericValueInRow(row);
      if (numericValue !== null) return numericValue;
    }
  }

  return 0;
}

// Extract NILAI MKBD YANG DIWAJIBKAN from VD59 (typically row 103)
export function extractVD59MKBDDiwajibkan(sheet: ProcessedSheet): number {
  const totalCol = findColumnOptional(sheet.headers, /^total$/i);
  
  for (const row of sheet.data) {
    const rowText = Object.values(row)
      .filter(v => typeof v === 'string')
      .join(' ');
    
    if (/nilai\s*mkbd\s*yang\s*diwajibkan/i.test(rowText) || /mkbd\s*diwajibkan/i.test(rowText)) {
      if (totalCol && row[totalCol] !== null) {
        return parseNumericValue(row[totalCol]);
      }
    }
  }
  
  return 25_000_000_000; // Default minimum
}

// Process VD510 and calculate Ranking Liabilities using TOTAL EKUITAS from VD52
export function calculateVD510RankingLiabilities(
  sheet: ProcessedSheet,
  totalEkuitas: number
): VD510CalculationDetail[] {
  const details: VD510CalculationDetail[] = [];
  const batas20Persen = totalEkuitas * 0.20;

  // Find relevant columns - support GRUP_NILAI_PASAR_WAJAR from enrichment
  const kodeEfekCol = findColumnByPattern(sheet.headers, /kode\s*efek/i);
  const grupNilaiPasarCol = findColumnByPattern(sheet.headers, /grup[_\s]*nilai[_\s]*pasar[_\s]*wajar/i);
  const nilaiPasarCol = grupNilaiPasarCol || findColumnByPattern(sheet.headers, /nilai\s*pasar\s*wajar/i);
  const grupEmitenCol = findColumnByPattern(sheet.headers, /grup[_\s]*emiten/i);
  
  // Calculate for each row
  let rowIndex = 0;
  for (const row of sheet.data) {
    rowIndex++;
    const kodeEfek = String(row[kodeEfekCol] ?? '');
    const namaEfek = String(row['Jenis_Efek'] ?? row['Nama_Akun'] ?? kodeEfek);
    
    // Use GRUP_NILAI_PASAR_WAJAR if available, otherwise use nilaiPasarWajar
    const grupNilaiPasarWajar = parseNumericValue(row['GRUP_NILAI_PASAR_WAJAR'] ?? row[nilaiPasarCol]);
    const nilaiPasarWajar = parseNumericValue(row[nilaiPasarCol]);
    const grupEmiten = String(row[grupEmitenCol] ?? row['GRUP_EMITEN'] ?? 'Non-Grup');
    
    // Skip empty rows and non-portfolio placeholders
    if (grupNilaiPasarWajar === 0 || kodeEfek.toLowerCase().includes('other')) {
      continue;
    }
    
    // FORMULA (sesuai instruksi): Nilai_Rangking_Liabilities = GRUP_NILAI_PASAR_WAJAR - (20% * TOTAL EKUITAS)
    const nilaiRankingLiabilities = grupNilaiPasarWajar - batas20Persen;
    
    // Calculate percentage
    const persentaseTerhadapModal = totalEkuitas > 0 
      ? (grupNilaiPasarWajar / totalEkuitas) * 100 
      : 0;
    
    const formula = `${formatNumber(grupNilaiPasarWajar)} - (20% × ${formatNumber(totalEkuitas)})`;
    
    details.push({
      rowIndex,
      kodeEfek,
      namaEfek,
      nilaiPasarWajar,
      grupEmiten,
      persentaseTerhadapModal,
      batas20Persen,
      nilaiRankingLiabilities,
      formula,
    });
  }

  return details;
}

// Get Total Ranking Liabilities from row 33 (Total Portofolio)
export function getTotalRankingLiabilitiesFromVD510(
  vd510Details: VD510CalculationDetail[]
): number {
  // Sum all ranking liabilities (this represents Total Portofolio at row 33)
  return vd510Details.reduce((sum, item) => sum + item.nilaiRankingLiabilities, 0);
}

function cloneSheet(sheet: ProcessedSheet): ProcessedSheet {
  return {
    ...sheet,
    headers: [...sheet.headers],
    data: sheet.data.map(r => ({ ...r })),
    metadata: { ...sheet.metadata },
  };
}

function findRowIndexByText(
  sheet: ProcessedSheet,
  pattern: RegExp,
  preferredRowNumber?: number
): number | null {
  // prefer exact row index (1-based) if provided and matches
  if (preferredRowNumber && preferredRowNumber >= 1 && preferredRowNumber <= sheet.data.length) {
    const idx = preferredRowNumber - 1;
    const row = sheet.data[idx];
    const rowText = Object.values(row)
      .filter(v => typeof v === 'string')
      .join(' ');
    if (pattern.test(rowText)) return idx;
  }

  for (let i = 0; i < sheet.data.length; i++) {
    const row = sheet.data[i];
    const rowText = Object.values(row)
      .filter(v => typeof v === 'string')
      .join(' ');
    if (pattern.test(rowText)) return i;
  }

  return null;
}

function applyVD510Overwrite(
  sheet: ProcessedSheet,
  details: VD510CalculationDetail[],
  totalPortofolio: number
): { sheet: ProcessedSheet; updated: boolean } {
  const updatedSheet = cloneSheet(sheet);
  const nilaiRLCol = findColumnOptional(updatedSheet.headers, /nilai\s*rangking\s*liabilit/i);
  const kodeEfekCol = findColumnOptional(updatedSheet.headers, /kode\s*efek/i);
  const grupNilaiCol = findColumnOptional(updatedSheet.headers, /grup[_\s]*nilai[_\s]*pasar[_\s]*wajar/i);
  const nilaiPasarCol = findColumnOptional(updatedSheet.headers, /nilai\s*pasar\s*wajar/i);

  if (!nilaiRLCol || !kodeEfekCol) return { sheet: updatedSheet, updated: false };

  // index details by kode efek + grup nilai pasar (to reduce collision)
  const keyOf = (kodeEfek: string, grupNilai: number) => `${kodeEfek}__${grupNilai}`;
  const detailMap = new Map<string, VD510CalculationDetail>();
  for (const d of details) {
    detailMap.set(keyOf(d.kodeEfek, d.nilaiPasarWajar || 0), d);
  }

  // overwrite per-item rows
  for (const row of updatedSheet.data) {
    const kodeEfek = String(row[kodeEfekCol] ?? '').trim();
    if (!kodeEfek || kodeEfek.toLowerCase().includes('other')) continue;

    const grupNilai = parseNumericValue(row['GRUP_NILAI_PASAR_WAJAR'] ?? (grupNilaiCol ? row[grupNilaiCol] : undefined));
    const nilaiPasar = parseNumericValue(nilaiPasarCol ? row[nilaiPasarCol] : undefined);
    const detail = detailMap.get(keyOf(kodeEfek, nilaiPasar)) ?? detailMap.get(keyOf(kodeEfek, grupNilai));

    if (!detail) continue;
    row[nilaiRLCol] = detail.nilaiRankingLiabilities;
  }

  // overwrite Total Portofolio row
  const totalPortIdx = findRowIndexByText(updatedSheet, /total\s*portofolio/i);
  if (totalPortIdx !== null) {
    updatedSheet.data[totalPortIdx][nilaiRLCol] = totalPortofolio;
  }

  return { sheet: updatedSheet, updated: true };
}

function applyVD59Overwrite(
  sheet: ProcessedSheet,
  totalRankingLiabilities: number,
  totalAsetLancar: number,
  totalLiabilitas: number
): { sheet: ProcessedSheet; updates: VD59Update[] } {
  const updatedSheet = cloneSheet(sheet);
  const jumlahCol = findColumnOptional(updatedSheet.headers, /jumlah/i);
  const totalCol = findColumnOptional(updatedSheet.headers, /^total$/i);

  const updates: VD59Update[] = [];

  const modalKerja = totalAsetLancar - totalLiabilitas - totalRankingLiabilities;

  const row12Idx = findRowIndexByText(updatedSheet, /total\s*ranking\s*liabilit/i, 12);
  const row13Idx = findRowIndexByText(updatedSheet, /total\s*modal\s*kerja\s*\(baris\s*9\s*dikurangi\s*baris\s*11\s*dan\s*baris\s*12\)/i, 13);
  const row15Idx = findRowIndexByText(updatedSheet, /total\s*modal\s*kerja\s*\(baris\s*13\)/i, 15);
  const row18Idx = findRowIndexByText(updatedSheet, /total\s*modal\s*kerja\s*bersih\s*\(baris\s*15\s*ditambah\s*baris\s*17\)/i, 18);
  const row20Idx = findRowIndexByText(updatedSheet, /total\s*modal\s*kerja\s*bersih\s*\(baris\s*18\)/i, 20);
  const row102Idx = findRowIndexByText(updatedSheet, /total\s*modal\s*kerja\s*bersih\s*disesuaikan/i, 102);
  const row104Idx = findRowIndexByText(updatedSheet, /lebih\s*\(kurang\)\s*mkbd/i, 104);
  const row103Idx = findRowIndexByText(updatedSheet, /nilai\s*mkbd\s*yang\s*diwajibkan/i, 103);

  // helper to write cell + record update
  const write = (rowIdx: number | null, col: string, newValue: number, formula: string) => {
    if (rowIdx === null) return;
    const row = updatedSheet.data[rowIdx];
    const oldValue = parseNumericValue(row[col]);
    row[col] = newValue;
    updates.push({
      rowIndex: rowIdx + 1,
      rowDescription: Object.values(row).filter(v => typeof v === 'string').join(' ').slice(0, 80),
      column: col,
      oldValue,
      newValue,
      formula,
    });
  };

  // A) Row 12 Jumlah = VD510 Total Portofolio (Nilai_Rangking_Liabilities)
  if (jumlahCol) {
    write(row12Idx, jumlahCol, totalRankingLiabilities, 'VD510: Total Portofolio (Nilai_Rangking_Liabilities)');
  }

  // B) Row 13 Jumlah = TOTAL_ASET_LANCAR - TOTAL_LIABILITAS - TOTAL_RANKING_LIABILITIES
  if (jumlahCol) {
    write(row13Idx, jumlahCol, modalKerja, 'TOTAL ASET LANCAR - TOTAL LIABILITAS - TOTAL RANKING LIABILITIES');
  }

  // C) Penyamaan nilai: Row 15 Jumlah, Row 18 Jumlah, Row 20 Total harus IDENTIK dengan modalKerja (row 13)
  if (jumlahCol) {
    write(row15Idx, jumlahCol, modalKerja, 'Sama persis dengan TOTAL MODAL KERJA (Row 13)');
    write(row18Idx, jumlahCol, modalKerja, 'Sama persis dengan TOTAL MODAL KERJA (Row 13)');
  }
  if (totalCol) {
    write(row20Idx, totalCol, modalKerja, 'Sama persis dengan TOTAL MODAL KERJA (Row 13)');
  }

  // D) Row 102 Total = Row 20 Total - SUM(Total rows 33..92)
  let haircutSum = 0;
  if (totalCol) {
    for (let i = 32; i < Math.min(92, updatedSheet.data.length); i++) {
      const r = updatedSheet.data[i];
      haircutSum += parseNumericValue(r[totalCol]);
    }
    const mkbdDisesuaikan = modalKerja - haircutSum;
    write(row102Idx, totalCol, mkbdDisesuaikan, 'TOTAL MODAL KERJA BERSIH (Row 18) - SUM(Total rows 33..92)');

    // E) Row 104 Total = Row 102 Total - NILAI MKBD YANG DIWAJIBKAN (Row 103 Total)
    const mkbdDiwajibkan = row103Idx !== null ? parseNumericValue(updatedSheet.data[row103Idx][totalCol]) : 0;
    const lebihKurang = mkbdDisesuaikan - mkbdDiwajibkan;
    write(row104Idx, totalCol, lebihKurang, 'TOTAL MODAL KERJA BERSIH DISESUAIKAN - NILAI MKBD YANG DIWAJIBKAN');
  }

  return { sheet: updatedSheet, updates };
}

/**
 * Force-overwrite nilai pada sheet VD510 & VD59 sesuai rule audit.
 * Output ini yang dipakai untuk preview & simpan database, sehingga tidak lagi mengikuti angka Excel lama (untuk bagian yang di-highlight).
 */
export function applyMKBDCorrections(sheets: ProcessedSheet[]): { sheets: ProcessedSheet[]; vd59Updates: VD59Update[] } {
  const out = sheets.map(cloneSheet);

  const vd51 = out.find(s => /vd5[\-_]?1\b|formulir[\-\_\s]*1\b/i.test(s.sheetName));
  const vd52 = out.find(s => /vd5[\-_]?2\b|formulir[\-\_\s]*2\b/i.test(s.sheetName));
  const vd59 = out.find(s => /vd5[\-_]?9\b|formulir[\-\_\s]*9\b/i.test(s.sheetName));
  const vd510 = out.find(s => /vd5[\-_]?10\b|formulir[\-\_\s]*10\b/i.test(s.sheetName));

  const totalAsetLancar = vd51 ? extractVD51TotalAsetLancar(vd51) : 0;
  const totalLiabilitas = vd52 ? extractVD52TotalLiabilitas(vd52) : 0;
  const totalEkuitas = vd52 ? extractVD52TotalEkuitas(vd52) : 0;

  const vd59Updates: VD59Update[] = [];

  if (vd510 && totalEkuitas !== 0) {
    const details = calculateVD510RankingLiabilities(vd510, totalEkuitas);
    const totalPortofolio = getTotalRankingLiabilitiesFromVD510(details);
    const res = applyVD510Overwrite(vd510, details, totalPortofolio);
    if (res.updated) {
      const idx = out.findIndex(s => s.sheetName === vd510.sheetName);
      out[idx] = res.sheet;

      if (vd59) {
        const vd59Res = applyVD59Overwrite(vd59, totalPortofolio, totalAsetLancar, totalLiabilitas);
        const idx59 = out.findIndex(s => s.sheetName === vd59.sheetName);
        out[idx59] = vd59Res.sheet;
        vd59Updates.push(...vd59Res.updates);
      }
    }
  }

  return { sheets: out, vd59Updates };
}

// Calculate VD59 updates based on the formulas
export function calculateVD59Updates(
  sheet: ProcessedSheet,
  totalAsetLancar: number,
  totalLiabilitas: number,
  totalRankingLiabilities: number,
  mkbdDiwajibkan: number
): { updates: VD59Update[]; haircutSum: number; mkbdDisesuaikan: number; lebihKurangMKBD: number } {
  const updates: VD59Update[] = [];
  const jumlahCol = findColumnByPattern(sheet.headers, /jumlah/i);
  const totalCol = findColumnByPattern(sheet.headers, /^total$/i);
  
  // Calculate TOTAL MODAL KERJA = TOTAL ASET LANCAR - TOTAL LIABILITAS - TOTAL RANKING LIABILITIES
  const totalModalKerja = totalAsetLancar - totalLiabilitas - totalRankingLiabilities;
  
  // Calculate haircut sum (rows 33-92 Total column)
  let haircutSum = 0;
  for (let i = 32; i < Math.min(92, sheet.data.length); i++) { // 0-indexed: 32 = row 33
    const row = sheet.data[i];
    if (totalCol && row[totalCol] !== null) {
      haircutSum += parseNumericValue(row[totalCol]);
    }
  }
  
  // MKBD Disesuaikan = TOTAL MODAL KERJA BERSIH - haircut sum
  const mkbdDisesuaikan = totalModalKerja - haircutSum;
  
  // Lebih/(Kurang) MKBD = MKBD Disesuaikan - MKBD Diwajibkan
  const lebihKurangMKBD = mkbdDisesuaikan - mkbdDiwajibkan;
  
  // Define rows to update based on user requirements
  const updateRules = [
    { 
      rowIndex: 12, 
      pattern: /total\s*ranking\s*liabilit/i,
      column: jumlahCol,
      newValue: totalRankingLiabilities,
      formula: 'VD510 Total Portofolio (Nilai_Rangking_Liabilities)',
    },
    {
      rowIndex: 13,
      pattern: /total\s*modal\s*kerja.*baris\s*9.*dikurangi.*baris\s*11.*baris\s*12/i,
      column: jumlahCol,
      newValue: totalModalKerja,
      formula: 'TOTAL ASET LANCAR - TOTAL LIABILITAS - TOTAL RANKING LIABILITIES',
    },
    {
      rowIndex: 15,
      pattern: /total\s*modal\s*kerja.*baris\s*9.*dikurangi.*baris\s*11.*baris\s*12/i,
      column: jumlahCol,
      newValue: totalModalKerja,
      formula: 'Same as Row 13 (TOTAL MODAL KERJA)',
    },
    {
      rowIndex: 18,
      pattern: /total\s*modal\s*kerja\s*bersih.*baris\s*15.*ditambah.*baris\s*17/i,
      column: jumlahCol,
      newValue: totalModalKerja,
      formula: 'Same as Row 13 (TOTAL MODAL KERJA)',
    },
    {
      rowIndex: 20,
      pattern: /total\s*modal\s*kerja\s*bersih.*baris\s*18/i,
      column: totalCol,
      newValue: totalModalKerja,
      formula: 'Same as Row 13 (TOTAL MODAL KERJA)',
    },
    {
      rowIndex: 102,
      pattern: /total\s*modal\s*kerja\s*bersih\s*disesuaikan/i,
      column: totalCol,
      newValue: mkbdDisesuaikan,
      formula: 'Row 20 - SUM(Rows 33-92 Total)',
    },
    {
      rowIndex: 104,
      pattern: /lebih.*kurang.*mkbd/i,
      column: totalCol,
      newValue: lebihKurangMKBD,
      formula: 'MKBD DISESUAIKAN - NILAI MKBD YANG DIWAJIBKAN',
    },
  ];
  
  // Apply updates
  for (const rule of updateRules) {
    const rowIdx = rule.rowIndex - 1; // Convert to 0-indexed
    if (rowIdx >= 0 && rowIdx < sheet.data.length && rule.column) {
      const row = sheet.data[rowIdx];
      const rowText = Object.values(row)
        .filter(v => typeof v === 'string')
        .join(' ');
      
      const oldValue = parseNumericValue(row[rule.column]);
      
      // Check if row matches pattern or use row index directly
      const matchesPattern = rule.pattern.test(rowText);
      
      updates.push({
        rowIndex: rule.rowIndex,
        rowDescription: matchesPattern ? rowText.substring(0, 60) : `Row ${rule.rowIndex}`,
        column: rule.column,
        oldValue,
        newValue: rule.newValue,
        formula: rule.formula,
      });
    }
  }
  
  return { updates, haircutSum, mkbdDisesuaikan, lebihKurangMKBD };
}

// Main calculation function - Multi-Pass
export function calculateMKBD(sheets: ProcessedSheet[]): MKBDCalculationResult {
  const calculationSteps: CalculationStep[] = [];
  const vd59Updates: VD59Update[] = [];
  
  // Find relevant sheets
  const vd51Sheet = sheets.find(s => /vd5[\-_]?1\b|formulir[\-_\s]*1\b/i.test(s.sheetName));
  const vd59Sheet = sheets.find(s => /vd5[\-_]?9\b|formulir[\-_\s]*9\b/i.test(s.sheetName));
  const vd510Sheet = sheets.find(s => /vd5[\-_]?10\b|formulir[\-_\s]*10\b/i.test(s.sheetName));
  const vd52Sheet = sheets.find(s => /vd5[\-_]?2\b|formulir[\-_\s]*2\b/i.test(s.sheetName));

  // === PASS 1: Extract key bases (VD51 + VD52 + VD59[mkbd diwajibkan]) ===
  let totalAsetLancar = 0;
  let totalEkuitas = 0;
  let totalLiabilitas = 0;
  let mkbdDiwajibkan = 25_000_000_000;

  if (vd51Sheet) {
    totalAsetLancar = extractVD51TotalAsetLancar(vd51Sheet);

    calculationSteps.push({
      id: 'pass1_aset_lancar',
      name: 'Total Aset Lancar (VD51, TOTAL ASET LANCAR kolom Saldo)',
      formula: 'Extract from VD51 row "TOTAL ASET LANCAR", column "Saldo"',
      inputValues: { value: totalAsetLancar },
      result: totalAsetLancar,
      source: 'VD51',
      editable: false,
    });
  }

  if (vd52Sheet) {
    totalLiabilitas = extractVD52TotalLiabilitas(vd52Sheet);
    totalEkuitas = extractVD52TotalEkuitas(vd52Sheet);

    calculationSteps.push({
      id: 'pass1_liabilitas',
      name: 'Total Liabilitas (VD52, TOTAL LIABILITAS kolom Saldo)',
      formula: 'Extract from VD52 row "TOTAL LIABILITAS", column "Saldo"',
      inputValues: { value: totalLiabilitas },
      result: totalLiabilitas,
      source: 'VD52',
      editable: false,
    });

    calculationSteps.push({
      id: 'pass1_total_ekuitas',
      name: 'Total Ekuitas (VD52, TOTAL EKUITAS kolom Saldo)',
      formula: 'Extract from VD52 row "TOTAL EKUITAS", column "Saldo"',
      inputValues: { value: totalEkuitas },
      result: totalEkuitas,
      source: 'VD52',
      editable: false,
    });
  }

  if (vd59Sheet) {
    mkbdDiwajibkan = extractVD59MKBDDiwajibkan(vd59Sheet);
    
    calculationSteps.push({
      id: 'pass1_mkbd_diwajibkan',
      name: 'NILAI MKBD YANG DIWAJIBKAN',
      formula: 'Extract from VD59 (minimum Rp 25 Miliar)',
      inputValues: { value: mkbdDiwajibkan },
      result: mkbdDiwajibkan,
      source: 'VD59',
      editable: false,
    });
  }

  // === PASS 2: Calculate VD510 Ranking Liabilities using TOTAL EKUITAS ===
  let vd510Details: VD510CalculationDetail[] = [];
  let totalRankingLiabilities = 0;

  if (vd510Sheet && totalEkuitas !== 0) {
    vd510Details = calculateVD510RankingLiabilities(vd510Sheet, totalEkuitas);
    
    // Get Total Ranking Liabilities (sum from VD510, represents row 33 Total Portofolio)
    totalRankingLiabilities = getTotalRankingLiabilitiesFromVD510(vd510Details);

    calculationSteps.push({
      id: 'pass2_ranking_liabilities',
      name: 'Total Ranking Liabilities (VD510 Row 33)',
      formula: 'SUM(GRUP_NILAI_PASAR_WAJAR - 20% × TOTAL_EKUITAS)',
      inputValues: { 
        totalEkuitas,
        batas20Persen: totalEkuitas * 0.20,
        itemCount: vd510Details.length,
      },
      result: totalRankingLiabilities,
      source: 'VD510',
      editable: true,
    });
  }

  // === PASS 3: Calculate VD59 Updates ===
  let modalKerja = totalAsetLancar - totalLiabilitas - totalRankingLiabilities;
  let haircutSum = 0;
  let mkbdDisesuaikan = modalKerja;
  let lebihKurangMKBD = mkbdDisesuaikan - mkbdDiwajibkan;
  
  if (vd59Sheet) {
    const vd59Result = calculateVD59Updates(
      vd59Sheet,
      totalAsetLancar,
      totalLiabilitas,
      totalRankingLiabilities,
      mkbdDiwajibkan
    );
    
    vd59Updates.push(...vd59Result.updates);
    haircutSum = vd59Result.haircutSum;
    mkbdDisesuaikan = vd59Result.mkbdDisesuaikan;
    lebihKurangMKBD = vd59Result.lebihKurangMKBD;
  }
  
  // Modal Kerja Bersih (same as modalKerja in this simplified version)
  const modalKerjaBersih = modalKerja;

  calculationSteps.push({
    id: 'pass3_modal_kerja',
    name: 'Total Modal Kerja (Row 13)',
    formula: 'TOTAL_ASET_LANCAR - TOTAL_LIABILITAS - TOTAL_RANKING_LIABILITIES',
    inputValues: { totalAsetLancar, totalLiabilitas, totalRankingLiabilities },
    result: modalKerja,
    source: 'Calculated',
    editable: true,
  });

  calculationSteps.push({
    id: 'pass3_haircut_sum',
    name: 'Haircut Sum (Rows 33-92)',
    formula: 'SUM(Total column, rows 33-92)',
    inputValues: { haircutSum },
    result: haircutSum,
    source: 'VD59',
    editable: false,
  });

  calculationSteps.push({
    id: 'pass3_mkbd_disesuaikan',
    name: 'MKBD Disesuaikan (Row 102)',
    formula: 'TOTAL_MODAL_KERJA_BERSIH - HAIRCUT_SUM',
    inputValues: { modalKerjaBersih, haircutSum },
    result: mkbdDisesuaikan,
    source: 'Calculated',
    editable: true,
  });

  calculationSteps.push({
    id: 'pass3_lebih_kurang',
    name: 'Lebih/(Kurang) MKBD (Row 104)',
    formula: 'MKBD_DISESUAIKAN - MKBD_DIWAJIBKAN',
    inputValues: { mkbdDisesuaikan, mkbdDiwajibkan },
    result: lebihKurangMKBD,
    source: 'Calculated',
    editable: false,
  });

  return {
    totalAsetLancar,
    totalEkuitas,
    totalLiabilitas,
    totalRankingLiabilities,
    modalKerja,
    modalKerjaBersih,
    mkbdDisesuaikan,
    mkbdDiwajibkan,
    lebihKurangMKBD,
    calculationSteps,
    vd510Details,
    vd59Updates,
    haircutSum,
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

function findColumnOptional(headers: string[], pattern: RegExp): string | undefined {
  return headers.find(h => pattern.test(h));
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
