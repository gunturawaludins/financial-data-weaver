// Types for ETL System

export interface SheetData {
  sheetName: string;
  headers: string[];
  rows: Record<string, unknown>[];
  originalHeaders: string[];
  rowCount: number;
}

export interface EnrichmentStats {
  kodeEfekColumn: string | null;
  nilaiPasarColumn: string | null;
  matchedCount: number;
  unmatchedCount: number;
  groupCount: number;
  totalGroupValue: number;
}

export interface ProcessedSheet {
  sheetName: string;
  tableName: string;
  headers: string[];
  data: Record<string, unknown>[];
  metadata: {
    fileName: string;
    uploadDate: string;
    originalRowCount: number;
    cleanedRowCount: number;
    enrichmentStats?: EnrichmentStats | null;
  };
}

export interface ETLResult {
  success: boolean;
  sheets: ProcessedSheet[];
  errors: string[];
  warnings: string[];
}

export interface DatabaseRecord extends Record<string, unknown> {
  _id?: number;
  _fileName: string;
  _uploadDate: string;
}

export interface StoredTable {
  tableName: string;
  records: DatabaseRecord[];
  lastUpdated: string;
}

// MKBD Calculation Types
export interface VD59Update {
  rowIndex: number;
  rowDescription: string;
  rowLabel?: string;
  column: string;
  oldValue: number;
  newValue: number;
  formula: string;
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

export interface CalculationStep {
  id: string;
  name: string;
  formula: string;
  inputValues: Record<string, number>;
  result: number;
  source: string;
  editable: boolean;
}

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
