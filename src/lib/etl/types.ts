// Types for ETL System

export interface SheetData {
  sheetName: string;
  headers: string[];
  rows: Record<string, unknown>[];
  originalHeaders: string[];
  rowCount: number;
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
