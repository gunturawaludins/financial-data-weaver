// Master Data Manager - Handles emiten reference data for grouping and aggregation
import * as XLSX from 'xlsx';

export interface EmitenMaster {
  kode: string;
  namaEmiten: string;
  afiliasiUtama: string;
  subAfiliasi: string;
  uboTokohKunci: string;
  kategori: string;
}

// In-memory storage for master data
let masterDataCache: Map<string, EmitenMaster> = new Map();
let isMasterDataLoaded = false;

/**
 * Parse master data from Excel file
 */
export async function loadMasterDataFromFile(file: File): Promise<{
  success: boolean;
  count: number;
  errors: string[];
}> {
  const errors: string[] = [];
  
  try {
    const arrayBuffer = await file.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: 'array' });
    
    // Assume first sheet contains the data
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    
    // Convert to JSON with headers
    const rawData = XLSX.utils.sheet_to_json<Record<string, unknown>>(worksheet);
    
    // Clear existing cache
    masterDataCache.clear();
    
    // Parse each row
    for (const row of rawData) {
      const kode = normalizeKode(row['Kode'] as string);
      
      if (!kode) {
        continue; // Skip rows without code
      }
      
      const emiten: EmitenMaster = {
        kode,
        namaEmiten: String(row['Nama Emiten'] || ''),
        afiliasiUtama: String(row['Afiliasi Utama'] || ''),
        subAfiliasi: String(row['Sub-Afiliasi'] || ''),
        uboTokohKunci: String(row['UBO / Tokoh Kunci'] || ''),
        kategori: String(row['Kategori'] || ''),
      };
      
      masterDataCache.set(kode, emiten);
    }
    
    isMasterDataLoaded = true;
    
    return {
      success: true,
      count: masterDataCache.size,
      errors,
    };
  } catch (error) {
    errors.push(`Gagal memuat master data: ${error}`);
    return {
      success: false,
      count: 0,
      errors,
    };
  }
}

/**
 * Load master data from pre-bundled file in public folder
 */
export async function loadMasterDataFromPublic(): Promise<{
  success: boolean;
  count: number;
  errors: string[];
}> {
  const errors: string[] = [];
  
  try {
    const response = await fetch('/data/master-emiten.xlsx');
    if (!response.ok) {
      throw new Error('Master data file not found');
    }
    
    const arrayBuffer = await response.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: 'array' });
    
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    const rawData = XLSX.utils.sheet_to_json<Record<string, unknown>>(worksheet);
    
    masterDataCache.clear();
    
    for (const row of rawData) {
      const kode = normalizeKode(row['Kode'] as string);
      
      if (!kode) continue;
      
      const emiten: EmitenMaster = {
        kode,
        namaEmiten: String(row['Nama Emiten'] || ''),
        afiliasiUtama: String(row['Afiliasi Utama'] || ''),
        subAfiliasi: String(row['Sub-Afiliasi'] || ''),
        uboTokohKunci: String(row['UBO / Tokoh Kunci'] || ''),
        kategori: String(row['Kategori'] || ''),
      };
      
      masterDataCache.set(kode, emiten);
    }
    
    isMasterDataLoaded = true;
    
    return {
      success: true,
      count: masterDataCache.size,
      errors,
    };
  } catch (error) {
    errors.push(`Gagal memuat master data: ${error}`);
    return {
      success: false,
      count: 0,
      errors,
    };
  }
}

/**
 * Normalize stock code for consistent lookup
 */
function normalizeKode(kode: unknown): string {
  if (kode === null || kode === undefined) return '';
  return String(kode).trim().toUpperCase();
}

/**
 * Lookup emiten by stock code
 */
export function lookupEmiten(kode: string): EmitenMaster | null {
  const normalizedKode = normalizeKode(kode);
  return masterDataCache.get(normalizedKode) || null;
}

/**
 * Get group name (Afiliasi Utama) for a stock code
 */
export function getGroupName(kode: string): string {
  const emiten = lookupEmiten(kode);
  return emiten?.afiliasiUtama || 'Non-Grup';
}

/**
 * Check if master data is loaded
 */
export function isMasterLoaded(): boolean {
  return isMasterDataLoaded;
}

/**
 * Get all master data entries
 */
export function getAllMasterData(): EmitenMaster[] {
  return Array.from(masterDataCache.values());
}

/**
 * Get master data statistics
 */
export function getMasterDataStats(): {
  totalEmiten: number;
  uniqueGroups: number;
  categories: Record<string, number>;
} {
  const categories: Record<string, number> = {};
  const groups = new Set<string>();
  
  for (const emiten of masterDataCache.values()) {
    groups.add(emiten.afiliasiUtama);
    categories[emiten.kategori] = (categories[emiten.kategori] || 0) + 1;
  }
  
  return {
    totalEmiten: masterDataCache.size,
    uniqueGroups: groups.size,
    categories,
  };
}

/**
 * Clear master data cache
 */
export function clearMasterData(): void {
  masterDataCache.clear();
  isMasterDataLoaded = false;
}
