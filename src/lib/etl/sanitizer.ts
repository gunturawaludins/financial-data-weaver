// Column Sanitizer - Handles duplicate, numeric, and empty column names

export interface SanitizedResult {
  sanitizedHeaders: string[];
  originalToSanitized: Map<number, string>;
  warnings: string[];
}

export function sanitizeColumnNames(headers: (string | number | null | undefined)[]): SanitizedResult {
  const sanitizedHeaders: string[] = [];
  const originalToSanitized = new Map<number, string>();
  const warnings: string[] = [];
  const usedNames = new Map<string, number>();

  headers.forEach((header, index) => {
    let sanitized = sanitizeSingleColumn(header, index);
    
    // Handle duplicates
    const baseName = sanitized;
    let counter = usedNames.get(baseName.toLowerCase()) || 0;
    
    if (counter > 0) {
      sanitized = `${baseName}_${counter}`;
      warnings.push(`Kolom duplikat ditemukan: "${header}" diubah menjadi "${sanitized}"`);
    }
    
    usedNames.set(baseName.toLowerCase(), counter + 1);
    sanitizedHeaders.push(sanitized);
    originalToSanitized.set(index, sanitized);
  });

  return { sanitizedHeaders, originalToSanitized, warnings };
}

function sanitizeSingleColumn(header: string | number | null | undefined, index: number): string {
  // Handle null, undefined, or empty
  if (header === null || header === undefined || header === '') {
    return `Unnamed_${index + 1}`;
  }

  const headerStr = String(header).trim();

  // Handle empty string after trim
  if (headerStr === '') {
    return `Unnamed_${index + 1}`;
  }

  // Handle pure numeric headers (e.g., "8.0", "7", "10")
  if (/^[\d.]+$/.test(headerStr)) {
    const numVal = parseFloat(headerStr);
    if (!isNaN(numVal)) {
      return `Meta_Baris_${Math.round(numVal)}`;
    }
  }

  // Clean special characters that could cause database issues
  let cleaned = headerStr
    .replace(/[^\w\s\u00C0-\u024F]/g, '_') // Replace special chars with underscore
    .replace(/\s+/g, '_') // Replace spaces with underscore
    .replace(/_+/g, '_') // Remove multiple underscores
    .replace(/^_|_$/g, ''); // Remove leading/trailing underscores

  // Ensure it doesn't start with a number
  if (/^\d/.test(cleaned)) {
    cleaned = `Col_${cleaned}`;
  }

  // Limit length to prevent extremely long column names
  if (cleaned.length > 64) {
    cleaned = cleaned.substring(0, 64);
  }

  return cleaned || `Unnamed_${index + 1}`;
}

export function sanitizeTableName(sheetName: string): string {
  let sanitized = sheetName
    .replace(/[^\w\s]/g, '_')
    .replace(/\s+/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_|_$/g, '')
    .toLowerCase();

  // Ensure it doesn't start with a number
  if (/^\d/.test(sanitized)) {
    sanitized = `table_${sanitized}`;
  }

  return sanitized || 'unknown_table';
}
