import * as XLSX from 'xlsx';

export interface ExtractedToken {
  rawHeader: string;      // e.g. "Roll Number "
  cleanHeader: string;    // e.g. "Roll Number"
  token: string;          // e.g. "<<Roll Number>>"
  sampleValue?: string;   // e.g. "25CSR220"
}

export interface ParsedExcelResult {
  headers: string[];
  tokens: ExtractedToken[];
  rows: Record<string, any>[];
  totalRows: number;
  fileName?: string;
}

export function parseExcelWorkbook(buffer: ArrayBuffer, fileName?: string): ParsedExcelResult {
  const data = new Uint8Array(buffer);
  const workbook = XLSX.read(data, { type: 'array', cellDates: true });
  const firstSheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[firstSheetName];

  // Convert to 2D array to inspect exact headers
  const jsonRows = XLSX.utils.sheet_to_json<Record<string, any>>(worksheet, { defval: '' });

  if (jsonRows.length === 0) {
    return {
      headers: [],
      tokens: [],
      rows: [],
      totalRows: 0,
      fileName
    };
  }

  // Get raw headers from first object keys
  const rawHeaders = Object.keys(jsonRows[0]);
  
  const tokens: ExtractedToken[] = rawHeaders.map((header) => {
    const clean = header.trim();
    return {
      rawHeader: header,
      cleanHeader: clean,
      token: `<<${clean}>>`,
      sampleValue: String(jsonRows[0][header] ?? '')
    };
  });

  return {
    headers: rawHeaders,
    tokens,
    rows: jsonRows,
    totalRows: jsonRows.length,
    fileName
  };
}

/**
 * Normalizes dynamic values from Excel rows (handles Date objects, nulls, numbers)
 */
export function formatCellValue(val: any): string {
  if (val === null || val === undefined) return '';
  if (val instanceof Date) {
    return val.toLocaleDateString('en-US', { year: 'numeric', month: 'short', day: 'numeric', timeZone: 'UTC' });
  }
  return String(val).trim();
}
