import * as XLSX from 'xlsx';
import { ProcessedFile } from '../types';

/**
 * Utility to repair garbled Chinese characters often found in older .xls files.
 * This should ONLY be applied if we are sure the string is misinterpreted Latin-1.
 */
const repairEncoding = (val: any, isOldFormat: boolean): any => {
  if (typeof val !== 'string' || !isOldFormat || val.length === 0) return val;
  
  // HEURISTIC: Misinterpreted Big5-as-Latin-1 strings ALWAYS consist 
  // ONLY of characters in the 0-255 range (the byte values).
  // If there is even one character > 255 (like a real Chinese Unicode char),
  // then the string is NOT garbled Latin-1 and we must not touch it.
  let hasNonAscii = false;
  for (let i = 0; i < val.length; i++) {
    const code = val.charCodeAt(i);
    if (code > 255) return val; // Already has real Unicode/Chinese, skip repair
    if (code > 127) hasNonAscii = true; // Found potential Big5 byte
  }

  // If it's pure ASCII (0-127), no repair needed
  if (!hasNonAscii) return val;

  try {
    // Convert the string characters back to their raw byte values (0-255)
    const bytes = new Uint8Array(val.split('').map(c => c.charCodeAt(0) & 0xFF));
    
    // Attempt to decode as Big5. 
    // fatal: true helps us detect if it's actually valid Big5 or just random binary data.
    const decoder = new TextDecoder('big5', { fatal: false });
    const decoded = decoder.decode(bytes);
    
    // Safety check: if the decoded result is empty or just replacement chars, fall back.
    if (!decoded || decoded.includes('\ufffd') && decoded.length < val.length / 2) {
      return val;
    }
    
    return decoded;
  } catch (e) {
    return val;
  }
};

/**
 * Reads a file and parses it into JSON
 */
export const parseExcelFile = (file: File): Promise<ProcessedFile> => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        const data = e.target?.result;
        if (!data) throw new Error("File is empty");

        const fileNameLower = file.name.toLowerCase();
        const isOldFormat = fileNameLower.endsWith('.xls') && !fileNameLower.endsWith('.xlsx');

        // Read the workbook. 
        const workbook = XLSX.read(data, { 
          type: 'array',
          cellDates: true,
          cellFormula: false,
          cellStyles: false,
          cellNF: true,
          // Only force codepage for truly old XLS files
          codepage: isOldFormat ? 950 : undefined 
        });
        
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        // Parse to JSON
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
        
        // Handle filename modernization
        let finalFileName = file.name;
        if (isOldFormat) {
          finalFileName = finalFileName.replace(/\.xls$/i, '.xlsx');
        }

        if (jsonData.length === 0) {
           resolve({
            id: crypto.randomUUID(),
            name: finalFileName,
            size: file.size,
            rowCount: 0,
            headers: [],
            data: [],
            error: "No data found in the first sheet"
          });
          return;
        }

        // Normalize data: 
        // 1. Repair encoding ONLY for old .xls formats
        // 2. Trim whitespace from headers
        const normalizedData = jsonData.map((row: any) => {
          const newRow: any = {};
          Object.keys(row).forEach(key => {
            // Repair the key (header) name - only if it's an old file
            const repairedKey = repairEncoding(key, isOldFormat).trim();
            // Repair the value content - only if it's an old file
            const repairedValue = repairEncoding(row[key], isOldFormat);
            newRow[repairedKey] = repairedValue;
          });
          return newRow;
        });

        // Extract headers from the first row
        const headers = Object.keys(normalizedData[0] as object);

        resolve({
          id: crypto.randomUUID(),
          name: finalFileName,
          size: file.size,
          rowCount: jsonData.length,
          headers,
          data: normalizedData,
        });
      } catch (err) {
        console.error("Error parsing excel:", err);
        reject(err);
      }
    };

    reader.onerror = (err) => reject(err);
    reader.readAsArrayBuffer(file);
  });
};

/**
 * Merges data from multiple files and sorts them
 */
export const mergeData = (files: ProcessedFile[], sortKey: string): any[] => {
  const allData = files.flatMap(file => file.data);
  const cleanSortKey = sortKey.trim();

  return allData.sort((a, b) => {
    const valA = a[cleanSortKey];
    const valB = b[cleanSortKey];

    if (valA === valB) return 0;
    if (valA === undefined || valA === null) return 1;
    if (valB === undefined || valB === null) return -1;

    const numA = Number(valA);
    const numB = Number(valB);

    if (!isNaN(numA) && !isNaN(numB) && valA !== '' && valB !== '') {
        return numA - numB;
    }

    return String(valA).localeCompare(String(valB));
  });
};

/**
 * Exports data to an XLSX file
 */
export const exportToExcel = (data: any[], filename: string) => {
  const worksheet = XLSX.utils.json_to_sheet(data);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Merged Data");

  const safeFilename = filename.endsWith('.xlsx') ? filename : `${filename}.xlsx`;

  // Write and download with compression
  XLSX.writeFile(workbook, safeFilename, { compression: true });
};