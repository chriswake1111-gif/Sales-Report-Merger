import * as XLSX from 'xlsx';
import { ProcessedFile } from '../types';

/**
 * Utility to repair garbled Chinese characters.
 * Useful when Big5 bytes are interpreted as Latin-1 (iso-8859-1).
 */
const repairEncoding = (val: any): any => {
  if (typeof val !== 'string' || val.length === 0) return val;
  
  // HEURISTIC: Check if the string consists ONLY of characters in the 0-255 range.
  // If there is any character > 255 (e.g. a valid Chinese char), we assume the string 
  // is already correctly decoded (or partially correct) and we shouldn't mess with it.
  let hasExtendedAscii = false;
  for (let i = 0; i < val.length; i++) {
    const code = val.charCodeAt(i);
    if (code > 255) return val; // Contains valid Unicode/Chinese, skip repair
    if (code > 127) hasExtendedAscii = true; // Contains bytes that might be Big5
  }

  // If strictly standard ASCII (0-127), no repair needed.
  if (!hasExtendedAscii) return val;

  try {
    // Convert Latin-1 characters back to raw bytes
    const bytes = new Uint8Array(val.split('').map(c => c.charCodeAt(0) & 0xFF));
    
    // Attempt to decode as Big5
    const decoder = new TextDecoder('big5', { fatal: false });
    const decoded = decoder.decode(bytes);
    
    // If decoding creates replacement chars () or is empty, it failed.
    // We compare with the original length/content to decide if it looks like a valid fix.
    if (!decoded || (decoded.includes('\ufffd') && decoded.length < val.length)) {
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
        
        // Handle .csv explicitly in logic if needed, but XLSX.read usually detects content.
        // We force codepage 950 (Big5) because:
        // 1. Genuine .xlsx (XML) files ignore this option (safe).
        // 2. Old .xls (BIFF8) files need it if headers are missing info.
        // 3. CSV files (often renamed to .xlsx) in Taiwan are almost always Big5.
        const workbook = XLSX.read(data, { 
          type: 'array',
          cellDates: true,
          cellFormula: false,
          cellStyles: false,
          cellNF: true,
          codepage: 950 // FORCE Big5 for any text/legacy fallback
        });
        
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        // Parse to JSON
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
        
        let finalFileName = file.name;
        // Normalize extension for output consistency if it was a legacy file
        if (fileNameLower.endsWith('.xls') || fileNameLower.endsWith('.csv')) {
          finalFileName = finalFileName.replace(/\.(xls|csv)$/i, '.xlsx');
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

        // Normalize data
        const normalizedData = jsonData.map((row: any) => {
          const newRow: any = {};
          Object.keys(row).forEach(key => {
            // Trim whitespace
            // Run repairEncoding just in case logic fell back to Latin-1
            const repairedKey = repairEncoding(key).trim();
            const repairedValue = repairEncoding(row[key]);
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