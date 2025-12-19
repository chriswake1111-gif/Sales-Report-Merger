import * as XLSX from 'xlsx';
import { ProcessedFile } from '../types';

/**
 * Utility to repair garbled Chinese characters often found in older .xls files.
 * This happens when Big5 bytes are misinterpreted as Latin-1 characters.
 */
const repairEncoding = (val: any): any => {
  if (typeof val !== 'string') return val;
  
  // Heuristic: Check if the string contains characters in the Latin-1 supplement range (0x80-0xFF)
  // which are typical symptoms of misinterpreted Big5 bytes (e.g., "單號" becoming "³æ¸¹").
  if (/[^\x00-\x7F]/.test(val)) {
    try {
      // Convert the string characters back to their raw byte values (0-255)
      const bytes = new Uint8Array(val.split('').map(c => c.charCodeAt(0) & 0xFF));
      
      // Attempt to decode as Big5 (Standard for Traditional Chinese in .xls)
      // 'big5' decoder is built-in to most modern browsers
      const decoded = new TextDecoder('big5').decode(bytes);
      
      // If the decoded string contains replacement characters (), the decoding might have failed.
      // Otherwise, return the repaired string.
      return decoded.includes('') ? val : decoded;
    } catch (e) {
      // If decoding fails, return the original string
      return val;
    }
  }
  return val;
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

        // Read the workbook. 
        // Note: For older XLS files, if the encoding isn't detected correctly, 
        // characters might be garbled. We will fix this in the normalization step.
        const workbook = XLSX.read(data, { 
          type: 'array',
          cellDates: true 
        });
        
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        // Parse to JSON
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
        
        // Handle filename modernization
        let finalFileName = file.name;
        if (finalFileName.toLowerCase().endsWith('.xls') && !finalFileName.toLowerCase().endsWith('.xlsx')) {
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
        // 1. Trim whitespace from keys
        // 2. Repair encoding for both keys and values
        const normalizedData = jsonData.map((row: any) => {
          const newRow: any = {};
          Object.keys(row).forEach(key => {
            // Repair the key (header) name
            const repairedKey = repairEncoding(key).trim();
            // Repair the value content
            const repairedValue = repairEncoding(row[key]);
            newRow[repairedKey] = repairedValue;
          });
          return newRow;
        });

        // Extract headers from the first row of normalized data
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