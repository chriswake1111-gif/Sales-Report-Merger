import * as XLSX from 'xlsx';
import { ProcessedFile } from '../types';

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

        const workbook = XLSX.read(data, { type: 'array' });
        
        // Assume data is in the first sheet
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        // Parse to JSON to get data and headers
        const jsonData = XLSX.utils.sheet_to_json(worksheet);
        
        if (jsonData.length === 0) {
           resolve({
            id: crypto.randomUUID(),
            name: file.name,
            size: file.size,
            rowCount: 0,
            headers: [],
            data: [],
            error: "No data found in the first sheet"
          });
          return;
        }

        // Extract headers from the first row object keys
        const headers = Object.keys(jsonData[0] as object);

        resolve({
          id: crypto.randomUUID(),
          name: file.name,
          size: file.size,
          rowCount: jsonData.length,
          headers,
          data: jsonData,
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
  // Combine all data arrays
  const allData = files.flatMap(file => file.data);

  // Sort based on the provided key (default: 單號)
  // We attempt to handle both string and number sorting
  return allData.sort((a, b) => {
    const valA = a[sortKey];
    const valB = b[sortKey];

    if (valA === valB) return 0;
    
    // Handle undefined/nulls (push to bottom)
    if (valA === undefined || valA === null) return 1;
    if (valB === undefined || valB === null) return -1;

    // Numeric sort
    if (typeof valA === 'number' && typeof valB === 'number') {
      return valA - valB;
    }

    // String sort
    return String(valA).localeCompare(String(valB));
  });
};

/**
 * Exports data to an XLSX file
 */
export const exportToExcel = (data: any[], filename: string) => {
  // Create a new workbook
  const worksheet = XLSX.utils.json_to_sheet(data);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Merged Data");

  // Ensure filename ends with .xlsx
  const safeFilename = filename.endsWith('.xlsx') ? filename : `${filename}.xlsx`;

  // Write and download
  XLSX.writeFile(workbook, safeFilename);
};
