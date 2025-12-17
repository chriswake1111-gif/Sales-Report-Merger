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

        // Use more robust options for reading to handle format differences
        // cellDates: true helps standardize date handling between XLS and XLSX,
        // ensuring they are parsed as Date objects instead of potentially inconsistent serial numbers.
        const workbook = XLSX.read(data, { 
          type: 'array',
          cellDates: true 
        });
        
        // Assume data is in the first sheet
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        // Parse to JSON to get data and headers
        // defval: "" ensures that empty cells produce an empty string value instead of being undefined
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
        
        // Handle "conversion" from XLS to XLSX
        // If the original file was XLS, we treat the parsed workbook as the converted XLSX structure
        // and update the filename to reflect this modernization.
        let finalFileName = file.name;
        if (finalFileName.toLowerCase().endsWith('.xls') && !finalFileName.toLowerCase().endsWith('.xlsx')) {
          // Replace .xls (case insensitive) with .xlsx
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

        // Normalize data: Trim whitespace from keys (headers)
        // This fixes issues where Excel headers are " 單號 " instead of "單號"
        const normalizedData = jsonData.map((row: any) => {
          const newRow: any = {};
          Object.keys(row).forEach(key => {
            const cleanKey = key.trim();
            newRow[cleanKey] = row[key];
          });
          return newRow;
        });

        // Extract headers from the first row object keys of the normalized data
        const headers = Object.keys(normalizedData[0] as object);

        resolve({
          id: crypto.randomUUID(),
          name: finalFileName, // Return the converted name (.xlsx)
          size: file.size,
          rowCount: jsonData.length,
          headers,
          data: normalizedData, // Use the data with trimmed keys
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
  const cleanSortKey = sortKey.trim();

  // Sort based on the provided key (default: 單號)
  // We attempt to handle both string and number sorting
  return allData.sort((a, b) => {
    const valA = a[cleanSortKey];
    const valB = b[cleanSortKey];

    if (valA === valB) return 0;
    
    // Handle undefined/nulls (push to bottom)
    if (valA === undefined || valA === null) return 1;
    if (valB === undefined || valB === null) return -1;

    // Numeric sort (handles Dates converted to numbers too)
    // Check if both values are valid numbers (even if they are strings in JSON)
    const numA = Number(valA);
    const numB = Number(valB);

    if (!isNaN(numA) && !isNaN(numB) && valA !== '' && valB !== '') {
        return numA - numB;
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