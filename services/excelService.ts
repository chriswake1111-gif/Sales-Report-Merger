import * as XLSX from 'xlsx';
import { ProcessedFile } from '../types';

/* load 'cpexcel' for codepages */
// @ts-ignore
import * as cpexcel from 'xlsx/dist/cpexcel.full.mjs';

// Monkey patch: Map 164 and 176 to 950 (Big5)
let cptable: any = {};

try {
  // Handle ESM default export options
  // @ts-ignore
  const rawCpexcel = cpexcel;

  // Clone to ensure extensibility
  cptable = { ...rawCpexcel };

  // Check for the utils.decode function (the core of all translations)
  if (cptable.utils && typeof cptable.utils.decode === 'function') {
    // Clone utils to avoid mutating frozen objects
    cptable.utils = { ...cptable.utils };
    const originalDecode = cptable.utils.decode;

    // INTERCEPTOR: Redirect 164/176/168 -> 950
    cptable.utils.decode = function (cp: number, data: any) {
      const table = cptable.cptable || cptable;
      const has950 = !!(table[950] || table['950']);

      if (cp === 164 || cp === 176 || cp === 168) {
        if (!has950) {
          console.error(`Sales-Merger Error: Redirecting CP${cp} to CP950 failed because CP950 is MISSING! Available keys count: ${Object.keys(table).length}`);
        }
        // Redirection for legacy pharmacy/system reports
        return originalDecode.call(this, 950, data);
      }
      return originalDecode.apply(this, arguments as any);
    };
    console.log("Sales-Merger: Successfully installed interceptor for CP164/176/168 -> CP950.");
    console.log("Sales-Merger: Initial CP950 check:", !!((cptable.cptable || cptable)[950]));
  } else {
    console.warn("Sales-Merger: Could not find utils.decode to patch!");
  }
} catch (e) {
  console.error("Sales-Merger: Failed to patch codepages:", e);
}

XLSX.set_cptable(cptable);

declare global {
  interface Window {
    electronAPI?: {
      parseExcel: (filePath: string) => Promise<any>;
      getPath: (file: File) => string;
      isElectron: boolean;
    };
  }
}

/**
 * Reads a file and parses it into JSON
 */
export const parseExcelFile = async (file: File): Promise<ProcessedFile> => {
  // Check if we are in Electron and get the physical file path
  let electronPath = (file as any).path;

  if (window.electronAPI && window.electronAPI.getPath) {
    try {
      electronPath = window.electronAPI.getPath(file);
    } catch (e) {
      console.warn("Sales-Merger: Failed to get path via getPath API, trying property:", e);
    }
  }

  console.log("Sales-Merger Debug: electronAPI present:", !!window.electronAPI);
  console.log("Sales-Merger Debug: file.path present:", !!electronPath);
  if (electronPath) console.log("Sales-Merger Debug: file.path content:", electronPath);

  if (window.electronAPI && electronPath) {
    try {
      console.log("Using Electron Python parser for:", electronPath);
      const result = await window.electronAPI.parseExcel(electronPath);

      if (!result.success) {
        if (result.rawOutput) console.error("Sales-Merger Debug: Raw Python Output:", result.rawOutput);
        if (result.stderr) console.error("Sales-Merger Debug: Python Stderr:", result.stderr);
        throw new Error(result.error || "Failed to parse file via Electron");
      }

      let fileNameLower = file.name.toLowerCase();
      let finalFileName = file.name;
      if (fileNameLower.endsWith('.xls') || fileNameLower.endsWith('.csv')) {
        finalFileName = finalFileName.replace(/\.(xls|csv)$/i, '.xlsx');
      }

      return {
        id: crypto.randomUUID(),
        name: finalFileName,
        size: file.size,
        rowCount: result.rowCount || result.data.length,
        headers: result.headers || (result.data.length > 0 ? Object.keys(result.data[0]) : []),
        data: result.data,
      };
    } catch (err) {
      console.error("Electron parsing failed, falling back to browser:", err);
      // Fallback to browser parsing below
    }
  }

  // BROWSER FALLBACK
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        const data = e.target?.result;
        if (!data) throw new Error("File is empty");

        const fileNameLower = file.name.toLowerCase();

        const workbook = XLSX.read(data, {
          type: 'array',
          cellDates: true,
          cellFormula: false,
          cellStyles: false,
          cellNF: true,
          codepage: 950
        });

        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];

        const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: "", raw: false });

        let finalFileName = file.name;
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

        const normalizedData = jsonData.map((row: any) => {
          const newRow: Record<string, any> = {};
          Object.keys(row).forEach(key => {
            const cleanKey = key.trim();
            const val = row[key];
            newRow[cleanKey] = typeof val === 'string' ? val.trim() : val;
          });
          return newRow;
        });

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
        console.error("Browser parsing error:", err);
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
