import * as XLSX from 'xlsx';
import { ProcessedFile } from '../types';

// Map of Windows-1252 characters (0x80-0x9F) that map to Unicode > 255.
// These often appear when Big5 bytes are interpreted as CP1252.
const CP1252_REV_MAP: { [key: number]: number } = {
  8364: 128, // €
  8218: 130, // ‚
  402: 131,  // ƒ
  8222: 132, // „
  8230: 133, // …
  8224: 134, // †
  8225: 135, // ‡
  710: 136,  // ˆ
  8240: 137, // ‰
  352: 138,  // Š
  8249: 139, // ‹
  338: 140,  // Œ
  381: 142,  // Ž
  8216: 145, // ‘
  8217: 146, // ’
  8220: 147, // “
  8221: 148, // ”
  8226: 149, // •
  8211: 150, // –
  8212: 151, // —
  732: 152,  // ˜
  8482: 153, // ™
  353: 154,  // š
  8250: 155, // ›
  339: 156,  // œ
  382: 158,  // ž
  376: 159   // Ÿ
};

/**
 * Utility to repair garbled Chinese characters.
 * Handles cases where Big5 bytes are interpreted as Latin-1 (ISO-8859-1) or Windows-1252.
 */
const repairEncoding = (val: any): any => {
  if (typeof val !== 'string' || val.length === 0) return val;
  
  // 1. If it contains valid CJK Unified Ideographs (Common Chinese chars), 
  // it is likely already correct. We trust it.
  if (/[\u4E00-\u9FFF]/.test(val)) return val;

  // 2. If strictly ASCII (0-127), no repair needed.
  // Big5 lead bytes are always > 0x80.
  let isAscii = true;
  for (let i = 0; i < val.length; i++) {
    if (val.charCodeAt(i) > 127) {
      isAscii = false;
      break;
    }
  }
  if (isAscii) return val;

  // 3. Attempt to convert from "Garbage" (Latin-1/CP1252) back to Raw Bytes.
  // We handle both ISO-8859-1 (1:1 mapping for < 256) and CP1252 specific chars.
  const bytes = new Uint8Array(val.length);
  for (let i = 0; i < val.length; i++) {
    const code = val.charCodeAt(i);
    
    if (code < 256) {
      bytes[i] = code;
    } else if (CP1252_REV_MAP[code]) {
      // It's a CP1252 special char (like € or š), map back to its byte value (128-159)
      bytes[i] = CP1252_REV_MAP[code];
    } else {
      // Encountered a high-unicode char that is NOT a CP1252 artifact.
      // This means the string contains other valid Unicode (e.g. Emoji, Greek, etc).
      // We should probably stop trying to repair it to avoid breaking valid text.
      return val;
    }
  }

  // 4. Decode the bytes as Big5
  try {
    const decoder = new TextDecoder('big5', { fatal: false });
    const decoded = decoder.decode(bytes);
    
    // 5. Validation Logic
    // If the decode process produced replacement characters (\ufffd), it means the bytes 
    // were not valid Big5 sequences. This likely wasn't Big5 data.
    // However, if the ORIGINAL also had replacement chars, we might just be passing them through.
    if (decoded.includes('\ufffd') && !val.includes('\ufffd')) {
      return val;
    }

    // CRITICAL: Only accept the repair if we found at least one valid CJK character.
    // This confirms we successfully uncovered the hidden Chinese text.
    if (/[\u4E00-\u9FFF]/.test(decoded)) {
        return decoded;
    }
    
    // If no Chinese found, revert to original.
    // This prevents turning "Café" (Latin-1) into random garbage.
    return val;

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
        
        // We force codepage 950 (Big5) for best effort on legacy files.
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
        
        // Parse to JSON
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
        
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

        // Normalize data
        const normalizedData = jsonData.map((row: any) => {
          const newRow: any = {};
          Object.keys(row).forEach(key => {
            // Apply repair to both Keys and Values
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