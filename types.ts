export interface ProcessedFile {
  id: string;
  name: string;
  size: number;
  rowCount: number;
  headers: string[];
  data: any[];
  error?: string;
}

export interface MergeConfig {
  sortColumn: string;
  outputFilename: string;
}

export type MergeStatus = 'idle' | 'processing' | 'success' | 'error';
