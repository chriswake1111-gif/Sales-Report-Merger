import React from 'react';
import { FileSpreadsheet, Trash2, AlertCircle } from 'lucide-react';
import { ProcessedFile } from '../types';

interface FileListProps {
  files: ProcessedFile[];
  onRemove: (id: string) => void;
}

const FileList: React.FC<FileListProps> = ({ files, onRemove }) => {
  if (files.length === 0) return null;

  return (
    <div className="mt-8 space-y-4">
      <h3 className="text-lg font-semibold text-gray-800 flex items-center justify-between">
        <span>已匯入檔案 ({files.length})</span>
        <span className="text-sm font-normal text-gray-500">
          總筆數: <span className="font-bold text-blue-600">{files.reduce((acc, f) => acc + f.rowCount, 0).toLocaleString()}</span>
        </span>
      </h3>
      
      <div className="bg-white rounded-lg shadow-sm border border-gray-200 overflow-hidden">
        <div className="grid grid-cols-12 gap-4 p-3 bg-gray-50 border-b border-gray-200 text-xs font-semibold text-gray-500 uppercase tracking-wider">
          <div className="col-span-6">檔案名稱</div>
          <div className="col-span-2 text-right">大小</div>
          <div className="col-span-3 text-right">資料筆數</div>
          <div className="col-span-1 text-center">操作</div>
        </div>
        <ul className="divide-y divide-gray-100">
          {files.map((file) => (
            <li key={file.id} className="grid grid-cols-12 gap-4 p-4 items-center hover:bg-gray-50 transition-colors">
              <div className="col-span-6 flex items-center space-x-3 overflow-hidden">
                <div className={`p-2 rounded-lg ${file.error ? 'bg-red-100 text-red-600' : 'bg-green-100 text-green-600'}`}>
                  <FileSpreadsheet size={20} />
                </div>
                <div className="truncate">
                  <p className="font-medium text-gray-900 truncate" title={file.name}>{file.name}</p>
                  {file.error && (
                     <p className="text-xs text-red-500 flex items-center mt-1">
                       <AlertCircle size={12} className="mr-1"/> {file.error}
                     </p>
                  )}
                </div>
              </div>
              <div className="col-span-2 text-right text-sm text-gray-500">
                {(file.size / 1024).toFixed(1)} KB
              </div>
              <div className="col-span-3 text-right font-medium text-gray-900">
                {file.rowCount.toLocaleString()} 筆
              </div>
              <div className="col-span-1 text-center">
                <button
                  onClick={() => onRemove(file.id)}
                  className="text-gray-400 hover:text-red-500 p-1 rounded-full hover:bg-red-50 transition-colors"
                  title="移除此檔案"
                >
                  <Trash2 size={18} />
                </button>
              </div>
            </li>
          ))}
        </ul>
      </div>
    </div>
  );
};

export default FileList;
