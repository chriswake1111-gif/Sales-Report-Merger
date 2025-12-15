import React, { useCallback } from 'react';
import { UploadCloud } from 'lucide-react';

interface DropzoneProps {
  onFilesSelected: (files: File[]) => void;
  isProcessing: boolean;
}

const Dropzone: React.FC<DropzoneProps> = ({ onFilesSelected, isProcessing }) => {
  const handleDrop = useCallback(
    (e: React.DragEvent<HTMLDivElement>) => {
      e.preventDefault();
      e.stopPropagation();
      if (isProcessing) return;

      const droppedFiles = Array.from(e.dataTransfer.files).filter(
        (file: File) => file.name.endsWith('.xls') || file.name.endsWith('.xlsx')
      );
      if (droppedFiles.length > 0) {
        onFilesSelected(droppedFiles);
      }
    },
    [onFilesSelected, isProcessing]
  );

  const handleDragOver = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    e.stopPropagation();
  };

  const handleFileInput = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      const selectedFiles = Array.from(e.target.files);
      onFilesSelected(selectedFiles);
    }
  };

  return (
    <div
      onDrop={handleDrop}
      onDragOver={handleDragOver}
      className={`border-2 border-dashed rounded-xl p-10 text-center transition-colors cursor-pointer ${
        isProcessing
          ? 'border-gray-300 bg-gray-50 cursor-not-allowed opacity-50'
          : 'border-blue-300 hover:border-blue-500 hover:bg-blue-50 bg-white'
      }`}
      onClick={() => !isProcessing && document.getElementById('fileInput')?.click()}
    >
      <input
        type="file"
        id="fileInput"
        multiple
        accept=".xls,.xlsx"
        className="hidden"
        onChange={handleFileInput}
        disabled={isProcessing}
      />
      <div className="flex flex-col items-center justify-center space-y-4">
        <div className="p-4 bg-blue-100 rounded-full text-blue-600">
          <UploadCloud size={40} />
        </div>
        <div>
          <h3 className="text-lg font-semibold text-gray-700">
            點擊或拖曳檔案至此
          </h3>
          <p className="text-sm text-gray-500 mt-1">
            支援 .xls 與 .xlsx 格式銷售報表
          </p>
        </div>
      </div>
    </div>
  );
};

export default Dropzone;