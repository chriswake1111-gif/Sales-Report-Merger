import React, { useState, useEffect } from 'react';
import { Download, Layers, FileOutput, RefreshCw, AlertTriangle } from 'lucide-react';
import Dropzone from './components/Dropzone';
import FileList from './components/FileList';
import { ProcessedFile, MergeStatus } from './types';
import { parseExcelFile, mergeData, exportToExcel } from './services/excelService';

const App: React.FC = () => {
  const [files, setFiles] = useState<ProcessedFile[]>([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [sortColumn, setSortColumn] = useState('單號');
  const [outputFilename, setOutputFilename] = useState('');
  const [mergeStatus, setMergeStatus] = useState<MergeStatus>('idle');
  const [totalRows, setTotalRows] = useState(0);

  // Auto-generate a default filename based on current date
  useEffect(() => {
    const today = new Date();
    const dateStr = today.toISOString().split('T')[0];
    setOutputFilename(`合併銷售報表_${dateStr}`);
  }, []);

  // Update total rows whenever files change
  useEffect(() => {
    const total = files.reduce((acc, file) => acc + file.rowCount, 0);
    setTotalRows(total);
  }, [files]);

  const handleFilesSelected = async (newFiles: File[]) => {
    setIsProcessing(true);
    setMergeStatus('idle');
    
    try {
      const processedPromises = newFiles.map(parseExcelFile);
      const processedResults = await Promise.all(processedPromises);
      
      setFiles(prev => [...prev, ...processedResults]);
    } catch (error) {
      console.error("Error processing files:", error);
      alert("讀取檔案時發生錯誤，請確認檔案格式是否正確。");
    } finally {
      setIsProcessing(false);
    }
  };

  const removeFile = (id: string) => {
    setFiles(prev => prev.filter(f => f.id !== id));
    setMergeStatus('idle');
  };

  const handleMergeAndDownload = () => {
    if (files.length === 0) return;
    
    setIsProcessing(true);
    setMergeStatus('processing');

    const targetColumn = sortColumn.trim();

    // Small timeout to allow UI to update to 'processing' state
    setTimeout(() => {
      try {
        // Validate sort column exists in at least one file
        // We trim the headers in the file service, so we trim the input here too
        const hasColumn = files.some(f => f.headers.includes(targetColumn));
        
        if (!hasColumn) {
          // Get a list of available headers from the first file to help user debug
          const availableHeaders = files[0]?.headers.slice(0, 5).join(', ') + (files[0]?.headers.length > 5 ? '...' : '');
          
          const proceed = window.confirm(
            `警告：在匯入的檔案中找不到「${targetColumn}」欄位。\n\n` +
            `偵測到的欄位範例：${availableHeaders}\n\n` +
            `合併將繼續，但排序可能無法正常運作。是否繼續？`
          );
          
          if (!proceed) {
            setIsProcessing(false);
            setMergeStatus('idle');
            return;
          }
        }

        const merged = mergeData(files, targetColumn);
        exportToExcel(merged, outputFilename);
        setMergeStatus('success');
      } catch (error) {
        console.error("Merge error:", error);
        setMergeStatus('error');
        alert("合併過程中發生錯誤");
      } finally {
        setIsProcessing(false);
      }
    }, 100);
  };

  const handleReset = () => {
    if (window.confirm('確定要清空所有已匯入的檔案嗎？')) {
      setFiles([]);
      setMergeStatus('idle');
    }
  };

  return (
    <div className="min-h-screen bg-gray-50 py-12 px-4 sm:px-6 lg:px-8 font-sans">
      <div className="max-w-4xl mx-auto">
        {/* Header */}
        <div className="text-center mb-10">
          <div className="flex justify-center mb-4">
            <div className="bg-blue-600 p-3 rounded-2xl shadow-lg">
              <Layers className="text-white h-8 w-8" />
            </div>
          </div>
          <h1 className="text-3xl font-bold text-gray-900 mb-2">銷售報表合併工具</h1>
          <p className="text-gray-600">
            輕鬆將多個 .xls/.xlsx 報表合併為一份，自動依「{sortColumn}」排序。
          </p>
        </div>

        {/* Main Card */}
        <div className="bg-white rounded-2xl shadow-xl overflow-hidden border border-gray-100">
          <div className="p-8">
            
            {/* 1. Upload Section */}
            <div className="mb-8">
              <h2 className="text-lg font-semibold text-gray-800 mb-4 flex items-center">
                <span className="bg-blue-100 text-blue-800 text-xs font-bold px-2 py-1 rounded mr-2">Step 1</span>
                匯入報表
              </h2>
              <Dropzone onFilesSelected={handleFilesSelected} isProcessing={isProcessing} />
            </div>

            {/* File List */}
            <FileList files={files} onRemove={removeFile} />

            {/* 2. Configuration Section */}
            {files.length > 0 && (
              <div className="mt-10 pt-8 border-t border-gray-100">
                <h2 className="text-lg font-semibold text-gray-800 mb-4 flex items-center">
                  <span className="bg-blue-100 text-blue-800 text-xs font-bold px-2 py-1 rounded mr-2">Step 2</span>
                  合併設定與匯出
                </h2>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-6 mb-6">
                  <div>
                    <label htmlFor="filename" className="block text-sm font-medium text-gray-700 mb-2">
                      匯出檔案名稱
                    </label>
                    <div className="relative rounded-md shadow-sm">
                      <div className="absolute inset-y-0 left-0 pl-3 flex items-center pointer-events-none">
                        <FileOutput size={16} className="text-gray-400" />
                      </div>
                      <input
                        type="text"
                        id="filename"
                        className="focus:ring-blue-500 focus:border-blue-500 block w-full pl-10 sm:text-sm border-gray-300 rounded-md py-2 border px-3"
                        placeholder="輸入檔案名稱"
                        value={outputFilename}
                        onChange={(e) => setOutputFilename(e.target.value)}
                      />
                      <div className="absolute inset-y-0 right-0 pr-3 flex items-center pointer-events-none">
                        <span className="text-gray-500 sm:text-sm">.xlsx</span>
                      </div>
                    </div>
                  </div>

                  <div>
                    <label htmlFor="sortCol" className="block text-sm font-medium text-gray-700 mb-2">
                      排序依據欄位
                    </label>
                    <input
                      type="text"
                      id="sortCol"
                      className="focus:ring-blue-500 focus:border-blue-500 block w-full sm:text-sm border-gray-300 rounded-md py-2 border px-3"
                      value={sortColumn}
                      onChange={(e) => setSortColumn(e.target.value)}
                      placeholder="例如: 單號"
                    />
                    <p className="mt-1 text-xs text-gray-500">系統將依此欄位進行升冪排序</p>
                  </div>
                </div>

                {/* Validation Info */}
                 <div className="bg-blue-50 border border-blue-100 rounded-lg p-4 mb-6 flex items-start">
                    <div className="mr-3 mt-0.5 text-blue-600">
                        <AlertTriangle size={20} />
                    </div>
                    <div className="text-sm text-blue-800">
                        <p className="font-semibold mb-1">合併預覽確認：</p>
                        <ul className="list-disc pl-4 space-y-1">
                            <li>將合併 <strong>{files.length}</strong> 份文件。</li>
                            <li>預計產出總筆數：<strong>{totalRows.toLocaleString()}</strong> 筆資料。</li>
                            <li>輸出格式為 <strong>.xlsx</strong> (Excel 活頁簿)。</li>
                        </ul>
                    </div>
                </div>

                {/* Actions */}
                <div className="flex items-center justify-end space-x-4">
                  <button
                    onClick={handleReset}
                    className="flex items-center px-4 py-2 border border-gray-300 shadow-sm text-sm font-medium rounded-md text-gray-700 bg-white hover:bg-gray-50 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-blue-500"
                  >
                    <RefreshCw size={16} className="mr-2" />
                    清空重來
                  </button>
                  <button
                    onClick={handleMergeAndDownload}
                    disabled={isProcessing || files.length === 0}
                    className={`flex items-center px-6 py-2 border border-transparent text-base font-medium rounded-md shadow-sm text-white focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-blue-500 ${
                      isProcessing || files.length === 0
                        ? 'bg-blue-400 cursor-not-allowed'
                        : 'bg-blue-600 hover:bg-blue-700'
                    }`}
                  >
                    {isProcessing ? (
                      <>處理中...</>
                    ) : (
                      <>
                        <Download size={18} className="mr-2" />
                        合併並下載
                      </>
                    )}
                  </button>
                </div>
              </div>
            )}
          </div>
          
          {/* Footer Status Bar */}
          <div className="bg-gray-50 px-8 py-4 border-t border-gray-200">
             <div className="flex justify-between items-center text-sm">
                <span className="text-gray-500">支援舊版 .xls 與新版 .xlsx 格式</span>
                {mergeStatus === 'success' && (
                    <span className="text-green-600 font-medium flex items-center">
                        ✓ 下載已開始
                    </span>
                )}
             </div>
          </div>
        </div>
      </div>
    </div>
  );
};

export default App;