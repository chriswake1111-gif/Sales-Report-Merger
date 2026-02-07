import React, { useState } from 'react';
import { Command } from '@tauri-apps/plugin-shell';
import { Terminal } from 'lucide-react';

const TauriDebug: React.FC = () => {
    const [output, setOutput] = useState<string>('');
    const [isLoading, setIsLoading] = useState(false);

    const testSidecar = async () => {
        setIsLoading(true);
        setOutput('Running sidecar...');
        try {
            // 呼叫我們在 tauri.conf.json 定義的外部執行檔
            // 注意：這裡不傳參數，預期 Python 會回傳 Error: No file path，證明執行成功
            const command = Command.sidecar('bin/excel_parser');
            const result = await command.execute();

            console.log('Sidecar Result:', result);

            if (result.stdout) {
                setOutput(`Success (stdout): ${result.stdout}`);
            } else if (result.stderr) {
                setOutput(`Output (stderr): ${result.stderr}`); // Python print to stderr for logs?
            } else {
                // It might have exited with code 1, which command.execute() catches? 
                // No, execute() resolves with {code, stdout, stderr}
                setOutput(`Exit Code: ${result.code}\nStdout: ${result.stdout}\nStderr: ${result.stderr}`);
            }

        } catch (e: any) {
            console.error(e);
            setOutput(`Error: ${e.message || e}`);
        } finally {
            setIsLoading(false);
        }
    };

    // 簡單檢查我們是否在 Tauri 環境 (需靠 window.__TAURI_INTERNALS__ 或類似機制，或只要不崩潰就好)
    // 這裡假設如果沒有 Tauri 環境，呼叫 Command.sidecar 可能會直接報錯，但不會讓 Component 渲染失敗

    return (
        <div className="bg-orange-50 border border-orange-200 rounded-lg p-4 mt-8">
            <h3 className="text-lg font-semibold text-orange-800 flex items-center mb-2">
                <Terminal size={20} className="mr-2" />
                Tauri Sidecar Debugger
            </h3>
            <p className="text-sm text-orange-600 mb-4">
                此區域僅供測試 Tauri Python Sidecar 整合。
                <br />
                點擊測試按鈕後，應顯示類似 <code>{`{"success": false, "error": "No file path"}`}</code> 的訊息，代表 Python 已成功執行。
            </p>

            <button
                onClick={testSidecar}
                disabled={isLoading}
                className="px-4 py-2 bg-orange-600 text-white rounded hover:bg-orange-700 disabled:opacity-50 transition-colors"
            >
                {isLoading ? 'Running...' : 'Test Python Sidecar'}
            </button>

            {output && (
                <div className="mt-4 p-3 bg-gray-900 text-green-400 font-mono text-xs rounded overflow-auto max-h-40 whitespace-pre-wrap">
                    {output}
                </div>
            )}
        </div>
    );
};

export default TauriDebug;
