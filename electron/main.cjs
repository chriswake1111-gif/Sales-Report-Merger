const { app, BrowserWindow, ipcMain } = require('electron');
const path = require('path');
const { spawn } = require('child_process');
const fs = require('fs');

function createWindow() {
    const win = new BrowserWindow({
        width: 1200,
        height: 800,
        webPreferences: {
            preload: path.join(__dirname, 'preload.cjs'),
            contextIsolation: true,
            nodeIntegration: false,
        },
        icon: path.join(__dirname, '../public/pwa-512x512.png'),
        title: "銷售報表合併工具 (桌面版)"
    });

    // Use app.isPackaged to detect dev vs prod
    if (!app.isPackaged) {
        const loadDevServer = () => {
            console.log("Attempting to connect to Vite dev server...");
            win.loadURL('http://localhost:5173').catch((e) => {
                console.log('Dev server not ready, retrying in 1s...', e.message);
                setTimeout(loadDevServer, 1000);
            });
        };
        loadDevServer();
        win.webContents.openDevTools();
    } else {
        win.loadFile(path.join(__dirname, '../dist/index.html'));
    }
}

app.whenReady().then(createWindow);

app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') app.quit();
});

// IPC Handler for Excel Parsing using Python
ipcMain.handle('parse-excel', async (event, filePath) => {
    return new Promise((resolve, reject) => {
        // Find python executable (standard 'python' or 'python3')
        const pythonPath = process.platform === 'win32' ? 'python' : 'python3';
        const scriptPath = path.join(__dirname, '../scripts/excel_parser.py');
        // Security: Validate file path and extension
        if (!filePath || typeof filePath !== 'string') {
            return resolve({ success: false, error: 'Invalid file path' });
        }
        const ext = path.extname(filePath).toLowerCase();
        if (!['.xls', '.xlsx', '.csv'].includes(ext)) {
            return resolve({ success: false, error: 'Unsupported file type' });
        }

        console.log(`Running Python parser for: ${filePath}`);
        const pyProg = spawn(pythonPath, [scriptPath, filePath]);
        let dataString = '';
        let errorString = '';

        pyProg.stdout.on('data', (data) => {
            dataString += data.toString();
        });

        pyProg.stderr.on('data', (data) => {
            errorString += data.toString();
        });

        pyProg.on('close', (code) => {
            if (code !== 0) {
                console.error(`Python script exited with code ${code}: ${errorString}`);
                return resolve({ success: false, error: errorString || `Exited with code ${code}` });
            }
            try {
                // Robust JSON extraction: Find the first { and last }
                // This prevents issues if Python packages print extra stuff to stdout
                const firstBrace = dataString.indexOf('{');
                const lastBrace = dataString.lastIndexOf('}');

                if (firstBrace === -1 || lastBrace === -1) {
                    throw new Error('No JSON object found in output');
                }

                const jsonString = dataString.substring(firstBrace, lastBrace + 1);
                const result = JSON.parse(jsonString);
                resolve(result);
            } catch (e) {
                console.error('Failed to parse Python output:', e);
                console.error('Raw Python Output (Full):', dataString);
                console.error('Python Stderr:', errorString);
                resolve({
                    success: false,
                    error: 'Invalid JSON response from parser',
                    rawOutput: dataString,
                    stderr: errorString
                });
            }
        });
    });
});

// IPC Handler to save file (dialog or direct)
ipcMain.handle('save-file-dialog', async (event, { defaultPath, data }) => {
    // For now, Front-end handles Blob saving. 
    // In full Electron mode we could use dialog.showSaveDialogSync
    return true;
});
