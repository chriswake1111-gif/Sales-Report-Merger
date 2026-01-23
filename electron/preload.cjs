const { contextBridge, ipcRenderer, webUtils } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
    parseExcel: (filePath) => ipcRenderer.invoke('parse-excel', filePath),
    getPath: (file) => webUtils.getPathForFile(file),
    isElectron: true
});
