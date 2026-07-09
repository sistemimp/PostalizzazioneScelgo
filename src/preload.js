const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('postaMassivaApi', {
  selectInputFile: () => ipcRenderer.invoke('select-input-file'),
  selectOutputDirectory: () => ipcRenderer.invoke('select-output-directory'),
  getExcelInfo: (filePath) => ipcRenderer.invoke('get-excel-info', filePath),
  processWorkbook: (options) => ipcRenderer.invoke('process-workbook', options),
  openPath: (targetPath) => ipcRenderer.invoke('open-path', targetPath)
});
