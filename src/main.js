const path = require('node:path');
const fs = require('node:fs');
const { app, BrowserWindow, dialog, ipcMain, shell } = require('electron');
const { getWorkbookInfo } = require('./lib/excelInfo');
const { processWorkbook } = require('./lib/processWorkbook');

function resolveDocsDir() {
  const candidate = path.resolve(__dirname, '..', 'docs');
  if (fs.existsSync(candidate)) {
    return candidate;
  }

  return path.resolve(process.cwd(), 'docs');
}

function createMainWindow() {
  const appIconPath = path.resolve(__dirname, '..', 'icon.ico');
  const mainWindow = new BrowserWindow({
    width: 1100,
    height: 860,
    minWidth: 960,
    minHeight: 720,
    icon: appIconPath,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false
    }
  });

  mainWindow.loadFile(path.join(__dirname, 'renderer', 'index.html'));
}

ipcMain.handle('select-input-file', async () => {
  const result = await dialog.showOpenDialog({
    properties: ['openFile'],
    filters: [{ name: 'Excel', extensions: ['xlsx', 'xls', 'xlsm'] }]
  });

  if (result.canceled || result.filePaths.length === 0) {
    return { canceled: true };
  }

  return { canceled: false, filePath: result.filePaths[0] };
});

ipcMain.handle('select-output-directory', async () => {
  const result = await dialog.showOpenDialog({
    properties: ['openDirectory']
  });

  if (result.canceled || result.filePaths.length === 0) {
    return { canceled: true };
  }

  return { canceled: false, directoryPath: result.filePaths[0] };
});

ipcMain.handle('get-excel-info', async (_, filePath) => {
  return getWorkbookInfo(filePath);
});

ipcMain.handle('process-workbook', async (_, userOptions) => {
  const options = {
    ...userOptions,
    docsDir: resolveDocsDir()
  };

  return processWorkbook(options);
});

ipcMain.handle('open-path', async (_, targetPath) => {
  if (!targetPath) {
    return;
  }
  await shell.showItemInFolder(targetPath);
});

app.whenReady().then(() => {
  createMainWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createMainWindow();
    }
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});
