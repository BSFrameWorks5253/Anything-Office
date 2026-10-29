const { app, BrowserWindow, Menu, ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs');
const { resolveChildAppLoaderConfig, showChildAppSplash } = require('./splash.cjs');

// Optimize Chromium startup performance
app.commandLine.appendSwitch('enable-gpu-rasterization');
app.commandLine.appendSwitch('enable-zero-copy');

// Strip default Electron application menu (File, Edit, View, Window, Help)
Menu.setApplicationMenu(null);

const APP_ID = 'anything-office';
const APP_NAME = 'Anything Office';

let mainWindow = null;
let ipcRegistered = false;

function createMainWindow() {
  mainWindow = new BrowserWindow({
    width: 1400,
    height: 900,
    minWidth: 1024,
    minHeight: 700,
    title: APP_NAME,
    frame: false, // Completely removes Windows default caption, titlebar, and border
    autoHideMenuBar: true, // Completely hides and prevents Windows menu bar
    backgroundColor: '#0f172a',
    show: false, // Hidden while splash displays or until content is ready
    webPreferences: {
      preload: path.join(__dirname, 'preload.cjs'),
      nodeIntegration: false,
      contextIsolation: true,
      sandbox: false,
      spellcheck: false, // Accelerates startup on Windows
      backgroundThrottling: false,
    },
  });

  // Explicitly strip menu bar at window level
  mainWindow.setMenuBarVisibility(false);
  if (typeof mainWindow.removeMenu === 'function') {
    mainWindow.removeMenu();
  }

  const prodIndex = path.join(__dirname, '../dist/index.html');
  const devUrl = 'http://localhost:5173';

  // Load production bundle for sub-second offline startup, fallback to dev server
  if (fs.existsSync(prodIndex)) {
    mainWindow.loadFile(prodIndex);
  } else {
    mainWindow.loadURL(devUrl);
  }

  // Debug logging
  mainWindow.webContents.on('did-fail-load', (event, errorCode, errorDescription, validatedURL) => {
    console.error('[Electron] Failed to load:', errorCode, errorDescription, validatedURL);
  });

  if (!ipcRegistered) {
    registerIpcHandlers();
    ipcRegistered = true;
  }

  return mainWindow;
}

function registerIpcHandlers() {
  // Window controls
  ipcMain.on('window-minimize', () => {
    if (mainWindow && !mainWindow.isDestroyed()) {
      mainWindow.minimize();
    }
  });

  ipcMain.on('window-maximize', () => {
    if (mainWindow && !mainWindow.isDestroyed()) {
      if (mainWindow.isMaximized()) {
        mainWindow.unmaximize();
      } else {
        mainWindow.maximize();
      }
    }
  });

  ipcMain.on('window-close', () => {
    if (mainWindow && !mainWindow.isDestroyed()) {
      mainWindow.close();
    }
  });

  // Native Open File Dialog
  ipcMain.handle('dialog-open-file', async () => {
    const targetWin = mainWindow && !mainWindow.isDestroyed() ? mainWindow : null;
    const result = await dialog.showOpenDialog(targetWin, {
      title: 'Open Document in Anything Office',
      properties: ['openFile'],
      filters: [
        {
          name: 'All Office & PDF Documents',
          extensions: ['docx', 'xlsx', 'pptx', 'pdf', 'odt', 'ods', 'odp', 'doc', 'xls', 'ppt', 'csv', 'tsv', 'anything'],
        },
        { name: 'Portable Document Format (*.pdf)', extensions: ['pdf'] },
        { name: 'Microsoft Word & LibreOffice Writer', extensions: ['docx', 'odt', 'doc', 'txt', 'md'] },
        { name: 'Microsoft Excel & LibreOffice Calc', extensions: ['xlsx', 'ods', 'xls', 'csv', 'tsv'] },
        { name: 'Microsoft PowerPoint & LibreOffice Impress', extensions: ['pptx', 'odp', 'ppt'] },
        { name: 'Anything Unified Projects', extensions: ['anything'] },
        { name: 'All Files', extensions: ['*'] },
      ],
    });

    if (result.canceled || result.filePaths.length === 0) {
      return null;
    }

    const filePath = result.filePaths[0];
    const buffer = fs.readFileSync(filePath);
    return {
      name: path.basename(filePath),
      path: filePath,
      buffer: Array.from(buffer),
    };
  });

  // Native Save File Dialog
  ipcMain.handle('dialog-save-file', async (event, { content, defaultName }) => {
    const targetWin = mainWindow && !mainWindow.isDestroyed() ? mainWindow : null;
    const result = await dialog.showSaveDialog(targetWin, {
      title: 'Save Document in Anything Office',
      defaultPath: defaultName || 'document.anything',
      filters: [
        { name: 'Anything Unified Project (.anything)', extensions: ['anything'] },
        { name: 'Microsoft Word (.docx)', extensions: ['docx'] },
        { name: 'Microsoft Excel (.xlsx)', extensions: ['xlsx'] },
        { name: 'Microsoft PowerPoint (.pptx)', extensions: ['pptx'] },
        { name: 'All Files', extensions: ['*'] },
      ],
    });

    if (result.canceled || !result.filePath) {
      return false;
    }

    fs.writeFileSync(result.filePath, content);
    return true;
  });
}

app.whenReady().then(() => {
  // 1. Resolve dynamic ecosystem loader configuration
  const loaderConfig = resolveChildAppLoaderConfig(APP_ID, APP_NAME);

  // 2. Preload main window in background while splash is displaying
  createMainWindow();

  // 3. Launch dynamic splash screen
  showChildAppSplash(loaderConfig, () => {
    // 4. Reveal main window seamlessly when splash finishes
    if (mainWindow && !mainWindow.isDestroyed()) {
      mainWindow.show();
      mainWindow.focus();
    }
  });

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createMainWindow();
      if (mainWindow && !mainWindow.isDestroyed()) {
        mainWindow.show();
        mainWindow.focus();
      }
    }
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});
