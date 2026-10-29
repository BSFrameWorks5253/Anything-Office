const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
  isElectron: true,
  minimize: () => ipcRenderer.send('window-minimize'),
  maximize: () => ipcRenderer.send('window-maximize'),
  close: () => ipcRenderer.send('window-close'),
  openNativeFileDialog: () => ipcRenderer.invoke('dialog-open-file'),
  saveNativeFileDialog: (content, defaultName) =>
    ipcRenderer.invoke('dialog-save-file', { content, defaultName }),
});
