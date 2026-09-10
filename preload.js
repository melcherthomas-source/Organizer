const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('desktopAPI', {
  saveDroppedFile: (payload) => ipcRenderer.invoke('save-dropped-file', payload),
  openAttachment: (filePath) => ipcRenderer.invoke('open-attachment', filePath)
});
