const {contextBridge,ipcRenderer}=require('electron');
contextBridge.exposeInMainWorld('desktopAPI',{
 inspectMailFile:p=>ipcRenderer.invoke('inspect-mail-file',p),
 saveDroppedFile:p=>ipcRenderer.invoke('save-dropped-file',p),
 openAttachment:p=>ipcRenderer.invoke('open-attachment',p)
});