const {contextBridge,ipcRenderer,webUtils}=require('electron');
contextBridge.exposeInMainWorld('desktopAPI',{
 inspectMailFile:p=>ipcRenderer.invoke('inspect-mail-file',p),
 saveDroppedFile:p=>ipcRenderer.invoke('save-dropped-file',p),
 openAttachment:p=>ipcRenderer.invoke('open-attachment',p),
 getPathForFile:file=>webUtils.getPathForFile(file)
});
