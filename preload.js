const {contextBridge,ipcRenderer,webUtils}=require('electron');

contextBridge.exposeInMainWorld('desktopAPI',{
 inspectMailFile:p=>ipcRenderer.invoke('inspect-mail-file',p),
 saveDroppedFile:p=>ipcRenderer.invoke('save-dropped-file',p),
 openAttachment:p=>ipcRenderer.invoke('open-attachment',p),
 getPathForFile:file=>webUtils.getPathForFile(file)
});

// Electron 32+ removed File.path. The existing renderer uses file.path,
// so provide a compatibility getter backed by webUtils.getPathForFile().
try{
 contextBridge.executeInMainWorld({
  func:()=>{
   if(!Object.getOwnPropertyDescriptor(File.prototype,'path')){
    Object.defineProperty(File.prototype,'path',{configurable:true,get(){
     return window.desktopAPI.getPathForFile(this);
    }});
   }
  }
 });
}catch(err){console.error('File.path compatibility:',err);}
