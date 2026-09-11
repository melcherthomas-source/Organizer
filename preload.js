const {contextBridge,ipcRenderer,webUtils}=require('electron');

const droppedFiles=new Map();
async function getPathOrToken(file){
 try{
  const p=webUtils.getPathForFile(file);
  if(p)return p;
 }catch(err){}
 try{
  const token='__dropped_file__'+Date.now()+'_'+Math.random().toString(36).slice(2);
  droppedFiles.set(token,await file.arrayBuffer());
  return token;
 }catch(err){return '';}
}

contextBridge.exposeInMainWorld('desktopAPI',{
 inspectMailFile:async p=>{
  if(typeof p==='string'&&p.startsWith('__dropped_file__')){
   const data=droppedFiles.get(p);
   if(!data)throw new Error('Die abgelegte Outlook-Datei ist nicht mehr verfügbar.');
   return ipcRenderer.invoke('inspect-mail-data',data);
  }
  return ipcRenderer.invoke('inspect-mail-file',p);
 },
 inspectMailData:data=>ipcRenderer.invoke('inspect-mail-data',data),
 saveDroppedFile:async p=>{
  if(p?.path&&String(p.path).startsWith('__dropped_file__')){
   const data=droppedFiles.get(p.path);
   if(data)return ipcRenderer.invoke('save-dropped-file',{name:p.name,data});
  }
  return ipcRenderer.invoke('save-dropped-file',p);
 },
 openAttachment:p=>ipcRenderer.invoke('open-attachment',p),
 openMailto:email=>ipcRenderer.invoke('open-mailto',email),
 getPathForFile:getPathOrToken
});

// Electron 32+ removed File.path. Keep compatibility for the existing renderer.
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
