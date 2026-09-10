const { app, BrowserWindow, ipcMain, shell } = require('electron');
const path = require('path');
const fs = require('fs');

function createWindow(){
  const win=new BrowserWindow({
    width:1500,height:950,minWidth:1000,minHeight:650,
    backgroundColor:'#111315',autoHideMenuBar:true,
    webPreferences:{
      contextIsolation:true,
      nodeIntegration:false,
      preload:path.join(__dirname,'preload.js')
    }
  });
  win.loadFile(path.join(__dirname,'index.html'));
}

function safeFileName(name){
  return String(name || 'Anhang').replace(/[<>:"/\\|?*\x00-\x1F]/g,'_').slice(0,180);
}

ipcMain.handle('save-dropped-file', async (_event, payload) => {
  const name=safeFileName(payload?.name || 'E-Mail-Anhang');
  const attachmentsDir=path.join(app.getPath('userData'),'attachments');
  fs.mkdirSync(attachmentsDir,{recursive:true});
  let target=path.join(attachmentsDir, `${Date.now()}-${name}`);
  if(payload?.path && fs.existsSync(payload.path)){
    fs.copyFileSync(payload.path,target);
  }else if(payload?.data){
    fs.writeFileSync(target, Buffer.from(payload.data));
  }else{
    throw new Error('Datei konnte nicht übernommen werden.');
  }
  return {name, path:target};
});

ipcMain.handle('open-attachment', async (_event, filePath) => {
  if(!filePath) return;
  return shell.openPath(filePath);
});

app.whenReady().then(()=>{
  createWindow();
  app.on('activate',()=>{if(BrowserWindow.getAllWindows().length===0)createWindow();});
});
app.on('window-all-closed',()=>{if(process.platform!=='darwin')app.quit();});
