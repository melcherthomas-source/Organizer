const { app, BrowserWindow } = require('electron');
const path = require('path');
function createWindow(){const win=new BrowserWindow({width:1500,height:950,minWidth:1000,minHeight:650,backgroundColor:'#111315',autoHideMenuBar:true,webPreferences:{contextIsolation:true,nodeIntegration:false}});win.loadFile(path.join(__dirname,'index.html'));}
app.whenReady().then(()=>{createWindow();app.on('activate',()=>{if(BrowserWindow.getAllWindows().length===0)createWindow();});});
app.on('window-all-closed',()=>{if(process.platform!=='darwin')app.quit();});
