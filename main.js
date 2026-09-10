const { app, BrowserWindow, ipcMain, shell } = require('electron');
const path=require('path'), fs=require('fs');
const MsgReader=require('@kenjiuno/msgreader');

function createWindow(){
 const win=new BrowserWindow({width:1200,height:800,minWidth:900,minHeight:600,backgroundColor:'#111315',autoHideMenuBar:true,
 webPreferences:{contextIsolation:true,nodeIntegration:false,preload:path.join(__dirname,'preload.js')}});
 win.loadFile(path.join(__dirname,'index.html'));
}
function safeFileName(n){return String(n||'Anhang').replace(/[<>:"/\\|?*\x00-\x1F]/g,'_').slice(0,180)}
function saveFile(sourcePath,name,data){
 const dir=path.join(app.getPath('userData'),'attachments');fs.mkdirSync(dir,{recursive:true});
 const target=path.join(dir,`${Date.now()}-${safeFileName(name)}`);
 if(sourcePath&&fs.existsSync(sourcePath))fs.copyFileSync(sourcePath,target);
 else if(data)fs.writeFileSync(target,Buffer.from(data));
 else throw new Error('Datei konnte nicht übernommen werden.');
 return {name:safeFileName(name),path:target};
}
function decodeQP(s){return s.replace(/=\r?\n/g,'').replace(/=([A-Fa-f0-9]{2})/g,(_,h)=>String.fromCharCode(parseInt(h,16)))}
function parseEml(p){
 const raw=fs.readFileSync(p,'utf8').replace(/\r\n/g,'\n'),parts=raw.split(/\n\n/),headers=parts.shift()||'',body=parts.join('\n\n');
 const h=n=>{const m=headers.match(new RegExp('^'+n+':\\s*(.*(?:\\n[ \\t]+.*)*)$','im'));return m?m[1].replace(/\n[ \t]+/g,' ').trim():''};
 const from=h('From'),subject=h('Subject'),email=(from.match(/<([^>]+)>/)||[])[1]||((from.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i)||[])[0]||'');
 const senderName=from.replace(/<[^>]+>/,'').replace(email,'').replace(/["']/g,'').trim();
 return {subject,senderName,senderEmail:email,body:/quoted-printable/i.test(headers)?decodeQP(body):body.replace(/--[-_A-Za-z0-9]+[\s\S]*$/,'').trim()};
}
ipcMain.handle('inspect-mail-file',async(_,filePath)=>{
 if(!filePath||!fs.existsSync(filePath))throw new Error('Maildatei nicht gefunden.');
 const ext=path.extname(filePath).toLowerCase();
 if(ext==='.msg'||ext==='.oft'){const msg=new MsgReader(fs.readFileSync(filePath));const i=msg.getFileData();return {subject:i.subject||'',body:i.body||'',senderName:i.senderName||'',senderEmail:i.senderEmail||''};}
 if(ext==='.eml')return parseEml(filePath);
 return null;
});
ipcMain.handle('save-dropped-file',async(_,p)=>saveFile(p?.path,p?.name,p?.data));
ipcMain.handle('open-attachment',async(_,p)=>p?shell.openPath(p):undefined);
app.whenReady().then(()=>{createWindow();app.on('activate',()=>{if(BrowserWindow.getAllWindows().length===0)createWindow()})});
app.on('window-all-closed',()=>{if(process.platform!=='darwin')app.quit()});
