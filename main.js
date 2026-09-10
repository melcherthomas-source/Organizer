const { app, BrowserWindow, ipcMain, shell } = require('electron');
const path=require('path'), fs=require('fs');
const MsgReader=require('@kenjiuno/msgreader');

function createWindow(){
 const win=new BrowserWindow({
  width:1200,height:800,minWidth:900,minHeight:600,backgroundColor:'#111315',autoHideMenuBar:true,
  webPreferences:{contextIsolation:true,nodeIntegration:false,preload:path.join(__dirname,'preload.js')}
 });
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
function headerValue(headers,name){
 const re=new RegExp('^'+name+':\\s*(.*(?:\\r?\\n[ \\t]+.*)*)$','im');
 const m=String(headers||'').match(re);
 return m?m[1].replace(/\\r?\\n[ \\t]+/g,' ').trim():'';
}
function parseFromHeader(headers){
 const from=headerValue(headers,'From');
 const email=(from.match(/<([^>]+)>/)||[])[1]||((from.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\\.[A-Z]{2,}/i)||[])[0]||'');
 const senderName=from.replace(/<[^>]+>/,'').replace(email,'').replace(/^\"|\"$/g,'').trim();
 return {senderName,senderEmail:email};
}
function cleanSenderName(name,email){
 let s=String(name||'').replace(/\\s+/g,' ').trim().replace(/^\"|\"$/g,'').trim();
 if(s && !/@/.test(s) && !/^(unknown|noreply|no-reply|mailbox)$/i.test(s))return s;
 const local=String(email||'').split('@')[0].replace(/[0-9]+$/,'');
 const parts=local.split(/[._-]+/).filter(Boolean);
 if(parts.length>=2)return parts.slice(0,3).join(' ');
 return '';
}
function parseEml(p){
 const raw=fs.readFileSync(p,'utf8').replace(/\\r\\n/g,'\\n'),parts=raw.split(/\\n\\n/);
 const headers=parts.shift()||'',body=parts.join('\\n\\n');
 const info=parseFromHeader(headers);
 return {subject:headerValue(headers,'Subject'),senderName:cleanSenderName(info.senderName,info.senderEmail),senderEmail:info.senderEmail,body};
}
ipcMain.handle('inspect-mail-file',async(_,filePath)=>{
 if(!filePath||!fs.existsSync(filePath))throw new Error('Maildatei nicht gefunden.');
 const ext=path.extname(filePath).toLowerCase();
 if(ext==='.msg'||ext==='.oft'){
  const msg=new MsgReader(fs.readFileSync(filePath));
  const i=msg.getFileData()||{};
  const h=parseFromHeader(i.headers||'');
  const senderEmail=String(i.senderEmail||'').trim()||h.senderEmail;
  const senderName=cleanSenderName(String(i.senderName||'').trim(),senderEmail)||cleanSenderName(h.senderName,senderEmail);
  return {subject:i.subject||headerValue(i.headers,'Subject')||'',body:i.body||'',senderName,senderEmail,headers:i.headers||''};
 }
 if(ext==='.eml')return parseEml(filePath);
 return null;
});
ipcMain.handle('save-dropped-file',async(_,p)=>saveFile(p?.path,p?.name,p?.data));
ipcMain.handle('open-attachment',async(_,p)=>p?shell.openPath(p):undefined);
app.whenReady().then(()=>{createWindow();app.on('activate',()=>{if(BrowserWindow.getAllWindows().length===0)createWindow()})});
app.on('window-all-closed',()=>{if(process.platform!=='darwin')app.quit()});
