const { app, BrowserWindow, ipcMain, shell } = require('electron');
const path=require('path'), fs=require('fs');
const MsgReader=require('@kenjiuno/msgreader').default;

const INTERNAL_DOMAIN='@caravan-spezialisten.de';
let firstNames=new Set();
try{
 const list=JSON.parse(fs.readFileSync(path.join(__dirname,'vornamen_de.json'),'utf8'));
 firstNames=new Set(list.map(x=>String(x).toLocaleLowerCase('de-DE')));
}catch(err){console.error('Vornamenliste konnte nicht geladen werden:',err)}

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
function isInternalEmail(email){return String(email||'').trim().toLowerCase().endsWith(INTERNAL_DOMAIN)}
function cleanSenderName(name,email){
 let s=String(name||'').replace(/\\s+/g,' ').trim().replace(/^\"|\"$/g,'').trim();
 if(isInternalEmail(email))return '';
 if(s && !/@/.test(s) && !/^(unknown|noreply|no-reply|mailbox|mailer-daemon)$/i.test(s))return s;
 const local=String(email||'').split('@')[0].replace(/[0-9]+$/,'');
 const parts=local.split(/[._-]+/).filter(Boolean);
 const first=parts.find(p=>firstNames.has(p.toLocaleLowerCase('de-DE')));
 if(first){
  const idx=parts.indexOf(first);
  const last=parts[idx+1]||parts[idx-1];
  if(last && !firstNames.has(last.toLocaleLowerCase('de-DE')))return first+' '+last;
 }
 if(parts.length>=2)return parts.slice(0,3).join(' ');
 return '';
}
function detectNameFromText(text,email=''){
 if(isInternalEmail(email))return '';
 const src=String(text||'').replace(/[<>(),;:]/g,' ');
 const words=src.split(/\\s+/).map(x=>x.replace(/^\"|\"$/g,'').trim()).filter(Boolean);
 for(let i=0;i<words.length;i++){
  const w=words[i].toLocaleLowerCase('de-DE');
  if(!firstNames.has(w))continue;
  const before=words[i-1]||'';
  const after=words[i+1]||'';
  const valid=x=>/^[A-Za-zÄÖÜäöüßÀ-ÿ'’-]{2,30}$/.test(x)&&!/^https?$/i.test(x)&&!x.includes('@');
  if(valid(after)&&!firstNames.has(after.toLocaleLowerCase('de-DE')))return words[i]+' '+after;
  if(valid(before)&&!firstNames.has(before.toLocaleLowerCase('de-DE')))return before+' '+words[i];
 }
 return '';
}
function rawMsgFallbackBuffer(raw){
 const text=Buffer.from(raw).toString('latin1').replace(/\x00/g,' ');
 const fromMatch=text.match(/From:\s*([^\r\n]{0,240})/i);
 const from=parseFromHeader(fromMatch?fromMatch[0]:'');
 const subject=(text.match(/Subject:\s*([^\r\n]{0,500})/i)||[])[1]||'';
 const emails=(text.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi)||[]);
 const senderEmail=from.senderEmail&&!isInternalEmail(from.senderEmail)?from.senderEmail:(emails.find(e=>!isInternalEmail(e))||'');
 const senderName=cleanSenderName(from.senderName,senderEmail)||detectNameFromText(text,senderEmail);
 const chunks=text.split(/\r?\n/).map(s=>s.trim()).filter(s=>s.length>=3 && /[A-Za-zÄÖÜäöüß]/.test(s));
 const body=chunks.filter(s=>!/^From:|^To:|^Cc:|^Subject:|^Date:|^Message-ID:|^MIME-Version:|^Content-/i.test(s)).slice(-80).join('\n');
 return {subject:subject.trim(),senderName,senderEmail,body,headers:fromMatch?fromMatch[0]:''};
}
function parseMsgBuffer(raw){
 const buffer=Buffer.from(raw);
 try{
  const msg=new MsgReader(buffer);
  const i=msg.getFileData()||{};
  const h=parseFromHeader(i.headers||'');
  let senderEmail=!isInternalEmail(h.senderEmail)?h.senderEmail:'';
  let senderName=cleanSenderName(String(i.senderName||'').trim(),senderEmail);
  if(!senderName) senderName=cleanSenderName(h.senderName,senderEmail);
  if(!senderName) senderName=detectNameFromText(h.senderName||'',senderEmail);
  if(!senderName) senderName=detectNameFromText(String(i.body||''),senderEmail);
  return {subject:i.subject||headerValue(i.headers,'Subject')||'',body:i.body||'',senderName,senderEmail,headers:i.headers||''};
 }catch(err){
  console.error('MSG-Parser fehlgeschlagen, verwende Fallback:',err);
  return rawMsgFallbackBuffer(buffer);
 }
}
function parseMsgFile(filePath){return parseMsgBuffer(fs.readFileSync(filePath));}
function parseEmlBuffer(raw){
 const text=Buffer.from(raw).toString('utf8').replace(/\r\n/g,'\n'),parts=text.split(/\n\n/);
 const headers=parts.shift()||'',body=parts.join('\n\n');
 const info=parseFromHeader(headers);
 const senderEmail=isInternalEmail(info.senderEmail)?'':info.senderEmail;
 const senderName=cleanSenderName(info.senderName,senderEmail)||detectNameFromText(body,senderEmail);
 return {subject:headerValue(headers,'Subject'),senderName,senderEmail,body,headers};
}
function parseEml(p){return parseEmlBuffer(fs.readFileSync(p));}
ipcMain.handle('inspect-mail-file',async(_,filePath)=>{
 if(!filePath||!fs.existsSync(filePath))throw new Error('Maildatei nicht gefunden.');
 const ext=path.extname(filePath).toLowerCase();
 if(ext==='.msg'||ext==='.oft')return parseMsgFile(filePath);
 if(ext==='.eml')return parseEml(filePath);
 return null;
});
ipcMain.handle('inspect-mail-data',async(_,data)=>{
 const raw=Buffer.from(data);
 const text=raw.toString('latin1',0,16);
 if(text.startsWith('MZ'))return null;
 if(raw.slice(0,8).equals(Buffer.from([0xD0,0xCF,0x11,0xE0,0xA1,0xB1,0x1A,0xE1])))return parseMsgBuffer(raw);
 return parseEmlBuffer(raw);
});
ipcMain.handle('save-dropped-file',async(_,p)=>saveFile(p?.path,p?.name,p?.data));
ipcMain.handle('open-attachment',async(_,p)=>p?shell.openPath(p):undefined);
app.whenReady().then(()=>{createWindow();app.on('activate',()=>{if(BrowserWindow.getAllWindows().length===0)createWindow()})});
app.on('window-all-closed',()=>{if(process.platform!=='darwin')app.quit()});
