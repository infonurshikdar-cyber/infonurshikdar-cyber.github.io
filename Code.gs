const STATS_SHEET='Stats';
const HISTORY_SHEET='History';
const ADMIN_PIN='44990';

const INITIAL=[
  ['NUR',39,14,25,628,61,1,0],['EASIN',46,24,22,504,62,0,0],
  ['RIFAT',44,29,15,752,1,0,0],['SABBIR',5,3,2,56,8,0,0],
  ['TAWMID',35,21,14,400,45,1,0],['RAFUN',46,22,24,200,16,0,0],
  ['JUBAYER',21,11,10,188,19,0,0],['MAHI',44,27,17,236,48,0,0]
];

function doPost(e){
  const p=e&&e.parameter?e.parameter:{};
  let result;
  try{setup_();const action=p.action||'';
    if(action==='updatePhoto') result=updatePhoto_(p);
    else if(action==='addPlayer') result=addPlayer_(p);
    else throw new Error('Unknown action');
  }catch(err){result={ok:false,error:String(err.message||err)}}
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}
function doGet(e){
  const p=e&&e.parameter?e.parameter:{};
  if(p.api==='1') return api_(p);
  return HtmlService.createHtmlOutput('<h2>Cricket Stats Backend</h2><p>API is running.</p>');
}
function api_(p){
  let result;
  try{setup_();const action=p.action||'get';
    if(action==='get') result={ok:true,players:readPlayers_()};
    else if(action==='checkPin'){auth_(p);result={ok:true,message:'Admin verified'}}
    else if(action==='save') result=save_(p);
    else if(action==='addPlayer') result=addPlayer_(p);
    else if(action==='updatePhoto') result=updatePhoto_(p);
    else if(action==='deletePlayer') result=deletePlayer_(p);
    else if(action==='resetPlayer') result=resetPlayer_(p);
    else if(action==='deleteHistory') result=deleteHistory_(p);
    else if(action==='correct') result=correct_(p);
    else throw new Error('Unknown action');
  }catch(err){result={ok:false,error:String(err.message||err)}}
  const body=JSON.stringify(result),cb=p.callback;
  if(cb&&/^[A-Za-z_$][\w$\.]*$/.test(cb)) return ContentService.createTextOutput(cb+'('+body+');').setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(body).setMimeType(ContentService.MimeType.JSON);
}
function setup_(){
  const ss=SpreadsheetApp.getActiveSpreadsheet();let sh=ss.getSheetByName(STATS_SHEET);
  if(!sh){sh=ss.insertSheet(STATS_SHEET);sh.appendRow(['Player','Total Match','Won','Lost','Total Run','Total Wicket','50','100','Photo']);INITIAL.forEach(r=>sh.appendRow(r.concat([''])));}
  else if(sh.getLastColumn()<9)sh.getRange(1,9).setValue('Photo');
  let h=ss.getSheetByName(HISTORY_SHEET);if(!h){h=ss.insertSheet(HISTORY_SHEET);h.appendRow(['ID','Player','Match','Won','Lost','Run','Wicket','50','100','CreatedAt']);}
}
function readPlayers_(){
  const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(STATS_SHEET),v=sh.getDataRange().getValues();
  return v.slice(1).filter(r=>r[0]).map(r=>({name:String(r[0]),match:+r[1]||0,won:+r[2]||0,lost:+r[3]||0,run:+r[4]||0,wicket:+r[5]||0,fifty:+r[6]||0,hundred:+r[7]||0,photo:String(r[8]||'')}));
}
function auth_(p){if(String(p.pin||'')!==ADMIN_PIN)throw new Error('ভুল Admin PIN')}
function save_(p){
  auth_(p);const name=String(p.player||''),sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(STATS_SHEET),vals=sh.getDataRange().getValues();
  let row=-1;for(let i=1;i<vals.length;i++)if(String(vals[i][0])===name){row=i+1;break}if(row<0)throw new Error('Player not found');
  const m=num_(p.match),w=num_(p.won),l=num_(p.lost),r=num_(p.run),wk=num_(p.wicket),f=num_(p.fifty),h=num_(p.hundred),cur=sh.getRange(row,2,1,7).getValues()[0];
  sh.getRange(row,2,1,7).setValues([[cur[0]+m,cur[1]+w,cur[2]+l,cur[3]+r,cur[4]+wk,cur[5]+f,cur[6]+h]]);
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HISTORY_SHEET).appendRow([Utilities.getUuid(),name,m,w,l,r,wk,f,h,new Date()]);
  return {ok:true,message:'সেভ হয়েছে',players:readPlayers_()};
}
function correct_(p){
  auth_(p);
  const name=String(p.player||'').trim();
  if(!name)throw new Error('Player নির্বাচন করুন');
  const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(STATS_SHEET);
  const vals=sh.getDataRange().getValues();
  let row=-1;
  for(let i=1;i<vals.length;i++)if(String(vals[i][0])===name){row=i+1;break}
  if(row<0)throw new Error('Player not found');
  const sub=[num_(p.match),num_(p.won),num_(p.lost),num_(p.run),num_(p.wicket),num_(p.fifty),num_(p.hundred)];
  if(sub.reduce((a,b)=>a+b,0)<=0)throw new Error('কমপক্ষে একটি সংখ্যা দিন');
  const cur=sh.getRange(row,2,1,7).getValues()[0].map(Number);
  for(let i=0;i<7;i++)if(sub[i]>cur[i])throw new Error('যত আছে তার চেয়ে বেশি কমানো যাবে না');
  const next=cur.map((x,i)=>x-sub[i]);
  sh.getRange(row,2,1,7).setValues([next]);
  // Negative history entry keeps the correction reversible through History delete.
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HISTORY_SHEET).appendRow([Utilities.getUuid(),name,-sub[0],-sub[1],-sub[2],-sub[3],-sub[4],-sub[5],-sub[6],new Date()]);
  return {ok:true,message:'ভুল Stats কমানো হয়েছে ✅',players:readPlayers_()};
}
function addPlayer_(p){
  auth_(p);const name=String(p.name||'').trim();if(!name)throw new Error('Player name দিন');
  const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(STATS_SHEET);
  if(readPlayers_().some(x=>x.name.toLowerCase()===name.toLowerCase()))throw new Error('এই player আগে থেকেই আছে');
  sh.appendRow([name,0,0,0,0,0,0,0,'']);return {ok:true,message:'Player যোগ হয়েছে',players:readPlayers_()};
}
function updatePhoto_(p){
  auth_(p);const name=String(p.name||''),data=String(p.photo||'');if(!data)throw new Error('ছবি পাওয়া যায়নি');
  const m=data.match(/^data:image\/(jpeg|jpg|png|webp);base64,(.+)$/i);if(!m)throw new Error('ছবির ফরম্যাট ঠিক নয়');
  const bytes=Utilities.base64Decode(m[2]);if(bytes.length>8388608)throw new Error('ছবিটি 8MB-এর বেশি। 8MB-এর মধ্যে ছবি দিন');
  const url=savePhotoToDrive_(data,name),sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(STATS_SHEET),v=sh.getDataRange().getValues();
  for(let i=1;i<v.length;i++)if(String(v[i][0])===name){
    const oldUrl=String(v[i][8]||'');sh.getRange(i+1,9).setValue(url);
    try{const mm=oldUrl.match(/(?:googleusercontent\.com\/d\/|drive\.google\.com\/(?:uc\?export=view\&id=|thumbnail\?id=))([^?&]+)/);if(mm&&mm[1])DriveApp.getFileById(mm[1]).setTrashed(true)}catch(e){}
    return {ok:true,message:'Player-এর ছবি সেভ হয়েছে',players:readPlayers_()};
  }throw new Error('Player পাওয়া যায়নি');
}
function savePhotoToDrive_(dataUrl,name){
  const m=dataUrl.match(/^data:([^;]+);base64,(.+)$/);if(!m)throw new Error('ছবির ফরম্যাট ঠিক নয়');
  const bytes=Utilities.base64Decode(m[2]);const mime=m[1].toLowerCase()==='image/jpg'?'image/jpeg':m[1];const ext=mime==='image/png'?'png':mime==='image/webp'?'webp':'jpg';const safe=String(name||'Player').replace(/[^A-Za-z0-9_-]/g,'_');
  const file=DriveApp.createFile(Utilities.newBlob(bytes,mime,safe+'_'+Date.now()+'.'+ext));
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK,DriveApp.Permission.VIEW);
  return 'https://drive.google.com/thumbnail?id='+file.getId()+'&sz=w3000';
}
function repairExistingPhotoCells_(){
  const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(STATS_SHEET);if(!sh)return;
  const v=sh.getDataRange().getValues();
  for(let i=1;i<v.length;i++){const name=String(v[i][0]||''),data=String(v[i][8]||'');if(!name||!data||data.indexOf('data:image/')!==0)continue;
    try{const m=data.match(/^data:([^;]+);base64,(.+)$/);if(!m)continue;const bytes=Utilities.base64Decode(m[2]);if(bytes.length>8388608)continue;const url=savePhotoToDrive_(data,name);sh.getRange(i+1,9).setValue(url);}catch(e){}
  }
}
function deletePlayer_(p){
  auth_(p);const name=String(p.name||''),sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(STATS_SHEET),v=sh.getDataRange().getValues();
  for(let i=1;i<v.length;i++)if(String(v[i][0])===name){sh.deleteRow(i+1);return {ok:true,message:'Player মুছে দেওয়া হয়েছে',players:readPlayers_()}}throw new Error('Player পাওয়া যায়নি');
}
function resetPlayer_(p){
  auth_(p);const name=String(p.name||''),sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(STATS_SHEET),v=sh.getDataRange().getValues();
  for(let i=1;i<v.length;i++)if(String(v[i][0])===name){sh.getRange(i+1,2,1,7).setValues([[0,0,0,0,0,0,0]]);return {ok:true,message:'Player stats reset হয়েছে',players:readPlayers_()}}throw new Error('Player পাওয়া যায়নি');
}
function deleteHistory_(p){
  auth_(p);const id=String(p.id||''),hs=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HISTORY_SHEET),v=hs.getDataRange().getValues();
  for(let i=1;i<v.length;i++)if(String(v[i][0])===id){const row=v[i],sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(STATS_SHEET),sv=sh.getDataRange().getValues();
    for(let j=1;j<sv.length;j++)if(String(sv[j][0])===String(row[1])){const nums=[1,2,3,4,5,6,7].map(k=>+row[k]||0),cur=sh.getRange(j+1,2,1,7).getValues()[0];sh.getRange(j+1,2,1,7).setValues([cur.map((x,k)=>x-nums[k])]);break}
    hs.deleteRow(i+1);return {ok:true,message:'শেষ ম্যাচটি মুছে দেওয়া হয়েছে',players:readPlayers_()}
  }throw new Error('History not found');
}
function num_(x){const n=Number(x);return isFinite(n)?n:0}
