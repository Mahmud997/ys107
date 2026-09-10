/**
 * School Attendance Pro — Google Apps Script backend
 *
 * Sheets required:
 * users    : login | password | role | name
 * students : id | name | class
 * records  : id | studentId | date | time | status | note | teacher | updatedAt
 * sessions : token | login | role | name | lastSeen
 *
 * Deploy: Web app -> Execute as Me -> Who has access: Anyone.
 */

const SHEETS = { users:'users', students:'students', records:'records', sessions:'sessions' };
const SESSION_TTL_MS = 2 * 60 * 1000;

function json_(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  return handle_(e && e.parameter ? e.parameter : {});
}

function doPost(e) {
  let p = {};
  try {
    if (e && e.postData && e.postData.contents) {
      p = JSON.parse(e.postData.contents);
    }
  } catch (_) {
    p = e && e.parameter ? e.parameter : {};
  }
  return handle_(p);
}

function handle_(p) {
  const action = String(p.action || '');
  try {
    if (action === 'login') return login_(p);
    if (action === 'getStudents') return withAuth_(p, getStudents_);
    if (action === 'getRecords') return withAuth_(p, getRecords_);
    if (action === 'addRecord') return withAuth_(p, addRecord_);
    if (action === 'addStudent') return withRole_(p, 'zavuch', addStudent_);
    if (action === 'importStudents') return withRole_(p, 'zavuch', importStudents_);
    if (action === 'removeStudent') return withRole_(p, 'zavuch', removeStudent_);
    if (action === 'removeStudentRecords') return withRole_(p, 'zavuch', removeStudentRecords_);
    if (action === 'resetAll') return withRole_(p, 'zavuch', resetAll_);
    if (action === 'heartbeat') return withAuth_(p, heartbeat_);
    if (action === 'getOnlineUsers') return withAuth_(p, getOnlineUsers_);
    if (action === 'logout') return withAuth_(p, logout_);
    if (action === 'health') return json_({ success:true, time:new Date().toISOString() });
    return json_({ success:false, error:'Неизвестное действие' });
  } catch (err) {
    console.error(err);
    return json_({ success:false, error:String(err.message || err) });
  }
}

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function sheet_(name) {
  const s = ss_().getSheetByName(name);
  if (!s) throw new Error('Не найден лист: ' + name);
  return s;
}

function ensureSessionsSheet_() {
  let s = ss_().getSheetByName(SHEETS.sessions);
  if (!s) {
    s = ss_().insertSheet(SHEETS.sessions);
    s.getRange(1,1,1,5).setValues([['token','login','role','name','lastSeen']]);
  }
  return s;
}

function makeId_() {
  return Utilities.getUuid();
}

function normalizeRole_(role) {
  const r = String(role || '').trim().toLowerCase().replace(/ё/g,'е');
  const normalized = r.replace(/[^a-zа-я0-9]+/gi,'');
  if (['zavuch','завуч','admin','administrator','администратор'].includes(normalized)) return 'zavuch';
  return 'teacher';
}

function login_(p) {
  const login = String(p.login || '').trim();
  const password = String(p.password || '');
  if (!login || !password) return json_({ success:false, error:'Введите логин и пароль' });

  const data = sheet_(SHEETS.users).getDataRange().getValues();
  for (let i=1; i<data.length; i++) {
    const rowLogin = String(data[i][0] || '').trim();
    const rowPassword = String(data[i][1] || '');
    if (rowLogin.toLowerCase() === login.toLowerCase() && rowPassword === password) {
      const token = Utilities.getUuid() + Utilities.getUuid().replace(/-/g,'');
      const s = ensureSessionsSheet_();
      cleanupSessions_(s);
      // Логин zavuch всегда получает роль завуча; в остальных случаях используется колонка role.
      const role = rowLogin.toLowerCase() === 'zavuch' ? 'zavuch' : normalizeRole_(data[i][2] || 'teacher');
      s.appendRow([token, rowLogin, role, String(data[i][3] || rowLogin), Date.now()]);
      return json_({ success:true, token, role, name:String(data[i][3] || rowLogin) });
    }
  }
  return json_({ success:false, error:'Неверный логин или пароль' });
}

function auth_(token) {
  token = String(token || '');
  if (!token) return null;
  const s = ensureSessionsSheet_();
  cleanupSessions_(s);
  const data = s.getDataRange().getValues();
  for (let i=1;i<data.length;i++) {
    if (String(data[i][0]) === token) {
      return { row:i+1, token:String(data[i][0]), login:String(data[i][1]), role:String(data[i][2]), name:String(data[i][3]), lastSeen:Number(data[i][4]) };
    }
  }
  return null;
}

function withAuth_(p, fn) {
  const user = auth_(p.token);
  if (!user) return json_({ success:false, error:'Сессия недействительна или истекла. Войдите снова.', auth:false });
  return fn(p, user);
}

function withRole_(p, role, fn) {
  const user = auth_(p.token);
  if (!user) return json_({ success:false, error:'Сессия недействительна или истекла.', auth:false });
  if (normalizeRole_(user.role) !== normalizeRole_(role)) return json_({ success:false, error:'Недостаточно прав. Войдите под учётной записью завуча.' });
  return fn(p, user);
}

function heartbeat_(p, user) {
  const s = ensureSessionsSheet_();
  s.getRange(user.row,5).setValue(Date.now());
  cleanupSessions_(s);
  return json_({ success:true, online:getOnlineUsersData_() });
}

function logout_(p, user) {
  ensureSessionsSheet_().deleteRow(user.row);
  return json_({ success:true });
}

function cleanupSessions_(s) {
  const data = s.getDataRange().getValues();
  const now = Date.now();
  const rows=[];
  for (let i=1;i<data.length;i++) {
    if (!data[i][0] || now - Number(data[i][4] || 0) > SESSION_TTL_MS) rows.push(i+1);
  }
  rows.sort((a,b)=>b-a).forEach(r=>s.deleteRow(r));
}

function getOnlineUsersData_() {
  const s=ensureSessionsSheet_();
  cleanupSessions_(s);
  const data=s.getDataRange().getValues();
  const unique={};
  for(let i=1;i<data.length;i++) unique[String(data[i][1])] = { login:String(data[i][1]), name:String(data[i][3]), role:String(data[i][2]) };
  return Object.values(unique);
}

function getOnlineUsers_(p,user) { return json_({ success:true, users:getOnlineUsersData_() }); }

function getStudents_(p,user) {
  const data=sheet_(SHEETS.students).getDataRange().getValues();
  return json_(data.slice(1).filter(r=>r[0] !== '').map(r=>({id:String(r[0]),name:String(r[1]||''),class:String(r[2]||'')})));
}

function getRecords_(p,user) {
  const data=sheet_(SHEETS.records).getDataRange().getValues();
  return json_(data.slice(1).filter(r=>r[0] !== '').map(r=>({
    id:String(r[0]), studentId:String(r[1]), date:formatDate_(r[2]), time:String(r[3]||''),
    status:String(r[4]||''), note:String(r[5]||''), teacher:String(r[6]||''), updatedAt:String(r[7]||'')
  })));
}

function formatDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v || '').slice(0,10);
}

function addRecord_(p,user) {
  const studentId=String(p.studentId||'').trim();
  const date=String(p.date||'').trim();
  const time=String(p.time||'').trim();
  const status=String(p.status||'present').trim();
  const note=String(p.note||'').trim();
  if(!studentId || !date) return json_({success:false,error:'Не указан ученик или дата.'});

  const lock=LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const s=sheet_(SHEETS.records);
    const data=s.getDataRange().getValues();
    // Один ученик — одна отметка в день. Повторная отметка обновляет существующую.
    for(let i=1;i<data.length;i++) {
      if(String(data[i][1])===studentId && formatDate_(data[i][2])===date) {
        const id=String(data[i][0]);
        s.getRange(i+1,3,1,6).setValues([[date,time,status,note,user.name + ' (#' + user.login + ')',new Date().toISOString()]]);
        return json_({success:true,updated:true,id});
      }
    }
    const id=makeId_();
    s.appendRow([id,studentId,date,time,status,note,user.name + ' (#' + user.login + ')',new Date().toISOString()]);
    return json_({success:true,updated:false,id});
  } finally { lock.releaseLock(); }
}

function addStudent_(p,user) {
  const name=String(p.name||'').trim();
  const cls=String(p.class||'').trim();
  if(!name || !cls) return json_({success:false,error:'Введите ФИО и класс.'});
  const s=sheet_(SHEETS.students);
  const data=s.getDataRange().getValues();
  if(data.slice(1).some(r=>String(r[1]).trim().toLowerCase()===name.toLowerCase() && String(r[2]).trim().toLowerCase()===cls.toLowerCase()))
    return json_({success:false,error:'Такой ученик уже есть в этом классе.'});
  const id=makeId_();
  s.appendRow([id,name,cls]);
  return json_({success:true,id,name,class:cls});
}

function importStudents_(p,user) {
  let students=[];
  try { students=JSON.parse(String(p.students||'[]')); } catch(e) { return json_({success:false,error:'Неверный формат данных Excel.'}); }
  if (!Array.isArray(students) || !students.length) return json_({success:false,error:'В Excel нет учеников.'});
  const s=sheet_(SHEETS.students);
  const data=s.getDataRange().getValues();
  const existing=new Set(data.slice(1).filter(r=>r[0]).map(r=>String(r[1]).trim().toLowerCase()+'|'+String(r[2]).trim().toLowerCase()));
  const lock=LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    let added=0, skipped=0, invalid=0;
    const out=[];
    students.forEach(item=>{
      const name=String(item.name||'').trim();
      const cls=String(item.class||'').trim();
      if(!name || !cls){ invalid++; return; }
      const key=name.toLowerCase()+'|'+cls.toLowerCase();
      if(existing.has(key)){ skipped++; return; }
      const id=makeId_();
      out.push([id,name,cls]);
      existing.add(key); added++;
    });
    if(out.length) s.getRange(s.getLastRow()+1,1,out.length,3).setValues(out);
    return json_({success:true,added,skipped,invalid});
  } finally { lock.releaseLock(); }
}

function removeStudent_(p,user) {
  const id=String(p.id||'');
  const s=sheet_(SHEETS.students); const data=s.getDataRange().getValues();
  for(let i=1;i<data.length;i++) if(String(data[i][0])===id){s.deleteRow(i+1); return json_({success:true});}
  return json_({success:false,error:'Ученик не найден.'});
}

function removeStudentRecords_(p,user) {
  const id=String(p.studentId||''); const s=sheet_(SHEETS.records); const data=s.getDataRange().getValues();
  const rows=[]; for(let i=1;i<data.length;i++) if(String(data[i][1])===id) rows.push(i+1);
  rows.sort((a,b)=>b-a).forEach(r=>s.deleteRow(r));
  return json_({success:true,deleted:rows.length});
}

function resetAll_(p,user) {
  const s=sheet_(SHEETS.records);
  const last=s.getLastRow();
  if(last>1) s.getRange(2,1,last-1,s.getLastColumn()).clearContent();
  return json_({success:true,deleted:Math.max(0,last-1)});
}

/** Run once manually after creating the spreadsheet. */
function setupSheets() {
  const ss=ss_();
  const defs={
    users:['login','password','role','name'],
    students:['id','name','class'],
    records:['id','studentId','date','time','status','note','teacher','updatedAt'],
    sessions:['token','login','role','name','lastSeen']
  };
  Object.keys(defs).forEach(name=>{
    let s=ss.getSheetByName(name); if(!s) s=ss.insertSheet(name);
    if(s.getLastRow()===0) s.getRange(1,1,1,defs[name].length).setValues([defs[name]]);
  });
}
