const QUIZ_VERSION = 'hanlin-ch3-3-4-v1';
const ANSWER_KEY = [1,3,0,2,0,1,3,2,1,0,2,3,0,2,3,1,2,3,1,0];
const POINTS_PER_QUESTION = 5;

function doGet() {
  return json_({ok:true,service:'hanlin-biology-assessment-scores',quizVersion:QUIZ_VERSION,questions:20,pointsPerQuestion:5});
}
function doPost(e) {
  try {
    const text = e && e.postData && e.postData.contents;
    if (!text || text.length > 10000) fail_('INVALID_REQUEST');
    const p = JSON.parse(text);
    if (p.action === 'testStatus' || p.action === 'cleanupTests') return json_(testAdmin_(p));
    return json_(submit_(p));
  } catch (err) {
    const known = ['INVALID_REQUEST','INVALID_STUDENT','INVALID_ANSWERS','INVALID_EVENT','WRONG_VERSION','EVENT_CONFLICT','NOT_AUTHORIZED'];
    const code = known.indexOf(err.message) >= 0 ? err.message : 'SERVER_ERROR';
    if (code === 'SERVER_ERROR') console.error('Grade sync failed: '+String(err));
    return json_({ok:false,code:code});
  }
}
function json_(value) {return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON);}
function fail_(code) {throw new Error(code);}
function normal_(value,max) {if(typeof value!=='string')fail_('INVALID_STUDENT');const s=value.normalize('NFKC').trim().replace(/\s+/g,' ');if(!s||s.length>max||/[\u0000-\u001f]/.test(s))fail_('INVALID_STUDENT');return s;}
function textCell_(value) {const s=String(value);return /^[=+\-@\t\r]/.test(s)?"'"+s:s;}
function hash_(value) {return Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,value,Utilities.Charset.UTF_8));}
function getBook_() {return SpreadsheetApp.openById(privateConfig_().spreadsheetId);}
function authorize() {const book=getBook_();console.log('Grade receiver ready: '+book.getName());}
function validated_(p) {
  if(!p||p.quizVersion!==QUIZ_VERSION)fail_('WRONG_VERSION');
  if(typeof p.eventId!=='string'||!/^[a-zA-Z0-9-]{16,80}$/.test(p.eventId))fail_('INVALID_EVENT');
  const className=normal_(p.className,20),seat=normal_(p.seat,6),name=normal_(p.name,50);
  if(!Array.isArray(p.answers)||p.answers.length!==20||p.answers.some(a=>!Number.isInteger(a)||a<0||a>3))fail_('INVALID_ANSWERS');
  const correct=p.answers.reduce((n,a,i)=>n+(a===ANSWER_KEY[i]?1:0),0);
  const keySeat=/^\d+$/.test(seat)?String(Number(seat)):seat;
  return {eventId:p.eventId,className:className,seat:seat,name:name,answers:p.answers,correct:correct,score:correct*POINTS_PER_QUESTION,key:hash_(JSON.stringify([className,keySeat,name])),fingerprint:hash_(JSON.stringify([QUIZ_VERSION,className,keySeat,name,p.answers]))};
}
function submit_(payload) {
  const p=validated_(payload),lock=LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const book=getBook_(),log=book.getSheetByName('作答紀錄'),summary=book.getSheetByName('成績');
    if(!log||!summary)throw new Error('Missing score sheets');
    let records=log.getDataRange().getValues().slice(1);
    const duplicate=records.find(row=>String(row[1])===p.eventId);
    if(duplicate&&String(duplicate[10])!==p.fingerprint)fail_('EVENT_CONFLICT');
    if(!duplicate){
      const row=[new Date(),p.eventId,p.key,textCell_(p.className),textCell_(p.seat),textCell_(p.name),p.correct,p.score,JSON.stringify(p.answers),QUIZ_VERSION,p.fingerprint];
      const next=log.getLastRow()+1;log.getRange(next,4,1,3).setNumberFormat('@');log.getRange(next,1,1,row.length).setValues([row]);log.getRange(next,1).setNumberFormat('yyyy-mm-dd hh:mm:ss');records.push(row);
    }
    // The event log is authoritative. Rebuild this student's summary on retry,
    // including a retry after a log write succeeded but the summary write failed.
    const studentRows=records.filter(row=>String(row[2])===p.key),latest=studentRows[studentRows.length-1];
    const best=Math.max.apply(null,studentRows.map(row=>Number(row[7])));
    const existing=summary.getDataRange().getValues().slice(1).findIndex(row=>String(row[9])===p.key);
    const target=existing<0?summary.getLastRow()+1:existing+2;
    const result=[latest[0],latest[3],latest[4],latest[5],latest[6],latest[7],best,studentRows.length,QUIZ_VERSION,p.key];
    summary.getRange(target,2,1,3).setNumberFormat('@');summary.getRange(target,1,1,result.length).setValues([result]);summary.getRange(target,1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
    SpreadsheetApp.flush();
    return {ok:true,eventId:p.eventId,score:p.score,correct:p.correct,duplicate:!!duplicate};
  } finally {lock.releaseLock();}
}
function testAdmin_(p) {
  if(typeof p.adminToken!=='string'||p.adminToken!==privateConfig_().adminToken)fail_('NOT_AUTHORIZED');
  if(typeof p.className!=='string'||!/^QA_[A-Za-z0-9_]{6,16}$/.test(p.className))fail_('INVALID_REQUEST');
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  try {
    const book=getBook_(),summary=book.getSheetByName('成績'),log=book.getSheetByName('作答紀錄');
    const matching=(row,c,n)=>String(row[c])===p.className&&String(row[n])==='整合測試學生';
    if(p.action==='cleanupTests'){
      let removed=0;
      for(const spec of [[summary,1,3],[log,3,5]]){const rows=spec[0].getDataRange().getValues();for(let i=rows.length-1;i>=1;i--)if(matching(rows[i],spec[1],spec[2])){spec[0].deleteRow(i+1);removed++;}}
      SpreadsheetApp.flush();return {ok:true,removed:removed};
    }
    return {ok:true,summary:summary.getDataRange().getValues().slice(1).filter(row=>matching(row,1,3)).map(row=>({correct:row[4],latest:row[5],best:row[6],attempts:row[7]})),events:log.getDataRange().getValues().slice(1).filter(row=>matching(row,3,5)).map(row=>({eventId:row[1],score:row[7]}))};
  }finally{lock.releaseLock();}
}
