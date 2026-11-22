/**
 * 教學原理互評系統 v5.2（2026 秋季）
 *
 * 關鍵修正：學生頁面改由 Apps Script HtmlService 提供，
 * 直接以 google.script.run 呼叫本程式，移除舊版跨網域 iframe 回傳逾時問題。
 * GitHub Pages index.html 僅嵌入此 /exec 頁面，不公開學生名單。
 *
 * 部署：綁定 Google Sheet 的 Apps Script，替換 Code.gs，新增 HTML 檔 Vote，
 * 執行 setup() 授權，管理部署 > 編輯 > 建立新版本，沿用原本 /exec。
 * 部署身分：「我」；存取：「任何人」。
 */
const TP_SETTINGS = Object.freeze({
  timezone: 'Asia/Taipei',
  replies: '互評回覆',
  roster: '互評正式名單2026', // 避免碰觸舊版「學生名單」
  summary: '統計總覽',
  spreadsheetProperty: 'TEACHING_PRINCIPLES_SHEET_2026',
  release: '2026-11-23-v5.2'
});
const TP_WEEKS = [
  {id:'2026-11-23',open:'2026-11-22',end:'2026-11-24',label:'第一週',groups:[
    {id:1,topic:'工作分析教學法'},{id:2,topic:'協同教學法'}]},
  {id:'2026-12-07',open:'2026-12-06',end:'2026-12-08',label:'第二週',groups:[
    {id:3,topic:'探究教學法／講述教學法'},{id:4,topic:'直接教學法'}]},
  {id:'2026-12-14',open:'2026-12-13',end:'2026-12-15',label:'第三週',groups:[
    {id:5,topic:'討論教學法'},{id:6,topic:'結構式教學法'}]}
];

// 教師提供的 34 人正式名單。只存在後端，不會放進 GitHub Pages 或送到學生端。
const TP_STUDENTS = [
  {name:"康凱浚",id:"g11433011",group:1},
  {name:"陳享恩",id:"g11433010",group:1},
  {name:"張又文",id:"u11246017",group:1},
  {name:"黃星凱",id:"u11246022",group:1},
  {name:"蔣雅婷",id:"u11246023",group:1},
  {name:"李玉蘭",id:"g11535003",group:1},
  {name:"陳莉穎",id:"g11433012",group:2},
  {name:"張心禹",id:"m11407111",group:2},
  {name:"黃意娟",id:"m11413123",group:2},
  {name:"楊子妮",id:"u11249037",group:2},
  {name:"吳宜璇",id:"g11307017",group:2},
  {name:"侯欣妤",id:"u11449038",group:3},
  {name:"黃子晏",id:"u11449019",group:3},
  {name:"黃得恩",id:"u11203014",group:3},
  {name:"羅宸誼",id:"u11344023",group:3},
  {name:"林妍潔",id:"u11444021",group:3},
  {name:"許舒斐",id:"u11344017",group:3},
  {name:"陳永馨",id:"g11413011",group:4},
  {name:"邱昱翔",id:"u11241052",group:4},
  {name:"陳鼎翔",id:"u11445013",group:4},
  {name:"盧季萱",id:"u11445017",group:4},
  {name:"張斌",id:"u11345015",group:4},
  {name:"陳祥悠",id:"u11445002",group:4},
  {name:"謝忱憲",id:"u11345046",group:5},
  {name:"陳伯易",id:"u11213003",group:5},
  {name:"趙冠東",id:"u11445151",group:5},
  {name:"黃元鉦",id:"u11145051",group:5},
  {name:"蕭亦庭",id:"u11448030",group:5},
  {name:"陳思吟",id:"m11005014",group:6},
  {name:"謝承辰",id:"d11401010",group:6},
  {name:"簡雅雪",id:"g11334005",group:6},
  {name:"鄭依旻",id:"g11334007",group:6},
  {name:"朱凌潔",id:"g11217003",group:6},
  {name:"蕭明晏",id:"g11505007",group:6}
];
const TP_NEW_HEADERS = ['提交時間','報告日期','週次','學號','姓名','獲選組別','報告主題','心得（至少30字）','提交編號（防重送）'];
// 相容 v5.1 原有表頭；只更新欄名，不刪除或改動學生心得。
const TP_PRIOR_NEW_HEADERS = ['提交時間','報告日期','週次','學號','姓名','獲選組別','報告主題','30字內心得','提交編號（防重送）'];
const TP_OLD_HEADERS = ['提交時間','報告日期','週次','獲選組別','報告主題','30字內心得','提交編號（防重送）'];

function doGet() {
  // 同時可直接開啟 /exec 或嵌入 GitHub Pages。
  return HtmlService.createHtmlOutputFromFile('Vote')
    .setTitle('教學原理｜跨週次報告互評')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** 教師執行一次；重複執行不會清除既有作答資料。 */
function setup() {
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('請從 Google Sheet 的「擴充功能 → Apps Script」執行 setup()。');
  PropertiesService.getScriptProperties().setProperty(TP_SETTINGS.spreadsheetProperty,ss.getId());
  tpPrepareReplies_(ss.getSheetByName(TP_SETTINGS.replies)||ss.insertSheet(TP_SETTINGS.replies));
  tpPrepareRoster_(ss.getSheetByName(TP_SETTINGS.roster)||ss.insertSheet(TP_SETTINGS.roster));
  tpBuildSummary_(ss.getSheetByName(TP_SETTINGS.summary)||ss.insertSheet(TP_SETTINGS.summary));
  Logger.log('v5.2 正式名單完成：34 人，6 組；第一週免評11人、第二週12人、第三週11人。舊「學生名單」保留。');
  Logger.log('試算表：'+ss.getUrl());
  return '名單與統計表已設定完成。';
}

function tpPrepareReplies_(sheet) {
  if (sheet.getLastRow()===0) {
    sheet.getRange(1,1,1,9).setValues([TP_NEW_HEADERS]);
  } else {
    const head=sheet.getRange(1,1,1,Math.max(9,sheet.getLastColumn())).getDisplayValues()[0];
    const isNew=TP_NEW_HEADERS.every((t,i)=>head[i]===t);
    const isPriorNew=TP_PRIOR_NEW_HEADERS.every((t,i)=>head[i]===t);
    const isOld=TP_OLD_HEADERS.every((t,i)=>head[i]===t);
    if (!isNew && !isPriorNew && !isOld) throw new Error('「互評回覆」欄位不符；已停止設定，原作答未更動。');
    if (isOld) {
      // 舊匿名版：只插入學號與姓名欄，保留原有心得及提交資料。
      sheet.insertColumnsBefore(4,2);
    }
    if (isPriorNew || isOld) {
      // 舊版欄名「30字內心得」改為「心得（至少30字）」；歷史紀錄一律保留。
      sheet.getRange(1,1,1,9).setValues([TP_NEW_HEADERS]);
    }
  }
  sheet.setFrozenRows(1);
  sheet.getRange(1,1,1,9).setFontColor('#fff').setBackground('#24405e').setFontWeight('bold');
  sheet.getRange('A:A').setNumberFormat('yyyy/mm/dd hh:mm:ss');
  sheet.getRange('B:B').setNumberFormat('@');
  sheet.getRange('D:E').setNumberFormat('@');
  sheet.getRange('H:H').setWrap(true);
  [176,128,95,142,110,100,255,305,265].forEach((n,i)=>sheet.setColumnWidth(i+1,n));
}

/** 另建專用「互評正式名單2026」，不修改既有「學生名單」。重複執行不會刪除已填學號。 */
function tpPrepareRoster_(sheet) {
  if (sheet.getLastRow()===0) {
    sheet.getRange(1,1,1,6).setValues([['姓名','學號（正式名單）','組別','報告日期','報告主題','免評週次']]);
  } else {
    const head=sheet.getRange(1,1,1,2).getDisplayValues()[0];
    if (head[0]!=='姓名' || !['學號（教師可補，建議）','學號（正式名單）'].includes(head[1])) {
      throw new Error('「互評正式名單2026」欄位不符；請勿清空此表，請教師檢查後重新執行。原有「學生名單」不受影響。');
    }
  }
  sheet.getRange(1,1,1,6).setValues([['姓名','學號（正式名單）','組別','報告日期','報告主題','免評週次']]);
  sheet.getRange('B:B').setNumberFormat('@');
  const old=sheet.getLastRow()>1?sheet.getRange(2,1,sheet.getLastRow()-1,2).getDisplayValues():[];
  const names=new Map();
  old.forEach((r,i)=>{
    const name=tpNormalizeName_(r[0]);
    if (!name) return;
    if (names.has(name)) throw new Error('「互評正式名單2026」姓名重複：'+name+'，請教師先核對。');
    names.set(name,{row:i+2,id:tpNormalizeId_(r[1])});
  });
  const problems=TP_STUDENTS.filter(s=>names.has(s.name)&&names.get(s.name).id&&names.get(s.name).id!==tpNormalizeId_(s.id));
  if (problems.length) throw new Error('下列學生原名單已有不同學號，未覆寫：'+problems.map(s=>s.name).join('、')+'。請教師確認後重新執行 setup()。');
  let nextRow=sheet.getLastRow()+1;
  for (const s of TP_STUDENTS) {
    const oldEntry=names.get(s.name);
    const row=oldEntry?oldEntry.row:nextRow++;
    const week=TP_WEEKS.find(w=>w.groups.some(g=>g.id===s.group));
    const group=week.groups.find(g=>g.id===s.group);
    sheet.getRange(row,1,1,6).setValues([[s.name,s.id,'第'+s.group+'組',week.id,group.topic,week.label]]);
  }
  sheet.getRange(1,1,1,6).setBackground('#e6eef9').setFontWeight('bold').setFontColor('#203b5b');
  sheet.setFrozenRows(1);
  [152,190,100,133,255,110].forEach((n,i)=>sheet.setColumnWidth(i+1,n));
  sheet.getRange('H1').setValue('正式名單：34人。學生以學號＋姓名確認，該週報告組別自動免填。');
  sheet.getRange('H1').setFontColor('#697b8e');
  sheet.setColumnWidth(8,540);
}

function tpBuildSummary_(sheet) {
  sheet.getRange(1,1,2,8).breakApart();
  sheet.clear();
  const rows=TP_WEEKS.flatMap(w=>w.groups.map(g=>[
    w.id,w.label,'第'+g.id+'組',g.topic
  ]));
  sheet.getRange(1,1,1,7).merge().setValue('教學原理｜報告互評統計（報告組員當週免填）');
  sheet.getRange(1,1).setFontSize(15).setFontWeight('bold').setFontColor('#203b5b');
  sheet.getRange(2,1,1,7).merge().setValue('同學每週只選一組；該週兩組報告者免評，預計有效填答母數：23 / 22 / 23 人。');
  sheet.getRange(4,1,1,7).setValues([['報告日期','週次','組別','主題','獲選票數','有效投票數','當週得票比例']])
    .setFontWeight('bold').setBackground('#e6eef9');
  sheet.getRange(5,1,rows.length,4).setNumberFormat('@').setValues(rows);
  for(let i=0;i<rows.length;i++){
    const r=5+i;
    sheet.getRange(r,5).setFormula('=COUNTIFS(\''+TP_SETTINGS.replies+'\'!$B:$B,A'+r+',\''+TP_SETTINGS.replies+'\'!$F:$F,C'+r+')');
    sheet.getRange(r,6).setFormula('=COUNTIF(\''+TP_SETTINGS.replies+'\'!$B:$B,A'+r+')');
    sheet.getRange(r,7).setFormula('=IFERROR(E'+r+'/F'+r+',0)');
  }
  sheet.getRange(5,7,rows.length,1).setNumberFormat('0.0%');
  sheet.setFrozenRows(4);
  [140,105,100,265,110,120,143].forEach((n,i)=>sheet.setColumnWidth(i+1,n));
}

/** 直接由 Vote.html 的 google.script.run 呼叫，不經跨網域 iframe POST。 */
function getDashboard(payload) {
  try {
    const student=tpValidateStudent_(payload);
    const ss=tpOpenSheet_();
    const records=tpReadEntries_(ss.getSheetByName(TP_SETTINGS.replies));
    tpValidateHistoryIdentity_(records,student);
    return tpDashboardResult_(student,records,'已取得歷次互評紀錄。');
  } catch(err) { return {ok:false,message:String(err&&err.message||err)}; }
}

function submitPeerVote(payload) {
  const lock=LockService.getScriptLock();
  try {
    const student=tpValidateStudent_(payload);
    if(!payload||!(/^[a-zA-Z0-9-]{16,90}$/).test(String(payload.requestId||''))) {
      throw new Error('提交編號無效，請重新整理再試。');
    }
    const week=TP_WEEKS.find(w=>w.id===String(payload.weekId||''));
    if(!week) throw new Error('報告週次不正確。');
    const group=week.groups.find(g=>g.id===Number(payload.groupId));
    if(!group) throw new Error('請選擇該週的報告組別。');
    const comment=String(payload.comment==null?'':payload.comment).replace(/\s+/g,' ').trim();
    if (Array.from(comment).length < 30) {
      throw new Error('心得至少需填寫 30 字，請補充後再送出。');
    }
    const studentPresentWeek=TP_WEEKS.find(w=>w.groups.some(g=>g.id===student.group));
    if(studentPresentWeek.id===week.id) {
      throw new Error('你是本週報告組員，依課堂規則免填本週互評。');
    }
    lock.waitLock(20000);
    const ss=tpOpenSheet_();
    const replies=ss.getSheetByName(TP_SETTINGS.replies);
    const records=tpReadEntries_(replies);
    tpValidateHistoryIdentity_(records,student);
    const id=tpNormalizeId_(student.id);
    const duplicateRequest=records.find(r=>r.requestId===payload.requestId);
    if(duplicateRequest) {
      if(duplicateRequest.studentId!==id||duplicateRequest.weekId!==week.id) {
        throw new Error('提交編號重複但內容不符，請重新登入。');
      }
      return tpDashboardResult_(student,records,'先前已成功送出，不會重複計票。');
    }
    const existing=records.find(r=>r.studentId===id&&r.weekId===week.id);
    if(existing) return tpDashboardResult_(student,records,'本週已有紀錄，不能修改或再次填答。');
    const today=Utilities.formatDate(new Date(),TP_SETTINGS.timezone,'yyyy-MM-dd');
    if(today<week.open||today>week.end) {
      throw new Error('本週僅於 '+week.open+' 00:00 至 '+week.end+' 23:59（臺灣時間）開放。');
    }
    const time=new Date();
    const requestId=String(payload.requestId);
    const row=replies.getLastRow()+1;
    const safeComment=/^[=+\-@]/.test(comment)?"'"+comment:comment;
    replies.getRange(row,4,1,2).setNumberFormat('@');
    replies.getRange(row,1,1,9).setValues([[
      time,week.id,week.label,id,student.name,'第'+group.id+'組',group.topic,safeComment,requestId
    ]]);
    SpreadsheetApp.flush();
    records.push({submittedAt:Utilities.formatDate(time,TP_SETTINGS.timezone,'yyyy/MM/dd HH:mm:ss'),
      weekId:week.id,weekLabel:week.label,studentId:id,studentName:student.name,
      groupLabel:'第'+group.id+'組',topic:group.topic,comment:comment,requestId:requestId});
    return tpDashboardResult_(student,records,'互評已寫入 Google Sheet，送出後不可修改。');
  } catch(err) { return {ok:false,message:String(err&&err.message||err)}; }
  finally { try { lock.releaseLock(); } catch(_) {} }
}

function tpDashboardResult_(student,records,message) {
  const week=TP_WEEKS.find(w=>w.groups.some(g=>g.id===student.group));
  const id=tpNormalizeId_(student.id);
  return {ok:true,message,studentName:student.name,studentId:id,
    presenterGroup:student.group,exemptWeekId:week.id,
    records:records.filter(r=>r.studentId===id&&r.studentName===student.name)
      .map(r=>({weekId:r.weekId,weekLabel:r.weekLabel,submittedAt:r.submittedAt,
        groupLabel:r.groupLabel,topic:r.topic,comment:r.comment}))
      .sort((a,b)=>b.weekId.localeCompare(a.weekId))};
}
function tpValidateStudent_(p) {
  const id=tpNormalizeId_(p&&p.studentId);
  const name=tpNormalizeName_(p&&p.studentName);
  if(!/^[A-Z0-9-]{4,24}$/.test(id)) throw new Error('學號請填入正確格式。');
  if(Array.from(name).length<2||Array.from(name).length>30) throw new Error('請填寫完整姓名。');
  const student=TP_STUDENTS.find(s=>tpNormalizeId_(s.id)===id);
  if(!student) throw new Error('此學號不在本課程名單中，請重新確認。');
  if(student.name!==name) throw new Error('學號與姓名不相符，請確認後重新輸入。');
  return student;
}
function tpOpenSheet_() {
  const id=PropertiesService.getScriptProperties().getProperty(TP_SETTINGS.spreadsheetProperty);
  if(!id) throw new Error('教師尚未執行 setup()，請通知授課教師。');
  const ss=SpreadsheetApp.openById(id);
  if(!ss.getSheetByName(TP_SETTINGS.replies)) throw new Error('找不到互評回覆工作表，請教師執行 setup()。');
  return ss;
}
function tpNormalizeId_(s) {return String(s==null?'':s).trim().toUpperCase();}
function tpNormalizeName_(s) {return String(s==null?'':s).replace(/\s+/g,'').trim();}
function tpReadEntries_(sheet) {
  if(sheet.getLastRow()<=1) return [];
  return sheet.getRange(2,1,sheet.getLastRow()-1,9).getDisplayValues().map(r=>({
    submittedAt:r[0],weekId:r[1],weekLabel:r[2],studentId:tpNormalizeId_(r[3]),
    studentName:tpNormalizeName_(r[4]),groupLabel:r[5],topic:r[6],comment:r[7],requestId:r[8]
  }));
}
function tpValidateHistoryIdentity_(records,student) {
  const id=tpNormalizeId_(student.id);
  for(const r of records) {
    if(!r.studentId||!r.studentName) continue; // 舊版匿名紀錄不連結到個人。
    if(r.studentId===id&&r.studentName!==student.name) {
      throw new Error('此學號曾出現不同姓名，請通知教師核對舊紀錄。');
    }
    if(r.studentName===student.name&&r.studentId!==id) {
      throw new Error('此姓名曾以其他學號提交，請通知教師核對舊紀錄。');
    }
  }
}

/** 教師可在 Apps Script 編輯器執行，確認試算表連結與名單對應；不寫入投票。 */
function checkSetup() {
  const ss=tpOpenSheet_();
  const studentSheet=ss.getSheetByName(TP_SETTINGS.roster);
  const names=studentSheet.getRange(2,1,studentSheet.getLastRow()-1,2).getDisplayValues();
  const valid=TP_STUDENTS.every(s=>names.some(r=>r[0]===s.name&&tpNormalizeId_(r[1])===tpNormalizeId_(s.id)));
  if(!valid) throw new Error('「互評正式名單2026」缺漏或學號不符，請執行 setup()。');
  Logger.log('檢查完成：'+TP_STUDENTS.length+' 人；第一週免評 11 人；第二週 12 人；第三週 11 人。');
  return 'v5.2 已連結 '+ss.getName()+'；34 位學號姓名匹配，舊名單未更動。';
}
