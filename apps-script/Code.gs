/**
 * AI융합학부 통합 브리핑 — 질문·의견 게시판 백엔드 (Google Apps Script)
 *
 * 이 스크립트는 처음 setup()을 실행할 때 만든 스프레드시트의 '응답' 시트에 글을 저장합니다.
 * - 이름·학번(전체)은 시트에만 저장되고, 페이지(doGet)에는 절대 내보내지 않습니다.
 * - 페이지에는 입학 학번(2X학번)·전공·학기수·종류·주제·내용·좋아요 수만 공개됩니다.
 * - '공개' 열을 N으로 바꾸면 그 글은 페이지에서 숨겨집니다.
 * - '답변' 열에 글을 쓰면 페이지에 "대표 답변"으로 함께 표시됩니다.
 *
 * 배포: 배포 → 새 배포 → 유형 '웹 앱' → 실행 사용자 '나' → 액세스 권한 '모든 사용자'
 */

var SHEET_NAME = '응답';
var HEAD = ['ID', '작성 시각', '종류', '주제', '내용', '이름', '학번', '입학 학번', '전공', '학기수', '좋아요', '공개', '답변'];
var COL = { id: 1, time: 2, type: 3, topic: 4, text: 5, name: 6, sid: 7, year: 8, major: 9, sem: 10, likes: 11, pub: 12, reply: 13 };

var TYPES = ['질문', '의견', '정정'];
var TOPICS = ['전공 존속·교육과정', '캡스톤·졸업요건', '시행 시기·선적용', '학부명·영문명', '경과조치·학부명 선택', '대표단·예산·공간 승계', '창업 교육·동문', 'GE 관련', '기타'];
var MAJORS = ['AI융합', '데이터사이언스', 'GE', 'ICT융합', 'ACE', '전공 미선택', '기타'];

function setup() {
  var sh = sheet_();
  sh.getRange(1, 1, 1, HEAD.length).setFontWeight('bold').setBackground('#EAF2FD');
  sh.setFrozenRows(1);
  sh.setColumnWidth(COL.text, 420);
  sh.setColumnWidth(COL.reply, 320);
  sh.getRange('E:E').setWrap(true);
  sh.getRange('M:M').setWrap(true);
  Logger.log(sh.getParent().getUrl());
}

function sheet_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SHEET_ID');
  var ss = id ? SpreadsheetApp.openById(id) : null;
  if (!ss) {
    ss = SpreadsheetApp.create('AI융합학부 통합 브리핑 - 질문·의견 응답');
    props.setProperty('SHEET_ID', ss.getId());
    ss.getSheets()[0].getRange(1, 1, 1, HEAD.length).setValues([HEAD]);
  }
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh && String(ss.getSheets()[0].getRange('A1').getValue()) === 'ID') {
    sh = ss.getSheets()[0].setName(SHEET_NAME);
  }
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME, 0);
    sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight('bold').setBackground('#EAF2FD');
    sh.setFrozenRows(1);
    sh.setColumnWidth(COL.text, 420);
    sh.setColumnWidth(COL.reply, 320);
    sh.getRange('E:E').setWrap(true);
    sh.getRange('M:M').setWrap(true);
  }
  return sh;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// 엑셀·시트에서 수식으로 해석되지 않도록
function safe_(s, max) {
  s = String(s == null ? '' : s).replace(/\r\n?/g, '\n').trim().slice(0, max);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function pick_(v, list) { return list.indexOf(v) >= 0 ? v : ''; }

function doGet() {
  var sh = sheet_();
  var n = sh.getLastRow() - 1;
  var items = [];
  if (n > 0) {
    var rows = sh.getRange(2, 1, n, HEAD.length).getValues();
    rows.forEach(function (r) {
      if (!r[0] || String(r[COL.pub - 1]).trim().toUpperCase() === 'N') return;
      var t = r[COL.time - 1];
      items.push({
        id: String(r[COL.id - 1]),
        t: t instanceof Date ? t.toISOString() : String(t),
        type: r[COL.type - 1], topic: r[COL.topic - 1], text: String(r[COL.text - 1]).replace(/^'/, ''),
        year: r[COL.year - 1], major: r[COL.major - 1], sem: r[COL.sem - 1],
        likes: Number(r[COL.likes - 1]) || 0,
        reply: String(r[COL.reply - 1] || '')
      });
      // 이름(F열)·학번(G열)은 의도적으로 내보내지 않음
    });
  }
  items.reverse();
  return json_({ ok: true, items: items });
}

function doPost(e) {
  var d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad request' }); }
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var sh = sheet_();
    if (d.action === 'like') return like_(sh, d);
    if (d.hp) return json_({ ok: true }); // 스팸 봇용 숨은 칸

    var type = pick_(d.type, TYPES), topic = pick_(d.topic, TOPICS), major = pick_(d.major, MAJORS);
    var text = safe_(d.text, 2000), name = safe_(d.name, 40), sid = safe_(d.sid, 20);
    var year = String(d.year || '').match(/^\d{2}$/) ? d.year + '학번' : '';
    var sem = String(d.sem || '').match(/^(\d{1,2}|9\+|졸업)$/) ? (d.sem === '졸업' ? '졸업' : d.sem + '학기') : '';
    if (!type || !topic || !major || !text || !name || !sid || !year || !sem) {
      return json_({ ok: false, error: '빠진 항목이 있습니다.' });
    }
    var id = Utilities.getUuid().replace(/-/g, '').slice(0, 10);
    sh.appendRow([id, new Date(), type, topic, text, name, "'" + sid, year, major, sem, 0, 'Y', '']);
    return json_({ ok: true, id: id });
  } finally {
    lock.releaseLock();
  }
}

function like_(sh, d) {
  var n = sh.getLastRow() - 1;
  if (n < 1) return json_({ ok: false, error: 'not found' });
  var ids = sh.getRange(2, COL.id, n, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(d.id)) {
      var c = sh.getRange(i + 2, COL.likes);
      var v = Math.max(0, (Number(c.getValue()) || 0) + (d.undo ? -1 : 1));
      c.setValue(v);
      return json_({ ok: true, likes: v });
    }
  }
  return json_({ ok: false, error: 'not found' });
}
