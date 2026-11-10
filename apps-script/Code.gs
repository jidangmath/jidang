/**
 * 지당 징검다리 - 구글 시트 자동 받기
 * 1) 아래 두 값을 선생님이 정해서 바꿔 주세요. (다른 사람에게 알려 주지 마세요: TEACHER_KEY)
 * 2) 배포 > 새 배포 > 웹 앱 > 실행: 나, 액세스: 모든 사용자
 */
var CLASS_CODE = '여기에-학생에게-알려줄-반-코드';   // 학생 링크에 들어가요
var TEACHER_KEY = '여기에-선생님만-아는-비밀-키';     // 교사용 화면에서만 써요
var SHEET = '현황';

function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET) || ss.insertSheet(SHEET);
  if (sh.getLastRow() === 0) {
    sh.appendRow(['별명', '시작일', '기록한 날', '연속', '최장 연속', '칭호', '퀘스트', '최근 에너지', '마지막 기록', '쿠폰', '보고 시각', '코드']);
  }
  return sh;
}

function decode_(code) {
  if (typeof code !== 'string' || code.indexOf('JD1.') !== 0 || code.length > 4000) return null;
  try {
    var bytes = Utilities.base64Decode(code.slice(4));
    var o = JSON.parse(Utilities.newBlob(bytes).getDataAsString('UTF-8'));
    if (!o || typeof o.n !== 'string' || typeof o.r !== 'number') return null;
    return o;
  } catch (e) { return null; }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var body = JSON.parse(e.postData.contents);
    if (body.k !== CLASS_CODE) return out_({ ok: false });
    var o = decode_(body.c);
    if (!o) return out_({ ok: false });
    var sh = sheet_();
    var rows = sh.getDataRange().getValues();
    var key = o.n + '|' + o.s, at = -1;
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0] + '|' + rows[i][1] === key) { at = i + 1; break; }
    }
    var cp = (o.c || []).map(function (c) { return c[0] + '일:' + (c[3] ? '사용' : '미사용'); }).join(' ');
    var row = [o.n, o.s, o.r, o.st, o.lg, o.lv, o.q, o.e, o.l, cp, new Date(o.t), body.c];
    if (at > 0) sh.getRange(at, 1, 1, row.length).setValues([row]);
    else sh.appendRow(row);
    return out_({ ok: true });
  } catch (err) {
    return out_({ ok: false });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function doGet(e) {
  if (!e || !e.parameter || e.parameter.key !== TEACHER_KEY) return out_({ ok: false });
  var rows = sheet_().getDataRange().getValues();
  var codes = [];
  for (var i = 1; i < rows.length; i++) if (rows[i][11]) codes.push(rows[i][11]);
  return out_({ ok: true, codes: codes });
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
