/**
 * 체험 대기 호출 — 구글시트 백엔드 (Google Apps Script)
 *
 * [설치]
 * 1) 구글 드라이브에서 새 스프레드시트를 만듭니다.
 * 2) 확장 프로그램 → Apps Script 를 열고, 이 파일 내용을 전부 붙여넣습니다.
 * 3) 배포 → 새 배포 → 유형: 웹 앱
 *      - 실행 계정: 나
 *      - 액세스 권한이 있는 사용자: 모든 사용자
 * 4) 만들어진 https://script.google.com/macros/s/.../exec 주소를
 *    체험대기호출.html 의 [설정 → 구글시트 연결] 칸에 붙여넣습니다.
 *
 * 코드를 수정하면 반드시 '새 배포'(새 버전)로 다시 배포해야 반영됩니다.
 * 시트의 '신청자'/'설정' 탭은 처음 호출될 때 자동으로 만들어집니다.
 */

var SHEET_ID = '';            // 비워두면 이 스크립트가 붙어 있는 시트를 사용합니다
var SHEET_ENTRIES = '신청자';
var SHEET_CONFIG = '설정';

var HEAD = ['id', '체험일', '대기번호', '이름', '연락처', '인원', '프로그램', '상태', '신청시각', '호출시각', '동의시각', '정렬키'];
var STATE_KO = { waiting: '대기', called: '호출', done: '입장완료', cancel: '취소' };
var STATE_EN = { '대기': 'waiting', '호출': 'called', '입장완료': 'done', '취소': 'cancel' };

var CFG_ROWS = [
  ['행사명', 'title', '체험 프로그램 접수'],
  ['장소', 'place', ''],
  ['시작일', 'startDate', ''],
  ['종료일', 'endDate', ''],
  ['하루 체험 가능 인원', 'dailyCap', 40],
  ['정원 기준(인원/팀)', 'capMode', '인원'],
  ['동시 진행 자리 수', 'lanes', 1],
  ['연락처 수집(전체/뒤4자리)', 'phoneMode', '전체'],
  ['대기 화면 안내', 'notice', '호출되면 이 화면이 노란색으로 바뀌고 알림이 울립니다. 부스 근처에서 대기해 주세요.'],
  ['개인정보 동의 문구', 'consentText', '체험 호출을 위한 기본 정보만 수집하며 전시회 종료 후 폐기합니다.'],
  ['체험 프로그램(JSON)', 'programs', '[{"id":"p1","name":"체험 프로그램","minutes":15}]']
];

/* ============ 엔트리 포인트 ============ */
function doGet(e) {
  try {
    var action = (e && e.parameter && e.parameter.action) || 'get';
    if (action === 'sheet') {
      return HtmlService.createHtmlOutput(
        '<script>location.href="' + book().getUrl() + '";</script>연결된 시트로 이동합니다…'
      );
    }
    return json(loadAll());
  } catch (err) { return json({ error: String(err.message || err) }); }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);
    var p = JSON.parse(e.postData.contents);
    switch (p.action) {
      case 'join':   return json(join(p.entry));
      case 'patch':  return json(patch(p.id, p.patch));
      case 'config': return json(saveConfig(p.config));
      case 'reset':  return json(resetEntries());
      default:       return json({ error: '알 수 없는 요청입니다: ' + p.action });
    }
  } catch (err) {
    return json({ error: String(err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ============ 시트 준비 ============ */
function book() { return SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet(); }

function entriesSheet() {
  var ss = book(), sh = ss.getSheetByName(SHEET_ENTRIES);
  if (!sh) {
    sh = ss.insertSheet(SHEET_ENTRIES);
    sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight('bold').setBackground('#F1EFF4');
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 120); sh.setColumnWidth(4, 110); sh.setColumnWidth(5, 130);
    sh.hideColumns(12); // 정렬키
  }
  return sh;
}

function configSheet() {
  var ss = book(), sh = ss.getSheetByName(SHEET_CONFIG);
  if (!sh) {
    sh = ss.insertSheet(SHEET_CONFIG);
    sh.getRange(1, 1, 1, 3).setValues([['항목', '값', '설명']]).setFontWeight('bold').setBackground('#F1EFF4');
    sh.setFrozenRows(1);
    var rows = CFG_ROWS.map(function (r) { return [r[0], r[2], '']; });
    sh.getRange(2, 1, rows.length, 3).setValues(rows);
    sh.getRange(2, 1, rows.length, 1).setFontWeight('bold');
    sh.setColumnWidth(1, 200); sh.setColumnWidth(2, 420);
    sh.getRange(2, 2, rows.length, 1).setNumberFormat('@'); // 날짜 자동변환 방지
  }
  return sh;
}

/* ============ 읽기 ============ */
function loadAll() {
  return { config: readConfig(), entries: readEntries(), seq: 0 };
}

function readConfig() {
  var sh = configSheet();
  var vals = sh.getRange(2, 1, Math.max(1, sh.getLastRow() - 1), 2).getValues();
  var map = {};
  vals.forEach(function (r) { if (r[0]) map[String(r[0]).trim()] = r[1]; });

  var cfg = {};
  CFG_ROWS.forEach(function (row) {
    var label = row[0], key = row[1], def = row[2];
    var v = map.hasOwnProperty(label) && map[label] !== '' ? map[label] : def;
    if (key === 'startDate' || key === 'endDate') cfg[key] = v ? ymd(v) : '';
    else if (key === 'dailyCap' || key === 'lanes') cfg[key] = Number(v) || Number(def);
    else if (key === 'capMode') cfg[key] = String(v).indexOf('팀') > -1 ? 'teams' : 'people';
    else if (key === 'phoneMode') cfg[key] = String(v).indexOf('4') > -1 ? 'last4' : 'full';
    else if (key === 'programs') {
      try { cfg[key] = JSON.parse(v); } catch (e) { cfg[key] = JSON.parse(def); }
      if (!cfg[key] || !cfg[key].length) cfg[key] = JSON.parse(def);
    }
    else cfg[key] = String(v);
  });
  cfg.askPeople = true;
  return cfg;
}

function readEntries() {
  var sh = entriesSheet();
  if (sh.getLastRow() < 2) return [];
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, HEAD.length).getValues();
  var out = [];
  for (var i = 0; i < vals.length; i++) {
    var r = vals[i];
    if (!r[0]) continue;
    out.push({
      id: String(r[0]),
      date: ymd(r[1]),
      num: Number(r[2]) || 0,
      name: String(r[3]),
      phone: String(r[4]),
      people: Number(r[5]) || 1,
      programId: String(r[6]),
      state: STATE_EN[String(r[7])] || 'waiting',
      ts: ms(r[8]),
      calledAt: ms(r[9]),
      consentAt: ms(r[10]),
      order: Number(r[11]) || ms(r[8])
    });
  }
  return out;
}

function ymd(v) {
  if (v instanceof Date) return Utilities.formatDate(v, tz(), 'yyyy-MM-dd');
  var s = String(v || '').trim();
  return s ? s.slice(0, 10) : '';
}
function ms(v) {
  if (v instanceof Date) return v.getTime();
  var n = Number(v);
  if (n > 0) return n;
  var d = new Date(v);
  return isNaN(d.getTime()) ? 0 : d.getTime();
}
function tz() { return book().getSpreadsheetTimeZone() || 'Asia/Seoul'; }

/* ============ 쓰기 ============ */
function join(entry) {
  var cfg = readConfig();
  var rows = readEntries();
  var date = entry.date || Utilities.formatDate(new Date(), tz(), 'yyyy-MM-dd');
  var people = Number(entry.people) || 1;

  // 정원 확인 (기간이 등록된 경우에만)
  if (cfg.startDate && cfg.endDate) {
    if (date < cfg.startDate || date > cfg.endDate) return { error: '체험 기간이 아닌 날짜입니다.' };
    var used = 0;
    rows.forEach(function (r) {
      if (r.date === date && r.state !== 'cancel') used += (cfg.capMode === 'teams' ? 1 : (r.people || 1));
    });
    var need = cfg.capMode === 'teams' ? 1 : people;
    if (used + need > cfg.dailyCap) return { error: '선택한 날짜는 마감되었습니다. 다른 날짜를 선택해 주세요.' };
  }

  // 대기번호는 날짜별로 1번부터
  var num = 0;
  rows.forEach(function (r) { if (r.date === date && r.num > num) num = r.num; });
  num += 1;

  var now = new Date();
  entriesSheet().appendRow([
    entry.id, date, num, entry.name, "'" + String(entry.phone || ''), people,
    entry.programId, STATE_KO.waiting, now, '', now, now.getTime()
  ]);
  return { ok: true, num: num, date: date };
}

function patch(id, p) {
  var sh = entriesSheet();
  if (sh.getLastRow() < 2) return { error: '신청 내역을 찾을 수 없습니다.' };
  var ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) {
      var row = i + 2;
      if (p.state) sh.getRange(row, 8).setValue(STATE_KO[p.state] || p.state);
      if (p.calledAt) sh.getRange(row, 10).setValue(new Date(p.calledAt));
      if (p.order) sh.getRange(row, 12).setValue(p.order);
      return { ok: true };
    }
  }
  return { error: '신청 내역을 찾을 수 없습니다.' };
}

function saveConfig(config) {
  var sh = configSheet();
  var last = Math.max(2, sh.getLastRow());
  var labels = sh.getRange(2, 1, last - 1, 1).getValues().map(function (r) { return String(r[0]).trim(); });

  CFG_ROWS.forEach(function (row) {
    var label = row[0], key = row[1];
    if (!config.hasOwnProperty(key)) return;
    var v = config[key];
    if (key === 'capMode') v = (v === 'teams' ? '팀' : '인원');
    if (key === 'phoneMode') v = (v === 'last4' ? '뒤4자리' : '전체');
    if (key === 'programs') v = JSON.stringify(v);
    var idx = labels.indexOf(label);
    if (idx === -1) { sh.appendRow([label, v, '']); labels.push(label); }
    else sh.getRange(idx + 2, 2).setValue(v);
  });
  return { ok: true };
}

function resetEntries() {
  var sh = entriesSheet();
  if (sh.getLastRow() > 1) sh.deleteRows(2, sh.getLastRow() - 1);
  return { ok: true };
}
