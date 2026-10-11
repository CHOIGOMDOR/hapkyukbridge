/**
 * 치즈볼의 덕질 공부 노트 - 구글 앱 스크립트 (웹페이지 + 시트 연결)
 *
 * [파일 구성] 이 프로젝트에 파일 2개를 만들어요
 *   Code.gs   ← 이 파일
 *   Index.html ← 저장소의 apps-script/Index.html 내용을 그대로 붙여넣기 (파일 추가 > HTML > 이름 "Index")
 *
 * [시트] 구글 시트에 탭 3개 (1행은 제목줄)
 *   계획 : 날짜 | 과목 | 단원 | 항목 | 범위 | 링크 | 완료
 *   기록 : 날짜 | 기분 | 개념1 | 개념2 | 개념3 | 다음공부 | 전사요약 | 사진
 *   설정 : A열 이름 / B열 값 → 목표제목, 시작일, 목표일, 목표설명, 이미지폴더ID, 사진폴더ID, 폰트파일ID
 *          (폰트파일ID = 드라이브에 올린 MemomentKkukkukk.woff2 파일의 ID)
 *
 * [배포] 배포 > 새 배포 > 유형: 웹 앱 > 실행: 나 / 액세스: 모든 사용자
 *        → 나오는 웹 앱 URL이 곧 페이지 주소예요. 카톡으로 보내면 열려요.
 *        코드를 고친 뒤에는 "배포 관리 > 수정 > 새 버전"으로 다시 배포해야 반영돼요.
 */
var TZ = 'Asia/Seoul';
var LOG_COLS = ['날짜', '기분', '개념1', '개념2', '개념3', '다음공부', '전사요약', '사진'];

function sheet_(name) { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name); }
function ymd_(v) { return v instanceof Date ? Utilities.formatDate(v, TZ, 'yyyy-MM-dd') : String(v || ''); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function settings_() {
  var o = {}, v = sheet_('설정').getDataRange().getValues();
  for (var i = 0; i < v.length; i++) if (v[i][0]) o[v[i][0]] = v[i][1];
  return o;
}

/** 웹 주소로 들어오면 페이지를 보여줘요. (?action=all 이면 데이터만 JSON으로) */
function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'all') return json_(getAll_());
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('치즈볼의 덕질 공부 노트')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** 페이지(google.script.run)가 부르는 함수들 */
function apiGet() { return getAll_(); }
function apiPost(bodyJson) { return handle_(JSON.parse(bodyJson)); }
function getFontBase64() {
  var id = settings_()['폰트파일ID'];
  if (!id) return '';
  return Utilities.base64Encode(DriveApp.getFileById(id).getBlob().getBytes());
}

function getAll_() {
  var st = settings_();
  var plan = sheet_('계획').getDataRange().getValues().slice(1).map(function (r, i) {
    return { id: i + 2, date: ymd_(r[0]), subject: r[1], unit: r[2], task: r[3], range: r[4], link: r[5], done: r[6] === true || r[6] === 'TRUE' || r[6] === '완료' };
  }).filter(function (r) { return r.date; });
  var log = sheet_('기록').getDataRange().getValues().slice(1).map(function (r) {
    return { date: ymd_(r[0]), mood: r[1], c1: r[2], c2: r[3], c3: r[4], next: r[5], summary: r[6], photos: String(r[7] || '').split(',').filter(String) };
  }).filter(function (r) { return r.date; });
  var images = [];
  if (st['이미지폴더ID']) {
    var it = DriveApp.getFolderById(st['이미지폴더ID']).getFiles();
    while (it.hasNext()) {
      var f = it.next();
      if (/^image\//.test(f.getMimeType())) images.push('https://drive.google.com/thumbnail?id=' + f.getId() + '&sz=w600');
    }
  }
  return {
    plan: plan, log: log, images: images,
    goal: { title: st['목표제목'] || '', start: ymd_(st['시작일']), end: ymd_(st['목표일']), sub: st['목표설명'] || '' }
  };
}

function handle_(b) {
  if (b.action === 'check') {
    sheet_('계획').getRange(b.row, 7).setValue(b.done ? '완료' : '');
  } else if (b.action === 'mood') {
    upsertLog(b.date, { 기분: b.mood });
  } else if (b.action === 'photo') {
    var folder = DriveApp.getFolderById(settings_()['사진폴더ID']);
    var blob = Utilities.newBlob(Utilities.base64Decode(b.data), b.mime, b.date + '_' + b.name);
    var f = folder.createFile(blob);
    f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var url = 'https://drive.google.com/thumbnail?id=' + f.getId() + '&sz=w600';
    var cur = (getLog_(b.date) || {})['사진'];
    upsertLog(b.date, { 사진: cur ? cur + ',' + url : url });
  }
  return { ok: true };
}
function doPost(e) { return json_(handle_(JSON.parse(e.postData.contents))); }

function getLog_(date) {
  var v = sheet_('기록').getDataRange().getValues();
  for (var i = 1; i < v.length; i++) if (ymd_(v[i][0]) === date) {
    var o = {}; LOG_COLS.forEach(function (c, k) { o[c] = v[i][k]; }); return o;
  }
  return null;
}

/** 날짜 행이 있으면 갱신, 없으면 추가. 클로바 전사 트리거 코드에서도 이 함수를 쓰면 돼요.
 *  예) upsertLog('2026-10-13', {개념1:'...', 개념2:'...', 개념3:'...', 다음공부:'...', 전사요약:'...'}) */
function upsertLog(date, fields) {
  var s = sheet_('기록'), v = s.getDataRange().getValues(), row = -1;
  for (var i = 1; i < v.length; i++) if (ymd_(v[i][0]) === date) { row = i + 1; break; }
  if (row < 0) { s.appendRow([date]); row = s.getLastRow(); s.getRange(row, 1).setNumberFormat('@').setValue(date); }
  for (var k in fields) {
    var c = LOG_COLS.indexOf(k);
    if (c >= 0) s.getRange(row, c + 1).setValue(fields[k]);
  }
}
