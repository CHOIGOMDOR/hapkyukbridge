/**
 * 치즈볼의 덕질 공부 노트 - 구글 앱 스크립트 (웹페이지 + 시트 연결)
 *
 * [처음 한 번만]
 *  1. 새 구글 시트를 하나 만들어요 (이름은 자유).
 *  2. 시트에서 확장 프로그램 > Apps Script 를 열어요.
 *  3. Code.gs 에 이 파일을 붙여넣고, 파일 추가(+) > HTML > 이름 "Index" 에 Index.html 내용을 붙여넣어요.
 *  4. 위쪽 함수 선택칸에서 setupSheets 를 고르고 ▶ 실행을 눌러요 (권한 허용).
 *     → 계획·단어장·기록·순공시간 등 필요한 탭이 알아서 만들어지고, 계획·단어장·한국사 문제가 채워져요.
 *  5. 배포 > 새 배포 > 유형: 웹 앱 > 실행: 나 / 액세스: 모든 사용자 > 배포
 *     → 나오는 웹 앱 주소가 곧 페이지 주소예요. 카톡으로 보내면 열려요.
 *  코드를 고친 뒤에는 "배포 관리 > 수정 > 새 버전"으로 다시 배포해야 반영돼요.
 *
 * [시트 탭] (setupSheets 가 만들어요)
 *  계획 : 날짜 | 과목 | 단원 | 항목 | 범위 | 링크 | 완료      ← 오늘 할 일 (완료 칸은 페이지가 채워요)
 *  단어장 : 유닛 | 유닛 제목 | 영어 | 품사 | 뜻 | 철자 시험 | 선생님 표시 | 시험일
 *  반복항목 : 항목 | 요일 | 주횟수 | 과목 | 메모           ← 요일이 "자유"면 주 N회 아무 날이나 (큐티)
 *  문제 : 과목 | 키 | 문제 | 보기1~4 | 정답(1~4) | 해설      ← 복습 퀴즈 문제 (키 = 항목 이름 또는 단원)
 *  기록 / 순공시간 / 반복체크 / 퀴즈기록 / 설정            ← 페이지가 쌓아요
 */
var TZ = 'Asia/Seoul';
// setupSheets 가 처음 채울 때 복사해 오는 원본 시트 (드라이브에 있는 파일 ID)
var SRC_PLAN = '1bivWFZdjZ2QOfV4z6-5V62wHkecc--fYdoX5MqcmLdc';   // 2026 학습 계획 v3
var SRC_VOCAB = '18gtRs_9dRpiAd1CkopYkyU2Mm8u2xOYS58DE7HSjW2k';  // 영어 단어장 전체 1~20
var SRC_QUESTIONS = '1OqEnDsWCYLbsU69oH_62o8X_2x7NVnmh9SNpRGpg0bI'; // 한국사 복습 문제 1~2권
// 새 자료가 생기면 아래 목록에 { tab: '넣을 탭', id: '재료 시트 ID' } 한 줄만 더하고 setupSheets 를 다시 실행해요.
// (이미 들어 있는 줄은 건너뛰어서, 몇 번을 실행해도 중복되지 않아요)
var EXTRA_SOURCES = [
  { tab: '문제', id: '1kNgeNBm_Q7GU3HtAm7yg86vbjdOqVAV1p2RGUemwQYk' },  // 한국사 복습 문제 3~6권
  { tab: '문제', id: '1Kio-FP1hnkjfQv5h5XCyGK3FyJTkFZAF9uGyBIalLuE' }   // 과학 복습 문제 6-2
];
var LOG_COLS = ['날짜', '기분', '개념1', '개념2', '개념3', '다음공부', '전사요약', '사진'];

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function sheet_(name) { return ss_().getSheetByName(name); }
function ymd_(v) { return v instanceof Date ? Utilities.formatDate(v, TZ, 'yyyy-MM-dd') : String(v || '').slice(0, 10); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function rows_(name) { var s = sheet_(name); return s ? s.getDataRange().getValues().slice(1) : []; }

/* ---------------- 처음 한 번: 탭 만들고 채우기 ---------------- */
function setupSheets() {
  var ss = ss_();
  importSheet_(ss, '계획', SRC_PLAN, 60);
  importSheet_(ss, '단어장', SRC_VOCAB, 100);
  importSheet_(ss, '문제', SRC_QUESTIONS, 20);
  EXTRA_SOURCES.forEach(function (x) { appendMore_(ss, x.tab, x.id); });
  ensure_(ss, '기록', LOG_COLS);
  var qt = ensure_(ss, '모르는문제', ['ID', '날짜', '사진URL', '질문', '답변(JSON)', '파일ID', '답변 보기']);
  qt.getRange(1, 4).setValue('질문'); qt.getRange(1, 7).setValue('답변 보기').setFontWeight('bold');
  ensure_(ss, '순공시간', ['날짜', '분', '입력시각']);
  ensure_(ss, '반복체크', ['날짜', '항목']);
  ensure_(ss, '퀴즈기록', ['날짜', '종류', '키', '정답']);
  ensure_(ss, '문제', ['과목', '키', '문제', '보기1', '보기2', '보기3', '보기4', '정답', '해설']); // 위에서 못 가져왔을 때만 빈 탭
  var rep = ensure_(ss, '반복항목', ['항목', '요일', '주횟수', '과목', '메모']);
  if (rep.getLastRow() < 2) {
    rep.appendRow(['큐티', '자유', 2, '큐티', '일요일 모임 준비 · 이번 주 2번']);
    rep.appendRow(['주제 글쓰기', '금', 1, '학교 숙제', '학교 글쓰기 숙제']);
  }
  var st = ensure_(ss, '설정', ['이름', '값']);
  var have = {}; st.getDataRange().getValues().forEach(function (r) { have[r[0]] = true; });
  [['목표제목', '중1 수학 1학기 완주'], ['시작일', '2026-10-12'], ['목표일', '2027-01-20'],
   ['목표설명', '2026.10.12 → 2027.1월 중후반 · 내 속도에 맞춰 분량을 조절해요.'],
   ['이미지폴더ID', ''], ['사진폴더ID', ''], ['폰트파일ID', ''],
   ['전사폴더ID', '137o9lPGy6DLn7x-LRxuEFDIQpNABjypv'], ['API키', ''], ['AI모델', 'gemini-2.5-flash']].forEach(function (r) { if (!have[r[0]]) st.appendRow(r); });
  ['Sheet1', '시트1'].forEach(function (n) { var d = ss.getSheetByName(n); if (d && ss.getSheets().length > 1 && d.getLastRow() === 0) ss.deleteSheet(d); });
  SpreadsheetApp.getUi && Logger.log('완료! 이제 배포 > 새 배포 > 웹 앱 으로 배포하세요.');
}
function ensure_(ss, name, header) {
  var s = ss.getSheetByName(name);
  if (!s) { s = ss.insertSheet(name); s.appendRow(header); s.setFrozenRows(1); s.getRange(1, 1, 1, header.length).setFontWeight('bold'); }
  return s;
}
/** 같은 형식의 다른 시트를 탭 맨 아래에 이어 붙여요. 이미 있는 줄(키·문제가 같은 줄)은 건너뛰어요 */
function appendMore_(ss, name, srcId) {
  var s = ss.getSheetByName(name);
  if (!s) return;
  var have = {};
  s.getDataRange().getValues().slice(1).forEach(function (r) { have[r[1] + '|' + r[2]] = true; });
  var add = SpreadsheetApp.openById(srcId).getSheets()[0].getDataRange().getValues().slice(1)
    .filter(function (r) { return r[1] && !have[r[1] + '|' + r[2]]; });
  if (add.length) s.getRange(s.getLastRow() + 1, 1, add.length, add[0].length).setValues(add);
  Logger.log(name + ': ' + add.length + '줄 추가');
}
/** 이미 데이터가 충분히 있으면(minRows 이상) 덮어쓰지 않아요 → 다시 실행해도 체크 기록이 안 지워져요 */
function importSheet_(ss, name, srcId, minRows) {
  var old = ss.getSheetByName(name);
  if (old && old.getLastRow() >= minRows) { Logger.log(name + ': 이미 있어서 그대로 둬요'); return; }
  var src = SpreadsheetApp.openById(srcId).getSheets()[0];
  var copy = src.copyTo(ss);
  if (old) ss.deleteSheet(old);
  copy.setName(name);
  copy.setFrozenRows(1);
}

/* ---------------- 웹 주소로 들어오면 페이지 ---------------- */
function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'all') return json_(getAll_());
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('치즈볼의 덕질 공부 노트')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function doPost(e) { return json_(handle_(JSON.parse(e.postData.contents))); }

/** 페이지(google.script.run)가 부르는 함수들 */
function apiGet() { return getAll_(); }
function apiPost(bodyJson) { return handle_(JSON.parse(bodyJson)); }
function getFontBase64() {
  var id = settings_()['폰트파일ID'];
  if (!id) return '';
  return Utilities.base64Encode(DriveApp.getFileById(id).getBlob().getBytes());
}
function settings_() {
  var o = {}; rows_('설정').forEach(function (r) { if (r[0]) o[r[0]] = r[1]; }); return o;
}

function getAll_() {
  var st = settings_(), since = ymd_(new Date(Date.now() - 150 * 864e5));
  var plan = rows_('계획').map(function (r, i) {
    return { id: i + 2, date: ymd_(r[0]), subject: r[1], unit: r[2], task: r[3], range: r[4], link: r[5], done: r[6] === true || r[6] === 'TRUE' || r[6] === '완료' };
  }).filter(function (r) { return r.date; });
  var log = rows_('기록').map(function (r) {
    return { date: ymd_(r[0]), mood: r[1], c1: r[2], c2: r[3], c3: r[4], next: r[5], summary: r[6], photos: String(r[7] || '').split(',').filter(String) };
  }).filter(function (r) { return r.date; });
  var vocab = rows_('단어장').filter(function (r) { return r[2]; })
    .map(function (r) { return [Number(r[0]), r[1], String(r[2]), String(r[3]), String(r[4]), String(r[5]), String(r[6] || '')]; });
  var repeats = rows_('반복항목').filter(function (r) { return r[0]; })
    .map(function (r) { return { name: String(r[0]), day: String(r[1] || '자유'), times: Number(r[2]) || 1, subject: String(r[3] || ''), memo: String(r[4] || '') }; });
  var rchecks = rows_('반복체크').filter(function (r) { return r[0] && ymd_(r[0]) >= since; })
    .map(function (r) { return { date: ymd_(r[0]), name: String(r[1]) }; });
  var minutes = rows_('순공시간').filter(function (r) { return r[0] && ymd_(r[0]) >= since; })
    .map(function (r) { return { date: ymd_(r[0]), min: Number(r[1]) || 0 }; });
  var qlog = rows_('퀴즈기록').filter(function (r) { return r[0] && ymd_(r[0]) >= since; })
    .map(function (r) { return { date: ymd_(r[0]), kind: String(r[1]), key: String(r[2]), ok: r[3] === 1 || r[3] === true || r[3] === '1' || r[3] === 'TRUE' }; });
  var questions = rows_('문제').filter(function (r) { return r[1] && r[2]; })
    .map(function (r) { return { subject: String(r[0]), key: String(r[1]), q: String(r[2]), opts: [r[3], r[4], r[5], r[6]].map(String), ans: Number(r[7]), exp: String(r[8] || '') }; });
  var qa = rows_('모르는문제').filter(function (r) { return r[0] && r[2]; }).map(function (r) {
    var ans = []; try { ans = JSON.parse(r[4] || '[]'); } catch (e) {}
    return { id: String(r[0]), date: ymd_(r[1]), url: String(r[2]), memo: String(r[3] || ''), answers: ans };
  });
  var images = [];
  if (st['이미지폴더ID']) {
    var it = DriveApp.getFolderById(st['이미지폴더ID']).getFiles();
    while (it.hasNext()) {
      var f = it.next();
      if (/^image\//.test(f.getMimeType())) images.push('https://drive.google.com/thumbnail?id=' + f.getId() + '&sz=w600');
    }
  }
  return {
    plan: plan, log: log, vocab: vocab, repeats: repeats, rchecks: rchecks, minutes: minutes, qlog: qlog, questions: questions, qa: qa, images: images,
    goal: { title: st['목표제목'] || '', start: ymd_(st['시작일']), end: ymd_(st['목표일']), sub: st['목표설명'] || '' }
  };
}

/* ---------------- 페이지에서 오는 저장 요청 ---------------- */
function handle_(b) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    if (b.action === 'check') {
      sheet_('계획').getRange(b.row, 7).setValue(b.done ? '완료' : '');
    } else if (b.action === 'mood') {
      upsertLog(b.date, { 기분: b.mood });
    } else if (b.action === 'photo') {
      var folder = DriveApp.getFolderById(settings_()['사진폴더ID']);
      var blob = Utilities.newBlob(Utilities.base64Decode(b.data), b.mime, b.date + '_' + b.name);
      var f = folder.createFile(blob);
      f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      var url = 'https://drive.google.com/thumbnail?id=' + f.getId() + '&sz=w1200';
      var id = b.id || ('q' + Date.now());
      appendText_('모르는문제', [id, b.date, url, b.memo || '', '[]', f.getId()]);
      return { ok: true, id: id, url: url };
    } else if (b.action === 'qaAnswer') {
      var qs = sheet_('모르는문제'), qv = qs.getDataRange().getValues(), entry = { who: b.who, text: b.text || '', at: b.at, img: '', fid: '' };
      if (b.data) {
        var af = DriveApp.getFolderById(settings_()['사진폴더ID']).createFile(Utilities.newBlob(Utilities.base64Decode(b.data), b.mime, b.date + '_답변_' + b.name));
        af.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        entry.img = 'https://drive.google.com/thumbnail?id=' + af.getId() + '&sz=w1200'; entry.fid = af.getId();
      }
      for (var i = 1; i < qv.length; i++) if (String(qv[i][0]) === String(b.id)) {
        var arr = []; try { arr = JSON.parse(qv[i][4] || '[]'); } catch (e) {}
        arr.push(entry);
        qs.getRange(i + 1, 5).setValue(JSON.stringify(arr));
        qs.getRange(i + 1, 7).setValue(arr.map(function (x) { return '[' + x.who + ' ' + x.at + '] ' + (x.text || '') + (x.img ? ' (사진: ' + x.img + ')' : ''); }).join('\n'));
        break;
      }
      return { ok: true, img: entry.img };
    } else if (b.action === 'qaDelete') {
      var ds = sheet_('모르는문제'), dv = ds.getDataRange().getValues();
      for (var j = dv.length - 1; j >= 1; j--) if (String(dv[j][0]) === String(b.id)) {
        try { if (dv[j][5]) DriveApp.getFileById(String(dv[j][5])).setTrashed(true); } catch (e) {}
        try { JSON.parse(dv[j][4] || '[]').forEach(function (x) { if (x.fid) { try { DriveApp.getFileById(x.fid).setTrashed(true); } catch (e2) {} } }); } catch (e) {}
        ds.deleteRow(j + 1);
      }
    } else if (b.action === 'minutes') {
      appendText_('순공시간', [b.date, Number(b.min), Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss')]);
    } else if (b.action === 'minutesUndo') {
      deleteLast_('순공시간', function (r) { return ymd_(r[0]) === b.date; });
    } else if (b.action === 'rcheck') {
      var s = sheet_('반복체크'), v = s.getDataRange().getValues();
      for (var i = v.length - 1; i >= 1; i--) if (ymd_(v[i][0]) === b.date && String(v[i][1]) === b.name) s.deleteRow(i + 1);
      if (b.on) appendText_('반복체크', [b.date, b.name]);
    } else if (b.action === 'quiz') {
      (b.rows || []).forEach(function (r) { appendText_('퀴즈기록', [b.date, r.kind, r.key, r.ok ? 1 : 0]); });
    }
  } finally { lock.releaseLock(); }
  return { ok: true };
}
function appendText_(name, row) {
  var s = sheet_(name); s.appendRow(row);
  s.getRange(s.getLastRow(), 1).setNumberFormat('@').setValue(String(row[0]));
}
function deleteLast_(name, pred) {
  var s = sheet_(name), v = s.getDataRange().getValues();
  for (var i = v.length - 1; i >= 1; i--) if (pred(v[i])) { s.deleteRow(i + 1); return; }
}

function getLog_(date) {
  var v = rows_('기록');
  for (var i = 0; i < v.length; i++) if (ymd_(v[i][0]) === date) {
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


/* ================= 밤 전사파일 → 공부 기록 자동 반영 =================
 * 1) 설정 탭에 전사폴더ID, API키(구글 Gemini 키)를 넣어요
 * 2) setupTrigger 를 한 번 실행하면 매일 밤 11시에 자동으로 돌아요
 * 3) 전사 txt(또는 구글 문서)를 '전사파일' 폴더에 날짜 이름으로 올리면 돼요. 예) 2026-10-13.txt
 * 지금 바로 시험해 보려면 processTranscripts 를 직접 실행하세요 */
function setupTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'processTranscripts') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('processTranscripts').timeBased().everyDays(1).atHour(23).inTimezone(TZ).create();
  Logger.log('완료! 매일 밤 11시에 전사파일을 확인해요.');
}

function processTranscripts() {
  var st = settings_();
  if (!st['전사폴더ID'] || !st['API키']) { Logger.log('설정 탭에 전사폴더ID, API키를 먼저 넣어 주세요.'); return; }
  var folder = DriveApp.getFolderById(st['전사폴더ID']);
  var doneIt = folder.getFoldersByName('처리완료'), done = doneIt.hasNext() ? doneIt.next() : folder.createFolder('처리완료');
  var it = folder.getFiles(), n = 0;
  while (it.hasNext()) {
    var f = it.next();
    try {
      var text = transcriptText_(f);
      if (!text) { Logger.log('읽을 수 없는 파일(txt 또는 구글 문서로 올려 주세요): ' + f.getName()); continue; }
      var m = f.getName().match(/(\d{4})[-._]?(\d{2})[-._]?(\d{2})/);
      var date = m ? m[1] + '-' + m[2] + '-' + m[3] : Utilities.formatDate(f.getDateCreated(), TZ, 'yyyy-MM-dd');
      var r = summarize_(text, st['API키'], st['AI모델'] || 'gemini-2.5-flash');
      upsertLog(date, { 개념1: r.c1, 개념2: r.c2, 개념3: r.c3, 전사요약: r.summary });
      f.moveTo(done); n++;
      Logger.log('반영: ' + f.getName() + ' → ' + date);
    } catch (e) { Logger.log('실패(다음에 다시 시도): ' + f.getName() + ' / ' + e); }
  }
  Logger.log('처리한 파일 ' + n + '개');
}

function transcriptText_(f) {
  var mt = f.getMimeType();
  if (mt === 'application/vnd.google-apps.document') return DocumentApp.openById(f.getId()).getBody().getText();
  if (/^text\//.test(mt)) return f.getBlob().getDataAsString('UTF-8');
  return '';
}

function summarize_(text, key, model) {
  var prompt = '아래는 중학교 1학년 학생이 공부한 내용을 말로 설명한 녹음의 전사 글입니다. 학생에게 보여 줄 공부 기록을 만들어 주세요.\n' +
    '- c1, c2, c3: 오늘 배운 핵심 개념 세 가지를 각각 한 문장으로. 학생이 읽기 쉬운 ~해요 체. 틀린 내용은 올바르게 고쳐서 쓰세요.\n' +
    '- summary: 학생이 어디서 헷갈렸는지와 잘한 점을 1~2문장으로.\n' +
    '전사 글에 없는 내용을 지어내지 마세요.\n\n[전사]\n' + text.slice(0, 60000);
  var body = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: {
      responseMimeType: 'application/json',
      responseSchema: { type: 'OBJECT', properties: { c1: { type: 'STRING' }, c2: { type: 'STRING' }, c3: { type: 'STRING' }, summary: { type: 'STRING' } }, required: ['c1', 'c2', 'c3', 'summary'] }
    }
  };
  var res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + model + ':generateContent', {
    method: 'post', contentType: 'application/json', headers: { 'x-goog-api-key': key }, payload: JSON.stringify(body), muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) throw new Error('AI 응답 오류 ' + res.getResponseCode() + ' ' + res.getContentText().slice(0, 200));
  var j = JSON.parse(res.getContentText());
  return JSON.parse(j.candidates[0].content.parts[0].text);
}
