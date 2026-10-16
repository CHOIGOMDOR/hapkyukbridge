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
  var ct = ensure_(ss, '개념', ['단원', '종류', '문장']);   // 오늘의 개념 세 문장의 원본(고쳐 써도 돼요)
  if (ct.getLastRow() < 2) ct.getRange(2, 1, CONCEPTS.length, 3).setValues(CONCEPTS);
  ensure_(ss, '분석기록', ANALYSIS_COLS);   // 학습 분석 누적(부모용, 웹페이지에는 안 나가요)
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
   ['전사폴더ID', '137o9lPGy6DLn7x-LRxuEFDIQpNABjypv'], ['API키', ''], ['AI모델', 'gemini-3.8-flash']].forEach(function (r) { if (!have[r[0]]) st.appendRow(r); });
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
function apiPost(bodyJson) {
  try { return handle_(JSON.parse(bodyJson)); }
  catch (e) {
    var m = String(e && e.message || e);
    if (/getFolderById|No item with the given ID|Invalid argument|잘못된 파일 또는 폴더|파일 또는 폴더 ID|찾을 수 없/i.test(m) && /photo|qaAnswer/.test(bodyJson.slice(0, 40))) m = '설정 탭의 사진폴더ID가 비어 있거나 잘못됐어요';
    return { ok: false, error: m };
  }
}
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
  var concepts = rows_('개념').filter(function (r) { return r[0] && r[2]; }).map(function (r) { return [String(r[0]), String(r[1]), String(r[2])]; });
  var images = [];
  if (st['이미지폴더ID']) {
    var it = DriveApp.getFolderById(st['이미지폴더ID']).getFiles();
    while (it.hasNext()) {
      var f = it.next();
      if (/^image\//.test(f.getMimeType())) images.push('https://drive.google.com/thumbnail?id=' + f.getId() + '&sz=w600');
    }
  }
  return {
    plan: plan, log: log, vocab: vocab, repeats: repeats, rchecks: rchecks, minutes: minutes, qlog: qlog, questions: questions, qa: qa, concepts: concepts, images: images,
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
      var folder = DriveApp.getFolderById(String(settings_()['사진폴더ID'] || '').trim());
      var blob = Utilities.newBlob(Utilities.base64Decode(b.data), b.mime, b.date + '_' + b.name);
      var f = folder.createFile(blob);
      f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      var url = 'https://drive.google.com/thumbnail?id=' + f.getId() + '&sz=w1200';
      var id = b.id || ('q' + Date.now());
      appendText_('모르는문제', [id, b.date, url, b.memo || '', '[]', f.getId()]);
      return { ok: true, id: id, url: url };
    } else if (b.action === 'transcribe') {
      return { ok: true, text: transcribeAudio_(b.data, b.mime || 'audio/mp4') };
    } else if (b.action === 'qaAnswer') {
      var qs = sheet_('모르는문제'), qv = qs.getDataRange().getValues(), entry = { who: b.who, text: b.text || '', at: b.at, img: '', fid: '' };
      if (b.data) {
        var af = DriveApp.getFolderById(String(settings_()['사진폴더ID'] || '').trim()).createFile(Utilities.newBlob(Utilities.base64Decode(b.data), b.mime, b.date + '_답변_' + b.name));
        af.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        entry.img = 'https://drive.google.com/thumbnail?id=' + af.getId() + '&sz=w1200'; entry.fid = af.getId();
      }
      for (var i = 1; i < qv.length; i++) if (String(qv[i][0]) === String(b.id)) {
        var arr = []; try { arr = JSON.parse(qv[i][4] || '[]'); } catch (e) {}
        arr.push(entry);
        writeAnswers_(qs, i + 1, arr);
        break;
      }
      return { ok: true, img: entry.img };
    } else if (b.action === 'qaAnswerEdit' || b.action === 'qaAnswerDelete') {
      var es = sheet_('모르는문제'), ev = es.getDataRange().getValues();
      for (var r = 1; r < ev.length; r++) if (String(ev[r][0]) === String(b.id)) {
        var list = []; try { list = JSON.parse(ev[r][4] || '[]'); } catch (e) {}
        var k = -1;
        if (b.aid) for (var m = 0; m < list.length; m++) if (list[m].aid === b.aid) { k = m; break; }
        if (k < 0 && b.idx >= 0 && b.idx < list.length) k = b.idx;
        if (k >= 0) {
          if (b.action === 'qaAnswerEdit') { list[k].text = b.text; list[k].edited = true; }
          else { var gone = list.splice(k, 1)[0]; if (gone && gone.fid) { try { DriveApp.getFileById(gone.fid).setTrashed(true); } catch (e) {} } }
          writeAnswers_(es, r + 1, list);
        }
        break;
      }
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
  var folder = DriveApp.getFolderById(String(st['전사폴더ID']).trim());
  var doneIt = folder.getFoldersByName('처리완료'), done = doneIt.hasNext() ? doneIt.next() : folder.createFolder('처리완료');
  var it = folder.getFiles(), n = 0;
  while (it.hasNext()) {
    var f = it.next();
    try {
      var text = transcriptText_(f);
      if (!text) { Logger.log('읽을 수 없는 파일(txt 또는 구글 문서로 올려 주세요): ' + f.getName()); continue; }
      var m = f.getName().match(/(\d{4})[-._]?(\d{2})[-._]?(\d{2})/);
      var date = m ? m[1] + '-' + m[2] + '-' + m[3] : Utilities.formatDate(f.getDateCreated(), TZ, 'yyyy-MM-dd');
      var r = summarize_(text, st['API키'], st['AI모델'] || 'gemini-3.8-flash');
      upsertLog(date, { 개념1: r.c1, 개념2: r.c2, 개념3: r.c3, 전사요약: r.summary });
      var okA = false;
      try { okA = analyzeTranscript_(text, date, f.getName(), st); } catch (e2) { Logger.log('분석 실패(내일 밤 다시 시도): ' + f.getName() + ' / ' + e2); }
      if (!okA) continue;                       // 분석이 끝나야 처리완료로 옮겨요
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
  var res, tried = [model, 'gemini-3.8-flash', 'gemini-flash-latest'];   // 모델 이름이 바뀌어 404가 나면 다음 이름을 써요
  for (var i = 0; i < tried.length; i++) {
    res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + tried[i] + ':generateContent', {
      method: 'post', contentType: 'application/json', headers: { 'x-goog-api-key': key }, payload: JSON.stringify(body), muteHttpExceptions: true
    });
    if (res.getResponseCode() !== 404) break;
  }
  if (res.getResponseCode() !== 200) throw new Error('AI 응답 오류 ' + res.getResponseCode() + ' ' + res.getContentText().slice(0, 200));
  var j = JSON.parse(res.getContentText());
  return JSON.parse(j.candidates[0].content.parts[0].text);
}


/** 폰에서 녹음한 짧은 소리를 한국어 글로 받아 적어요 (설정 탭 API키 사용) */
function transcribeAudio_(b64, mime) {
  var st = settings_(), key = st['API키'];
  if (!key) throw new Error('설정 탭에 API키를 넣어 주세요');
  var body = { contents: [{ parts: [
    { text: '다음 음성을 한국어로 그대로 받아 적어 주세요. 학생이 모르는 문제나 설명을 말한 내용입니다. 받아 적은 글만 출력하고 다른 말은 쓰지 마세요.' },
    { inline_data: { mime_type: mime, data: b64 } }
  ] }] };
  var res, tried = [st['AI모델'] || 'gemini-3.8-flash', 'gemini-3.8-flash', 'gemini-flash-latest'];
  for (var i = 0; i < tried.length; i++) {
    res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + tried[i] + ':generateContent', {
      method: 'post', contentType: 'application/json', headers: { 'x-goog-api-key': key }, payload: JSON.stringify(body), muteHttpExceptions: true
    });
    if (res.getResponseCode() !== 404) break;
  }
  if (res.getResponseCode() !== 200) throw new Error('음성 변환 오류 ' + res.getResponseCode() + ' ' + res.getContentText().slice(0, 160));
  var j = JSON.parse(res.getContentText());
  return String(((j.candidates || [{}])[0].content || { parts: [{}] }).parts[0].text || '').trim();
}


/** 답변 목록을 시트에 저장(JSON + 읽기 쉬운 글) */
function writeAnswers_(sheet, row, list) {
  sheet.getRange(row, 5).setValue(JSON.stringify(list));
  sheet.getRange(row, 7).setValue(list.map(function (x) {
    return '[' + x.who + ' ' + (x.at || '') + (x.edited ? ' 수정됨' : '') + '] ' + (x.text || '') + (x.img ? ' (사진: ' + x.img + ')' : '');
  }).join('\n'));
}

/* ================= 학습 분석 (6영역 + 독립 메타인지) =================
 * 기준서: 6영역(개념 20·사고 20·적용·확장 25·오류 수정 15·기억·인출 10·집중·조절 10) + 별도 메타인지 5항목(각 20점)
 * AI는 근거와 점수 판단만 하고, 가중 종합·메타인지 총점·근거 충실도·변화량은 아래 코드가 직접 계산해요.
 * 결과는 시트 '분석기록'(누적, 덮어쓰지 않고 추가)과 '분석현황'(최신 요약)에만 쌓여요. 웹페이지에는 내보내지 않아요. */
var AREAS = [
  { key: 'concept', name: '개념 이해', w: 0.20 }, { key: 'thinking', name: '사고 과정', w: 0.20 },
  { key: 'apply', name: '적용·확장', w: 0.25 }, { key: 'error', name: '오류 수정', w: 0.15 },
  { key: 'memory', name: '기억·인출', w: 0.10 }, { key: 'focus', name: '집중·조절', w: 0.10 }
];
var METAS = [
  { key: 'm1', name: '자기상태 인식' }, { key: 'm2', name: '확신도·정확성 일치' }, { key: 'm3', name: '계획·전략 선택' },
  { key: 'm4', name: '수행 중 점검·오류 발견' }, { key: 'm5', name: '수정·검증·전이' }
];
var ANALYSIS_COLS = ['분석ID', '분석일', '자료일', '회차', '과목·주제', '원본파일',
  '개념 이해', '사고 과정', '적용·확장', '오류 수정', '기억·인출', '집중·조절', '가중 종합',
  '자기상태 인식', '확신도·정확성', '계획·전략', '수행 중 점검', '수정·검증·전이', '메타인지 총점', '근거 충실도',
  '난도', '도움', '최고 영역', '핵심 병목', '대표 근거1', '대표 근거2',
  '보충', '확인', '통과 기준', '메타인지 질문',
  '직전 종합', '종합 변화', '직전 메타', '메타 변화', '메타 최근3회 평균', '상태·해석 제한'];

function analysisPrompt_(text) {
  return '당신은 학습 분석가입니다. 아래는 중학교 1학년 학생이 혼자 공부하며 클로바 노트로 녹음한 전사 글입니다(교사나 부모의 질문은 없고, 학생 혼자 말합니다). ' +
    '아래 기준으로 학생의 이 회차 학습을 분석해 JSON으로만 답하세요.\n\n' +
    '[기본 6영역, 0~100점. 관찰 근거가 없으면 점수 -1(미관찰)]\n' +
    '- concept 개념 이해의 정확성: 용어를 자기 말로 설명, 조건과 결론 구분, 사용 조건 확인\n' +
    '- thinking 사고 과정과 막힘 지점: 무엇을 구하는지, 개념을 고른 이유, 풀이 단계 사이의 이유, 막힌 지점을 구체적으로\n' +
    '- apply 사고 확장과 적용: 유사·변형 문제에 원리 적용, 조건이 바뀌면 전략 조정\n' +
    '- error 오류 인식·수정·극복: 틀린 것을 알아차리고 원인을 특정, 고친 뒤 검증\n' +
    '- memory 기억 유지와 인출: 자료 없이 설명하고 꺼내 쓰는지\n' +
    '- focus 집중 상태와 학습조절: 학습 진입, 집중 저하 신호를 스스로 말하고 조절. 침묵이나 소음만으로 집중 저하라고 단정하지 말 것\n\n' +
    '[독립 메타인지 5항목, 각각 0·5·10·15·20 중 하나. 기회가 없어 관찰하지 못했으면 -1(미관찰). 관찰하지 못한 항목을 0점으로 쓰지 말 것]\n' +
    '- m1 자기상태 인식: 아는 것·모르는 것·막힌 지점·필요한 도움을 구분해 말함\n' +
    '- m2 확신도와 정확성 일치: 말한 확신(예: 70%)이 실제 정확성과 맞는지, 불확실할 때 검증하는지\n' +
    '- m3 계획과 전략 선택: 첫 행동과 방법을 고른 이유\n' +
    '- m4 수행 중 점검과 오류 발견: 풀이 도중 조건·의미·계산 점검\n' +
    '- m5 수정과 검증 및 전이: 고친 뒤 다시 확인하고 새 문제에 적용\n' +
    '(확신도 퍼센트는 학생의 자신감이며 메타인지 점수와 다릅니다.)\n\n' +
    '[기타] difficulty: D1(기초 재생) D2(기본 개념) D3(학년 수준 적용) D4(복합 조건) D5(심화·전이) 중 하나. ' +
    'help: 혼자 학습이므로 대체로 H0이며, 교재·영상·정답을 보고 풀었다는 말이 있으면 H1~H3, 판단 불가면 H?.\n' +
    'quotes: 학생이 실제로 한 말을 전사 글에서 한 글자도 바꾸지 말고 그대로 복사한 대표 근거 2개(없으면 빈 배열).\n' +
    'next: 다음에 바로 할 일. repair(보충: 다시 지도할 개념·단계), check(확인: 다음에 쓸 문항·질문), pass(통과 기준: 문항 수·정확도·난도·허용 도움), metaQuestion(학생이 말해야 할 막힘·전략·검산 질문 한 개).\n' +
    'bottleneck: 가장 큰 병목 한 줄, best: 가장 좋았던 점 한 줄. subject: 과목과 주제 한 줄. limits: 해석상 한계(예: 전사가 짧아 관찰 부족)를 한 줄.\n' +
    '전사에 없는 내용은 지어내지 말고, 점수에는 근거가 있어야 합니다. 평가는 학생을 비난하지 않는 교육용 지표입니다.\n\n[전사]\n' + text.slice(0, 60000);
}

function analysisSchema_() {
  var item = { type: 'OBJECT', properties: { score: { type: 'INTEGER' }, evidence: { type: 'STRING' } }, required: ['score', 'evidence'] };
  var areas = { type: 'OBJECT', properties: {}, required: [] }, metas = { type: 'OBJECT', properties: {}, required: [] };
  AREAS.forEach(function (a) { areas.properties[a.key] = item; areas.required.push(a.key); });
  METAS.forEach(function (m) { metas.properties[m.key] = item; metas.required.push(m.key); });
  return { type: 'OBJECT', properties: {
    subject: { type: 'STRING' }, difficulty: { type: 'STRING' }, help: { type: 'STRING' },
    areas: areas, metas: metas, quotes: { type: 'ARRAY', items: { type: 'STRING' } },
    bottleneck: { type: 'STRING' }, best: { type: 'STRING' }, limits: { type: 'STRING' },
    next: { type: 'OBJECT', properties: { repair: { type: 'STRING' }, check: { type: 'STRING' }, pass: { type: 'STRING' }, metaQuestion: { type: 'STRING' } }, required: ['repair', 'check', 'pass', 'metaQuestion'] }
  }, required: ['subject', 'difficulty', 'help', 'areas', 'metas', 'quotes', 'bottleneck', 'best', 'limits', 'next'] };
}

/** 공통 Gemini 호출(JSON 응답). 모델 이름이 바뀌어 404가 나면 다음 이름을 써요 */
function geminiJson_(prompt, schema, key, model) {
  var body = { contents: [{ parts: [{ text: prompt }] }], generationConfig: { responseMimeType: 'application/json', responseSchema: schema } };
  var res, tried = [model, 'gemini-3.8-flash', 'gemini-flash-latest'];
  for (var i = 0; i < tried.length; i++) {
    res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + tried[i] + ':generateContent', {
      method: 'post', contentType: 'application/json', headers: { 'x-goog-api-key': key }, payload: JSON.stringify(body), muteHttpExceptions: true
    });
    if (res.getResponseCode() !== 404) break;
  }
  if (res.getResponseCode() !== 200) throw new Error('AI 응답 오류 ' + res.getResponseCode() + ' ' + res.getContentText().slice(0, 200));
  return JSON.parse(JSON.parse(res.getContentText()).candidates[0].content.parts[0].text);
}

function snap5_(v) { v = Number(v); if (isNaN(v) || v < 0) return null; return Math.min(20, Math.round(v / 5) * 5); }
function area100_(v) { v = Number(v); if (isNaN(v) || v < 0) return null; return Math.min(100, Math.round(v)); }
function norm_(s) { return String(s || '').replace(/\s+/g, ' ').trim(); }

/** AI가 준 판단(raw)을 기준서 공식대로 계산. prev: 직전까지의 분석기록 객체 배열(오래된 순) */
function scoreAnalysis_(raw, transcript, prev) {
  var notes = [], area = {}, wsum = 0, wtot = 0, observedAreas = 0;
  AREAS.forEach(function (a) {
    var v = area100_((raw.areas[a.key] || {}).score);
    area[a.key] = v;
    if (v !== null) { wsum += v * a.w; wtot += a.w; observedAreas++; }
  });
  var composite = observedAreas >= 3 ? Math.round(wsum / wtot * 10) / 10 : null;   // 3개 미만이면 종합을 내지 않아요(관찰 부족)
  if (observedAreas < 6 && observedAreas >= 3) notes.push('6영역 중 ' + observedAreas + '개만 관찰되어 관찰된 영역의 가중치로 환산했어요');
  if (observedAreas < 3) notes.push('6영역 관찰 부족(' + observedAreas + '/6) — 가중 종합은 계산하지 않았어요');
  var meta = {}, msum = 0, mobs = 0;
  METAS.forEach(function (m) {
    var v = snap5_((raw.metas[m.key] || {}).score);
    meta[m.key] = v;
    if (v !== null) { msum += v; mobs++; }
  });
  var metaTotal = mobs >= 3 ? Math.round(msum / (mobs * 20) * 1000) / 10 : null;
  var fidelity = Math.round(mobs / 5 * 100);
  if (mobs < 3) notes.push('메타인지 관찰 부족(' + mobs + '/5)');
  // 인용은 전사에 실제로 있는 문장만 인정
  var tn = norm_(transcript), quotes = (raw.quotes || []).filter(function (q) { return norm_(q).length > 3 && tn.indexOf(norm_(q)) >= 0; }).slice(0, 2);
  if ((raw.quotes || []).length && quotes.length < Math.min(2, raw.quotes.length)) notes.push('전사에서 확인되지 않은 인용은 뺐어요');
  if (/침묵|소음/.test(String(raw.limits || ''))) notes.push('집중은 침묵만으로 단정하지 않았어요');
  if (raw.limits) notes.push(String(raw.limits));
  // 최고 영역(관찰된 것 중)
  var best = null, bestV = -1;
  AREAS.forEach(function (a) { if (area[a.key] !== null && area[a.key] > bestV) { bestV = area[a.key]; best = a.name; } });
  // 직전 대비
  var lastC = null, lastM = null;
  for (var i = prev.length - 1; i >= 0; i--) { if (lastC === null && typeof prev[i].composite === 'number') lastC = prev[i].composite; if (lastM === null && typeof prev[i].metaTotal === 'number') lastM = prev[i].metaTotal; }
  var recent = prev.map(function (x) { return x.metaTotal; }).filter(function (x) { return typeof x === 'number'; });
  if (metaTotal !== null) recent.push(metaTotal);
  recent = recent.slice(-3);
  var avg3 = recent.length ? Math.round(recent.reduce(function (a, b) { return a + b; }, 0) / recent.length * 10) / 10 : null;
  return {
    area: area, composite: composite, meta: meta, metaTotal: metaTotal, fidelity: fidelity, quotes: quotes, best: best,
    lastC: lastC, dC: (composite !== null && lastC !== null) ? Math.round((composite - lastC) * 10) / 10 : null,
    lastM: lastM, dM: (metaTotal !== null && lastM !== null) ? Math.round((metaTotal - lastM) * 10) / 10 : null,
    avg3: avg3, notes: notes, baseline: lastC === null
  };
}

function analysisRows_() {
  var s = sheet_('분석기록'); if (!s) return [];
  var v = s.getDataRange().getValues().slice(1);
  return v.filter(function (r) { return r[0]; }).map(function (r) {
    return { id: r[0], date: ymd_(r[2]), file: String(r[5] || ''), composite: r[12] === '' ? null : Number(r[12]), metaTotal: (r[18] === '' || isNaN(Number(r[18]))) ? null : Number(r[18]), row: r };
  });
}

/** 전사 한 개를 분석해 '분석기록'에 한 줄 추가. 같은 원본 파일은 다시 집계하지 않아요 */
function analyzeTranscript_(text, date, fileName, st) {   // 성공하면 true
  var ss = ss_(), tab = ensure_(ss, '분석기록', ANALYSIS_COLS);
  var prev = analysisRows_();
  if (prev.some(function (x) { return x.file === fileName && x.date === date; })) { Logger.log('이미 분석한 자료라 건너뛰어요: ' + fileName); return true; }
  var raw = geminiJson_(analysisPrompt_(text), analysisSchema_(), st['API키'], st['AI모델'] || 'gemini-3.8-flash');
  var r = scoreAnalysis_(raw, text, prev);
  var n = prev.length + 1, nx = raw.next || {};
  var ev = function (o, k) { var e = ((o || {})[k] || {}).evidence; return e ? String(e) : ''; };
  var row = [
    date + '#' + n, Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'), date, n, raw.subject || '', fileName,
    r.area.concept, r.area.thinking, r.area.apply, r.area.error, r.area.memory, r.area.focus, r.composite,
    r.meta.m1, r.meta.m2, r.meta.m3, r.meta.m4, r.meta.m5, r.metaTotal === null ? '관찰 부족' : r.metaTotal, r.fidelity + '%',
    raw.difficulty || '', raw.help || '', r.best || '', raw.bottleneck || '', r.quotes[0] || '', r.quotes[1] || '',
    nx.repair || '', nx.check || '', nx.pass || '', nx.metaQuestion || '',
    r.lastC === null ? '기초선' : r.lastC, r.dC === null ? '' : r.dC, r.lastM === null ? '' : r.lastM, r.dM === null ? '' : r.dM, r.avg3 === null ? '' : r.avg3,
    r.notes.join(' / ')
  ];
  var sh = tab; sh.appendRow(row);
  var rr = sh.getLastRow();
  sh.getRange(rr, 1, 1, 3).setNumberFormat('@').setValues([[row[0], row[1], row[2]]]);
  // 각 영역의 근거 문장은 메모로 붙여 두어요(셀에 마우스를 올리면 보여요)
  AREAS.forEach(function (a, i) { var e = ev(raw.areas, a.key); if (e) sh.getRange(rr, 7 + i).setNote(e); });
  METAS.forEach(function (m, i) { var e = ev(raw.metas, m.key); if (e) sh.getRange(rr, 14 + i).setNote(e); });
  updateAnalysisSummary_();
  return true;
}

/** '분석현황': 최신 회차 요약과 최근 추이(부모용). 실행할 때마다 새로 그려요 */
function updateAnalysisSummary_() {
  var rows = analysisRows_(); if (!rows.length) return;
  var ss = ss_(), sh = ss.getSheetByName('분석현황') || ss.insertSheet('분석현황');
  sh.clear();
  var last = rows[rows.length - 1].row, H = ANALYSIS_COLS;
  var g = function (name) { return last[H.indexOf(name)]; };
  var out = [
    ['학습 분석 현황 (부모용 · 교육용 지표이며 학교 시험점수와 다릅니다)', '', '', ''],
    ['최신 회차', g('회차') + '회차', '자료일', g('자료일')],
    ['과목·주제', g('과목·주제'), '난도 / 도움', g('난도') + ' / ' + g('도움')],
    ['', '', '', ''],
    ['영역', '점수', '비중', '']
  ];
  AREAS.forEach(function (a) { var v = g(a.name === '개념 이해' ? '개념 이해' : a.name); out.push([a.name, v === '' ? '미관찰' : v, Math.round(a.w * 100) + '%', '']); });
  out.push(['가중 종합', g('가중 종합'), '직전 ' + g('직전 종합'), '변화 ' + g('종합 변화')]);
  out.push(['', '', '', '']);
  out.push(['독립 메타인지', '점수(20점 만점)', '', '']);
  ['자기상태 인식', '확신도·정확성', '계획·전략', '수행 중 점검', '수정·검증·전이'].forEach(function (n) { var v = g(n); out.push([n, v === '' ? '미관찰' : v, '', '']); });
  out.push(['메타인지 총점', g('메타인지 총점'), '근거 충실도 ' + g('근거 충실도'), '최근 3회 평균 ' + g('메타 최근3회 평균')]);
  out.push(['', '', '', '']);
  out.push(['최고 영역', g('최고 영역'), '', '']);
  out.push(['핵심 병목', g('핵심 병목'), '', '']);
  out.push(['보충', g('보충'), '', '']);
  out.push(['확인', g('확인'), '', '']);
  out.push(['통과 기준', g('통과 기준'), '', '']);
  out.push(['메타인지 질문', g('메타인지 질문'), '', '']);
  out.push(['해석 제한', g('상태·해석 제한'), '', '']);
  out.push(['', '', '', '']);
  out.push(['최근 추이', '자료일', '가중 종합', '메타인지 총점']);
  rows.slice(-8).forEach(function (x) { out.push([x.row[3] + '회차', x.date, x.composite === null ? '' : x.composite, x.metaTotal === null ? '관찰 부족' : x.metaTotal]); });
  sh.getRange(1, 1, out.length, 4).setValues(out);
  sh.getRange(1, 1).setFontWeight('bold').setFontSize(13);
  sh.setColumnWidth(1, 150); sh.setColumnWidth(2, 360); sh.setColumnWidth(3, 170); sh.setColumnWidth(4, 170);
  sh.getRange(1, 1, out.length, 4).setWrap(true).setVerticalAlignment('top');
}

/* ===CONCEPTS-BEGIN=== */
var CONCEPTS = [["소수와 합성수","개념","소수는 1보다 큰 자연수 중에서 약수가 1과 자기 자신뿐인 수예요."],["소수와 합성수","개념","합성수는 약수가 3개 이상인 수예요. 1은 소수도 합성수도 아니에요."],["소수와 합성수","개념","소수 중에서 짝수는 2 하나뿐이에요."],["소수와 합성수","개념","약수는 어떤 수를 나누어떨어지게 하는 수이고, 배수는 그 수를 1배, 2배, 3배… 한 수예요."],["소수와 합성수","연산","12의 약수는 1, 2, 3, 4, 6, 12예요. 짝을 지어 빠짐없이 찾아요."],["소수와 합성수","연산","20 이하의 소수는 2, 3, 5, 7, 11, 13, 17, 19예요."],["소수와 합성수","연산","에라토스테네스의 체: 2부터 시작해 그 수의 배수를 차례로 지우면 소수만 남아요."],["소수와 합성수","연산","어떤 수가 소수인지 볼 때는 작은 소수 2, 3, 5, 7로 나누어떨어지는지 확인해요."],["소인수분해","개념","거듭제곱은 같은 수를 여러 번 곱한 것을 2³처럼 밑과 지수로 나타내는 거예요."],["소인수분해","개념","소인수는 어떤 수의 약수 중에서 소수인 수예요."],["소인수분해","개념","소인수분해는 자연수를 소인수들만의 곱으로 나타내는 거예요."],["소인수분해","개념","소인수분해 결과는 곱하는 순서만 달라질 뿐 하나로 정해져요."],["소인수분해","연산","2×2×2×3 = 2³×3처럼 같은 수의 곱은 지수로 정리해요."],["소인수분해","연산","12는 12÷2=6, 6÷2=3으로 나누어 2²×3이에요."],["소인수분해","연산","72는 2³×3²이에요. 작은 소수부터 차례로 나눠요."],["소인수분해","연산","소인수분해한 식에서 지수에 1을 더해 곱하면 약수의 개수예요. 12=2²×3은 (2+1)×(1+1)=6개예요."],["최대공약수·최소공배수(1)","개념","공약수는 두 수의 공통인 약수이고, 그중 가장 큰 수가 최대공약수예요."],["최대공약수·최소공배수(1)","개념","두 수의 공약수는 최대공약수의 약수예요."],["최대공약수·최소공배수(1)","개념","최대공약수가 1인 두 수를 서로소라고 해요."],["최대공약수·최소공배수(1)","개념","공배수는 두 수의 공통인 배수이고, 그중 가장 작은 수가 최소공배수예요."],["최대공약수·최소공배수(1)","연산","12와 18의 최대공약수는 6이에요."],["최대공약수·최소공배수(1)","연산","소인수분해로 최대공약수를 구할 때는 공통인 소인수에 작은 지수를 곱해요."],["최대공약수·최소공배수(1)","연산","소인수분해로 최소공배수를 구할 때는 모든 소인수에 큰 지수를 곱해요."],["최대공약수·최소공배수(1)","연산","2²×3과 2×3²의 최대공약수는 2×3=6, 최소공배수는 2²×3²=36이에요."],["최소공배수·소인수분해 단원 마무리","개념","두 수 A, B의 최대공약수 G와 최소공배수 L 사이에는 A×B = G×L이 성립해요."],["최소공배수·소인수분해 단원 마무리","개념","최대공약수는 \"똑같이 나누기\", 최소공배수는 \"동시에 다시 만나기\" 문제에 써요."],["최소공배수·소인수분해 단원 마무리","개념","두 수를 최대공약수로 나눈 몫은 서로소예요."],["최소공배수·소인수분해 단원 마무리","개념","소수, 소인수분해, 최대공약수, 최소공배수는 서로 이어져 있어요."],["최소공배수·소인수분해 단원 마무리","연산","6과 8의 최소공배수는 24, 최대공약수는 2이고 6×8=24×2=48이에요."],["최소공배수·소인수분해 단원 마무리","연산","가로 12cm, 세로 18cm 직사각형을 크기가 같은 가장 큰 정사각형으로 자르면 한 변이 6cm예요."],["최소공배수·소인수분해 단원 마무리","연산","4일, 6일마다 오는 두 버스가 처음 같이 만나는 날은 최소공배수인 12일 뒤예요."],["최소공배수·소인수분해 단원 마무리","연산","소인수분해는 작은 소수부터 나누고, 지수를 꼭 확인해요."],["정수와 유리수의 뜻·대소 관계","개념","양수는 0보다 큰 수, 음수는 0보다 작은 수이고 0은 양수도 음수도 아니에요."],["정수와 유리수의 뜻·대소 관계","개념","정수는 양의 정수, 0, 음의 정수를 모두 합한 수예요."],["정수와 유리수의 뜻·대소 관계","개념","유리수는 분수 a/b (a, b는 정수, b≠0)로 나타낼 수 있는 수예요."],["정수와 유리수의 뜻·대소 관계","개념","수직선 위에서 오른쪽에 있을수록 큰 수예요."],["정수와 유리수의 뜻·대소 관계","연산","절댓값은 수직선에서 0까지의 거리예요. |-3|=3, |3|=3이에요."],["정수와 유리수의 뜻·대소 관계","연산","음수끼리는 절댓값이 클수록 작아요. -5 < -2예요."],["정수와 유리수의 뜻·대소 관계","연산","양수는 음수보다 항상 커요."],["정수와 유리수의 뜻·대소 관계","연산","분수는 통분해서, 소수는 분수로 바꿔서 크기를 비교해요."],["정수와 유리수의 덧셈·뺄셈","개념","부호가 같은 두 수를 더할 때는 절댓값을 더하고 공통 부호를 붙여요."],["정수와 유리수의 덧셈·뺄셈","개념","부호가 다른 두 수를 더할 때는 절댓값이 큰 수에서 작은 수를 빼고 절댓값이 큰 수의 부호를 붙여요."],["정수와 유리수의 덧셈·뺄셈","개념","뺄셈은 빼는 수의 부호를 바꿔서 더하는 거예요."],["정수와 유리수의 덧셈·뺄셈","개념","덧셈은 교환법칙과 결합법칙이 성립해요."],["정수와 유리수의 덧셈·뺄셈","연산","(+3)+(+4)=+7, (-3)+(-4)=-7이에요."],["정수와 유리수의 덧셈·뺄셈","연산","(+5)+(-8)=-3이에요. 8이 더 크니까 부호는 음수예요."],["정수와 유리수의 덧셈·뺄셈","연산","(+2)-(+5)=(+2)+(-5)=-3이에요."],["정수와 유리수의 덧셈·뺄셈","연산","(-4)-(-6)=(-4)+(+6)=+2예요."],["정수와 유리수의 곱셈·나눗셈","개념","부호가 같은 두 수의 곱은 양수, 다른 두 수의 곱은 음수예요."],["정수와 유리수의 곱셈·나눗셈","개념","음수를 짝수 번 곱하면 양수, 홀수 번 곱하면 음수예요."],["정수와 유리수의 곱셈·나눗셈","개념","나눗셈은 나누는 수의 역수를 곱하는 것과 같아요."],["정수와 유리수의 곱셈·나눗셈","개념","곱셈은 교환법칙·결합법칙이 성립하고, 분배법칙도 계산에 써요."],["정수와 유리수의 곱셈·나눗셈","연산","(-3)×(-4)=+12, (-3)×(+4)=-12예요."],["정수와 유리수의 곱셈·나눗셈","연산","(-2)³=-8, (-2)²=4예요. 부호와 지수를 함께 살펴요."],["정수와 유리수의 곱셈·나눗셈","연산","12÷(-3)=-4예요. 부호가 다르면 몫은 음수예요."],["정수와 유리수의 곱셈·나눗셈","연산","계산 순서는 괄호 → 거듭제곱 → 곱셈·나눗셈 → 덧셈·뺄셈이에요."],["정수와 유리수 단원 마무리·복습","개념","정수와 유리수는 수직선 위의 점으로 나타내고 크기를 비교해요."],["정수와 유리수 단원 마무리·복습","개념","덧셈·뺄셈은 부호와 절댓값, 곱셈·나눗셈은 부호 규칙이 핵심이에요."],["정수와 유리수 단원 마무리·복습","개념","뺄셈은 덧셈으로, 나눗셈은 곱셈으로 바꾸면 계산이 하나로 이어져요."],["정수와 유리수 단원 마무리·복습","개념","계산 순서는 괄호 → 거듭제곱 → 곱셈·나눗셈 → 덧셈·뺄셈이에요."],["정수와 유리수 단원 마무리·복습","연산","틀리기 쉬운 곳은 부호예요. 계산하기 전에 부호부터 정해요."],["정수와 유리수 단원 마무리·복습","연산","-2²=-4이고 (-2)²=4예요. 괄호가 있는지 꼭 확인해요."],["정수와 유리수 단원 마무리·복습","연산","분수의 나눗셈은 역수를 곱해요. (2/3)÷(4/5)=(2/3)×(5/4)=5/6이에요."],["정수와 유리수 단원 마무리·복습","연산","풀이 끝에 부호와 약분을 한 번 더 점검해요."],["문자의 사용·식의 값","개념","문자를 사용하면 수량 사이의 관계를 식으로 간단하게 나타낼 수 있어요."],["문자의 사용·식의 값","개념","곱셈 기호 ×는 생략하고, 수는 문자 앞에 써요. a×3은 3a예요."],["문자의 사용·식의 값","개념","나눗셈 기호 ÷는 분수 꼴로 나타내요. a÷3은 a/3이에요."],["문자의 사용·식의 값","개념","식의 값은 문자에 수를 대입해서 계산한 값이에요."],["문자의 사용·식의 값","연산","한 개에 500원인 사탕 x개의 값은 500x원이에요."],["문자의 사용·식의 값","연산","a×a는 a²이고, a×b×3은 3ab예요."],["문자의 사용·식의 값","연산","x=2일 때 3x+1의 값은 3×2+1=7이에요."],["문자의 사용·식의 값","연산","음수를 대입할 때는 괄호를 씌워요. x=-2일 때 5x는 5×(-2)=-10이에요."],["일차식의 곱셈·나눗셈, 덧셈·뺄셈","개념","항은 수나 문자의 곱으로만 이루어진 식이고, 계수는 문자 앞의 수예요."],["일차식의 곱셈·나눗셈, 덧셈·뺄셈","개념","문자와 차수가 같은 항을 동류항이라고 해요."],["일차식의 곱셈·나눗셈, 덧셈·뺄셈","개념","일차식은 차수가 1인 다항식이에요."],["일차식의 곱셈·나눗셈, 덧셈·뺄셈","개념","괄호 앞이 음수이면 괄호를 풀 때 안의 부호가 모두 바뀌어요."],["일차식의 곱셈·나눗셈, 덧셈·뺄셈","연산","3x+2x=5x처럼 동류항은 계수끼리 더해요."],["일차식의 곱셈·나눗셈, 덧셈·뺄셈","연산","2(x+3)=2x+6처럼 분배법칙으로 괄호를 풀어요."],["일차식의 곱셈·나눗셈, 덧셈·뺄셈","연산","-(x-4)=-x+4예요."],["일차식의 곱셈·나눗셈, 덧셈·뺄셈","연산","(6x+9)÷3=2x+3이에요. 모든 항을 나눠요."],["방정식과 그 해·일차방정식의 풀이","개념","등식은 등호(=)로 나타낸 식이고, 방정식은 미지수의 값에 따라 참·거짓이 달라지는 등식이에요."],["방정식과 그 해·일차방정식의 풀이","개념","방정식을 참이 되게 하는 미지수의 값을 해(근)라고 해요."],["방정식과 그 해·일차방정식의 풀이","개념","등식의 양변에 같은 수를 더하거나 빼거나, 곱하거나 나누어도(0으로 나누기 제외) 등식은 성립해요."],["방정식과 그 해·일차방정식의 풀이","개념","이항은 항의 부호를 바꿔서 등호의 다른 쪽으로 옮기는 거예요."],["방정식과 그 해·일차방정식의 풀이","연산","x+3=7이면 x=7-3=4예요."],["방정식과 그 해·일차방정식의 풀이","연산","2x=10이면 양변을 2로 나눠 x=5예요."],["방정식과 그 해·일차방정식의 풀이","연산","3x-4=11이면 -4를 이항해 3x=15, x=5예요."],["방정식과 그 해·일차방정식의 풀이","연산","계수가 분수이면 양변에 분모의 최소공배수를 곱해 정수로 만들어요."],["일차방정식의 활용·단원 마무리","개념","문장제는 구하려는 것을 x로 놓고, 문장에서 식을 세운 뒤 풀어요."],["일차방정식의 활용·단원 마무리","개념","풀고 나면 해가 문제의 조건에 맞는지 꼭 확인해요."],["일차방정식의 활용·단원 마무리","개념","거리 = 속력 × 시간이에요."],["일차방정식의 활용·단원 마무리","개념","일차방정식 풀이는 괄호 풀기 → 이항 → 정리 → 계수로 나누기 순서예요."],["일차방정식의 활용·단원 마무리","연산","어떤 수의 2배에서 3을 뺀 수가 9이면 2x-3=9이고 x=6이에요."],["일차방정식의 활용·단원 마무리","연산","연속하는 세 자연수의 합이 33이면 x-1, x, x+1로 놓아 3x=33이고 x=11이에요."],["일차방정식의 활용·단원 마무리","연산","시속 4km로 x시간 걸으면 4x km예요."],["일차방정식의 활용·단원 마무리","연산","소금의 양 = 농도(%)/100 × 소금물의 양으로 식을 세워요."],["순서쌍과 좌표·그래프","개념","좌표평면은 x축과 y축이 원점에서 수직으로 만나는 평면이에요."],["순서쌍과 좌표·그래프","개념","순서쌍 (a, b)에서 a는 x좌표, b는 y좌표예요."],["순서쌍과 좌표·그래프","개념","좌표평면은 x축과 y축에 의해 4개의 사분면으로 나뉘어요."],["순서쌍과 좌표·그래프","개념","그래프는 두 양 사이의 관계를 점과 선으로 나타낸 거예요."],["순서쌍과 좌표·그래프","연산","점 (3, 2)는 원점에서 오른쪽으로 3, 위로 2만큼 간 점이에요."],["순서쌍과 좌표·그래프","연산","제1사분면은 (+, +), 제2사분면은 (-, +), 제3사분면은 (-, -), 제4사분면은 (+, -)예요."],["순서쌍과 좌표·그래프","연산","x축 위의 점은 y좌표가 0이고, y축 위의 점은 x좌표가 0이에요."],["순서쌍과 좌표·그래프","연산","그래프를 읽을 때는 가로축과 세로축이 무엇을 뜻하는지 먼저 확인해요."],["정비례·반비례","개념","y가 x에 정비례하면 y=ax (a≠0)예요."],["정비례·반비례","개념","정비례 그래프는 원점을 지나는 직선이에요."],["정비례·반비례","개념","y가 x에 반비례하면 y=a/x (a≠0)예요."],["정비례·반비례","개념","반비례 그래프는 원점을 지나지 않는 한 쌍의 매끄러운 곡선이에요."],["정비례·반비례","연산","x가 2배, 3배가 될 때 y도 2배, 3배가 되면 정비례예요."],["정비례·반비례","연산","x=2일 때 y=6이면 정비례식은 y=3x예요."],["정비례·반비례","연산","x=2일 때 y=6이면 반비례식은 y=12/x예요."],["정비례·반비례","연산","정비례는 y÷x가 일정하고, 반비례는 x×y가 일정해요."],["좌표평면과 그래프 단원 마무리·전체 복습","개념","1학기 수학은 수 → 식 → 방정식 → 그래프 순서로 이어져요."],["좌표평면과 그래프 단원 마무리·전체 복습","개념","소인수분해, 정수와 유리수, 문자와 식, 일차방정식, 좌표와 그래프가 핵심 단원이에요."],["좌표평면과 그래프 단원 마무리·전체 복습","개념","식으로 나타낸 것을 그래프로 그려 보면 관계가 한눈에 보여요."],["좌표평면과 그래프 단원 마무리·전체 복습","개념","틀린 문제는 어떤 개념을 놓쳤는지 적어 두고 다시 풀어 봐요."],["좌표평면과 그래프 단원 마무리·전체 복습","연산","소인수분해는 작은 소수부터 나눠요."],["좌표평면과 그래프 단원 마무리·전체 복습","연산","정수의 계산은 부호부터 정해요."],["좌표평면과 그래프 단원 마무리·전체 복습","연산","방정식은 이항해서 x 혼자 남기고, 풀고 나면 대입해서 확인해요."],["좌표평면과 그래프 단원 마무리·전체 복습","연산","정비례는 y=ax, 반비례는 y=a/x로 구분해요."]];
/* ===CONCEPTS-END=== */
