/**
 * 다이어트 벌금장 - Google Sheets 백엔드 (Apps Script 웹앱)
 *
 * 설치: 시트 > 확장 프로그램 > Apps Script > 이 코드 붙이기 > 저장
 *       배포 > 새 배포 > 유형: 웹 앱 > 실행 주체: 나 > 액세스: 모든 사용자 > 배포
 *       나온 웹앱 URL을 index.html 의 API_URL 에 넣거나, 페이지 설정에서 붙여넣기.
 *
 * 탭: 기록(날짜,이름,체중,단식,운동,저장시각) / 멤버 / 규칙 / 요약(자동 계산)
 */

var SHEETS = {
  logs:    { name: '기록',  header: ['날짜', '이름', '체중', '단식', '운동', '저장시각'] },
  members: { name: '멤버',  header: ['이름', '단식방식', '운동', '시작체중', '목표체중'] },
  rules:   { name: '규칙',  header: ['항목', '금액(원)', '설명'] },
  summary: { name: '요약',  header: ['이름', '최근날짜', '최근체중', '시작대비', '미기록일', '누적벌금'] }
};
var DEFAULT_MEMBERS = ['멤버1', '멤버2', '멤버3', '멤버4', '멤버5'];
var DEFAULT_RULES = [
  ['miss', 1000, '기록 안 한 날'],
  ['fast', 1000, '단식 실패'],
  ['ex',   1000, '운동 안 함'],
  ['gain', 1000, '전 기록보다 체중 증가']
];

function doGet(e) { return json_(readAll_()); }

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var body = JSON.parse(e.postData.contents || '{}');
    var ss = ss_();
    if (body.action === 'log') upsertLog_(ss, body);
    else if (body.action === 'deleteLog') deleteLog_(ss, body);
    else if (body.action === 'member') updateMember_(ss, body);
    else if (body.action === 'rules') updateRules_(ss, body);
    else throw new Error('unknown action');
    var data = readAll_();
    writeSummary_(ss, data);
    return json_(data);
  } catch (err) {
    return json_({ error: String(err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

// ---------- read ----------
function readAll_() {
  var ss = ss_();
  ensure_(ss);
  var members = rows_(ss, SHEETS.members).map(function (r) {
    return { name: s_(r[0]), fast: s_(r[1]), ex: s_(r[2]), start: n_(r[3]), goal: n_(r[4]) };
  }).filter(function (m) { return m.name; });
  var rules = {};
  rows_(ss, SHEETS.rules).forEach(function (r) { if (s_(r[0])) rules[s_(r[0])] = n_(r[1]) || 0; });
  var logs = rows_(ss, SHEETS.logs).map(function (r) {
    return { date: date_(r[0]), name: s_(r[1]), w: n_(r[2]), fast: b_(r[3]), ex: b_(r[4]) };
  }).filter(function (l) { return l.date && l.name && l.w != null; });
  return { members: members, rules: rules, logs: logs, today: fmt_(new Date()) };
}

// ---------- write ----------
function upsertLog_(ss, b) {
  var sh = sheet_(ss, SHEETS.logs), name = s_(b.name), date = s_(b.date);
  var w = Number(b.weight);
  if (!name || !/^\d{4}-\d{2}-\d{2}$/.test(date) || isNaN(w)) throw new Error('잘못된 기록');
  var row = [date, name, Math.round(w * 10) / 10, b.fast ? 'O' : 'X', b.ex ? 'O' : 'X', fmt_(new Date(), true)];
  var vals = sh.getDataRange().getValues();
  for (var i = 1; i < vals.length; i++) {
    if (date_(vals[i][0]) === date && s_(vals[i][1]) === name) { sh.getRange(i + 1, 1, 1, row.length).setValues([row]); return; }
  }
  sh.appendRow(row);
  // 시작 체중 비어 있으면 첫 기록으로 채움
  var ms = sheet_(ss, SHEETS.members), mv = ms.getDataRange().getValues();
  for (var j = 1; j < mv.length; j++) if (s_(mv[j][0]) === name && mv[j][3] === '') ms.getRange(j + 1, 4).setValue(row[2]);
}
function deleteLog_(ss, b) {
  var sh = sheet_(ss, SHEETS.logs), vals = sh.getDataRange().getValues();
  for (var i = vals.length - 1; i >= 1; i--) if (date_(vals[i][0]) === s_(b.date) && s_(vals[i][1]) === s_(b.name)) sh.deleteRow(i + 1);
}
function updateMember_(ss, b) {
  var sh = sheet_(ss, SHEETS.members), idx = Number(b.index);
  if (!(idx >= 0 && idx < 5)) throw new Error('잘못된 멤버');
  var vals = sh.getDataRange().getValues(), oldName = vals[idx + 1] ? s_(vals[idx + 1][0]) : '';
  var newName = s_(b.name) || DEFAULT_MEMBERS[idx];
  sh.getRange(idx + 2, 1, 1, 5).setValues([[newName, s_(b.fast), s_(b.ex), b.start == null || b.start === '' ? '' : Number(b.start), b.goal == null || b.goal === '' ? '' : Number(b.goal)]]);
  if (oldName && oldName !== newName) { // 이름 바꾸면 기록도 따라감
    var ls = sheet_(ss, SHEETS.logs), lv = ls.getDataRange().getValues();
    for (var i = 1; i < lv.length; i++) if (s_(lv[i][1]) === oldName) ls.getRange(i + 1, 2).setValue(newName);
  }
}
function updateRules_(ss, b) {
  var sh = sheet_(ss, SHEETS.rules), vals = sh.getDataRange().getValues();
  for (var i = 1; i < vals.length; i++) {
    var k = s_(vals[i][0]);
    if (b[k] != null && !isNaN(Number(b[k]))) sh.getRange(i + 1, 2).setValue(Math.max(0, Number(b[k])));
  }
}

// ---------- fine engine (index.html 과 동일한 규칙) ----------
function compute_(logs, rules, today) {
  var byDate = {}; logs.forEach(function (l) { byDate[l.date] = l; });
  var dates = Object.keys(byDate).sort(); var rows = [];
  if (!dates.length) return rows;
  var end = dates[dates.length - 1] > today ? dates[dates.length - 1] : today;
  var prevW = null, d = dates[0];
  while (d <= end) {
    var log = byDate[d];
    if (log) {
      var fee = 0, delta = prevW != null ? Math.round((log.w - prevW) * 10) / 10 : null;
      if (delta != null && delta > 0) fee += rules.gain || 0;
      if (!log.fast) fee += rules.fast || 0;
      if (!log.ex) fee += rules.ex || 0;
      rows.push({ date: d, missed: false, w: log.w, fee: fee });
      prevW = log.w;
    } else if (d < today) {
      rows.push({ date: d, missed: true, fee: rules.miss || 0 });
    }
    d = addDays_(d, 1);
  }
  return rows;
}
function writeSummary_(ss, data) {
  var sh = sheet_(ss, SHEETS.summary);
  var out = data.members.map(function (m) {
    var rows = compute_(data.logs.filter(function (l) { return l.name === m.name; }), data.rules, data.today);
    var last = null; for (var i = rows.length - 1; i >= 0; i--) if (!rows[i].missed) { last = rows[i]; break; }
    var missed = rows.filter(function (r) { return r.missed; }).length;
    var total = rows.reduce(function (a, r) { return a + r.fee; }, 0);
    return [m.name, last ? last.date : '', last ? last.w : '', (last && m.start != null) ? Math.round((last.w - m.start) * 10) / 10 : '', missed, total];
  });
  sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 6).clearContent();
  if (out.length) sh.getRange(2, 1, out.length, 6).setValues(out);
}

// ---------- setup / helpers ----------
function ensure_(ss) {
  var created = false;
  Object.keys(SHEETS).forEach(function (k) {
    var def = SHEETS[k], sh = ss.getSheetByName(def.name);
    if (!sh) { sh = ss.insertSheet(def.name); created = true; }
    if (sh.getLastRow() === 0) {
      sh.appendRow(def.header); sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, def.header.length).setFontWeight('bold');
      if (k === 'logs') sh.getRange('A:A').setNumberFormat('@');
      if (k === 'members') DEFAULT_MEMBERS.forEach(function (n) { sh.appendRow([n, '16:8', '걷기 30분', '', '']); });
      if (k === 'rules') DEFAULT_RULES.forEach(function (r) { sh.appendRow(r); });
      if (k === 'summary') sh.getRange('B:B').setNumberFormat('@');
    }
  });
  var first = ss.getSheets()[0];
  if (created && first.getName() === 'Sheet1' && first.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(first);
}
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function sheet_(ss, def) { ensure_(ss); return ss.getSheetByName(def.name); }
function rows_(ss, def) { var sh = sheet_(ss, def), n = sh.getLastRow(); return n < 2 ? [] : sh.getRange(2, 1, n - 1, def.header.length).getValues(); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function s_(v) { return v == null ? '' : String(v).trim(); }
function n_(v) { if (v === '' || v == null) return null; var n = Number(v); return isNaN(n) ? null : n; }
function b_(v) { v = s_(v).toUpperCase(); return v === 'O' || v === 'TRUE' || v === '1' || v === 'Y'; }
function pad_(n) { return (n < 10 ? '0' : '') + n; }
function fmt_(d, withTime) {
  var s = d.getFullYear() + '-' + pad_(d.getMonth() + 1) + '-' + pad_(d.getDate());
  return withTime ? s + ' ' + pad_(d.getHours()) + ':' + pad_(d.getMinutes()) : s;
}
function date_(v) { if (v instanceof Date) return fmt_(v); var s = s_(v); var m = s.match(/^(\d{4})[-./](\d{1,2})[-./](\d{1,2})/); return m ? m[1] + '-' + pad_(+m[2]) + '-' + pad_(+m[3]) : ''; }
function addDays_(s, n) { var p = s.split('-'); var d = new Date(+p[0], +p[1] - 1, +p[2]); d.setDate(d.getDate() + n); return fmt_(d); }
