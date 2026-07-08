/**
 * 교회 재정실 헌금관리 시스템 - Apps Script 백엔드
 * -----------------------------------------------------
 * 이 스크립트는 두 가지 방식으로 쓸 수 있습니다.
 *  A) 스프레드시트에 바인딩 (확장 프로그램 > Apps Script) → SPREADSHEET_ID를 비워두면 됩니다.
 *  B) 독립 실행형 스크립트 (script.google.com에서 새 프로젝트) → 아래 SPREADSHEET_ID에
 *     스프레드시트 URL의 /d/ 와 /edit 사이 부분(ID)을 붙여넣으세요.
 *
 * [처음 설치 시 1회 실행]
 * 1) 이 코드를 Apps Script 편집기에 붙여넣기 (필요하면 SPREADSHEET_ID 입력)
 * 2) 상단 함수 선택 드롭다운에서 setupSheet 선택 후 실행(▶) → 권한 승인
 * 3) 배포 > 새 배포 > 웹 앱으로 배포
 *    - 실행 계정: 나(Me)
 *    - 액세스 권한: 전체(Anyone)
 * 4) 배포 후 나오는 웹 앱 URL을 index.html 설정 화면에 입력
 *
 * [통신 방식]
 * GitHub Pages 등 외부 도메인에서 fetch()로 호출하면 브라우저 CORS 정책에
 * 막히는 경우가 있어, 모든 요청(조회/저장/삭제 포함)을 GET + JSONP 방식으로
 * 처리합니다. entries/offeringTypes 같은 배열 값은 JSON 문자열로 전달됩니다.
 */

const SPREADSHEET_ID = ''; // 독립 실행형으로 쓸 경우 여기에 스프레드시트 ID를 입력하세요
const SHEET_NAME = '헌금기록';
const SETTINGS_SHEET_NAME = '설정';
const DEFAULT_OFFERING_TYPES = ['십일조', '감사헌금', '주일헌금', '선교헌금', '건축헌금'];

function getSpreadsheet_() {
  if (SPREADSHEET_ID) return SpreadsheetApp.openById(SPREADSHEET_ID);
  return SpreadsheetApp.getActiveSpreadsheet();
}

// ---------- 초기 설정 ----------

function setupSheet() {
  const ss = getSpreadsheet_();

  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['ID', '날짜', '년도', '월', '주차(월중)', '헌금종류', '금액', '비고', '입력시각']);
    sheet.setFrozenRows(1);
  }

  let settingsSheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet(SETTINGS_SHEET_NAME);
    settingsSheet.appendRow(['항목', '값']);
    settingsSheet.appendRow(['헌금종류', DEFAULT_OFFERING_TYPES.join(',')]);
    settingsSheet.appendRow(['보고서수신이메일', '']);
  }

  SpreadsheetApp.flush();
  return 'setup-complete';
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('헌금관리')
    .addItem('초기 설정 실행', 'setupSheet')
    .addItem('비밀번호 재설정', 'resetPasswordPrompt')
    .addToUi();
}

function resetPasswordPrompt() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('새 비밀번호를 입력하세요 (재정실 공용 비밀번호)');
  if (res.getSelectedButton() === ui.Button.OK) {
    const pw = res.getResponseText().trim();
    if (pw) {
      PropertiesService.getScriptProperties().setProperty('ADMIN_PASSWORD', pw);
      ui.alert('비밀번호가 변경되었습니다.');
    }
  }
}

// ---------- 공통 유틸 ----------

function getSheet_() {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('시트가 초기화되지 않았습니다. setupSheet를 먼저 실행하세요.');
  return sheet;
}

function getSettingsMap_() {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
  const map = {};
  if (!sheet) return map;
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    map[values[i][0]] = values[i][1];
  }
  return map;
}

function setSetting_(key, value) {
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SETTINGS_SHEET_NAME);
    sheet.appendRow(['항목', '값']);
  }
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sheet.appendRow([key, value]);
}

function getOfferingTypes_() {
  const map = getSettingsMap_();
  if (map['헌금종류']) {
    return String(map['헌금종류']).split(',').map(s => s.trim()).filter(Boolean);
  }
  return DEFAULT_OFFERING_TYPES;
}

function checkPassword_(password) {
  const props = PropertiesService.getScriptProperties();
  const stored = props.getProperty('ADMIN_PASSWORD');
  if (!stored) {
    if (!password) return false;
    props.setProperty('ADMIN_PASSWORD', password);
    return true;
  }
  return stored === password;
}

function toDateKey_(dateInput) {
  const d = (dateInput instanceof Date) ? dateInput : new Date(dateInput + 'T00:00:00');
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function weekOfMonth_(d) {
  return Math.ceil(d.getDate() / 7);
}

// ---------- 데이터 읽기 헬퍼 ----------

function readAllRecords_() {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
  return values.map((row, idx) => ({
    rowIndex: idx + 2,
    id: row[0],
    date: row[1] instanceof Date ? Utilities.formatDate(row[1], Session.getScriptTimeZone(), 'yyyy-MM-dd') : row[1],
    year: row[2],
    month: row[3],
    weekOfMonth: row[4],
    type: row[5],
    amount: Number(row[6]) || 0,
    note: row[7],
    enteredAt: row[8]
  })).filter(r => r.date);
}

function sumByType_(records, types) {
  const result = {};
  types.forEach(t => result[t] = 0);
  let total = 0;
  records.forEach(r => {
    if (result[r.type] === undefined) result[r.type] = 0;
    result[r.type] += r.amount;
    total += r.amount;
  });
  return { byType: result, total: total };
}

// ---------- 액션 구현 ----------

function actionAddRecord_(payload) {
  const sheet = getSheet_();
  const dateKey = toDateKey_(payload.date);
  const d = new Date(dateKey + 'T00:00:00');
  const year = d.getFullYear();
  const month = d.getMonth() + 1;
  const wom = weekOfMonth_(d);
  const entries = payload.entries || [];
  const now = new Date();
  let added = 0;

  entries.forEach(entry => {
    const amount = Number(entry.amount);
    if (!entry.type || !amount || amount <= 0) return;
    const id = Utilities.getUuid();
    sheet.appendRow([id, dateKey, year, month, wom, entry.type, amount, payload.note || '', now]);
    added++;
  });

  return { added: added, date: dateKey };
}

function actionDeleteRecord_(payload) {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { deleted: false };
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] === payload.id) {
      sheet.deleteRow(i + 2);
      return { deleted: true };
    }
  }
  return { deleted: false };
}

function actionGetWeek_(payload) {
  const dateKey = toDateKey_(payload.date);
  const all = readAllRecords_();
  const records = all.filter(r => r.date === dateKey);
  const types = getOfferingTypes_();
  const summary = sumByType_(records, types);
  return { date: dateKey, records: records, summary: summary, offeringTypes: types };
}

function actionGetMonth_(payload) {
  const year = Number(payload.year);
  const month = Number(payload.month);
  const all = readAllRecords_();
  const records = all.filter(r => r.year === year && r.month === month);
  const types = getOfferingTypes_();

  const byDate = {};
  records.forEach(r => {
    if (!byDate[r.date]) byDate[r.date] = [];
    byDate[r.date].push(r);
  });
  const weeks = Object.keys(byDate).sort().map(date => {
    const s = sumByType_(byDate[date], types);
    return { date: date, byType: s.byType, total: s.total };
  });

  const summary = sumByType_(records, types);
  return { year: year, month: month, weeks: weeks, summary: summary, offeringTypes: types };
}

function actionGetYear_(payload) {
  const year = Number(payload.year);
  const all = readAllRecords_();
  const records = all.filter(r => r.year === year);
  const types = getOfferingTypes_();

  const byMonth = {};
  for (let m = 1; m <= 12; m++) byMonth[m] = [];
  records.forEach(r => { byMonth[r.month].push(r); });

  const months = [];
  for (let m = 1; m <= 12; m++) {
    const s = sumByType_(byMonth[m], types);
    months.push({ month: m, byType: s.byType, total: s.total });
  }

  const summary = sumByType_(records, types);
  return { year: year, months: months, summary: summary, offeringTypes: types };
}

function actionGetMonthsWithData_(payload) {
  const year = Number(payload.year);
  const all = readAllRecords_();
  const set = {};
  all.filter(r => r.year === year).forEach(r => { set[r.month] = true; });
  return { months: Object.keys(set).map(Number).sort((a,b)=>a-b) };
}

function actionGetSettings_() {
  const map = getSettingsMap_();
  return {
    offeringTypes: getOfferingTypes_(),
    reportEmail: map['보고서수신이메일'] || ''
  };
}

function actionUpdateSettings_(payload) {
  if (payload.offeringTypes) {
    setSetting_('헌금종류', payload.offeringTypes.join(','));
  }
  if (payload.reportEmail !== undefined) {
    setSetting_('보고서수신이메일', payload.reportEmail);
  }
  return { updated: true };
}

function buildReportText_(kind, payload) {
  const types = getOfferingTypes_();
  const fmt = n => Number(n || 0).toLocaleString('ko-KR');
  let lines = [];

  if (kind === 'week') {
    const data = actionGetWeek_(payload);
    lines.push('[주별 헌금 보고서] ' + data.date);
    types.forEach(t => lines.push(t + ': ' + fmt(data.summary.byType[t]) + '원'));
    lines.push('합계: ' + fmt(data.summary.total) + '원');
  } else if (kind === 'month') {
    const data = actionGetMonth_(payload);
    lines.push('[월별 헌금 보고서] ' + data.year + '년 ' + data.month + '월');
    data.weeks.forEach(w => {
      lines.push('- ' + w.date + ' 합계: ' + fmt(w.total) + '원');
    });
    lines.push('');
    lines.push('[월 합계]');
    types.forEach(t => lines.push(t + ': ' + fmt(data.summary.byType[t]) + '원'));
    lines.push('총 합계: ' + fmt(data.summary.total) + '원');
  } else if (kind === 'year') {
    const data = actionGetYear_(payload);
    lines.push('[연간 헌금 보고서] ' + data.year + '년');
    data.months.forEach(m => {
      if (m.total > 0) lines.push('- ' + m.month + '월 합계: ' + fmt(m.total) + '원');
    });
    lines.push('');
    lines.push('[연간 합계]');
    types.forEach(t => lines.push(t + ': ' + fmt(data.summary.byType[t]) + '원'));
    lines.push('총 합계: ' + fmt(data.summary.total) + '원');
  }
  return lines.join('\n');
}

function actionEmailReport_(payload) {
  const map = getSettingsMap_();
  const to = payload.email || map['보고서수신이메일'];
  if (!to) throw new Error('수신 이메일이 설정되어 있지 않습니다.');
  const text = buildReportText_(payload.kind, payload);
  const subjectMap = { week: '주별 헌금 보고서', month: '월별 헌금 보고서', year: '연간 헌금 보고서' };
  MailApp.sendEmail(to, subjectMap[payload.kind] || '헌금 보고서', text);
  return { sent: true, to: to };
}

// ---------- 라우팅 ----------

function routeAction_(action, payload) {
  switch (action) {
    case 'ping': return { ok: true };
    case 'addRecord': return actionAddRecord_(payload);
    case 'deleteRecord': return actionDeleteRecord_(payload);
    case 'getWeek': return actionGetWeek_(payload);
    case 'getMonth': return actionGetMonth_(payload);
    case 'getYear': return actionGetYear_(payload);
    case 'getMonthsWithData': return actionGetMonthsWithData_(payload);
    case 'getSettings': return actionGetSettings_();
    case 'updateSettings': return actionUpdateSettings_(payload);
    case 'emailReport': return actionEmailReport_(payload);
    default: throw new Error('알 수 없는 action: ' + action);
  }
}

// ---------- 응답 처리 (JSONP 지원) ----------

function normalizeParams_(raw) {
  const params = {};
  Object.keys(raw).forEach(k => { params[k] = raw[k]; });
  ['entries', 'offeringTypes'].forEach(key => {
    if (typeof params[key] === 'string' && params[key]) {
      try { params[key] = JSON.parse(params[key]); } catch (e) { /* 그대로 둠 */ }
    }
  });
  return params;
}

function respond_(rawParams, obj) {
  const json = JSON.stringify(obj);
  const callback = rawParams && rawParams.callback;
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  var raw = (e && e.parameter) || {};
  try {
    setupIfNeeded_();
    var params = normalizeParams_(raw);
    var action = params.action;

    if (action !== 'ping' && !checkPassword_(params.password)) {
      return respond_(raw, { error: 'unauthorized' });
    }
    var result = routeAction_(action, params);
    return respond_(raw, { result: result });
  } catch (err) {
    return respond_(raw, { error: String(err) });
  }
}

function doPost(e) {
  var raw = (e && e.parameter) || {};
  try {
    setupIfNeeded_();
    var body = {};
    if (e.postData && e.postData.contents) {
      body = JSON.parse(e.postData.contents);
    }
    var action = body.action;
    if (!checkPassword_(body.password)) {
      return respond_(raw, { error: 'unauthorized' });
    }
    var result = routeAction_(action, body);
    return respond_(raw, { result: result });
  } catch (err) {
    return respond_(raw, { error: String(err) });
  }
}

function setupIfNeeded_() {
  var ss = getSpreadsheet_();
  if (!ss.getSheetByName(SHEET_NAME)) {
    setupSheet();
  }
}
