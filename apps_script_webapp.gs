/**
 * Apps Script Web App — 영업프로젝트 + 일정 CRUD
 *
 * 배포 방법 (1회):
 *   1. Apps Script 편집기 열기
 *   2. 배포 → 새 배포 → 유형: 웹 앱
 *   3. 설명: 대시보드 API
 *      다음 사용자로 실행: 나
 *      액세스 권한: 모든 사용자
 *   4. 배포 → 웹 앱 URL 복사
 *   5. index.html의 WEBAPP_URL 상수에 붙여넣기
 *
 * 재배포 (코드 변경 시):
 *   배포 → 배포 관리 → 편집(연필 아이콘) → 버전: 새 버전 → 배포
 */

var SPREADSHEET_ID = '1F5C-XAGHAUlyPWDwEcfje__ynzeGNeQ3CWUHNW80ES4';
var SALES_SHEET    = '영업프로젝트';
var EVENTS_SHEET   = '일정';

var SALES_HEADERS  = ['_id','name','stage','owner','amount','due','note','createdAt','updatedAt'];
var EVENTS_HEADERS = ['_id','title','date','time','location','note','createdAt','updatedAt'];

function doPost(e) {
  try {
    var params = JSON.parse(e.postData.contents);
    var action = params.action;
    var data   = params.data || {};
    var ss     = SpreadsheetApp.openById(SPREADSHEET_ID);
    var result = {};

    if      (action === 'saleAdd')      result = saleAdd_(ss, data);
    else if (action === 'saleUpdate')   result = saleUpdate_(ss, data);
    else if (action === 'saleDelete')   result = saleDelete_(ss, data);
    else if (action === 'eventAdd')     result = eventAdd_(ss, data);
    else if (action === 'eventUpdate')  result = eventUpdate_(ss, data);
    else if (action === 'eventDelete')  result = eventDelete_(ss, data);
    else if (action === 'contractRenew') result = contractRenew_(ss, data);
    else throw new Error('알 수 없는 action: ' + action);

    return jsonOk_(result);
  } catch (err) {
    return jsonErr_(err.message);
  }
}

/* ======== SALES ======== */
function saleAdd_(ss, d) {
  var sh  = getOrCreate_(ss, SALES_SHEET, SALES_HEADERS);
  var id  = Utilities.getUuid();
  var now = Date.now();
  sh.appendRow([id, d.name, d.stage||'제안', d.owner||'', d.amount||0, d.due||'', d.note||'', now, now]);
  return { id: id };
}

function saleUpdate_(ss, d) {
  var sh   = getOrCreate_(ss, SALES_SHEET, SALES_HEADERS);
  var rows = sh.getDataRange().getValues();
  for (var r = 1; r < rows.length; r++) {
    if (String(rows[r][0]) === String(d._id)) {
      sh.getRange(r + 1, 2, 1, 8).setValues([[
        d.name, d.stage||'제안', d.owner||'', d.amount||0,
        d.due||'', d.note||'', rows[r][7], Date.now()
      ]]);
      return {};
    }
  }
  throw new Error('항목을 찾을 수 없음: ' + d._id);
}

function saleDelete_(ss, d) {
  deleteById_(getOrCreate_(ss, SALES_SHEET, SALES_HEADERS), d._id);
  return {};
}

/* ======== EVENTS ======== */
function eventAdd_(ss, d) {
  var sh  = getOrCreate_(ss, EVENTS_SHEET, EVENTS_HEADERS);
  var id  = Utilities.getUuid();
  var now = Date.now();
  sh.appendRow([id, d.title, d.date, d.time||'', d.location||'', d.note||'', now, now]);
  return { id: id };
}

function eventUpdate_(ss, d) {
  var sh   = getOrCreate_(ss, EVENTS_SHEET, EVENTS_HEADERS);
  var rows = sh.getDataRange().getValues();
  for (var r = 1; r < rows.length; r++) {
    if (String(rows[r][0]) === String(d._id)) {
      sh.getRange(r + 1, 2, 1, 7).setValues([[
        d.title, d.date, d.time||'', d.location||'', d.note||'', rows[r][6], Date.now()
      ]]);
      return {};
    }
  }
  throw new Error('항목을 찾을 수 없음: ' + d._id);
}

function eventDelete_(ss, d) {
  deleteById_(getOrCreate_(ss, EVENTS_SHEET, EVENTS_HEADERS), d._id);
  return {};
}

/* ======== HELPERS ======== */
function getOrCreate_(ss, name, headers) {
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function deleteById_(sh, id) {
  var rows = sh.getDataRange().getValues();
  for (var r = rows.length - 1; r >= 1; r--) {
    if (String(rows[r][0]) === String(id)) {
      sh.deleteRow(r + 1);
      return;
    }
  }
  throw new Error('항목을 찾을 수 없음: ' + id);
}

function jsonOk_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(Object.assign({ ok: true }, obj)))
    .setMimeType(ContentService.MimeType.JSON);
}

function jsonErr_(msg) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: false, error: msg }))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ======== CONTRACT RENEWAL ======== */
function renewalDate_(value) {
  var s = String(value || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) throw new Error('날짜를 YYYY-MM-DD 형식으로 입력하세요.');
  var dt = new Date(s + 'T12:00:00+09:00');
  if (isNaN(dt.getTime()) || Utilities.formatDate(dt, 'Asia/Seoul', 'yyyy-MM-dd') !== s)
    throw new Error('존재하지 않는 날짜입니다.');
  return dt;
}
function contractRenew_(ss, d) {
  var start = renewalDate_(d.startDate), end = renewalDate_(d.endDate);
  if (end < start) throw new Error('새 만료일은 갱신 시작일 이후여야 합니다.');
  var rowNum = Number(d.row);
  if (!Number.isInteger(rowNum) || rowNum < 3) throw new Error('계약 행이 올바르지 않습니다.');
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sh = ss.getSheetByName('계약 현황');
    if (!sh || rowNum > sh.getLastRow()) throw new Error('계약을 찾을 수 없습니다. 새로고침해 주세요.');
    var row = sh.getRange(rowNum, 1, 1, 10).getValues()[0];
    var forms = sh.getRange(rowNum, 1, 1, 10).getFormulas()[0];
    var text = function(v) { return String(v || '').trim(); };
    var dateText = function(v) { return v instanceof Date ? Utilities.formatDate(v, ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd') : text(v); };
    if (text(row[2]) !== text(d.name) || text(row[6]) !== text(d.product) ||
        dateText(row[4]) !== text(d.expectedStart) || dateText(row[5]) !== text(d.expectedEnd))
      throw new Error('계약 내용이 변경되었습니다. 새로고침 후 다시 입력해 주세요.');
    var history = '재갱신 ' + d.startDate + ' / 만료 ' + (dateText(row[5]) || '미입력') + ' → ' + d.endDate;
    var note = text(row[9]);
    var today = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
    var status = d.endDate < today ? '만료' : d.startDate > today ? '갱신 예정' : '진행중';
    // D열 최초 계약일, G/H열 제품/수량, 기존 상태 수식과 비고를 보존한다.
    sh.getRange(rowNum, 5, 1, 6).setValues([[start, end, forms[6] || row[6], forms[7] || row[7], forms[8] || status, (note ? note + '\n' : '') + history]]);
    sh.getRange(rowNum, 5, 1, 2).setNumberFormat('yyyy-mm-dd');
    SpreadsheetApp.flush();
    return {row:rowNum, startDate:d.startDate, endDate:d.endDate};
  } finally { lock.releaseLock(); }
}
