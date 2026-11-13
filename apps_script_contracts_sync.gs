/**
 * 제휴처 계약 갱신일자 → GitHub contracts.json 동기화
 *
 * 구글시트 컬럼 (1행 헤더, 순서 무관 — 헤더 이름으로 매핑):
 *   거래처명 | 업종 | 제품 | 계약 시작일 | 계약 만료일 | 담당자 | 비고
 *
 * 동작:
 *   1. 지정 시트에서 계약 데이터 읽기
 *   2. GitHub repo의 contracts.json에 PUT (Contents API)
 *   3. GitHub Pages 1~3분 내 자동 재배포 → 대시보드 "계약 관리" 섹션 반영
 *
 * 사전 조건:
 *   - 스크립트 속성 GITHUB_TOKEN 등록
 *   - CONTRACT_SHEET_NAME 에 실제 시트 이름 입력
 *
 * 사용:
 *   메뉴 [📋 계약 관리] → [▶ GitHub로 동기화]
 */

var CON_GH_OWNER  = 'hulkyaya';
var CON_GH_REPO   = 'brandon';
var CON_GH_BRANCH = 'main';
var CON_GH_PATH   = 'contracts.json';

// ★ 실제 시트 이름으로 변경하세요
var CONTRACT_SHEET_NAME = '계약관리';

// 헤더 이름 → 내부 필드 매핑 (한/영 모두 허용)
var CON_HEADER_MAP = {
  '거래처명':   'name',
  '업체명':     'name',
  '업종':       'industry',
  '제품':       'product',
  '제품명':     'product',
  '계약시작일': 'startDate',
  '계약 시작일':'startDate',
  '시작일':     'startDate',
  '계약만료일': 'endDate',
  '계약 만료일':'endDate',
  '만료일':     'endDate',
  '종료일':     'endDate',
  '담당자':     'manager',
  '비고':       'note',
  '메모':       'note'
};

function conAddMenu_() {
  SpreadsheetApp.getUi()
    .createMenu('📋 계약 관리')
    .addItem('▶ GitHub로 동기화', 'syncContractsToGitHub')
    .addItem('🔍 시트 헤더 미리보기', 'previewContractHeaders')
    .addToUi();
}

function previewContractHeaders() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(CONTRACT_SHEET_NAME);
  if (!sh) {
    SpreadsheetApp.getUi().alert('시트를 찾을 수 없습니다: ' + CONTRACT_SHEET_NAME +
      '\n\nCONTRACT_SHEET_NAME 변수를 실제 시트 이름으로 수정하세요.');
    return;
  }
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  SpreadsheetApp.getUi().alert(
    '시트 헤더 목록:\n\n' + headers.filter(Boolean).join(', ') +
    '\n\n매핑 안 된 컬럼은 무시됩니다.'
  );
}

function conParseDate_(val) {
  if (!val) return null;
  if (val instanceof Date) {
    var y = val.getFullYear(), m = val.getMonth() + 1, d = val.getDate();
    return y + '-' + String(m).padStart(2, '0') + '-' + String(d).padStart(2, '0');
  }
  var s = String(val).trim();
  // yyyy-mm-dd or yyyy/mm/dd
  var m1 = s.match(/^(\d{4})[-\/\.](\d{1,2})[-\/\.](\d{1,2})$/);
  if (m1) return m1[1] + '-' + m1[2].padStart(2, '0') + '-' + m1[3].padStart(2, '0');
  return s || null;
}

function conPutGitHub_(token, contracts) {
  var payload = {
    contracts: contracts,
    updatedAt: Date.now(),
    source: 'sheets-sync'
  };
  var contentB64 = Utilities.base64Encode(
    JSON.stringify(payload, null, 1), Utilities.Charset.UTF_8
  );

  var getUrl = 'https://api.github.com/repos/' + CON_GH_OWNER + '/' + CON_GH_REPO +
               '/contents/' + CON_GH_PATH + '?ref=' + CON_GH_BRANCH;
  var getRes = UrlFetchApp.fetch(getUrl, {
    method: 'get',
    headers: { 'Authorization': 'Bearer ' + token, 'Accept': 'application/vnd.github+json' },
    muteHttpExceptions: true
  });
  var sha = null;
  if (getRes.getResponseCode() === 200) {
    sha = JSON.parse(getRes.getContentText()).sha;
  } else if (getRes.getResponseCode() !== 404) {
    throw new Error('GitHub 파일 조회 실패 (' + getRes.getResponseCode() + ')');
  }

  var putUrl = 'https://api.github.com/repos/' + CON_GH_OWNER + '/' + CON_GH_REPO +
               '/contents/' + CON_GH_PATH;
  var body = {
    message: 'auto: contracts sync ' + contracts.length + ' (' + new Date().toISOString() + ')',
    content: contentB64,
    branch: CON_GH_BRANCH
  };
  if (sha) body.sha = sha;

  var putRes = UrlFetchApp.fetch(putUrl, {
    method: 'put',
    contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + token, 'Accept': 'application/vnd.github+json' },
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });
  var code = putRes.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error('GitHub PUT 실패 (' + code + '): ' + putRes.getContentText().substring(0, 300));
  }
}

function syncContractsToGitHub() {
  var ui = SpreadsheetApp.getUi();
  var token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  if (!token) {
    ui.alert('GITHUB_TOKEN 스크립트 속성이 없습니다.\n\nApps Script 편집기 → ⚙️ 프로젝트 설정 → 스크립트 속성에서 GITHUB_TOKEN 추가 후 다시 시도하세요.');
    return;
  }

  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(CONTRACT_SHEET_NAME);
  if (!sh) {
    ui.alert('시트를 찾을 수 없습니다: ' + CONTRACT_SHEET_NAME +
      '\n\nCONTRACT_SHEET_NAME 변수를 실제 시트 이름으로 수정하세요.');
    return;
  }

  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) {
    ui.alert('시트에 데이터가 없습니다.\n1행: 헤더, 2행~: 데이터');
    return;
  }

  var rawData = sh.getRange(1, 1, lastRow, lastCol).getValues();
  var headers = rawData[0];

  // 헤더 → 필드 인덱스 매핑
  var colMap = {};
  headers.forEach(function(h, i) {
    var key = CON_HEADER_MAP[String(h).trim()];
    if (key && colMap[key] === undefined) colMap[key] = i;
  });

  var contracts = [];
  for (var r = 1; r < rawData.length; r++) {
    var row = rawData[r];
    var name = colMap['name'] !== undefined ? String(row[colMap['name']] || '').trim() : '';
    if (!name) continue;

    var endDate = colMap['endDate'] !== undefined ? conParseDate_(row[colMap['endDate']]) : null;

    contracts.push({
      id:         r,
      name:       name,
      industry:   colMap['industry']  !== undefined ? String(row[colMap['industry']]  || '').trim() : '',
      product:    colMap['product']   !== undefined ? String(row[colMap['product']]   || '').trim() : '',
      startDate:  colMap['startDate'] !== undefined ? conParseDate_(row[colMap['startDate']]) : null,
      endDate:    endDate,
      manager:    colMap['manager']   !== undefined ? String(row[colMap['manager']]   || '').trim() : '',
      note:       colMap['note']      !== undefined ? String(row[colMap['note']]      || '').trim() : ''
    });
  }

  if (!contracts.length) {
    ui.alert('거래처명이 있는 행이 없습니다. 시트 데이터를 확인하세요.');
    return;
  }

  try {
    ss.toast('GitHub에 업로드 중 (' + contracts.length + '건)...', '⏳', 5);
    conPutGitHub_(token, contracts);
    ui.alert('✅ 계약 데이터 → GitHub 동기화 완료\n\n' +
             '총 ' + contracts.length + '건 업로드\n\n' +
             '1~3분 후 대시보드 새로고침하면 반영됩니다.');
  } catch (e) {
    ui.alert('❌ 동기화 실패\n\n' + e.message);
    Logger.log(e.stack || e);
  }
}
