/**
 * 오프라인 제휴 현황 → GitHub data.json 동기화
 *
 * 시트 구조 (자동 감지):
 *   Row 2     : 에어리소닉 | 소닉플로우 | 슈퍼플로우  (제품 헤더)
 *   Row 4     : 대분류 | 지점수 | 제휴 수량(개)       (KPI 서브헤더)
 *   Row 5-11  : 업종별 집계 행
 *   Row 12    : 총 계
 *   Row 29    : 분류 | 업체명 | 출고 수량             (개별 거래처 헤더)
 *   Row 30+   : 개별 거래처 행
 *
 * ── 사전 설정 ─────────────────────────────────────────────────────
 *   Apps Script 편집기 → ⚙️ 프로젝트 설정 → 스크립트 속성
 *   GITHUB_TOKEN : github_pat_... 또는 ghp_...
 *
 * ── Code.gs onOpen() 에 아래 한 줄 추가 ──────────────────────────
 *   syncAddMenu_();
 */

var SYNC_SHEET     = '오프라인 제휴 현황';
var GH_OWNER       = 'hulkyaya';
var GH_REPO        = 'brandon';
var GH_BRANCH      = 'main';
var GH_FILE        = 'data.json';
var BLOCK_PRODUCTS = ['에어리소닉', '소닉플로우', '슈퍼플로우'];

/* ── 메뉴 ─────────────────────────────────────────────────────── */
function syncAddMenu_() {
  SpreadsheetApp.getUi()
    .createMenu('📤 data.json 동기화')
    .addItem('▶ 지금 동기화',               'syncToDataJson')
    .addSeparator()
    .addItem('⏰ 1시간마다 자동 동기화 켜기', 'installAutoSync')
    .addItem('⏹ 자동 동기화 끄기',           'removeAutoSync')
    .addSeparator()
    .addItem('🔍 GITHUB_TOKEN 확인',         'checkToken')
    .addItem('🔎 시트 헤더 확인',            'checkSheetHeaders')
    .addToUi();
}

/* ── 시트 헤더 진단 ───────────────────────────────────────────── */
function checkSheetHeaders() {
  var sh = SpreadsheetApp.getActive().getSheetByName(SYNC_SHEET);
  if (!sh) { SpreadsheetApp.getUi().alert('"' + SYNC_SHEET + '" 시트가 없습니다.'); return; }
  var values = sh.getDataRange().getValues();
  var msg = '📋 "' + SYNC_SHEET + '" 시트 처음 5행:\n\n';
  for (var r = 0; r < Math.min(5, values.length); r++) {
    var cells = values[r].slice(0, 12).map(function(c) { return '"' + String(c || '').trim() + '"'; });
    msg += 'Row ' + (r + 1) + ': ' + cells.join(', ') + '\n';
  }
  SpreadsheetApp.getUi().alert(msg);
}

/* ── 메인 동기화 ──────────────────────────────────────────────── */
function syncToDataJson() {
  var ui    = SpreadsheetApp.getUi();
  var token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  if (!token) {
    ui.alert('❌ GITHUB_TOKEN 미설정\n\nApps Script 편집기 → ⚙️ 프로젝트 설정 → 스크립트 속성에 GITHUB_TOKEN을 추가하세요.');
    return;
  }

  var sh = SpreadsheetApp.getActive().getSheetByName(SYNC_SHEET);
  if (!sh) { ui.alert('❌ "' + SYNC_SHEET + '" 시트가 없습니다.'); return; }

  var values = sh.getDataRange().getValues();

  /* 1. KPI 요약 블록 파싱 (총 계 행) */
  var kpi = parseBlockKpi_(values);
  if (!kpi) {
    ui.alert('❌ KPI 블록 파싱 실패\n\n에어리소닉/소닉플로우/슈퍼플로우 헤더와 "총 계" 행이 있는지 확인하세요.');
    return;
  }

  /* 2. 개별 거래처 파싱 */
  var entries = parseIndividualPartners_(values);
  Logger.log('개별 거래처: ' + entries.length + '개');

  /* 3. 기존 data.json 에서 좌표 보존 */
  var existing = loadExistingData_(token);
  var coordMap = {};
  (existing.partners || []).forEach(function(p) {
    if (p.name && p.lat && p.lng) coordMap[p.name + '|' + (p.product || '')] = { lat: p.lat, lng: p.lng };
  });

  /* 4. 파트너 배열 빌드 */
  var seen = {}, partners = [];
  entries.forEach(function(e) {
    var key = e.name + '|' + e.product;
    if (!e.name || seen[key]) return;
    seen[key] = true;
    var p = { name: e.name, product: e.product, category: e.category, qty: e.qty };
    var coord = coordMap[key];
    if (coord) { p.lat = coord.lat; p.lng = coord.lng; }
    partners.push(p);
  });

  /* 5. GitHub 업로드 */
  var json = { partners: partners, kpi: kpi, updatedAt: Date.now(), source: 'apps-script-sync' };
  try {
    uploadToGitHub_(token, json, partners.length);
  } catch (e) {
    ui.alert('❌ GitHub 업로드 실패\n\n' + e.message);
    return;
  }

  var withCoord = partners.filter(function(p) { return p.lat && p.lng; }).length;
  SpreadsheetApp.getActive().toast(
    'KPI: ' + kpi.partners + '개소 / ' + kpi.qty + '대\n' +
    '거래처: ' + partners.length + '개 (좌표 ' + withCoord + '개)\n' +
    '1~3분 내 대시보드에 반영됩니다.',
    '✅ data.json 동기화 완료', 10
  );
  Logger.log('완료 — KPI: ' + kpi.partners + '/' + kpi.qty + ', 거래처: ' + partners.length + '개');
}

/* ── KPI 요약 블록 파싱 ───────────────────────────────────────── */
/**
 * 제품명 헤더(에어리소닉 등) → 대분류/지점수/제휴수량 서브헤더 → 총 계 행 구조를 파싱
 */
function parseBlockKpi_(values) {
  var toInt = function(v) {
    if (!v && v !== 0) return 0;
    var n = Number(String(v).replace(/[,\s]/g, ''));
    return isFinite(n) ? Math.round(n) : 0;
  };

  var totalPartners = 0, totalQty = 0;
  var byCategory = {}, byProduct = {};

  for (var pi = 0; pi < BLOCK_PRODUCTS.length; pi++) {
    var prod = BLOCK_PRODUCTS[pi];
    var hRow = -1, hCol = -1;

    // 제품 헤더 셀 찾기 (첫 10행)
    outer:
    for (var r = 0; r < Math.min(10, values.length); r++) {
      for (var c = 0; c < (values[r] || []).length; c++) {
        if (String(values[r][c] || '').trim() === prod) { hRow = r; hCol = c; break outer; }
      }
    }
    if (hRow < 0) continue;

    // 서브헤더 행에서 지점수/제휴수량/대분류 열 찾기
    var colP = -1, colQ = -1, catCol = -1, subR = -1;
    for (var rr = hRow; rr < Math.min(hRow + 6, values.length); rr++) {
      for (var cc = Math.max(0, hCol - 2); cc < Math.min((values[rr] || []).length, hCol + 4); cc++) {
        var v = String(values[rr][cc] || '').replace(/\s/g, '');
        if (v === '지점수')             { colP = cc; subR = rr; }
        if (v.indexOf('제휴수량') === 0) { colQ = cc; subR = rr; }
        if (v === '대분류')             { catCol = cc; }
      }
      if (colP >= 0 && colQ >= 0) break;
    }
    if (colP < 0 || colQ < 0) continue;
    if (catCol < 0) catCol = hCol - 1;

    // 카테고리 행 → 총 계까지 읽기
    var prodByCat = {}, totals = null;
    for (var rrr = subR + 1; rrr < Math.min(subR + 30, values.length); rrr++) {
      var rawLabel = String((values[rrr] || [])[catCol] || '').trim();
      var label    = rawLabel.replace(/\s/g, '');
      if (!label) continue;
      if (label === '총계' || label === '합계') {
        totals = { partners: toInt((values[rrr] || [])[colP]), qty: toInt((values[rrr] || [])[colQ]) };
        break;
      }
      if (BLOCK_PRODUCTS.indexOf(rawLabel) >= 0 && rawLabel !== prod) break;
      var p = toInt((values[rrr] || [])[colP]);
      var q = toInt((values[rrr] || [])[colQ]);
      if (p === 0 && q === 0) continue;
      prodByCat[rawLabel] = { partners: p, qty: q };
    }
    if (!totals) continue;

    byProduct[prod] = { partners: totals.partners, qty: totals.qty, byCategory: prodByCat };
    totalPartners += totals.partners;
    totalQty      += totals.qty;
    Object.keys(prodByCat).forEach(function(cat) {
      if (!byCategory[cat]) byCategory[cat] = { partners: 0, qty: 0 };
      byCategory[cat].partners += prodByCat[cat].partners;
      byCategory[cat].qty      += prodByCat[cat].qty;
    });
  }

  if (totalPartners === 0) return null;
  return { partners: totalPartners, qty: totalQty, byCategory: byCategory, byProduct: byProduct };
}

/* ── 개별 거래처 파싱 ─────────────────────────────────────────── */
/**
 * 분류/업체명/출고수량 헤더 행 이하의 개별 거래처 행을 파싱
 * 분류 셀이 빈 경우 위 행의 분류를 이어받음 (fill-down)
 */
function parseIndividualPartners_(values) {
  // 제품 헤더 행 찾기
  var prodRow = -1;
  outer2:
  for (var r = 0; r < Math.min(10, values.length); r++) {
    for (var c = 0; c < (values[r] || []).length; c++) {
      if (BLOCK_PRODUCTS.indexOf(String(values[r][c] || '').trim()) >= 0) { prodRow = r; break outer2; }
    }
  }
  if (prodRow < 0) return [];

  // 각 제품의 열 위치: 제품명 열이 nameCol, catCol = nameCol-1, qtyCol = nameCol+1
  var blocks = [];
  for (var c2 = 0; c2 < (values[prodRow] || []).length; c2++) {
    var prodName = String(values[prodRow][c2] || '').trim();
    if (BLOCK_PRODUCTS.indexOf(prodName) >= 0) {
      blocks.push({ product: prodName, catCol: c2 - 1, nameCol: c2, qtyCol: c2 + 1 });
    }
  }
  if (!blocks.length) return [];

  // 개별 거래처 헤더 행 찾기 (업체명/사업장명/거래처명 셀)
  var dataHeaderRow = -1;
  for (var r2 = prodRow + 1; r2 < values.length; r2++) {
    for (var ci = 0; ci < (values[r2] || []).length; ci++) {
      var val = String(values[r2][ci] || '').replace(/\s/g, '');
      if (val === '업체명' || val === '사업장명' || val === '거래처명') { dataHeaderRow = r2; break; }
    }
    if (dataHeaderRow >= 0) break;
  }
  if (dataHeaderRow < 0) return [];

  // 데이터 행 파싱 (블록별 fill-down category)
  var result = [];
  var lastCat = {};
  blocks.forEach(function(b) { lastCat[b.product] = ''; });

  for (var r3 = dataHeaderRow + 1; r3 < values.length; r3++) {
    var row = values[r3];
    blocks.forEach(function(b) {
      var catVal = String(row[b.catCol] || '').trim();
      var name   = String(row[b.nameCol] || '').trim();
      var rawQty = String(row[b.qtyCol] || '').replace(/[,\s]/g, '');
      var qty    = isFinite(Number(rawQty)) ? Math.round(Number(rawQty)) : 0;

      if (catVal) lastCat[b.product] = catVal;
      if (!name) return;
      var normName = name.replace(/\s/g, '');
      if (normName === '총계' || normName === '합계' || normName === '소계') return;

      result.push({
        name:     name,
        product:  b.product,
        category: lastCat[b.product] || '미분류',
        qty:      qty
      });
    });
  }
  return result;
}

/* ── 기존 data.json 로드 ──────────────────────────────────────── */
function loadExistingData_(token) {
  var url = 'https://raw.githubusercontent.com/' +
            GH_OWNER + '/' + GH_REPO + '/' + GH_BRANCH + '/' + GH_FILE + '?_=' + Date.now();
  var res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) return {};
  try { return JSON.parse(res.getContentText()); } catch (e) { return {}; }
}

/* ── GitHub Contents API PUT ──────────────────────────────────── */
function uploadToGitHub_(token, json, partnerCount) {
  var baseUrl = 'https://api.github.com/repos/' + GH_OWNER + '/' + GH_REPO + '/contents/' + GH_FILE;
  var headers = { Authorization: 'Bearer ' + token, Accept: 'application/vnd.github+json' };

  var getRes = UrlFetchApp.fetch(baseUrl + '?ref=' + GH_BRANCH, {
    method: 'get', headers: headers, muteHttpExceptions: true
  });
  var sha = (getRes.getResponseCode() === 200) ? JSON.parse(getRes.getContentText()).sha : null;

  var body = {
    message: 'auto: partners sync ' + partnerCount + ' (' + new Date().toISOString() + ')',
    content: Utilities.base64Encode(JSON.stringify(json, null, 1), Utilities.Charset.UTF_8),
    branch:  GH_BRANCH
  };
  if (sha) body.sha = sha;

  var putRes = UrlFetchApp.fetch(baseUrl, {
    method: 'put', contentType: 'application/json',
    headers: headers, payload: JSON.stringify(body), muteHttpExceptions: true
  });
  var code = putRes.getResponseCode();
  if (code < 200 || code >= 300) throw new Error('HTTP ' + code + ': ' + putRes.getContentText().substring(0, 300));
}

/* ── 자동 동기화 트리거 ───────────────────────────────────────── */
function installAutoSync() {
  removeAutoSync();
  ScriptApp.newTrigger('syncToDataJson').timeBased().everyHours(1).create();
  try { SpreadsheetApp.getUi().alert('⏰ 자동 동기화 ON\n\n1시간마다 시트 → data.json → 대시보드 자동 반영됩니다.'); } catch (e) {}
}

function removeAutoSync() {
  var removed = 0;
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'syncToDataJson') { ScriptApp.deleteTrigger(t); removed++; }
  });
  if (removed > 0) {
    try { SpreadsheetApp.getUi().alert('⏹ 자동 동기화 OFF (' + removed + '개 트리거 제거)'); } catch (e) {}
  }
}

/* ── 토큰 확인 ───────────────────────────────────────────────── */
function checkToken() {
  var token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  if (!token) { SpreadsheetApp.getUi().alert('❌ GITHUB_TOKEN 미설정'); return; }
  var res = UrlFetchApp.fetch('https://api.github.com/repos/' + GH_OWNER + '/' + GH_REPO, {
    method: 'get',
    headers: { Authorization: 'Bearer ' + token, Accept: 'application/vnd.github+json' },
    muteHttpExceptions: true
  });
  if (res.getResponseCode() === 200) {
    var info = JSON.parse(res.getContentText());
    SpreadsheetApp.getUi().alert('✅ 토큰 정상\n\nRepo: ' + info.full_name + '\nBranch: ' + GH_BRANCH + '\nFile: ' + GH_FILE);
  } else {
    SpreadsheetApp.getUi().alert('❌ 인증 실패 (HTTP ' + res.getResponseCode() + ')\n' + res.getContentText().substring(0, 300));
  }
}
