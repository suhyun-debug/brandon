/**
 * ERP 재고 → GitHub inventory.json 동기화 (Firebase 의존 없음)
 *
 * 동작:
 *   1. ERP /v2/inventory/stock/paging 호출 (전 페이지 스캔)
 *   2. 제품명에 INV_FAMILY_KEYWORDS 포함된 것만 추출
 *   3. GitHub repo의 inventory.json에 PUT (Contents API)
 *   4. GitHub Pages 1~3분 내 자동 재배포 → 대시보드 "제품 재고 현황" 반영
 *   5. 6분 실행 한도 도달 시 진행 상태 저장 → 메뉴 재클릭으로 이어받기
 *
 * 사전 조건:
 *   - 기존 Apps Script에 login(), authGet() 함수 (ERP 호출용)
 *   - ERP 계정 등록 (메뉴 ⚙️ 내 ERP 계정 등록)
 *   - 스크립트 속성 GITHUB_TOKEN 등록 (partners 동기화와 동일 토큰)
 *
 * 사용:
 *   메뉴 [📦 ERP 재고] → [▶ GitHub로 동기화]
 *   또는 [⏰ 1시간 자동 동기화 켜기] (트리거 기반, 사람이 안 눌러도 자동 실행)
 */

// GitHub 설정 (partners 동기화와 동일)
var INV_GH_OWNER  = 'hulkyaya';
var INV_GH_REPO   = 'brandon';
var INV_GH_BRANCH = 'main';
var INV_GH_PATH   = 'inventory.json';

// 자동 동기화 트리거 설정
var INV_AUTOSYNC_FN    = 'autoSyncInventoryToGitHub';
var INV_AUTOSYNC_HOURS = 1;

// 동기화 대상 제품군 — 이 키워드가 productName에 포함된 것만 업로드
var INV_FAMILY_KEYWORDS = ['에어리소닉', '소닉플로우'];

// 재실행 시 이어받기용 키
var INV_RESUME_KEY = 'INV_SYNC_RESUME_GH';
var INV_BUFFER_KEY = 'INV_SYNC_BUFFER_GH';
var INV_MAX_RUN_MS = 4.5 * 60 * 1000;

// onOpen()에 invAddMenu_() 한 줄 추가
function invAddMenu_() {
  SpreadsheetApp.getUi()
    .createMenu('📦 ERP 재고')
    .addItem('▶ GitHub로 동기화', 'syncInventoryToGitHub')
    .addSeparator()
    .addItem('⏰ ' + INV_AUTOSYNC_HOURS + '시간 자동 동기화 켜기', 'installInventoryAutoSync')
    .addItem('⏹ 자동 동기화 끄기', 'removeInventoryAutoSync')
    .addSeparator()
    .addItem('🔍 ERP 재고 데이터 미리보기 (첫 5건)', 'previewErpInventory')
    .addItem('🗑️ 동기화 진행상태 초기화', 'resetInventoryResume')
    .addToUi();
}

function autoSyncInventoryToGitHub() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    Logger.log('autoSyncInventoryToGitHub: 다른 실행이 진행 중이라 건너뜀');
    return;
  }
  try {
    syncInventoryToGitHub();
  } catch (e) {
    Logger.log('autoSyncInventoryToGitHub 실패: ' + (e && e.message));
  } finally {
    lock.releaseLock();
  }
}

function installInventoryAutoSync() {
  removeInventoryAutoSync();
  ScriptApp.newTrigger(INV_AUTOSYNC_FN)
    .timeBased()
    .everyHours(INV_AUTOSYNC_HOURS)
    .create();
  try {
    SpreadsheetApp.getUi().alert(
      '⏰ 재고 자동 동기화 ON\n\n' +
      INV_AUTOSYNC_HOURS + '시간마다 ERP 재고를 GitHub inventory.json에 자동 업로드합니다.\n\n' +
      'ERP 재고가 바뀌면 최대 ' + INV_AUTOSYNC_HOURS + '시간 + (GitHub Pages 빌드 1~3분) 안에 대시보드에 반영됩니다.'
    );
  } catch (e) {}
}

function removeInventoryAutoSync() {
  var triggers = ScriptApp.getProjectTriggers();
  var removed = 0;
  triggers.forEach(function(t) {
    if (t.getHandlerFunction() === INV_AUTOSYNC_FN) {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });
  if (removed > 0) {
    try { SpreadsheetApp.getUi().alert('⏹ 재고 자동 동기화 OFF (' + removed + '개 트리거 제거)'); } catch (e) {}
  }
}

// 트리거(자동 실행) 컨텍스트에는 UI가 없어 SpreadsheetApp.getUi()가 예외를 던짐 —
// 항상 Logger에는 남기고, UI가 있을 때만 다이얼로그도 띄움
function invAlert_(ui, msg) {
  Logger.log(msg);
  if (ui) { try { ui.alert(msg); } catch (e) {} }
}

function previewErpInventory() {
  var data = authGet(ERP_PRODUCT_LIST_PATH + '?page=1&size=5');
  var list = ((data.response_data || {}).response_list) || [];
  if (!list.length) { SpreadsheetApp.getUi().alert('ERP 응답에 데이터가 없습니다.'); return; }
  var first = list[0];
  var msg = 'ERP 재고 응답 필드 구조 (첫 항목):\n\n';
  Object.keys(first).sort().forEach(function(k) {
    var v = first[k];
    if (v && typeof v === 'object') v = '(object)';
    msg += k + ' = ' + String(v).substring(0, 50) + '\n';
  });
  SpreadsheetApp.getUi().alert(msg);
}

function resetInventoryResume() {
  var props = PropertiesService.getDocumentProperties();
  props.deleteProperty(INV_RESUME_KEY);
  props.deleteProperty(INV_BUFFER_KEY);
  SpreadsheetApp.getUi().alert('🗑️ 진행 상태 초기화 완료');
}

function invExtractQty_(item) {
  var candidates = ['stock_qty','stock_quantity','current_stock','current_qty','qty','quantity','inventory_qty','available_qty','on_hand'];
  for (var i = 0; i < candidates.length; i++) {
    if (item[candidates[i]] !== undefined && item[candidates[i]] !== null) {
      var n = Number(item[candidates[i]]);
      if (!isNaN(n)) return n;
    }
  }
  if (item.stock && typeof item.stock === 'object') {
    for (var j = 0; j < candidates.length; j++) {
      if (item.stock[candidates[j]] !== undefined) {
        var m = Number(item.stock[candidates[j]]);
        if (!isNaN(m)) return m;
      }
    }
  }
  return 0;
}

function invIsTargetFamily_(name) {
  var s = String(name || '');
  for (var i = 0; i < INV_FAMILY_KEYWORDS.length; i++) {
    if (s.indexOf(INV_FAMILY_KEYWORDS[i]) >= 0) return true;
  }
  return false;
}

function invPutGithub_(token, items) {
  var json = {
    items: items,
    updatedAt: Date.now(),
    source: 'erp-manual'
  };
  var contentB64 = Utilities.base64Encode(JSON.stringify(json, null, 1), Utilities.Charset.UTF_8);

  var getUrl = 'https://api.github.com/repos/' + INV_GH_OWNER + '/' + INV_GH_REPO +
               '/contents/' + INV_GH_PATH + '?ref=' + INV_GH_BRANCH;
  var getRes = UrlFetchApp.fetch(getUrl, {
    method: 'get',
    headers: { 'Authorization': 'Bearer ' + token, 'Accept': 'application/vnd.github+json' },
    muteHttpExceptions: true
  });
  var sha = null;
  if (getRes.getResponseCode() === 200) {
    sha = JSON.parse(getRes.getContentText()).sha;
  } else if (getRes.getResponseCode() !== 404) {
    throw new Error('GitHub 파일 조회 실패 (' + getRes.getResponseCode() + '): ' +
                    getRes.getContentText().substring(0, 300));
  }

  var putUrl = 'https://api.github.com/repos/' + INV_GH_OWNER + '/' + INV_GH_REPO +
               '/contents/' + INV_GH_PATH;
  var body = {
    message: 'auto: inventory sync ' + items.length + ' (' + new Date().toISOString() + ')',
    content: contentB64,
    branch: INV_GH_BRANCH
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
    throw new Error('GitHub PUT 실패 (' + code + '): ' + putRes.getContentText().substring(0, 400));
  }
}

function syncInventoryToGitHub() {
  var ui = null;
  try { ui = SpreadsheetApp.getUi(); } catch (e) {} // 트리거(자동 실행) 컨텍스트에는 UI 없음
  var ss = SpreadsheetApp.getActive();
  var props = PropertiesService.getDocumentProperties();
  var startTime = Date.now();

  var token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  if (!token) {
    invAlert_(ui, 'GITHUB_TOKEN 스크립트 속성이 없습니다.\n\nApps Script 편집기 → ⚙️ 프로젝트 설정 → 스크립트 속성에서 GITHUB_TOKEN 추가 후 다시 시도하세요.');
    return;
  }

  try {
    // 이어받기 상태 로드 (30분 지난 상태는 재고 값이 낡았으므로 버리고 처음부터)
    var resumeState = JSON.parse(props.getProperty(INV_RESUME_KEY) || '{"page":1,"totalScanned":0}');
    var buffer = JSON.parse(props.getProperty(INV_BUFFER_KEY) || '[]');
    if (resumeState.savedAt && Date.now() - resumeState.savedAt > 30 * 60 * 1000) {
      props.deleteProperty(INV_RESUME_KEY);
      props.deleteProperty(INV_BUFFER_KEY);
      resumeState = { page: 1, totalScanned: 0 };
      buffer = [];
    }
    var startPage = resumeState.page || 1;
    var totalScanned = resumeState.totalScanned || 0;

    if (startPage > 1) {
      ss.toast('이어받기: 페이지 ' + startPage + '부터 (이전 버퍼 ' + buffer.length + '개)', '⏯️ 재개', 5);
    } else {
      ss.toast('ERP 재고 가져오는 중...', '🔄 시작', 5);
    }

    var PAGE_SIZE = 200;
    var page = startPage, totalPage = 1;
    var timedOut = false;
    while (page <= totalPage || totalPage === 1) {
      if (Date.now() - startTime > INV_MAX_RUN_MS) {
        timedOut = true;
        break;
      }
      var data = authGet(ERP_PRODUCT_LIST_PATH + '?page=' + page + '&size=' + PAGE_SIZE);
      totalPage = (data.page_data || {}).total_page || 1;
      ((data.response_data && data.response_data.response_list) || []).forEach(function(item) {
        if (!item.product_id) return;
        totalScanned++;
        if (!invIsTargetFamily_(item.product_name || '')) return;
        buffer.push({
          productId:   String(item.product_id),
          productCode: item.product_code || '',
          productName: item.product_name || '',
          modelCode:   item.model_code || '',
          brandId:     item.brand_id || '',
          qty:         invExtractQty_(item)
        });
      });
      if (page % 3 === 0 || page === totalPage) ss.toast('페이지 ' + page + ' / ' + totalPage + ' (대상 ' + buffer.length + '개)', '🔄 진행', 3);
      page++;
      if (page > totalPage) break;
      Utilities.sleep(30);
    }

    if (timedOut) {
      props.setProperty(INV_RESUME_KEY, JSON.stringify({ page: page, totalScanned: totalScanned, savedAt: Date.now() }));
      props.setProperty(INV_BUFFER_KEY, JSON.stringify(buffer));
      invAlert_(ui, '⏸️ 시간 초과로 일시 중단\n\n' +
               '현재까지: 페이지 ' + (page - 1) + ' / ' + totalPage + ', 대상 ' + buffer.length + '개\n\n' +
               '다시 메뉴 [▶ GitHub로 동기화] 클릭하거나(또는 자동 동기화 다음 실행 시) 페이지 ' + page + '부터 이어받습니다.');
      return;
    }

    // 페이지 완료 → 이어받기 상태 클리어
    props.deleteProperty(INV_RESUME_KEY);
    props.deleteProperty(INV_BUFFER_KEY);

    if (!buffer.length) {
      invAlert_(ui, 'ERP에서 가져온 ' + totalScanned + '개 중 대상 제품(' +
               INV_FAMILY_KEYWORDS.join('/') + ')이 0개입니다.');
      return;
    }

    // 중복 제거 (product_id 기준)
    var uniqueMap = {};
    buffer.forEach(function(it) { if (!uniqueMap[it.productId]) uniqueMap[it.productId] = it; });
    var unique = Object.values(uniqueMap);

    // GitHub inventory.json에 PUT
    ss.toast('GitHub에 업로드 중...', '⏳', 5);
    invPutGithub_(token, unique);

    var withQty = unique.filter(function(x) { return x.qty > 0; }).length;
    var brandonLines = unique
      .filter(function(x) { return String(x.productName || '').indexOf('브랜드온') >= 0; })
      .map(function(x) { return '  · ' + x.productName + ' (' + x.productCode + '): ' + x.qty + '개'; })
      .join('\n');
    invAlert_(ui, '✅ ERP 재고 → GitHub 동기화 완료\n\n' +
             'ERP 전체 스캔: ' + totalScanned + '개\n' +
             '대상 제품 (' + INV_FAMILY_KEYWORDS.join('/') + '): ' + unique.length + '개\n' +
             '재고 보유: ' + withQty + '개\n' +
             '재고 0: ' + (unique.length - withQty) + '개\n' +
             (brandonLines ? '\n브랜드온 품목 업로드 수량:\n' + brandonLines + '\n' : '') +
             '\n1~3분 후 대시보드 새로고침하면 반영됩니다.');
  } catch (e) {
    invAlert_(ui, '❌ 동기화 실패\n\n' + e.message);
    Logger.log(e.stack || e);
  }
}
