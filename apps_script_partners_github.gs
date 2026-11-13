/**
 * 거래처(partners) → GitHub data.json 동기화
 *
 * 동작:
 *   1. "오프라인 제휴 현황" 시트에서 거래처 목록 파싱
 *   2. 기존 GitHub data.json에서 좌표(lat/lng) 가져와 매칭 (이름+제품 키)
 *      → 시트에 없는 거래처는 자동 삭제됨 (시트가 단일 진실 소스)
 *   3. GitHub repo의 data.json에 PUT (Contents API)
 *   4. GitHub Pages가 1~3분 내 자동 재배포 → 대시보드 반영
 *
 * 사전 설정 (1회):
 *   1. GitHub Personal Access Token 발급
 *      https://github.com/settings/tokens?type=beta
 *      - Repository access: hulkyaya/brandon 만 선택
 *      - Permissions → Contents: Read and write
 *      - 발급된 토큰 복사
 *   2. Apps Script 편집기 → 프로젝트 설정 → 스크립트 속성에 추가:
 *      이름: GITHUB_TOKEN   값: ghp_... 또는 github_pat_...
 *
 * 사용:
 *   - 메뉴 [🌐 GitHub 동기화] → [▶ 거래처 data.json 업로드]
 *   - 또는 [⏰ 1시간 자동 동기화 켜기]
 */

var GH_OWNER  = 'hulkyaya';
var GH_REPO   = 'brandon';
var GH_BRANCH = 'main';
var GH_PATH   = 'data.json';
var GH_AUTOSYNC_FN = 'autoSyncPartnersToGitHub';
var GH_AUTOSYNC_HOURS = 1;

function ghAddMenu_() {
  SpreadsheetApp.getUi()
    .createMenu('🌐 GitHub 동기화')
    .addItem('▶ 거래처 data.json 업로드', 'syncPartnersToGitHub')
    .addItem('📍 주소 → 좌표 변환 후 업로드', 'geocodeAndSync')
    .addSeparator()
    .addItem('⏰ 1시간 자동 동기화 켜기', 'installPartnersGithubAutoSync')
    .addItem('⏹ 자동 동기화 끄기', 'removePartnersGithubAutoSync')
    .addSeparator()
    .addItem('🔍 GITHUB_TOKEN 확인', 'ghCheckToken')
    .addToUi();
}
// onOpen()에 ghAddMenu_() 호출 한 줄 추가하거나, 별도 트리거로 등록
// 기존 onOpen 함수가 있다면 그 안에 ghAddMenu_(); 한 줄만 추가

function ghCheckToken() {
  var token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  if (!token) {
    SpreadsheetApp.getUi().alert(
      '❌ GITHUB_TOKEN 미설정\n\n' +
      'Apps Script 편집기 → ⚙️ 프로젝트 설정 → 스크립트 속성에서\n' +
      'GITHUB_TOKEN 추가 후 다시 시도하세요.\n\n' +
      '토큰 발급: https://github.com/settings/tokens?type=beta'
    );
    return;
  }
  // 토큰으로 repo 접근 가능한지 확인
  var url = 'https://api.github.com/repos/' + GH_OWNER + '/' + GH_REPO;
  var res = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { 'Authorization': 'Bearer ' + token, 'Accept': 'application/vnd.github+json' },
    muteHttpExceptions: true
  });
  if (res.getResponseCode() === 200) {
    var info = JSON.parse(res.getContentText());
    SpreadsheetApp.getUi().alert(
      '✅ 토큰 정상\n\n' +
      'Repo: ' + info.full_name + '\n' +
      'Default branch: ' + info.default_branch + '\n' +
      'Private: ' + info.private
    );
  } else {
    SpreadsheetApp.getUi().alert(
      '❌ 토큰 인증 실패 (HTTP ' + res.getResponseCode() + ')\n\n' +
      res.getContentText().substring(0, 300)
    );
  }
}

function syncPartnersToGitHub() {
  var token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  if (!token) throw new Error('GITHUB_TOKEN 스크립트 속성이 설정되지 않았습니다. 메뉴 [🔍 GITHUB_TOKEN 확인]을 먼저 실행하세요.');

  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('오프라인 제휴 현황');
  if (!sh) throw new Error('"오프라인 제휴 현황" 시트를 찾을 수 없습니다.');
  var values = sh.getDataRange().getValues();

  // 1) 시트에서 거래처 파싱 (기존 parseOfflineDetails_ 재사용)
  var entries = parseOfflineDetails_(values);
  if (!entries.length) throw new Error('시트에서 거래처 0건 파싱됨. "오프라인 제휴 현황" 시트 구조를 확인하세요.');

  // 2) 기존 GitHub data.json에서 좌표 가져오기 (Firestore 의존 없음)
  var coordMap = ghLoadExistingCoords_(token);

  // 3) 거래처 배열 빌드 (이름+제품 dedup) — 시트가 단일 진실 소스
  var seen = {};
  var partners = [];
  entries.forEach(function(e) {
    var key = (e.name || '').trim() + '|' + (e.product || '').trim();
    if (!key || seen[key]) return;
    seen[key] = true;
    var p = {
      name: e.name,
      product: e.product,
      category: e.category || '미분류',
      qty: e.qty || 0
    };
    var co = coordMap[key];
    if (co && co.lat && co.lng) {
      p.lat = co.lat;
      p.lng = co.lng;
    }
    partners.push(p);
  });

  var json = {
    partners: partners,
    updatedAt: Date.now(),
    source: 'apps-script-sheet-sync'
  };

  // 4) GitHub Contents API로 PUT
  var contentB64 = Utilities.base64Encode(JSON.stringify(json, null, 1), Utilities.Charset.UTF_8);
  var getUrl = 'https://api.github.com/repos/' + GH_OWNER + '/' + GH_REPO +
               '/contents/' + GH_PATH + '?ref=' + GH_BRANCH;
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

  var putUrl = 'https://api.github.com/repos/' + GH_OWNER + '/' + GH_REPO + '/contents/' + GH_PATH;
  var body = {
    message: 'auto: partners sync ' + partners.length + ' (' + new Date().toISOString() + ')',
    content: contentB64,
    branch: GH_BRANCH
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

  var withCoord = partners.filter(function(p) { return p.lat && p.lng; }).length;
  ss.toast('GitHub 동기화 완료: ' + partners.length + '개 거래처 (좌표 ' + withCoord + '개)', '✅', 8);
  Logger.log('partners synced: ' + partners.length + ' (with coords: ' + withCoord + ')');
}

function ghLoadExistingCoords_(token) {
  // 기존 data.json을 GitHub raw로 읽어 이름+제품 → 좌표 맵 빌드
  var url = 'https://raw.githubusercontent.com/' + GH_OWNER + '/' + GH_REPO + '/' + GH_BRANCH + '/' + GH_PATH;
  var res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  var coordMap = {};
  if (res.getResponseCode() !== 200) {
    Logger.log('기존 data.json 없음 (첫 동기화) — 좌표 0개');
    return coordMap;
  }
  try {
    var data = JSON.parse(res.getContentText());
    (data.partners || []).forEach(function(p) {
      if (!p.name || !p.lat || !p.lng) return;
      var key = (p.name || '').trim() + '|' + (p.product || '').trim();
      coordMap[key] = { lat: Number(p.lat), lng: Number(p.lng) };
    });
    Logger.log('기존 좌표 ' + Object.keys(coordMap).length + '개 로드');
  } catch (e) {
    Logger.log('data.json 파싱 실패: ' + e.message);
  }
  return coordMap;
}

function autoSyncPartnersToGitHub() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    Logger.log('autoSyncPartnersToGitHub: 다른 실행이 진행 중이라 건너뜀');
    return;
  }
  try {
    syncPartnersToGitHub();
  } catch (e) {
    Logger.log('autoSyncPartnersToGitHub 실패: ' + (e && e.message));
  } finally {
    lock.releaseLock();
  }
}

function installPartnersGithubAutoSync() {
  removePartnersGithubAutoSync();
  ScriptApp.newTrigger(GH_AUTOSYNC_FN)
    .timeBased()
    .everyHours(GH_AUTOSYNC_HOURS)
    .create();
  try {
    SpreadsheetApp.getUi().alert(
      '⏰ GitHub 자동 동기화 ON\n\n' +
      GH_AUTOSYNC_HOURS + '시간마다 시트의 거래처를 GitHub data.json에 자동 업로드합니다.\n\n' +
      '시트에 새 거래처를 입력하면 최대 ' + GH_AUTOSYNC_HOURS + '시간 + (GitHub Pages 빌드 1~3분) 안에 대시보드에 반영됩니다.'
    );
  } catch (e) {}
}

function removePartnersGithubAutoSync() {
  var triggers = ScriptApp.getProjectTriggers();
  var removed = 0;
  triggers.forEach(function(t) {
    if (t.getHandlerFunction() === GH_AUTOSYNC_FN) {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });
  if (removed > 0) {
    try { SpreadsheetApp.getUi().alert('⏹ 자동 동기화 OFF (' + removed + '개 트리거 제거)'); } catch (e) {}
  }
}

/* ============ 주소 → 좌표 변환 후 data.json 업로드 ============
 *
 * 사전 설정 (1회):
 *   Apps Script → ⚙️ 프로젝트 설정 → 스크립트 속성에 추가
 *   이름: KAKAO_REST_KEY   값: 카카오 REST API 키
 *   (https://developers.kakao.com → 내 애플리케이션 → 앱 키 → REST API 키)
 *
 * 동작:
 *   1. "오프라인 제휴 현황" 시트에서 거래처 목록 파싱
 *   2. "에어리소닉/소닉플로우 출고 이력" 시트의 주소 컬럼으로 거래처명→주소 맵 빌드
 *   3. 기존 data.json 좌표 보존, 좌표 없는 거래처만 Kakao Local API로 변환
 *   4. 전체 파트너 data.json으로 GitHub 업로드 (1~3분 내 대시보드 반영)
 */
function geocodeAndSync() {
  var token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  if (!token) throw new Error('GITHUB_TOKEN 스크립트 속성이 설정되지 않았습니다.');
  var kakaoKey = PropertiesService.getScriptProperties().getProperty('KAKAO_REST_KEY');
  if (!kakaoKey) throw new Error(
    'KAKAO_REST_KEY 스크립트 속성이 설정되지 않았습니다.\n\n' +
    '카카오 개발자 콘솔(https://developers.kakao.com)에서\n' +
    'REST API 키 발급 후 Apps Script 스크립트 속성에 저장하세요.'
  );

  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) throw new Error('시트 없음: ' + SHEET_NAME);

  // 1. 파트너 목록 파싱 (오프라인 제휴 현황)
  var entries = parseOfflineDetails_(sh.getDataRange().getValues());
  if (!entries.length) throw new Error('파트너 0건. 시트 구조를 확인하세요.');
  Logger.log('파트너 ' + entries.length + '건 파싱');

  // 2. 출고 이력 시트에서 거래처명 → 주소 맵 빌드
  var addressMap = buildAddressMap_();
  Logger.log('주소 맵: ' + Object.keys(addressMap).length + '개 거래처');

  // 3. 기존 data.json 좌표 로드 (이미 좌표 있는 거래처는 재변환 안 함)
  var coordMap = ghLoadExistingCoords_(token);
  Logger.log('기존 좌표: ' + Object.keys(coordMap).length + '개');

  // 4. 미좌표 거래처 Kakao API 변환
  var seenGeo = {};
  var cnt = { already: 0, geocoded: 0, noAddr: 0, failed: 0 };

  entries.forEach(function(e) {
    var nm   = (e.name    || '').trim();
    var key  = nm + '|' + (e.product || '').trim();
    if (!nm) return;

    if (coordMap[key] && coordMap[key].lat) { cnt.already++; return; }

    // 같은 거래처명을 이미 시도했으면 결과 재사용
    if (seenGeo.hasOwnProperty(nm)) {
      if (seenGeo[nm]) coordMap[key] = seenGeo[nm];
      return;
    }

    var addr = addressMap[nm];
    if (!addr) { seenGeo[nm] = null; cnt.noAddr++; return; }

    var coord = kakaoGeocode_(addr, kakaoKey);
    if (coord) {
      seenGeo[nm] = coord;
      coordMap[key] = coord;
      cnt.geocoded++;
      Logger.log('✓ ' + nm + ' → ' + coord.lat.toFixed(5) + ',' + coord.lng.toFixed(5) + ' [' + addr + ']');
    } else {
      seenGeo[nm] = null;
      cnt.failed++;
      Logger.log('✗ 좌표 없음: ' + nm + ' [' + addr + ']');
    }
    Utilities.sleep(150); // 카카오 API 초당 10건 제한 대응
  });

  Logger.log('기존:' + cnt.already + ' 신규:' + cnt.geocoded + ' 주소없음:' + cnt.noAddr + ' 실패:' + cnt.failed);

  // 5. Partners 배열 빌드
  var seen = {}, partners = [];
  entries.forEach(function(e) {
    var key = (e.name || '').trim() + '|' + (e.product || '').trim();
    if (!key || seen[key]) return;
    seen[key] = true;
    var p = { name: e.name, product: e.product, category: e.category || '미분류', qty: e.qty || 0 };
    var co = coordMap[key];
    if (co && co.lat && co.lng) { p.lat = co.lat; p.lng = co.lng; }
    partners.push(p);
  });

  // 6. GitHub data.json 업로드
  var json = { partners: partners, updatedAt: Date.now(), source: 'apps-script-geocode-sync' };
  var contentB64 = Utilities.base64Encode(JSON.stringify(json, null, 1), Utilities.Charset.UTF_8);

  var getUrl = 'https://api.github.com/repos/' + GH_OWNER + '/' + GH_REPO + '/contents/' + GH_PATH + '?ref=' + GH_BRANCH;
  var getRes = UrlFetchApp.fetch(getUrl, {
    method: 'get',
    headers: { Authorization: 'Bearer ' + token, Accept: 'application/vnd.github+json' },
    muteHttpExceptions: true
  });
  var sha = (getRes.getResponseCode() === 200) ? JSON.parse(getRes.getContentText()).sha : null;

  var body = {
    message: 'auto: geocode ' + cnt.geocoded + ' new coords (' + partners.length + ' partners total)',
    content: contentB64,
    branch:  GH_BRANCH
  };
  if (sha) body.sha = sha;

  var putRes = UrlFetchApp.fetch(
    'https://api.github.com/repos/' + GH_OWNER + '/' + GH_REPO + '/contents/' + GH_PATH,
    {
      method: 'put',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + token, Accept: 'application/vnd.github+json' },
      payload: JSON.stringify(body),
      muteHttpExceptions: true
    }
  );

  var code = putRes.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error('GitHub PUT 실패 (' + code + '): ' + putRes.getContentText().substring(0, 300));
  }

  var withCoord = partners.filter(function(p) { return p.lat && p.lng; }).length;
  ss.toast(
    '거래처 ' + partners.length + '개 / 좌표 있음 ' + withCoord + '개 / 신규 변환 ' + cnt.geocoded + '개\n' +
    '1~3분 내 대시보드 지도에 반영됩니다.',
    '📍 좌표 변환 완료', 10
  );
  Logger.log('업로드 완료: ' + partners.length + '개 (좌표 ' + withCoord + '개)');
}

/**
 * 출고 이력 시트(에어리소닉/소닉플로우)에서 거래처명 → 주소 맵 빌드
 * SHIPMENT_SHEETS, detectShipmentCols_ 는 apps_script_kpi_sync.gs 에 정의됨
 */
function buildAddressMap_() {
  var addressMap = {};
  var ADDR_KEYS = ['주소', '지점주소', '상세주소', '도로명주소', '지번주소', '업체주소', '사업장주소'];

  SHIPMENT_SHEETS.forEach(function(cfg) {
    var sh = SpreadsheetApp.getActive().getSheetByName(cfg.name);
    if (!sh) { Logger.log(cfg.name + ': 시트 없음'); return; }
    var values = sh.getDataRange().getValues();
    if (values.length < 2) return;

    var col = detectShipmentCols_(values);
    if (col.name < 0) { Logger.log(cfg.name + ': 거래처명 컬럼 없음'); return; }

    var hdr = values[col.headerRow] || [];
    var addrCol = -1;
    for (var i = 0; i < hdr.length; i++) {
      if (ADDR_KEYS.indexOf(String(hdr[i] || '').replace(/\s/g, '')) >= 0) { addrCol = i; break; }
    }
    if (addrCol < 0) {
      Logger.log(cfg.name + ': 주소 컬럼 없음 — 헤더: ' + hdr.slice(0, 12).join(', '));
      return;
    }

    var before = Object.keys(addressMap).length;
    for (var r = col.headerRow + 1; r < values.length; r++) {
      var nm   = String(values[r][col.name]  || '').trim();
      var addr = String(values[r][addrCol]   || '').trim();
      if (nm && addr && !addressMap[nm]) addressMap[nm] = addr;
    }
    Logger.log(cfg.name + ': ' + (Object.keys(addressMap).length - before) + '개 주소 수집');
  });
  return addressMap;
}

/**
 * Kakao Local API 주소→좌표 변환
 * 1차: 주소 검색 / 2차: 키워드 검색 (주소 실패 시)
 */
function kakaoGeocode_(address, key) {
  var headers = { Authorization: 'KakaoAK ' + key };
  try {
    var r = UrlFetchApp.fetch(
      'https://dapi.kakao.com/v2/local/search/address.json?query=' + encodeURIComponent(address),
      { headers: headers, muteHttpExceptions: true }
    );
    if (r.getResponseCode() === 200) {
      var docs = JSON.parse(r.getContentText()).documents;
      if (docs && docs.length) return { lat: Number(docs[0].y), lng: Number(docs[0].x) };
    }
  } catch (e) { Logger.log('주소검색 오류: ' + e.message); }

  try {
    var r2 = UrlFetchApp.fetch(
      'https://dapi.kakao.com/v2/local/search/keyword.json?query=' + encodeURIComponent(address),
      { headers: headers, muteHttpExceptions: true }
    );
    if (r2.getResponseCode() === 200) {
      var docs2 = JSON.parse(r2.getContentText()).documents;
      if (docs2 && docs2.length) return { lat: Number(docs2[0].y), lng: Number(docs2[0].x) };
    }
  } catch (e) { Logger.log('키워드검색 오류: ' + e.message); }

  return null;
}
