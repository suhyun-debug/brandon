/**
 * 오프라인 제휴 현황 시트 → Firestore (kpiSummary/offline) 동기화
 *
 * 사용:
 *   - 구글시트의 Apps Script 편집기에서 본 파일 내용을 추가
 *   - 메뉴 "🔄 KPI 동기화" 실행 또는 트리거로 자동 실행
 *
 * 시트 구조 가정 (스크린샷 기준):
 *   각 제품(에어리소닉/소닉플로우/슈퍼플로우) 블록의 마지막 행이 "총 계"
 *   해당 행의 같은 블록 안에서 "지점수" / "제휴 수량(개)" 컬럼 값을 읽음
 */

var SHEET_NAME = '오프라인 제휴 현황';
var PRODUCTS   = ['에어리소닉', '소닉플로우', '슈퍼플로우'];

// Firebase 프로젝트 정보 (브라우저 코드와 동일)
var FB_PROJECT = 'brandon-2dbe9';
var FB_API_KEY = 'AIzaSyCmc__AgQak-is4LpKsNqArvTtTbXRFtjQ';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🔄 KPI 동기화')
    .addItem('오프라인 제휴 현황 → Firestore', 'syncOfflineKpi')
    .addItem('출고 이력 (에어리소닉/소닉플로우) → Firestore', 'syncShipments')
    .addToUi();
}

/* ============ 출고 이력 동기화 ============
 * "에어리소닉 출고 이력" / "소닉플로우 출고 이력" 시트 → Firestore shipments 컬렉션
 *   - 헤더 자동 인식: 날짜/출고일, 사업장명/거래처명, 수량, 제품(없으면 시트명 기준 기본값), 업종/카테고리
 *   - 문서 ID = "{시트명}_row{행번호}" → 재실행 시 upsert (중복 방지)
 *   - 시트에서 행 삭제된 경우의 stale 데이터는 자동 정리하지 않음(필요 시 수동)
 */
var SHIPMENT_SHEETS = [
  { name: '에어리소닉 출고 이력', defaultProduct: '에어리소닉' },
  { name: '소닉플로우 출고 이력', defaultProduct: '소닉플로우' }
];

function syncShipments() {
  var summary = [];
  var totalCount = 0;
  SHIPMENT_SHEETS.forEach(function(cfg) {
    var r = syncShipmentSheet_(cfg.name, cfg.defaultProduct);
    summary.push(cfg.name + ': ' + r.count + '건 (cols=' + r.colsStr + ')');
    totalCount += r.count;
  });
  Logger.log(summary.join('\n'));
  SpreadsheetApp.getActive().toast('출고 이력 동기화 완료: ' + totalCount + '건\n' + summary.join('\n'), '✅', 8);
}

function syncShipmentSheet_(sheetName, defaultProduct) {
  var sh = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sh) throw new Error('시트 없음: ' + sheetName);
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return { count: 0, colsStr: '(no data)' };

  var col = detectShipmentCols_(values);
  if (col.date < 0 || col.name < 0 || col.qty < 0) {
    throw new Error('필수 컬럼 미발견 [' + sheetName + ']: 날짜/사업장명/수량 헤더 확인 필요. 감지된 헤더: ' + JSON.stringify(values[col.headerRow]));
  }
  var colsStr = '날짜=' + colLet_(col.date) + ',사업장=' + colLet_(col.name) + ',수량=' + colLet_(col.qty) +
                (col.product>=0?',제품='+colLet_(col.product):'') +
                (col.category>=0?',업종='+colLet_(col.category):'') +
                (col.address>=0?',주소='+colLet_(col.address):'');

  var rows = [];
  for (var r = col.headerRow + 1; r < values.length; r++) {
    var v = values[r];
    var dateStr = formatDate_(v[col.date]);
    var nm = String(v[col.name] || '').trim();
    var q = toInt_(v[col.qty]);
    if (!dateStr || !nm || !q) continue;
    rows.push({
      docId: sheetName.replace(/\s/g, '_') + '_row' + (r + 1),
      partnerName: nm,
      date: dateStr,
      qty: q,
      product: col.product >= 0 ? (String(v[col.product] || '').trim() || defaultProduct) : defaultProduct,
      category: col.category >= 0 ? String(v[col.category] || '').trim() : '',
      address: col.address >= 0 ? String(v[col.address] || '').trim() : ''
    });
  }

  for (var i = 0; i < rows.length; i += 100) {
    commitShipmentBatch_(rows.slice(i, i + 100), sheetName);
  }
  return { count: rows.length, colsStr: colsStr };
}

function detectShipmentCols_(values) {
  // 헤더 행 찾기 (위에서 5행 이내에서 "날짜" 또는 "출고일"이 등장하는 행)
  var headerRow = 0;
  for (var r = 0; r < Math.min(5, values.length); r++) {
    var found = false;
    for (var c = 0; c < values[r].length; c++) {
      var h = String(values[r][c] || '').replace(/\s/g, '');
      if (h === '날짜' || h === '출고일' || h === '출고일자' || h === '등록일시') { headerRow = r; found = true; break; }
    }
    if (found) break;
  }
  var hdr = values[headerRow] || [];
  var col = { headerRow: headerRow, date: -1, name: -1, qty: -1, product: -1, category: -1, address: -1 };
  hdr.forEach(function(raw, i) {
    var h = String(raw || '').replace(/\s/g, '');
    if (col.date < 0 && (h === '날짜' || h === '출고일' || h === '출고일자' || h === '등록일시')) col.date = i;
    if (col.name < 0 && (h === '사업장명' || h === '사업장' || h === '거래처명' || h === '거래처' || h === '업체명' || h === '업체')) col.name = i;
    if (col.qty < 0 && (h === '수량' || h === '출고수량' || h === '대수')) col.qty = i;
    if (col.product < 0 && (h === '제품' || h === '제품명' || h === '품목')) col.product = i;
    if (col.category < 0 && (h === '업종' || h === '카테고리' || h === '분류' || h === '대분류')) col.category = i;
    if (col.address < 0 && (h === '주소' || h === '지점주소' || h === '상세주소' || h === '도로명주소' || h === '지번주소' || h === '업체주소' || h === '사업장주소')) col.address = i;
  });
  return col;
}

function colLet_(i) {
  var s = '';
  i = Number(i);
  while (i >= 0) { s = String.fromCharCode(65 + (i % 26)) + s; i = Math.floor(i / 26) - 1; }
  return s;
}

function formatDate_(v) {
  if (v == null || v === '') return null;
  if (v instanceof Date && !isNaN(v)) {
    return v.getFullYear() + '-' + pad2_(v.getMonth() + 1) + '-' + pad2_(v.getDate());
  }
  var s = String(v).trim();
  var m = s.match(/^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})/);
  if (m) return m[1] + '-' + pad2_(m[2]) + '-' + pad2_(m[3]);
  // 엑셀 직렬값
  var n = Number(s);
  if (isFinite(n) && n > 30000 && n < 80000) {
    var d = new Date(Math.round((n - 25569) * 86400 * 1000));
    if (!isNaN(d)) return d.getFullYear() + '-' + pad2_(d.getMonth() + 1) + '-' + pad2_(d.getDate());
  }
  return null;
}
function pad2_(n) { n = String(n); return n.length < 2 ? '0' + n : n; }

function commitShipmentBatch_(rows, sheetName) {
  if (!rows.length) return;
  var now = Date.now();
  var src = 'sheet:' + sheetName;
  var writes = rows.map(function(r) {
    return {
      update: {
        name: 'projects/' + FB_PROJECT + '/databases/(default)/documents/shipments/' + r.docId,
        fields: {
          partnerName: { stringValue: r.partnerName },
          date:        { stringValue: r.date },
          qty:         { integerValue: String(r.qty) },
          product:     { stringValue: r.product },
          category:    { stringValue: r.category },
          address:     { stringValue: r.address || '' },
          source:      { stringValue: src },
          updatedAt:   { integerValue: String(now) }
        }
      }
    };
  });
  var url = 'https://firestore.googleapis.com/v1/projects/' + FB_PROJECT +
            '/databases/(default)/documents:commit?key=' + FB_API_KEY;
  var res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ writes: writes }),
    muteHttpExceptions: true
  });
  var code = res.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error('Commit 실패 (' + code + '): ' + res.getContentText());
  }
}

/* ============ 오프라인 제휴 현황 KPI 동기화 ============ */

function syncOfflineKpi() {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  if (!sh) throw new Error('시트를 찾을 수 없음: ' + SHEET_NAME);

  var values = sh.getDataRange().getValues();
  var byProduct = {};
  var totalPartners = 0, totalQty = 0;
  var byCategoryAll = {}; // 전 제품 합산 카테고리 집계

  PRODUCTS.forEach(function(prod) {
    var hit = findProductDetails_(values, prod);
    if (hit) {
      byProduct[prod] = { partners: hit.partners, qty: hit.qty, byCategory: hit.byCategory || {} };
      totalPartners += hit.partners;
      totalQty      += hit.qty;
      Object.keys(hit.byCategory || {}).forEach(function(cat) {
        var v = hit.byCategory[cat];
        if (!byCategoryAll[cat]) byCategoryAll[cat] = { partners: 0, qty: 0 };
        byCategoryAll[cat].partners += v.partners;
        byCategoryAll[cat].qty      += v.qty;
      });
    } else {
      byProduct[prod] = { partners: 0, qty: 0, byCategory: {} };
    }
  });

  function catMapFields_(obj) {
    return Object.keys(obj).reduce(function(acc, c) {
      acc[c] = { mapValue: { fields: {
        partners: { integerValue: String(obj[c].partners) },
        qty:      { integerValue: String(obj[c].qty) }
      }}};
      return acc;
    }, {});
  }

  var payload = {
    fields: {
      partners:  { integerValue: String(totalPartners) },
      qty:       { integerValue: String(totalQty) },
      updatedAt: { integerValue: String(Date.now()) },
      byCategory: {
        mapValue: { fields: catMapFields_(byCategoryAll) }
      },
      byProduct: {
        mapValue: { fields: Object.keys(byProduct).reduce(function(acc, p) {
          acc[p] = { mapValue: { fields: {
            partners:   { integerValue: String(byProduct[p].partners) },
            qty:        { integerValue: String(byProduct[p].qty) },
            byCategory: { mapValue: { fields: catMapFields_(byProduct[p].byCategory || {}) } }
          }}};
          return acc;
        }, {}) }
      }
    }
  };

  var url = 'https://firestore.googleapis.com/v1/projects/' + FB_PROJECT +
            '/databases/(default)/documents/kpiSummary/offline?key=' + FB_API_KEY;
  var res = UrlFetchApp.fetch(url, {
    method: 'patch',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  var code = res.getResponseCode();
  var body = res.getContentText();
  if (code >= 200 && code < 300) {
    Logger.log('동기화 완료: 총 ' + totalPartners + '개소 / ' + totalQty + '대');
    SpreadsheetApp.getActive().toast('KPI 동기화 완료: ' + totalPartners + '개소 / ' + totalQty + '대');
  } else {
    Logger.log('실패 ' + code + ': ' + body);
    throw new Error('Firestore 쓰기 실패 (' + code + '): ' + body);
  }
}

/**
 * 시트 2차원 배열에서 특정 제품 블록을 찾아 카테고리별/총계 지점수·제휴수량 반환
 * 제품 헤더 셀의 컬럼을 블록 기준으로 잡고, 서브헤더(지점수/제휴수량) 아래의 행들을
 * 카테고리 라벨로 보아 집계하다가 "총 계"를 만나면 종료
 */
function findProductDetails_(values, productName) {
  var headerRow = -1, headerCol = -1;
  for (var r = 0; r < values.length && headerRow < 0; r++) {
    for (var c = 0; c < values[r].length; c++) {
      var cell = String(values[r][c] || '').trim();
      if (cell === productName) { headerRow = r; headerCol = c; break; }
    }
  }
  if (headerRow < 0) return null;

  // "지점수" / "제휴 수량" 컬럼 위치 (헤더 아래 1~3행 내에서 탐색)
  var colCntPartners = -1, colCntQty = -1, subHeaderRow = -1;
  for (var rr = headerRow; rr < Math.min(headerRow + 4, values.length); rr++) {
    for (var cc = Math.max(0, headerCol - 1); cc < Math.min(values[rr].length, headerCol + 6); cc++) {
      var v = String(values[rr][cc] || '').replace(/\s/g, '');
      if (v === '지점수') { colCntPartners = cc; subHeaderRow = rr; }
      if (v.indexOf('제휴수량') === 0) { colCntQty = cc; subHeaderRow = rr; }
    }
    if (colCntPartners >= 0 && colCntQty >= 0) break;
  }
  if (colCntPartners < 0 || colCntQty < 0) return null;

  var byCategory = {};
  var totals = null;
  for (var rrr = subHeaderRow + 1; rrr < Math.min(subHeaderRow + 40, values.length); rrr++) {
    var rawLabel = String(values[rrr][headerCol] || '').trim();
    var label = rawLabel.replace(/\s/g, '');
    if (!label) continue;
    if (label === '총계' || label === '합계') {
      totals = {
        partners: toInt_(values[rrr][colCntPartners]),
        qty:      toInt_(values[rrr][colCntQty])
      };
      break;
    }
    // 다음 제품 블록 헤더를 만나면 중단
    if (PRODUCTS.indexOf(rawLabel) >= 0 && rawLabel !== productName) break;
    var partners = toInt_(values[rrr][colCntPartners]);
    var qty      = toInt_(values[rrr][colCntQty]);
    if (partners === 0 && qty === 0) continue;
    byCategory[rawLabel] = { partners: partners, qty: qty };
  }
  if (!totals) return null;
  return { partners: totals.partners, qty: totals.qty, byCategory: byCategory };
}

// 하위 호환을 위해 이전 이름도 유지 (호출부 없으면 제거 가능)
function findProductTotals_(values, productName) {
  var d = findProductDetails_(values, productName);
  return d ? { partners: d.partners, qty: d.qty } : null;
}

function toInt_(v) {
  if (v === '' || v == null || v === '-') return 0;
  var n = Number(String(v).replace(/[,\s]/g, ''));
  return isFinite(n) ? Math.round(n) : 0;
}
