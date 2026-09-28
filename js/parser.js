/*
 * 오픈마켓 주문 데이터 → 당일발주 행 변환 로직 (UI 무관, 브라우저/Node 공용)
 */
(function (root) {
  'use strict';

  // 출력 양식 (첨부된 당일발주 템플릿과 동일한 순서/헤더)
  var OUTPUT_COLUMNS = [
    { key: 'name', header: '수취자명' },
    { key: 'item', header: '구입품목 ' },
    { key: 'qty', header: '수량' },
    { key: 'phone', header: '전화번호' },
    { key: 'address', header: '주소' },
    { key: 'memo', header: '배송메세지' }
  ];

  // 필드별 헤더 후보 (우선순위 순). 공백 제거·소문자 비교.
  var FIELD_SYNONYMS = {
    name: ['수취인명', '수취인이름', '수취자명', '수취인', '받는분', '받는분이름', '받는사람', '수령인', '수령자', '수령인명', '수하인명', '수하인', '이름', '성명'],
    product: ['상품명', '등록상품명', '구입품목', '품목명', '제품명', '노출상품명(옵션명)', '노출상품명', '상품'],
    option: ['옵션정보', '등록옵션명', '옵션명', '옵션', '옵션내용', '선택옵션'],
    qty: ['수량', '구매수(수량)', '구매수', '주문수량', '수량(개)', '개수'],
    phone: ['수취인연락처1', '수취인전화번호', '수취인휴대폰', '수취인휴대전화', '수취인핸드폰', '받는분전화번호', '받는분연락처', '받는분휴대폰', '수령인연락처', '수령인전화번호', '수하인전화번호', '휴대폰번호', '휴대폰', '핸드폰', '휴대전화', '전화번호', '연락처'],
    phone2: ['수취인연락처2', '수취인전화번호2', '받는분전화번호2'],
    address: ['통합배송지', '수취인주소', '받는분주소', '수령인주소', '배송지주소', '배송주소', '배송지', '주소', '수하인주소'],
    addressBase: ['기본배송지', '기본주소'],
    addressDetail: ['상세배송지', '상세주소'],
    memo: ['배송메세지', '배송메시지', '배송요청사항', '배송메모', '요청사항', '배송시요청사항', '배송요청메세지'],
    status: ['주문상태'],
    status2: ['주문세부상태'],
    claim: ['클레임상태', '클레임'],
    orderItemNo: ['상품주문번호'],
    orderNo: ['주문번호'],
    optionId: ['옵션id', '옵션관리코드']
  };

  var FIELD_LABEL = {
    name: '수취자명', product: '상품명', option: '옵션', qty: '수량', phone: '전화번호',
    address: '주소', memo: '배송메세지'
  };

  // 수취인 정보는 구매자/주문자 컬럼에서 가져오지 않는다
  var RECIPIENT_FIELDS = { name: 1, phone: 1, phone2: 1, address: 1 };
  var BUYER_RE = /구매자|주문자|구매인|주문인/;

  var REQUIRED = ['name', 'product', 'phone', 'address'];
  var HEADER_SCAN_ROWS = 50;

  function norm(s) {
    return String(s == null ? '' : s).replace(/\s+/g, '').toLowerCase();
  }

  function cellText(v) {
    if (v == null) return '';
    if (v instanceof Date) return v.toISOString().slice(0, 10);
    return String(v).trim();
  }

  var KNOWN = {};
  Object.keys(FIELD_SYNONYMS).forEach(function (f) {
    FIELD_SYNONYMS[f].forEach(function (s) { KNOWN[norm(s)] = true; });
  });

  // 헤더 행 자동 탐지: 상위 50행 중 알려진 헤더가 가장 많이 포함된 행
  function findHeader(rows) {
    var best = -1, bestScore = 0;
    var limit = Math.min(rows.length, HEADER_SCAN_ROWS);
    for (var i = 0; i < limit; i++) {
      var score = 0;
      (rows[i] || []).forEach(function (c) { if (KNOWN[norm(c)]) score++; });
      if (score > bestScore) { bestScore = score; best = i; }
    }
    return { index: bestScore >= 2 ? best : -1, score: bestScore };
  }

  function findHeaderRow(rows) { return findHeader(rows).index; }

  function mapColumns(headers) {
    var normed = headers.map(norm);
    var mapping = {};
    Object.keys(FIELD_SYNONYMS).forEach(function (field) {
      var syns = FIELD_SYNONYMS[field];
      for (var i = 0; i < syns.length; i++) {
        for (var j = 0; j < normed.length; j++) {
          if (normed[j] !== norm(syns[i])) continue;
          if (RECIPIENT_FIELDS[field] && BUYER_RE.test(normed[j])) continue;
          mapping[field] = j;
          return;
        }
      }
    });
    return mapping;
  }

  function detectMarket(headers) {
    var set = {};
    headers.forEach(function (h) { set[norm(h)] = true; });
    if (set['상품주문번호'] && (set['수취인연락처1'] || set['통합배송지'])) return 'smartstore';
    if (set['묶음배송번호'] || set['구매수(수량)'] || (set['수취인이름'] && set['노출상품명(옵션명)'])) return 'coupang';
    return 'etc';
  }

  var MARKET_LABEL = { smartstore: '스마트스토어', coupang: '쿠팡', etc: '기타' };

  var PHONE_PATTERNS = [
    [/^(050\d)(\d{4})(\d{4})$/, '$1-$2-$3'],     // 안심번호 0504-xxxx-xxxx
    [/^(01[016789])(\d{3,4})(\d{4})$/, '$1-$2-$3'], // 휴대폰
    [/^(02)(\d{3,4})(\d{4})$/, '$1-$2-$3'],      // 서울
    [/^(0[3-6]\d)(\d{3,4})(\d{4})$/, '$1-$2-$3'], // 지역번호
    [/^(070|080)(\d{3,4})(\d{4})$/, '$1-$2-$3'],  // 인터넷전화
    [/^(1[5-9]\d{2})(\d{4})$/, '$1-$2']           // 대표번호
  ];

  function matchPhone(d) {
    for (var i = 0; i < PHONE_PATTERNS.length; i++) {
      if (PHONE_PATTERNS[i][0].test(d)) return d.replace(PHONE_PATTERNS[i][0], PHONE_PATTERNS[i][1]);
    }
    return null;
  }

  /** 전화번호 정규화. 반환: { value, valid } */
  function normalizePhone(v) {
    var raw = cellText(v);
    if (!raw) return { value: '', valid: false };
    var d = raw.replace(/\D/g, '');
    if (/^\s*\+?\s*82/.test(raw)) d = '0' + d.slice(2).replace(/^0/, ''); // +82 10-... 국제 표기
    var hit = matchPhone(d);
    // 엑셀 숫자형으로 저장되며 앞자리 0 이 사라진 경우 복원
    if (!hit && d.charAt(0) !== '0' && d.length >= 8 && d.length <= 11) {
      var restored = matchPhone('0' + d);
      if (restored && !/^1[5-9]\d{6}$/.test(d)) hit = restored;
    }
    return hit ? { value: hit, valid: true } : { value: raw, valid: false };
  }

  function formatPhone(v) { return normalizePhone(v).value; }

  function parseQty(v) {
    var s = cellText(v).replace(/[,\s개]/g, '');
    if (!/^\d+$/.test(s)) return null;
    return parseInt(s, 10);
  }

  function buildItem(product, option) {
    product = cellText(product);
    option = cellText(option);
    if (!option) return product;
    if (!product) return option;
    // 쿠팡 '노출상품명(옵션명)'처럼 상품명 끝에 옵션이 이미 붙어 있는 경우만 생략
    var p = norm(product), o = norm(option);
    if (p === o) return product;
    if (o.length >= 2 && p.length > o.length && p.slice(-o.length) === o && /[,/(\s]/.test(product.charAt(product.length - option.length - 1) || '')) {
      return product;
    }
    return product + ' / ' + option;
  }

  // 발주에서 제외할 주문 상태 판정 (취소철회·교환재배송 등은 정상 출고)
  var KEEP_STATUS_RE = /철회|거부|재배송|교환완료/;
  var EXCLUDE_STATUS_RE = /취소|반품|환불|교환/;
  var SHIPPED_STATUS_RE = /배송중|배송완료|구매확정/;

  function excludeReason(statuses) {
    for (var i = 0; i < statuses.length; i++) {
      var s = statuses[i];
      if (!s || KEEP_STATUS_RE.test(s)) continue;
      if (EXCLUDE_STATUS_RE.test(s) || SHIPPED_STATUS_RE.test(s)) return s;
    }
    return '';
  }

  function isShippedStatus(s) { return SHIPPED_STATUS_RE.test(s || '') && !EXCLUDE_STATUS_RE.test(s || ''); }

  function oneLine(s) {
    return cellText(s).replace(/\s*[\r\n]+\s*/g, ' ');
  }

  /* ---------- 품목 분류 ---------- */

  // 지역명만(예산, 구포) 또는 흔한 채소명만(쪽파, 미나리)으로는 분류하지 않는다 — 다른 산지·상품 오분류 방지.
  // 'A+B' 키워드는 A 와 B 가 둘 다 들어 있을 때만 일치 (예: '[청도] 미나리', '기장 햇쪽파').
  // 키워드는 앱의 '품목 관리'에서 판매자가 직접 고칠 수 있다.
  var DEFAULT_CATALOG = [
    { name: '수제비', keywords: ['수제비'] },
    { name: '칼국수', keywords: ['칼국수'] },
    { name: '구포국수', keywords: ['구포국수', '구포+국수', '구포+소면'] },
    { name: '예산국수', keywords: ['예산국수', '예산+국수'] },
    { name: '기장 쪽파', keywords: ['기장쪽파', '기장+쪽파'] },
    { name: '여수 돌산갓', keywords: ['돌산갓', '여수+갓'] },
    { name: '청도 미나리', keywords: ['청도미나리', '한재미나리', '청도+미나리'] },
    { name: '곡물면', keywords: ['곡물면', '곡물국수'] }
  ];

  function keywordsOf(entry) {
    if (!entry) return [];
    var list = Array.isArray(entry.keywords) ? entry.keywords : String(entry.keywords || '').split(',');
    var ks = list.map(norm).filter(function (k) { return k.replace(/\+/g, ''); });
    if (ks.indexOf(norm(entry.name)) === -1) ks.unshift(norm(entry.name));
    return ks;
  }

  /** 키워드 일치 길이 (조합 키워드는 가장 긴 조각 길이). 불일치면 0 */
  function matchLen(t, k) {
    if (k.indexOf('+') === -1) return t.indexOf(k) !== -1 ? k.length : 0;
    var parts = k.split('+').filter(Boolean), longest = 0;
    for (var i = 0; i < parts.length; i++) {
      if (t.indexOf(parts[i]) === -1) return 0;
      longest = Math.max(longest, parts[i].length);
    }
    return longest;
  }

  /**
   * 텍스트에서 품목을 찾는다.
   * @returns {{name: string, matches: string[]}} 여러 품목이 걸리면 matches.length > 1 (확인 필요)
   */
  function classifyText(text, catalog) {
    var t = norm(text), best = '', bestLen = 0, matches = [], hits = {};
    (catalog || []).forEach(function (entry) {
      if (!entry || !entry.name) return;
      var hit = 0, hitKeys = [];
      keywordsOf(entry).forEach(function (k) {
        var len = matchLen(t, k);
        if (len) hitKeys.push(k);
        if (len > hit) hit = len;
      });
      if (!hit) return;
      matches.push(entry.name);
      hits[entry.name] = hitKeys;
      if (hit > bestLen) { best = entry.name; bestLen = hit; }
    });
    // 다른 품목의 일치가 전부 1위 품목 키워드 안에 포함된 글자라면 (예: '곡물국수' 안의 '국수') 겹침으로 보지 않는다
    if (matches.length > 1) {
      var bestPlain = hits[best].filter(function (k) { return k.indexOf('+') === -1; });
      matches = matches.filter(function (m) {
        if (m === best) return true;
        return !hits[m].every(function (k) {
          return k.indexOf('+') === -1 && bestPlain.some(function (b) { return b !== k && b.indexOf(k) !== -1; });
        });
      });
    }
    return { name: best, matches: matches };
  }

  /** 상품명을 먼저 보고, 상품명에서 못 찾을 때만 옵션을 본다 */
  function classifyProduct(product, option, catalog) {
    var r = classifyText(product, catalog);
    if (!r.name && option) r = classifyText(option, catalog);
    return r;
  }

  function classify(text, catalog) { return classifyText(text, catalog).name; }

  /**
   * 2차원 배열(시트) → 변환 결과
   * @returns {{market, marketLabel, headerRow, mapping, missing, rows}}
   *   rows[i] = { name,item,qty,phone,address,memo, source, key, excluded, issues[] }
   */
  function convertSheet(rows, opts) {
    opts = opts || {};
    var hr = findHeaderRow(rows);
    if (hr === -1) {
      return { error: '주문 목록을 찾지 못했습니다. 스마트스토어 ‘발주/발송관리’ 또는 쿠팡 ‘배송관리’에서 받은 엑셀인지 확인해 주세요.', code: 'NO_HEADER' };
    }
    var headers = rows[hr].map(cellText);
    var mapping = mapColumns(headers);
    var market = detectMarket(headers);
    var missing = REQUIRED.filter(function (f) {
      if (f === 'address') return mapping.address == null && mapping.addressBase == null;
      return mapping[f] == null;
    });
    var out = [];
    for (var r = hr + 1; r < rows.length; r++) {
      var row = rows[r] || [];
      var get = function (f) { return mapping[f] == null ? '' : row[mapping[f]]; };
      var name = cellText(get('name'));
      var product = cellText(get('product'));
      if (!name && !product) continue; // 빈 행

      var address = oneLine(get('address'));
      if (!address) {
        address = [oneLine(get('addressBase')), oneLine(get('addressDetail'))].filter(Boolean).join(' ');
      }
      var ph = normalizePhone(get('phone'));
      if (!ph.value) ph = normalizePhone(get('phone2'));
      var issues = [];
      var qty;
      if (mapping.qty == null || cellText(get('qty')) === '') {
        qty = 1;
        issues.push('qty-default');
      } else {
        qty = parseQty(get('qty'));
        if (qty == null || qty <= 0) { issues.push('qty'); qty = cellText(get('qty')); }
      }
      if (!name) issues.push('name');
      if (!product) issues.push('item');
      if (!ph.value) issues.push('phone');
      else if (!ph.valid) issues.push('phone-format');
      if (!address) issues.push('address');

      var item = buildItem(product, get('option'));
      var orderItemNo = cellText(get('orderItemNo'));
      var key = orderItemNo
        ? 'o:' + orderItemNo
        : cellText(get('orderNo'))
          ? 'n:' + cellText(get('orderNo')) + '|' + cellText(get('optionId')) + '|' + norm(item)
          : 'c:' + [name, ph.value, address, norm(item), qty].join('|');

      out.push({
        name: name,
        item: item,
        qty: qty,
        phone: ph.value,
        address: address,
        memo: oneLine(get('memo')),
        product: product,
        option: cellText(get('option')),
        source: MARKET_LABEL[market],
        market: market,
        sourceRow: r + 1,
        key: key,
        keyKind: key.charAt(0), // o/n: 주문번호 기준, c: 내용 기준(주문번호 없음)
        excluded: excludeReason([cellText(get('status')), cellText(get('status2')), cellText(get('claim'))]),
        issues: issues
      });
    }
    return {
      market: market,
      marketLabel: MARKET_LABEL[market],
      headerRow: hr + 1,
      headers: headers,
      mapping: mapping,
      missing: missing,
      missingLabels: missing.map(function (f) { return FIELD_LABEL[f]; }),
      rows: out
    };
  }

  /** 여러 시트 중 주문 헤더 점수가 가장 높은 시트를 변환 */
  function convertSheets(sheets) {
    var best = null, bestScore = -1;
    sheets.forEach(function (s) {
      var h = findHeader(s.rows);
      if (h.index !== -1 && h.score > bestScore) { best = s; bestScore = h.score; }
    });
    if (!best) return convertSheet([]);
    var res = convertSheet(best.rows);
    res.sheetName = best.name;
    return res;
  }

  var ISSUE_LABEL = {
    name: '수취자명 없음', item: '품목 없음', phone: '전화번호 없음', 'phone-format': '전화번호 형식 확인',
    address: '주소 없음', qty: '수량 오류', 'qty-default': '수량 없음 → 1로 입력',
    'dup-suspect': '같은 주문이 또 있음(확인)', ambiguous: '여러 품목에 해당(품목 확인)'
  };
  // 다운로드 전 확인이 필요한 심각한 문제
  var BLOCKING = { name: 1, item: 1, phone: 1, address: 1, qty: 1 };

  var api = {
    OUTPUT_COLUMNS: OUTPUT_COLUMNS,
    FIELD_SYNONYMS: FIELD_SYNONYMS,
    FIELD_LABEL: FIELD_LABEL,
    MARKET_LABEL: MARKET_LABEL,
    ISSUE_LABEL: ISSUE_LABEL,
    BLOCKING: BLOCKING,
    DEFAULT_CATALOG: DEFAULT_CATALOG,
    findHeaderRow: findHeaderRow,
    mapColumns: mapColumns,
    detectMarket: detectMarket,
    normalizePhone: normalizePhone,
    formatPhone: formatPhone,
    parseQty: parseQty,
    buildItem: buildItem,
    excludeReason: excludeReason,
    isShippedStatus: isShippedStatus,
    classify: classify,
    classifyText: classifyText,
    classifyProduct: classifyProduct,
    convertSheet: convertSheet,
    convertSheets: convertSheets
  };

  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.POParser = api;
})(typeof self !== 'undefined' ? self : this);
