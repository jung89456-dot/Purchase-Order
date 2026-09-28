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
    name: ['수취인명', '수취인이름', '수취자명', '수취인', '받는분', '받는분이름', '받는사람', '수령인', '수령자', '수령인명', '수하인명', '수하인'],
    product: ['상품명', '등록상품명', '구입품목', '품목명', '제품명', '노출상품명(옵션명)', '노출상품명', '상품'],
    option: ['옵션정보', '등록옵션명', '옵션명', '옵션', '옵션내용', '선택옵션'],
    qty: ['수량', '구매수(수량)', '구매수', '주문수량', '수량(개)', '개수'],
    phone: ['수취인연락처1', '수취인전화번호', '수취인휴대폰', '수취인휴대전화', '수취인핸드폰', '받는분전화번호', '받는분연락처', '받는분휴대폰', '수령인연락처', '수령인전화번호', '수하인전화번호', '전화번호', '연락처', '휴대폰번호', '휴대폰'],
    phone2: ['수취인연락처2', '수취인전화번호2', '받는분전화번호2'],
    address: ['통합배송지', '수취인주소', '받는분주소', '수령인주소', '배송지주소', '배송주소', '배송지', '주소', '수하인주소'],
    addressBase: ['기본배송지', '기본주소'],
    addressDetail: ['상세배송지', '상세주소'],
    memo: ['배송메세지', '배송메시지', '배송요청사항', '배송메모', '요청사항', '배송시요청사항', '배송요청메세지'],
    status: ['주문상태', '주문세부상태']
  };

  var REQUIRED = ['name', 'product', 'phone', 'address'];

  function norm(s) {
    return String(s == null ? '' : s).replace(/\s+/g, '').toLowerCase();
  }

  function cellText(v) {
    if (v == null) return '';
    if (v instanceof Date) return v.toISOString().slice(0, 10);
    return String(v).trim();
  }

  // 헤더 행 자동 탐지: 상위 20행 중 알려진 헤더가 가장 많이 포함된 행
  function findHeaderRow(rows) {
    var known = {};
    Object.keys(FIELD_SYNONYMS).forEach(function (f) {
      FIELD_SYNONYMS[f].forEach(function (s) { known[norm(s)] = true; });
    });
    var best = -1, bestScore = 0;
    var limit = Math.min(rows.length, 20);
    for (var i = 0; i < limit; i++) {
      var score = 0;
      (rows[i] || []).forEach(function (c) { if (known[norm(c)]) score++; });
      if (score > bestScore) { bestScore = score; best = i; }
    }
    return bestScore >= 2 ? best : -1;
  }

  function mapColumns(headers) {
    var normed = headers.map(norm);
    var mapping = {};
    Object.keys(FIELD_SYNONYMS).forEach(function (field) {
      var syns = FIELD_SYNONYMS[field];
      for (var i = 0; i < syns.length; i++) {
        var idx = normed.indexOf(norm(syns[i]));
        if (idx !== -1) { mapping[field] = idx; return; }
      }
    });
    return mapping;
  }

  function detectMarket(headers) {
    var set = {};
    headers.forEach(function (h) { set[norm(h)] = true; });
    if (set['상품주문번호'] && (set['수취인연락처1'] || set['통합배송지'])) return 'smartstore';
    if (set['묶음배송번호'] || set['구매수(수량)'] || set['수취인이름'] && set['노출상품명(옵션명)']) return 'coupang';
    return 'etc';
  }

  var MARKET_LABEL = { smartstore: '스마트스토어', coupang: '쿠팡', etc: '기타' };

  function formatPhone(v) {
    var raw = cellText(v);
    if (!raw) return '';
    var d = raw.replace(/\D/g, '');
    // 엑셀 숫자형으로 저장되며 앞자리 0 이 사라진 경우 복원
    if (/^1[016789]\d{7,8}$/.test(d)) d = '0' + d;
    if (/^050\d{9}$/.test(d)) return d.replace(/^(\d{4})(\d{4})(\d{4})$/, '$1-$2-$3');
    if (/^02\d{7,8}$/.test(d)) return d.replace(/^(02)(\d{3,4})(\d{4})$/, '$1-$2-$3');
    if (/^0\d{9,10}$/.test(d)) return d.replace(/^(\d{3})(\d{3,4})(\d{4})$/, '$1-$2-$3');
    if (/^1[5-9]\d{6}$/.test(d)) return d.replace(/^(\d{4})(\d{4})$/, '$1-$2');
    return raw;
  }

  function parseQty(v) {
    var s = cellText(v).replace(/[,\s개]/g, '');
    var n = parseInt(s, 10);
    return isNaN(n) ? null : n;
  }

  function buildItem(product, option) {
    product = cellText(product);
    option = cellText(option);
    if (!option) return product;
    if (!product) return option;
    if (product.indexOf(option) !== -1) return product;
    return product + ' / ' + option;
  }

  /**
   * 2차원 배열(시트) → 변환 결과
   * @returns {{market, marketLabel, headerRow, mapping, missing, rows, skipped}}
   */
  function convertSheet(rows, opts) {
    opts = opts || {};
    var hr = findHeaderRow(rows);
    if (hr === -1) {
      return { error: '주문 데이터의 헤더(수취인명, 상품명 등)를 찾지 못했습니다.' };
    }
    var headers = rows[hr].map(cellText);
    var mapping = mapColumns(headers);
    var market = detectMarket(headers);
    var missing = REQUIRED.filter(function (f) {
      if (f === 'address') return mapping.address == null && mapping.addressBase == null;
      return mapping[f] == null;
    });
    var out = [], skipped = [];
    for (var r = hr + 1; r < rows.length; r++) {
      var row = rows[r] || [];
      var get = function (f) { return mapping[f] == null ? '' : row[mapping[f]]; };
      var name = cellText(get('name'));
      var product = cellText(get('product'));
      if (!name && !product) continue; // 빈 행
      var status = cellText(get('status'));
      if (opts.excludeCanceled !== false && /취소|반품|교환/.test(status)) {
        skipped.push({ row: r + 1, name: name, reason: status });
        continue;
      }
      var address = cellText(get('address'));
      if (!address) {
        address = [cellText(get('addressBase')), cellText(get('addressDetail'))].filter(Boolean).join(' ');
      }
      var phone = formatPhone(get('phone')) || formatPhone(get('phone2'));
      var qty = mapping.qty == null ? 1 : parseQty(get('qty'));
      out.push({
        name: name,
        item: buildItem(product, get('option')),
        qty: qty,
        phone: phone,
        address: address,
        memo: cellText(get('memo')),
        source: MARKET_LABEL[market]
      });
    }
    return {
      market: market,
      marketLabel: MARKET_LABEL[market],
      headerRow: hr + 1,
      headers: headers,
      mapping: mapping,
      missing: missing,
      rows: out,
      skipped: skipped
    };
  }

  var api = {
    OUTPUT_COLUMNS: OUTPUT_COLUMNS,
    FIELD_SYNONYMS: FIELD_SYNONYMS,
    MARKET_LABEL: MARKET_LABEL,
    findHeaderRow: findHeaderRow,
    mapColumns: mapColumns,
    detectMarket: detectMarket,
    formatPhone: formatPhone,
    parseQty: parseQty,
    buildItem: buildItem,
    convertSheet: convertSheet
  };

  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.POParser = api;
})(typeof self !== 'undefined' ? self : this);
