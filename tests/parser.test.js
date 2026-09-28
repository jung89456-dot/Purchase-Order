const test = require('node:test');
const assert = require('node:assert');
const path = require('path');
const XLSX = require('xlsx');
const P = require('../js/parser.js');

function sheetRows(file) {
  const wb = XLSX.readFile(path.join(__dirname, '..', 'samples', file));
  return XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: false, defval: '' });
}

test('스마트스토어 샘플 변환', () => {
  const r = P.convertSheet(sheetRows('스마트스토어_주문샘플.xlsx'));
  assert.strictEqual(r.market, 'smartstore');
  assert.deepStrictEqual(r.missing, []);
  assert.strictEqual(r.rows.length, 5);
  assert.strictEqual(r.rows.filter((x) => x.excluded).length, 1); // 취소요청
  const first = r.rows[0];
  assert.strictEqual(first.name, '김하늘');
  assert.strictEqual(first.item, '[손칼국수] 쫄깃한 생칼국수 1kg / 4인분');
  assert.strictEqual(first.qty, 1);
  assert.strictEqual(first.phone, '010-1234-5678');
  assert.strictEqual(first.address, '서울특별시 강남구 테헤란로 123 4층 401호');
  assert.strictEqual(first.memo, '부재 시 문 앞에 놓아주세요');
  assert.strictEqual(r.rows[1].phone, '010-9876-5432');
  assert.strictEqual(r.rows[2].phone, '0504-1111-2222');
  assert.strictEqual(r.rows[3].item, '옛날 구포국수 3kg');
});

test('쿠팡 샘플 변환 — 구매자가 아닌 수취인 정보 사용', () => {
  const r = P.convertSheet(sheetRows('쿠팡_주문샘플.xlsx'));
  assert.strictEqual(r.market, 'coupang');
  assert.deepStrictEqual(r.missing, []);
  assert.strictEqual(r.rows.length, 3);
  assert.strictEqual(r.rows[0].name, '홍길순');
  assert.strictEqual(r.rows[0].phone, '0502-3333-4444');
  assert.strictEqual(r.rows[0].item, '여수 돌산갓 2kg / 생갓');
  assert.strictEqual(r.rows[0].qty, 2);
  assert.strictEqual(r.rows[2].qty, 4);
});

test('헤더 위에 제목행이 30행 있어도 탐지', () => {
  const rows = [['발주발송관리']].concat(Array.from({ length: 29 }, () => ['안내'])).concat([
    ['수취인명', '상품명', '수량', '전화번호', '주소'],
    ['가', '나', '2', '1012345678', '서울'],
  ]);
  const r = P.convertSheet(rows);
  assert.strictEqual(r.headerRow, 31);
  assert.strictEqual(r.rows[0].phone, '010-1234-5678');
});

test('여러 시트 중 주문 시트 선택', () => {
  const r = P.convertSheets([
    { name: '안내', rows: [['이 파일은 주문 목록입니다']] },
    { name: '주문', rows: [['수취인명', '상품명', '수량', '전화번호', '주소'], ['가', '칼국수', '1', '01011112222', '서울']] },
  ]);
  assert.strictEqual(r.sheetName, '주문');
  assert.strictEqual(r.rows.length, 1);
});

test('전화번호 포맷 · 앞자리 0 복원', () => {
  assert.strictEqual(P.formatPhone('0212345678'), '02-1234-5678');
  assert.strictEqual(P.formatPhone('031-123-4567'), '031-123-4567');
  assert.strictEqual(P.formatPhone('15881234'), '1588-1234');
  assert.strictEqual(P.formatPhone(''), '');
  assert.strictEqual(P.formatPhone('50411112222'), '0504-1111-2222');
  assert.strictEqual(P.formatPhone('212345678'), '02-1234-5678');
  assert.strictEqual(P.formatPhone('311234567'), '031-123-4567');
  assert.strictEqual(P.formatPhone('7012345678'), '070-1234-5678');
  assert.strictEqual(P.normalizePhone('12345').valid, false);
});

test('취소/반품 상태 판정', () => {
  assert.ok(P.excludeReason(['결제완료', '취소요청']));
  assert.ok(P.excludeReason(['반품완료']));
  assert.strictEqual(P.excludeReason(['취소철회']), '');
  assert.strictEqual(P.excludeReason(['교환재배송']), '');
  assert.strictEqual(P.excludeReason(['결제완료', '']), '');
});

test('구매자 연락처 대신 수취인 휴대폰 사용', () => {
  const r = P.convertSheet([
    ['이름', '상품명', '수량', '주문자 연락처', '휴대폰', '주소'],
    ['가', '수제비', '1', '010-1111-1111', '010-2222-2222', '서울'],
  ]);
  assert.strictEqual(r.rows[0].phone, '010-2222-2222');
});

test('품목명 결합 — 옵션이 상품명 끝에 붙은 경우만 생략', () => {
  assert.strictEqual(P.buildItem('사과 1kg', '1'), '사과 1kg / 1');
  assert.strictEqual(P.buildItem('블루베리 1kg, 냉동', '냉동'), '블루베리 1kg, 냉동');
  assert.strictEqual(P.buildItem('칼국수', ''), '칼국수');
});

test('수량 누락/오류 표시', () => {
  const r = P.convertSheet([
    ['수취인명', '상품명', '수량', '전화번호', '주소'],
    ['가', '칼국수', '', '01011112222', '서울'],
    ['나', '칼국수', '두개', '01011112222', '서울'],
  ]);
  assert.strictEqual(r.rows[0].qty, 1);
  assert.ok(r.rows[0].issues.includes('qty-default'));
  assert.ok(r.rows[1].issues.includes('qty'));
});

test('품목 분류', () => {
  const cat = P.DEFAULT_CATALOG;
  assert.strictEqual(P.classify('[손칼국수] 쫄깃한 생칼국수 1kg', cat), '칼국수');
  assert.strictEqual(P.classify('감자 수제비 반죽', cat), '수제비');
  assert.strictEqual(P.classify('청도 한재 미나리 1kg', cat), '청도 미나리');
  assert.strictEqual(P.classify('여수돌산갓 2kg', cat), '여수 돌산갓');
  assert.strictEqual(P.classify('기장 쪽파 1단', cat), '기장 쪽파');
  assert.strictEqual(P.classify('옛날 구포국수', cat), '구포국수');
  assert.strictEqual(P.classify('오곡 곡물면', cat), '곡물면');
  assert.strictEqual(P.classify('오곡 곡물 국수', cat), '곡물면');
  assert.strictEqual(P.classify('국산 들기름', cat), '');
  assert.strictEqual(P.classify('국산 들기름', cat.concat([{ name: '들기름', keywords: ['들기름'] }])), '들기름');
});

test('품목 오분류 방지 — 지역명·흔한 채소명만으로는 분류하지 않음', () => {
  const cat = P.DEFAULT_CATALOG;
  for (const t of ['예산 사과 5kg', '구포시장 수제 어묵', '진도 쪽파', '돌미나리']) {
    assert.strictEqual(P.classify(t, cat), '', t);
  }
  assert.strictEqual(P.classify('구포 칼국수', cat), '칼국수');
  // 산지+품목 조합 키워드
  assert.strictEqual(P.classify('[청도] 미나리 1kg', cat), '청도 미나리');
  assert.strictEqual(P.classify('기장 햇쪽파', cat), '기장 쪽파');
  assert.strictEqual(P.classify('예산 옛날국수 3kg', cat), '예산국수');
  assert.strictEqual(P.classify('구포 소면', cat), '구포국수');
  // 여러 품목에 걸리면 확인 필요
  assert.deepStrictEqual(P.classifyText('칼국수+수제비 세트', cat).matches.sort(), ['수제비', '칼국수']);
  // 상품명 우선, 옵션은 보조
  assert.strictEqual(P.classifyProduct('생칼국수 1kg', '수제비 추가', cat).name, '칼국수');
  assert.strictEqual(P.classifyProduct('산지직송 채소', '기장쪽파 1kg', cat).name, '기장 쪽파');
  // 저장된 키워드가 문자열이어도 동작
  assert.strictEqual(P.classify('들기름', [{ name: '기름', keywords: '들기름, 참기름' }]), '기름');
});

test('이미 발송된 주문 제외, +82 번호 정규화', () => {
  assert.strictEqual(P.excludeReason(['배송완료']), '배송완료');
  assert.ok(P.isShippedStatus('구매확정'));
  assert.strictEqual(P.formatPhone('+82 10-1234-5678'), '010-1234-5678');
});

test('숫자로 저장된 긴 주문번호도 그대로 키로 사용', () => {
  const r = P.convertSheet([
    ['상품주문번호', '수취인명', '상품명', '수량', '수취인연락처1', '통합배송지'],
    [2026092812345601, '가', '칼국수', 1, '01011112222', '서울'],
    [2026092812345602, '나', '칼국수', 1, '01033334444', '부산'],
  ]);
  assert.notStrictEqual(r.rows[0].key, r.rows[1].key);
  assert.strictEqual(r.rows[0].keyKind, 'o');
});

test('헤더를 못 찾으면 오류', () => {
  assert.ok(P.convertSheet([['a', 'b'], [1, 2]]).error);
});
