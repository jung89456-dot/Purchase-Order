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
  assert.strictEqual(r.rows.length, 4); // 취소요청 1건 제외
  assert.strictEqual(r.skipped.length, 1);
  const first = r.rows[0];
  assert.strictEqual(first.name, '김하늘');
  assert.strictEqual(first.item, '국내산 햇사과 5kg / 크기: 중과(18~20과)');
  assert.strictEqual(first.qty, 1);
  assert.strictEqual(first.phone, '010-1234-5678');
  assert.strictEqual(first.address, '서울특별시 강남구 테헤란로 123 4층 401호');
  assert.strictEqual(first.memo, '부재 시 문 앞에 놓아주세요');
  assert.strictEqual(r.rows[1].phone, '010-9876-5432');
  assert.strictEqual(r.rows[2].phone, '0504-1111-2222');
  assert.strictEqual(r.rows[3].item, '유기농 고구마 10kg');
});

test('쿠팡 샘플 변환 — 구매자가 아닌 수취인 정보 사용', () => {
  const r = P.convertSheet(sheetRows('쿠팡_주문샘플.xlsx'));
  assert.strictEqual(r.market, 'coupang');
  assert.deepStrictEqual(r.missing, []);
  assert.strictEqual(r.rows.length, 3);
  assert.strictEqual(r.rows[0].name, '홍길순');
  assert.strictEqual(r.rows[0].phone, '0502-3333-4444');
  assert.strictEqual(r.rows[0].item, '무농약 블루베리 1kg / 냉동, 1kg');
  assert.strictEqual(r.rows[0].qty, 2);
  assert.strictEqual(r.rows[2].qty, 4);
});

test('헤더 위에 제목행이 있어도 탐지', () => {
  const rows = [['발주발송관리'], [], ['수취인명', '상품명', '수량', '전화번호', '주소'], ['가', '나', '2', '1012345678', '서울']];
  const r = P.convertSheet(rows);
  assert.strictEqual(r.headerRow, 3);
  assert.strictEqual(r.rows[0].phone, '010-1234-5678');
});

test('전화번호 포맷', () => {
  assert.strictEqual(P.formatPhone('0212345678'), '02-1234-5678');
  assert.strictEqual(P.formatPhone('031-123-4567'), '031-123-4567');
  assert.strictEqual(P.formatPhone('15881234'), '1588-1234');
  assert.strictEqual(P.formatPhone(''), '');
});

test('헤더를 못 찾으면 오류', () => {
  assert.ok(P.convertSheet([['a', 'b'], [1, 2]]).error);
});
