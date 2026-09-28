// 실제 브라우저(Chromium)로 index.html 을 열어 업로드 → 품목 선택 → 다운로드까지 검증
const test = require('node:test');
const assert = require('node:assert');
const path = require('path');
const fs = require('fs');
const os = require('os');
const { chromium } = require('playwright');
const ExcelJS = require('exceljs');

const ROOT = path.join(__dirname, '..');
const url = 'file://' + path.join(ROOT, 'index.html');

// Playwright 는 한글 경로를 setInputFiles 로 넘기면 조용히 실패하므로 버퍼로 전달
function sample(name) {
  return {
    name,
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: fs.readFileSync(path.join(ROOT, 'samples', name)),
  };
}

async function withPage(fn) {
  // LANG 미설정 컨테이너에서는 크로미움이 한글 다운로드 파일명을 'download' 로 바꾸므로 UTF-8 로케일 지정
  const browser = await chromium.launch({ env: { ...process.env, LANG: 'C.UTF-8', LC_ALL: 'C.UTF-8' } });
  try {
    const page = await browser.newPage({ acceptDownloads: true });
    const errors = [];
    page.on('pageerror', (e) => errors.push(e.message));
    await page.goto(url);
    await fn(page);
    assert.deepStrictEqual(errors, []);
  } finally {
    await browser.close();
  }
}

const included = (page) => page.evaluate(() => window.__po.state.rows.filter((r) => r.included).length);
const chip = (page, name) => page.locator('#itemChips label.chip', { hasText: name }).locator('input[type=checkbox]');

async function upload(page, files) {
  const before = await page.evaluate(() => window.__po.state.files.length);
  await page.setInputFiles('#fileInput', files);
  return before;
}

test('업로드 → 품목 선택 → 당일발주 엑셀 다운로드', async () => {
  await withPage(async (page) => {
    await upload(page, [sample('스마트스토어_주문샘플.xlsx'), sample('쿠팡_주문샘플.xlsx')]);
    await page.waitForFunction(() => window.__po.state.files.length === 2 && !window.__po.state.busy);
    // 스마트스토어 5건 중 취소 1건 제외 + 쿠팡 3건 중 기타(들기름, 기본 해제) 1건 제외
    assert.strictEqual(await included(page), 6);
    assert.strictEqual(await page.locator('#previewTable tbody tr').count(), 6);
    // 빠진 주문 안내 (품목 미선택 → 경고)
    assert.match(await page.textContent('#excludeNotice'), /빠진 주문 2건/);
    assert.match(await page.textContent('#excludeNotice'), /체크하지 않은 품목: 기타 1건/);
    // 올린 뒤 업로드 영역은 한 줄로, 샘플 버튼 숨김
    assert.ok(await page.locator('#dropzone.compact').isVisible());
    assert.ok(await page.locator('#demoBtn').isHidden());

    // 칼국수 해제 → 5건, 기타 체크 → 6건
    await chip(page, '칼국수').uncheck();
    assert.strictEqual(await included(page), 5);
    await chip(page, '기타').check();
    assert.strictEqual(await included(page), 6);

    // 같은 파일 재업로드는 무시
    await upload(page, [sample('쿠팡_주문샘플.xlsx')]);
    await page.waitForSelector('#toast:not([hidden])');
    assert.match(await page.textContent('#toast'), /이미 올린 파일/);
    assert.strictEqual(await page.evaluate(() => window.__po.state.files.length), 2);

    const [download] = await Promise.all([page.waitForEvent('download'), page.click('#downloadBtn')]);
    assert.match(download.suggestedFilename(), /^당일발주_\d{8}_\d{4}\.xlsx$/);
    const out = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'po-')), 'out.xlsx');
    await download.saveAs(out);

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(out);
    const ws = wb.getWorksheet('당일발주');
    assert.ok(ws, '당일발주 시트 존재');
    assert.deepStrictEqual(ws.getRow(1).values.slice(1), ['수취자명', '구입품목 ', '수량', '전화번호', '주소', '배송메세지']);
    assert.strictEqual(ws.getCell('A1').fill.fgColor.argb, 'FF9BC2E6');
    assert.strictEqual(ws.getCell('A1').font.bold, true);
    assert.strictEqual(ws.rowCount, 7);
    const names = [2, 3, 4, 5, 6, 7].map((i) => ws.getCell('A' + i).value);
    assert.ok(!names.includes('김하늘'), '칼국수 주문 제외');
    assert.ok(names.includes('윤소희'), '기타(들기름) 주문 포함');
    assert.ok(!names.includes('정예린'), '취소 주문 제외');
    assert.strictEqual(ws.getCell('D2').value, '010-9876-5432');
  });
});

test('비밀번호 걸린 파일: 틀린 비번 → 재입력 → 열림, 같은 주문은 중복 제외', async () => {
  await withPage(async (page) => {
    await upload(page, [sample('스마트스토어_주문샘플.xlsx')]);
    await page.waitForFunction(() => window.__po.state.files.length === 1 && !window.__po.state.busy);
    await upload(page, [sample('스마트스토어_암호걸린샘플_비번1234.xlsx')]);
    await page.waitForSelector('#pwDialog[open]');
    await page.fill('#pwInput', '0000');
    await page.click('#pwOk');
    await page.waitForSelector('#pwError:not([hidden])');
    assert.match(await page.textContent('#pwError'), /맞지 않습니다/);
    await page.fill('#pwInput', '1234');
    await page.click('#pwOk');
    await page.waitForFunction(() => window.__po.state.files.length === 2 && !window.__po.state.busy, null, { timeout: 20000 });
    const f = await page.evaluate(() => window.__po.state.files[1]);
    assert.ok(!f.error, f.error);
    assert.strictEqual(f.count, 5);
    // 같은 상품주문번호 → 중복으로 제외되어 발주 건수는 그대로 4건
    assert.strictEqual(await included(page), 4);
  });
});

test('잘못된 파일은 친절한 안내, 키보드로 파일 선택 가능', async () => {
  await withPage(async (page) => {
    // 단계 버튼 3개 다음에 파일 선택 버튼
    for (let i = 0; i < 4; i++) await page.keyboard.press('Tab');
    const focused = await page.evaluate(() => document.activeElement.id);
    assert.strictEqual(focused, 'pickBtn');
    await upload(page, [{ name: 'photo.jpg', mimeType: 'image/jpeg', buffer: Buffer.from([0xff, 0xd8, 0xff, 0xe0, 1, 2, 3]) }]);
    await page.waitForSelector('#fileList li.error');
    assert.match(await page.textContent('#fileList li.error'), /엑셀 파일\(.xlsx, .xls, .csv\)이 아닙니다/);
  });
});

test('품목 추가 → 분류되어 포함, 삭제는 되돌리기 가능', async () => {
  await withPage(async (page) => {
    await page.click('#demoBtn');
    await page.click('#selectNoneBtn');
    assert.strictEqual(await included(page), 0);
    await page.click('#manageBtn');
    await page.fill('#newItemInput', '들기름');
    await page.click('#addItemForm button[type=submit]');
    assert.strictEqual(await included(page), 1);
    // 삭제 → 기타로 이동 알림 → 되돌리기
    await page.click('#manageList li:has-text("들기름") button');
    assert.strictEqual(await included(page), 0);
    assert.match(await page.textContent('#toastText'), /기타/);
    await page.click('#toastAction');
    assert.strictEqual(await included(page), 1);
    // 저장된 품목 목록은 새로고침 후에도 유지
    await page.reload();
    assert.ok(await chip(page, '들기름').isChecked());
    assert.ok(!(await chip(page, '수제비').isChecked()));
  });
});

test('샘플은 실제 파일을 올리면 자동으로 사라짐', async () => {
  await withPage(async (page) => {
    await page.click('#demoBtn');
    assert.strictEqual(await page.evaluate(() => window.__po.state.files[0].demo), true);
    await upload(page, [sample('쿠팡_주문샘플.xlsx')]);
    await page.waitForFunction(() => window.__po.state.files.length === 1 && !window.__po.state.files[0].demo && !window.__po.state.busy);
    assert.strictEqual(await page.evaluate(() => window.__po.state.rows.length), 3);
  });
});

test('다른 파일의 취소가 반영되고, 같은 파일 안의 같은 내용은 빼지 않고 경고만', async () => {
  const XLSX = require('xlsx');
  const mk = (rows) => {
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), 'S');
    return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  };
  const H = ['상품주문번호', '주문상태', '수취인명', '상품명', '수량', '수취인연락처1', '통합배송지'];
  const am = mk([H, [2026092800000001, '결제완료', '가', '칼국수', 1, '01011112222', '서울'], [2026092800000002, '결제완료', '나', '수제비', 1, '01033334444', '부산']]);
  const pm = mk([H, [2026092800000001, '취소요청', '가', '칼국수', 1, '01011112222', '서울']]);
  const etc = mk([['수취인명', '상품명', '수량', '전화번호', '주소'], ['다', '칼국수', 1, '01055556666', '대구'], ['다', '칼국수', 1, '01055556666', '대구']]);
  const t = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  await withPage(async (page) => {
    await upload(page, [{ name: 'am.xlsx', mimeType: t, buffer: am }, { name: 'pm.xlsx', mimeType: t, buffer: pm }, { name: 'etc.xlsx', mimeType: t, buffer: etc }]);
    await page.waitForFunction(() => window.__po.state.files.length === 3 && !window.__po.state.busy);
    const rows = await page.evaluate(() => window.__po.state.rows.map((r) => ({ n: r.name, inc: r.included, w: r.warns })));
    assert.deepStrictEqual(rows.map((r) => r.inc), [false, true, false, true, true]);
    assert.ok(rows[4].w.includes('dup-suspect'));
  });
});

test('표 방향키 이동 · Enter 로 편집 · 되돌리기', async () => {
  await withPage(async (page) => {
    await page.click('#demoBtn');
    await page.focus('#previewTable tbody td[tabindex="0"]');
    assert.strictEqual(await page.evaluate(() => document.activeElement.dataset.col), 'name');
    await page.keyboard.press('ArrowRight');
    await page.keyboard.press('ArrowDown');
    assert.strictEqual(await page.evaluate(() => document.activeElement.dataset.col), 'item');
    await page.keyboard.press('ArrowRight');
    await page.keyboard.press('Enter');
    await page.keyboard.type('5');
    await page.keyboard.press('Enter');
    assert.strictEqual(await page.evaluate(() => window.__po.state.rows[1].qty), 5);
    assert.ok(await page.locator('#previewTable tbody tr').nth(1).locator('td.edited').count());
    await page.locator('#previewTable tbody tr').nth(1).locator('.undo').click();
    assert.strictEqual(await page.evaluate(() => window.__po.state.rows[1].qty), 3);
  });
});
