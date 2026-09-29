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
// 품목 입력칸에 적고 Enter
async function typeTerms(page, text) {
  await page.fill('#termInput', text);
  await page.press('#termInput', 'Enter');
}

async function upload(page, files) {
  const before = await page.evaluate(() => window.__po.state.files.length);
  await page.setInputFiles('#fileInput', files);
  return before;
}

test('업로드 → 품목 입력 → 당일발주 엑셀 다운로드', async () => {
  await withPage(async (page) => {
    await upload(page, [sample('스마트스토어_주문샘플.xlsx'), sample('쿠팡_주문샘플.xlsx')]);
    await page.waitForFunction(() => window.__po.state.files.length === 2 && !window.__po.state.busy);
    // 품목을 비워 두면 전체 주문 (스마트스토어 5건 중 취소 1건 제외 + 쿠팡 3건)
    assert.strictEqual(await included(page), 7);
    assert.match(await page.textContent('#termResults'), /전체 주문 7건/);
    // 올린 뒤 업로드 영역은 한 줄로
    assert.ok(await page.locator('#dropzone.compact').isVisible());

    // 칼국수·들기름을 빼고 적기
    await typeTerms(page, '수제비, 구포국수, 예산국수, 기장 쪽파, 여수 돌산갓, 청도 미나리');
    assert.strictEqual(await included(page), 5);
    assert.strictEqual(await page.locator('#previewTable tbody tr').count(), 5);
    const results = await page.textContent('#termResults');
    assert.match(results, /수제비 1건/);
    assert.match(results, /청도 미나리 0건/); // 미나리 주문은 취소됨
    assert.match(results, /해당 없음 2건/);
    assert.match(await page.textContent('#excludeNotice'), /아닌 주문 2건/);

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
    assert.strictEqual(ws.rowCount, 6);
    const names = [2, 3, 4, 5, 6].map((i) => ws.getCell('A' + i).value);
    assert.ok(!names.includes('김하늘'), '칼국수 주문 제외');
    assert.ok(!names.includes('윤소희'), '들기름 주문 제외');
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
    // 데이터가 없을 때 단계 버튼은 Tab 순서에서 빠지므로 첫 Tab 이 파일 선택 버튼
    await page.keyboard.press('Tab');
    const focused = await page.evaluate(() => document.activeElement.id);
    assert.strictEqual(focused, 'pickBtn');
    await upload(page, [{ name: 'photo.jpg', mimeType: 'image/jpeg', buffer: Buffer.from([0xff, 0xd8, 0xff, 0xe0, 1, 2, 3]) }]);
    await page.waitForSelector('#fileList li.error');
    assert.match(await page.textContent('#fileList li.error'), /엑셀 파일\(.xlsx, .xls, .csv\)이 아닙니다/);
  });
});

test('품목 입력은 바로 반영되고, 다음에 열어도 남아 있음', async () => {
  await withPage(async (page) => {
    await page.click('#demoBtnBig');
    // 타이핑만 해도(Enter 없이) 잠시 뒤 반영
    await page.fill('#termInput', '들기름');
    await page.waitForFunction(() => window.__po.state.rows.filter((r) => r.included).length === 1);
    assert.match(await page.textContent('#termResults'), /들기름 1건/);
    await typeTerms(page, '들기름, 없는품목');
    assert.ok(await page.locator('#termResults .term-chip.zero', { hasText: '없는품목' }).count());
    await page.reload();
    assert.strictEqual(await page.inputValue('#termInput'), '들기름, 없는품목');
    // 지우기 → 전체
    await page.click('#demoBtnBig');
    await page.click('#termClear');
    assert.strictEqual(await included(page), 6);
  });
});
test('샘플은 실제 파일을 올리면 자동으로 사라짐', async () => {
  await withPage(async (page) => {
    await page.click('#demoBtnBig');
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
    await page.click('#demoBtnBig');
    await page.focus('#previewTable tbody td[tabindex="0"]');
    assert.strictEqual(await page.evaluate(() => document.activeElement.dataset.col), 'name');
    await page.keyboard.press('ArrowRight');
    await page.keyboard.press('ArrowDown');
    assert.strictEqual(await page.evaluate(() => document.activeElement.dataset.col), 'item');
    await page.keyboard.press('ArrowRight'); // 품목
    await page.keyboard.press('ArrowRight'); // 수량
    await page.keyboard.press('Enter');
    await page.keyboard.type('5');
    await page.keyboard.press('Enter');
    assert.strictEqual(await page.evaluate(() => window.__po.state.rows[1].qty), 5);
    assert.ok(await page.locator('#previewTable tbody tr').nth(1).locator('td.edited').count());
    await page.locator('#previewTable tbody tr').nth(1).locator('.undo').click();
    assert.strictEqual(await page.evaluate(() => window.__po.state.rows[1].qty), 3);
  });
});

const XLSXN = require('xlsx');
const T = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const mkXlsx = (rows) => {
  const wb = XLSXN.utils.book_new();
  XLSXN.utils.book_append_sheet(wb, XLSXN.utils.aoa_to_sheet(rows), 'S');
  return XLSXN.write(wb, { type: 'buffer', bookType: 'xlsx' });
};

test('품목 메뉴: 키보드로 바꾸고, 해당 없음으로 바꾸면 빠진다는 안내', async () => {
  await withPage(async (page) => {
    await page.click('#demoBtnBig');
    await typeTerms(page, '칼국수, 수제비');
    await page.focus('#previewTable tbody td[tabindex="0"]');
    await page.keyboard.press('ArrowRight');
    await page.keyboard.press('ArrowRight');
    assert.strictEqual(await page.evaluate(() => document.activeElement.dataset.col), 'cat');
    await page.keyboard.press('Enter');
    await page.waitForSelector('.cat-menu');
    // 맨 아래 '해당 없음'으로 이동
    await page.keyboard.press('End');
    await page.keyboard.press('Enter');
    const r = await page.evaluate(() => window.__po.state.rows[0]);
    assert.strictEqual(r.categoryManual, '');
    assert.strictEqual(r.included, false);
    assert.match(await page.textContent('#toastText'), /빠집니다/);
    assert.strictEqual(await page.evaluate(() => document.activeElement.dataset.col), 'cat');
  });
});

test('편집 중 Tab 은 저장 후 옆 칸으로', async () => {
  await withPage(async (page) => {
    await page.click('#demoBtnBig');
    await page.click('#previewTable tbody tr:first-child td[data-col="name"]');
    await page.keyboard.type('님');
    await page.keyboard.press('Tab');
    assert.strictEqual(await page.evaluate(() => window.__po.state.rows[0].name), '김하늘님');
    assert.strictEqual(await page.evaluate(() => document.activeElement.dataset.col), 'item');
  });
});

test('입력에서 지운 품목으로 직접 지정한 주문은 자동 분류로, 다시 적으면 복구', async () => {
  await withPage(async (page) => {
    await page.click('#demoBtnBig');
    await typeTerms(page, '칼국수, 수제비');
    // 들기름 주문을 칼국수로 직접 지정
    await page.click('#previewTable tbody tr:has-text("한지우") .cat-tag').catch(() => {});
    await page.evaluate(() => { const r = window.__po.state.rows.find((x) => x.name === '한지우'); r.categoryManual = '칼국수'; });
    await typeTerms(page, '칼국수, 수제비');
    assert.strictEqual(await page.evaluate(() => window.__po.state.rows.find((x) => x.name === '한지우').included), true);
    await typeTerms(page, '수제비');
    const r = await page.evaluate(() => window.__po.state.rows.find((x) => x.name === '한지우'));
    assert.strictEqual(r.category, '');
    assert.strictEqual(r.included, false);
    await typeTerms(page, '칼국수, 수제비');
    assert.strictEqual(await page.evaluate(() => window.__po.state.rows.find((x) => x.name === '한지우').category), '칼국수');
  });
});
test('주문번호 없는 파일의 사본은 파일째 중복으로 빠지고, 다운로드 전 점검창이 뜬다', async () => {
  const rows = [['수취인명', '상품명', '수량', '전화번호', '주소'], ['가', '칼국수', 1, '01011112222', '서울'], ['나', '수제비', 2, '01033334444', '부산']];
  const buf = mkXlsx(rows);
  const buf2 = mkXlsx(rows.concat([['다', '국산 들기름', 1, '01055556666', '대구']]));
  await withPage(async (page) => {
    await typeTerms(page, '칼국수, 수제비');
    await upload(page, [{ name: 'a.xlsx', mimeType: T, buffer: buf }, { name: 'a 사본.xlsx', mimeType: T, buffer: buf }]);
    await page.waitForFunction(() => window.__po.state.files.length === 2 && !window.__po.state.busy);
    assert.strictEqual(await included(page), 2);
    await upload(page, [{ name: 'b.xlsx', mimeType: T, buffer: buf2 }]);
    await page.waitForFunction(() => window.__po.state.files.length === 3 && !window.__po.state.busy);
    // b 는 전부 같지는 않으므로 빼지 않고 경고만 → 가·나 가 한 번 더 (4건)
    assert.strictEqual(await included(page), 4);
    await page.click('#downloadBtn');
    await page.waitForSelector('#confirmDialog[open]');
    const text = await page.textContent('#confirmBody');
    assert.match(text, /같은 내용이 또 있는 주문 2건/);
    assert.match(text, /입력한 품목이 아닌 주문 1건/);
    await page.click('#confirmCancel');
    // 취소하면 '확인 필요' 보기로
    await page.waitForFunction(() => window.__po.state.view === 'check');
  });
});

test('HTML 표로 된 .xls 와 UTF-16 텍스트도 읽음', async () => {
  const html = '<html><body><table><tr><td>수취인명</td><td>상품명</td><td>수량</td><td>전화번호</td><td>주소</td></tr>' +
    '<tr><td>가</td><td>칼국수</td><td>1</td><td>010-1111-2222</td><td>서울</td></tr></table></body></html>';
  const tsv = '수취인명\t상품명\t수량\t전화번호\t주소\r\n나\t수제비\t2\t010-3333-4444\t부산\r\n';
  const utf16 = Buffer.concat([Buffer.from([0xff, 0xfe]), Buffer.from(tsv, 'utf16le')]);
  await withPage(async (page) => {
    await upload(page, [
      { name: 'order.xls', mimeType: 'application/vnd.ms-excel', buffer: Buffer.from(html) },
      { name: 'order.txt', mimeType: 'text/plain', buffer: utf16 },
    ]);
    await page.waitForFunction(() => window.__po.state.files.length === 2 && !window.__po.state.busy);
    const files = await page.evaluate(() => window.__po.state.files.map((f) => f.error || f.count));
    assert.deepStrictEqual(files, [1, 1]);
  });
});

test('3,000건도 빠르게 (처음 200건만 그림)', async () => {
  const H = ['상품주문번호', '수취인명', '상품명', '수량', '수취인연락처1', '통합배송지'];
  const rows = [H];
  for (let i = 0; i < 3000; i++) rows.push([String(2026092800000000 + i), '고객' + i, i % 2 ? '생칼국수 1kg' : '감자 수제비', 1, '010' + String(10000000 + i), '서울시 어딘가 ' + i]);
  await withPage(async (page) => {
    const t0 = Date.now();
    await upload(page, [{ name: 'big.xlsx', mimeType: T, buffer: mkXlsx(rows) }]);
    await page.waitForFunction(() => window.__po.state.files.length === 1 && !window.__po.state.busy, null, { timeout: 20000 });
    const loadMs = Date.now() - t0;
    assert.strictEqual(await included(page), 3000);
    assert.strictEqual(await page.locator('#previewTable tbody tr[data-uid]').count(), 200);
    const t1 = Date.now();
    await typeTerms(page, '칼국수');
    assert.strictEqual(await included(page), 1500);
    await typeTerms(page, '');
    const toggleMs = Date.now() - t1;
    assert.ok(loadMs < 8000, 'load ' + loadMs);
    assert.ok(toggleMs < 2500, 'toggle ' + toggleMs);
    await page.click('.more-row button');
    assert.strictEqual(await page.locator('#previewTable tbody tr[data-uid]').count(), 400);
  });
});

test('샘플 2개 파일은 점검창 없이 바로 다운로드', async () => {
  await withPage(async (page) => {
    await upload(page, [sample('스마트스토어_주문샘플.xlsx'), sample('쿠팡_주문샘플.xlsx')]);
    await page.waitForFunction(() => window.__po.state.files.length === 2 && !window.__po.state.busy);
    const [download] = await Promise.all([page.waitForEvent('download', { timeout: 5000 }), page.click('#downloadBtn')]);
    assert.ok(download);
    assert.ok(!(await page.locator('#confirmDialog[open]').count()));
    // 확인 필요가 없으면 초록 표시
    assert.match(await page.textContent('#tab-check'), /확인 필요 없음/);
  });
});

test('품목 메뉴는 표 밖에 떠서 잘리지 않음', async () => {
  await withPage(async (page) => {
    await page.click('#demoBtnBig');
    await typeTerms(page, '칼국수, 수제비, 구포국수, 예산국수, 기장 쪽파, 여수 돌산갓, 청도 미나리, 곡물면');
    await page.click('#previewTable tbody tr:last-child .cat-tag');
    const box = await page.evaluate(() => {
      const m = document.querySelector('.cat-menu');
      const r = m.getBoundingClientRect();
      return { parentIsBody: m.parentElement === document.body, h: r.height, top: r.top, bottom: r.bottom, vh: innerHeight, full: m.scrollHeight <= m.clientHeight + 1 };
    });
    assert.ok(box.parentIsBody);
    assert.ok(box.top >= 0 && box.bottom <= box.vh, JSON.stringify(box));
    assert.ok(box.full, '메뉴 항목이 모두 보임 ' + JSON.stringify(box));
  });
});

test('적지 않은 품목은 여러 품목에 걸려도 섞이지 않음', async () => {
  const rows = [['수취인명', '상품명', '수량', '전화번호', '주소'],
    ['가', '구포시장 칼국수', 1, '01011112222', '서울'],
    ['나', '감자 수제비', 1, '01033334444', '부산']];
  await withPage(async (page) => {
    await upload(page, [{ name: 's.xlsx', mimeType: T, buffer: mkXlsx(rows) }]);
    await page.waitForFunction(() => window.__po.state.files.length === 1 && !window.__po.state.busy);
    await typeTerms(page, '수제비, 구포 국수');
    const r = await page.evaluate(() => window.__po.state.rows.map((x) => [x.category, x.included]));
    // '구포 국수'는 두 단어가 모두 있으면 해당 → '구포시장 칼국수'도 걸림(사용자가 적은 대로)
    assert.deepStrictEqual(r, [['구포 국수', true], ['수제비', true]]);
    await typeTerms(page, '수제비');
    assert.deepStrictEqual(await page.evaluate(() => window.__po.state.rows.map((x) => x.included)), [false, true]);
  });
});
test('탭은 방향키로 이동, 빈 필수 칸에는 안내 문구', async () => {
  const rows = [['수취인명', '상품명', '수량', '전화번호', '주소'], ['가', '칼국수', 1, '01011112222', '']];
  await withPage(async (page) => {
    await upload(page, [{ name: 'n.xlsx', mimeType: T, buffer: mkXlsx(rows) }]);
    await page.waitForFunction(() => window.__po.state.files.length === 1 && !window.__po.state.busy);
    const hint = await page.evaluate(() => {
      const td = document.querySelector('td[data-field="address"]');
      return td.dataset.hint + '|' + getComputedStyle(td, '::before').content;
    });
    assert.match(hint, /주소 입력 필요\|"주소 입력 필요"/);
    assert.match(await page.textContent('.c-cat .row-reason'), /주소 없음/);
    await page.focus('#tab-included');
    await page.keyboard.press('ArrowRight');
    assert.strictEqual(await page.evaluate(() => [window.__po.state.view, document.activeElement.id].join()), 'check,tab-check');
  });
});

test('여러 날 중복: 어제 받은 주문이 오늘 다시 오면 기본으로 빠짐 (기록 지우기 가능)', async () => {
  const browser = await chromium.launch({ env: { ...process.env, LANG: 'C.UTF-8', LC_ALL: 'C.UTF-8' } });
  try {
    const ctx = await browser.newContext({ acceptDownloads: true });
    const page = await ctx.newPage();
    await page.goto(url);
    await upload(page, [sample('스마트스토어_주문샘플.xlsx')]);
    await page.waitForFunction(() => window.__po.state.files.length === 1 && !window.__po.state.busy);
    const [d] = await Promise.all([page.waitForEvent('download'), page.click('#downloadBtn')]);
    assert.match(d.suggestedFilename(), /^당일발주_/);
    // 완료 상태: 다음 할 일 + 다시 받기/새로 시작
    assert.ok(await page.locator('#doneBox').isVisible());
    assert.strictEqual((await page.textContent('#downloadBtn')).trim(), '다시 받기');
    const hist = await page.evaluate(() => localStorage.getItem('po.history.v1'));
    assert.ok(hist && !hist.includes('김하늘') && !hist.includes('2026092812345'), '해시만 저장');
    // 고치면 '받은 뒤 내용이 바뀌었습니다'
    await page.click('#previewTable tbody tr:first-child td[data-col="memo"]');
    await page.keyboard.type('!');
    await page.keyboard.press('Enter');
    assert.match(await page.textContent('#actionText'), /바뀌었습니다/);

    // 다음 날(새로 연 페이지)에 같은 파일 → 전부 '지난 발주에 있음'
    const page2 = await ctx.newPage();
    await page2.goto(url);
    await upload(page2, [sample('스마트스토어_주문샘플.xlsx')]);
    await page2.waitForFunction(() => window.__po.state.files.length === 1 && !window.__po.state.busy);
    assert.strictEqual(await included(page2), 0);
    assert.match(await page2.textContent('#excludeNotice'), /지난 발주에 이미 넣은 주문 4건/);
    // 기록 지우기
    await page2.click('#clearHistoryBtn');
    await page2.click('#confirmOk');
    await page2.waitForFunction(() => window.__po.state.rows.filter((r) => r.included).length === 4);

    // 고친 뒤 다시 받기 → 완료 상태 → '새로 시작' 후 같은 파일을 다시 올려도 걸러짐
    await Promise.all([page.waitForEvent('download'), page.click('#downloadBtn')]);
    await page.click('#newStartBtn');
    await upload(page, [sample('쿠팡_주문샘플.xlsx'), sample('스마트스토어_주문샘플.xlsx')]);
    await page.waitForFunction(() => window.__po.state.files.length === 2 && !window.__po.state.busy);
    const ss = await page.evaluate(() => window.__po.state.rows.filter((r) => r.source === '스마트스토어' && r.included).length);
    assert.strictEqual(ss, 0);
  } finally {
    await browser.close();
  }
});

test('샘플 상태에서는 파일 이름에 샘플_ 이 붙고 기록에 남지 않음', async () => {
  await withPage(async (page) => {
    await page.click('#demoBtnBig');
    assert.ok(await page.locator('#demoBand').isVisible());
    assert.strictEqual((await page.textContent('#downloadBtn')).trim(), '샘플 발주서 받아보기');
    // 샘플에는 주소가 빠진 주문이 있어 점검창이 뜬다 → 그대로 다운로드
    await page.click('#downloadBtn');
    await page.waitForSelector('#confirmDialog[open]');
    const [d] = await Promise.all([page.waitForEvent('download'), page.click('#confirmOk')]);
    assert.match(d.suggestedFilename(), /^샘플_당일발주_/);
    assert.strictEqual(await page.evaluate(() => localStorage.getItem('po.history.v1')), null);
    assert.ok(await page.locator('#doneBox').isHidden());
  });
});

test('세트 상품은 조합 품목으로, 짧은 전화번호는 차단', async () => {
  const rows = [['수취인명', '상품명', '수량', '전화번호', '주소'],
    ['가', '칼국수+수제비 세트', 1, '01011112222', '서울'],
    ['나', '생칼국수', 1, '12345', '부산']];
  await withPage(async (page) => {
    await upload(page, [{ name: 'set.xlsx', mimeType: T, buffer: mkXlsx(rows) }]);
    await page.waitForFunction(() => window.__po.state.files.length === 1 && !window.__po.state.busy);
    await typeTerms(page, '칼국수, 수제비');
    // 세트는 '칼국수+수제비' 로 분류되어 확인 필요가 아님 → 짧은 전화번호 1건만
    assert.match(await page.textContent('#tab-check'), /1/);
    assert.strictEqual(await page.evaluate(() => window.__po.state.rows[0].category), '칼국수+수제비');
    await page.click('#downloadBtn');
    await page.waitForSelector('#confirmDialog[open]');
    const t = await page.textContent('#confirmBody');
    assert.doesNotMatch(t, /여러 품목/);
    assert.match(t, /전화번호가 너무 짧음/);
  });
});

test('열린 품목 메뉴의 태그를 다시 누르면 닫힘', async () => {
  await withPage(async (page) => {
    await page.click('#demoBtnBig');
    await typeTerms(page, '칼국수, 수제비');
    const tag = page.locator('#previewTable tbody tr:first-child .cat-tag');
    await tag.click();
    await page.waitForSelector('.cat-menu');
    await tag.click();
    assert.strictEqual(await page.locator('.cat-menu').count(), 0);
  });
});

test('여러 품목이 함께 든 주문은 수제비+칼국수 로 표시되고 메뉴에서도 고를 수 있음', async () => {
  const rows = [['수취인명', '상품명', '수량', '전화번호', '주소'],
    ['가', '칼국수+수제비 세트', 1, '01011112222', '서울'],
    ['나', '생칼국수 1kg', 1, '01033334444', '부산']];
  await withPage(async (page) => {
    await upload(page, [{ name: 'combo.xlsx', mimeType: T, buffer: mkXlsx(rows) }]);
    await page.waitForFunction(() => window.__po.state.files.length === 1 && !window.__po.state.busy);
    await typeTerms(page, '수제비, 칼국수');
    const r = await page.evaluate(() => window.__po.state.rows.map((x) => [x.category, x.included, x.warns.includes('ambiguous')]));
    assert.deepStrictEqual(r, [['수제비+칼국수', true, false], ['칼국수', true, false]]);
    assert.match(await page.textContent('#termResults'), /수제비\+칼국수 1건/);
    assert.match(await page.textContent('#tab-check'), /확인 필요 없음/);
    // 칼국수 주문을 메뉴에서 '수제비+칼국수'로 바꾸기
    await page.click('#previewTable tbody tr:nth-child(2) .cat-tag');
    await page.click('.cat-menu .cat-opt:has-text("함께 주문")');
    assert.strictEqual(await page.evaluate(() => window.__po.state.rows[1].category), '수제비+칼국수');
    // 수제비만 적어도 조합 주문은 포함
    await typeTerms(page, '수제비');
    assert.deepStrictEqual(await page.evaluate(() => window.__po.state.rows.map((x) => x.included)), [true, true]);
  });
});

test('④ 추가 데이터: 이름을 적으면 이름·연락처·주소가 당일발주 시트에 줄로 들어감', async () => {
  const browser = await chromium.launch({ env: { ...process.env, LANG: 'C.UTF-8', LC_ALL: 'C.UTF-8' } });
  try {
    const ctx = await browser.newContext({ acceptDownloads: true });
    const page = await ctx.newPage();
    const errors = [];
    page.on('pageerror', (e) => errors.push(e.message));
    await page.goto(url);
    await upload(page, [sample('스마트스토어_주문샘플.xlsx'), sample('쿠팡_주문샘플.xlsx')]);
    await page.waitForFunction(() => window.__po.state.files.length === 2 && !window.__po.state.busy);
    await page.setInputFiles('#dataInput', [sample('단골고객_샘플.xlsx')]);
    await page.waitForSelector('#dataBody:not([hidden])');
    assert.match(await page.textContent('#dataCols'), /이름 ← ‘고객명’ 열 · 연락처 ← ‘연락처’ 열 · 주소 ← ‘주소’ 열/);
    // 이름 적기 전에는 발주서에 아무것도 추가되지 않음
    assert.strictEqual(await included(page), 7);
    await page.fill('#dataQuery', '문지호, 송하윤, 박민수');
    await page.press('#dataQuery', 'Enter');
    const sum = await page.textContent('#dataSummary');
    assert.match(sum, /2명 찾음/);
    assert.match(sum, /‘박민수’은\(는\) 찾지 못했습니다/);
    // 발주서 표 맨 아래에 두 줄 추가 (품목 필터와 상관없이 포함)
    assert.strictEqual(await included(page), 9);
    const extra = await page.evaluate(() => window.__po.state.rows.filter((r) => r.extra).map((r) => [r.name, r.phone, r.address, r.item]));
    assert.deepStrictEqual(extra, [
      ['송하윤', '010-8080-9090', '서울특별시 마포구 월드컵북로 400', ''],
      ['문지호', '010-6060-7070', '부산광역시 수영구 광안해변로 219', ''],
    ]);
    assert.match(await page.textContent('#actionText'), /추가 데이터 2건 포함/);
    // 표에서 구입품목 적기
    const last = page.locator('#previewTable tbody tr[data-uid]').last();
    await last.locator('td[data-field="item"]').click();
    await page.keyboard.type('수제비 2봉');
    await page.keyboard.press('Enter');
    // 다운로드: 품목이 빈 줄이 하나 남아 점검창 → 그대로 다운로드
    await page.click('#downloadBtn');
    await page.waitForSelector('#confirmDialog[open]');
    assert.match(await page.textContent('#confirmBody'), /송하윤 — 품목 없음/);
    const [d] = await Promise.all([page.waitForEvent('download'), page.click('#confirmOk')]);
    const out = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'po-')), 'po.xlsx');
    await d.saveAs(out);
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(out);
    assert.deepStrictEqual(wb.worksheets.map((w) => w.name), ['당일발주']);
    const ws = wb.getWorksheet('당일발주');
    assert.strictEqual(ws.rowCount, 10);
    assert.deepStrictEqual(ws.getRow(9).values.slice(1, 6), ['송하윤', '', 1, '010-8080-9090', '서울특별시 마포구 월드컵북로 400']);
    assert.deepStrictEqual(ws.getRow(10).values.slice(1, 6), ['문지호', '수제비 2봉', 1, '010-6060-7070', '부산광역시 수영구 광안해변로 219']);
    // 끄면 빠짐
    await page.uncheck('#dataWithOrder');
    assert.strictEqual(await included(page), 7);
    // 다시 열면 올린 파일은 남고, 적은 이름은 비어 있음 (다음 날 같은 고객이 또 들어가지 않게)
    const page2 = await ctx.newPage();
    await page2.goto(url);
    await page2.waitForSelector('#dataBody:not([hidden])');
    assert.strictEqual(await page2.inputValue('#dataQuery'), '');
    assert.deepStrictEqual(errors, []);
  } finally {
    await browser.close();
  }
});
