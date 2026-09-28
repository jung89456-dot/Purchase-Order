// 버전별 화면 캡처: node scripts/screenshot.js v1
// docs/screenshots/<버전>-empty.png, <버전>-result.png, <버전>-mobile.png 생성
const path = require('path');
const fs = require('fs');
const { chromium } = require('playwright');

const ROOT = path.join(__dirname, '..');
const version = process.argv[2] || 'dev';
const outDir = path.join(ROOT, 'docs', 'screenshots');
const samples = ['스마트스토어_주문샘플.xlsx', '쿠팡_주문샘플.xlsx'].map((name) => ({
  name,
  mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  buffer: fs.readFileSync(path.join(ROOT, 'samples', name)),
}));

(async () => {
  fs.mkdirSync(outDir, { recursive: true });
  const browser = await chromium.launch({ env: { ...process.env, LANG: 'C.UTF-8', LC_ALL: 'C.UTF-8' } });
  const url = 'file://' + path.join(ROOT, 'index.html');
  try {
    for (const [suffix, viewport] of [['', { width: 1280, height: 800 }], ['-mobile', { width: 390, height: 844 }]]) {
      const page = await browser.newPage({ viewport, deviceScaleFactor: 1, acceptDownloads: true });
      await page.goto(url);
      if (!suffix) await page.screenshot({ path: path.join(outDir, `${version}-empty.png`), fullPage: true });
      await page.setInputFiles('#fileInput', samples);
      await page.waitForFunction(() => !window.__po || (window.__po.state.files.length === 2 && !window.__po.state.busy));
      await page.waitForTimeout(300);
      await page.screenshot({ path: path.join(outDir, `${version}${suffix || '-result'}.png`), fullPage: true });
      // 다운로드 완료 화면 (v6 부터)
      if (!suffix && (await page.locator('#doneBox').count())) {
        await Promise.all([page.waitForEvent('download'), page.click('#downloadBtn')]);
        await page.waitForTimeout(300);
        await page.screenshot({ path: path.join(outDir, `${version}-done.png`) });
      }
      await page.close();
    }
  } finally {
    await browser.close();
  }
  console.log('screenshots saved for', version);
})();
