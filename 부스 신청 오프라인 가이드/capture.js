const puppeteer = require('puppeteer-core');
const path = require('path');
const fs = require('fs');

const CHROME = 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe';
const URL = 'https://www.nc26-bnikorea.com/ticket.html';
const OUT = path.join(__dirname, 'screenshots');
const sleep = ms => new Promise(r => setTimeout(r, ms));

(async () => {
  fs.mkdirSync(OUT, { recursive: true });

  const browser = await puppeteer.launch({
    executablePath: CHROME,
    headless: 'new',
    defaultViewport: { width: 1366, height: 900, deviceScaleFactor: 2 },
    args: ['--lang=ko-KR']
  });

  const page = await browser.newPage();
  await page.setExtraHTTPHeaders({ 'Accept-Language': 'ko-KR,ko;q=0.9' });

  console.log('1. Goto ticket.html');
  await page.goto(URL, { waitUntil: 'networkidle2', timeout: 60000 });
  await sleep(2000);

  // Step 1: 부스 전시 카드 섹션 스크린샷
  console.log('2. Capture booth section');
  await page.evaluate(() => {
    const h = [...document.querySelectorAll('h3')].find(el => el.textContent.includes('부스 전시'));
    if (h) h.scrollIntoView({ block: 'start' });
    window.scrollBy(0, -80);
  });
  await sleep(1200);
  await page.screenshot({ path: path.join(OUT, '01-booth-section.png'), fullPage: false });

  // Step 2: 부스 선택하기 버튼 클릭
  console.log('3. Click 부스 선택하기');
  await page.evaluate(() => {
    const btn = [...document.querySelectorAll('button')].find(b => b.textContent.includes('부스 선택하기'));
    if (btn) btn.click();
  });
  // 부스 상태(sold/admin) 로드 대기
  await sleep(3500);
  await page.screenshot({ path: path.join(OUT, '02-booth-map.png'), fullPage: false });

  // Step 3: 첫 번째 가능한 부스 선택
  console.log('4. Select first available booth');
  const selectedId = await page.evaluate(() => {
    const all = [...document.querySelectorAll('#booth-map-modal button[data-booth]')];
    const desktopBtns = all.filter(b => {
      let p = b.parentElement;
      while (p) { if (p.id === 'zone-a-mobile-row1' || p.id === 'zone-a-mobile-row2' || p.id === 'zone-a-mobile-row3') return false; p = p.parentElement; }
      return !b.disabled;
    });
    if (desktopBtns.length === 0) return null;
    desktopBtns[0].click();
    return desktopBtns[0].dataset.booth;
  });
  console.log('   selected:', selectedId);
  await sleep(800);

  // 하단 선택 바가 보이도록 모달 본문 스크롤 다운
  await page.evaluate(() => {
    const modal = document.getElementById('booth-map-modal');
    if (modal) {
      const scrollBox = modal.querySelector('.overflow-y-auto') || modal;
      scrollBox.scrollTop = scrollBox.scrollHeight;
    }
  });
  await sleep(800);
  await page.screenshot({ path: path.join(OUT, '03-booth-selected.png'), fullPage: false });

  // Step 4: 신청서 작성 버튼 클릭
  console.log('5. Click 신청서 작성');
  await page.evaluate(() => {
    const btn = [...document.querySelectorAll('button')].find(b => b.textContent.includes('신청서 작성'));
    if (btn) btn.click();
  });
  await sleep(1800);

  // 신청서 모달 상단 캡처
  await page.evaluate(() => {
    const m = document.getElementById('booth-modal');
    if (m) {
      const sb = m.querySelector('.overflow-y-auto') || m;
      sb.scrollTop = 0;
    }
  });
  await sleep(500);
  await page.screenshot({ path: path.join(OUT, '04-application-form-top.png'), fullPage: false });

  // 신청서 중간 캡처
  await page.evaluate(() => {
    const m = document.getElementById('booth-modal');
    if (m) {
      const sb = m.querySelector('.overflow-y-auto') || m;
      sb.scrollTop = sb.scrollHeight * 0.45;
    }
  });
  await sleep(500);
  await page.screenshot({ path: path.join(OUT, '05-application-form-middle.png'), fullPage: false });

  // 신청서 하단 (제출 버튼) 캡처
  await page.evaluate(() => {
    const m = document.getElementById('booth-modal');
    if (m) {
      const sb = m.querySelector('.overflow-y-auto') || m;
      sb.scrollTop = sb.scrollHeight;
    }
  });
  await sleep(500);
  await page.screenshot({ path: path.join(OUT, '06-application-form-bottom.png'), fullPage: false });

  // Step 6: 결제 페이지 캡처 (실제 폼 제출 없이 결제 URL 직접 이동)
  console.log('6. Goto payment page');
  await page.goto('https://pay.bnikorea.com/linkpay/1774064635225x621795363979827700', { waitUntil: 'networkidle2', timeout: 60000 }).catch(e => console.log('payment goto err:', e.message));
  await sleep(3500);
  await page.screenshot({ path: path.join(OUT, '07-payment-page.png'), fullPage: false });

  // 결제 페이지 전체 (스크롤 길이 캡처)
  await page.screenshot({ path: path.join(OUT, '08-payment-page-full.png'), fullPage: true });

  await browser.close();
  console.log('Done.');
})();
