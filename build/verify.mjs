import { chromium } from 'playwright-core';
const EXE = '/opt/pw-browsers/chromium-1194/chrome-linux/chrome';
const FILE = 'file://' + process.cwd() + '/deliverables/USC_Dashboard_v2.html';
const OUT = '/tmp/claude-0/-home-user-Claude/8455871d-1e33-59c7-adbb-f9c39ca13e7f/scratchpad';
const proxy = process.env.HTTPS_PROXY || process.env.https_proxy;

const args = ['--no-sandbox','--disable-dev-shm-usage'];
const launch = { executablePath: EXE, args };
if (proxy) launch.proxy = { server: proxy };

const browser = await chromium.launch(launch);
const page = await browser.newPage({ viewport: { width: 1440, height: 900 }, ignoreHTTPSErrors: true });
const errors = [], warns = [];
page.on('pageerror', e => errors.push('PAGEERROR: ' + e.message));
page.on('console', m => { if (m.type()==='error') errors.push('CONSOLE.ERR: ' + m.text().slice(0,200)); });

try {
  await page.goto(FILE, { waitUntil: 'load', timeout: 45000 });
  await page.waitForTimeout(3500); // let CDNs + boot run
  const diag = await page.evaluate(() => ({
    raw: (window.RAW||[]).length,
    meta: window.DASH_META ? { rows: DASH_META.rowCount, zones: (DASH_META.zones||[]).length, cats: (DASH_META.categoryOrder||[]).length } : null,
    chartjs: typeof window.Chart !== 'undefined',
    tailwind: !!document.querySelector('script[src*="tailwind"]'),
    status: (document.getElementById('statusText')||{}).textContent,
    kpi1Records: (document.getElementById('kpi1Records')||{}).textContent,
    kpi4Pay: (document.getElementById('kpi4Pay')||{}).textContent,
    msRings: document.querySelectorAll('#milestoneMatrix .ring').length,
    msCols: document.querySelectorAll('#milestoneMatrix thead th').length,
  }));
  console.log('DIAG', JSON.stringify(diag, null, 2));

  const tabs = [['tab0','Category Milestone'],['tab1','Overall'],['tab2','Quantity'],['tab3','Auction'],['tab4','Payment']];
  for (const [t,label] of tabs) {
    await page.click(`.tab-btn[data-tab="${t}"]`);
    await page.waitForTimeout(900);
    await page.screenshot({ path: `${OUT}/shot_${t}.png`, fullPage: true });
    console.log(`shot ${t} (${label}) saved`);
  }
} catch (e) {
  console.log('FATAL', e.message);
}
console.log('--- ERRORS (' + errors.length + ') ---');
errors.slice(0,25).forEach(e => console.log(e));
await browser.close();
