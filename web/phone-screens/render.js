// Usage: node render.js <scenes.json> <outdir>
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');
(async () => {
  const scenes = JSON.parse(fs.readFileSync(process.argv[2], 'utf8'));
  const out = process.argv[3];
  fs.mkdirSync(out, { recursive: true });
  const browser = await chromium.launch();
  const page = await browser.newPage({ viewport: { width: 480, height: 970 }, deviceScaleFactor: 1 });
  await page.goto('file://' + path.resolve(__dirname, 'phone.html'));
  await page.evaluate(() => document.fonts.ready);
  for (const [name, S] of Object.entries(scenes)) {
    await page.evaluate((s) => window.render(s), S);
    await page.evaluate(() => document.fonts.ready);
    await page.waitForTimeout(60);
    await page.screenshot({ path: path.join(out, name + '.png'), omitBackground: true });
    console.log('rendered', name);
  }
  await browser.close();
})();
