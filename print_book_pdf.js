const puppeteer = require('puppeteer');
const path = require('path');
const fs = require('fs');

(async () => {
  const inputHtml = path.resolve(process.argv[2]);
  const outputPdf = path.resolve(process.argv[3]);
  const title = process.argv[4] || '中文译制版';
  fs.mkdirSync(path.dirname(outputPdf), { recursive: true });

  const browser = await puppeteer.launch({
    args: ['--no-sandbox', '--disable-setuid-sandbox', '--allow-file-access-from-files'],
  });
  const page = await browser.newPage();
  page.setDefaultNavigationTimeout(300000);
  await page.goto('file:///' + inputHtml.replace(/\\/g, '/'), { waitUntil: 'load' });
  await page.waitForFunction(() => window.rendered === true, { timeout: 300000 });
  await page.pdf({
    path: outputPdf,
    format: 'A4',
    printBackground: true,
    displayHeaderFooter: false,
    margin: { top: '0mm', bottom: '0mm', left: '0mm', right: '0mm' },
    timeout: 0,
  });
  await browser.close();
  console.log(`PDF rendered: ${outputPdf}`);
})();
