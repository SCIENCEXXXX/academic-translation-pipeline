const puppeteer = require('puppeteer');
const path = require('path');
const fs = require('fs');

(async () => {
  const browser = await puppeteer.launch({
    args: ['--no-sandbox', '--disable-setuid-sandbox', '--allow-file-access-from-files'],
  });
  const page = await browser.newPage();
  page.setDefaultNavigationTimeout(300000);

  const inputHtmlPath = process.env.BOOK_OUTPUT_HTML
    ? path.resolve(process.env.BOOK_OUTPUT_HTML)
    : path.join(__dirname, 'full_book.html');
  const finalPdfPath = process.env.BOOK_OUTPUT_PDF
    ? path.resolve(process.env.BOOK_OUTPUT_PDF)
    : path.join(__dirname, 'output', 'BOOK_PIPELINE.pdf');

  const outputDir = path.dirname(finalPdfPath);
  if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir, { recursive: true });

  await page.goto('file:///' + inputHtmlPath.replace(/\\/g, '/'), { waitUntil: 'load' });
  await page.waitForFunction(() => window.rendered === true, { timeout: 300000 });

  await page.pdf({
    path: finalPdfPath,
    format: 'A4',
    printBackground: true,
    displayHeaderFooter: true,
    headerTemplate:
      '<div style="font-size:8px;width:100%;text-align:center;color:#9aa4b2;">344 Olympic weightlifting training exercises - 中文翻译版</div>',
    footerTemplate:
      '<div style="font-size:8px;width:100%;text-align:center;color:#9aa4b2;"><span class="pageNumber"></span> / <span class="totalPages"></span></div>',
    margin: { top: '42px', bottom: '48px', left: '42px', right: '42px' },
    timeout: 0,
  });

  await browser.close();
  console.log(`PDF rendered: ${finalPdfPath}`);
})();
