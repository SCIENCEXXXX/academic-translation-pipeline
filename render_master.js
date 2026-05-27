const fs = require('fs');
const path = require('path');
const { marked } = require('marked');
const katex = require('katex');
const cheerio = require('cheerio');

const projectRoot = __dirname;
const mdPath = process.env.BOOK_INPUT_MD
  ? path.resolve(process.env.BOOK_INPUT_MD)
  : path.join(projectRoot, 'translated', 'BOOK_MASTER_PIPELINE.md');
const imagesDir = process.env.BOOK_IMAGES_DIR
  ? path.resolve(process.env.BOOK_IMAGES_DIR)
  : path.join(projectRoot, 'images', 'pages');
const htmlOutputPath = process.env.BOOK_OUTPUT_HTML
  ? path.resolve(process.env.BOOK_OUTPUT_HTML)
  : path.join(projectRoot, 'full_book.html');

const bookTitle = process.env.BOOK_TITLE || '344 个中国举重运动员使用的奥运举重训练动作';
const bookSubtitle = process.env.BOOK_SUBTITLE || '中文翻译版';
const coverImage = process.env.BOOK_COVER_IMAGE || path.join(imagesDir, 'page_0001.jpg');

let markdown = fs.readFileSync(mdPath, 'utf8');

markdown = markdown.replace(/\$\$([\s\S]*?)\$\$/g, (match, formula) => {
  try {
    return `<div class="katex-display">${katex.renderToString(formula.trim(), {
      displayMode: true,
      throwOnError: false,
    })}</div>`;
  } catch {
    return match;
  }
});

markdown = markdown.replace(/\$([^\$\n]+?)\$/g, (match, formula) => {
  try {
    return katex.renderToString(formula.trim(), { displayMode: false, throwOnError: false });
  } catch {
    return match;
  }
});

const imageQueue = [];
markdown = markdown.replace(/!\[(.*?)\]\((.*?)\)/g, (_match, alt, filename) => {
  const placeholder = `@@IMG_PH_${imageQueue.length}@@`;
  imageQueue.push({ placeholder, alt, filename });
  return placeholder;
});

const rawBodyHtml = marked.parse(markdown);
const $ = cheerio.load(rawBodyHtml);
const toc = [];
const headerIds = new Set();

$('h1, h2').each((i, el) => {
  const heading = $(el);
  const text = heading.text().trim();
  let id = text
    .toLowerCase()
    .replace(/[^\w\u4e00-\u9fa5]+/g, '-')
    .replace(/^-+|-+$/g, '');
  if (!id || headerIds.has(id)) id = `section-${i + 1}`;
  headerIds.add(id);
  heading.attr('id', id);
  toc.push({ level: el.name === 'h1' ? 1 : 2, text, id });
});

let bodyHtml = $.html();

imageQueue.forEach((img) => {
  const cleanName = path.basename(img.filename);
  const candidates = [
    path.resolve(path.dirname(mdPath), img.filename),
    path.join(imagesDir, cleanName),
    path.join(projectRoot, 'images', cleanName),
  ];
  const imgPath = candidates.find((candidate) => fs.existsSync(candidate));
  if (!imgPath) {
    bodyHtml = bodyHtml.replace(img.placeholder, `<p class="missing-image">[图片缺失：${cleanName}]</p>`);
    return;
  }

  const fileUrl = 'file:///' + imgPath.replace(/\\/g, '/');
  const caption = img.alt ? `<div class="caption">${img.alt}</div>` : '';
  bodyHtml = bodyHtml.replace(
    img.placeholder,
    `<figure class="page-image"><img src="${fileUrl}" alt="${img.alt || ''}">${caption}</figure>`,
  );
});

const coverImageHtml = fs.existsSync(coverImage)
  ? `<img class="cover-image" src="file:///${coverImage.replace(/\\/g, '/')}" alt="原书封面">`
  : '';

const tocItems = toc
  .filter((item) => !/^第\s*\d+\s*页$/.test(item.text))
  .slice(0, 80)
  .map((item) => `<div class="toc-item depth-${item.level}"><a href="#${item.id}">${item.text}</a></div>`)
  .join('');

const tocHtml = tocItems
  ? `<section class="toc-page"><h1>目录</h1>${tocItems}</section>`
  : '';

const htmlContent = `<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8">
  <style>
    @page { size: A4; margin: 18mm 18mm 20mm; }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      color: #1f2933;
      background: #fff;
      font-family: "Microsoft YaHei", "Noto Sans SC", "PingFang SC", Arial, sans-serif;
      font-size: 11.5pt;
      line-height: 1.68;
    }
    .cover-page {
      min-height: 255mm;
      page-break-after: always;
      display: grid;
      grid-template-columns: 42% 1fr;
      gap: 24px;
      align-items: center;
    }
    .cover-image {
      width: 100%;
      max-height: 210mm;
      object-fit: contain;
      border: 1px solid #d7dde8;
    }
    .cover-title {
      font-size: 28pt;
      line-height: 1.24;
      color: #172033;
      margin: 0 0 18px;
      font-weight: 800;
      letter-spacing: 0;
    }
    .cover-subtitle {
      font-size: 16pt;
      color: #475569;
      margin: 0 0 8px;
      font-weight: 600;
    }
    .cover-note {
      font-size: 10.5pt;
      color: #64748b;
      margin-top: 26px;
      border-top: 1px solid #d7dde8;
      padding-top: 12px;
    }
    .toc-page {
      page-break-after: always;
    }
    .toc-page h1 {
      page-break-before: auto;
      text-align: left;
      margin-top: 0;
      border-bottom: 2px solid #172033;
    }
    .toc-item {
      margin: 7px 0;
      border-bottom: 1px dotted #cbd5e1;
      padding-bottom: 4px;
    }
    .toc-item a {
      color: #1f2933;
      text-decoration: none;
    }
    .toc-item.depth-2 {
      margin-left: 18px;
      color: #475569;
    }
    .main-body {
      max-width: 176mm;
      margin: 0 auto;
    }
    h1 {
      page-break-before: always;
      font-size: 18pt;
      line-height: 1.3;
      margin: 0 0 14px;
      padding-bottom: 8px;
      border-bottom: 1px solid #cbd5e1;
      color: #172033;
      font-weight: 800;
    }
    h1:first-child {
      page-break-before: auto;
    }
    h2 {
      font-size: 15pt;
      margin: 22px 0 10px;
      color: #172033;
      font-weight: 700;
    }
    h3 {
      font-size: 13pt;
      margin: 18px 0 8px;
      color: #243447;
    }
    p {
      margin: 0 0 9px;
      text-align: left;
      text-indent: 0;
    }
    ul, ol {
      margin: 6px 0 12px 24px;
      padding: 0;
    }
    li {
      margin: 3px 0;
    }
    .page-image {
      margin: 0 0 16px;
      padding: 0;
      page-break-inside: avoid;
      text-align: center;
    }
    .page-image img {
      max-width: 100%;
      max-height: 120mm;
      object-fit: contain;
      border: 1px solid #d7dde8;
      background: white;
    }
    .caption {
      margin-top: 5px;
      font-size: 9pt;
      color: #64748b;
      text-align: center;
    }
    .missing-image {
      color: #b91c1c;
      text-align: center;
      border: 1px solid #fecaca;
      padding: 8px;
    }
    table {
      border-collapse: collapse;
      width: 100%;
      margin: 12px 0;
      font-size: 10pt;
    }
    th, td {
      border: 1px solid #cbd5e1;
      padding: 6px 8px;
      vertical-align: top;
    }
    th {
      background: #eef2f7;
      font-weight: 700;
    }
    code {
      font-family: Consolas, monospace;
    }
  </style>
</head>
<body>
  <section class="cover-page">
    <div>${coverImageHtml}</div>
    <div>
      <h1 class="cover-title">${bookTitle}</h1>
      <p class="cover-subtitle">${bookSubtitle}</p>
      <p class="cover-note">以原书 PDF 为依据翻译；正文按原书页码组织，术语按中国大陆举重训练语境处理。</p>
    </div>
  </section>
  ${tocHtml}
  <main class="main-body">${bodyHtml}</main>
  <script>
    window.tocData = ${JSON.stringify(toc)};
    window.rendered = true;
  </script>
</body>
</html>`;

fs.writeFileSync(htmlOutputPath, htmlContent, 'utf8');
console.log(`HTML rendered: ${htmlOutputPath}`);
