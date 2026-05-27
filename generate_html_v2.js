
const fs = require('fs');
const path = require('path');
const { marked } = require('marked');

const projectRoot = __dirname;
const mdPath = path.join(projectRoot, 'translated', 'Chapter_02_Part1_CN.md');

if (!fs.existsSync(mdPath)) {
    console.error('❌ 找不到完整翻译文件');
    process.exit(1);
}

let markdown = fs.readFileSync(mdPath, 'utf8');

const renderer = new marked.Renderer();

// 💡 核心：将图片转换为 Base64，确保 100% 在 PDF 中显示
renderer.image = ({ href, text }) => {
    try {
        const imgPath = path.join(projectRoot, href.replace('../', ''));
        const ext = path.extname(imgPath).substring(1);
        const buffer = fs.readFileSync(imgPath);
        const base64 = buffer.toString('base64');
        return `
        <figure class="img-box">
            <img src="data:image/${ext};base64,${base64}" alt="${text}">
            <figcaption>${text}</figcaption>
        </figure>`;
    } catch (e) {
        return `<p style="color:red">图片加载失败: ${href}</p>`;
    }
};

marked.setOptions({ renderer });
const htmlBody = marked.parse(markdown);

const htmlContent = `
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <link rel="stylesheet" href="file://${path.join(projectRoot, 'pdf-style.css')}">
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Noto+Serif+SC:wght@400;700&family=Noto+Sans+SC:wght@400;700&display=swap');
        body { font-family: 'Noto Serif SC', serif; font-size: 13pt; line-height: 1.8; color: #333; text-align: justify; }
        .main-body { padding: 3cm; max-width: 210mm; margin: 0 auto; }
        h1, h2, h3 { color: #1e3a8a; font-family: 'Noto Sans SC', sans-serif; page-break-after: avoid; }
        h1 { font-size: 3em; text-align: center; border-bottom: 3px solid #1e3a8a; padding-bottom: 0.5em; margin-bottom: 2em; page-break-before: always; }
        h2 { font-size: 1.8em; border-left: 10px solid #1e3a8a; padding-left: 15px; background: #f1f5f9; margin-top: 2em; }
        .img-box { margin: 2em 0; text-align: center; break-inside: avoid; }
        img { max-width: 75%; max-height: 450px; height: auto; border: 1px solid #ddd; padding: 5px; box-shadow: 0 5px 15px rgba(0,0,0,0.1); }
        figcaption { font-size: 13px; color: #1e3a8a; font-weight: bold; margin-top: 10px; }
        p { text-indent: 2em; margin-bottom: 1.5em; }
        ul, ol { margin-left: 2em; margin-bottom: 1.5em; }
    </style>
</head>
<body>
    <div id="content" class="main-body">
        ${htmlBody}
    </div>
    <script>window.rendered = true;</script>
</body>
</html>
`;

fs.writeFileSync(path.join(projectRoot, 'full_book.html'), htmlContent);
console.log('✅ 终极 Base64 HTML 已生成。');
