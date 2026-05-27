
const fs = require('fs');
const path = require('path');
const { marked } = require('marked');

// 绝对路径定位
const projectRoot = __dirname;
const mdPath = path.join(projectRoot, 'translated', 'Chapter_01_Actual_Content.md');
const imageDir = path.join(projectRoot, 'images');

if (!fs.existsSync(mdPath)) {
    console.error('❌ 找不到翻译文件:', mdPath);
    process.exit(1);
}

let markdown = fs.readFileSync(mdPath, 'utf8');

// 💡 修正图片引用，确保路径正确
markdown = markdown.replace(/images\//g, imageDir + '/');

const renderer = new marked.Renderer();
// 💡 使用绝对路径 src
renderer.image = ({ href, text }) => {
    const absolutePath = href.startsWith('/') ? href : path.join(projectRoot, href);
    console.log('Resolving image:', absolutePath);
    return `
    <figure class="img-box">
        <img src="file://${absolutePath}" alt="${text}">
        <figcaption style="font-weight:bold; margin-top:10px;">${text}</figcaption>
    </figure>`;
};

marked.setOptions({ renderer });
const htmlBody = marked.parse(markdown);

const htmlContent = `
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <link rel="stylesheet" href="file://${path.join(projectRoot, 'pdf-style.css')}">
    <style>
        .page-marker { text-align: right; font-size: 10px; color: #ccc; border-bottom: 1px solid #eee; margin: 3em 0 1em; padding-bottom: 5px; page-break-after: avoid; }
        .img-box { margin: 3em 0; text-align: center; break-inside: avoid; }
        img { max-width: 90%; height: auto; border: 1px solid #eee; padding: 5px; background: #fff; }
        figcaption { font-size: 14px; color: #1e3a8a; }
        .main-body { padding: 2.5cm 3cm; max-width: 210mm; margin: 0 auto; line-height: 1.8; }
        h1, h2, h3 { color: #1e3a8a; }
        p { text-align: justify; text-indent: 2em; margin-bottom: 1.2em; }
        h1 + p, h2 + p, h3 + p, .img-box + p { text-indent: 0; }
    </style>
</head>
<body>
    <div id="content" class="main-body">
        ${htmlBody}
    </div>
    <script>
        // 确保所有图片加载后再标记完成
        window.onload = () => {
            const imgs = document.getElementsByTagName('img');
            const promises = Array.from(imgs).map(img => {
                if (img.complete) return Promise.resolve();
                return new Promise(resolve => { img.onload = resolve; img.onerror = resolve; });
            });
            Promise.all(promises).then(() => {
                window.rendered = true;
            });
        };
    </script>
</body>
</html>
`;

fs.writeFileSync(path.join(projectRoot, 'full_book.html'), htmlContent);
console.log('✅ HTML 已成功生成，已注入图片绝对路径。');
