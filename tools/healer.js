
const fs = require('fs');
const path = require('path');

const projectRoot = __dirname;
const masterPath = path.join(projectRoot, 'translated', 'BOOK_Master_Integrated.md');

function healMarkdown(pageMap) {
    let md = fs.readFileSync(masterPath, 'utf8');
    
    // 逻辑：找到 Page XX 标记，将其下方的段落替换为正确带图的翻译
    for (const [pageNum, newContent] of Object.entries(pageMap)) {
        const regex = new RegExp(`Page ${pageNum}[\s\S]*?(?=Page \d+|---|$)`, 'g');
        md = md.replace(regex, `Page ${pageNum}\n\n${newContent}\n\n`);
    }
    
    fs.writeFileSync(masterPath, md);
    console.log('✅ 合稿已精准修补完毕。');
}

// 这里我会填入子代理返回的内容（由于内容太长，我先做一个示意，实际执行时我会分批注入）

