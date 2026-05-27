# Academic Translation Pipeline

面向体育训练、体能训练、举重、足球科学等专业书籍的 PDF 翻译与中文书稿排版工作流。目标不是生成逐页翻译草稿，而是尽量输出可阅读、可打印、便于后续人工校对的中文终稿。

## 核心能力

- 从英文 PDF 提取文本、图片和参考文献信息。
- 调用 OpenAI 兼容接口进行分页翻译，支持 DeepSeek。
- 按书籍类型启用术语表和翻译提示词。
- 重建中文章节层级，减少 PDF 抽取导致的目录乱序、标题错位和硬换行。
- 将原书图片回插到译文中，避免整页截图式翻译。
- 保留参考文献原文格式，不强行翻译英文文献条目。
- 输出 `final_zh.docx` 和 `final_zh.pdf`。

## 目录结构

```text
config/        翻译提示词、术语表和书籍类型配置
tools/         PDF 解析、翻译、排版、质量检查和终稿生成脚本
*.js           旧版 HTML/PDF 渲染脚本，保留兼容
package.json   Node 依赖和兼容脚本
requirements.txt
```

以下目录用于本地运行，不应提交到 GitHub：

```text
source_pdfs/
runs/
output/
translated/
images/
source_full/
```

## 环境准备

```powershell
python -m pip install -r requirements.txt
npm install
```

如需从 DOCX 导出 PDF，Windows 环境建议安装 Microsoft Word。若没有 Word，可先生成 DOCX，再用其他工具手动导出 PDF。

## API 配置

复制环境变量模板：

```powershell
Copy-Item .env.example .env
```

DeepSeek 示例：

```text
DEEPSEEK_API_KEY=你的 DeepSeek Key
DEEPSEEK_MODEL=deepseek-chat
```

也可以使用 OpenAI 或其他 OpenAI 兼容服务：

```text
OPENAI_API_KEY=你的 API Key
OPENAI_MODEL=gpt-4.1-mini
OPENAI_BASE_URL=https://api.openai.com/v1
```

不要把 `.env` 提交到 GitHub。

## 推荐运行流程

1. 将原始 PDF 放入本地目录，例如：

```text
source_pdfs/book.pdf
```

2. 翻译 PDF：

```powershell
python tools/translate_pdf.py --pdf source_pdfs/book.pdf --output-dir runs/book_translated
```

常用参数：

```powershell
python tools/translate_pdf.py --pdf source_pdfs/book.pdf --start 1 --end 50 --output-dir runs/book_translated
python tools/translate_pdf.py --pdf source_pdfs/book.pdf --config config/soviet_weightlifting_translation.json --output-dir runs/book_translated
```

3. 生成中文终稿：

```powershell
python tools/pipeline.py --mode final --input source_pdfs/book.pdf --translated-dir runs/book_translated --output runs/book_final --title "中文书名"
```

终稿输出：

```text
runs/book_final/final_zh.docx
runs/book_final/final_zh.pdf
```

## 书籍类型配置

当前已包含多套体育训练领域配置，例如：

- `config/soviet_weightlifting_translation.json`
- `config/chinese_weightlifting_translation.json`
- `config/science_soccer_translation.json`
- `config/strength_training_translation.json`
- `config/sports_training_principles_translation.json`
- `config/gilbert_coaching_translation.json`
- `config/gordon_coaching_science_translation.json`
- `config/ltad_translation.json`
- `config/platonov_periodization_translation.json`

新增书籍时，优先复制相近配置并修改术语表、翻译风格和章节识别规则。

## 质量控制

终稿生成阶段会检查常见问题：

- 是否残留 “原书第 X 页” 等内部标记；
- 是否残留 Markdown 表格符号、`<br>` 或代码块；
- 是否出现模型客套语；
- 是否出现错误章节标记；
- 是否存在无效中间文件。

注意：如果原书中的训练计划表或复杂图表本身是图片，工作流会优先保留清晰原图。若要把图片表格完全重建为中文真表格，需要额外安装 OCR 能力并进行人工校对。
