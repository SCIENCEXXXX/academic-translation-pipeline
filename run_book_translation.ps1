# 填好 .env 后可由 Codex 运行本脚本完成整本翻译和排版。
$ErrorActionPreference = "Stop"
$pdf = "D:\谷歌浏览器\344 Olympic weightlifting training exercises used by Chinese weightlifters (Manuel Buitrago) ).pdf"
$desktop = [Environment]::GetFolderPath("Desktop")
$out = Join-Path $desktop "344举重训练动作_中文翻译成品"

npm run translate -- --pdf $pdf --source-dir source_full --translated-dir translated --manifest manifest.json --target-language "简体中文" --overwrite
npm run pipeline

Copy-Item -LiteralPath "translated\BOOK_MASTER_PIPELINE.md" -Destination (Join-Path $out "344举重训练动作_中文翻译.md") -Force
Copy-Item -LiteralPath "full_book.html" -Destination (Join-Path $out "344举重训练动作_中文翻译.html") -Force
Copy-Item -LiteralPath "output\BOOK_PIPELINE.pdf" -Destination (Join-Path $out "344举重训练动作_中文翻译.pdf") -Force
Copy-Item -LiteralPath "translation_run.json" -Destination (Join-Path $out "translation_run.json") -Force
