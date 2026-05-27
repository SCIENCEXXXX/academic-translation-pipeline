from __future__ import annotations

import html
import json
import subprocess
from collections import defaultdict
from pathlib import Path

import fitz

try:
    from .extract_pdf_images import ExtractedImage
    from .layout_analyzer import RebuiltPage
except ImportError:
    from extract_pdf_images import ExtractedImage
    from layout_analyzer import RebuiltPage


def render_cover_image(input_pdf: Path, output_dir: Path) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    cover_path = output_dir / "cover.jpg"
    with fitz.open(input_pdf) as doc:
        pix = doc.load_page(0).get_pixmap(matrix=fitz.Matrix(2.2, 2.2), alpha=False)
        pix.save(cover_path)
    return cover_path


def file_url(path: Path) -> str:
    return "file:///" + str(path.resolve()).replace("\\", "/")


def paragraph(text: str) -> str:
    return f"<p>{html.escape(text)}</p>"


def render_page(page: RebuiltPage, page_images: list[ExtractedImage]) -> str:
    parts = [f'<div class="source-page">—— 原书第 {page.page_number} 页 ——</div>']
    for item in page.intro:
        parts.append(paragraph(item))

    for kind, value in page.elements:
        if kind == "section":
            parts.append(f'<h2 class="section-title">{html.escape(str(value))}</h2>')
        else:
            parts.append(f'<h3 class="action-title">{html.escape(value.number)}. {html.escape(value.title)}</h3>')
            parts.append(paragraph(value.body))

    if not page.elements:
        for section in page.sections:
            parts.append(f'<h2 class="section-title">{html.escape(section)}</h2>')
        for action in page.actions:
            parts.append(f'<h3 class="action-title">{html.escape(action.number)}. {html.escape(action.title)}</h3>')
            parts.append(paragraph(action.body))

    for idx, image in enumerate(page_images, start=1):
        image_path = Path(image.path)
        if image_path.exists():
            parts.append(
                '<figure>'
                f'<img src="{file_url(image_path)}" alt="图 {page.page_number}-{idx}">'
                f'<figcaption>图 {page.page_number}-{idx} 原书第 {page.page_number} 页插图</figcaption>'
                '</figure>'
            )

    return "\n".join(parts)


def render_book_html(
    output_html: Path,
    input_pdf: Path,
    pages: list[RebuiltPage],
    images: list[ExtractedImage],
    title: str,
) -> Path:
    asset_dir = output_html.parent / "assets"
    cover_path = render_cover_image(input_pdf, asset_dir)
    by_page: dict[int, list[ExtractedImage]] = defaultdict(list)
    for image in images:
        by_page[image.page_number].append(image)

    body = "\n".join(render_page(page, by_page.get(page.page_number, [])) for page in pages)
    html_content = f"""<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<title>{html.escape(title)}</title>
<style>
  @page {{
    size: A4;
    margin: 24mm 24mm 24mm 28mm;
  }}
  * {{ box-sizing: border-box; }}
  body {{
    margin: 0;
    color: #111827;
    background: #fff;
    font-family: SimSun, "Songti SC", "Noto Serif SC", serif;
    font-size: 11pt;
    line-height: 1.85;
    text-rendering: optimizeLegibility;
  }}
  .cover {{
    page-break-after: always;
    height: 297mm;
    display: flex;
    align-items: stretch;
    justify-content: center;
    background: #090909;
  }}
  .cover img {{
    width: 100%;
    height: 100%;
    object-fit: contain;
  }}
  .title-page {{
    page-break-after: always;
    min-height: 235mm;
    display: flex;
    flex-direction: column;
    justify-content: center;
    text-align: center;
  }}
  .title-page h1 {{
    font-family: SimHei, "Microsoft YaHei", sans-serif;
    font-size: 24pt;
    line-height: 1.35;
    margin: 0 0 18mm;
    font-weight: 700;
  }}
  .title-page .sub {{
    font-size: 12pt;
    color: #374151;
  }}
  .preface {{
    page-break-after: always;
  }}
  h1 {{
    font-family: SimHei, "Microsoft YaHei", sans-serif;
    font-size: 16pt;
    margin: 0 0 9mm;
    line-height: 1.4;
  }}
  h2.section-title {{
    font-family: SimHei, "Microsoft YaHei", sans-serif;
    font-size: 13pt;
    margin: 7mm 0 2mm;
    line-height: 1.45;
    page-break-after: avoid;
  }}
  h3.action-title {{
    font-family: SimHei, "Microsoft YaHei", sans-serif;
    font-size: 11.5pt;
    margin: 5mm 0 1.5mm;
    line-height: 1.45;
    page-break-after: avoid;
  }}
  p {{
    margin: 0;
    text-indent: 2em;
    text-align: justify;
    widows: 2;
    orphans: 2;
  }}
  .source-page {{
    text-align: center;
    color: #6b7280;
    font-size: 9.5pt;
    margin: 6mm 0 5mm;
    page-break-after: avoid;
  }}
  figure {{
    margin: 7mm auto;
    text-align: center;
    page-break-inside: avoid;
  }}
  figure img {{
    max-width: 90%;
    max-height: 120mm;
    object-fit: contain;
  }}
  figcaption {{
    margin-top: 2mm;
    color: #4b5563;
    font-size: 9.5pt;
    text-align: center;
  }}
  .content {{
    counter-reset: page;
  }}
</style>
</head>
<body>
  <section class="cover"><img src="{file_url(cover_path)}" alt="封面"></section>
  <section class="title-page">
    <h1>{html.escape(title)}</h1>
    <div class="sub">中文译制版</div>
  </section>
  <section class="preface">
    <h1>译制说明</h1>
    <p>本书稿依据原书 PDF 文本进行中文译制，面向中国大陆举重训练、体能训练与专项力量训练读者。译文按中文训练教材的阅读习惯重建段落层级，保留原书页码来源，并将原书插图作为正文图像嵌入。</p>
    <p>本工作流参考中文运动训练类译著的排版方式，避免逐页截图式翻译，尽量呈现连续、自然、可校对的中文书稿版式。</p>
  </section>
  <main class="content">{body}</main>
  <script>window.rendered = true;</script>
</body>
</html>
"""
    output_html.parent.mkdir(parents=True, exist_ok=True)
    output_html.write_text(html_content, encoding="utf-8")
    return output_html


def render_book_pdf(
    input_pdf: Path,
    output_pdf: Path,
    pages: list[RebuiltPage],
    images: list[ExtractedImage],
    title: str,
    project_root: Path,
) -> Path:
    output_html = output_pdf.with_suffix(".html")
    render_book_html(output_html, input_pdf, pages, images, title)
    subprocess.run(
        ["node", "print_book_pdf.js", str(output_html), str(output_pdf), title],
        cwd=project_root,
        check=True,
    )
    return output_pdf
