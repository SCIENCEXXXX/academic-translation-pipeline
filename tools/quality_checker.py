from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

try:
    from .extract_pdf_images import ExtractedImage
    from .layout_analyzer import RebuiltPage
    from .text_polisher import is_ai_hallucination
    from .terminology import TERMINOLOGY
except ImportError:
    from extract_pdf_images import ExtractedImage
    from layout_analyzer import RebuiltPage
    from text_polisher import is_ai_hallucination
    from terminology import TERMINOLOGY


@dataclass
class QualityIssue:
    page_number: int
    issue_type: str
    description: str
    suggestion: str


def check_pages(pages: list[RebuiltPage], images: list[ExtractedImage]) -> list[QualityIssue]:
    issues: list[QualityIssue] = []
    image_pages = {image.page_number for image in images}

    for page in pages:
        # Blank/front-matter separator pages are common in books and should not fail output.
        for action in page.actions:
            if _looks_like_reference_or_toc_entry(action.title):
                continue
            if is_ai_hallucination(action.body):
                issues.append(QualityIssue(page.page_number, "AI对话腔残留", f"{action.number}. {action.title} 含模型说明或伪内容。", "删除该段或回到原文重译。"))
            if "：" in action.title or ":" in action.title:
                issues.append(QualityIssue(page.page_number, "标题正文混排", f"{action.number} 的标题仍含冒号。", "将动作名称和说明拆开。"))
            if len(action.body) < 12:
                issues.append(QualityIssue(page.page_number, "说明过短", f"{action.number}. {action.title} 的正文可能未完整识别。", "检查段落合并规则。"))
            if re.search(r"\b(clean pull|hang clean|split clean|jump shrug)\b", action.body, flags=re.I):
                issues.append(QualityIssue(page.page_number, "术语未翻译", f"{action.number}. {action.title} 可能残留英文术语。", "优先使用术语表译名。"))
        for para in page.intro:
            if is_ai_hallucination(para):
                issues.append(QualityIssue(page.page_number, "AI对话腔残留", "段落中含模型说明或请求用户补充内容。", "删除该段并回到原文重译。"))
            short_lines = [line for line in para.splitlines() if 0 < len(line.strip()) < 12]
            if len(short_lines) >= 3:
                issues.append(QualityIssue(page.page_number, "连续短行", "疑似保留了 PDF 原始断行。", "重新合并段落。"))
            for source in TERMINOLOGY:
                if re.search(rf"\b{re.escape(source)}\b", para, flags=re.I):
                    issues.append(QualityIssue(page.page_number, "术语未翻译", f"段落中残留英文术语：{source}", "使用术语表统一替换或重译。"))
                    break

    return issues


def _looks_like_reference_or_toc_entry(title: str) -> bool:
    if re.search(r"\b(19|20)\d{2}\b", title):
        return True
    if re.search(r"\b[A-Z][a-z]+,\s*[A-Z]\.", title):
        return True
    if "《" in title and "》" in title and re.search(r"[A-Za-z]", title):
        return True
    if re.match(r"^\d+(\.\d+)+\s+", title):
        return True
    if re.match(r"^\d+\s*[.．]\s*", title):
        return True
    if re.match(r"^\d+\s+", title) and len(title) < 40:
        return True
    if title in {"训练效应", "训练课的类型"}:
        return True
    if "建模" in title and len(title) < 40:
        return True
    return False


def write_quality_report(path: Path, issues: list[QualityIssue]) -> None:
    lines = ["# 中文排版质量检查报告", ""]
    if not issues:
        lines.append("未发现阻断性问题。")
    else:
        lines.extend(["| 页码 | 问题类型 | 问题说明 | 建议修复方式 |", "|---:|---|---|---|"])
        for issue in issues:
            lines.append(
                f"| {issue.page_number} | {issue.issue_type} | {issue.description} | {issue.suggestion} |"
            )
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
