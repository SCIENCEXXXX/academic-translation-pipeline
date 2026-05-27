from __future__ import annotations

import re

try:
    from .terminology import enforce_terms
except ImportError:
    from terminology import enforce_terms


META_PATTERNS = [
    r"以下是.*?翻译[，,。]?.*?(?=\n|$)",
    r"以下是.*?(?:OCR|原文|页面).*?(?:翻译|整理).*?(?=\n|$)",
    r"好的，?这是.*?(?:翻译|整理).*?(?=\n|$)",
    r"抱歉，?您提供的文本内容.*?(?=\n|$)",
    r"抱歉，?.*?(?:没有提供|未提供|无法识别|无法翻译).*?(?=\n|$)",
    r"请提供完整的.*?(?=\n|$)",
    r"请(?:将|粘贴|提供).*?(?:OCR|原文|文本|内容).*?(?=\n|$)",
    r"无法根据当前.*?(?=\n|$)",
    r"无法识别.*?(?=\n|$)",
    r"无法翻译.*?(?=\n|$)",
    r"^原文翻译[:：]\s*",
    r"^译文[:：]\s*",
    r"如需.*?(?:告知|提供).*?(?=\n|$)",
    r"如果您需要.*?(?:告知|提供).*?(?=\n|$)",
    r"如果您能提供.*?(?=\n|$)",
    r"若需完整翻译.*?(?=\n|$)",
    r"^\（?注[:：].*?示例翻译.*?$",
    r"^\*?说明\*?[:：].*?(?=\n|$)",
    r"^\*?OCR修复\*?[:：].*?(?=\n|$)",
    r"^\*?术语统一\*?[:：].*?(?=\n|$)",
    r"^\*?格式\*?[:：].*?(?=\n|$)",
    r"^\*?语言风格\*?[:：].*?(?=\n|$)",
    r"^\*?内容忠实\*?[:：].*?(?=\n|$)",
]


def strip_markdown_marks(text: str) -> str:
    text = text.replace("```", "")
    text = re.sub(r"\*\*(.*?)\*\*", r"\1", text)
    text = re.sub(r"\*(.*?)\*", r"\1", text)
    text = re.sub(r"`([^`]+)`", r"\1", text)
    text = text.replace("---", "")
    return text


def remove_ai_meta(text: str) -> str:
    for pattern in META_PATTERNS:
        text = re.sub(pattern, "", text, flags=re.IGNORECASE | re.MULTILINE)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def normalize_title(title: str) -> str:
    title = strip_markdown_marks(title).strip()
    title = re.sub(r"\s*[（(][A-Za-z][^)）]{1,80}[)）]\s*", "", title).strip()
    title = title.replace("半程翻铃", "半蹲翻铃")
    title = title.replace("三级翻铃", "三段翻铃")
    title = title.replace("翻铃上拉", "翻铃拉")
    return enforce_terms(title)


def polish_body(title: str, body: str) -> str:
    body = strip_markdown_marks(body)
    body = remove_ai_meta(body)
    body = enforce_terms(body)
    body = re.sub(r"[（(][A-Za-z][A-Za-z0-9 %×@,./+-]{1,80}[)）]", "", body)
    body = re.sub(r"\bHigh Hang Pull\b", "高位悬垂拉", body, flags=re.IGNORECASE)
    body = re.sub(r"\bHigh Hang Power Snatch\b", "高位悬垂力量抓举", body, flags=re.IGNORECASE)
    body = re.sub(r"\bSnatch\b", "抓举", body, flags=re.IGNORECASE)
    body = re.sub(r"\bClean and Jerk\b", "挺举", body, flags=re.IGNORECASE)
    body = re.sub(r"\bClean\b", "翻铃", body, flags=re.IGNORECASE)
    body = re.sub(r"\bSquat\b", "深蹲", body, flags=re.IGNORECASE)
    body = re.sub(r"\bPull\b", "拉", body, flags=re.IGNORECASE)
    body = re.sub(r"\s+", "", body)
    body = body.replace("此动作", "该动作")
    body = body.replace("本动作", "该动作")
    body = body.replace("可以作为", "可作为")
    body = body.replace("几乎总是使用", "通常会使用")
    body = body.replace("保持躯干挺直", "保持躯干直立")
    body = body.replace("重心平衡于前脚掌", "重心保持在前脚掌附近")

    title = normalize_title(title)
    starters = (
        "可作为",
        "是",
        "一种",
        "传统动作",
        "翻铃学习",
        "抓举学习",
        "挺举学习",
        "用于",
        "适用于",
        "针对",
        "另一",
        "因",
    )
    if body.startswith("可作为"):
        body = f"{title}{body}"
    elif body.startswith("一种"):
        body = f"{title}是{body}"
    elif body.startswith("传统动作"):
        body = f"{title}是一种{body}"
    elif body.startswith("翻铃学习") or body.startswith("抓举学习") or body.startswith("挺举学习"):
        body = f"{title}是{body}"
    elif body.startswith("用于") or body.startswith("适用于") or body.startswith("针对"):
        body = f"{title}{body}"
    elif body.startswith("另一"):
        body = f"{title}是{body}"
    elif body.startswith("因"):
        body = f"{title}{body}"
    elif not body.startswith(title) and any(body.startswith(s) for s in starters):
        body = f"{title}{body}"

    body = re.sub(r"([。！？；])", r"\1", body)
    body = body.replace("。但", "。但")
    body = body.replace("。由于", "。由于")
    return body.strip()


def is_ai_hallucination(text: str) -> bool:
    bad_signals = [
        "如果您需要",
        "若需完整翻译",
        "请提供",
        "示例翻译",
        "以下为基于常见训练场景",
        "训练计划 77",
        "训练计划77",
    ]
    return any(signal in text for signal in bad_signals)
