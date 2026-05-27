from __future__ import annotations

import re
from dataclasses import dataclass, field

try:
    from .text_polisher import is_ai_hallucination, normalize_title, polish_body, remove_ai_meta, strip_markdown_marks
except ImportError:
    from text_polisher import is_ai_hallucination, normalize_title, polish_body, remove_ai_meta, strip_markdown_marks


@dataclass
class ActionItem:
    number: str
    title: str
    body: str


@dataclass
class RebuiltPage:
    page_number: int
    intro: list[str] = field(default_factory=list)
    actions: list[ActionItem] = field(default_factory=list)
    sections: list[str] = field(default_factory=list)
    elements: list[tuple[str, object]] = field(default_factory=list)


ACTION_RE = re.compile(
    r"(?ms)(?P<number>\d{1,3})[.、]\s*(?P<title>[^：:\n]{2,80})[：:]\s*(?P<body>.*?)(?=\n\s*\d{1,3}[.、]\s*[^：:\n]{2,80}[：:]|\n\s*[○●o]\s+|$)"
)


def clean_translated_markdown(markdown: str) -> str:
    text = re.sub(r"^#\s*第\s*\d+\s*页\s*", "", markdown, flags=re.MULTILINE)
    text = re.sub(r"!\[[^\]]+\]\([^)]+\)", "", text)
    text = re.sub(r"<!--.*?-->", "", text, flags=re.DOTALL)
    text = re.sub(r"^\s*\d+\s*$", "", text, flags=re.MULTILINE)
    text = remove_ai_meta(text)
    text = strip_markdown_marks(text)
    text = text.replace("○", "\n○").replace("●", "\n●")
    text = re.sub(r"[ \t]+\n", "\n", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def compact_body(text: str) -> str:
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    joined = "".join(lines)
    joined = re.sub(r"[（(][A-Za-z][A-Za-z0-9 %×@,./+-]{1,80}[)）]", "", joined)
    joined = re.sub(r"(?m)^[A-Za-z][A-Za-z0-9 %×@,./+-]{2,80}$", "", joined)
    joined = re.sub(r"\s+", " ", joined)
    joined = joined.replace(" 。", "。").replace(" ，", "，")
    return joined.strip()


def analyze_translated_page(page_number: int, markdown: str) -> RebuiltPage:
    text = clean_translated_markdown(markdown)
    page = RebuiltPage(page_number=page_number)

    consumed: list[tuple[int, int]] = []
    for match in ACTION_RE.finditer(text):
        title = normalize_title(match.group("title").strip())
        body = polish_body(title, compact_body(match.group("body")))
        if title and body:
            action = ActionItem(match.group("number"), title, body)
            page.actions.append(action)
            page.elements.append(("action", action))
            consumed.append(match.span())

    if consumed:
        intro_text = text[: consumed[0][0]].strip()
    else:
        intro_text = text

    for section_match in re.finditer(r"(?m)^[○●o]\s*(.+)$", text):
        section = section_match.group(1).strip()
        if section:
            page.sections.append(section)
            page.elements.append(("section", section))

    page.elements.sort(key=lambda item: _element_position(text, item))

    for para in re.split(r"\n{2,}", intro_text):
        para = compact_body(para)
        para = remove_ai_meta(para)
        if para and not is_ai_hallucination(para):
            page.intro.append(para)

    return page


def _element_position(text: str, item: tuple[str, object]) -> int:
    kind, value = item
    if kind == "section":
        pos = text.find(str(value))
        return pos if pos >= 0 else 10**9
    action = value
    if isinstance(action, ActionItem):
        pattern = f"{action.number}."
        pos = text.find(pattern)
        return pos if pos >= 0 else 10**9
    return 10**9
