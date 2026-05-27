from __future__ import annotations

import argparse
import json
import re
from dataclasses import dataclass, asdict
from pathlib import Path

import fitz

try:
    from text_encoding_repair import repair_pdf_glyph_text
except ImportError:
    from .text_encoding_repair import repair_pdf_glyph_text


@dataclass
class PageText:
    page_number: int
    text: str


def normalize_pdf_text(text: str) -> str:
    text = repair_pdf_glyph_text(text)
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"(\w)-\n(\w)", r"\1\2", text)
    text = re.sub(r"[ \t]+\n", "\n", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def extract_pdf_text(pdf_path: Path, start: int | None = None, end: int | None = None) -> list[PageText]:
    pages: list[PageText] = []
    with fitz.open(pdf_path) as doc:
        first = max(start or 1, 1)
        last = min(end or doc.page_count, doc.page_count)
        for index in range(first - 1, last):
            text = normalize_pdf_text(doc.load_page(index).get_text("text"))
            pages.append(PageText(page_number=index + 1, text=text))
    return pages


def main() -> int:
    parser = argparse.ArgumentParser(description="Extract normalized text from a PDF.")
    parser.add_argument("--input", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--start", type=int)
    parser.add_argument("--end", type=int)
    args = parser.parse_args()

    pages = extract_pdf_text(args.input, args.start, args.end)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(
        json.dumps([asdict(page) for page in pages], ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
