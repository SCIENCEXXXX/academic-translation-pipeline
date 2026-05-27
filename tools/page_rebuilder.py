from __future__ import annotations

from pathlib import Path

try:
    from .layout_analyzer import RebuiltPage, analyze_translated_page
    from .terminology import enforce_terms
except ImportError:
    from layout_analyzer import RebuiltPage, analyze_translated_page
    from terminology import enforce_terms


def load_translated_page(translated_dir: Path, page_number: int) -> str | None:
    path = translated_dir / f"page_{page_number:04d}.md"
    if not path.exists():
        return None
    return path.read_text(encoding="utf-8")


def rebuild_page_from_translation(translated_dir: Path, page_number: int) -> RebuiltPage | None:
    markdown = load_translated_page(translated_dir, page_number)
    if markdown is None:
        return None
    markdown = enforce_terms(markdown)
    return analyze_translated_page(page_number, markdown)
