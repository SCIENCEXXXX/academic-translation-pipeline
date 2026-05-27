from __future__ import annotations

import argparse
import json
import os
import re
from pathlib import Path

from dotenv import load_dotenv
from tqdm import tqdm

try:
    from translate_pdf import Translator, load_config, translate_page
except ImportError:
    from .translate_pdf import Translator, load_config, translate_page


def page_number_from_path(path: Path) -> int | None:
    stem = path.stem
    parts = stem.split("_")
    for part in reversed(parts):
        if part.isdigit():
            return int(part)
    return None


def clean_source_page_markdown(text: str) -> str:
    """Remove local OCR bookkeeping before the text is sent to the model."""
    text = re.sub(r"(?m)^#\s*OCR\s*第\s*\d+\s*页\s*$", "", text)
    text = re.sub(r"<!--\s*ocr elapsed=.*?-->", "", text, flags=re.IGNORECASE | re.DOTALL)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def main() -> int:
    load_dotenv()
    parser = argparse.ArgumentParser(description="Translate OCR/source page Markdown files.")
    parser.add_argument("--source-dir", required=True, type=Path)
    parser.add_argument("--translated-dir", required=True, type=Path)
    parser.add_argument("--manifest", required=True, type=Path)
    parser.add_argument("--config", required=True, type=Path)
    parser.add_argument("--start", type=int, default=None)
    parser.add_argument("--end", type=int, default=None)
    parser.add_argument("--overwrite", action="store_true")
    args = parser.parse_args()

    config = load_config(args.config)
    translator = Translator(config)
    args.translated_dir.mkdir(parents=True, exist_ok=True)

    files = sorted(args.source_dir.glob("page_*.md"))
    selected: list[tuple[int, Path]] = []
    for file in files:
        number = page_number_from_path(file)
        if number is None:
            continue
        if args.start is not None and number < args.start:
            continue
        if args.end is not None and number > args.end:
            continue
        selected.append((number, file))

    written: list[str] = []
    skipped = 0
    for page_number, source_file in tqdm(selected, desc="Translating OCR pages"):
        output_file = args.translated_dir / f"page_{page_number:04d}.md"
        if output_file.exists() and not args.overwrite:
            skipped += 1
            written.append(str(output_file))
            continue
        text = clean_source_page_markdown(source_file.read_text(encoding="utf-8", errors="ignore"))
        translated = translate_page(page_number, text, translator, config)
        output_file.write_text(f"# 第 {page_number} 页\n\n{translated}", encoding="utf-8")
        written.append(str(output_file))

    args.manifest.parent.mkdir(parents=True, exist_ok=True)
    args.manifest.write_text(
        json.dumps(
            {
                "source_dir": str(args.source_dir),
                "translated_dir": str(args.translated_dir),
                "pages": len(selected),
                "skipped": skipped,
                "files": written,
            },
            ensure_ascii=False,
            indent=2,
        )
        + "\n",
        encoding="utf-8",
    )
    print(json.dumps({"pages": len(selected), "skipped": skipped, "manifest": str(args.manifest)}, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
