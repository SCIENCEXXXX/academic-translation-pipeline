import argparse
import json
import os
import re
import time
from dataclasses import dataclass
from pathlib import Path

from dotenv import load_dotenv
from openai import OpenAI
from tqdm import tqdm

try:
    import fitz
except ImportError as exc:
    raise SystemExit("Missing dependency: pymupdf. Run `python -m pip install -r requirements.txt`.") from exc

try:
    from text_encoding_repair import repair_pdf_glyph_text
except ImportError:
    from .text_encoding_repair import repair_pdf_glyph_text


@dataclass(frozen=True)
class TranslationConfig:
    source_language: str
    target_language: str
    chunk_size: int
    chunk_overlap: int
    temperature: float
    system_prompt: str
    user_prompt_template: str


def load_config(path: Path) -> TranslationConfig:
    data = {}
    if path.exists():
        data = json.loads(path.read_text(encoding="utf-8"))

    return TranslationConfig(
        source_language=str(data.get("source_language", "English")),
        target_language=str(data.get("target_language", "简体中文")),
        chunk_size=int(data.get("chunk_size", 5000)),
        chunk_overlap=int(data.get("chunk_overlap", 200)),
        temperature=float(data.get("temperature", 0.2)),
        system_prompt=str(data.get("system_prompt", "")),
        user_prompt_template=str(
            data.get(
                "user_prompt_template",
                "Translate from {source_language} to {target_language}:\n\n{text}",
            )
        ),
    )


def extract_pdf_pages(pdf_path: Path, start: int | None, end: int | None) -> list[tuple[int, str]]:
    pages: list[tuple[int, str]] = []
    with fitz.open(pdf_path) as doc:
        first = max((start or 1), 1)
        last = min((end or doc.page_count), doc.page_count)
        for page_index in range(first - 1, last):
            page_number = page_index + 1
            text = doc.load_page(page_index).get_text("text").strip()
            pages.append((page_number, normalize_text(text)))
    return pages


def render_pdf_pages(
    pdf_path: Path,
    output_dir: Path,
    start: int | None,
    end: int | None,
    zoom: float,
) -> list[tuple[int, Path]]:
    output_dir.mkdir(parents=True, exist_ok=True)
    rendered: list[tuple[int, Path]] = []
    with fitz.open(pdf_path) as doc:
        first = max((start or 1), 1)
        last = min((end or doc.page_count), doc.page_count)
        matrix = fitz.Matrix(zoom, zoom)
        for page_index in range(first - 1, last):
            page_number = page_index + 1
            image_path = output_dir / f"page_{page_number:04d}.jpg"
            pix = doc.load_page(page_index).get_pixmap(matrix=matrix, alpha=False)
            pix.save(image_path)
            rendered.append((page_number, image_path))
    return rendered


def normalize_text(text: str) -> str:
    text = repair_pdf_glyph_text(text)
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = remove_quark_line_number_artifacts(text)
    text = re.sub(r"[ \t]+\n", "\n", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def remove_quark_line_number_artifacts(text: str) -> str:
    lines = text.splitlines()
    cleaned: list[str] = []
    for line in lines:
        stripped = line.strip()
        if re.fullmatch(r"\d{1,3}", stripped):
            continue
        if re.fullmatch(r"\d{1,3}\s*/\s*[^/]+", stripped):
            continue
        cleaned.append(line)
    return "\n".join(cleaned)


def split_text(text: str, max_chars: int, overlap: int) -> list[str]:
    if max_chars <= 0:
        raise ValueError("--chunk-size must be greater than 0")
    if overlap < 0 or overlap >= max_chars:
        raise ValueError("--chunk-overlap must be >= 0 and smaller than --chunk-size")
    if not text.strip():
        return []

    paragraphs = [p.strip() for p in text.split("\n\n") if p.strip()]
    chunks: list[str] = []
    current = ""

    for paragraph in paragraphs:
        candidate = f"{current}\n\n{paragraph}".strip() if current else paragraph
        if len(candidate) <= max_chars:
            current = candidate
            continue

        if current:
            chunks.append(current)
            current = current[-overlap:].strip() if overlap else ""

        while len(paragraph) > max_chars:
            piece = paragraph[:max_chars].strip()
            chunks.append(piece)
            prefix = piece[-overlap:].strip() if overlap else ""
            paragraph = f"{prefix} {paragraph[max_chars:]}".strip()

        current = f"{current}\n\n{paragraph}".strip() if current else paragraph

    if current:
        chunks.append(current)

    return chunks


class Translator:
    def __init__(self, config: TranslationConfig) -> None:
        api_key = os.getenv("OPENAI_API_KEY") or os.getenv("DEEPSEEK_API_KEY", "")
        if not api_key:
            raise RuntimeError("OPENAI_API_KEY or DEEPSEEK_API_KEY is required unless --dry-run is used.")

        self.client = OpenAI(
            api_key=api_key,
            base_url=os.getenv("OPENAI_BASE_URL") or ("https://api.deepseek.com" if os.getenv("DEEPSEEK_API_KEY") else None),
            timeout=float(os.getenv("OPENAI_TIMEOUT", "180")),
            max_retries=int(os.getenv("OPENAI_MAX_RETRIES", "2")),
        )
        self.model = os.getenv("OPENAI_MODEL") or os.getenv("DEEPSEEK_MODEL") or "deepseek-chat"
        self.config = config

    def translate(self, text: str) -> str:
        prompt = self.config.user_prompt_template.format(
            source_language=self.config.source_language,
            target_language=self.config.target_language,
            text=text,
        )
        attempts = int(os.getenv("TRANSLATION_MAX_ATTEMPTS", "6"))
        last_error: Exception | None = None
        for attempt in range(1, attempts + 1):
            try:
                response = self.client.chat.completions.create(
                    model=self.model,
                    temperature=self.config.temperature,
                    messages=[
                        {"role": "system", "content": self.config.system_prompt},
                        {"role": "user", "content": prompt},
                    ],
                )
                break
            except Exception as exc:
                last_error = exc
                if attempt >= attempts:
                    raise
                wait_seconds = min(90, 5 * attempt * attempt)
                print(f"Translation request failed, retrying in {wait_seconds}s ({attempt}/{attempts}): {type(exc).__name__}")
                time.sleep(wait_seconds)
        else:
            raise last_error or RuntimeError("Translation request failed")
        content = response.choices[0].message.content
        return content.strip() if content else ""


def translate_page(page_number: int, text: str, translator: Translator | None, config: TranslationConfig) -> str:
    chunks = split_text(text, max_chars=config.chunk_size, overlap=config.chunk_overlap)
    if not chunks:
        return f"<!-- Page {page_number}: no extractable text -->\n"

    translated_chunks = []
    for chunk_index, chunk in enumerate(chunks, start=1):
        if translator is None:
            translated = chunk
        else:
            translated = translator.translate(chunk)
        translated_chunks.append(f"<!-- page {page_number}, chunk {chunk_index}/{len(chunks)} -->\n\n{translated}")

    return "\n\n".join(translated_chunks).strip() + "\n"


def write_manifest(path: Path, files: list[str]) -> None:
    path.write_text(json.dumps(files, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def main() -> int:
    load_dotenv()

    parser = argparse.ArgumentParser(description="Extract and translate a book PDF into paged Markdown files.")
    parser.add_argument("--pdf", required=True, type=Path, help="Input PDF path")
    parser.add_argument("--config", default=Path("config/translation.json"), type=Path)
    parser.add_argument("--source-dir", default=Path("source_full"), type=Path)
    parser.add_argument("--translated-dir", default=Path("translated"), type=Path)
    parser.add_argument("--manifest", default=Path("manifest.json"), type=Path)
    parser.add_argument("--start", type=int, default=None, help="First page number, 1-based")
    parser.add_argument("--end", type=int, default=None, help="Last page number, 1-based")
    parser.add_argument("--source-language", default=None)
    parser.add_argument("--target-language", default=None)
    parser.add_argument("--chunk-size", type=int, default=None)
    parser.add_argument("--chunk-overlap", type=int, default=None)
    parser.add_argument("--dry-run", action="store_true", help="Extract and write source text without API calls")
    parser.add_argument("--overwrite", action="store_true", help="Overwrite existing translated page files")
    parser.add_argument("--page-images-dir", default=Path("images/pages"), type=Path)
    parser.add_argument("--include-page-images", action="store_true", help="Render original PDF pages and include them as reference images")
    parser.add_argument("--image-zoom", type=float, default=2.0, help="Original page image render scale")
    args = parser.parse_args()

    pdf_path = args.pdf.resolve()
    if not pdf_path.exists():
        raise FileNotFoundError(f"PDF not found: {pdf_path}")

    config = load_config(args.config)
    config = TranslationConfig(
        source_language=args.source_language or config.source_language,
        target_language=args.target_language or config.target_language,
        chunk_size=args.chunk_size or config.chunk_size,
        chunk_overlap=args.chunk_overlap if args.chunk_overlap is not None else config.chunk_overlap,
        temperature=config.temperature,
        system_prompt=config.system_prompt,
        user_prompt_template=config.user_prompt_template,
    )

    args.source_dir.mkdir(parents=True, exist_ok=True)
    args.translated_dir.mkdir(parents=True, exist_ok=True)
    Path("output").mkdir(exist_ok=True)

    pages = extract_pdf_pages(pdf_path, args.start, args.end)
    rendered_pages: dict[int, Path] = {}
    if args.include_page_images:
        rendered_pages = dict(
            render_pdf_pages(
                pdf_path=pdf_path,
                output_dir=args.page_images_dir,
                start=args.start,
                end=args.end,
                zoom=args.image_zoom,
            )
        )
    if not any(text for _, text in pages):
        raise RuntimeError("No text was extracted. This may be a scanned PDF; run OCR first.")

    translator = None if args.dry_run else Translator(config)
    manifest_files: list[str] = []
    translated_pages = 0
    skipped_pages = 0

    for page_number, text in tqdm(pages, desc="Preparing" if args.dry_run else "Translating"):
        raw_name = f"page_{page_number:04d}.txt"
        md_name = f"page_{page_number:04d}.md"
        raw_path = args.source_dir / raw_name
        md_path = args.translated_dir / md_name

        raw_path.write_text(text + "\n", encoding="utf-8")
        manifest_files.append(md_name)

        if md_path.exists() and not args.overwrite:
            skipped_pages += 1
            continue

        heading = f"# 第 {page_number} 页\n\n"
        page_image = ""
        if page_number in rendered_pages:
            image_name = rendered_pages[page_number].name
            page_image = f"![原书第 {page_number} 页]({image_name})\n\n"
        body = translate_page(page_number, text, translator, config)
        md_path.write_text(heading + page_image + body, encoding="utf-8")
        translated_pages += 1

    write_manifest(args.manifest, manifest_files)
    run_info = {
        "pdf": str(pdf_path),
        "pages_seen": len(pages),
        "pages_written": translated_pages,
        "pages_skipped": skipped_pages,
        "dry_run": args.dry_run,
        "source_language": config.source_language,
        "target_language": config.target_language,
        "chunk_size": config.chunk_size,
        "chunk_overlap": config.chunk_overlap,
    }
    Path("translation_run.json").write_text(json.dumps(run_info, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

    print(json.dumps(run_info, ensure_ascii=False, indent=2))
    print(f"Manifest written to: {args.manifest.resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
