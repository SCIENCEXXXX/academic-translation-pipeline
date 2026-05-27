from __future__ import annotations

import argparse
import json
from dataclasses import dataclass, asdict
from pathlib import Path

import fitz


@dataclass
class ExtractedImage:
    page_number: int
    image_index: int
    path: str
    width: int
    height: int


def extract_pdf_images(
    pdf_path: Path,
    output_dir: Path,
    start: int | None = None,
    end: int | None = None,
    min_width: int = 80,
    min_height: int = 80,
) -> list[ExtractedImage]:
    output_dir.mkdir(parents=True, exist_ok=True)
    images: list[ExtractedImage] = []

    with fitz.open(pdf_path) as doc:
        first = max(start or 1, 1)
        last = min(end or doc.page_count, doc.page_count)
        seen: set[tuple[int, int]] = set()

        for page_index in range(first - 1, last):
            page_number = page_index + 1
            page = doc.load_page(page_index)
            for image_index, image in enumerate(page.get_images(full=True), start=1):
                xref = image[0]
                if (page_number, xref) in seen:
                    continue
                seen.add((page_number, xref))
                base = doc.extract_image(xref)
                width = int(base.get("width", 0))
                height = int(base.get("height", 0))
                if width < min_width or height < min_height:
                    continue
                ext = base.get("ext", "png")
                image_path = output_dir / f"page_{page_number:04d}_img_{image_index:02d}.{ext}"
                image_path.write_bytes(base["image"])
                images.append(
                    ExtractedImage(
                        page_number=page_number,
                        image_index=image_index,
                        path=str(image_path),
                        width=width,
                        height=height,
                    )
                )
    return images


def main() -> int:
    parser = argparse.ArgumentParser(description="Extract embedded images from a PDF.")
    parser.add_argument("--input", required=True, type=Path)
    parser.add_argument("--output-dir", required=True, type=Path)
    parser.add_argument("--manifest", type=Path)
    parser.add_argument("--start", type=int)
    parser.add_argument("--end", type=int)
    args = parser.parse_args()

    images = extract_pdf_images(args.input, args.output_dir, args.start, args.end)
    if args.manifest:
        args.manifest.parent.mkdir(parents=True, exist_ok=True)
        args.manifest.write_text(
            json.dumps([asdict(image) for image in images], ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
