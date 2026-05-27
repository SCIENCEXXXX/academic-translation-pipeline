from __future__ import annotations

import argparse
import json
import math
import os
import re
from concurrent.futures import ProcessPoolExecutor, as_completed
from pathlib import Path
from typing import Any

import fitz
import numpy as np
from PIL import Image
from rapidocr_onnxruntime import RapidOCR
from tqdm import tqdm


_OCR: RapidOCR | None = None


def init_worker() -> None:
    global _OCR
    _OCR = RapidOCR()


def pixmap_to_array(pix: fitz.Pixmap) -> np.ndarray:
    channels = pix.n
    arr = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, channels)
    if channels == 4:
        arr = arr[:, :, :3]
    return arr


def box_rect(points: list[list[float]]) -> tuple[float, float, float, float]:
    xs = [p[0] for p in points]
    ys = [p[1] for p in points]
    return min(xs), min(ys), max(xs), max(ys)


def is_page_number(text: str, y0: float, y1: float, height: int) -> bool:
    stripped = text.strip()
    if not re.fullmatch(r"\d{1,4}|[ivxlcdm]{1,8}", stripped, flags=re.I):
        return False
    cy = (y0 + y1) / 2
    return cy < height * 0.08 or cy > height * 0.92


def clean_ocr_line(text: str) -> str:
    text = text.replace("\u00ad", "")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def sorted_ocr_lines(result: list[Any], width: int, height: int, min_conf: float) -> list[str]:
    boxes: list[dict[str, Any]] = []
    for item in result or []:
        points, text, conf = item
        text = clean_ocr_line(str(text))
        if not text or float(conf) < min_conf:
            continue
        x0, y0, x1, y1 = box_rect(points)
        if is_page_number(text, y0, y1, height):
            continue
        boxes.append({"text": text, "x0": x0, "y0": y0, "x1": x1, "y1": y1, "cx": (x0 + x1) / 2, "w": x1 - x0})
    if not boxes:
        return []

    left = [b for b in boxes if b["cx"] < width * 0.48 and not re.fullmatch(r"\d{1,4}", b["text"])]
    right = [b for b in boxes if b["cx"] > width * 0.52 and not re.fullmatch(r"\d{1,4}", b["text"])]
    wide = [b for b in boxes if b["w"] > width * 0.58]
    two_columns = len(left) >= 8 and len(right) >= 8 and len(wide) < max(12, len(boxes) * 0.28)

    if not two_columns:
        ordered = sorted(boxes, key=lambda b: (round(b["y0"] / 8), b["x0"]))
    else:
        top = [b for b in boxes if b["y0"] < height * 0.10 or b["w"] > width * 0.60 and b["y0"] < height * 0.25]
        bottom_wide = [b for b in boxes if b["w"] > width * 0.60 and b not in top]
        column_boxes = [b for b in boxes if b not in top and b not in bottom_wide]
        left_col = [b for b in column_boxes if b["cx"] < width * 0.52]
        right_col = [b for b in column_boxes if b["cx"] >= width * 0.52]
        ordered = (
            sorted(top, key=lambda b: (round(b["y0"] / 8), b["x0"]))
            + sorted(left_col, key=lambda b: (round(b["y0"] / 8), b["x0"]))
            + sorted(right_col, key=lambda b: (round(b["y0"] / 8), b["x0"]))
            + sorted(bottom_wide, key=lambda b: (round(b["y0"] / 8), b["x0"]))
        )
    return [b["text"] for b in ordered]


def caption_entries(result: list[Any], width: int, height: int) -> list[dict[str, Any]]:
    entries: list[dict[str, Any]] = []
    for item in result or []:
        points, text, conf = item
        text = clean_ocr_line(str(text))
        if float(conf) < 0.55:
            continue
        if not re.match(r"^(?:FIGURA|Figura|Fig\.|TABELUL|Tabelul|TABEL|Tabel)\s+\d+", text):
            continue
        x0, y0, x1, y1 = box_rect(points)
        is_table = text.lower().startswith(("tabel", "tabelul"))
        if is_table:
            crop_y0 = max(0, y1 - 4)
            crop_y1 = min(height, y1 + height * 0.32)
        else:
            crop_y0 = max(0, y0 - height * 0.34)
            crop_y1 = max(crop_y0 + 40, y0 - 4)
        if x1 - x0 > width * 0.55:
            crop_x0, crop_x1 = width * 0.06, width * 0.94
        elif (x0 + x1) / 2 < width / 2:
            crop_x0, crop_x1 = width * 0.05, width * 0.52
        else:
            crop_x0, crop_x1 = width * 0.48, width * 0.95
        if crop_y1 - crop_y0 < 60:
            continue
        entries.append(
            {
                "caption": text,
                "is_table": is_table,
                "clip": [int(crop_x0), int(crop_y0), int(crop_x1), int(crop_y1)],
            }
        )
    return entries


def save_crop(image: np.ndarray, clip: list[int], path: Path) -> bool:
    x0, y0, x1, y1 = clip
    h, w = image.shape[:2]
    x0, y0 = max(0, x0), max(0, y0)
    x1, y1 = min(w, x1), min(h, y1)
    if x1 - x0 < 80 or y1 - y0 < 60:
        return False
    crop = image[y0:y1, x0:x1]
    # Avoid saving nearly blank crops.
    if crop.size == 0 or float(np.std(crop)) < 3.5:
        return False
    Image.fromarray(crop).save(path)
    return True


def ocr_one_page(args: tuple[str, int, float, str, str, bool, float]) -> dict[str, Any]:
    pdf_path, page_number, zoom, output_dir, figures_dir, overwrite, min_conf = args
    source_path = Path(output_dir) / f"page_{page_number:04d}.md"
    if source_path.exists() and not overwrite:
        return {"page": page_number, "source": str(source_path), "figures": [], "skipped": True}

    global _OCR
    if _OCR is None:
        _OCR = RapidOCR()

    with fitz.open(pdf_path) as doc:
        page = doc.load_page(page_number - 1)
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
    image = pixmap_to_array(pix)
    result, elapsed = _OCR(image)
    lines = sorted_ocr_lines(result or [], pix.width, pix.height, min_conf)

    source_path.parent.mkdir(parents=True, exist_ok=True)
    text = "\n".join(lines).strip()
    source_path.write_text(
        f"# OCR 第 {page_number} 页\n\n<!-- ocr elapsed={elapsed} -->\n\n{text}\n",
        encoding="utf-8",
    )

    figures: list[dict[str, Any]] = []
    fig_dir = Path(figures_dir)
    fig_dir.mkdir(parents=True, exist_ok=True)
    for index, entry in enumerate(caption_entries(result or [], pix.width, pix.height), start=1):
        fig_path = fig_dir / f"page_{page_number:04d}_figure_{index:02d}.png"
        if save_crop(image, entry["clip"], fig_path):
            figures.append({"page": page_number, "path": str(fig_path.resolve()), "caption": entry["caption"]})
    return {"page": page_number, "source": str(source_path), "figures": figures, "skipped": False}


def main() -> int:
    parser = argparse.ArgumentParser(description="OCR an image-only scanned PDF into page Markdown files.")
    parser.add_argument("--pdf", required=True, type=Path)
    parser.add_argument("--output-dir", required=True, type=Path)
    parser.add_argument("--figures-dir", required=True, type=Path)
    parser.add_argument("--manifest", required=True, type=Path)
    parser.add_argument("--start", type=int, default=1)
    parser.add_argument("--end", type=int, default=None)
    parser.add_argument("--zoom", type=float, default=1.15)
    parser.add_argument("--workers", type=int, default=1)
    parser.add_argument("--min-confidence", type=float, default=0.48)
    parser.add_argument("--overwrite", action="store_true")
    args = parser.parse_args()

    with fitz.open(args.pdf) as doc:
        last = min(args.end or doc.page_count, doc.page_count)
    pages = list(range(max(args.start, 1), last + 1))
    tasks = [
        (str(args.pdf.resolve()), page, args.zoom, str(args.output_dir), str(args.figures_dir), args.overwrite, args.min_confidence)
        for page in pages
    ]

    results: list[dict[str, Any]] = []
    if args.workers <= 1:
        init_worker()
        for task in tqdm(tasks, desc="OCR"):
            results.append(ocr_one_page(task))
    else:
        with ProcessPoolExecutor(max_workers=args.workers, initializer=init_worker) as pool:
            futures = [pool.submit(ocr_one_page, task) for task in tasks]
            for future in tqdm(as_completed(futures), total=len(futures), desc="OCR"):
                results.append(future.result())

    results.sort(key=lambda item: item["page"])
    all_figures = [fig for result in results for fig in result.get("figures", [])]
    manifest = {
        "pdf": str(args.pdf.resolve()),
        "pages": len(results),
        "source_files": [result["source"] for result in results],
        "figures": all_figures,
    }
    args.manifest.parent.mkdir(parents=True, exist_ok=True)
    args.manifest.write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    (args.figures_dir / "figures_manifest.json").write_text(
        json.dumps(all_figures, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    print(json.dumps({"pages": len(results), "figures": len(all_figures), "manifest": str(args.manifest)}, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
