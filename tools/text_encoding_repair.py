from __future__ import annotations


PDF_GLYPH_MAP = {
    "\x03": " ",
    "\x04": "A",
    "\x11": "B",
    "\x12": "C",
    "\x18": "D",
    "\x1c": "E",
    "&": "F",
    "'": "G",
    ",": "H",
    "/": "I",
    ":": "J",
    "<": "K",
    ">": "L",
    "D": "M",
    "E": "N",
    "K": "O",
    "W": "P",
    "Z": "R",
    "^": "S",
    "d": "T",
    "h": "U",
    "s": "V",
    "t": "W",
    "z": "Y",
    "Ă": "a",
    "ď": "b",
    "Đ": "c",
    "Ě": "d",
    "Ğ": "e",
    "Ĩ": "f",
    "Ő": "g",
    "Ś": "h",
    "ŝ": "i",
    "ũ": "j",
    "Ŭ": "k",
    "ů": "l",
    "ŵ": "m",
    "Ŷ": "n",
    "Ž": "o",
    "Ɖ": "p",
    "Ƌ": "q",
    "ƌ": "r",
    "Ɛ": "s",
    "ƚ": "t",
    "Ƶ": "u",
    "ǀ": "v",
    "ǁ": "w",
    "ǆ": "x",
    "Ǉ": "y",
    "ǌ": "z",
    "Ϭ": "0",
    "ϭ": "1",
    "Ϯ": "2",
    "ϯ": "3",
    "ϰ": "4",
    "ϱ": "5",
    "ϲ": "6",
    "ϳ": "7",
    "ϴ": "8",
    "ϵ": "9",
    "͕": ",",
    "͘": ".",
    "͗": ":",
    ";": "(",
    "Ϳ": ")",
    "Ͳ": "-",
    "ʹ": "'",
    "͛": "'",
    "ͬ": "/",
    "Ξ": "©",
    "й": "%",
}


def looks_like_pdf_glyph_text(text: str) -> bool:
    if not text:
        return False
    hits = sum(1 for ch in text if ch in PDF_GLYPH_MAP)
    latin_ext = sum(1 for ch in text if "\u0100" <= ch <= "\u03ff")
    return hits >= 8 and latin_ext / max(len(text), 1) > 0.08


def repair_pdf_glyph_text(text: str) -> str:
    if not looks_like_pdf_glyph_text(text):
        return text
    return "".join(PDF_GLYPH_MAP.get(ch, ch) for ch in text)
