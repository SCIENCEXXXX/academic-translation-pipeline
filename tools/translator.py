from __future__ import annotations

import os
from dataclasses import dataclass

from openai import OpenAI

try:
    from .terminology import glossary_lines
except ImportError:
    from terminology import glossary_lines


BOOK_TRANSLATION_SYSTEM_PROMPT = f"""
You are translating an Olympic weightlifting training book for mainland Chinese coaches and athletes.
Write like a formal Chinese sports training textbook, not like line-by-line notes.

Rules:
- Translate naturally into Simplified Chinese.
- Rebuild paragraphs according to Chinese reading habits.
- Preserve exercise numbers.
- Put action names and explanations into clear hierarchy.
- Split long English sentences into shorter Chinese sentences when helpful.
- Do not omit technical meaning.
- Prefer these glossary terms whenever applicable:
{chr(10).join(glossary_lines())}
""".strip()


@dataclass
class TranslatorConfig:
    model: str = "deepseek-chat"
    base_url: str = "https://api.deepseek.com"
    temperature: float = 0.2


class BookTranslator:
    def __init__(self, config: TranslatorConfig | None = None) -> None:
        self.config = config or TranslatorConfig(
            model=os.getenv("OPENAI_MODEL", "deepseek-chat"),
            base_url=os.getenv("OPENAI_BASE_URL", "https://api.deepseek.com"),
        )
        api_key = os.getenv("OPENAI_API_KEY", "")
        if not api_key:
            raise RuntimeError("OPENAI_API_KEY is required for translation calls.")
        self.client = OpenAI(api_key=api_key, base_url=self.config.base_url)

    def translate_page(self, text: str) -> str:
        response = self.client.chat.completions.create(
            model=self.config.model,
            temperature=self.config.temperature,
            messages=[
                {"role": "system", "content": BOOK_TRANSLATION_SYSTEM_PROMPT},
                {"role": "user", "content": text},
            ],
        )
        content = response.choices[0].message.content
        return content.strip() if content else ""
