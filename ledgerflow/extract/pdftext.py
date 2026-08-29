"""Thin wrapper over pdfplumber that keeps word geometry available.

Bank statements are laid out in columns, and the only reliable way to tell a
deposit from a withdrawal is which column an amount sits in. So text is kept
together with x-positions rather than flattened to strings too early.
"""

from __future__ import annotations

from dataclasses import dataclass, field

import pdfplumber


@dataclass
class Word:
    text: str
    x0: float
    x1: float
    top: float
    size: float = 0.0

    @property
    def mid(self) -> float:
        return (self.x0 + self.x1) / 2


@dataclass
class Line:
    """One visual line of a page, with its words in reading order."""

    page: int
    top: float
    words: list[Word]

    @property
    def text(self) -> str:
        return " ".join(w.text for w in self.words)

    @property
    def size(self) -> float:
        """Largest glyph size on the line, used to spot headings such as a vendor name."""
        return max((w.size for w in self.words), default=0.0)


@dataclass
class Page:
    number: int
    lines: list[Line]
    tables: list[list[list[str | None]]] = field(default_factory=list)
    width: float = 0.0

    @property
    def text(self) -> str:
        return "\n".join(line.text for line in self.lines)


@dataclass
class Document:
    path: str
    pages: list[Page]
    has_text_layer: bool

    @property
    def text(self) -> str:
        return "\n".join(page.text for page in self.pages)


def _group_lines(words: list[dict], page_number: int, tolerance: float = 3.0) -> list[Line]:
    """Group words into visual lines by their vertical position."""
    lines: list[Line] = []
    for word in sorted(words, key=lambda w: (round(w["top"], 1), w["x0"])):
        item = Word(
            text=word["text"],
            x0=word["x0"],
            x1=word["x1"],
            top=word["top"],
            size=float(word.get("size") or 0),
        )
        if lines and abs(lines[-1].top - item.top) <= tolerance:
            lines[-1].words.append(item)
        else:
            lines.append(Line(page=page_number, top=item.top, words=[item]))
    for line in lines:
        line.words.sort(key=lambda w: w.x0)
    return lines


def read_pdf(path: str) -> Document:
    """Read a PDF into pages of positioned words plus any ruled tables.

    A PDF with no text layer is a scan. It is reported as such rather than
    returning empty pages, so the caller can route it to visual review instead
    of silently producing nothing.
    """
    pages: list[Page] = []
    with pdfplumber.open(path) as pdf:
        for index, page in enumerate(pdf.pages, start=1):
            words = page.extract_words(
                    use_text_flow=False, keep_blank_chars=False, extra_attrs=["size"]
                ) or []
            try:
                tables = page.extract_tables() or []
            except Exception:
                tables = []
            pages.append(
                Page(
                    number=index,
                    lines=_group_lines(words, index),
                    tables=tables,
                    width=float(page.width or 0),
                )
            )
    has_text = any(page.lines for page in pages)
    return Document(path=path, pages=pages, has_text_layer=has_text)
