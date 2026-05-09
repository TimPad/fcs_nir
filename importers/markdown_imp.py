"""Импортёр Markdown-подобной разметки (обёртка над GostTextParser)."""

from typing import Optional

from importers import ImportResult
from parser import GostTextParser


def import_markdown(data: bytes, filename: Optional[str] = None) -> ImportResult:
    text = data.decode("utf-8", errors="replace")
    return import_markdown_text(text)


def import_markdown_text(text: str) -> ImportResult:
    parser = GostTextParser()
    elements = parser.parse(text)
    return ImportResult(elements=elements, images={}, warnings=[])
