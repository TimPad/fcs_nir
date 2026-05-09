"""Импортёр DOCX (обёртка над GostDocxParser)."""

from typing import Optional

from importers import ImportResult
from docx_parser import GostDocxParser


def import_docx(data: bytes, filename: Optional[str] = None) -> ImportResult:
    parser = GostDocxParser()
    elements, images = parser.parse(data)
    return ImportResult(elements=elements, images=images, warnings=[])
