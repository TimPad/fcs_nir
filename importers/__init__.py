"""
Импортёры — преобразуют входной файл/текст в список DocElement.

Все импортёры возвращают ImportResult с полями:
- elements: List[DocElement]
- images:   Dict[str, bytes]   (имя_файла → содержимое)

Реестр выбирает импортёр по расширению файла.
"""

from dataclasses import dataclass, field
from typing import Callable, Dict, List, Optional

from elements import DocElement
from bibliography import Reference


@dataclass
class ImportResult:
    elements: List[DocElement] = field(default_factory=list)
    images: Dict[str, bytes] = field(default_factory=dict)
    # Источники, извлечённые из самого файла (например, \begin{thebibliography})
    references: List[Reference] = field(default_factory=list)
    # Сообщения для пользователя (warnings) — не ошибки
    warnings: List[str] = field(default_factory=list)


# Импортёр: (data, filename) -> ImportResult
Importer = Callable[[bytes, Optional[str]], ImportResult]


# ──────────────────────────────────────────
#  Реестр импортёров (по расширению)
# ──────────────────────────────────────────
def _ext(filename: str) -> str:
    name = (filename or "").lower().strip()
    if "." not in name:
        return ""
    return name.rsplit(".", 1)[1]


def get_importer_for_filename(filename: str) -> Optional[Importer]:
    """Подбирает импортёр по расширению файла. Возвращает None, если нет."""
    ext = _ext(filename)
    if ext == "docx":
        from importers.docx_imp import import_docx
        return import_docx
    if ext in ("md", "markdown", "txt"):
        from importers.markdown_imp import import_markdown
        return import_markdown
    if ext == "tex":
        from importers.latex_imp import import_latex
        return import_latex
    if ext in ("xlsx", "xls"):
        from importers.xlsx_imp import import_xlsx
        return import_xlsx
    if ext == "csv":
        from importers.xlsx_imp import import_csv
        return import_csv
    return None


def list_supported_extensions() -> List[str]:
    return ["docx", "md", "markdown", "txt", "tex", "xlsx", "csv"]
