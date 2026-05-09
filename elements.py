"""
Единый набор типов структурных элементов отчёта.

Все парсеры/импортёры (Markdown, DOCX, LaTeX, XLSX) возвращают список
этих элементов; билдер DOCX рендерит документ из них.
"""

from dataclasses import dataclass, field
from typing import List, Optional, Union


@dataclass
class Heading:
    level: int          # 1, 2, 3
    text: str
    number: str = ""    # "1", "1.1", "1.1.1"


@dataclass
class Paragraph:
    text: str


@dataclass
class FigureRef:
    path: str           # имя файла или placeholder
    caption: str = ""
    image_data: Optional[bytes] = None


@dataclass
class TableElement:
    rows: List[List[str]] = field(default_factory=list)
    caption: str = ""
    has_header: bool = True


@dataclass
class ListItem:
    text: str
    ordered: bool = False
    number: int = 1


@dataclass
class PageBreak:
    pass


@dataclass
class FormulaElement:
    text: str
    number: str = ""


DocElement = Union[
    Heading, Paragraph, FigureRef, TableElement,
    ListItem, PageBreak, FormulaElement,
]


# ─────────────────────────────────────────────
#  Cпециальные разделы (без сквозной нумерации)
# ─────────────────────────────────────────────
SPECIAL_HEADINGS = frozenset({
    "РЕФЕРАТ",
    "СОДЕРЖАНИЕ",
    "ОПРЕДЕЛЕНИЯ",
    "ОБОЗНАЧЕНИЯ И СОКРАЩЕНИЯ",
    "ВВЕДЕНИЕ",
    "ЗАКЛЮЧЕНИЕ",
    "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ",
    "СПИСОК ИСТОЧНИКОВ",
    "ПРИЛОЖЕНИЯ",
    "ПРИЛОЖЕНИЕ",
})


def is_special_heading(text: str) -> bool:
    """Проверяет, является ли заголовок специальным разделом ГОСТ."""
    upper = (text or "").upper().strip()
    if not upper:
        return False
    if upper in SPECIAL_HEADINGS:
        return True
    return any(upper.startswith(s) for s in SPECIAL_HEADINGS)


def auto_number(elements: List[DocElement]) -> List[DocElement]:
    """Сквозная нумерация: разделы, подразделы, рисунки, таблицы.

    Применяется ко всему документу (включая результат merge() из частей)
    — так нумерация остаётся непрерывной.

    H2/H3 получают номер только под нумерованным H1; иначе номер очищается.
    Изменяет элементы in-place и возвращает тот же список.
    """
    h1 = h2 = h3 = 0
    fig = tbl = 0
    in_numbered_section = False
    for el in elements:
        if isinstance(el, Heading):
            if el.level == 1:
                if is_special_heading(el.text):
                    el.number = ""
                    in_numbered_section = False
                    h2 = h3 = 0
                else:
                    h1 += 1
                    h2 = h3 = 0
                    el.number = str(h1)
                    in_numbered_section = True
            elif el.level == 2:
                if in_numbered_section:
                    h2 += 1
                    h3 = 0
                    el.number = f"{h1}.{h2}"
                else:
                    el.number = ""
            elif el.level == 3:
                if in_numbered_section and h2 > 0:
                    h3 += 1
                    el.number = f"{h1}.{h2}.{h3}"
                else:
                    el.number = ""
        elif isinstance(el, FigureRef):
            fig += 1
            el._number = fig  # type: ignore[attr-defined]
        elif isinstance(el, TableElement):
            tbl += 1
            el._number = tbl  # type: ignore[attr-defined]
    return elements
