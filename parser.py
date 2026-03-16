"""
Парсер входного текста → список структурных элементов документа.
Поддерживает Markdown-подобный синтаксис + специальные теги ГОСТ.
"""

import re
from dataclasses import dataclass, field
from typing import List, Optional


# ─────────────────────────────────────────────
#  Типы элементов
# ─────────────────────────────────────────────
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
    caption: str

@dataclass
class TableElement:
    rows: List[List[str]]
    caption: str
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
class SpecialSection:
    """Реферат, Содержание, Введение и т.д."""
    title: str
    paragraphs: List[str] = field(default_factory=list)

@dataclass
class FormulaElement:
    text: str
    number: str = ""


DocElement = (Heading | Paragraph | FigureRef | TableElement |
              ListItem | PageBreak | SpecialSection | FormulaElement)


# ─────────────────────────────────────────────
#  Парсер
# ─────────────────────────────────────────────
class GostTextParser:

    # Регулярки
    RE_H1      = re.compile(r'^#\s+(.+)$')
    RE_H2      = re.compile(r'^##\s+(.+)$')
    RE_H3      = re.compile(r'^###\s+(.+)$')
    RE_FIGURE  = re.compile(r'^\[рисунок\s*:\s*([^\|]+)\|([^\]]+)\]', re.I)
    RE_FIGURE2 = re.compile(r'^\[рисунок\s*:\s*([^\]]+)\]', re.I)
    RE_TABLE_S = re.compile(r'^\[таблица\]', re.I)
    RE_TABLE_E = re.compile(r'^\[/таблица\s*(?:\|\s*(.+))?\]', re.I)
    RE_LIST_U  = re.compile(r'^[-*•]\s+(.+)$')
    RE_LIST_O  = re.compile(r'^(\d+)[.)]\s+(.+)$')
    RE_FORMULA = re.compile(r'^\[формула\s*:\s*([^\|]+)(?:\|([^\]]+))?\]', re.I)
    RE_BREAK   = re.compile(r'^---+$')

    # Заголовки специальных разделов (без нумерации)
    SPECIAL = {
        "РЕФЕРАТ", "СОДЕРЖАНИЕ", "ОПРЕДЕЛЕНИЯ",
        "ОБОЗНАЧЕНИЯ И СОКРАЩЕНИЯ", "ВВЕДЕНИЕ",
        "ЗАКЛЮЧЕНИЕ", "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ",
    }

    @staticmethod
    def _normalize(s: str) -> str:
        """Схлопывает множественные пробелы, убирает пробелы по краям,
        заменяет неразрывный пробел и табуляцию на обычный пробел."""
        s = s.replace('\u00a0', ' ').replace('\t', ' ')
        s = re.sub(r' {2,}', ' ', s)
        return s.strip()

    def parse(self, text: str) -> List:
        # Нормализация входного текста:
        # 1. Убираем лишние пустые строки (более двух подряд → одна пустая)
        text = re.sub(r'\n{3,}', '\n\n', text)
        lines = text.splitlines()
        elements = []
        i = 0
        table_rows = []
        in_table = False

        while i < len(lines):
            line = self._normalize(lines[i])

            # --- Пустая строка ---
            if not line.strip():
                i += 1
                continue

            # --- Таблица: начало ---
            if self.RE_TABLE_S.match(line):
                in_table = True
                table_rows = []
                i += 1
                continue

            # --- Таблица: конец ---
            if in_table and self.RE_TABLE_E.match(line):
                m = self.RE_TABLE_E.match(line)
                caption = m.group(1).strip() if m.group(1) else "Таблица"
                if table_rows:
                    elements.append(TableElement(
                        rows=table_rows,
                        caption=caption,
                        has_header=True
                    ))
                in_table = False
                i += 1
                continue

            # --- Таблица: строка данных ---
            if in_table:
                cells = [c.strip() for c in line.split("|")]
                # Убираем разделитель |---|---| (Markdown-таблица)
                if not all(re.match(r'^[-:]+$', c.replace(' ', '')) for c in cells if c):
                    table_rows.append(cells)
                i += 1
                continue

            # --- Разрыв страницы ---
            if self.RE_BREAK.match(line):
                elements.append(PageBreak())
                i += 1
                continue

            # --- Рисунок ---
            m = self.RE_FIGURE.match(line)
            if m:
                elements.append(FigureRef(
                    path=m.group(1).strip(),
                    caption=m.group(2).strip()
                ))
                i += 1
                continue

            m = self.RE_FIGURE2.match(line)
            if m:
                elements.append(FigureRef(
                    path=m.group(1).strip(),
                    caption=""
                ))
                i += 1
                continue

            # --- Формула ---
            m = self.RE_FORMULA.match(line)
            if m:
                elements.append(FormulaElement(
                    text=m.group(1).strip(),
                    number=m.group(2).strip() if m.group(2) else ""
                ))
                i += 1
                continue

            # --- Заголовок H1 ---
            m = self.RE_H1.match(line)
            if m:
                title = m.group(1).strip()
                upper = title.upper()
                if upper in self.SPECIAL or any(upper.startswith(s) for s in self.SPECIAL):
                    elements.append(Heading(level=1, text=title))
                else:
                    elements.append(Heading(level=1, text=title))
                i += 1
                continue

            # --- Заголовок H2 ---
            m = self.RE_H2.match(line)
            if m:
                elements.append(Heading(level=2, text=m.group(1).strip()))
                i += 1
                continue

            # --- Заголовок H3 ---
            m = self.RE_H3.match(line)
            if m:
                elements.append(Heading(level=3, text=m.group(1).strip()))
                i += 1
                continue

            # --- Маркированный список ---
            m = self.RE_LIST_U.match(line)
            if m:
                elements.append(ListItem(text=m.group(1).strip(), ordered=False))
                i += 1
                continue

            # --- Нумерованный список ---
            m = self.RE_LIST_O.match(line)
            if m:
                elements.append(ListItem(
                    text=m.group(2).strip(),
                    ordered=True,
                    number=int(m.group(1))
                ))
                i += 1
                continue

            # --- Обычный абзац (собираем многострочный) ---
            para_lines = [line]
            i += 1
            while i < len(lines):
                next_line = lines[i].rstrip()
                if not next_line.strip():
                    break
                # Если следующая строка — спецэлемент, останавливаемся
                if (self.RE_H1.match(next_line) or self.RE_H2.match(next_line) or
                        self.RE_H3.match(next_line) or self.RE_FIGURE.match(next_line) or
                        self.RE_TABLE_S.match(next_line) or self.RE_LIST_U.match(next_line) or
                        self.RE_LIST_O.match(next_line) or self.RE_BREAK.match(next_line)):
                    break
                para_lines.append(next_line)
                i += 1

            full_text = self._normalize(" ".join(para_lines))
            if full_text:
                elements.append(Paragraph(text=full_text))

        return elements


    def auto_number(self, elements: List) -> List:
        """Автоматически нумерует заголовки, рисунки, таблицы."""
        h1_count = 0
        h2_count = 0
        h3_count = 0
        fig_count = 0
        tbl_count = 0

        SPECIAL_UPPER = {
            "РЕФЕРАТ", "СОДЕРЖАНИЕ", "ОПРЕДЕЛЕНИЯ",
            "ОБОЗНАЧЕНИЯ И СОКРАЩЕНИЯ", "ВВЕДЕНИЕ",
            "ЗАКЛЮЧЕНИЕ", "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ",
        }

        for el in elements:
            if isinstance(el, Heading):
                text_upper = el.text.upper().strip()
                is_special = text_upper in SPECIAL_UPPER or any(
                    text_upper.startswith(s) for s in SPECIAL_UPPER
                )
                if el.level == 1:
                    if not is_special:
                        h1_count += 1
                        h2_count = 0
                        h3_count = 0
                        el.number = str(h1_count)
                    else:
                        el.number = ""
                elif el.level == 2:
                    h2_count += 1
                    h3_count = 0
                    el.number = f"{h1_count}.{h2_count}"
                elif el.level == 3:
                    h3_count += 1
                    el.number = f"{h1_count}.{h2_count}.{h3_count}"

            elif isinstance(el, FigureRef):
                fig_count += 1
                el._number = fig_count

            elif isinstance(el, TableElement):
                tbl_count += 1
                el._number = tbl_count

        return elements
