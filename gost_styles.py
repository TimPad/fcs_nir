"""
ГОСТ 7.32-2017 — все константы стилей для отчёта о НИР
"""

from docx.shared import Pt, Mm, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT

# ─────────────────────────────────────────────
#  СТРАНИЦА (ГОСТ 7.32-2017, п. 6.1)
# ─────────────────────────────────────────────
PAGE = {
    "width":  Mm(210),
    "height": Mm(297),
    "margin_left":   Mm(30),
    "margin_right":  Mm(10),
    "margin_top":    Mm(20),
    "margin_bottom": Mm(20),
    "orientation": WD_ORIENT.PORTRAIT,
}

# ─────────────────────────────────────────────
#  ОСНОВНОЙ ТЕКСТ (ГОСТ 7.32-2017, п. 6.3)
# ─────────────────────────────────────────────
BODY = {
    "font_name":       "Times New Roman",
    "font_size":       Pt(14),
    "line_spacing":    Pt(21),          # полуторный ≈ 14 * 1.5
    "first_line_indent": Cm(1.25),
    "alignment":       WD_ALIGN_PARAGRAPH.JUSTIFY,
    "space_before":    Pt(0),
    "space_after":     Pt(0),
}

# ─────────────────────────────────────────────
#  ЗАГОЛОВКИ (ГОСТ 7.32-2017, п. 6.5)
# ─────────────────────────────────────────────
HEADING1 = {
    "font_name":    "Times New Roman",
    "font_size":    Pt(14),
    "bold":         True,
    "alignment":    WD_ALIGN_PARAGRAPH.CENTER,
    "space_before": Pt(0),
    "space_after":  Pt(0),
    "caps":         True,           # ЗАГЛАВНЫЕ БУКВЫ
    "page_break_before": True,
    "first_line_indent": Cm(0),
}

HEADING2 = {
    "font_name":    "Times New Roman",
    "font_size":    Pt(14),
    "bold":         True,
    "alignment":    WD_ALIGN_PARAGRAPH.LEFT,
    "space_before": Pt(0),
    "space_after":  Pt(0),
    "first_line_indent": Cm(1.25),
}

HEADING3 = {
    "font_name":    "Times New Roman",
    "font_size":    Pt(14),
    "bold":         True,
    "italic":       True,
    "alignment":    WD_ALIGN_PARAGRAPH.LEFT,
    "space_before": Pt(0),
    "space_after":  Pt(0),
    "first_line_indent": Cm(1.25),
}

# ─────────────────────────────────────────────
#  ПОДПИСИ К РИСУНКАМ (ГОСТ 7.32-2017, п. 6.12)
# ─────────────────────────────────────────────
FIGURE_CAPTION = {
    "font_name":    "Times New Roman",
    "font_size":    Pt(14),
    "bold":         False,
    "alignment":    WD_ALIGN_PARAGRAPH.CENTER,
    "space_before": Pt(6),
    "space_after":  Pt(6),
    "first_line_indent": Cm(0),
}

# ─────────────────────────────────────────────
#  ПОДПИСИ К ТАБЛИЦАМ (ГОСТ 7.32-2017, п. 6.13)
# ─────────────────────────────────────────────
TABLE_CAPTION = {
    "font_name":    "Times New Roman",
    "font_size":    Pt(14),
    "bold":         False,
    "alignment":    WD_ALIGN_PARAGRAPH.LEFT,
    "space_before": Pt(6),
    "space_after":  Pt(3),
    "first_line_indent": Cm(0),
}

TABLE_CELL = {
    "font_name": "Times New Roman",
    "font_size": Pt(12),
}

# ─────────────────────────────────────────────
#  КОЛОНТИТУЛЫ (ГОСТ 7.32-2017, п. 6.4)
#  Нижний: номер страницы по центру, 12 пт
# ─────────────────────────────────────────────
FOOTER = {
    "font_name": "Times New Roman",
    "font_size": Pt(12),
    "alignment": WD_ALIGN_PARAGRAPH.CENTER,
}

# ─────────────────────────────────────────────
#  РЕФЕРАТ / СОДЕРЖАНИЕ / СПИСОК ИСТОЧНИКОВ
# ─────────────────────────────────────────────
SECTION_TITLES = {
    "РЕФЕРАТ":              True,
    "СОДЕРЖАНИЕ":           True,
    "ОПРЕДЕЛЕНИЯ":          True,
    "ОБОЗНАЧЕНИЯ И СОКРАЩЕНИЯ": True,
    "ВВЕДЕНИЕ":             True,
    "ЗАКЛЮЧЕНИЕ":           True,
    "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ": True,
    "ПРИЛОЖЕНИЕ":           True,
}

# ─────────────────────────────────────────────
#  СПИСОК (маркированный и нумерованный)
# ─────────────────────────────────────────────
LIST_ITEM = {
    "font_name":  "Times New Roman",
    "font_size":  Pt(14),
    "alignment":  WD_ALIGN_PARAGRAPH.JUSTIFY,
    "left_indent": Cm(1.25),
    "first_line_indent": Cm(-0.5),
    "space_before": Pt(0),
    "space_after":  Pt(0),
}
