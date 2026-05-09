"""
LaTeX-импортёр.

Регекспный парсер LaTeX без внешних зависимостей. Поддерживает основные
конструкции, встречающиеся в типичном отчёте:

- \\section / \\subsection / \\subsubsection (+ \\section* и т.д.)
- \\paragraph как H3
- itemize / enumerate / description
- figure (\\includegraphics + \\caption)
- table (tabular + \\caption)
- \\cite{key} → метка [@key] в тексте
- thebibliography + \\bibitem → источники

Ограничения: вложенные среды, нестандартные пакеты, кастомные команды
не разворачиваются. Сложные таблицы (multirow, multicolumn) импортируются
как простые с потерей слияний ячеек.
"""

import re
from typing import List, Optional, Tuple

from importers import ImportResult
from elements import (
    Heading, Paragraph, FigureRef, TableElement,
    ListItem, FormulaElement,
)
from bibliography import Reference


# ──────────────────────────────────────────
#  Регекспы
# ──────────────────────────────────────────
RE_COMMENT = re.compile(r'(?<!\\)%.*$', re.M)
RE_BEGIN_DOC = re.compile(r'\\begin\s*\{document\}')
RE_END_DOC = re.compile(r'\\end\s*\{document\}')

RE_SECTION = re.compile(r'\\section\*?\s*\{(.+?)\}')
RE_SUBSECTION = re.compile(r'\\subsection\*?\s*\{(.+?)\}')
RE_SUBSUBSECTION = re.compile(r'\\subsubsection\*?\s*\{(.+?)\}')
RE_PARAGRAPH_CMD = re.compile(r'\\paragraph\s*\{(.+?)\}')

RE_ENV = re.compile(
    r'\\begin\s*\{(?P<name>[a-zA-Z*]+)\}(?P<opts>(?:\[[^\]]*\])*)\s*'
    r'(?P<body>.*?)\\end\s*\{(?P=name)\}',
    re.S,
)

RE_INCLUDE_GRAPHICS = re.compile(r'\\includegraphics(?:\[[^\]]*\])?\s*\{([^}]+)\}')
RE_CAPTION = re.compile(r'\\caption\s*\{(.+?)\}', re.S)
RE_LABEL = re.compile(r'\\label\s*\{[^}]*\}')

RE_BIBITEM = re.compile(r'\\bibitem\s*(?:\[[^\]]*\])?\s*\{([^}]+)\}', re.S)
RE_CITE = re.compile(r'\\cite[a-zA-Z]*\s*(?:\[[^\]]*\])?\s*\{([^}]+)\}')

RE_ITEM = re.compile(r'\\item\b\s*')
RE_DOUBLE_BACKSLASH = re.compile(r'\\\\(?:\s*\[[^\]]*\])?')

# «Шумовые» команды, которые выкидываем целиком (с аргументами)
RE_SIMPLE_NOISE = re.compile(
    r'\\(?:maketitle|tableofcontents|newpage|clearpage|pagebreak|noindent|hline|toprule|midrule|bottomrule|centering|small|large|footnotesize)\b\s*'
)
RE_LATEX_ARG_TEXT = {
    "textbf": re.compile(r'\\textbf\s*\{([^{}]+)\}'),
    "textit": re.compile(r'\\textit\s*\{([^{}]+)\}'),
    "emph": re.compile(r'\\emph\s*\{([^{}]+)\}'),
    "underline": re.compile(r'\\underline\s*\{([^{}]+)\}'),
    "texttt": re.compile(r'\\texttt\s*\{([^{}]+)\}'),
    "url": re.compile(r'\\url\s*\{([^{}]+)\}'),
    "href": re.compile(r'\\href\s*\{[^{}]+\}\s*\{([^{}]+)\}'),
}

# Простые символьные подстановки
LATEX_REPLACEMENTS = [
    (re.compile(r'\\&'), '&'),
    (re.compile(r'\\%'), '%'),
    (re.compile(r'\\\$'), '$'),
    (re.compile(r'\\#'), '#'),
    (re.compile(r'\\_'), '_'),
    (re.compile(r'\\textendash\s*'), '–'),
    (re.compile(r'\\textemdash\s*'), '—'),
    (re.compile(r'\\textquoteleft\s*'), '‘'),
    (re.compile(r'\\textquoteright\s*'), '’'),
    (re.compile(r'``'), '«'),
    (re.compile(r"''"), '»'),
    (re.compile(r'~'), ' '),
    (re.compile(r'---'), '—'),
    (re.compile(r'--'), '–'),
    (re.compile(r'\\ldots\b'), '…'),
    (re.compile(r'\\dots\b'), '…'),
]


# ──────────────────────────────────────────
#  Утилиты
# ──────────────────────────────────────────
def _strip_comments(text: str) -> str:
    return RE_COMMENT.sub("", text)


def _trim_to_document(text: str) -> str:
    """Оставляем содержимое между \\begin{document}…\\end{document}, если есть."""
    m_begin = RE_BEGIN_DOC.search(text)
    if not m_begin:
        return text
    m_end = RE_END_DOC.search(text, m_begin.end())
    if m_end:
        return text[m_begin.end():m_end.start()]
    return text[m_begin.end():]


def _clean_inline(text: str) -> str:
    """Разворачивает простые форматирующие команды и подставляет литералы."""
    # \cite → [@key]
    text = RE_CITE.sub(lambda m: "[@" + m.group(1).strip() + "]", text)

    # Текстовые команды-обёртки
    # Несколько проходов для вложенностей вида \textbf{\textit{…}} (без скобок внутри)
    for _ in range(3):
        for rx in RE_LATEX_ARG_TEXT.values():
            text = rx.sub(lambda m: m.group(1), text)

    # Снимаем \label{...}
    text = RE_LABEL.sub("", text)

    # Шумовые команды
    text = RE_SIMPLE_NOISE.sub("", text)

    # Подстановки символов
    for rx, replacement in LATEX_REPLACEMENTS:
        text = rx.sub(replacement, text)

    # Сжимаем пробелы
    text = re.sub(r'[ \t]+', ' ', text)
    return text


def _split_paragraphs(text: str) -> List[str]:
    chunks = re.split(r'\n\s*\n', text)
    return [c.strip() for c in chunks if c.strip()]


def _parse_table_body(body: str) -> List[List[str]]:
    """Парсит тело tabular: разбивает на строки по \\\\, ячейки — по &."""
    # Удаляем спецификацию столбцов сразу после tabular: {l|c|r}
    # (она уже была частью opts, но на всякий случай)
    body = re.sub(r'^\s*\{[^{}]*\}', '', body, count=1)
    # Удаляем строки-разделители
    body = re.sub(r'\\(?:hline|toprule|midrule|bottomrule|cline\s*\{[^}]*\})', '', body)
    rows: List[List[str]] = []
    raw_rows = RE_DOUBLE_BACKSLASH.split(body)
    for raw in raw_rows:
        raw = raw.strip()
        if not raw:
            continue
        cells = [c.strip() for c in re.split(r'(?<!\\)&', raw)]
        cells = [_clean_inline(c) for c in cells]
        if any(c for c in cells):
            rows.append(cells)
    return rows


def _parse_list_env(body: str, ordered: bool) -> List[ListItem]:
    """Парсит itemize/enumerate, возвращает список ListItem."""
    items: List[ListItem] = []
    # Разбиваем по \item, отбрасываем содержимое до первого \item
    parts = RE_ITEM.split(body)
    if not parts:
        return items
    parts = parts[1:]  # до первого \item — мусор / опции
    for raw in parts:
        text = _clean_inline(raw).strip().rstrip(";").strip()
        if not text:
            continue
        items.append(ListItem(text=text, ordered=ordered))
    return items


def _extract_bibitems(body: str) -> List[Reference]:
    """Парсит \\begin{thebibliography}{...}…\\end{thebibliography}."""
    refs: List[Reference] = []
    # Удаляем заглушку аргумента {99} и т.п.
    body = re.sub(r'^\s*\{[^}]*\}', '', body, count=1)
    # Разбиваем по \bibitem, попутно ловим ключ
    splits = re.split(r'\\bibitem\s*(?:\[[^\]]*\])?\s*\{([^}]+)\}', body)
    # splits = [пред-текст, key1, body1, key2, body2, ...]
    if len(splits) < 3:
        return refs
    pairs = list(zip(splits[1::2], splits[2::2]))
    for key, raw in pairs:
        text = _clean_inline(raw).strip().rstrip(",.;").strip()
        if key.strip() and text:
            refs.append(Reference(key=key.strip(), text=text))
    return refs


# ──────────────────────────────────────────
#  Основной импортёр
# ──────────────────────────────────────────
def import_latex(data: bytes, filename: Optional[str] = None) -> ImportResult:
    raw = data.decode("utf-8", errors="replace")
    return _import_latex_text(raw)


def _import_latex_text(raw: str) -> ImportResult:
    elements: List = []
    references: List[Reference] = []
    warnings: List[str] = []

    text = _strip_comments(raw)
    text = _trim_to_document(text)

    # 1) Извлекаем все среды (figure, table, itemize, enumerate, thebibliography…)
    #    и заменяем на маркеры-плейсхолдеры, чтобы остался плоский текст с заголовками.
    placeholders: List = []  # элементы или списки элементов в порядке появления
    bib_extracted = False

    def _on_env(match: re.Match) -> str:
        name = match.group("name")
        body = match.group("body")
        idx = len(placeholders)
        marker = f"\n@@ENV{idx}@@\n"

        if name in ("figure", "figure*"):
            path_m = RE_INCLUDE_GRAPHICS.search(body)
            cap_m = RE_CAPTION.search(body)
            placeholders.append(FigureRef(
                path=(path_m.group(1).strip() if path_m else "image.png"),
                caption=(_clean_inline(cap_m.group(1)).strip() if cap_m else ""),
            ))
            return marker

        if name in ("table", "table*"):
            cap_m = RE_CAPTION.search(body)
            tab_m = re.search(
                r'\\begin\s*\{tabular\*?\}\s*(?:\[[^\]]*\])?\s*\{[^{}]*\}'
                r'(?P<tbody>.*?)\\end\s*\{tabular\*?\}',
                body, re.S,
            )
            rows: List[List[str]] = []
            if tab_m:
                rows = _parse_table_body(tab_m.group("tbody"))
            placeholders.append(TableElement(
                rows=rows,
                caption=(_clean_inline(cap_m.group(1)).strip() if cap_m else ""),
                has_header=True,
            ))
            return marker

        if name == "tabular":
            placeholders.append(TableElement(
                rows=_parse_table_body(body),
                caption="",
                has_header=True,
            ))
            return marker

        if name in ("itemize", "description"):
            placeholders.append(_parse_list_env(body, ordered=False))
            return marker

        if name == "enumerate":
            placeholders.append(_parse_list_env(body, ordered=True))
            return marker

        if name == "thebibliography":
            nonlocal bib_extracted
            bib_extracted = True
            references.extend(_extract_bibitems(body))
            return ""  # удаляем из тела документа

        if name in ("equation", "equation*", "align", "align*", "displaymath"):
            placeholders.append(FormulaElement(text=_clean_inline(body).strip()))
            return marker

        if name == "abstract":
            # Маппим в РЕФЕРАТ
            placeholders.append([
                Heading(level=1, text="РЕФЕРАТ"),
                Paragraph(text=_clean_inline(body).strip()),
            ])
            return marker

        # Прочие среды — оставляем тело как есть
        return body

    # Применяем итеративно (вложенные среды разворачиваются за несколько проходов)
    prev = None
    while prev != text:
        prev = text
        text = RE_ENV.sub(_on_env, text)

    # 2) Разбиваем плоский текст на абзацы и вкрапляем плейсхолдеры по маркерам
    chunks = re.split(r'@@ENV(\d+)@@', text)
    # chunks = [текст0, idx1, текст1, idx2, текст2, ...]

    def _emit_text_block(block: str):
        block = block.strip()
        if not block:
            return
        # Заголовки: \section{...}, \subsection{...}, \subsubsection{...}, \paragraph{...}
        # Идём по строкам, чтобы корректно разбить на абзацы между заголовками.
        # Разбиваем по строкам с заголовками.
        i = 0
        # Собираем накопленный текст до встречи заголовка
        buf: List[str] = []

        def _flush():
            nonlocal buf
            if buf:
                joined = "\n".join(buf).strip()
                buf = []
                if joined:
                    for para in _split_paragraphs(joined):
                        cleaned = _clean_inline(para).strip()
                        if cleaned:
                            elements.append(Paragraph(text=cleaned))

        # Обрабатываем построчно
        for line in block.splitlines():
            stripped = line.strip()
            m = RE_SECTION.search(stripped)
            if m:
                _flush()
                elements.append(Heading(level=1, text=_clean_inline(m.group(1)).strip()))
                continue
            m = RE_SUBSECTION.search(stripped)
            if m:
                _flush()
                elements.append(Heading(level=2, text=_clean_inline(m.group(1)).strip()))
                continue
            m = RE_SUBSUBSECTION.search(stripped)
            if m:
                _flush()
                elements.append(Heading(level=3, text=_clean_inline(m.group(1)).strip()))
                continue
            m = RE_PARAGRAPH_CMD.search(stripped)
            if m:
                _flush()
                elements.append(Heading(level=3, text=_clean_inline(m.group(1)).strip()))
                continue
            buf.append(line)
        _flush()

    for k, chunk in enumerate(chunks):
        if k % 2 == 0:
            _emit_text_block(chunk)
        else:
            ph = placeholders[int(chunk)]
            if isinstance(ph, list):
                elements.extend(ph)
            else:
                elements.append(ph)

    if not elements:
        warnings.append("LaTeX: не удалось распознать структуру документа.")

    if bib_extracted:
        warnings.append(
            f"LaTeX: извлечено {len(references)} источник(ов) из thebibliography."
        )

    return ImportResult(
        elements=elements,
        images={},
        references=references,
        warnings=warnings,
    )
