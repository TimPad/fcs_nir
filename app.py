"""
Streamlit-приложение: форматирование текста по ГОСТ 7.32-2017 (отчёт о НИР).
"""

import streamlit as st
import os, sys

sys.path.insert(0, os.path.dirname(__file__))

from parser import GostTextParser
from docx_builder import GostDocxBuilder

# ─────────────────────────────────────────────
#  ПРИМЕР ТЕКСТА
# ─────────────────────────────────────────────
EXAMPLE_TEXT = """\
# 1 ВВЕДЕНИЕ

Настоящий отчёт посвящён исследованию методов машинного обучения применительно к задачам обработки естественного языка. В ходе работы были рассмотрены основные подходы и алгоритмы.

Актуальность работы определяется возрастающим объёмом неструктурированных текстовых данных в различных прикладных областях.

## 1.1 Цель и задачи исследования

Целью настоящей работы является разработка и исследование метода автоматической классификации текстовых документов.

Для достижения цели поставлены следующие задачи:

- провести анализ существующих методов обработки естественного языка;
- разработать алгоритм классификации документов;
- провести экспериментальное исследование.

### 1.1.1 Ограничения исследования

В рамках данной работы рассматриваются только тексты на русском языке объёмом не менее 100 символов.

# 2 ОСНОВНАЯ ЧАСТЬ

## 2.1 Методы и материалы

В качестве основного метода выбран метод опорных векторов (SVM), показавший высокую эффективность на стандартных бенчмарках.

[рисунок: diagram.png | Архитектура предложенного метода классификации]

Результаты сравнительного анализа методов представлены в таблице 1.

[таблица]
Метод | Точность, % | Полнота, % | F1-мера
SVM | 92.3 | 89.1 | 90.7
BERT | 95.6 | 94.2 | 94.9
Наивный Байес | 81.4 | 78.3 | 79.8
[/таблица | Сравнение методов классификации]

## 2.2 Результаты экспериментов

Эксперименты проводились на датасете из 10 000 документов, разделённых на обучающую (80%) и тестовую (20%) выборки.

# 3 ЗАКЛЮЧЕНИЕ

В результате проведённых исследований разработан и апробирован метод автоматической классификации текстовых документов. Предложенный подход обеспечивает точность классификации не менее 92 %.

## 3.1 Выводы

1) Метод SVM показал устойчивые результаты на всех тестовых наборах.
2) Дальнейшее улучшение возможно за счёт применения предобученных языковых моделей.
"""

# ─────────────────────────────────────────────
#  КОНФИГУРАЦИЯ СТРАНИЦЫ
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Форматирование НИР по ГОСТ 7.32-2017",
    page_icon="📄",
    layout="wide",
)

# CSS для улучшения внешнего вида
st.markdown("""
<style>
    .gost-header {
        background: linear-gradient(135deg, #1a3c6e 0%, #2d6bb5 100%);
        color: white;
        padding: 1.5rem 2rem;
        border-radius: 8px;
        margin-bottom: 1.5rem;
    }
    .gost-header h1 { margin: 0; font-size: 1.6rem; }
    .gost-header p  { margin: 0.3rem 0 0; opacity: 0.85; font-size: 0.9rem; }

    .gost-rule {
        background: #f0f4fa;
        border-left: 4px solid #2d6bb5;
        padding: 0.5rem 1rem;
        border-radius: 0 6px 6px 0;
        margin: 0.3rem 0;
        font-size: 0.85rem;
    }
    .check-ok  { color: #1a7a3c; font-weight: bold; }
    .check-warn { color: #b86a00; font-weight: bold; }
    .check-err  { color: #c0392b; font-weight: bold; }

    .preview-item {
        padding: 2px 8px;
        border-radius: 4px;
        font-size: 0.88rem;
        margin: 1px 0;
    }
    .preview-h1 { background:#e8f0fe; font-weight:bold; border-left:4px solid #2d6bb5; }
    .preview-h2 { background:#f5f7ff; border-left:3px solid #6890d4; padding-left:16px; }
    .preview-h3 { background:#fafbff; border-left:2px solid #a0b8e8; padding-left:24px; font-style:italic; }
    .preview-fig { background:#fff8e1; border-left:3px solid #f0a500; }
    .preview-tbl { background:#e8f5e9; border-left:3px solid #34a853; }
    .preview-para { color:#555; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
#  ЗАГОЛОВОК
# ─────────────────────────────────────────────
st.markdown("""
<div class="gost-header">
  <h1>📄 Форматирование НИР по ГОСТ 7.32-2017</h1>
  <p>Автоматическое оформление отчёта о научно-исследовательской работе</p>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
#  БОКОВАЯ ПАНЕЛЬ — МЕТАДАННЫЕ
# ─────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Сведения об отчёте")

    ministry = st.text_input("Министерство / ведомство",
        value="Министерство науки и высшего образования Российской Федерации")
    org = st.text_input("Организация",
        value="ФГБОУ ВО «Название университета»")
    title = st.text_input("Название НИР",
        value="Исследование методов машинного обучения")
    theme_code = st.text_input("Шифр темы", value="НИР-2024-01")
    udc = st.text_input("УДК", value="004.8")
    inv_number = st.text_input("Инв. №", value="")
    city = st.text_input("Город", value="Москва")
    year = st.text_input("Год", value="2024")

    st.divider()
    st.subheader("👤 Руководитель")
    head_position = st.text_input("Должность руководителя", value="д-р техн. наук, проф.")
    head_name     = st.text_input("ФИО руководителя", value="И.О. Фамилия")
    approver_name = st.text_input("ФИО утверждающего", value="И.О. Фамилия")
    approver_position = st.text_input("Должность утверждающего", value="Директор")

    st.divider()
    st.subheader("👥 Исполнители")
    authors_raw = st.text_area(
        "Исполнители (каждый с новой строки: Должность | ФИО)",
        value="Ст. науч. сотр. | И.О. Фамилия\nМл. науч. сотр. | И.О. Фамилия",
        height=100
    )

    st.divider()
    st.subheader("📌 Параметры генерации")
    add_title   = st.checkbox("Генерировать титульный лист", value=True)
    add_abstract = st.checkbox("Добавить заготовку реферата", value=True)
    add_toc     = st.checkbox("Автоматическое содержание (TOC)", value=True)

# ─────────────────────────────────────────────
#  ОСНОВНАЯ ОБЛАСТЬ
# ─────────────────────────────────────────────
col_input, col_preview = st.columns([3, 2], gap="large")

with col_input:
    st.subheader("📝 Входной текст")

    with st.expander("📖 Синтаксис разметки", expanded=False):
        st.markdown("""
| Элемент | Синтаксис |
|---|---|
| Раздел (H1) | `# 1 НАЗВАНИЕ РАЗДЕЛА` |
| Подраздел (H2) | `## 1.1 Название подраздела` |
| Пункт (H3) | `### 1.1.1 Название пункта` |
| Рисунок | `[рисунок: имя_файла.png \\| Подпись]` |
| Таблица | `[таблица]` ... `[/таблица \\| Название]` |
| Маркированный список | `- Элемент списка` |
| Нумерованный список | `1) Первый пункт` |
| Разрыв страницы | `---` |

**Столбцы таблицы** разделяются символом `|`.
        """)

    text_input = st.text_area(
        "Введите или вставьте текст НИР:",
        value=EXAMPLE_TEXT,
        height=500,
        placeholder="# 1 НАЗВАНИЕ РАЗДЕЛА\n\nТекст абзаца...",
        label_visibility="collapsed"
    )

    # Загрузка изображений
    st.subheader("🖼️ Изображения")
    uploaded_images = st.file_uploader(
        "Загрузите рисунки (PNG, JPG, JPEG)",
        type=["png", "jpg", "jpeg"],
        accept_multiple_files=True
    )

# ─────────────────────────────────────────────
#  ПРЕВЬЮ СТРУКТУРЫ
# ─────────────────────────────────────────────
with col_preview:
    st.subheader("👁️ Структура документа")

    parser = GostTextParser()

    if text_input.strip():
        elements = parser.parse(text_input)
        elements = parser.auto_number(elements)

        # Отображение структуры
        from parser import Heading, Paragraph, FigureRef, TableElement, ListItem, FormulaElement

        preview_html = []
        for el in elements:
            if isinstance(el, Heading):
                num = f"{el.number} " if el.number else ""
                lvl = f"preview-h{el.level}"
                icon = {"1": "📌", "2": "📎", "3": "▸"}.get(str(el.level), "")
                preview_html.append(
                    f'<div class="preview-item {lvl}">{icon} {num}{el.text}</div>'
                )
            elif isinstance(el, FigureRef):
                n = getattr(el, "_number", "?")
                preview_html.append(
                    f'<div class="preview-item preview-fig">🖼️ Рисунок {n} — {el.caption or el.path}</div>'
                )
            elif isinstance(el, TableElement):
                n = getattr(el, "_number", "?")
                preview_html.append(
                    f'<div class="preview-item preview-tbl">📊 Таблица {n} — {el.caption}</div>'
                )
            elif isinstance(el, Paragraph):
                snippet = el.text[:60] + ("…" if len(el.text) > 60 else "")
                preview_html.append(
                    f'<div class="preview-item preview-para">📝 {snippet}</div>'
                )
            elif isinstance(el, ListItem):
                prefix = "1)" if el.ordered else "–"
                preview_html.append(
                    f'<div class="preview-item preview-para" style="padding-left:20px">'
                    f'{prefix} {el.text[:50]}…</div>'
                )

        st.markdown("\n".join(preview_html), unsafe_allow_html=True)

        # ── Чеклист соответствия ГОСТ ──────────
        st.divider()
        st.subheader("✅ Чеклист ГОСТ 7.32-2017")

        from parser import Heading
        headings_text = [el.text.upper() for el in elements if isinstance(el, Heading) and el.level == 1]
        fig_count = sum(1 for el in elements if isinstance(el, FigureRef))
        tbl_count = sum(1 for el in elements if isinstance(el, TableElement))

        checks = [
            ("Титульный лист", add_title, True),
            ("Реферат",   add_abstract or "РЕФЕРАТ" in headings_text, True),
            ("Содержание", add_toc, True),
            ("Введение",  any("ВВЕДЕНИЕ" in h for h in headings_text), True),
            ("Заключение", any("ЗАКЛЮЧЕНИЕ" in h for h in headings_text), False),
            ("Список источников",
             any("СПИСОК" in h for h in headings_text), False),
            (f"Рисунков: {fig_count}", fig_count > 0, False),
            (f"Таблиц: {tbl_count}", tbl_count > 0, False),
        ]

        for name, ok, required in checks:
            if ok:
                st.markdown(f'<span class="check-ok">✔</span> {name}', unsafe_allow_html=True)
            elif required:
                st.markdown(f'<span class="check-err">✘</span> {name} <i>(обязательный элемент)</i>', unsafe_allow_html=True)
            else:
                st.markdown(f'<span class="check-warn">⚠</span> {name} <i>(рекомендуется)</i>', unsafe_allow_html=True)

    else:
        st.info("Введите текст слева для отображения структуры.")

# ─────────────────────────────────────────────
#  СПРАВОЧНИК ГОСТ (раскрываемый)
# ─────────────────────────────────────────────
with st.expander("📏 Параметры оформления по ГОСТ 7.32-2017", expanded=False):
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown("**Страница (А4)**")
        st.markdown('<div class="gost-rule">Поле левое: 30 мм</div>', unsafe_allow_html=True)
        st.markdown('<div class="gost-rule">Поле правое: 10 мм</div>', unsafe_allow_html=True)
        st.markdown('<div class="gost-rule">Поле верхнее: 20 мм</div>', unsafe_allow_html=True)
        st.markdown('<div class="gost-rule">Поле нижнее: 20 мм</div>', unsafe_allow_html=True)
    with c2:
        st.markdown("**Текст**")
        st.markdown('<div class="gost-rule">Шрифт: Times New Roman 14 пт</div>', unsafe_allow_html=True)
        st.markdown('<div class="gost-rule">Межстрочный: полуторный</div>', unsafe_allow_html=True)
        st.markdown('<div class="gost-rule">Отступ абзаца: 1,25 см</div>', unsafe_allow_html=True)
        st.markdown('<div class="gost-rule">Выравнивание: по ширине</div>', unsafe_allow_html=True)
    with c3:
        st.markdown("**Заголовки**")
        st.markdown('<div class="gost-rule">H1: 14 пт, жирный, ПРОПИСНЫЕ, с новой стр.</div>', unsafe_allow_html=True)
        st.markdown('<div class="gost-rule">H2: 14 пт, жирный, с отступом</div>', unsafe_allow_html=True)
        st.markdown('<div class="gost-rule">H3: 14 пт, жирный курсив</div>', unsafe_allow_html=True)
        st.markdown('<div class="gost-rule">Нижний колонтитул: № стр. по центру, 12 пт</div>', unsafe_allow_html=True)

# ─────────────────────────────────────────────
#  КНОПКА ГЕНЕРАЦИИ
# ─────────────────────────────────────────────
st.divider()
col_btn, col_status = st.columns([1, 3])

with col_btn:
    generate = st.button("🔄 Сформировать DOCX", type="primary", use_container_width=True)

if generate:
    if not text_input.strip():
        st.error("❌ Введите текст документа!")
    else:
        with st.spinner("Формирование документа по ГОСТ 7.32-2017…"):

            # Парсинг текста
            parser = GostTextParser()
            elements = parser.parse(text_input)
            elements = parser.auto_number(elements)

            # Сбор изображений
            images = {}
            if uploaded_images:
                for f in uploaded_images:
                    images[f.name] = f.read()

            # Метаданные
            authors = []
            for line in authors_raw.strip().splitlines():
                parts = [p.strip() for p in line.split("|")]
                if len(parts) == 2:
                    authors.append({"position": parts[0], "name": parts[1]})
                elif parts[0]:
                    authors.append({"position": parts[0], "name": ""})

            meta = {
                "ministry":  ministry,
                "org":       org,
                "title":     title,
                "theme_code": theme_code,
                "udc":       udc,
                "inv_number": inv_number,
                "city":      city,
                "year":      year,
                "head_position": head_position,
                "head_name":     head_name,
                "approver_name": approver_name,
                "approver_position": approver_position,
                "authors":   authors,
            }

            # Сборка документа
            builder = GostDocxBuilder(meta)

            if add_title:
                builder.add_title_page()
            if add_abstract:
                builder.add_abstract_placeholder()
            if add_toc:
                builder.add_toc()

            docx_bytes = builder.build(elements, images)

        st.success("✅ Документ сформирован по ГОСТ 7.32-2017!")

        # Кнопка скачивания
        safe_title = "".join(c if c.isalnum() or c in " _-" else "_" for c in title)[:40]
        filename = f"НИР_{safe_title}_{year}.docx"

        st.download_button(
            label="📥 Скачать DOCX",
            data=docx_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
            use_container_width=False
        )

        with st.expander("📊 Статистика документа"):
            from parser import Heading, Paragraph, FigureRef, TableElement, ListItem
            h1s = [el for el in elements if isinstance(el, Heading) and el.level == 1]
            h2s = [el for el in elements if isinstance(el, Heading) and el.level == 2]
            paras = [el for el in elements if isinstance(el, Paragraph)]
            figs  = [el for el in elements if isinstance(el, FigureRef)]
            tbls  = [el for el in elements if isinstance(el, TableElement)]
            words = sum(len(el.text.split()) for el in paras)

            c1, c2, c3, c4, c5 = st.columns(5)
            c1.metric("Разделов", len(h1s))
            c2.metric("Подразделов", len(h2s))
            c3.metric("Рисунков", len(figs))
            c4.metric("Таблиц", len(tbls))
            c5.metric("Слов ~", words)

# ─────────────────────────────────────────────
#  ПОДВАЛ
# ─────────────────────────────────────────────
st.markdown("---")
st.caption("Документ формируется строго по **ГОСТ 7.32-2017** «Система разработки и постановки продукции на производство. Отчёт о научно-исследовательской работе».")
