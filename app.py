"""
Streamlit-приложение для подготовки отчёта о НИР по ГОСТ 7.32-2017.

Архитектура:
    - Report     — упорядоченный набор частей + источники + метаданные.
    - importers/ — Markdown / DOCX / LaTeX / XLSX / CSV.
    - bibliography — сквозная нумерация ссылок [@key] → [N].
    - docx_builder — рендерит итоговый DOCX по ГОСТ.
"""

import os
import sys

import streamlit as st

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from elements import (
    Heading, Paragraph, FigureRef, TableElement, ListItem,
    is_special_heading,
)
from bibliography import Bibliography, Reference
from document import Report, ReportPart
from importers import (
    ImportResult, get_importer_for_filename, list_supported_extensions,
)
from importers.markdown_imp import import_markdown_text
from docx_builder import GostDocxBuilder


# ─────────────────────────────────────────────
#  ШАБЛОН ПРИМЕРА (Markdown)
# ─────────────────────────────────────────────
EXAMPLE_TEXT = """\
# 1 ВВЕДЕНИЕ

Настоящий отчёт посвящён исследованию методов машинного обучения. Согласно [@ivanov2020], классические методы остаются конкурентоспособными при ограниченных данных, что подтверждено и в [@petrov2021].

## 1.1 Цель и задачи

Цель — разработка и исследование метода автоматической классификации.

Для достижения цели поставлены задачи:

- провести анализ существующих методов;
- разработать алгоритм;
- провести экспериментальное исследование.

# 2 ОСНОВНАЯ ЧАСТЬ

## 2.1 Методы

В качестве основного выбран метод опорных векторов (SVM) [@petrov2021].

[таблица]
Метод | Точность, % | F1
SVM | 92.3 | 90.7
BERT | 95.6 | 94.9
[/таблица | Сравнение методов]

# 3 ЗАКЛЮЧЕНИЕ

Метод обеспечивает точность не ниже 92 %, что согласуется с данными [@ivanov2020, petrov2021].
"""

EXAMPLE_BIB = """\
# Формат: ключ | оформление по ГОСТ Р 7.0.5-2008
ivanov2020 | Иванов И. И. Методы машинного обучения. — М.: Наука, 2020. — 320 с.
petrov2021 | Петров П. П. Классификация текстов на основе SVM // Вестник ВУЗа. — 2021. — № 4. — С. 12–25.
"""


# ─────────────────────────────────────────────
#  КОНФИГУРАЦИЯ СТРАНИЦЫ
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Отчёт НИР по ГОСТ 7.32-2017",
    page_icon="📄",
    layout="wide",
)

st.markdown("""
<style>
    .gost-header {
        background: linear-gradient(135deg, #1a3c6e 0%, #2d6bb5 100%);
        color: white; padding: 1.2rem 1.6rem; border-radius: 8px;
        margin-bottom: 1rem;
    }
    .gost-header h1 { margin: 0; font-size: 1.4rem; }
    .gost-header p  { margin: 0.2rem 0 0; opacity: 0.85; font-size: 0.85rem; }

    .part-card {
        border: 1px solid #d8def0; border-radius: 8px;
        padding: 0.6rem 0.9rem; margin: 0.4rem 0;
        background: #fbfcff;
    }
    .part-disabled { opacity: 0.55; }
    .part-meta { color: #5a6478; font-size: 0.8rem; }

    .check-ok  { color: #1a7a3c; font-weight: 600; }
    .check-warn { color: #b86a00; font-weight: 600; }
    .check-err  { color: #c0392b; font-weight: 600; }

    .ref-unknown {
        background: #fff3f0; border-left: 3px solid #c0392b;
        padding: 0.3rem 0.7rem; border-radius: 0 4px 4px 0;
        font-size: 0.85rem; margin: 2px 0;
    }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="gost-header">
  <h1>📄 Отчёт о НИР — оформление по ГОСТ 7.32-2017</h1>
  <p>Сборка отчёта из нескольких источников (Word / LaTeX / Markdown / Excel / CSV) со сквозной нумерацией.</p>
</div>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
#  СОСТОЯНИЕ
# ─────────────────────────────────────────────
def _init_state():
    if "parts" not in st.session_state:
        st.session_state.parts: list[ReportPart] = []
    if "bibliography_text" not in st.session_state:
        st.session_state.bibliography_text = EXAMPLE_BIB
    if "uploader_nonce" not in st.session_state:
        # Чтобы можно было «программно» очистить file_uploader
        st.session_state.uploader_nonce = 0
    if "md_editor" not in st.session_state:
        st.session_state.md_editor = EXAMPLE_TEXT
    if "extra_images" not in st.session_state:
        st.session_state.extra_images: dict[str, bytes] = {}


_init_state()


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
    approver_position = st.text_input("Должность утверждающего", value="Директор")
    approver_name     = st.text_input("ФИО утверждающего", value="И.О. Фамилия")

    st.divider()
    st.subheader("👥 Исполнители")
    authors_raw = st.text_area(
        "Каждый с новой строки: «Должность | ФИО»",
        value="Ст. науч. сотр. | И.О. Фамилия\nМл. науч. сотр. | И.О. Фамилия",
        height=90,
    )

    st.divider()
    st.subheader("📌 Структура")
    add_title    = st.checkbox("Титульный лист", value=True)
    add_abstract = st.checkbox("Реферат (заглушка)", value=True)
    add_toc      = st.checkbox("Автособираемое содержание", value=True)


# ─────────────────────────────────────────────
#  ВКЛАДКИ
# ─────────────────────────────────────────────
tab_parts, tab_sources, tab_check = st.tabs(
    ["📥 Части отчёта", "📚 Источники", "✅ Проверка ГОСТ"]
)


# ============================================================
#  ВКЛАДКА: ЧАСТИ ОТЧЁТА
# ============================================================
with tab_parts:
    st.markdown(
        "Загрузите файлы разных форматов — они склеятся в один отчёт со "
        "сквозной нумерацией разделов, рисунков и таблиц."
    )

    col_upload, col_md = st.columns([1, 1], gap="large")

    # ── Загрузка файлов ───────────────────────
    with col_upload:
        st.subheader("📁 Загрузить файл(ы)")
        uploader_key = f"file_uploader_{st.session_state.uploader_nonce}"
        uploaded = st.file_uploader(
            "Поддержка: " + ", ".join(list_supported_extensions()),
            type=list_supported_extensions(),
            accept_multiple_files=True,
            key=uploader_key,
        )

        page_break = st.checkbox(
            "Разрыв страницы после каждой загруженной части",
            value=True,
            key="upload_pgbreak",
        )

        if st.button("➕ Добавить файлы как части", type="primary",
                     use_container_width=True):
            if not uploaded:
                st.warning("Сначала выберите файлы.")
            else:
                added = 0
                for f in uploaded:
                    importer = get_importer_for_filename(f.name)
                    if importer is None:
                        st.error(f"❌ Неподдерживаемый формат: {f.name}")
                        continue
                    try:
                        result = importer(f.read(), f.name)
                    except Exception as e:
                        st.error(f"❌ {f.name}: ошибка импорта — {e}")
                        continue

                    ext = f.name.rsplit(".", 1)[-1].lower() if "." in f.name else "?"
                    part = ReportPart.from_import(
                        label=f.name,
                        source_kind=ext,
                        result=result,
                        page_break_after=page_break,
                    )
                    st.session_state.parts.append(part)

                    # Если файл принёс источники — дописываем в текстовое поле
                    if result.references:
                        existing = st.session_state.bibliography_text or ""
                        appended = "\n".join(
                            f"{r.key} | {r.text}" for r in result.references
                            if f"{r.key} |" not in existing
                        )
                        if appended:
                            st.session_state.bibliography_text = (
                                existing.rstrip() + "\n" + appended + "\n"
                            )
                    added += 1

                if added:
                    st.session_state.uploader_nonce += 1  # сбрасываем uploader
                    st.success(f"✅ Добавлено частей: {added}")
                    st.rerun()

        st.markdown("---")
        st.subheader("🖼️ Дополнительные изображения")
        st.caption(
            "Имена файлов должны совпадать с тем, что указано в `[рисунок: имя.png …]` "
            "в Markdown- или LaTeX-частях."
        )
        extra_imgs = st.file_uploader(
            "PNG / JPG",
            type=["png", "jpg", "jpeg"],
            accept_multiple_files=True,
            key="extra_imgs",
        )
        if extra_imgs:
            for f in extra_imgs:
                st.session_state.extra_images[f.name] = f.read()
            st.success(f"Доступно изображений: {len(st.session_state.extra_images)}")

    # ── Markdown-часть «на лету» ────────────────
    with col_md:
        st.subheader("📝 Текстовая часть (Markdown)")
        st.caption(
            "Шпаргалка по разметке — в раскрывающемся блоке ниже. "
            "Ссылки на источники: `[@ключ]` или `[@ключ1, ключ2]`."
        )
        st.text_area(
            "Введите или вставьте текст:",
            key="md_editor",
            height=320,
        )
        with st.expander("📖 Синтаксис разметки"):
            st.markdown("""
| Элемент | Синтаксис |
|---|---|
| Раздел (H1) | `# 1 НАЗВАНИЕ` |
| Подраздел (H2) | `## 1.1 Название` |
| Пункт (H3) | `### 1.1.1 Название` |
| Рисунок | `[рисунок: имя.png \\| Подпись]` |
| Таблица | `[таблица]` … `[/таблица \\| Название]` |
| Список | `- пункт` или `1) пункт` |
| Ссылка на источник | `[@ключ]` |
| Разрыв страницы | `---` |
            """)
        if st.button("➕ Добавить как часть", use_container_width=True):
            text = st.session_state.md_editor.strip()
            if not text:
                st.warning("Текст пуст.")
            else:
                result = import_markdown_text(text)
                part = ReportPart.from_import(
                    label="Markdown (вручную)",
                    source_kind="markdown",
                    result=result,
                    page_break_after=False,
                )
                st.session_state.parts.append(part)
                st.success("Добавлено.")
                st.rerun()

    # ── Список частей ────────────────────────────
    st.markdown("---")
    st.subheader("📋 Состав отчёта")
    if not st.session_state.parts:
        st.info("Пока ни одной части. Загрузите файл или добавьте Markdown-часть.")
    else:
        for idx, part in enumerate(st.session_state.parts):
            disabled_cls = "" if part.enabled else "part-disabled"
            n_elements = len(part.elements)
            n_imgs = len(part.images)
            with st.container():
                st.markdown(
                    f'<div class="part-card {disabled_cls}">'
                    f'<b>{idx + 1}. {part.label}</b> '
                    f'<span class="part-meta">— {part.source_kind}, '
                    f'элементов: {n_elements}, изображений: {n_imgs}'
                    f'{", разрыв страницы" if part.page_break_after else ""}</span>'
                    f'</div>',
                    unsafe_allow_html=True,
                )
                c1, c2, c3, c4, c5 = st.columns([1, 1, 1, 1, 6])
                if c1.button("▲", key=f"up_{idx}", help="Выше",
                             disabled=(idx == 0)):
                    st.session_state.parts[idx], st.session_state.parts[idx - 1] = (
                        st.session_state.parts[idx - 1], st.session_state.parts[idx]
                    )
                    st.rerun()
                if c2.button("▼", key=f"down_{idx}", help="Ниже",
                             disabled=(idx == len(st.session_state.parts) - 1)):
                    st.session_state.parts[idx], st.session_state.parts[idx + 1] = (
                        st.session_state.parts[idx + 1], st.session_state.parts[idx]
                    )
                    st.rerun()
                toggle_label = "Выключить" if part.enabled else "Включить"
                if c3.button(toggle_label, key=f"tog_{idx}"):
                    part.enabled = not part.enabled
                    st.rerun()
                if c4.button("✕", key=f"del_{idx}", help="Удалить"):
                    st.session_state.parts.pop(idx)
                    st.rerun()
                # Предупреждения от импортёра
                if part.warnings:
                    for w in part.warnings:
                        c5.caption(f"⚠️ {w}")

        if st.button("🗑 Очистить все части"):
            st.session_state.parts = []
            st.rerun()


# ============================================================
#  ВКЛАДКА: ИСТОЧНИКИ
# ============================================================
with tab_sources:
    st.markdown(
        "Один источник на строку, формат: `ключ | Полное оформление по ГОСТ Р 7.0.5`. "
        "Ссылка в тексте: `[@ключ]` (можно перечислять через запятую: `[@a, b]`)."
    )

    st.text_area(
        "Источники",
        key="bibliography_text",
        height=300,
    )

    bib_preview = Bibliography.from_text(st.session_state.bibliography_text)
    st.caption(f"Распознано источников: **{len(bib_preview)}**")

    # Какие ключи реально использованы в загруженных частях?
    if st.session_state.parts:
        merged_for_check: list = []
        for p in st.session_state.parts:
            if p.enabled:
                merged_for_check.extend(p.elements)
        used_keys = bib_preview.collect_used(merged_for_check)
        defined_keys = {r.key for r in bib_preview.all()}
        missing = [k for k in used_keys if k not in defined_keys]
        unused  = [k for k in defined_keys if k not in used_keys]

        col_a, col_b = st.columns(2)
        col_a.metric("Использовано в тексте", len(used_keys))
        col_b.metric("Описано в списке", len(defined_keys))

        if missing:
            st.markdown("**❌ Отсутствуют в списке источников:**")
            for k in missing:
                st.markdown(
                    f'<div class="ref-unknown">{k}</div>',
                    unsafe_allow_html=True,
                )
        if unused:
            st.caption(
                "ℹ️ Описаны, но не использованы (в финальный список не попадут): "
                + ", ".join(unused)
            )


# ============================================================
#  ВКЛАДКА: ПРОВЕРКА ГОСТ
# ============================================================
with tab_check:
    if not st.session_state.parts:
        st.info("Добавьте хотя бы одну часть, чтобы запустить проверку.")
    else:
        merged: list = []
        for p in st.session_state.parts:
            if p.enabled:
                merged.extend(p.elements)

        h1 = [el for el in merged if isinstance(el, Heading) and el.level == 1]
        h2 = [el for el in merged if isinstance(el, Heading) and el.level == 2]
        figs = [el for el in merged if isinstance(el, FigureRef)]
        tbls = [el for el in merged if isinstance(el, TableElement)]

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Разделы", len(h1))
        c2.metric("Подразделы", len(h2))
        c3.metric("Рисунки", len(figs))
        c4.metric("Таблицы", len(tbls))

        st.markdown("---")
        st.subheader("Чеклист")

        upper_titles = [el.text.upper() for el in h1]
        bib = Bibliography.from_text(st.session_state.bibliography_text)
        used_keys = bib.collect_used(merged)
        missing = [k for k in used_keys if not bib.get(k)]

        checks = [
            ("Титульный лист", add_title, True),
            ("Реферат", add_abstract or any("РЕФЕРАТ" in t for t in upper_titles), True),
            ("Содержание", add_toc, True),
            ("Введение", any("ВВЕДЕНИЕ" in t for t in upper_titles), True),
            ("Заключение", any("ЗАКЛЮЧЕНИЕ" in t for t in upper_titles), False),
            ("Источники описаны", len(bib) > 0, False),
            ("Все ссылки [@…] определены", len(missing) == 0, True),
            (f"Хотя бы один рисунок", len(figs) > 0, False),
            (f"Хотя бы одна таблица", len(tbls) > 0, False),
        ]
        for name, ok, required in checks:
            if ok:
                st.markdown(f'<span class="check-ok">✔</span> {name}',
                            unsafe_allow_html=True)
            elif required:
                st.markdown(
                    f'<span class="check-err">✘</span> {name} '
                    f'<i>(обязательный элемент)</i>',
                    unsafe_allow_html=True,
                )
            else:
                st.markdown(
                    f'<span class="check-warn">⚠</span> {name} '
                    f'<i>(рекомендуется)</i>',
                    unsafe_allow_html=True,
                )


# ─────────────────────────────────────────────
#  КНОПКА ГЕНЕРАЦИИ
# ─────────────────────────────────────────────
st.divider()
gen_col, _ = st.columns([1, 3])
generate = gen_col.button("🔄 Сформировать DOCX", type="primary",
                          use_container_width=True,
                          disabled=not st.session_state.parts)

if generate:
    # 1. Метаданные
    authors = []
    for line in authors_raw.strip().splitlines():
        parts = [p.strip() for p in line.split("|")]
        if len(parts) == 2:
            authors.append({"position": parts[0], "name": parts[1]})
        elif parts[0]:
            authors.append({"position": parts[0], "name": ""})

    metadata = {
        "ministry": ministry, "org": org, "title": title,
        "theme_code": theme_code, "udc": udc, "inv_number": inv_number,
        "city": city, "year": year,
        "head_position": head_position, "head_name": head_name,
        "approver_position": approver_position, "approver_name": approver_name,
        "authors": authors,
    }

    # 2. Сборка модели Report
    bib = Bibliography.from_text(st.session_state.bibliography_text)
    report = Report(
        metadata=metadata,
        parts=list(st.session_state.parts),
        bibliography=bib,
    )

    with st.spinner("Сборка отчёта…"):
        elements, images, ordered_refs, unknown = report.merge()
        # Дополнительные изображения, загруженные пользователем
        for name, blob in st.session_state.extra_images.items():
            images.setdefault(name, blob)

        builder = GostDocxBuilder(metadata)
        if add_title:
            builder.add_title_page()
        if add_abstract:
            builder.add_abstract_placeholder()
        if add_toc:
            builder.add_toc()

        docx_bytes = builder.build(elements, images, references=ordered_refs)

    st.success(
        f"✅ Готово. Источников в списке: {len(ordered_refs)}. "
        f"Неизвестных ключей: {len(unknown)}."
    )
    if unknown:
        st.warning("Неизвестные ключи (отрендерены как `[?]`): " + ", ".join(unknown))

    safe_title = "".join(c if c.isalnum() or c in " _-" else "_" for c in title)[:40]
    filename = f"НИР_{safe_title}_{year}.docx"

    st.download_button(
        "📥 Скачать DOCX",
        data=docx_bytes,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        type="primary",
    )

st.markdown("---")
st.caption(
    "Документ формируется по **ГОСТ 7.32-2017** «Отчёт о научно-исследовательской работе». "
    "Список источников оформляется в порядке упоминания в тексте."
)
