"""
Модель отчёта: набор частей (каждая — с собственным источником) +
библиография + метаданные. Слияние возвращает плоский список элементов
для билдера.
"""

from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

from elements import DocElement, auto_number
from bibliography import Bibliography, Reference
from importers import ImportResult


@dataclass
class ReportPart:
    """Один кусок отчёта, импортированный из конкретного файла или текста."""

    label: str                          # человекочитаемое имя в UI
    source_kind: str                    # "markdown" | "docx" | "latex" | "xlsx" | "csv"
    elements: List[DocElement] = field(default_factory=list)
    images: Dict[str, bytes] = field(default_factory=dict)
    enabled: bool = True
    page_break_after: bool = False      # вставлять разрыв страницы после части
    warnings: List[str] = field(default_factory=list)

    @classmethod
    def from_import(cls, label: str, source_kind: str,
                    result: ImportResult,
                    page_break_after: bool = False) -> "ReportPart":
        return cls(
            label=label,
            source_kind=source_kind,
            elements=list(result.elements),
            images=dict(result.images),
            warnings=list(result.warnings),
            page_break_after=page_break_after,
        )


@dataclass
class Report:
    """Отчёт о НИР: упорядоченный набор частей + источники + метаданные."""

    metadata: Dict = field(default_factory=dict)
    parts: List[ReportPart] = field(default_factory=list)
    bibliography: Bibliography = field(default_factory=Bibliography)

    # ──────────────────────────────────────────
    #  СБОРКА ПЛОСКОГО ДОКУМЕНТА
    # ──────────────────────────────────────────
    def merge(self) -> Tuple[List[DocElement], Dict[str, bytes], List[Reference], List[str]]:
        """
        Сливает все включённые части в один список элементов.
        Префиксует имена изображений именем части (избегаем коллизий).
        Возвращает: (elements, images, ordered_references, unknown_keys).

        Сквозная нумерация (разделы/рисунки/таблицы) применяется ПОСЛЕ слияния,
        что и обеспечивает непрерывность нумерации между частями.
        """
        from elements import PageBreak, FigureRef

        merged_elements: List[DocElement] = []
        merged_images: Dict[str, bytes] = {}

        for idx, part in enumerate(self.parts):
            if not part.enabled or not part.elements:
                continue

            # Префикс к имени изображения, чтобы части с одинаковыми
            # filename'ами не перетирали друг друга.
            prefix = f"part{idx}_"
            for name, blob in part.images.items():
                merged_images[prefix + name] = blob

            for el in part.elements:
                # Обновляем ссылку на путь, если это рисунок
                if isinstance(el, FigureRef) and el.path:
                    new_path = prefix + el.path
                    if new_path in merged_images:
                        el = FigureRef(
                            path=new_path,
                            caption=el.caption,
                            image_data=el.image_data,
                        )
                merged_elements.append(el)

            if part.page_break_after and idx != len(self.parts) - 1:
                merged_elements.append(PageBreak())

        # Замена меток [@key] → [N]
        ordered_refs, unknown = self.bibliography.renumber_and_replace(merged_elements)

        # Сквозная нумерация заголовков/рисунков/таблиц
        auto_number(merged_elements)

        return merged_elements, merged_images, ordered_refs, unknown

    # ──────────────────────────────────────────
    #  УПРАВЛЕНИЕ ЧАСТЯМИ
    # ──────────────────────────────────────────
    def add_part(self, part: ReportPart) -> None:
        self.parts.append(part)

    def move_part(self, idx: int, delta: int) -> None:
        new_idx = idx + delta
        if 0 <= idx < len(self.parts) and 0 <= new_idx < len(self.parts):
            self.parts[idx], self.parts[new_idx] = self.parts[new_idx], self.parts[idx]

    def remove_part(self, idx: int) -> None:
        if 0 <= idx < len(self.parts):
            self.parts.pop(idx)
