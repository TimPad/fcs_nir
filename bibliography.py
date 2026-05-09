"""
Менеджер источников и сквозная нумерация ссылок.

Ссылка в тексте: [@key]. После сборки документа все вхождения
заменяются на [N], где N — порядковый номер источника
по первому упоминанию (как требует ГОСТ Р 7.0.5-2008).
"""

import re
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple


# Паттерн ссылки в тексте: [@ключ] или [@ключ, ключ2] (через запятую)
RE_REF_MARK = re.compile(r'\[@([^\]]+)\]')


@dataclass
class Reference:
    """Один источник в списке использованных источников."""
    key: str        # короткий идентификатор: ivanov2020, gost-7-32, и т.д.
    text: str       # полностью оформленная по ГОСТ Р 7.0.5 запись

    def __post_init__(self):
        self.key = self.key.strip()
        self.text = self.text.strip()


class Bibliography:
    """
    Хранит все источники и осуществляет:
    1. Поиск меток [@key] во всех текстовых полях элементов.
    2. Присвоение сквозных номеров по первому упоминанию.
    3. Замену меток на [N] в тексте элементов.
    4. Возврат упорядоченного списка использованных источников.
    """

    def __init__(self, references: Optional[List[Reference]] = None):
        # OrderedDict-семантика: храним в порядке добавления
        self._refs: Dict[str, Reference] = {}
        if references:
            for ref in references:
                self.add(ref)

    # ──────────────────────────────────────────
    #  УПРАВЛЕНИЕ ИСТОЧНИКАМИ
    # ──────────────────────────────────────────
    def add(self, ref: Reference) -> None:
        if not ref.key:
            return
        self._refs[ref.key] = ref

    def all(self) -> List[Reference]:
        return list(self._refs.values())

    def get(self, key: str) -> Optional[Reference]:
        return self._refs.get(key)

    def __len__(self) -> int:
        return len(self._refs)

    # ──────────────────────────────────────────
    #  ПАРСИНГ ИЗ ТЕКСТА
    # ──────────────────────────────────────────
    @classmethod
    def from_text(cls, text: str) -> "Bibliography":
        """
        Считывает источники из текстового блока.

        Поддерживаемый формат:
            ключ | Полное оформление по ГОСТ Р 7.0.5
            ключ2 | Другая запись...

        Альтернативно — двухстрочный формат:
            [ключ]
            Полное оформление...

        Пустые строки и строки, начинающиеся с #, пропускаются.
        """
        bib = cls()
        if not text or not text.strip():
            return bib

        lines = text.splitlines()
        i = 0
        while i < len(lines):
            line = lines[i].strip()
            if not line or line.startswith("#"):
                i += 1
                continue

            # Формат "ключ | текст"
            if "|" in line:
                key, _, body = line.partition("|")
                bib.add(Reference(key=key.strip(), text=body.strip()))
                i += 1
                continue

            # Формат "[ключ]\n текст..."
            m = re.match(r'^\[([^\]]+)\]\s*$', line)
            if m:
                key = m.group(1).strip()
                body_lines: List[str] = []
                i += 1
                while i < len(lines):
                    nxt = lines[i].rstrip()
                    if not nxt.strip():
                        i += 1
                        break
                    if re.match(r'^\[[^\]]+\]\s*$', nxt.strip()):
                        break
                    if "|" in nxt and "|" in nxt.split(None, 1)[0]:
                        # начало новой записи в формате key | text
                        break
                    body_lines.append(nxt)
                    i += 1
                bib.add(Reference(key=key, text=" ".join(body_lines).strip()))
                continue

            # Не распознано — пропускаем строку
            i += 1

        return bib

    def to_text(self) -> str:
        """Сериализация для повторной правки в UI (формат key | text)."""
        return "\n".join(f"{r.key} | {r.text}" for r in self._refs.values())

    # ──────────────────────────────────────────
    #  ПОИСК ИСПОЛЬЗОВАННЫХ КЛЮЧЕЙ
    # ──────────────────────────────────────────
    @staticmethod
    def extract_keys(text: str) -> List[str]:
        """Возвращает все ключи из вхождений [@key] в тексте, по порядку.
        Поддерживает [@k1, k2] — несколько ключей через запятую."""
        keys: List[str] = []
        for match in RE_REF_MARK.finditer(text):
            inner = match.group(1)
            for k in inner.split(","):
                k = k.strip()
                if k:
                    keys.append(k)
        return keys

    def collect_used(self, elements: List) -> List[str]:
        """Сканирует все элементы документа, возвращает уникальные ключи
        в порядке первого упоминания."""
        seen: Dict[str, None] = {}  # порядок добавления = порядок упоминания
        for el in elements:
            for text_field in self._iter_text_fields(el):
                for key in self.extract_keys(text_field):
                    if key not in seen:
                        seen[key] = None
        return list(seen.keys())

    @staticmethod
    def _iter_text_fields(el) -> List[str]:
        """Возвращает текстовые поля элемента, в которых могут быть метки."""
        fields: List[str] = []
        for attr in ("text", "caption"):
            v = getattr(el, attr, None)
            if isinstance(v, str):
                fields.append(v)
        # Для таблиц: ячейки
        rows = getattr(el, "rows", None)
        if isinstance(rows, list):
            for row in rows:
                if isinstance(row, list):
                    for cell in row:
                        if isinstance(cell, str):
                            fields.append(cell)
        return fields

    # ──────────────────────────────────────────
    #  ЗАМЕНА МЕТОК НА НОМЕРА
    # ──────────────────────────────────────────
    def renumber_and_replace(self, elements: List) -> Tuple[List[Reference], List[str]]:
        """
        Сквозная нумерация: находит все [@key], присваивает номера
        по первому упоминанию, заменяет в тексте на [N].

        Возвращает: (упорядоченный список Reference для печати,
                    список неизвестных ключей).

        Изменяет элементы in-place (поля text, caption и ячейки таблиц).
        """
        used_keys = self.collect_used(elements)

        numbering: Dict[str, int] = {}
        ordered_refs: List[Reference] = []
        unknown: List[str] = []

        for key in used_keys:
            ref = self._refs.get(key)
            if ref is None:
                if key not in unknown:
                    unknown.append(key)
                continue
            numbering[key] = len(ordered_refs) + 1
            ordered_refs.append(ref)

        def replace_in(s: str) -> str:
            def sub(match: re.Match) -> str:
                inner = match.group(1)
                nums: List[str] = []
                for k in (x.strip() for x in inner.split(",")):
                    if not k:
                        continue
                    if k in numbering:
                        nums.append(str(numbering[k]))
                    else:
                        nums.append("?")
                return f"[{', '.join(nums)}]" if nums else match.group(0)
            return RE_REF_MARK.sub(sub, s)

        for el in elements:
            for attr in ("text", "caption"):
                v = getattr(el, attr, None)
                if isinstance(v, str) and "[@" in v:
                    setattr(el, attr, replace_in(v))
            rows = getattr(el, "rows", None)
            if isinstance(rows, list):
                for r_idx, row in enumerate(rows):
                    if isinstance(row, list):
                        rows[r_idx] = [
                            replace_in(c) if isinstance(c, str) and "[@" in c else c
                            for c in row
                        ]

        return ordered_refs, unknown
