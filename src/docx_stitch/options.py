"""Типы и параметры склейки."""
from __future__ import annotations

import os
from dataclasses import dataclass
from typing import BinaryIO, Literal, Union

from .errors import StitchError

__all__ = [
    "ALIGNMENTS",
    "BREAK_TYPES",
    "SECTION_BREAK_TYPES",
    "STYLE_MODES",
    "Alignment",
    "BreakType",
    "DocumentInput",
    "OutputTarget",
    "Source",
    "StyleMode",
]

#: Разрывы раздела OOXML (значения ``w:type``): каждый документ сохраняет
#: собственные параметры страницы и колонтитулы.
SECTION_BREAK_TYPES: frozenset[str] = frozenset({
    "nextPage", "continuous", "evenPage", "oddPage", "nextColumn",
})

#: Все допустимые типы стыка: разрывы раздела, а также ``"page"`` —
#: обычный разрыв страницы без нового раздела и ``"none"`` — без разрыва.
BREAK_TYPES: frozenset[str] = SECTION_BREAK_TYPES | {"page", "none"}

#: ``"isolate"`` — документ выглядит в точности как в собственном файле
#: (его стили копируются под уникальными id); ``"base"`` — документ
#: перенимает одноимённые стили главного документа.
STYLE_MODES: frozenset[str] = frozenset({"isolate", "base"})

ALIGNMENTS: frozenset[str] = frozenset({"left", "center", "right", "both"})

BreakType = Literal[
    "nextPage", "continuous", "evenPage", "oddPage", "nextColumn", "page", "none",
]
StyleMode = Literal["isolate", "base"]
Alignment = Literal["left", "center", "right", "both"]

#: Путь к файлу, его содержимое или открытый бинарный поток.
DocumentInput = Union[str, "os.PathLike[str]", bytes, bytearray, memoryview, BinaryIO]
OutputTarget = Union[str, "os.PathLike[str]", BinaryIO]


def validate_break_type(value: object, what: str = "break_type") -> str:
    if not isinstance(value, str) or value not in BREAK_TYPES:
        valid = ", ".join(sorted(BREAK_TYPES))
        raise StitchError(
            f"Недопустимое значение {what} {value!r}. Допустимые значения: {valid}."
        )
    return value


def validate_style_mode(value: object, what: str = "style_mode") -> str:
    if not isinstance(value, str) or value not in STYLE_MODES:
        valid = ", ".join(sorted(STYLE_MODES))
        raise StitchError(
            f"Недопустимое значение {what} {value!r}. Допустимые значения: {valid}."
        )
    return value


def validate_alignment(value: object) -> str:
    if not isinstance(value, str) or value not in ALIGNMENTS:
        valid = ", ".join(sorted(ALIGNMENTS))
        raise StitchError(
            f"Недопустимое выравнивание разделителя {value!r}. Допустимые значения: {valid}."
        )
    return value


@dataclass
class Source:
    """
    Документ для склейки с индивидуальными параметрами.

    Любое поле со значением ``None`` берётся из общих параметров
    :func:`~docx_stitch.stitch` / :class:`~docx_stitch.Stitcher`.

    Attributes:
        document: путь, ``bytes`` или бинарный поток с содержимым .docx.
        break_after: тип стыка после этого документа (см. :data:`BREAK_TYPES`).
            Для последнего документа игнорируется.
        style_mode: ``"isolate"`` или ``"base"`` (см. :data:`STYLE_MODES`).
            Для главного документа не применяется.
        style_prefix: префикс id стилей в режиме ``"isolate"``
            (по умолчанию ``"_docN_"``, где N — номер документа с 1).
        keep_headers_footers: переносить ли колонтитулы документа. Если
            ``False``, его разделы наследуют колонтитулы предыдущего раздела.
        exclude_from_toc: убрать уровни структуры, чтобы абзацы документа
            не попали в оглавление и область навигации.
        page_number_start: начать нумерацию страниц документа с этого числа.
        name: имя документа для сообщений (по умолчанию — имя файла).
    """

    document: DocumentInput
    break_after: str | None = None
    style_mode: str | None = None
    style_prefix: str | None = None
    keep_headers_footers: bool | None = None
    exclude_from_toc: bool | None = None
    page_number_start: int | None = None
    name: str | None = None

    def __post_init__(self) -> None:
        if self.break_after is not None:
            validate_break_type(self.break_after, "break_after")
        if self.style_mode is not None:
            validate_style_mode(self.style_mode)
        if self.style_prefix is not None and (
            not isinstance(self.style_prefix, str) or not self.style_prefix.strip("_")
        ):
            raise StitchError(
                f"style_prefix должен быть непустой строкой, получено {self.style_prefix!r}."
            )
        if self.page_number_start is not None and (
            isinstance(self.page_number_start, bool)
            or not isinstance(self.page_number_start, int)
            or self.page_number_start < 0
        ):
            raise StitchError(
                "page_number_start должен быть неотрицательным целым числом, "
                f"получено {self.page_number_start!r}."
            )
