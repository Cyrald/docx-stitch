"""
docx-stitch — склейка произвольного числа документов .docx в один за один
проход на уровне пакета OOXML, с сохранением стилей, нумерации, сносок,
колонтитулов и изображений каждого документа.

Быстрый старт::

    from docx_stitch import stitch

    stitch(["title.docx", "body.docx", "appendix.docx"], "report.docx", base=1)
"""
from ._version import __version__
from .errors import StitchError
from .options import (
    ALIGNMENTS,
    BREAK_TYPES,
    SECTION_BREAK_TYPES,
    STYLE_MODES,
    Alignment,
    BreakType,
    DocumentInput,
    OutputTarget,
    Source,
    StyleMode,
)
from .stitcher import TITLE_STYLE_PREFIX, Stitcher, attach_title_page, stitch

__all__ = [
    "ALIGNMENTS",
    "BREAK_TYPES",
    "SECTION_BREAK_TYPES",
    "STYLE_MODES",
    "TITLE_STYLE_PREFIX",
    "Alignment",
    "BreakType",
    "DocumentInput",
    "OutputTarget",
    "Source",
    "StitchError",
    "Stitcher",
    "StyleMode",
    "__version__",
    "attach_title_page",
    "stitch",
]
