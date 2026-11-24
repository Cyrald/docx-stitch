"""Командная строка: ``docx-stitch`` / ``python -m docx_stitch``."""
from __future__ import annotations

import argparse
import logging
import sys
from collections.abc import Sequence
from pathlib import Path

from ._version import __version__
from .errors import StitchError
from .options import ALIGNMENTS, BREAK_TYPES, STYLE_MODES, Source
from .stitcher import stitch

_EPILOG = """\
Номера документов (N) считаются с 1 в порядке перечисления; отрицательные
номера считаются с конца (-1 — последний документ); их передают через «=»:
--base=-1, --page-start=-1=1.

Примеры:
  docx-stitch title.docx body.docx appendix.docx -o report.docx --base 2
  docx-stitch a.docx b.docx c.docx -o out.docx --break page --separator "* * *"
  docx-stitch title.docx body.docx -o out.docx --base 2 --exclude-from-toc-for 1 \\
      --break-after 1=nextPage --page-start 2=1
"""


def _parse_index(raw: str, count: int, option: str) -> int:
    try:
        number = int(raw)
    except ValueError:
        raise StitchError(f"{option}: номер документа должен быть целым числом, получено {raw!r}.") from None
    if number == 0 or abs(number) > count:
        raise StitchError(f"{option}: номер документа {number} вне диапазона 1..{count}.")
    return number - 1 if number > 0 else count + number


def _parse_pairs(values: list[str], count: int, option: str) -> list[tuple[int, str]]:
    pairs: list[tuple[int, str]] = []
    for value in values:
        if "=" not in value:
            raise StitchError(f"{option}: ожидается N=ЗНАЧЕНИЕ, получено {value!r}.")
        raw_index, raw_value = value.split("=", 1)
        pairs.append((_parse_index(raw_index, count, option), raw_value))
    return pairs


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="docx-stitch",
        description="Склеивает несколько документов .docx в один в заданном порядке.",
        epilog=_EPILOG,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("inputs", nargs="+", metavar="FILE", help="документы .docx в итоговом порядке")
    parser.add_argument("-o", "--output", required=True, metavar="FILE", help="путь итогового .docx")
    parser.add_argument(
        "-b", "--base", default="1", metavar="N",
        help="главный документ: его настройки, тема и встроенные стили — основа результата (по умолчанию 1)",
    )
    parser.add_argument(
        "--break", dest="break_type", default="nextPage", choices=sorted(BREAK_TYPES),
        help="стык между документами по умолчанию (по умолчанию nextPage)",
    )
    parser.add_argument(
        "--break-after", action="append", default=[], metavar="N=TYPE",
        help="стык после документа N (можно повторять)",
    )
    parser.add_argument(
        "--style-mode", default="isolate", choices=sorted(STYLE_MODES),
        help="isolate — документы выглядят как в своих файлах; base — перенимают стили главного",
    )
    parser.add_argument(
        "--style-mode-for", action="append", default=[], metavar="N=MODE",
        help="режим стилей для документа N (можно повторять)",
    )
    parser.add_argument(
        "--prefix", action="append", default=[], metavar="N=PREFIX",
        help="префикс id стилей документа N в режиме isolate (можно повторять)",
    )
    parser.add_argument("--separator", metavar="TEXT", help="текст между документами, например \"* * *\"")
    parser.add_argument("--separator-style", metavar="STYLE", help="id или имя стиля разделителя из главного документа")
    parser.add_argument(
        "--separator-align", default="center", choices=sorted(ALIGNMENTS),
        help="выравнивание разделителя (по умолчанию center)",
    )
    parser.add_argument(
        "--no-headers-footers", action="store_true",
        help="не переносить колонтитулы присоединяемых документов",
    )
    parser.add_argument(
        "--exclude-from-toc", action="store_true",
        help="убрать все присоединяемые документы из оглавления и навигации",
    )
    parser.add_argument(
        "--exclude-from-toc-for", action="append", default=[], metavar="N",
        help="убрать документ N из оглавления и навигации (можно повторять)",
    )
    parser.add_argument(
        "--page-start", action="append", default=[], metavar="N=START",
        help="начать нумерацию страниц документа N с числа START (можно повторять)",
    )
    parser.add_argument("-f", "--force", action="store_true", help="перезаписать существующий выходной файл")
    verbosity = parser.add_mutually_exclusive_group()
    verbosity.add_argument("-v", "--verbose", action="store_true", help="подробный журнал")
    verbosity.add_argument("-q", "--quiet", action="store_true", help="только ошибки")
    parser.add_argument("-V", "--version", action="version", version=f"%(prog)s {__version__}")
    return parser


def _build_sources(args: argparse.Namespace) -> tuple[list[Source], int, bool]:
    count = len(args.inputs)
    options: list[dict] = [{} for _ in range(count)]

    for index, value in _parse_pairs(args.break_after, count, "--break-after"):
        if value not in BREAK_TYPES:
            raise StitchError(
                f"--break-after: неизвестный тип {value!r}; допустимые: {', '.join(sorted(BREAK_TYPES))}."
            )
        options[index]["break_after"] = value
    for index, value in _parse_pairs(args.style_mode_for, count, "--style-mode-for"):
        options[index]["style_mode"] = value
    for index, value in _parse_pairs(args.prefix, count, "--prefix"):
        options[index]["style_prefix"] = value
    for index, value in _parse_pairs(args.page_start, count, "--page-start"):
        try:
            options[index]["page_number_start"] = int(value)
        except ValueError:
            raise StitchError(f"--page-start: START должен быть целым числом, получено {value!r}.") from None

    for value in args.exclude_from_toc_for:
        options[_parse_index(value, count, "--exclude-from-toc-for")]["exclude_from_toc"] = True

    sources = [Source(path, **opts) for path, opts in zip(args.inputs, options)]
    return sources, _parse_index(args.base, count, "--base"), args.exclude_from_toc


def main(argv: Sequence[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    level = logging.DEBUG if args.verbose else logging.ERROR if args.quiet else logging.WARNING
    logging.basicConfig(level=level, format="docx-stitch: %(levelname)s: %(message)s")

    try:
        output = Path(args.output)
        if output.exists() and not args.force:
            raise StitchError(f"Файл '{output}' уже существует. Используйте --force для перезаписи.")
        sources, base, exclude_all = _build_sources(args)
        stitch(
            sources,
            output,
            base=base,
            break_type=args.break_type,
            style_mode=args.style_mode,
            separator=args.separator.replace("\\n", "\n") if args.separator else None,
            separator_style=args.separator_style,
            separator_alignment=args.separator_align,
            keep_headers_footers=not args.no_headers_footers,
            exclude_from_toc=exclude_all,
        )
    except StitchError as exc:
        print(f"docx-stitch: ошибка: {exc}", file=sys.stderr)
        return 1
    except KeyboardInterrupt:
        return 130

    if not args.quiet:
        print(f"Готово: {output} ({len(sources)} документ(ов))")
    return 0


if __name__ == "__main__":
    sys.exit(main())
