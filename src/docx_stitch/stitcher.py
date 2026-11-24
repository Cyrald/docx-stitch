"""
Склейка произвольного числа документов .docx за один проход.

Один из документов назначается главным (``base``): его пакет становится
основой результата — настройки, тема, docDefaults, встроенные стили
(Heading1, TOC1, Hyperlink, …) и колонтитулы сохраняются как есть.
Все остальные документы импортируются в этот пакет по очереди, а затем
тела всех документов выстраиваются в заданном порядке — без промежуточных
файлов вида ``A + B = tmp; tmp + C = D``.
"""
from __future__ import annotations

import copy
import dataclasses
import io
import os
import posixpath
import uuid
import zipfile
from collections.abc import Iterable, Iterator, Sequence
from dataclasses import dataclass
from pathlib import Path
from typing import BinaryIO

from lxml import etree as ET

from . import _ooxml as X
from .errors import StitchError
from .options import (
    SECTION_BREAK_TYPES,
    DocumentInput,
    OutputTarget,
    Source,
    validate_alignment,
    validate_break_type,
    validate_style_mode,
)

__all__ = ["TITLE_STYLE_PREFIX", "Stitcher", "attach_title_page", "stitch"]

logger = X.logger

_W = f"{{{X._WNS}}}"

#: Префикс id стилей титульного листа в :func:`attach_title_page`;
#: отображаемые имена его стилей получают «title: ».
TITLE_STYLE_PREFIX = "_title_"

_HF_REL_SUFFIXES = (X._HEADER_REL_SUFFIX, X._FOOTER_REL_SUFFIX)

# Элементы sectPr, которые по схеме идут после w:pgNumType.
_AFTER_PG_NUM_TYPE = frozenset({
    "cols", "formProt", "vAlign", "noEndnote", "titlePg", "textDirection",
    "bidi", "rtlGutter", "docGrid", "printerSettings", "sectPrChange",
})


# ── Входные данные ──────────────────────────────────────────────────────────

@dataclass
class _Resolved:
    """Документ с загруженным содержимым и окончательными параметрами."""

    index: int
    name: str
    data: bytes
    path: Path | None
    break_after: str
    style_mode: str
    style_prefix: str
    keep_headers_footers: bool
    exclude_from_toc: bool
    page_number_start: int | None


@dataclass
class _Block:
    """Тело одного документа, готовое к расстановке в итоговом порядке."""

    name: str
    elems: list[ET._Element]
    sect_pr: ET._Element
    filler_style: str | None


def _load_input(document: DocumentInput, index: int) -> tuple[str, bytes, Path | None]:
    """Читает документ из пути, байтов или бинарного потока."""
    fallback_name = f"документ №{index + 1}"
    path: Path | None = None

    if isinstance(document, (bytes, bytearray, memoryview)):
        data = bytes(document)
        name = fallback_name
    elif hasattr(document, "read"):
        raw = document.read()
        if not isinstance(raw, (bytes, bytearray)):
            raise StitchError(
                f"Поток для {fallback_name} должен быть открыт в двоичном режиме ('rb')."
            )
        data = bytes(raw)
        stream_name = getattr(document, "name", None)
        name = (
            Path(stream_name).name
            if isinstance(stream_name, (str, os.PathLike)) else fallback_name
        )
    else:
        try:
            path = Path(document)  # type: ignore[arg-type]
        except TypeError as exc:
            raise StitchError(
                f"Некорректный документ №{index + 1}: ожидается путь, bytes "
                f"или бинарный поток, получено {type(document).__name__}."
            ) from exc
        X._validate_input_file(path)
        try:
            data = path.read_bytes()
        except OSError as exc:
            raise StitchError(f"Не удалось прочитать '{path}': {exc}") from exc
        name = path.name

    if not zipfile.is_zipfile(io.BytesIO(data)):
        raise StitchError(
            f"'{name}' не является корректным файлом .docx (это не ZIP-архив). "
            "Возможно, файл повреждён."
        )
    return name, data, path


def _as_source(document: DocumentInput | Source, options: dict) -> Source:
    if isinstance(document, Source):
        return dataclasses.replace(document, **options) if options else document
    return Source(document, **options)


def _resolve_sources(
    sources: Sequence[Source],
    *,
    break_type: str,
    style_mode: str,
    keep_headers_footers: bool,
    exclude_from_toc: bool,
) -> list[_Resolved]:
    resolved: list[_Resolved] = []
    for i, src in enumerate(sources):
        name, data, path = _load_input(src.document, i)
        resolved.append(_Resolved(
            index=i,
            name=src.name or name,
            data=data,
            path=path,
            break_after=src.break_after or break_type,
            style_mode=src.style_mode or style_mode,
            style_prefix=src.style_prefix or f"_doc{i + 1}_",
            keep_headers_footers=(
                keep_headers_footers if src.keep_headers_footers is None
                else src.keep_headers_footers
            ),
            exclude_from_toc=(
                exclude_from_toc if src.exclude_from_toc is None
                else src.exclude_from_toc
            ),
            page_number_start=src.page_number_start,
        ))
    return resolved


# ── Общий пакет ─────────────────────────────────────────────────────────────

class _Package:
    """
    Пакет главного документа, в который последовательно импортируются
    остальные. Хранит текущее состояние общих частей (стили, нумерация,
    сноски, связи, типы содержимого) и новые файлы пакета.
    """

    def __init__(self, data: bytes, name: str) -> None:
        self.data = data
        self.name = name

        with zipfile.ZipFile(io.BytesIO(data)) as zf:
            names = set(zf.namelist())
            for required in (X._DOC_PATH, X._CT_PATH, X._STYLES_PATH):
                if required not in names:
                    raise StitchError(
                        f"В '{name}' отсутствует {required}. Возможно, файл повреждён."
                    )

            def read(part: str) -> bytes | None:
                return X._read_if_present(zf, names, part)

            self.doc_xml: bytes = zf.read(X._DOC_PATH)
            self.styles: bytes = zf.read(X._STYLES_PATH)
            self.ct: bytes = zf.read(X._CT_PATH)
            self.rels = read(X._RELS_PATH)
            self.numbering = read(X._NUMBERING_PATH)
            self.font = read(X._FONT_TABLE_PATH)
            self.footnotes = read(X._FOOTNOTES_PATH)
            self.endnotes = read(X._ENDNOTES_PATH)
            self.footnotes_rels = read(X._FOOTNOTES_RELS_PATH)
            self.endnotes_rels = read(X._ENDNOTES_RELS_PATH)

            # Колонтитулы и сноски главного документа нужны только для того,
            # чтобы id рисунков и закладок импортируемых документов
            # не пересеклись и с ними.
            aux_roots: list[ET._Element] = []
            aux_parts: list[bytes] = [
                p for p in (self.footnotes, self.endnotes) if p is not None
            ]
            if self.rels is not None:
                for rtype, target, mode in X._parse_rels(self.rels).values():
                    if mode != "External" and rtype.endswith(_HF_REL_SUFFIXES):
                        part = X._read_if_present(zf, names, X._resolve_zip_path(target))
                        if part is not None:
                            aux_parts.append(part)
            for part in aux_parts:
                try:
                    aux_roots.append(X._fromstring(part))
                except StitchError:
                    logger.warning(
                        "'%s': повреждённая часть пакета пропущена при проверке "
                        "уникальности id.", name,
                    )

        self.original_parts = {
            "numbering": self.numbering is not None,
            "footnotes": self.footnotes is not None,
            "endnotes": self.endnotes is not None,
            "font": self.font is not None,
        }

        styles = X._parse_styles(self.styles)
        self.style_ids: set[str] = set(styles)
        self.style_names: dict[str, str] = {}
        for sid, el in styles.items():
            name_el = el.find(f"{_W}name")
            if name_el is not None and name_el.get(f"{_W}val"):
                self.style_names.setdefault(name_el.get(f"{_W}val").lower(), sid)
        self.doc_defaults = X._read_doc_defaults(self.styles, None)

        elems, sect_pr = X._parse_body_elems(self.doc_xml, name)
        self.base_block = _Block(name, elems, sect_pr, None)
        self.collision: list[ET._Element] = list(elems) + [sect_pr] + aux_roots
        self.src_doc_xmls: list[bytes] = []
        self.extra_files: dict[str, bytes] = {}

    # -- стили ---------------------------------------------------------------

    def add_styles(
        self,
        styles: dict[str, ET._Element],
        rename_map: dict[str, str],
        display_prefix: str | None,
    ) -> None:
        """Добавляет стили в styles.xml под новыми id и уникальными именами."""
        if not styles:
            return
        root = X._fromstring(self.styles)
        for sid, style_el in styles.items():
            new_el = copy.deepcopy(style_el)
            new_sid = rename_map[sid]
            new_el.set(f"{_W}styleId", new_sid)
            new_el.attrib.pop(f"{_W}default", None)
            name_el = new_el.find(f"{_W}name")
            if name_el is not None and name_el.get(f"{_W}val"):
                base_name = f"{display_prefix or ''}{name_el.get(f'{_W}val')}"
                new_name, n = base_name, 1
                while new_name.lower() in self.style_names:
                    n += 1
                    new_name = f"{base_name} ({n})"
                name_el.set(f"{_W}val", new_name)
                self.style_names[new_name.lower()] = new_sid
            root.append(new_el)
            self.style_ids.add(new_sid)
        self.styles = X._xml_bytes(root)

    def find_style(self, style: str) -> str | None:
        """id стиля главного документа по id или отображаемому имени."""
        if style in self.style_ids:
            return style
        return self.style_names.get(style.lower())

    # -- связи ---------------------------------------------------------------

    @staticmethod
    def _extend_rels(
        current: bytes | None,
        entries: list[tuple[str, str, str, str | None]],
    ) -> bytes | None:
        if not entries:
            return current
        existing = X._parse_rels(current) if current else {}
        return X._build_rels_xml(existing, entries)


def _import_source(pkg: _Package, src: _Resolved) -> _Block:
    """
    Импортирует документ ``src`` в общий пакет и возвращает его тело.

    Повторяет шаги исходного алгоритма md2gost (перенумерация связей,
    очистка, переименование стилей, перенумерация id, слияние нумерации,
    сносок, шрифтов и стилей), но изменяет накопленное состояние пакета,
    поэтому каждый следующий документ учитывает все предыдущие.
    """
    isolate = src.style_mode == "isolate"
    skip_hf = not src.keep_headers_footers

    with zipfile.ZipFile(io.BytesIO(src.data)) as szf:
        s_names = set(szf.namelist())
        if X._DOC_PATH not in s_names:
            raise StitchError(
                f"В '{src.name}' отсутствует {X._DOC_PATH}. Возможно, файл повреждён."
            )

        def read(part: str) -> bytes | None:
            return X._read_if_present(szf, s_names, part)

        s_doc_xml = szf.read(X._DOC_PATH)
        s_styles_xml = read(X._STYLES_PATH)
        s_rels_xml = read(X._RELS_PATH)
        s_ct_xml = read(X._CT_PATH)
        s_media = {n: szf.read(n) for n in s_names if n.startswith("word/media/")}
        s_numbering_xml = read(X._NUMBERING_PATH)
        s_font_xml = read(X._FONT_TABLE_PATH)
        s_footnotes_xml = read(X._FOOTNOTES_PATH)
        s_endnotes_xml = read(X._ENDNOTES_PATH)
        s_footnotes_rels_xml = read(X._FOOTNOTES_RELS_PATH)
        s_endnotes_rels_xml = read(X._ENDNOTES_RELS_PATH)

        s_rels = X._parse_rels(s_rels_xml) if s_rels_xml is not None else {}
        hf_rids = {
            rid for rid, (rtype, _t, mode) in s_rels.items()
            if mode != "External" and rtype.endswith(_HF_REL_SUFFIXES)
        }

        s_hf_raw: dict[str, bytes] = {}
        if not skip_hf:
            for rid in hf_rids:
                hf_zip_path = X._resolve_zip_path(s_rels[rid][1])
                if hf_zip_path in s_names:
                    s_hf_raw[hf_zip_path] = szf.read(hf_zip_path)
                hf_rels_path = (
                    f"{posixpath.dirname(hf_zip_path)}/_rels/"
                    f"{posixpath.basename(hf_zip_path)}.rels"
                )
                if hf_rels_path in s_names:
                    s_hf_raw[hf_rels_path] = szf.read(hf_rels_path)

        s_theme_xml: bytes | None = None
        for rtype, target, mode in s_rels.values():
            if mode != "External" and rtype.endswith("/theme"):
                s_theme_xml = read(X._resolve_zip_path(target))
                break

    # 1. Новые имена медиафайлов и колонтитулов, перенумерация rId.
    fname_map = {
        mf: f"word/media/{uuid.uuid4().hex}_{Path(mf).name}" for mf in s_media
    }
    rid_map, new_entries, hf_fname_map, hf_content_types = X._compute_rels_remap(
        fname_map, s_rels_xml, pkg.rels,
        s_header_footer=s_hf_raw, skip_header_footer=skip_hf,
    )
    foot_rid_map, foot_new_entries, _, _ = X._compute_rels_remap(
        fname_map, s_footnotes_rels_xml, pkg.footnotes_rels,
    )
    end_rid_map, end_new_entries, _, _ = X._compute_rels_remap(
        fname_map, s_endnotes_rels_xml, pkg.endnotes_rels,
    )

    src_elems, src_sect_pr = X._parse_body_elems(s_doc_xml, src.name)
    hf_elems: dict[str, ET._Element] = {
        zip_path: X._fromstring(raw)
        for zip_path, raw in s_hf_raw.items()
        if not zip_path.endswith(".rels")
    }
    hf_list = list(hf_elems.values())
    body_and_sect = src_elems + [src_sect_pr]

    # 2. Очистка: printerSettings, непереносимые объекты, примечания.
    X._strip_printer_settings_refs(src_elems, [src_sect_pr], hf_list)
    dropped_ids = X._dropped_rel_ids(s_rels_xml, rid_map)
    if skip_hf:
        X._strip_dropped_rel_refs(body_and_sect, dropped_ids & hf_rids)
        dropped_ids -= hf_rids
    dropped_refs = X._strip_dropped_rel_refs(body_and_sect, dropped_ids)
    if dropped_refs:
        logger.warning(
            "'%s': %d объект(ов) неподдерживаемого типа (диаграммы, OLE-объекты, "
            "SmartArt и т.п.) удалены из итогового документа.",
            src.name, dropped_refs,
        )
    if X._strip_comment_anchors(src_elems + hf_list):
        logger.warning(
            "'%s': примечания (комментарии) Word не переносятся и удалены "
            "из итогового документа.", src.name,
        )

    src_styles = X._parse_styles(s_styles_xml)
    theme_fonts = X._read_theme_fonts(s_theme_xml) if isolate else {}
    style_parts = src_elems + hf_list + list(src_styles.values())
    if isolate:
        X._detach_theme_formatting(style_parts, theme_fonts)
    if src.exclude_from_toc:
        X._strip_outline_levels(style_parts)

    # 3. Переименование стилей.
    rename_map: dict[str, str] = {}
    if isolate:
        styles_to_add = src_styles
        used = set(pkg.style_ids)
        for sid in src_styles:
            candidate, suffix = src.style_prefix + sid, 0
            while candidate in used:
                suffix += 1
                candidate = f"{src.style_prefix}{sid}_{suffix}"
            rename_map[sid] = candidate
            used.add(candidate)
    else:
        styles_to_add = {}
        for sid, el in src_styles.items():
            rename_map[sid] = sid
            if sid not in pkg.style_ids:
                styles_to_add[sid] = el

    implicit_target: str | None = None
    if isolate:
        default_sid = X._find_default_paragraph_style_id(src_styles)
        if default_sid is not None:
            implicit_target = rename_map[default_sid]
        else:
            implicit_target = X._add_synthetic_default_style(
                src_styles, pkg.style_ids, rename_map, src.style_prefix,
            )

    for elem in body_and_sect:
        X._rename_style_refs(elem, rename_map)

    if isolate:
        targets = src_elems + hf_list
        X._inject_implicit_style(targets, implicit_target, include_styled_tables=True)
        char_sid = X._find_default_style_id(src_styles, "character", "DefaultParagraphFont")
        X._inject_implicit_run_style(
            targets, rename_map.get(char_sid) if char_sid is not None else None,
        )
        table_sid = X._find_default_style_id(src_styles, "table", "TableNormal")
        X._inject_implicit_table_style(
            targets, rename_map.get(table_sid) if table_sid is not None else None,
        )

    # 4. Уникальность id рисунков и закладок во всём пакете.
    X._remap_drawing_ids(src_elems + hf_list, pkg.collision)
    bookmark_elems = src_elems + hf_list
    bookmark_name_map = X._remap_bookmark_ids(bookmark_elems, pkg.collision)
    X._rename_bookmark_field_refs(bookmark_elems, bookmark_name_map)

    for hf_root in hf_list:
        X._rename_style_refs(hf_root, rename_map)
    for style_el in src_styles.values():
        X._rename_style_refs(style_el, rename_map)

    declared_ids = (
        set(rename_map.values()) if isolate
        else pkg.style_ids | set(rename_map.values())
    )
    X._resolve_undeclared_style_refs(
        src_elems + hf_list + list(styles_to_add.values()),
        declared_ids=declared_ids,
        paragraph_fallback=implicit_target,
    )

    if rid_map:
        X._rename_rids(body_and_sect, rid_map)

    # 5. Нумерация, сноски, стили, шрифты.
    merged_numbering, num_remap = X._merge_numbering(
        pkg.numbering, s_numbering_xml, src_elems, rename_map,
        extra_elems=hf_list + list(styles_to_add.values()),
    )
    if merged_numbering is not None:
        pkg.numbering = merged_numbering

    collision = pkg.collision + src_elems + hf_list
    for attr, src_notes, ref_local, n_rid_map, n_rels_xml in (
        ("footnotes", s_footnotes_xml, "footnoteReference", foot_rid_map, s_footnotes_rels_xml),
        ("endnotes", s_endnotes_xml, "endnoteReference", end_rid_map, s_endnotes_rels_xml),
    ):
        merged_notes = X._merge_notes(
            getattr(pkg, attr), src_notes, src_elems,
            ref_tag=f"{_W}{ref_local}",
            rid_map=n_rid_map,
            dropped_rids=X._dropped_rel_ids(n_rels_xml, n_rid_map),
            rename_map=rename_map,
            existing_elems=collision,
            style_declared_ids=declared_ids,
            style_fallback=implicit_target,
            num_remap=num_remap,
        )
        if merged_notes is not None:
            setattr(pkg, attr, merged_notes)
    pkg.collision = collision
    pkg.footnotes_rels = pkg._extend_rels(pkg.footnotes_rels, foot_new_entries)
    pkg.endnotes_rels = pkg._extend_rels(pkg.endnotes_rels, end_new_entries)

    if isolate:
        X._isolate_root_paragraph_styles(
            src_styles,
            title_defaults=X._read_doc_defaults(s_styles_xml, theme_fonts),
            document_defaults=pkg.doc_defaults,
        )
    pkg.add_styles(
        styles_to_add, rename_map,
        display_prefix=(src.style_prefix.strip("_") + ": ") if isolate else None,
    )

    merged_font = X._merge_font_table(pkg.font, s_font_xml)
    if merged_font is not None:
        pkg.font = merged_font

    # 6. Связи, типы содержимого и новые файлы пакета.
    pkg.rels = pkg._extend_rels(pkg.rels, new_entries)
    if fname_map:
        pkg.ct = X._merge_content_types(pkg.ct, s_ct_xml, list(s_media), fname_map)
        pkg.extra_files.update({fname_map[zp]: data for zp, data in s_media.items()})
    if hf_fname_map:
        pkg.ct = X._add_hf_content_types(pkg.ct, hf_fname_map, hf_content_types)
        pkg.extra_files.update(
            X._build_hf_extra_files(hf_fname_map, hf_elems, s_hf_raw, fname_map)
        )

    pkg.src_doc_xmls.append(s_doc_xml)
    return _Block(src.name, src_elems, src_sect_pr, implicit_target)


# ── Расстановка блоков ──────────────────────────────────────────────────────

def _first_section_sect_pr(block: _Block) -> ET._Element:
    """sectPr первого раздела блока (тот, что определяет начало блока)."""
    for el in block.elems:
        for node in el.iter(f"{_W}sectPr"):
            parent = node.getparent()
            if parent is not None and parent.tag == f"{_W}pPr":
                return node
    return block.sect_pr


def _set_page_number_start(sect_pr: ET._Element, start: int) -> None:
    el = sect_pr.find(f"{_W}pgNumType")
    if el is None:
        el = ET.Element(f"{_W}pgNumType")
        for i, child in enumerate(sect_pr):
            if X._local(child.tag) in _AFTER_PG_NUM_TYPE:
                sect_pr.insert(i, el)
                break
        else:
            sect_pr.append(el)
    el.set(f"{_W}start", str(start))


def _add_page_break(prev: _Block, nxt: _Block) -> None:
    """Обычный разрыв страницы перед блоком nxt (без нового раздела)."""
    first = nxt.elems[0] if nxt.elems else None
    if first is not None and first.tag == f"{_W}p":
        p_pr = first.find(f"{_W}pPr")
        if p_pr is None:
            p_pr = ET.Element(f"{_W}pPr")
            first.insert(0, p_pr)
        existing = p_pr.find(f"{_W}pageBreakBefore")
        if existing is not None:
            existing.attrib.pop(f"{_W}val", None)
        else:
            ET.SubElement(p_pr, f"{_W}pageBreakBefore")
            X._sort_children(p_pr, X._PPR_ORDER)
        return
    filler = X._make_filler_paragraph(prev.filler_style)
    run = ET.SubElement(filler, f"{_W}r")
    ET.SubElement(run, f"{_W}br").set(f"{_W}type", "page")
    prev.elems.append(filler)


def _make_separator(text: str, style_id: str | None, alignment: str) -> list[ET._Element]:
    paragraphs: list[ET._Element] = []
    for line in text.split("\n"):
        p = ET.Element(f"{_W}p")
        p_pr = ET.SubElement(p, f"{_W}pPr")
        if style_id:
            ET.SubElement(p_pr, f"{_W}pStyle").set(f"{_W}val", style_id)
        ET.SubElement(p_pr, f"{_W}jc").set(f"{_W}val", alignment)
        if line:
            run = ET.SubElement(p, f"{_W}r")
            t = ET.SubElement(run, f"{_W}t")
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            try:
                t.text = line
            except ValueError as exc:
                raise StitchError(
                    f"Текст разделителя содержит недопустимые для XML символы: {exc}"
                ) from exc
        paragraphs.append(p)
    return paragraphs


def _arrange(
    blocks: list[_Block],
    sources: list[_Resolved],
    separator: str | None,
    separator_style: str | None,
    separator_alignment: str,
) -> ET._Element:
    """Расставляет разрывы и разделители; возвращает итоговый sectPr тела."""
    last = len(blocks) - 1
    first_sects = [_first_section_sect_pr(b) for b in blocks]

    for i, (block, src) in enumerate(zip(blocks, sources)):
        if src.page_number_start is None:
            continue
        target = first_sects[i]
        if i < last and src.break_after in ("page", "none") and target is block.sect_pr:
            logger.warning(
                "'%s': page_number_start проигнорирован — после документа стоит "
                "стык %r, и его раздел сливается со следующим.",
                block.name, src.break_after,
            )
            continue
        _set_page_number_start(target, src.page_number_start)

    for i in range(last):
        block, brk = blocks[i], sources[i].break_after
        if separator is not None:
            block.elems.extend(_make_separator(separator, separator_style, separator_alignment))
        if brk in SECTION_BREAK_TYPES:
            # w:type описывает, как раздел начинается относительно
            # предыдущего, поэтому тип стыка задаётся первому разделу
            # следующего блока.
            X._set_section_type(first_sects[i + 1], brk)
            X._embed_sect_pr(block.elems, block.sect_pr, filler_style_id=block.filler_style)
        elif brk == "page":
            _add_page_break(block, blocks[i + 1])

    return blocks[last].sect_pr


def _build_document(pkg: _Package, blocks: list[_Block], final_sect_pr: ET._Element) -> bytes:
    root = X._fromstring(pkg.doc_xml)
    for src_xml in pkg.src_doc_xmls:
        try:
            src_root = X._fromstring(src_xml)
            extra_nsmap = X._resolve_extra_nsmap(root.nsmap, X._collect_full_nsmap(src_root))
            if extra_nsmap:
                new_root = ET.Element(root.tag, nsmap={**root.nsmap, **extra_nsmap})
                for attr, val in root.attrib.items():
                    new_root.set(attr, val)
                for child in list(root):
                    new_root.append(child)
                root = new_root
            if X._MC_NS in root.nsmap.values():
                X._merge_ignorable(root, src_root)
        except Exception as exc:
            logger.warning(
                "Не удалось расширить карту пространств имён: %s", exc, exc_info=True,
            )

    body = root.find(f"{_W}body")
    if body is None:
        raise StitchError(
            f"В '{pkg.name}' отсутствует элемент <w:body>. Возможно, файл повреждён."
        )
    for child in list(body):
        body.remove(child)
    for block in blocks:
        for el in block.elems:
            body.append(el)
    body.append(final_sect_pr)
    return X._xml_bytes(root)


def _finalize(pkg: _Package, document_xml: bytes) -> bytes:
    """Регистрирует новые части и записывает итоговый архив в память."""
    out_rels = X._parse_rels(pkg.rels) if pkg.rels else {}
    used = set(out_rels)
    next_n = X._max_rid(out_rels) + 1
    rels_to_add: list[tuple[str, str, str, str | None]] = []
    for attr, rel_type, target, ct_type in (
        ("numbering", X._NUMBERING_REL_TYPE, "numbering.xml", X._NUMBERING_CT_TYPE),
        ("footnotes", X._FOOTNOTES_REL_TYPE, "footnotes.xml", X._FOOTNOTES_CT_TYPE),
        ("endnotes", X._ENDNOTES_REL_TYPE, "endnotes.xml", X._ENDNOTES_CT_TYPE),
        ("font", X._FONT_TABLE_REL_TYPE, "fontTable.xml", X._FONT_TABLE_CT_TYPE),
    ):
        if pkg.original_parts[attr] or getattr(pkg, attr) is None:
            continue
        rid, next_n = X._next_free_rid(used, next_n)
        used.add(rid)
        rels_to_add.append((rid, rel_type, target, None))
        pkg.ct = X._add_ct_override(pkg.ct, f"/word/{target}", ct_type)
    if rels_to_add:
        pkg.rels = X._build_rels_xml(out_rels, rels_to_add)

    replacements: dict[str, bytes] = {
        X._DOC_PATH: document_xml,
        X._STYLES_PATH: pkg.styles,
        X._CT_PATH: pkg.ct,
    }
    for part, data in (
        (X._RELS_PATH, pkg.rels),
        (X._NUMBERING_PATH, pkg.numbering),
        (X._FONT_TABLE_PATH, pkg.font),
        (X._FOOTNOTES_PATH, pkg.footnotes),
        (X._ENDNOTES_PATH, pkg.endnotes),
        (X._FOOTNOTES_RELS_PATH, pkg.footnotes_rels),
        (X._ENDNOTES_RELS_PATH, pkg.endnotes_rels),
    ):
        if data is not None:
            replacements[part] = data

    buf = io.BytesIO()
    with (
        zipfile.ZipFile(io.BytesIO(pkg.data), "r") as base_zip,
        zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as out_zip,
    ):
        written: set[str] = set()
        for item in base_zip.infolist():
            if item.filename in written:
                continue
            data = replacements.get(item.filename)
            out_zip.writestr(item, data if data is not None else base_zip.read(item))
            written.add(item.filename)
        for zip_path, data in {**replacements, **pkg.extra_files}.items():
            if zip_path not in written:
                out_zip.writestr(zip_path, data)
                written.add(zip_path)
    return buf.getvalue()


# ── Публичный API ───────────────────────────────────────────────────────────

def _prepare_output(output: OutputTarget | None, inputs: Iterable[_Resolved]) -> Path | BinaryIO | None:
    if output is None or hasattr(output, "write"):
        return output  # type: ignore[return-value]
    try:
        path = Path(output)  # type: ignore[arg-type]
    except TypeError as exc:
        raise StitchError(f"Некорректный путь сохранения: {exc}") from exc
    X._validate_output_path(path)
    resolved = path.resolve()
    if any(src.path is not None and src.path.resolve() == resolved for src in inputs):
        raise StitchError(
            f"Выходной файл '{path}' не должен совпадать ни с одним из входных документов."
        )
    return path


def _write_output(data: bytes, target: Path | BinaryIO | None) -> bytes | None:
    if target is None:
        return data
    if isinstance(target, Path):
        tmp = target.with_name(f"{target.name}.{uuid.uuid4().hex}.tmp")
        try:
            tmp.write_bytes(data)
            tmp.replace(target)
        except BaseException:
            tmp.unlink(missing_ok=True)
            raise
        return None
    target.write(data)
    return None


def _run(
    sources: Sequence[Source],
    output: OutputTarget | None,
    *,
    base: int | Source,
    break_type: str,
    style_mode: str,
    separator: str | None,
    separator_style: str | None,
    separator_alignment: str,
    keep_headers_footers: bool,
    exclude_from_toc: bool,
) -> bytes | None:
    validate_break_type(break_type)
    validate_style_mode(style_mode)
    validate_alignment(separator_alignment)
    if separator is not None and not isinstance(separator, str):
        raise StitchError(f"separator должен быть строкой, получено {type(separator).__name__}.")
    if not sources:
        raise StitchError("Не передано ни одного документа для склейки.")

    if isinstance(base, Source):
        base_index = next((i for i, s in enumerate(sources) if s is base), None)
        if base_index is None:
            raise StitchError("Главный документ (base) отсутствует в списке документов.")
    elif isinstance(base, int) and not isinstance(base, bool):
        base_index = base + len(sources) if base < 0 else base
        if not 0 <= base_index < len(sources):
            raise StitchError(
                f"Индекс главного документа base={base} вне диапазона "
                f"(документов: {len(sources)})."
            )
    else:
        raise StitchError(f"base должен быть индексом документа, получено {base!r}.")

    resolved = _resolve_sources(
        sources,
        break_type=break_type,
        style_mode=style_mode,
        keep_headers_footers=keep_headers_footers,
        exclude_from_toc=exclude_from_toc,
    )
    target = _prepare_output(output, resolved)

    base_src = resolved[base_index]
    logger.info(
        "Склейка %d документ(ов); главный: '%s'.", len(resolved), base_src.name,
    )
    pkg = _Package(base_src.data, base_src.name)

    sep_style: str | None = None
    if separator_style is not None:
        sep_style = pkg.find_style(separator_style)
        if sep_style is None:
            logger.warning(
                "Стиль разделителя %r не найден в главном документе '%s' — "
                "используется стиль абзаца по умолчанию.", separator_style, base_src.name,
            )

    blocks: list[_Block] = []
    for src in resolved:
        if src.index == base_index:
            blocks.append(pkg.base_block)
        else:
            logger.debug("Импорт '%s' (режим стилей: %s).", src.name, src.style_mode)
            blocks.append(_import_source(pkg, src))

    final_sect_pr = _arrange(blocks, resolved, separator, sep_style, separator_alignment)
    document_xml = _build_document(pkg, blocks, final_sect_pr)
    return _write_output(_finalize(pkg, document_xml), target)


def stitch(
    documents: Sequence[DocumentInput | Source],
    output: OutputTarget | None = None,
    *,
    base: int = 0,
    break_type: str = "nextPage",
    style_mode: str = "isolate",
    separator: str | None = None,
    separator_style: str | None = None,
    separator_alignment: str = "center",
    keep_headers_footers: bool = True,
    exclude_from_toc: bool = False,
) -> bytes | None:
    """
    Склеивает документы ``documents`` в указанном порядке за один проход.

    Args:
        documents: документы в итоговом порядке — пути, ``bytes``, бинарные
            потоки или :class:`Source` с индивидуальными параметрами.
            Один и тот же файл может встречаться несколько раз.
        output: путь или бинарный поток для результата. Если ``None``,
            результат возвращается как ``bytes``.
        base: индекс главного документа (можно отрицательный). Его пакет —
            основа результата: настройки, тема, docDefaults, встроенные стили
            и колонтитулы. Позиция главного документа в тексте не меняется.
        break_type: стык между документами по умолчанию —
            ``"nextPage"``, ``"continuous"``, ``"evenPage"``, ``"oddPage"``,
            ``"nextColumn"`` (разрыв раздела: у каждого документа свои поля,
            ориентация и колонтитулы), ``"page"`` (разрыв страницы без нового
            раздела) или ``"none"`` (вплотную).
        style_mode: ``"isolate"`` — каждый документ выглядит в точности как
            в своём файле; ``"base"`` — одноимённые стили берутся из главного.
        separator: текст, вставляемый между документами (например ``"* * *"``);
            ``"\\n"`` разбивает его на несколько абзацев.
        separator_style: id или имя стиля главного документа для разделителя.
        separator_alignment: ``"left"``, ``"center"``, ``"right"`` или ``"both"``.
        keep_headers_footers: переносить колонтитулы присоединяемых документов.
        exclude_from_toc: убрать присоединяемые документы из оглавления
            и области навигации.

    Raises:
        StitchError: некорректные входные данные или параметры
            (подкласс :class:`ValueError`).
    """
    if isinstance(documents, (str, bytes, bytearray, os.PathLike)) or not isinstance(
        documents, Iterable
    ):
        raise StitchError(
            "documents должен быть списком документов, например [\"a.docx\", \"b.docx\"]."
        )
    sources = [_as_source(d, {}) for d in documents]
    return _run(
        sources, output,
        base=base,
        break_type=break_type,
        style_mode=style_mode,
        separator=separator,
        separator_style=separator_style,
        separator_alignment=separator_alignment,
        keep_headers_footers=keep_headers_footers,
        exclude_from_toc=exclude_from_toc,
    )


class Stitcher:
    """
    Построитель склейки с изменяемым списком документов.

    >>> (Stitcher(break_type="nextPage")
    ...     .add("title.docx", exclude_from_toc=True)
    ...     .add("body.docx", base=True)
    ...     .add("appendix.docx", page_number_start=1)
    ...     .save("report.docx"))

    Общие параметры конструктора совпадают с :func:`stitch`;
    параметры отдельных документов — с полями :class:`Source`.
    """

    def __init__(
        self,
        documents: Iterable[DocumentInput | Source] = (),
        *,
        base: int = 0,
        break_type: str = "nextPage",
        style_mode: str = "isolate",
        separator: str | None = None,
        separator_style: str | None = None,
        separator_alignment: str = "center",
        keep_headers_footers: bool = True,
        exclude_from_toc: bool = False,
    ) -> None:
        validate_break_type(break_type)
        validate_style_mode(style_mode)
        validate_alignment(separator_alignment)
        self.break_type = break_type
        self.style_mode = style_mode
        self.separator = separator
        self.separator_style = separator_style
        self.separator_alignment = separator_alignment
        self.keep_headers_footers = keep_headers_footers
        self.exclude_from_toc = exclude_from_toc
        self._items: list[Source] = [_as_source(doc, {}) for doc in documents]
        self._base: int | Source = base

    # -- список документов -----------------------------------------------------

    def add(self, document: DocumentInput | Source, *, base: bool = False, **options: object) -> Stitcher:
        """Добавляет документ в конец. ``base=True`` делает его главным."""
        return self.insert(len(self._items), document, base=base, **options)

    append = add

    def prepend(self, document: DocumentInput | Source, *, base: bool = False, **options: object) -> Stitcher:
        """Добавляет документ в начало (например, титульный лист)."""
        return self.insert(0, document, base=base, **options)

    def insert(
        self, index: int, document: DocumentInput | Source, *, base: bool = False, **options: object,
    ) -> Stitcher:
        """Вставляет документ в позицию ``index``."""
        if isinstance(self._base, int):
            # Запоминаем главный документ по ссылке, чтобы вставка
            # не сдвинула выбор на другой документ.
            self._pin_base()
        src = _as_source(document, options)
        self._items.insert(index, src)
        if base:
            self._base = src
        return self

    def remove(self, index: int) -> Source:
        """Удаляет документ из позиции ``index`` и возвращает его."""
        self._pin_base()
        src = self._items.pop(index)
        if src is self._base:
            self._base = 0
        return src

    def move(self, old_index: int, new_index: int) -> Stitcher:
        """Перемещает документ на новую позицию."""
        self._pin_base()
        self._items.insert(new_index, self._items.pop(old_index))
        return self

    def set_base(self, index: int) -> Stitcher:
        """Делает главным документ с индексом ``index``."""
        self._base = self._items[index]
        return self

    def _pin_base(self) -> None:
        if isinstance(self._base, int) and self._items:
            try:
                self._base = self._items[self._base]
            except IndexError:
                pass

    @property
    def documents(self) -> list[Source]:
        """Копия текущего списка документов."""
        return list(self._items)

    @property
    def base_index(self) -> int:
        """Индекс главного документа."""
        if isinstance(self._base, Source):
            return next(i for i, s in enumerate(self._items) if s is self._base)
        return self._base if self._base >= 0 else self._base + len(self._items)

    def __len__(self) -> int:
        return len(self._items)

    def __iter__(self) -> Iterator[Source]:
        return iter(self._items)

    # -- результат -------------------------------------------------------------

    def _execute(self, output: OutputTarget | None) -> bytes | None:
        return _run(
            self._items, output,
            base=self._base,
            break_type=self.break_type,
            style_mode=self.style_mode,
            separator=self.separator,
            separator_style=self.separator_style,
            separator_alignment=self.separator_alignment,
            keep_headers_footers=self.keep_headers_footers,
            exclude_from_toc=self.exclude_from_toc,
        )

    def save(self, output: OutputTarget) -> None:
        """Записывает результат в файл или бинарный поток."""
        if output is None:
            raise StitchError("Не указан путь или поток для сохранения.")
        self._execute(output)

    def to_bytes(self) -> bytes:
        """Возвращает результат как ``bytes``."""
        result = self._execute(None)
        assert result is not None
        return result


def attach_title_page(
    title: str | os.PathLike[str],
    document: str | os.PathLike[str],
    output: str | os.PathLike[str],
    section_break: str = "nextPage",
) -> None:
    """
    Добавляет ``title`` перед ``document`` и записывает результат в ``output``.

    Совместимая обёртка над :func:`stitch` с прежним API md2gost
    (``core.merger.attach_title_page``). Основой остаётся ``document``;
    титульный лист импортируется изолированным блоком со стилями
    ``_title_*`` и не попадает в оглавление.
    """
    try:
        title_path = Path(title)
        doc_path = Path(document)
        out = Path(output)
    except TypeError as exc:
        raise StitchError(f"Некорректный путь: {exc}") from exc

    X._validate_input_file(title_path)
    X._validate_input_file(doc_path)
    X._validate_output_path(out)
    validate_break_type(section_break, "section_break")

    if out.resolve() in (title_path.resolve(), doc_path.resolve()):
        raise StitchError(
            f"Выходной файл '{out}' не должен совпадать ни с одним из входных документов."
        )
    if title_path.resolve() == doc_path.resolve():
        raise StitchError("Титульный лист и документ должны быть разными файлами.")

    stitch(
        [
            Source(title_path, style_prefix=TITLE_STYLE_PREFIX, exclude_from_toc=True),
            doc_path,
        ],
        out,
        base=1,
        break_type=section_break,
        style_mode="isolate",
    )
