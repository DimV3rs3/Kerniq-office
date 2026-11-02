"""Create, read, and edit OOXML and ODF documents. Stdlib only.

Word, Excel, and PowerPoint open the OOXML files. LibreOffice opens those and
the ODF files. Tables and paragraphs are numbered from 1. Paragraph numbers
do not count text that lives inside a table.
"""

from __future__ import annotations

import ast
import html
import json
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
PR = "http://schemas.openxmlformats.org/package/2006/relationships"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
XML = "http://www.w3.org/XML/1998/namespace"

OFFICE = "urn:oasis:names:tc:opendocument:xmlns:office:1.0"
TEXT = "urn:oasis:names:tc:opendocument:xmlns:text:1.0"
TABLE = "urn:oasis:names:tc:opendocument:xmlns:table:1.0"
STYLE = "urn:oasis:names:tc:opendocument:xmlns:style:1.0"
FO = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
DRAW = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
MANIFEST = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"

EDITABLE = {".docx", ".xlsx", ".pptx", ".odt", ".ods", ".odp"}
LEGACY = {".doc": ".docx or .odt", ".xls": ".xlsx or .ods", ".ppt": ".pptx or .odp"}
MIME = {
    ".odt": "application/vnd.oasis.opendocument.text",
    ".ods": "application/vnd.oasis.opendocument.spreadsheet",
    ".odp": "application/vnd.oasis.opendocument.presentation",
}

for prefix, uri in (
    ("w", W), ("r", R), ("a", A), ("p", P), ("x", S),
    ("office", OFFICE), ("text", TEXT), ("table", TABLE),
    ("style", STYLE), ("fo", FO), ("draw", DRAW), ("manifest", MANIFEST),
):
    ET.register_namespace(prefix, uri)


class OfficeError(ValueError):
    pass


def q(ns: str, tag: str) -> str:
    return f"{{{ns}}}{tag}"


def _insert_before(parent: ET.Element, child: ET.Element, fresh: ET.Element) -> None:
    parent.insert(list(parent).index(child), fresh)


def canon_style(name: object) -> str | None:
    if name is None or str(name).strip() == "":
        return None
    key = str(name).strip().lower().replace(" ", "").replace("_", "").replace("-", "")
    return {
        "normal": "Normal",
        "standard": "Normal",
        "heading1": "Heading1",
        "heading2": "Heading2",
        "title": "Title",
    }.get(key, str(name).strip())


def _odf_style_name(style: str | None) -> str:
    mapped = canon_style(style) or "Normal"
    return {"Normal": "Standard", "Heading1": "Heading 1", "Heading2": "Heading 2", "Title": "Title"}.get(mapped, mapped)


def _xml(root: ET.Element) -> bytes:
    if root.tag in {q(CT, "Types"), q(PR, "Relationships")}:
        ns = CT if root.tag == q(CT, "Types") else PR
        return _default_ns(root, ns)
    return b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + ET.tostring(root, encoding="utf-8")


def _default_ns(root: ET.Element, ns: str) -> bytes:
    """Content types and relationships use a default xmlns. Word rejects a prefix there."""

    def one(node: ET.Element, *, declare: str | None) -> str:
        name = node.tag.split("}", 1)[-1]
        attrs: list[str] = []
        if declare:
            attrs.append(f'xmlns="{declare}"')
        for key, value in node.attrib.items():
            local = key.split("}", 1)[-1]
            if key.startswith("{") and key[1:].startswith(XML):
                local = "xml:" + local
            attrs.append(f'{local}="{html.escape(str(value), quote=True)}"')
        attr = (" " + " ".join(attrs)) if attrs else ""
        children = list(node)
        text = html.escape(node.text or "")
        if not children and not text:
            return f"<{name}{attr}/>"
        inner = text + "".join(one(child, declare=None) for child in children)
        return f"<{name}{attr}>{inner}</{name}>"

    body = one(root, declare=ns)
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + body).encode("utf-8")


def _toggle(node: ET.Element | None) -> bool:
    if node is None:
        return False
    val = str(node.get(q(W, "val")) or node.get(q(S, "val")) or node.get("val") or "true").lower()
    return val not in {"0", "false", "off", "none"}


def _parse(data: bytes) -> ET.Element:
    return ET.fromstring(data)


def _open_zip(path: Path) -> dict[str, bytes]:
    with zipfile.ZipFile(path) as archive:
        return {info.filename: archive.read(info.filename) for info in archive.infolist()}


def _save_zip(path: Path, parts: dict[str, bytes], *, mime: str | None = None) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + ".writing")
    with zipfile.ZipFile(tmp, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        if mime is not None:
            info = zipfile.ZipInfo("mimetype")
            info.compress_type = zipfile.ZIP_STORED
            archive.writestr(info, mime.encode("ascii"))
        for name in sorted(parts):
            if name == "mimetype":
                continue
            archive.writestr(name, parts[name])
    tmp.replace(path)


def _kind(path: Path) -> str:
    suffix = path.suffix.lower()
    if suffix in LEGACY:
        raise OfficeError(
            f"{path.name} is a legacy binary Office file. "
            f"Use {LEGACY[suffix]} with office_create, office_read, and office_edit."
        )
    if suffix not in EDITABLE:
        raise OfficeError(
            f"{path.name} is not an Office document. "
            "Use .docx, .xlsx, .pptx, .odt, .ods, or .odp."
        )
    return suffix


def _flag(spec: dict, key: str, default: bool) -> bool:
    if not isinstance(spec, dict) or key not in spec or spec.get(key) is None:
        return default
    value = spec.get(key)
    if isinstance(value, str):
        token = value.strip().lower()
        if token in {"", "false", "no", "0", "off"}:
            return False
        if token in {"true", "yes", "1", "on"}:
            return True
    return bool(value)


def _font_of(spec: dict, *, style: str | None) -> tuple[str, float, bool, bool]:
    named = canon_style(style) or "Normal"
    size, bold = {"Normal": (11.0, False), "Heading1": (16.0, True), "Heading2": (13.0, True), "Title": (28.0, True)}.get(named, (11.0, False))
    if isinstance(spec, dict) and spec.get("size_pt") not in (None, ""):
        token = spec["size_pt"]
        if isinstance(token, str):
            token = token.lower().replace("pt", "").strip()
        try:
            size = float(token)
        except (TypeError, ValueError):
            pass
    font = str(spec.get("font") or "Calibri") if isinstance(spec, dict) else "Calibri"
    return font, size, _flag(spec, "bold", bold), _flag(spec, "italic", False)


def _text_of(node: ET.Element, *tags: str) -> str:
    chunks: list[str] = []
    for child in node.iter():
        if child.tag in tags and child.text:
            chunks.append(child.text)
    return "".join(chunks)


def create_document(path: str | Path, spec: dict | None = None) -> dict:
    target = Path(path)
    suffix = _kind(target)
    if target.exists():
        raise OfficeError(f"Already exists: {target.name}. Use office_edit to change it.")
    body = dict(spec or {})
    if suffix == ".docx":
        parts = _docx_parts(body)
        mime = None
    elif suffix == ".xlsx":
        parts = _xlsx_parts(body)
        mime = None
    elif suffix == ".pptx":
        parts = _pptx_parts(body)
        mime = None
    else:
        parts = _odf_parts(suffix, body)
        mime = MIME[suffix]
    _save_zip(target, parts, mime=mime)
    return read_document(target)


def read_document(path: str | Path, *, limit: int = 0) -> dict:
    target = Path(path)
    if not target.is_file():
        raise OfficeError(f"File not found: {target.name}")
    suffix = _kind(target)
    parts = _open_zip(target)
    if suffix == ".docx":
        doc = _read_docx(parts)
    elif suffix == ".xlsx":
        doc = _read_xlsx(parts)
    elif suffix == ".pptx":
        doc = _read_pptx(parts)
    elif suffix == ".odt":
        doc = _read_odt(parts)
    elif suffix == ".ods":
        doc = _read_ods(parts)
    else:
        doc = _read_odp(parts)
    doc["path"] = target.name
    doc["kind"] = suffix[1:]
    # A limit used to cut the document. The model needs the whole file, so it is ignored.
    return _clip(doc, 0)


def _words(text: object) -> int:
    return len(str(text or "").split())


def _remember(bucket: list, value: object) -> None:
    if value in (None, ""):
        return
    if value not in bucket:
        bucket.append(value)


def inspect_document(path: str | Path) -> dict:
    """Outline only. Cell text stays out so the model can see the shape before it reads."""
    doc = read_document(path, limit=0)
    words = 0
    fonts: list = []
    for paragraph in doc.get("paragraphs") or []:
        words += _words(paragraph.get("text"))
        _remember(fonts, paragraph.get("font"))
    table_list = []
    for table in doc.get("tables") or []:
        cells = table.get("cells") or []
        for row in cells:
            values = row if isinstance(row, (list, tuple)) else [row]
            for value in values:
                if isinstance(value, dict):
                    words += _words(value.get("text") or value.get("value"))
                else:
                    words += _words(value)
        for font in table.get("fonts") or []:
            _remember(fonts, font)
        table_list.append({
            "n": table.get("n"),
            "rows": int(table.get("rows") or len(cells)),
            "cols": int(table.get("cols") or 0),
            "fonts": list(table.get("fonts") or []),
            "sizes_pt": list(table.get("sizes_pt") or []),
        })
    sheets = []
    for sheet in doc.get("sheets") or []:
        for cell in sheet.get("cells") or []:
            if isinstance(cell, dict):
                words += _words(cell.get("value"))
                _remember(fonts, cell.get("font"))
            else:
                words += _words(cell)
        for font in sheet.get("fonts") or []:
            _remember(fonts, font)
        sheets.append({
            "name": sheet.get("name"),
            "rows": sheet.get("rows"),
            "cols": sheet.get("cols"),
            "fonts": list(sheet.get("fonts") or []),
            "sizes_pt": list(sheet.get("sizes_pt") or []),
        })
    slides = []
    for slide in doc.get("slides") or []:
        words += _words(slide.get("title"))
        words += _words(slide.get("body"))
        _remember(fonts, slide.get("font"))
        slides.append({"n": slide.get("n"), "title": str(slide.get("title") or "")})
    return {
        "path": doc.get("path"),
        "kind": doc.get("kind"),
        "inspected": True,
        "words": words,
        "paragraphs": len(doc.get("paragraphs") or []),
        "styles": list(doc.get("styles") or []),
        "fonts": fonts,
        "tables": len(table_list),
        "table_list": table_list,
        "sheets": sheets,
        "slides": slides,
        "truncated": False,
        "omitted": [],
    }


def edit_document(path: str | Path, ops: list) -> dict:
    target = Path(path)
    if not target.is_file():
        raise OfficeError(f"File not found: {target.name}")
    suffix = _kind(target)
    if isinstance(ops, dict):
        ops = [ops]
    if not isinstance(ops, list) or not ops:
        raise OfficeError("ops must be a non-empty list of edits.")
    clean = []
    for op in ops:
        if not isinstance(op, dict):
            raise OfficeError(
                "Each edit must be an object with an op name. "
                "Example: {\"op\": \"set_sheet_cell\", \"sheet\": 1, \"row\": 1, \"col\": 1, \"value\": \"text\"}."
            )
        clean.append(dict(op))
    ops = clean
    parts = _open_zip(target)
    mime = MIME.get(suffix)
    if suffix in {".docx", ".odt"}:
        _edit_text(suffix, parts, ops)
    elif suffix in {".xlsx", ".ods"}:
        _edit_sheet(suffix, parts, ops)
    else:
        _edit_show(suffix, parts, ops)
    _save_zip(target, parts, mime=mime)
    return read_document(target)


def _clip(doc: dict, limit: int) -> dict:
    doc["truncated"] = False
    doc["omitted"] = []
    if not limit or int(limit) <= 0:
        return doc
    limit = int(limit)
    notes = []
    paragraphs = doc.get("paragraphs") or []
    if len(paragraphs) > limit:
        doc["paragraphs"] = paragraphs[:limit]
        notes.append(f"paragraphs {limit + 1}–{len(paragraphs)}")
    for table in doc.get("tables") or []:
        cells = table.get("cells") or []
        if len(cells) > limit:
            table["rows_total"] = len(cells)
            table["cells"] = cells[:limit]
            notes.append(f"table {table.get('n')} rows {limit + 1}–{len(cells)}")
    for sheet in doc.get("sheets") or []:
        cells = sheet.get("cells") or []
        if len(cells) > limit:
            sheet["cells_total"] = len(cells)
            sheet["cells"] = cells[:limit]
            notes.append(f"sheet {sheet.get('name')} cells {limit + 1}–{len(cells)}")
    doc["truncated"] = bool(notes)
    doc["omitted"] = notes
    return doc


def _rels(target: str, pairs: list[tuple[str, str]]) -> bytes:
    root = ET.Element(q(PR, "Relationships"))
    for rel_id, (rel_type, dest) in zip(
        (f"rId{i}" for i in range(1, len(pairs) + 1)),
        pairs,
        strict=False,
    ):
        ET.SubElement(root, q(PR, "Relationship"), {"Id": rel_id, "Type": rel_type, "Target": dest})
    return _xml(root)


def _content_types(defaults: list[tuple[str, str]], overrides: list[tuple[str, str]]) -> bytes:
    root = ET.Element(q(CT, "Types"))
    for ext, value in defaults:
        ET.SubElement(root, q(CT, "Default"), {"Extension": ext, "ContentType": value})
    for part, value in overrides:
        ET.SubElement(root, q(CT, "Override"), {"PartName": part, "ContentType": value})
    return _xml(root)


def _docx_paragraph(spec: dict) -> ET.Element:
    style, = (canon_style(spec.get("style")) or "Normal",)
    font, size, bold, italic = _font_of(spec, style=style)
    paragraph = ET.Element(q(W, "p"))
    props = ET.SubElement(paragraph, q(W, "pPr"))
    ET.SubElement(props, q(W, "pStyle"), {q(W, "val"): style})
    run = ET.SubElement(paragraph, q(W, "r"))
    run_props = ET.SubElement(run, q(W, "rPr"))
    ET.SubElement(run_props, q(W, "rFonts"), {q(W, "ascii"): font, q(W, "hAnsi"): font})
    if bold:
        ET.SubElement(run_props, q(W, "b"))
    if italic:
        ET.SubElement(run_props, q(W, "i"))
    ET.SubElement(run_props, q(W, "sz"), {q(W, "val"): str(int(round(size * 2)))})
    value = spec.get("text")
    if value is None:
        value = ""
    text = ET.SubElement(run, q(W, "t"))
    value = _as_text(value)
    if value[:1].isspace() or value[-1:].isspace():
        text.set(q(XML, "space"), "preserve")
    text.text = value
    return paragraph


_PAGE_WIDTH = 12240
_PAGE_MARGIN = 1440
_TEXT_WIDTH = _PAGE_WIDTH - (_PAGE_MARGIN * 2)


def _docx_table(spec: dict) -> ET.Element:
    rows = _table_grid(spec)
    width = max(len(row) if isinstance(row, list) else 1 for row in rows)
    width = max(width, 1)
    table = ET.Element(q(W, "tbl"))
    props = ET.SubElement(table, q(W, "tblPr"))
    ET.SubElement(props, q(W, "tblW"), {q(W, "w"): str(_TEXT_WIDTH), q(W, "type"): "dxa"})
    ET.SubElement(props, q(W, "tblLayout"), {q(W, "type"): "fixed"})
    ET.SubElement(props, q(W, "tblInd"), {q(W, "w"): "0", q(W, "type"): "dxa"})
    borders = ET.SubElement(props, q(W, "tblBorders"))
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        ET.SubElement(borders, q(W, edge), {q(W, "val"): "single", q(W, "sz"): "4", q(W, "space"): "0", q(W, "color"): "666666"})
    margins = ET.SubElement(props, q(W, "tblCellMar"))
    for edge, pad in (("top", "40"), ("left", "80"), ("bottom", "40"), ("right", "80")):
        ET.SubElement(margins, q(W, edge), {q(W, "w"): pad, q(W, "type"): "dxa"})
    base, extra = divmod(_TEXT_WIDTH, width)
    col_widths = [base + (1 if index < extra else 0) for index in range(width)]
    grid = ET.SubElement(table, q(W, "tblGrid"))
    for col_width in col_widths:
        ET.SubElement(grid, q(W, "gridCol"), {q(W, "w"): str(col_width)})
    for row in rows:
        values = row if isinstance(row, list) else [row]
        tr = ET.SubElement(table, q(W, "tr"))
        for col in range(width):
            raw = values[col] if col < len(values) else ""
            cell_spec = raw if isinstance(raw, dict) else {"text": "" if raw is None else str(raw)}
            tc = ET.SubElement(tr, q(W, "tc"))
            props = ET.SubElement(tc, q(W, "tcPr"))
            ET.SubElement(props, q(W, "tcW"), {q(W, "w"): str(col_widths[col]), q(W, "type"): "dxa"})
            tc.append(_docx_paragraph(cell_spec))
    return table


def _docx_styles() -> bytes:
    root = ET.Element(q(W, "styles"))
    for style, size, bold in (("Normal", 22, False), ("Heading1", 32, True), ("Heading2", 26, True), ("Title", 56, True)):
        node = ET.SubElement(root, q(W, "style"), {q(W, "type"): "paragraph", q(W, "styleId"): style})
        ET.SubElement(node, q(W, "name"), {q(W, "val"): style})
        run = ET.SubElement(node, q(W, "rPr"))
        ET.SubElement(run, q(W, "rFonts"), {q(W, "ascii"): "Calibri", q(W, "hAnsi"): "Calibri"})
        if bold:
            ET.SubElement(run, q(W, "b"))
        ET.SubElement(run, q(W, "sz"), {q(W, "val"): str(size)})
    return _xml(root)


def _core_props() -> bytes:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:creator>Kerniq</dc:creator></cp:coreProperties>'
    ).encode("utf-8")


def _package_rels(main: str) -> bytes:
    return _rels("", [
        ("http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", main),
        ("http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties", "docProps/core.xml"),
    ])


def _docx_parts(spec: dict) -> dict[str, bytes]:
    document = ET.Element(q(W, "document"))
    body = ET.SubElement(document, q(W, "body"))
    paragraphs = _paragraphs_of(spec.get("paragraphs"))
    if spec.get("title"):
        paragraphs.insert(0, {"text": spec["title"], "style": "Title"})
    if not paragraphs:
        paragraphs = [{"text": ""}]
    for item in paragraphs:
        body.append(_docx_paragraph(item))
    for table in _as_list(spec.get("tables")):
        body.append(_docx_table(table if isinstance(table, dict) else {"cells": table}))
    sect = ET.SubElement(body, q(W, "sectPr"))
    ET.SubElement(sect, q(W, "pgSz"), {q(W, "w"): str(_PAGE_WIDTH), q(W, "h"): "15840"})
    ET.SubElement(sect, q(W, "pgMar"), {
        q(W, "top"): str(_PAGE_MARGIN),
        q(W, "right"): str(_PAGE_MARGIN),
        q(W, "bottom"): str(_PAGE_MARGIN),
        q(W, "left"): str(_PAGE_MARGIN),
        q(W, "header"): "720",
        q(W, "footer"): "720",
        q(W, "gutter"): "0",
    })
    return {
        "[Content_Types].xml": _content_types(
            [("rels", "application/vnd.openxmlformats-package.relationships+xml"), ("xml", "application/xml")],
            [
                ("/word/document.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"),
                ("/word/styles.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"),
                ("/docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml"),
            ],
        ),
        "_rels/.rels": _package_rels("word/document.xml"),
        "docProps/core.xml": _core_props(),
        "word/_rels/document.xml.rels": _rels("styles", [("http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles", "styles.xml")]),
        "word/document.xml": _xml(document),
        "word/styles.xml": _docx_styles(),
    }


def _run_info(run: ET.Element) -> dict:
    props = run.find(q(W, "rPr"))
    info = {"text": _text_of(run, q(W, "t")), "font": None, "size_pt": None, "bold": False, "italic": False}
    if props is None:
        return info
    fonts = props.find(q(W, "rFonts"))
    if fonts is not None:
        info["font"] = fonts.get(q(W, "ascii")) or fonts.get(q(W, "hAnsi"))
    size = props.find(q(W, "sz"))
    if size is not None and size.get(q(W, "val")):
        info["size_pt"] = float(size.get(q(W, "val"))) / 2
    info["bold"] = _toggle(props.find(q(W, "b")))
    info["italic"] = _toggle(props.find(q(W, "i")))
    return info


def _docx_style_defaults(parts: dict[str, bytes]) -> dict[str, dict]:
    if "word/styles.xml" not in parts:
        return {}
    found: dict[str, dict] = {}
    for style in _parse(parts["word/styles.xml"]).findall(q(W, "style")):
        style_id = canon_style(style.get(q(W, "styleId"))) or style.get(q(W, "styleId"))
        props = style.find(q(W, "rPr"))
        if not style_id or props is None:
            continue
        info = {"font": None, "size_pt": None, "bold": False, "italic": False}
        fonts = props.find(q(W, "rFonts"))
        if fonts is not None:
            info["font"] = fonts.get(q(W, "ascii")) or fonts.get(q(W, "hAnsi"))
        size = props.find(q(W, "sz"))
        if size is not None and size.get(q(W, "val")):
            info["size_pt"] = float(size.get(q(W, "val"))) / 2
        info["bold"] = _toggle(props.find(q(W, "b")))
        info["italic"] = _toggle(props.find(q(W, "i")))
        found[style_id] = info
    return found


def _paragraph_info(paragraph: ET.Element, number: int, styles: dict | None = None) -> dict:
    style = None
    props = paragraph.find(q(W, "pPr"))
    if props is not None:
        named = props.find(q(W, "pStyle"))
        if named is not None:
            style = canon_style(named.get(q(W, "val")))
    runs = [_run_info(run) for run in paragraph.findall(q(W, "r"))]
    first = next((run for run in runs if run.get("text")), runs[0] if runs else {})
    inherited = (styles or {}).get(style or "Normal") or {}
    explicit = bool(first.get("font") or first.get("size_pt") is not None or first.get("bold") or first.get("italic"))
    return {
        "n": number,
        "style": style or "Normal",
        "text": "".join(run["text"] for run in runs),
        "font": first.get("font") or inherited.get("font"),
        "size_pt": first.get("size_pt") if first.get("size_pt") is not None else inherited.get("size_pt"),
        "bold": bool(first.get("bold")) if explicit else bool(inherited.get("bold")),
        "italic": bool(first.get("italic")) if explicit else bool(inherited.get("italic")),
    }


def _read_docx(parts: dict[str, bytes]) -> dict:
    root = _parse(parts["word/document.xml"])
    body = root.find(q(W, "body"))
    paragraphs = []
    tables = []
    styles = []
    defined = _docx_style_defaults(parts)
    if body is None:
        return {"paragraphs": [], "tables": [], "styles": []}
    for child in list(body):
        if child.tag == q(W, "p"):
            info = _paragraph_info(child, len(paragraphs) + 1, defined)
            paragraphs.append(info)
            if info["style"] not in styles:
                styles.append(info["style"])
        elif child.tag == q(W, "tbl"):
            rows = []
            fonts, sizes = [], []
            for tr in child.findall(q(W, "tr")):
                row = []
                for tc in tr.findall(q(W, "tc")):
                    text = []
                    for paragraph in tc.findall(q(W, "p")):
                        info = _paragraph_info(paragraph, 0, defined)
                        text.append(info["text"])
                        if info["font"] and info["font"] not in fonts:
                            fonts.append(info["font"])
                        if info["size_pt"] and info["size_pt"] not in sizes:
                            sizes.append(info["size_pt"])
                    row.append("\n".join(part for part in text if part))
                rows.append(row)
            width = max((len(row) for row in rows), default=0)
            tables.append({"n": len(tables) + 1, "rows": len(rows), "cols": width, "cells": rows, "fonts": fonts, "sizes_pt": sizes})
    return {"paragraphs": paragraphs, "tables": tables, "styles": styles}


def _docx_body(parts: dict[str, bytes]) -> tuple[ET.Element, ET.Element]:
    root = _parse(parts["word/document.xml"])
    body = root.find(q(W, "body"))
    if body is None:
        raise OfficeError("This document has no body.")
    return root, body


def _edit_docx(parts: dict[str, bytes], ops: list) -> None:
    root, body = _docx_body(parts)
    for op in ops:
        name = str(op.get("op") or "")
        if name in {"set_paragraph", "set_style"}:
            items = [child for child in list(body) if child.tag == q(W, "p")]
            index = int(op.get("index") or 0)
            if index < 1 or index > len(items):
                raise OfficeError(f"Paragraph {index} is outside 1..{len(items)}.")
            current = _paragraph_info(items[index - 1], index)
            spec = dict(op)
            if name == "set_style":
                spec["text"] = current["text"]
            spec.setdefault("text", current["text"])
            spec.setdefault("style", current["style"])
            spec.setdefault("font", current["font"])
            spec.setdefault("size_pt", current["size_pt"])
            spec.setdefault("bold", current["bold"])
            spec.setdefault("italic", current["italic"])
            items[index - 1].clear()
            fresh = _docx_paragraph(spec)
            items[index - 1].extend(list(fresh))
            items[index - 1].attrib.update(fresh.attrib)
        elif name == "add_paragraph":
            fresh = _docx_paragraph(op)
            sect = body.find(q(W, "sectPr"))
            if sect is not None:
                _insert_before(body, sect, fresh)
            else:
                body.append(fresh)
        elif name == "set_cell":
            _docx_set_cell(body, op)
        elif name == "add_table":
            sect = body.find(q(W, "sectPr"))
            table = _docx_table(op)
            if sect is not None:
                _insert_before(body, sect, table)
            else:
                body.append(table)
        else:
            raise OfficeError(f"{name or 'missing op'} is not a document edit. Use set_paragraph, add_paragraph, set_cell, add_table, or set_style.")
    parts["word/document.xml"] = _xml(root)


def _docx_set_cell(body: ET.Element, op: dict) -> None:
    tables = [child for child in list(body) if child.tag == q(W, "tbl")]
    table_n = int(op.get("table") or 0)
    row_n = int(op.get("row") or 0)
    col_n = int(op.get("col") or 0)
    if table_n < 1 or table_n > len(tables):
        raise OfficeError(f"Table {table_n} is outside 1..{len(tables)}.")
    rows = tables[table_n - 1].findall(q(W, "tr"))
    if row_n < 1 or row_n > len(rows):
        raise OfficeError(f"Row {row_n} is outside 1..{len(rows)} in table {table_n}.")
    cells = rows[row_n - 1].findall(q(W, "tc"))
    if col_n < 1 or col_n > len(cells):
        raise OfficeError(f"Column {col_n} is outside 1..{len(cells)} in table {table_n}.")
    cell = cells[col_n - 1]
    for child in list(cell):
        cell.remove(child)
    cell.append(_docx_paragraph(op))


def _col_letter(number: int) -> str:
    letters = ""
    while number:
        number, rest = divmod(number - 1, 26)
        letters = chr(65 + rest) + letters
    return letters


def _col_number(letters: str) -> int:
    value = 0
    for char in letters:
        if char.isalpha():
            value = value * 26 + ord(char.upper()) - 64
    return value


def _split_ref(ref: str) -> tuple[int, int]:
    letters = "".join(ch for ch in ref if ch.isalpha())
    digits = "".join(ch for ch in ref if ch.isdigit())
    return int(digits or "0"), _col_number(letters)


def _xlsx_styles(fonts: list[tuple[str, float, bool, bool]]) -> bytes:
    root = ET.Element(q(S, "styleSheet"))
    font_node = ET.SubElement(root, q(S, "fonts"), {"count": str(len(fonts))})
    for name, size, bold, italic in fonts:
        font = ET.SubElement(font_node, q(S, "font"))
        if bold:
            ET.SubElement(font, q(S, "b"))
        if italic:
            ET.SubElement(font, q(S, "i"))
        ET.SubElement(font, q(S, "sz"), {"val": str(size)})
        ET.SubElement(font, q(S, "name"), {"val": name})
    fills = ET.SubElement(root, q(S, "fills"), {"count": "2"})
    ET.SubElement(ET.SubElement(fills, q(S, "fill")), q(S, "patternFill"), {"patternType": "none"})
    ET.SubElement(ET.SubElement(fills, q(S, "fill")), q(S, "patternFill"), {"patternType": "gray125"})
    borders = ET.SubElement(root, q(S, "borders"), {"count": "1"})
    border = ET.SubElement(borders, q(S, "border"))
    for edge in ("left", "right", "top", "bottom", "diagonal"):
        ET.SubElement(border, q(S, edge))
    ET.SubElement(root, q(S, "cellStyleXfs"), {"count": "1"}).append(
        ET.Element(q(S, "xf"), {"fontId": "0", "fillId": "0", "borderId": "0"})
    )
    xfs = ET.SubElement(root, q(S, "cellXfs"), {"count": str(len(fonts))})
    for index in range(len(fonts)):
        ET.SubElement(xfs, q(S, "xf"), {"fontId": str(index), "fillId": "0", "borderId": "0", "applyFont": "1"})
    return _xml(root)


def _as_text(value) -> str:
    if value is None:
        return ""
    if isinstance(value, dict):
        return _as_text(value.get("text", value.get("value")))
    if isinstance(value, (list, tuple)):
        return "\n".join(_as_text(item) for item in value)
    return str(value)


def _int_or_none(value):
    if isinstance(value, bool) or value is None or isinstance(value, (list, tuple, dict)):
        return None
    if isinstance(value, int):
        return value
    if isinstance(value, float) and value.is_integer():
        return int(value)
    if isinstance(value, str) and value.strip().isdigit():
        return int(value.strip())
    return None


def _cell_value(value):
    if isinstance(value, dict):
        return _cell_value(value.get("value", value.get("text")))
    if isinstance(value, (list, tuple)):
        return _as_text(value)
    return value



_BROKEN_CELL = re.compile(r"""^[\{\[\s]*['"]?\d+\s*,\s*\d+['"]?\s*$""")
_RC_REF = re.compile(r"^[rR](\d+)\s*[cC](\d+)$")


def _broken_cell(value) -> bool:
    return isinstance(value, str) and bool(_BROKEN_CELL.match(value.strip()))


def _loads_cell_blob(blob: str):
    body = str(blob or "").strip()
    if body[:6].lower() == "cells=":
        body = body[6:].strip()
    if not body or body[:1] not in "{[":
        return None
    for loader in (json.loads, ast.literal_eval):
        try:
            value = loader(body)
        except (ValueError, SyntaxError, MemoryError):
            continue
        if isinstance(value, (dict, list)):
            return value
    return None


def _take_style(spec: dict, source) -> None:
    if not isinstance(source, dict):
        return
    for key in ("bold", "italic", "size_pt"):
        if key in source and source.get(key) not in (None, ""):
            spec[key] = source[key]
    font = source.get("font")
    if isinstance(font, bool) or font in (None, ""):
        return
    if isinstance(font, (int, float)) or (isinstance(font, str) and font.strip().replace(".", "", 1).isdigit()):
        spec.setdefault("size_pt", font)
        return
    if isinstance(font, str) and font.strip():
        spec["font"] = font


def _ref_pair(ref) -> tuple[int, int] | None:
    token = str(ref or "").strip().strip("{}[]()").strip().strip("'\"").strip()
    if not token:
        return None
    match = _RC_REF.fullmatch(token)
    if match:
        return int(match.group(1)), int(match.group(2))
    if token[:1].isalpha():
        row, col = _split_ref(token)
        if row >= 1 and col >= 1:
            return row, col
        return None
    bits = [bit.strip().strip("'\"") for bit in token.replace(";", ",").split(",")]
    bits = [bit for bit in bits if bit]
    if len(bits) == 2 and _int_or_none(bits[0]) and _int_or_none(bits[1]):
        return _int_or_none(bits[0]), _int_or_none(bits[1])
    return None


def _is_address_row(item) -> bool:
    if not isinstance(item, (list, tuple)) or len(item) < 3:
        return False
    if _int_or_none(item[0]) is None or _int_or_none(item[1]) is None:
        return False
    return not isinstance(item[2], (list, tuple))


def _cell_from_raw(raw, row: int = 1, col: int = 1) -> dict:
    if isinstance(raw, dict):
        spec = {"row": _int_or_none(raw.get("row")) or row, "col": _int_or_none(raw.get("col")) or col}
        if "value" in raw:
            inner = raw.get("value")
        elif "text" in raw:
            inner = raw.get("text")
        else:
            inner = ""
        if isinstance(inner, dict):
            _take_style(spec, inner)
            inner = inner.get("value", inner.get("text", ""))
        spec["value"] = _cell_value(inner)
        _take_style(spec, raw)
        return spec
    if _is_address_row(raw):
        inner = raw[2]
        spec = {"row": _int_or_none(raw[0]) or row, "col": _int_or_none(raw[1]) or col}
        if isinstance(inner, dict):
            _take_style(spec, inner)
            spec["value"] = _cell_value(inner.get("value", inner.get("text", "")))
        else:
            spec["value"] = _cell_value(inner)
        extra = raw[3] if len(raw) > 3 else None
        if isinstance(extra, dict):
            _take_style(spec, extra)
        elif isinstance(extra, str) and extra.strip() and not extra.strip().replace(".", "", 1).isdigit():
            spec["font"] = extra
        return spec
    if isinstance(raw, (list, tuple)):
        return {"row": row, "col": col, "value": _as_text(raw)}
    return {"row": row, "col": col, "value": _cell_value(raw)}


def _looks_like_cell(value) -> bool:
    if not isinstance(value, dict):
        return not isinstance(value, (list, tuple))
    return any(key in value for key in ("text", "value", "bold", "italic", "font", "size_pt", "row", "col"))


def _flat_cells(raw) -> list[dict] | None:
    if len(raw) >= 3 and len(raw) % 3 == 0:
        triples = []
        for index in range(0, len(raw), 3):
            row = _int_or_none(raw[index])
            col = _int_or_none(raw[index + 1])
            if not row or not col:
                triples = []
                break
            value = raw[index + 2]
            payload = value if isinstance(value, dict) else {"value": value}
            triples.append(_cell_from_raw(payload, row, col))
        if triples:
            return triples
    if len(raw) >= 2 and len(raw) % 2 == 0:
        pairs = []
        for index in range(0, len(raw), 2):
            col = _int_or_none(raw[index])
            value = raw[index + 1]
            if not col or (isinstance(value, (int, float)) and not isinstance(value, bool)):
                return None
            payload = value if isinstance(value, dict) else {"value": value}
            pairs.append(_cell_from_raw(payload, 1, col))
        if pairs:
            return pairs
    return None


def _has_payload(raw) -> bool:
    return not (raw is None or raw == "" or raw == [] or raw == {})


def _cells_of(raw) -> list[dict]:
    if isinstance(raw, str):
        parsed = _loads_cell_blob(raw)
        if parsed is not None:
            return _cells_of(parsed)
    cells = _cells_from(raw)
    if cells and all(_broken_cell(cell.get("value")) for cell in cells):
        raise OfficeError(
            "Those sheet cells were not a grid. Use [row, col, text], starting at 1. "
            'Example: {"name": "Sheet1", "cells": [[1, 1, "File"], [1, 2, "Indicator"]]}.'
        )
    if _has_payload(raw) and not cells and not isinstance(raw, str):
        raise OfficeError(
            "Sheet cells were not understood. Use [row, col, text], starting at 1. "
            'Example: {"name": "Sheet1", "cells": [[1, 1, "File"], [1, 2, "Indicator"]]}.'
        )
    return cells


def _cells_from(raw) -> list[dict]:
    if raw is None or raw == "":
        return []
    if isinstance(raw, str):
        return [_cell_from_raw(raw)]
    if isinstance(raw, dict):
        address_keys = []
        for key in raw:
            if key in {"name", "title"}:
                continue
            if _ref_pair(key) or (_int_or_none(str(key).strip().strip("'\"")) and key not in {"row", "col", "font", "size_pt"}):
                address_keys.append(key)
        if not address_keys and any(key in raw for key in ("value", "text", "row", "col")):
            return [_cell_from_raw(raw)]
        cells = []
        for key, value in raw.items():
            if key in {"name", "title"}:
                continue
            pair = _ref_pair(key)
            if pair:
                row, col = pair
                payload = value if isinstance(value, dict) else {"value": value}
                cells.append(_cell_from_raw(payload, row, col))
                continue
            number = _int_or_none(str(key).strip().strip("'\""))
            if not number or number < 1 or key in {"row", "col", "font", "size_pt"}:
                continue
            if isinstance(value, dict) and not _looks_like_cell(value):
                nested = []
                for col_key, col_value in value.items():
                    col = _int_or_none(str(col_key).strip().strip("'\""))
                    if not col:
                        continue
                    payload = col_value if isinstance(col_value, dict) else {"value": col_value}
                    nested.append(_cell_from_raw(payload, number, col))
                if nested:
                    cells.extend(nested)
                    continue
            payload = value if isinstance(value, dict) else {"value": value}
            cells.append(_cell_from_raw(payload, 1, number))
        if cells:
            return cells
        if any(key in raw for key in ("value", "text", "row", "col")):
            return [_cell_from_raw(raw)]
        return []
    if isinstance(raw, (list, tuple)):
        if raw and all(isinstance(item, (list, tuple)) for item in raw):
            if any(_is_address_row(item) for item in raw):
                return [_cell_from_raw(item) for item in raw]
            cells = []
            for row_n, row in enumerate(raw, start=1):
                for col_n, value in enumerate(list(row), start=1):
                    payload = value if isinstance(value, dict) else {"value": value}
                    cells.append(_cell_from_raw(payload, row_n, col_n))
            return cells
        flat = _flat_cells(list(raw))
        if flat is not None:
            return flat
        kept = [item for item in raw if isinstance(item, (dict, str))]
        return [_cell_from_raw(item) for item in kept]
    return [_cell_from_raw(raw)]


def _name_only(item) -> str | None:
    if isinstance(item, str) and item.strip():
        return item.strip()
    if not isinstance(item, dict):
        return None
    name = item.get("name") or item.get("title")
    if not name:
        return None
    extra = [key for key in item if key not in {"name", "title"} and item.get(key) not in (None, "", [], {})]
    if extra:
        return None
    return str(name)


def _table_grid(spec) -> list:
    if not isinstance(spec, dict):
        spec = {"cells": spec}
    rows = spec.get("cells")
    if rows is None and isinstance(spec.get("rows"), list):
        rows = spec.get("rows")
    if not rows and spec.get("rows") and spec.get("cols"):
        width = max(1, int(spec.get("cols") or 1))
        return [[""] * width for _ in range(int(spec["rows"]))]
    cells = _cells_of(rows)
    if not cells:
        return [[""]]
    height = max(int(cell["row"]) for cell in cells)
    width = max(int(cell["col"]) for cell in cells)
    grid = [["" for _ in range(width)] for _ in range(height)]
    for cell in cells:
        value = cell.get("value")
        styled = any(cell.get(key) not in (None, "", False) for key in ("font", "size_pt", "bold", "italic"))
        row_n = int(cell["row"]) - 1
        col_n = int(cell["col"]) - 1
        if row_n < 0 or col_n < 0:
            continue
        if styled:
            payload = {"text": "" if value is None else value}
            for key in ("font", "size_pt", "bold", "italic"):
                if cell.get(key) not in (None, ""):
                    payload[key] = cell[key]
            grid[row_n][col_n] = payload
        else:
            grid[row_n][col_n] = "" if value is None else value
    return grid


_BAD_SHEET = str.maketrans({":": " ", "\\": " ", "/": " ", "?": " ", "*": " ", "[": " ", "]": " "})


def _sheet_name(name, fallback: str, used: set[str]) -> str:
    text = str(name or "").translate(_BAD_SHEET).strip() or fallback
    text = text[:31].strip() or fallback
    base = text
    number = 2
    while text in used:
        suffix = f" {number}"
        text = (base[: 31 - len(suffix)] + suffix).strip()
        number += 1
    used.add(text)
    return text


def _sheets_of(raw) -> list[dict]:
    if isinstance(raw, dict):
        if any(key in raw for key in ("cells", "rows", "name", "title")) or _name_only(raw):
            raw = [raw]
        else:
            raw = [{"name": name, "cells": value} for name, value in raw.items()]
    elif isinstance(raw, str):
        raw = [raw]
    elif not isinstance(raw, list):
        raw = []
    pending: list[tuple] = []
    for item in raw:
        label = _name_only(item)
        if label and not (isinstance(item, dict) and _has_payload(item.get("cells", item.get("rows")))):
            if pending and not pending[-1][1]:
                _old, _explicit, cells = pending[-1]
                pending[-1] = (label, True, cells)
            else:
                pending.append((label, True, []))
            continue
        if isinstance(item, dict):
            explicit = item.get("name") or item.get("title")
            pending.append((explicit, bool(explicit), _cells_of(item.get("cells", item.get("rows")))))
            continue
        if isinstance(item, list):
            pending.append((None, False, _cells_of(item)))
            continue
    if not pending:
        return [{"name": "Sheet1", "cells": []}]
    used: set[str] = set()
    sheets = []
    for index, (name, _explicit, cells) in enumerate(pending, start=1):
        sheets.append({"name": _sheet_name(name, f"Sheet{index}", used), "cells": cells})
    return sheets



def _slides_of(raw, title: str = "") -> list[dict]:
    if not raw:
        return [{"title": title or "Slide 1", "body": ""}]
    if isinstance(raw, dict):
        raw = [raw]
    elif isinstance(raw, str):
        raw = [raw]
    elif not isinstance(raw, list):
        raw = [raw]
    slides = []
    for item in raw:
        if isinstance(item, str):
            slides.append({"title": item, "body": ""})
        elif isinstance(item, (list, tuple)):
            slides.append({
                "title": "" if not item else _as_text(item[0]),
                "body": "" if len(item) < 2 else _as_text(item[1]),
            })
        elif isinstance(item, dict):
            slide = dict(item)
            slide["title"] = _as_text(slide.get("title", slide.get("name")) or "")
            slide["body"] = _as_text(slide.get("body", slide.get("text")) or "")
            slides.append(slide)
        else:
            slides.append({"title": _as_text(item), "body": ""})
    return slides or [{"title": title or "Slide 1", "body": ""}]


def _paragraphs_of(raw) -> list[dict]:
    if not raw:
        return []
    if isinstance(raw, dict):
        raw = [raw]
    elif isinstance(raw, str):
        raw = [raw]
    elif not isinstance(raw, list):
        raw = [raw]
    items = []
    for item in raw:
        if isinstance(item, str):
            items.append({"text": item})
        elif isinstance(item, dict):
            spec = dict(item)
            if spec.get("text") is None:
                spec["text"] = spec.get("value") if spec.get("value") is not None else ""
            spec["text"] = _as_text(spec.get("text"))
            items.append(spec)
        else:
            items.append({"text": _as_text(item)})
    return items


def _as_list(raw):
    if raw is None or raw == "":
        return []
    if isinstance(raw, dict):
        return [raw]
    if isinstance(raw, (list, tuple)):
        return list(raw)
    return [raw]


def _cell_spec(raw, row: int, col: int) -> dict:
    return _cell_from_raw(raw, row, col)


def _xlsx_width(text: str) -> float:
    return float(min(80, max(10, len(text) + 2)))


def _xlsx_cols(widths: dict[int, float]) -> ET.Element:
    cols = ET.Element(q(S, "cols"))
    for col in sorted(widths):
        ET.SubElement(cols, q(S, "col"), {
            "min": str(col),
            "max": str(col),
            "width": f"{widths[col]:.2f}",
            "customWidth": "1",
        })
    return cols


def _is_formula(value) -> bool:
    return isinstance(value, str) and value.startswith("=") and len(value) > 1


def _write_xlsx_value(node: ET.Element, value) -> None:
    """Put a value into a cell without discarding the rest of the worksheet."""
    for child in list(node):
        local = child.tag.split("}", 1)[-1]
        if local in {"v", "is", "f"}:
            node.remove(child)
    if isinstance(value, dict):
        value = _cell_value(value)
    if _is_formula(value):
        node.attrib.pop("t", None)
        ET.SubElement(node, q(S, "f")).text = str(value)[1:]
        return
    if isinstance(value, bool):
        node.set("t", "b")
        ET.SubElement(node, q(S, "v")).text = "1" if value else "0"
        return
    if value is None or value == "":
        node.attrib.pop("t", None)
        return
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        node.attrib.pop("t", None)
        ET.SubElement(node, q(S, "v")).text = str(value)
        return
    node.set("t", "inlineStr")
    ET.SubElement(ET.SubElement(node, q(S, "is")), q(S, "t")).text = str(value)


def _style_requested(op: dict) -> bool:
    return any(op.get(key) not in (None, "") for key in ("font", "bold", "italic", "size_pt"))


def _xlsx_ensure_style(parts: dict[str, bytes], key: tuple) -> str | None:
    """Append one font and one cell xf. Never rebuild the stylesheet."""
    if "xl/styles.xml" not in parts:
        return None
    root = _parse(parts["xl/styles.xml"])
    fonts = root.find(q(S, "fonts"))
    xfs = root.find(q(S, "cellXfs"))
    if fonts is None or xfs is None:
        return None
    font_nodes = list(fonts)
    xf_nodes = list(xfs)
    want_name, want_size, want_bold, want_italic = key
    for index, xf in enumerate(xf_nodes):
        font_id = int(xf.get("fontId") or 0)
        if font_id < 0 or font_id >= len(font_nodes):
            continue
        font = font_nodes[font_id]
        name = font.find(q(S, "name"))
        size = font.find(q(S, "sz"))
        current = (
            name.get("val") if name is not None else "Calibri",
            float(size.get("val")) if size is not None and size.get("val") else 11.0,
            font.find(q(S, "b")) is not None,
            font.find(q(S, "i")) is not None,
        )
        if current == (want_name, float(want_size), bool(want_bold), bool(want_italic)):
            return str(index)
    font = ET.SubElement(fonts, q(S, "font"))
    if want_bold:
        ET.SubElement(font, q(S, "b"))
    if want_italic:
        ET.SubElement(font, q(S, "i"))
    ET.SubElement(font, q(S, "sz"), {"val": str(want_size)})
    ET.SubElement(font, q(S, "name"), {"val": str(want_name)})
    fonts.set("count", str(len(list(fonts))))
    xf = ET.SubElement(xfs, q(S, "xf"), {
        "fontId": str(len(list(fonts)) - 1),
        "fillId": "0",
        "borderId": "0",
        "applyFont": "1",
    })
    xfs.set("count", str(len(list(xfs))))
    parts["xl/styles.xml"] = _xml(root)
    return str(len(list(xfs)) - 1)


def _xlsx_sheet(cells: list[dict], style_ids: dict[tuple, int]) -> bytes:
    root = ET.Element(q(S, "worksheet"))
    widths: dict[int, float] = {}
    if cells:
        max_row = max(int(cell["row"]) for cell in cells)
        max_col = max(int(cell["col"]) for cell in cells)
        ET.SubElement(root, q(S, "dimension"), {"ref": f"A1:{_col_letter(max_col)}{max_row}"})
        for cell in cells:
            text = "" if cell.get("value") is None else str(cell["value"])
            col = int(cell["col"])
            widths[col] = max(widths.get(col, 10.0), _xlsx_width(text))
    if widths:
        root.append(_xlsx_cols(widths))
    data = ET.SubElement(root, q(S, "sheetData"))
    rows: dict[int, list[dict]] = {}
    for cell in cells:
        rows.setdefault(int(cell["row"]), []).append(cell)
    for row_n in sorted(rows):
        row = ET.SubElement(data, q(S, "row"), {"r": str(row_n)})
        for cell in sorted(rows[row_n], key=lambda item: int(item["col"])):
            font, size, bold, italic = _font_of(cell, style="Normal")
            ref = f"{_col_letter(int(cell['col']))}{row_n}"
            node = ET.SubElement(row, q(S, "c"), {"r": ref, "s": str(style_ids[(font, size, bold, italic)])})
            _write_xlsx_value(node, cell.get("value"))
    return _xml(root)


def _xlsx_parts(spec: dict) -> dict[str, bytes]:
    sheets = _sheets_of(spec.get("sheets"))
    fonts: list[tuple[str, float, bool, bool]] = [("Calibri", 11.0, False, False)]
    built = []
    for sheet in sheets:
        cells = []
        for item in sheet["cells"]:
            if not isinstance(item, dict):
                continue
            cell = _cell_spec(item, int(item.get("row") or 1), int(item.get("col") or 1))
            key = _font_of(cell, style="Normal")
            if key not in fonts:
                fonts.append(key)
            cells.append(cell)
        built.append((sheet["name"], cells))
    style_ids = {font: index for index, font in enumerate(fonts)}
    workbook = ET.Element(q(S, "workbook"))
    views = ET.SubElement(workbook, q(S, "bookViews"))
    ET.SubElement(views, q(S, "workbookView"), {"activeTab": "0"})
    sheets_node = ET.SubElement(workbook, q(S, "sheets"))
    parts = {
        "[Content_Types].xml": _content_types(
            [("rels", "application/vnd.openxmlformats-package.relationships+xml"), ("xml", "application/xml")],
            [("/xl/workbook.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"),
             ("/xl/styles.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"),
             ("/docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml")],
        ),
        "_rels/.rels": _package_rels("xl/workbook.xml"),
        "docProps/core.xml": _core_props(),
        "xl/styles.xml": _xlsx_styles(fonts),
        "xl/workbook.xml": b"",
    }
    rels = []
    overrides = []
    for index, (name, cells) in enumerate(built, start=1):
        sheet_path = f"xl/worksheets/sheet{index}.xml"
        parts[sheet_path] = _xlsx_sheet(cells, style_ids)
        rels.append(("http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet", f"worksheets/sheet{index}.xml"))
        ET.SubElement(sheets_node, q(S, "sheet"), {"name": name, "sheetId": str(index), q(R, "id"): f"rId{index}"})
        overrides.append((f"/xl/worksheets/sheet{index}.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"))
    rels.append(("http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles", "styles.xml"))
    parts["xl/_rels/workbook.xml.rels"] = _rels("", rels)
    parts["xl/workbook.xml"] = _xml(workbook)
    types = _parse(parts["[Content_Types].xml"])
    for part, value in overrides:
        ET.SubElement(types, q(CT, "Override"), {"PartName": part, "ContentType": value})
    parts["[Content_Types].xml"] = _xml(types)
    return parts


def _read_xlsx_styles(parts: dict[str, bytes]) -> list[dict]:
    if "xl/styles.xml" not in parts:
        return [{"font": "Calibri", "size_pt": 11, "bold": False, "italic": False}]
    root = _parse(parts["xl/styles.xml"])
    fonts = []
    for font in root.find(q(S, "fonts")) or []:
        name = font.find(q(S, "name"))
        size = font.find(q(S, "sz"))
        fonts.append({
            "font": name.get("val") if name is not None else None,
            "size_pt": float(size.get("val")) if size is not None and size.get("val") else None,
            "bold": font.find(q(S, "b")) is not None,
            "italic": font.find(q(S, "i")) is not None,
        })
    xfs = []
    for xf in root.find(q(S, "cellXfs")) or []:
        font_id = int(xf.get("fontId") or 0)
        xfs.append(fonts[font_id] if 0 <= font_id < len(fonts) else {})
    return xfs or [{"font": "Calibri", "size_pt": 11, "bold": False, "italic": False}]


def _shared_strings(parts: dict[str, bytes]) -> list[str]:
    if "xl/sharedStrings.xml" not in parts:
        return []
    root = _parse(parts["xl/sharedStrings.xml"])
    values = []
    for item in root.findall(q(S, "si")):
        values.append("".join(node.text or "" for node in item.iter(q(S, "t"))))
    return values


def _sheet_cells(sheet: ET.Element, styles: list[dict], shared: list[str] | None = None) -> list[dict]:
    cells = []
    shared = shared or []
    for node in sheet.findall(".//" + q(S, "c")):
        ref = node.get("r") or "A1"
        row, col = _split_ref(ref)
        style = styles[int(node.get("s") or 0)] if styles else {}
        if int(node.get("s") or 0) >= len(styles):
            style = {}
        kind = node.get("t")
        formula = node.find(q(S, "f"))
        value: object
        if formula is not None and formula.text:
            value = "=" + formula.text
        elif kind == "inlineStr":
            value = _text_of(node, q(S, "t"))
        elif kind == "s":
            raw = node.find(q(S, "v"))
            index = int(raw.text) if raw is not None and str(raw.text or "").isdigit() else -1
            value = shared[index] if 0 <= index < len(shared) else ""
        elif kind == "b":
            raw = node.find(q(S, "v"))
            value = str(raw.text if raw is not None else "") in {"1", "true", "TRUE"}
        else:
            raw = node.find(q(S, "v"))
            value = raw.text if raw is not None else ""
            if kind != "s":
                try:
                    value = float(value) if value and "." in str(value) else int(value) if str(value).lstrip("-").isdigit() else value
                except ValueError:
                    pass
        cells.append({"row": row, "col": col, "value": value, "font": style.get("font"), "size_pt": style.get("size_pt"), "bold": bool(style.get("bold")), "italic": bool(style.get("italic"))})
    return cells


def _workbook_sheets(parts: dict[str, bytes]) -> list[tuple[str, str]]:
    workbook = _parse(parts["xl/workbook.xml"])
    rels = _parse(parts["xl/_rels/workbook.xml.rels"])
    targets = {node.get("Id"): node.get("Target") for node in rels}
    found = []
    for sheet in workbook.findall(".//" + q(S, "sheet")):
        target = targets.get(sheet.get(q(R, "id")) or "")
        if target and not target.startswith("xl/"):
            target = "xl/" + target.lstrip("/")
        found.append((sheet.get("name") or "Sheet", target or ""))
    return found


def _read_xlsx(parts: dict[str, bytes]) -> dict:
    styles = _read_xlsx_styles(parts)
    shared = _shared_strings(parts)
    sheets = []
    for name, target in _workbook_sheets(parts):
        if target not in parts:
            continue
        cells = _sheet_cells(_parse(parts[target]), styles, shared)
        rows = max((cell["row"] for cell in cells), default=0)
        cols = max((cell["col"] for cell in cells), default=0)
        fonts = []
        sizes = []
        for cell in cells:
            if cell["font"] and cell["font"] not in fonts:
                fonts.append(cell["font"])
            if cell["size_pt"] and cell["size_pt"] not in sizes:
                sizes.append(cell["size_pt"])
        sheets.append({"name": name, "rows": rows, "cols": cols, "cells": cells, "fonts": fonts, "sizes_pt": sizes})
    return {"sheets": sheets}


def _xlsx_refresh_widths(root: ET.Element, shared: list[str] | None = None) -> None:
    shared = shared or []
    old = root.find(q(S, "cols"))
    if old is not None:
        root.remove(old)
    widths: dict[int, float] = {}
    for node in root.findall(".//" + q(S, "c")):
        _row, col = _split_ref(node.get("r") or "A1")
        kind = node.get("t")
        if kind == "inlineStr":
            text = _text_of(node, q(S, "t"))
        else:
            raw = node.find(q(S, "v"))
            text = "" if raw is None or raw.text is None else str(raw.text)
            if kind == "s" and str(text).isdigit():
                index = int(text)
                text = shared[index] if 0 <= index < len(shared) else ""
        widths[col] = max(widths.get(col, 10.0), _xlsx_width(text))
    if not widths:
        return
    cols = _xlsx_cols(widths)
    data = root.find(q(S, "sheetData"))
    if data is None:
        root.append(cols)
    else:
        root.insert(list(root).index(data), cols)


def _sheet_target(parts: dict[str, bytes], wanted: str) -> str:
    sheets = _workbook_sheets(parts)
    for index, (name, target) in enumerate(sheets, start=1):
        if wanted == name or wanted == str(index):
            if target not in parts:
                break
            return target
    raise OfficeError(f"Sheet {wanted} was not found.")


def _used_rows(root: ET.Element) -> int:
    rows = [int(item.get("r") or 0) for item in root.findall(".//" + q(S, "row"))]
    return max(rows) if rows else 0


def _expand_sheet_ops(parts: dict[str, bytes], ops: list) -> list:
    expanded = []
    for op in ops:
        name = str(op.get("op") or "set_sheet_cell")
        sheet = op.get("sheet") or op.get("name") or "1"
        if name in {"set_sheet_cell", "set_cell"}:
            item = dict(op)
            item["op"] = "set_sheet_cell"
            item["sheet"] = sheet
            expanded.append(item)
            continue
        if name in {"set_cells", "write_cells"}:
            for cell in _cells_of(op.get("cells") or op.get("values") or []):
                expanded.append({
                    "op": "set_sheet_cell",
                    "sheet": sheet,
                    "row": cell.get("row"),
                    "col": cell.get("col"),
                    "value": cell.get("value"),
                    "font": cell.get("font", op.get("font")),
                    "size_pt": cell.get("size_pt", op.get("size_pt")),
                    "bold": cell.get("bold", op.get("bold")),
                    "italic": cell.get("italic", op.get("italic")),
                })
            continue
        if name in {"append_row", "add_row"}:
            target = _sheet_target(parts, str(sheet))
            root = _parse(parts[target])
            row_n = _used_rows(root) + 1
            values = op.get("values") if op.get("values") is not None else op.get("cells") or []
            if isinstance(values, dict):
                values = _cells_of(values)
            for col, value in enumerate(list(values), start=1):
                if isinstance(value, dict):
                    expanded.append({
                        "op": "set_sheet_cell",
                        "sheet": sheet,
                        "row": value.get("row") or row_n,
                        "col": value.get("col") or col,
                        "value": value.get("value", value.get("text")),
                        "font": value.get("font", op.get("font")),
                        "bold": value.get("bold", op.get("bold")),
                        "italic": value.get("italic", op.get("italic")),
                        "size_pt": value.get("size_pt", op.get("size_pt")),
                    })
                else:
                    expanded.append({
                        "op": "set_sheet_cell",
                        "sheet": sheet,
                        "row": row_n,
                        "col": col,
                        "value": value,
                    })
            continue
        raise OfficeError(
            "Spreadsheets take set_sheet_cell, set_cells, or append_row. "
            "A formula is a value that starts with =."
        )
    return expanded


def _edit_xlsx(parts: dict[str, bytes], ops: list) -> None:
    pending = [dict(op) for op in ops]
    while pending:
        op = pending.pop(0)
        name = str(op.get("op") or "set_sheet_cell")
        if name in {"append_row", "add_row", "set_cells", "write_cells"}:
            pending = _expand_sheet_ops(parts, [op]) + pending
            continue
        if name not in {"set_sheet_cell", "set_cell"}:
            raise OfficeError(
                "Spreadsheets take set_sheet_cell, set_cells, or append_row. "
                "A formula is a value that starts with =."
            )
        wanted = str(op.get("sheet") or op.get("name") or "1")
        match = _sheet_target(parts, wanted)
        root = _parse(parts[match])
        data = root.find(q(S, "sheetData"))
        if data is None:
            data = ET.SubElement(root, q(S, "sheetData"))
        row_n = int(op.get("row") or 1)
        col_n = int(op.get("col") or 1)
        if row_n < 1 or col_n < 1:
            raise OfficeError("Row and column start at 1.")
        ref = f"{_col_letter(col_n)}{row_n}"
        row = next((item for item in data.findall(q(S, "row")) if item.get("r") == str(row_n)), None)
        if row is None:
            row = ET.Element(q(S, "row"), {"r": str(row_n)})
            inserted = False
            for existing in list(data):
                if int(existing.get("r") or 0) > row_n:
                    _insert_before(data, existing, row)
                    inserted = True
                    break
            if not inserted:
                data.append(row)
        cell = next((item for item in row.findall(q(S, "c")) if item.get("r") == ref), None)
        if cell is None:
            cell = ET.SubElement(row, q(S, "c"), {"r": ref})
        if _style_requested(op):
            key = _font_of(op, style="Normal")
            style_id = _xlsx_ensure_style(parts, key)
            if style_id is not None:
                cell.set("s", style_id)
        _write_xlsx_value(cell, _cell_value(op.get("value", op.get("text"))))
        parts[match] = _xml(root)


_THEME = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Kerniq">
  <a:themeElements>
    <a:clrScheme name="Kerniq">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="1F4E79"/></a:dk2>
      <a:lt2><a:srgbClr val="EEECE1"/></a:lt2>
      <a:accent1><a:srgbClr val="5B9BD5"/></a:accent1>
      <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
      <a:accent3><a:srgbClr val="A9D08E"/></a:accent3>
      <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
      <a:accent5><a:srgbClr val="4472C4"/></a:accent5>
      <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Kerniq">
      <a:majorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
      <a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Kerniq">
      <a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>
      <a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>
      <a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>
      <a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
</a:theme>
"""

_SLIDE_MASTER = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld name="Kerniq"><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
  </p:spTree></p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/></p:sldLayoutIdLst>
</p:sldMaster>
"""

_SLIDE_LAYOUT = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="title">
  <p:cSld name="Title"><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
  </p:spTree></p:cSld>
</p:sldLayout>
"""


def _slide_shape(parent: ET.Element, index: int, name: str, text: str, size: float, bold: bool, font: str, x: int, y: int, cx: int, cy: int) -> None:
    shape = ET.SubElement(parent, q(P, "sp"))
    nv = ET.SubElement(shape, q(P, "nvSpPr"))
    ET.SubElement(nv, q(P, "cNvPr"), {"id": str(index), "name": name})
    ET.SubElement(nv, q(P, "cNvSpPr"))
    ET.SubElement(nv, q(P, "nvPr"))
    props = ET.SubElement(shape, q(P, "spPr"))
    xfrm = ET.SubElement(props, q(A, "xfrm"))
    ET.SubElement(xfrm, q(A, "off"), {"x": str(x), "y": str(y)})
    ET.SubElement(xfrm, q(A, "ext"), {"cx": str(cx), "cy": str(cy)})
    tx = ET.SubElement(shape, q(P, "txBody"))
    ET.SubElement(tx, q(A, "bodyPr"))
    ET.SubElement(tx, q(A, "lstStyle"))
    paragraph = ET.SubElement(tx, q(A, "p"))
    run = ET.SubElement(paragraph, q(A, "r"))
    run_props = ET.SubElement(run, q(A, "rPr"), {"sz": str(int(round(size * 100))), "b": "1" if bold else "0"})
    ET.SubElement(run_props, q(A, "latin"), {"typeface": font})
    ET.SubElement(run, q(A, "t")).text = text or ""


def _slide_xml(title: str, body: str, *, font: str, title_pt: float, body_pt: float, bold: bool) -> bytes:
    root = ET.Element(q(P, "sld"))
    tree = ET.SubElement(ET.SubElement(root, q(P, "cSld")), q(P, "spTree"))
    group = ET.SubElement(tree, q(P, "nvGrpSpPr"))
    ET.SubElement(group, q(P, "cNvPr"), {"id": "1", "name": ""})
    ET.SubElement(group, q(P, "cNvGrpSpPr"))
    ET.SubElement(group, q(P, "nvPr"))
    grp = ET.SubElement(tree, q(P, "grpSpPr"))
    xfrm = ET.SubElement(grp, q(A, "xfrm"))
    ET.SubElement(xfrm, q(A, "off"), {"x": "0", "y": "0"})
    ET.SubElement(xfrm, q(A, "ext"), {"cx": "0", "cy": "0"})
    ET.SubElement(xfrm, q(A, "chOff"), {"x": "0", "y": "0"})
    ET.SubElement(xfrm, q(A, "chExt"), {"cx": "0", "cy": "0"})
    _slide_shape(tree, 2, "Title", title, title_pt, True, font, 457200, 274638, 8229600, 1143000)
    _slide_shape(tree, 3, "Body", body, body_pt, bold, font, 457200, 1600200, 8229600, 4525963)
    return _xml(root)


def _pptx_support() -> dict[str, bytes]:
    return {
        "ppt/theme/theme1.xml": _THEME.encode("utf-8"),
        "ppt/slideMasters/slideMaster1.xml": _SLIDE_MASTER.encode("utf-8"),
        "ppt/slideLayouts/slideLayout1.xml": _SLIDE_LAYOUT.encode("utf-8"),
        "ppt/slideMasters/_rels/slideMaster1.xml.rels": _rels("", [
            ("http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout", "../slideLayouts/slideLayout1.xml"),
            ("http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme", "../theme/theme1.xml"),
        ]),
        "ppt/slideLayouts/_rels/slideLayout1.xml.rels": _rels("", [
            ("http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster", "../slideMasters/slideMaster1.xml"),
        ]),
        "ppt/presProps.xml": b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<p:presentationPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>',
        "ppt/viewProps.xml": b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<p:viewPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:normalViewPr/><p:slideViewPr><p:cSldViewPr/></p:slideViewPr></p:viewPr>',
        "ppt/tableStyles.xml": b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<p:tblStyleLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>',
        "docProps/core.xml": _core_props(),
    }


def _pptx_parts(spec: dict) -> dict[str, bytes]:
    slides = _slides_of(spec.get("slides"), str(spec.get("title") or ""))
    presentation = ET.Element(q(P, "presentation"))
    listing = ET.SubElement(presentation, q(P, "sldIdLst"))
    parts: dict[str, bytes] = _pptx_support()
    rels = []
    overrides = [
        ("/ppt/presentation.xml", "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"),
        ("/ppt/slideMasters/slideMaster1.xml", "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"),
        ("/ppt/slideLayouts/slideLayout1.xml", "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"),
        ("/ppt/theme/theme1.xml", "application/vnd.openxmlformats-officedocument.theme+xml"),
        ("/ppt/presProps.xml", "application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"),
        ("/ppt/viewProps.xml", "application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"),
        ("/ppt/tableStyles.xml", "application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"),
        ("/docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml"),
    ]
    for index, slide in enumerate(slides, start=1):
        font = str(slide.get("font") or "Calibri")
        parts[f"ppt/slides/slide{index}.xml"] = _slide_xml(
            str(slide.get("title") or ""),
            str(slide.get("body") or ""),
            font=font,
            title_pt=float(slide.get("title_pt") or slide.get("size_pt") or 32),
            body_pt=float(slide.get("body_pt") or 18),
            bold=bool(slide.get("bold")),
        )
        parts[f"ppt/slides/_rels/slide{index}.xml.rels"] = _rels("", [
            ("http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout", "../slideLayouts/slideLayout1.xml"),
        ])
        rels.append(("http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide", f"slides/slide{index}.xml"))
        ET.SubElement(listing, q(P, "sldId"), {"id": str(255 + index), q(R, "id"): f"rId{index}"})
        overrides.append((f"/ppt/slides/slide{index}.xml", "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"))
    rels.extend([
        ("http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster", "slideMasters/slideMaster1.xml"),
        ("http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps", "presProps.xml"),
        ("http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps", "viewProps.xml"),
        ("http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles", "tableStyles.xml"),
    ])
    ET.SubElement(presentation, q(P, "sldSz"), {"cx": "9144000", "cy": "6858000", "type": "screen4x3"})
    ET.SubElement(presentation, q(P, "notesSz"), {"cx": "6858000", "cy": "9144000"})
    parts["ppt/presentation.xml"] = _xml(presentation)
    parts["ppt/_rels/presentation.xml.rels"] = _rels("", rels)
    parts["_rels/.rels"] = _package_rels("ppt/presentation.xml")
    parts["[Content_Types].xml"] = _content_types(
        [("rels", "application/vnd.openxmlformats-package.relationships+xml"), ("xml", "application/xml")],
        overrides,
    )
    return parts


def _read_pptx(parts: dict[str, bytes]) -> dict:
    slides = []
    names = sorted(
        (name for name in parts if name.startswith("ppt/slides/slide") and name.endswith(".xml") and "/_rels/" not in name),
        key=lambda name: int("".join(ch for ch in Path(name).stem if ch.isdigit()) or "0"),
    )
    for index, name in enumerate(names, start=1):
        root = _parse(parts[name])
        texts = []
        font = None
        size = None
        bold = False
        for shape in root.iter(q(P, "sp")):
            texts.append(_text_of(shape, q(A, "t")))
            props = shape.find(".//" + q(A, "rPr"))
            if props is not None and font is None:
                latin = props.find(q(A, "latin"))
                font = latin.get("typeface") if latin is not None else None
                if props.get("sz"):
                    size = float(props.get("sz")) / 100
                bold = props.get("b") == "1"
        slides.append({
            "n": index,
            "title": texts[0] if texts else "",
            "body": "\n".join(texts[1:]),
            "font": font,
            "size_pt": size,
            "bold": bold,
        })
    return {"slides": slides}


def _edit_pptx(parts: dict[str, bytes], ops: list) -> None:
    for op in ops:
        name = str(op.get("op") or "")
        names = sorted(
            (item for item in parts if item.startswith("ppt/slides/slide") and item.endswith(".xml") and "/_rels/" not in item),
            key=lambda name: int("".join(ch for ch in Path(name).stem if ch.isdigit()) or "0"),
        )
        if name == "set_slide":
            number = int(op.get("number") or op.get("index") or 0)
            if number < 1 or number > len(names):
                raise OfficeError(f"Slide {number} is outside 1..{len(names)}.")
            current = _read_pptx(parts)["slides"][number - 1]
            parts[names[number - 1]] = _slide_xml(
                str(op.get("title", current["title"])),
                str(op.get("body", current["body"])),
                font=str(op.get("font") or current.get("font") or "Calibri"),
                title_pt=float(op.get("title_pt") or op.get("size_pt") or current.get("size_pt") or 32),
                body_pt=float(op.get("body_pt") or 18),
                bold=_flag(op, "bold", bool(current.get("bold"))),
            )
        elif name == "add_slide":
            index = len(names) + 1
            parts[f"ppt/slides/slide{index}.xml"] = _slide_xml(
                str(op.get("title") or ""),
                str(op.get("body") or ""),
                font=str(op.get("font") or "Calibri"),
                title_pt=float(op.get("title_pt") or op.get("size_pt") or 32),
                body_pt=float(op.get("body_pt") or 18),
                bold=bool(op.get("bold")),
            )
            parts[f"ppt/slides/_rels/slide{index}.xml.rels"] = _rels("", [
                ("http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout", "../slideLayouts/slideLayout1.xml"),
            ])
            rels = _parse(parts["ppt/_rels/presentation.xml.rels"])
            used = {node.get("Id") for node in rels}
            number = 1
            while f"rId{number}" in used:
                number += 1
            rel_id = f"rId{number}"
            ET.SubElement(rels, q(PR, "Relationship"), {
                "Id": rel_id,
                "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
                "Target": f"slides/slide{index}.xml",
            })
            parts["ppt/_rels/presentation.xml.rels"] = _xml(rels)
            presentation = _parse(parts["ppt/presentation.xml"])
            listing = presentation.find(q(P, "sldIdLst"))
            if listing is None:
                listing = ET.Element(q(P, "sldIdLst"))
                presentation.insert(0, listing)
            ET.SubElement(listing, q(P, "sldId"), {"id": str(255 + index), q(R, "id"): rel_id})
            parts["ppt/presentation.xml"] = _xml(presentation)
            types = _parse(parts["[Content_Types].xml"])
            ET.SubElement(types, q(CT, "Override"), {
                "PartName": f"/ppt/slides/slide{index}.xml",
                "ContentType": "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
            })
            parts["[Content_Types].xml"] = _xml(types)
        else:
            raise OfficeError(f"{name or 'missing op'} is not a slide edit. Use set_slide or add_slide.")


def _odf_automatic(name: str, parent: str, font: str, size: float, bold: bool, italic: bool) -> ET.Element:
    style = ET.Element(q(STYLE, "style"), {
        q(STYLE, "name"): name,
        q(STYLE, "family"): "paragraph",
        q(STYLE, "parent-style-name"): parent,
    })
    ET.SubElement(style, q(STYLE, "text-properties"), {
        q(FO, "font-size"): f"{size:g}pt",
        q(FO, "font-weight"): "bold" if bold else "normal",
        q(FO, "font-style"): "italic" if italic else "normal",
        q(STYLE, "font-name"): font,
    })
    return style


def _odf_paragraph(spec: dict, auto: ET.Element, counter: list[int]) -> ET.Element:
    style = canon_style(spec.get("style")) or "Normal"
    font, size, bold, italic = _font_of(spec, style=style)
    counter[0] += 1
    name = f"K{counter[0]}"
    auto.append(_odf_automatic(name, _odf_style_name(style), font, size, bold, italic))
    if style in {"Heading1", "Heading2", "Title"}:
        level = {"Heading1": "1", "Heading2": "2", "Title": "1"}[style]
        node = ET.Element(q(TEXT, "h"), {q(TEXT, "style-name"): name, q(TEXT, "outline-level"): level})
    else:
        node = ET.Element(q(TEXT, "p"), {q(TEXT, "style-name"): name})
    node.text = _as_text(spec.get("text"))
    return node


def _cm_for_text(text: str) -> str:
    chars = min(60, max(10, len(text) + 2))
    return f"{chars * 0.22:.2f}cm"


def _odf_column_style(auto: ET.Element, counter: list[int], width: str) -> str:
    counter[0] += 1
    name = f"KC{counter[0]}"
    style = ET.SubElement(auto, q(STYLE, "style"), {
        q(STYLE, "name"): name,
        q(STYLE, "family"): "table-column",
    })
    ET.SubElement(style, q(STYLE, "table-column-properties"), {q(STYLE, "column-width"): width})
    return name


def _odf_table(spec: dict, auto: ET.Element, counter: list[int]) -> ET.Element:
    rows = _table_grid(spec)
    width = max(len(row) if isinstance(row, list) else 1 for row in rows)
    width = max(width, 1)
    each = round(16.0 / width, 2)
    total = round(each * width, 2)
    counter[0] += 1
    table_style = f"KT{counter[0]}"
    style = ET.SubElement(auto, q(STYLE, "style"), {
        q(STYLE, "name"): table_style,
        q(STYLE, "family"): "table",
    })
    ET.SubElement(style, q(STYLE, "table-properties"), {
        q(STYLE, "width"): f"{total:.2f}cm",
        q(TABLE, "align"): "margins",
    })
    column_style = _odf_column_style(auto, counter, f"{each:.2f}cm")
    table = ET.Element(q(TABLE, "table"), {
        q(TABLE, "name"): spec.get("name") or "Table",
        q(TABLE, "style-name"): table_style,
    })
    for _ in range(width):
        ET.SubElement(table, q(TABLE, "table-column"), {q(TABLE, "style-name"): column_style})
    for row in rows:
        values = row if isinstance(row, list) else [row]
        tr = ET.SubElement(table, q(TABLE, "table-row"))
        for col in range(width):
            raw = values[col] if col < len(values) else ""
            cell_spec = raw if isinstance(raw, dict) else {"text": "" if raw is None else str(raw)}
            tc = ET.SubElement(tr, q(TABLE, "table-cell"))
            tc.append(_odf_paragraph(cell_spec, auto, counter))
    return table


def _odf_styles() -> bytes:
    root = ET.Element(q(OFFICE, "document-styles"), {q(OFFICE, "version"): "1.2"})
    faces = ET.SubElement(root, q(OFFICE, "font-face-decls"))
    for name in ("Liberation Serif", "Liberation Sans", "Calibri", "Times New Roman"):
        ET.SubElement(faces, q(STYLE, "font-face"), {q(STYLE, "name"): name})
    styles = ET.SubElement(root, q(OFFICE, "styles"))
    for name, size, bold in (("Standard", 11, False), ("Heading 1", 16, True), ("Heading 2", 13, True), ("Title", 28, True)):
        style = ET.SubElement(styles, q(STYLE, "style"), {q(STYLE, "name"): name, q(STYLE, "family"): "paragraph"})
        ET.SubElement(style, q(STYLE, "text-properties"), {
            q(FO, "font-size"): f"{size}pt",
            q(FO, "font-weight"): "bold" if bold else "normal",
            q(STYLE, "font-name"): "Liberation Serif" if name == "Standard" else "Liberation Sans",
        })
    return _xml(root)


def _odf_manifest(mime: str) -> bytes:
    root = ET.Element(q(MANIFEST, "manifest"), {q(MANIFEST, "version"): "1.2"})
    ET.SubElement(root, q(MANIFEST, "file-entry"), {q(MANIFEST, "full-path"): "/", q(MANIFEST, "media-type"): mime})
    ET.SubElement(root, q(MANIFEST, "file-entry"), {q(MANIFEST, "full-path"): "content.xml", q(MANIFEST, "media-type"): "text/xml"})
    ET.SubElement(root, q(MANIFEST, "file-entry"), {q(MANIFEST, "full-path"): "styles.xml", q(MANIFEST, "media-type"): "text/xml"})
    ET.SubElement(root, q(MANIFEST, "file-entry"), {q(MANIFEST, "full-path"): "meta.xml", q(MANIFEST, "media-type"): "text/xml"})
    return _xml(root)


def _odf_meta() -> bytes:
    return (
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        '<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" '
        'xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" office:version="1.2">'
        '<office:meta><meta:generator>Kerniq Office</meta:generator></office:meta>'
        '</office:document-meta>'
    ).encode("utf-8")


def _odf_parts(suffix: str, spec: dict) -> dict[str, bytes]:
    root = ET.Element(q(OFFICE, "document-content"), {q(OFFICE, "version"): "1.2"})
    ET.SubElement(root, q(OFFICE, "font-face-decls"))
    auto = ET.SubElement(root, q(OFFICE, "automatic-styles"))
    office_body = ET.SubElement(root, q(OFFICE, "body"))
    counter = [0]
    if suffix == ".odt":
        host = ET.SubElement(office_body, q(OFFICE, "text"))
        paragraphs = _paragraphs_of(spec.get("paragraphs"))
        if spec.get("title"):
            paragraphs.insert(0, {"text": spec["title"], "style": "Title"})
        if not paragraphs:
            paragraphs = [{"text": ""}]
        for item in paragraphs:
            host.append(_odf_paragraph(item, auto, counter))
        for table in _as_list(spec.get("tables")):
            host.append(_odf_table(table if isinstance(table, dict) else {"cells": table}, auto, counter))
    elif suffix == ".ods":
        host = ET.SubElement(office_body, q(OFFICE, "spreadsheet"))
        for sheet in _sheets_of(spec.get("sheets")):
            table = ET.SubElement(host, q(TABLE, "table"), {q(TABLE, "name"): sheet["name"]})
            cells = sheet["cells"]
            if not cells:
                ET.SubElement(table, q(TABLE, "table-column"))
                row = ET.SubElement(table, q(TABLE, "table-row"))
                ET.SubElement(row, q(TABLE, "table-cell"))
                continue
            width = max(int(item.get("col") or 1) for item in cells)
            height = max(int(item.get("row") or 1) for item in cells)
            grid = {(int(item.get("row") or 1), int(item.get("col") or 1)): item for item in cells}
            for col_n in range(1, width + 1):
                longest = ""
                for row_n in range(1, height + 1):
                    item = grid.get((row_n, col_n))
                    if not item:
                        continue
                    value = item.get("value", item.get("text"))
                    text = "" if value is None else str(value)
                    if len(text) > len(longest):
                        longest = text
                style_name = _odf_column_style(auto, counter, _cm_for_text(longest))
                ET.SubElement(table, q(TABLE, "table-column"), {q(TABLE, "style-name"): style_name})
            for row_n in range(1, height + 1):
                row = ET.SubElement(table, q(TABLE, "table-row"))
                for col_n in range(1, width + 1):
                    item = grid.get((row_n, col_n), {"value": ""})
                    tc = ET.SubElement(row, q(TABLE, "table-cell"))
                    value = item.get("value", item.get("text"))
                    if isinstance(value, (int, float)) and not isinstance(value, bool):
                        tc.set(q(OFFICE, "value-type"), "float")
                        tc.set(q(OFFICE, "value"), str(value))
                    else:
                        tc.set(q(OFFICE, "value-type"), "string")
                    tc.append(_odf_paragraph({"text": "" if value is None else str(value), **item}, auto, counter))
    else:
        host = ET.SubElement(office_body, q(OFFICE, "presentation"))
        for index, slide in enumerate(_slides_of(spec.get("slides"), str(spec.get("title") or "")), start=1):
            page = ET.SubElement(host, q(DRAW, "page"), {q(DRAW, "name"): f"page{index}"})
            frame = ET.SubElement(page, q(DRAW, "frame"))
            box = ET.SubElement(frame, q(DRAW, "text-box"))
            box.append(_odf_paragraph({"text": slide.get("title") or "", "style": "Title", "font": slide.get("font"), "size_pt": slide.get("size_pt") or 32}, auto, counter))
            box.append(_odf_paragraph({"text": slide.get("body") or "", "font": slide.get("font"), "size_pt": slide.get("body_pt") or 18, "bold": slide.get("bold")}, auto, counter))
    return {
        "content.xml": _xml(root),
        "styles.xml": _odf_styles(),
        "meta.xml": _odf_meta(),
        "META-INF/manifest.xml": _odf_manifest(MIME[suffix]),
    }


def _style_props(auto: ET.Element, styles_root: ET.Element | None, name: str | None) -> dict:
    info = {"style": "Normal", "font": None, "size_pt": None, "bold": False, "italic": False}
    if not name:
        return info
    node = None
    for candidate in list(auto) + list(styles_root or []):
        if candidate.get(q(STYLE, "name")) == name:
            node = candidate
            break
    parent = node.get(q(STYLE, "parent-style-name")) if node is not None else name
    info["style"] = canon_style(parent) or canon_style(name) or "Normal"
    props = None if node is None else node.find(q(STYLE, "text-properties"))
    if props is not None:
        info["font"] = props.get(q(STYLE, "font-name"))
        size = props.get(q(FO, "font-size"))
        if size and size.endswith("pt"):
            info["size_pt"] = float(size[:-2])
        info["bold"] = props.get(q(FO, "font-weight")) == "bold"
        info["italic"] = props.get(q(FO, "font-style")) == "italic"
    return info


def _odf_auto_and_styles(parts: dict[str, bytes]) -> tuple[ET.Element, ET.Element, ET.Element | None]:
    root = _parse(parts["content.xml"])
    auto = root.find(q(OFFICE, "automatic-styles"))
    if auto is None:
        auto = ET.Element(q(OFFICE, "automatic-styles"))
        root.insert(0, auto)
    styles_root = None
    if "styles.xml" in parts:
        styles_root = _parse(parts["styles.xml"]).find(q(OFFICE, "styles"))
    return root, auto, styles_root


def _read_odt(parts: dict[str, bytes]) -> dict:
    root, auto, styles_root = _odf_auto_and_styles(parts)
    host = root.find(q(OFFICE, "body"))
    text = None if host is None else host.find(q(OFFICE, "text"))
    paragraphs = []
    tables = []
    seen = []
    if text is None:
        return {"paragraphs": [], "tables": [], "styles": []}
    for child in list(text):
        if child.tag in {q(TEXT, "p"), q(TEXT, "h")}:
            props = _style_props(auto, styles_root, child.get(q(TEXT, "style-name")))
            if child.tag == q(TEXT, "h") and props["style"] not in {"Title", "Heading1", "Heading2"}:
                level = child.get(q(TEXT, "outline-level"))
                props["style"] = "Heading2" if level == "2" else "Heading1"
            paragraphs.append({"n": len(paragraphs) + 1, "text": "".join(child.itertext()), **props})
            if props["style"] not in seen:
                seen.append(props["style"])
        elif child.tag == q(TABLE, "table"):
            rows = []
            fonts, sizes = [], []
            for tr in child.findall(q(TABLE, "table-row")):
                row = []
                for tc in tr.findall(q(TABLE, "table-cell")):
                    chunks = []
                    for paragraph in list(tc):
                        if paragraph.tag not in {q(TEXT, "p"), q(TEXT, "h")}:
                            continue
                        props = _style_props(auto, styles_root, paragraph.get(q(TEXT, "style-name")))
                        chunks.append("".join(paragraph.itertext()))
                        if props["font"] and props["font"] not in fonts:
                            fonts.append(props["font"])
                        if props["size_pt"] and props["size_pt"] not in sizes:
                            sizes.append(props["size_pt"])
                    row.append("\n".join(chunks))
                rows.append(row)
            tables.append({
                "n": len(tables) + 1,
                "rows": len(rows),
                "cols": max((len(row) for row in rows), default=0),
                "cells": rows,
                "fonts": fonts,
                "sizes_pt": sizes,
            })
    return {"paragraphs": paragraphs, "tables": tables, "styles": seen}


def _edit_odt(parts: dict[str, bytes], ops: list) -> None:
    root, auto, _styles = _odf_auto_and_styles(parts)
    text = root.find(q(OFFICE, "body")).find(q(OFFICE, "text"))
    counter = [len(list(auto))]
    for op in ops:
        name = str(op.get("op") or "")
        blocks = [child for child in list(text) if child.tag in {q(TEXT, "p"), q(TEXT, "h")}]
        if name in {"set_paragraph", "set_style"}:
            index = int(op.get("index") or 0)
            if index < 1 or index > len(blocks):
                raise OfficeError(f"Paragraph {index} is outside 1..{len(blocks)}.")
            current_props = _style_props(auto, None, blocks[index - 1].get(q(TEXT, "style-name")))
            spec = dict(op)
            if name == "set_style":
                spec["text"] = "".join(blocks[index - 1].itertext())
            spec.setdefault("text", "".join(blocks[index - 1].itertext()))
            spec.setdefault("style", current_props["style"])
            spec.setdefault("font", current_props["font"])
            spec.setdefault("size_pt", current_props["size_pt"])
            spec.setdefault("bold", current_props["bold"])
            spec.setdefault("italic", current_props["italic"])
            fresh = _odf_paragraph(spec, auto, counter)
            blocks[index - 1].clear()
            blocks[index - 1].tag = fresh.tag
            blocks[index - 1].attrib.clear()
            blocks[index - 1].attrib.update(fresh.attrib)
            blocks[index - 1].text = fresh.text
            blocks[index - 1].extend(list(fresh))
        elif name == "add_paragraph":
            text.append(_odf_paragraph(op, auto, counter))
        elif name == "add_table":
            text.append(_odf_table(op, auto, counter))
        elif name == "set_cell":
            tables = [child for child in list(text) if child.tag == q(TABLE, "table")]
            table_n = int(op.get("table") or 0)
            row_n = int(op.get("row") or 0)
            col_n = int(op.get("col") or 0)
            if table_n < 1 or table_n > len(tables):
                raise OfficeError(f"Table {table_n} is outside 1..{len(tables)}.")
            rows = tables[table_n - 1].findall(q(TABLE, "table-row"))
            if row_n < 1 or row_n > len(rows):
                raise OfficeError(f"Row {row_n} is outside 1..{len(rows)} in table {table_n}.")
            cells = rows[row_n - 1].findall(q(TABLE, "table-cell"))
            if col_n < 1 or col_n > len(cells):
                raise OfficeError(f"Column {col_n} is outside 1..{len(cells)} in table {table_n}.")
            cell = cells[col_n - 1]
            for child in list(cell):
                cell.remove(child)
            cell.append(_odf_paragraph(op, auto, counter))
        else:
            raise OfficeError(f"{name or 'missing op'} is not a document edit. Use set_paragraph, add_paragraph, set_cell, add_table, or set_style.")
    parts["content.xml"] = _xml(root)


def _read_ods(parts: dict[str, bytes]) -> dict:
    root, auto, _styles = _odf_auto_and_styles(parts)
    host = root.find(q(OFFICE, "body")).find(q(OFFICE, "spreadsheet"))
    sheets = []
    if host is None:
        return {"sheets": []}
    for table in host.findall(q(TABLE, "table")):
        cells = []
        fonts, sizes = [], []
        for row_n, row in enumerate(table.findall(q(TABLE, "table-row")), start=1):
            col_n = 1
            for cell in row.findall(q(TABLE, "table-cell")):
                repeat = int(cell.get(q(TABLE, "number-columns-repeated")) or 1)
                formula = cell.get(q(TABLE, "formula"))
                value = cell.get(q(OFFICE, "value"))
                if formula:
                    text = formula[3:] if str(formula).startswith("of:") else str(formula)
                    value = text if str(text).startswith("=") else "=" + str(text).lstrip("=")
                elif cell.get(q(OFFICE, "value-type")) == "boolean":
                    value = str(cell.get(q(OFFICE, "boolean-value")) or "").lower() == "true"
                elif value is None:
                    value = "".join(cell.itertext())
                else:
                    try:
                        value = float(value) if "." in value else int(value)
                    except ValueError:
                        pass
                paragraph = next((item for item in cell if item.tag in {q(TEXT, "p"), q(TEXT, "h")}), None)
                props = _style_props(auto, None, None if paragraph is None else paragraph.get(q(TEXT, "style-name")))
                for _ in range(repeat):
                    if value != "":
                        cells.append({"row": row_n, "col": col_n, "value": value, **{k: props[k] for k in ("font", "size_pt", "bold", "italic")}})
                        if props["font"] and props["font"] not in fonts:
                            fonts.append(props["font"])
                        if props["size_pt"] and props["size_pt"] not in sizes:
                            sizes.append(props["size_pt"])
                    col_n += 1
        sheets.append({
            "name": table.get(q(TABLE, "name")) or f"Sheet{len(sheets) + 1}",
            "rows": max((cell["row"] for cell in cells), default=0),
            "cols": max((cell["col"] for cell in cells), default=0),
            "cells": cells,
            "fonts": fonts,
            "sizes_pt": sizes,
        })
    return {"sheets": sheets}


def _edit_ods(parts: dict[str, bytes], ops: list) -> None:
    root, auto, _styles = _odf_auto_and_styles(parts)
    host = root.find(q(OFFICE, "body")).find(q(OFFICE, "spreadsheet"))
    counter = [len(list(auto))]
    tables = [] if host is None else host.findall(q(TABLE, "table"))
    prepared = []
    for op in ops:
        name = str(op.get("op") or "set_sheet_cell")
        sheet = op.get("sheet") or "1"
        if name in {"set_cells", "write_cells"}:
            for cell in _cells_of(op.get("cells") or op.get("values") or []):
                prepared.append({
                    "op": "set_sheet_cell",
                    "sheet": sheet,
                    "row": cell.get("row"),
                    "col": cell.get("col"),
                    "value": cell.get("value"),
                })
        elif name in {"append_row", "add_row", "set_sheet_cell", "set_cell"}:
            item = dict(op)
            item["op"] = name if name in {"append_row", "add_row"} else "set_sheet_cell"
            item["sheet"] = sheet
            prepared.append(item)
        else:
            raise OfficeError("Spreadsheets take set_sheet_cell, set_cells, or append_row.")
    for op in prepared:
        wanted = str(op.get("sheet") or "1")
        table = None
        for index, candidate in enumerate(tables, start=1):
            if wanted == (candidate.get(q(TABLE, "name")) or "") or wanted == str(index):
                table = candidate
                break
        if table is None:
            raise OfficeError(f"Sheet {wanted} was not found.")
        if op["op"] == "append_row":
            row_n = len(table.findall(q(TABLE, "table-row"))) + 1
            values = op.get("values") if op.get("values") is not None else op.get("cells") or []
            for col, value in enumerate(list(values), start=1):
                prepared.append({
                    "op": "set_sheet_cell",
                    "sheet": wanted,
                    "row": row_n,
                    "col": col,
                    "value": value.get("value") if isinstance(value, dict) else value,
                })
            continue
        row_n = int(op.get("row") or 1)
        col_n = int(op.get("col") or 1)
        rows = table.findall(q(TABLE, "table-row"))
        while len(rows) < row_n:
            rows.append(ET.SubElement(table, q(TABLE, "table-row")))
        row = rows[row_n - 1]
        cells = row.findall(q(TABLE, "table-cell"))
        while len(cells) < col_n:
            cells.append(ET.SubElement(row, q(TABLE, "table-cell")))
        cell = cells[col_n - 1]
        for child in list(cell):
            cell.remove(child)
        value = _cell_value(op.get("value", op.get("text")))
        for child in list(cell):
            cell.remove(child)
        if _is_formula(value):
            cell.set(q(OFFICE, "value-type"), "float")
            cell.set(q(TABLE, "formula"), "of:" + str(value))
            spec = dict(op)
            spec["text"] = str(value)
        elif isinstance(value, bool):
            cell.set(q(OFFICE, "value-type"), "boolean")
            cell.set(q(OFFICE, "boolean-value"), "true" if value else "false")
            spec = dict(op)
            spec["text"] = "TRUE" if value else "FALSE"
        elif isinstance(value, (int, float)) and not isinstance(value, bool):
            cell.set(q(OFFICE, "value-type"), "float")
            cell.set(q(OFFICE, "value"), str(value))
            spec = dict(op)
            spec["text"] = str(value)
        else:
            cell.set(q(OFFICE, "value-type"), "string")
            if q(OFFICE, "value") in cell.attrib:
                del cell.attrib[q(OFFICE, "value")]
            spec = dict(op)
            spec["text"] = "" if value is None else str(value)
        cell.append(_odf_paragraph(spec, auto, counter))
        _ods_widen(table, auto, counter, col_n, spec["text"])
    parts["content.xml"] = _xml(root)


def _ods_widen(table: ET.Element, auto: ET.Element, counter: list[int], col_n: int, text: str) -> None:
    columns = [child for child in list(table) if child.tag == q(TABLE, "table-column")]
    while len(columns) < col_n:
        fresh = ET.Element(q(TABLE, "table-column"))
        first_row = next((child for child in list(table) if child.tag == q(TABLE, "table-row")), None)
        if first_row is None:
            table.append(fresh)
        else:
            _insert_before(table, first_row, fresh)
        columns.append(fresh)
    needed = _cm_for_text(text)
    column = columns[col_n - 1]
    current = column.get(q(TABLE, "style-name"))
    if current:
        style = next((node for node in list(auto) if node.get(q(STYLE, "name")) == current), None)
        props = None if style is None else style.find(q(STYLE, "table-column-properties"))
        old = "" if props is None else str(props.get(q(STYLE, "column-width")) or "")
        try:
            if old.endswith("cm") and float(old[:-2]) >= float(needed[:-2]):
                return
        except ValueError:
            pass
    column.set(q(TABLE, "style-name"), _odf_column_style(auto, counter, needed))


def _read_odp(parts: dict[str, bytes]) -> dict:
    root, auto, _styles = _odf_auto_and_styles(parts)
    host = root.find(q(OFFICE, "body")).find(q(OFFICE, "presentation"))
    slides = []
    if host is None:
        return {"slides": []}
    for index, page in enumerate(host.findall(q(DRAW, "page")), start=1):
        paragraphs = [node for node in page.iter() if node.tag in {q(TEXT, "p"), q(TEXT, "h")}]
        props = _style_props(auto, None, paragraphs[0].get(q(TEXT, "style-name"))) if paragraphs else {"font": None, "size_pt": None, "bold": False}
        slides.append({
            "n": index,
            "title": "".join(paragraphs[0].itertext()) if paragraphs else "",
            "body": "\n".join("".join(node.itertext()) for node in paragraphs[1:]),
            "font": props.get("font"),
            "size_pt": props.get("size_pt"),
            "bold": bool(props.get("bold")),
        })
    return {"slides": slides}


def _edit_odp(parts: dict[str, bytes], ops: list) -> None:
    root, auto, _styles = _odf_auto_and_styles(parts)
    host = root.find(q(OFFICE, "body")).find(q(OFFICE, "presentation"))
    counter = [len(list(auto))]
    for op in ops:
        name = str(op.get("op") or "")
        pages = host.findall(q(DRAW, "page"))
        if name == "set_slide":
            number = int(op.get("number") or op.get("index") or 0)
            if number < 1 or number > len(pages):
                raise OfficeError(f"Slide {number} is outside 1..{len(pages)}.")
            page = pages[number - 1]
            paragraphs = [node for node in page.iter() if node.tag in {q(TEXT, "p"), q(TEXT, "h")}]
            old_title = "".join(paragraphs[0].itertext()) if paragraphs else ""
            old_body = "\n".join("".join(node.itertext()) for node in paragraphs[1:])
            for child in list(page):
                page.remove(child)
            frame = ET.SubElement(page, q(DRAW, "frame"))
            box = ET.SubElement(frame, q(DRAW, "text-box"))
            box.append(_odf_paragraph({"text": op.get("title", old_title), "style": "Title", "font": op.get("font"), "size_pt": op.get("size_pt") or 32}, auto, counter))
            box.append(_odf_paragraph({"text": op.get("body", old_body), "font": op.get("font"), "size_pt": op.get("body_pt") or 18, "bold": op.get("bold")}, auto, counter))
        elif name == "add_slide":
            page = ET.SubElement(host, q(DRAW, "page"), {q(DRAW, "name"): f"page{len(pages) + 1}"})
            frame = ET.SubElement(page, q(DRAW, "frame"))
            box = ET.SubElement(frame, q(DRAW, "text-box"))
            box.append(_odf_paragraph({"text": op.get("title") or "", "style": "Title", "font": op.get("font"), "size_pt": op.get("size_pt") or 32}, auto, counter))
            box.append(_odf_paragraph({"text": op.get("body") or "", "font": op.get("font"), "size_pt": op.get("body_pt") or 18, "bold": op.get("bold")}, auto, counter))
        else:
            raise OfficeError(f"{name or 'missing op'} is not a slide edit. Use set_slide or add_slide.")
    parts["content.xml"] = _xml(root)


def _edit_text(suffix: str, parts: dict[str, bytes], ops: list) -> None:
    if suffix == ".docx":
        _edit_docx(parts, ops)
    else:
        _edit_odt(parts, ops)


def _edit_sheet(suffix: str, parts: dict[str, bytes], ops: list) -> None:
    if suffix == ".xlsx":
        _edit_xlsx(parts, ops)
    else:
        _edit_ods(parts, ops)


def _edit_show(suffix: str, parts: dict[str, bytes], ops: list) -> None:
    if suffix == ".pptx":
        _edit_pptx(parts, ops)
    else:
        _edit_odp(parts, ops)
