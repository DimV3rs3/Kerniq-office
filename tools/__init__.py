"""Office commands. Offered to the model because the manifest says so."""

from __future__ import annotations

import importlib.util
import sys
from pathlib import Path
from typing import Any, Dict

from Kerniq.interfaces.tools import Tool, ToolResult


def _engine():
    name = "kerniq_office_engine"
    cached = sys.modules.get(name)
    if cached is not None:
        return cached
    path = Path(__file__).resolve().parents[1] / "engine.py"
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        raise RuntimeError("Office engine is missing")
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


def _workspace(context: Any) -> Path:
    """Real workspace directory. Never rewrite it through the virtual-path helper."""
    from Kerniq.security.roots import model_roots
    from Kerniq.server.homes import jail_root

    if jail_root() is not None:
        return model_roots(".").workspace
    raw = getattr(context, "workspace", None) or ""
    if not raw:
        meta = getattr(context, "metadata", None)
        if isinstance(meta, dict):
            raw = meta.get("workspace") or ""
    if raw:
        try:
            return model_roots(raw).workspace
        except Exception:
            path = Path(str(raw)).expanduser()
            if path.is_absolute():
                return path.resolve()
    return model_roots(Path.cwd() / "workspace").workspace


def _safe(path_str: str, context: Any, *, must_exist: bool) -> Path:
    from Kerniq.server.homes import jail_root, logical_path
    from Kerniq.security.roots import model_roots, resolve_virtual

    raw = str(path_str or "").strip()
    if not raw or "\x00" in raw:
        raise PermissionError("A path inside the workspace is required.")
    base = _workspace(context)
    try:
        virtual = resolve_virtual(raw, base)
    except PermissionError as exc:
        raise PermissionError("That path is outside the workspace.") from exc
    if virtual is not None:
        target = virtual
    else:
        jail = jail_root()
        if jail is not None:
            target = logical_path(raw, jail)
        else:
            candidate = Path(raw)
            target = candidate.resolve() if candidate.is_absolute() else (base / candidate).resolve()
            roots = model_roots(base)
            allowed = False
            for folder in roots.as_list():
                try:
                    target.relative_to(folder.resolve())
                    allowed = True
                    break
                except ValueError:
                    continue
            if not allowed:
                raise PermissionError("That path is outside the workspace.")
    if must_exist and not target.is_file():
        raise FileNotFoundError(f"File not found: {raw}")
    return target


def _virtual_path_for_return(path: Path, context: Any) -> str:
    """Return the path the model asked for (virtual form) so UI and logs stay consistent."""
    from Kerniq.security.pathmask import rewrite_model_path
    try:
        from Kerniq.security.roots import shown_virtual
        v = shown_virtual(path, getattr(context, "workspace", None) or ".")
        if v:
            return v
    except Exception:
        pass
    return rewrite_model_path(str(path)) or str(path)


class _OfficeTool(Tool):
    def __init__(self, plugin) -> None:
        self._plugin = plugin

    @property
    def risk_level(self) -> str:
        return "medium"

    @property
    def category(self) -> str:
        return "office"

    def settings_schema(self):
        from Kerniq.interfaces.settings_schema import SettingsSchema

        return SettingsSchema(
            component_id=self.tool_id,
            component_type="tool",
            title=self.name,
            fields=[],
        )

    def _run(self, call_id: str, fn):
        try:
            output = fn()
        except Exception as exc:
            return ToolResult(
                tool_call_id=call_id,
                tool_id=self.tool_id,
                status="error",
                success=False,
                error=str(exc),
                human_readable=str(exc),
            )
        return ToolResult(
            tool_call_id=call_id,
            tool_id=self.tool_id,
            status="success",
            output=output,
            human_readable=_readable(output),
        )


def _readable(output: dict) -> str:
    if not isinstance(output, dict):
        return str(output)
    if output.get("inspected"):
        return _inspect_markdown(output)
    chunks = [f"# {output.get('path') or 'document'}\n\n`{output.get('kind') or 'document'}`"]
    paragraphs = output.get("paragraphs") or []
    if paragraphs:
        chunks.append("## Paragraphs")
        for paragraph in paragraphs:
            chunks.append(_paragraph_line(paragraph))
    for table in output.get("tables") or []:
        rows = table.get("cells") or []
        total = int(table.get("rows_total") or table.get("rows") or len(rows))
        title = f"## Table {table.get('n')} · {total}×{table.get('cols')}"
        meta = []
        if table.get("fonts"):
            meta.append("fonts " + ", ".join(str(item) for item in table["fonts"]))
        if table.get("sizes_pt"):
            meta.append("sizes " + ", ".join(_point(item) for item in table["sizes_pt"]))
        if meta:
            title += "\n\n*" + " · ".join(meta) + "*"
        chunks.append(title)
        chunks.extend(_grid_markdown(rows))
    for sheet in output.get("sheets") or []:
        lines, notes = _sheet_markdown(sheet.get("cells") or [])
        title = f"## Sheet {sheet.get('name')} · {sheet.get('rows')}×{sheet.get('cols')}"
        if sheet.get("fonts"):
            title += "\n\n*fonts " + ", ".join(str(item) for item in sheet["fonts"]) + "*"
        chunks.append(title)
        chunks.extend(lines)
        if notes:
            chunks.append("*Formatting: " + "; ".join(notes) + ".*")
    for slide in output.get("slides") or []:
        font = []
        if slide.get("font"):
            font.append(str(slide["font"]))
        if slide.get("size_pt") is not None:
            font.append(_point(slide["size_pt"]))
        block = f"## Slide {slide.get('n')}\n\n**{_md_cell(slide.get('title') or '')}**\n\n{_md_cell(slide.get('body') or '')}"
        if font:
            block += "\n\n*" + " ".join(font) + "*"
        chunks.append(block)
    return "\n".join(chunks)


def _inspect_markdown(output: dict) -> str:
    lines = [
        f"# {output.get('path') or 'document'}",
        "",
        f"`{output.get('kind') or 'document'}`",
        "",
        f"- Words: {int(output.get('words') or 0)}",
        f"- Paragraphs: {int(output.get('paragraphs') or 0)} (text outside tables)",
        f"- Tables: {int(output.get('tables') or 0)}",
    ]
    styles = output.get("styles") or []
    if styles:
        lines.append("- Styles: " + ", ".join(str(item) for item in styles))
    fonts = output.get("fonts") or []
    if fonts:
        lines.append("- Fonts: " + ", ".join(str(item) for item in fonts))
    for table in output.get("table_list") or []:
        bits = [f"Table {table.get('n')}: {table.get('rows')}×{table.get('cols')}"]
        if table.get("fonts"):
            bits.append("fonts " + ", ".join(str(item) for item in table["fonts"]))
        if table.get("sizes_pt"):
            bits.append("sizes " + ", ".join(_point(item) for item in table["sizes_pt"]))
        lines.append("- " + " · ".join(bits))
    for sheet in output.get("sheets") or []:
        bits = [f"Sheet {sheet.get('name')}: {sheet.get('rows')}×{sheet.get('cols')}"]
        if sheet.get("fonts"):
            bits.append("fonts " + ", ".join(str(item) for item in sheet["fonts"]))
        if sheet.get("sizes_pt"):
            bits.append("sizes " + ", ".join(_point(item) for item in sheet["sizes_pt"]))
        lines.append("- " + " · ".join(bits))
    for slide in output.get("slides") or []:
        title = _md_cell(slide.get("title") or "")
        lines.append(f"- Slide {slide.get('n')}: {title}" if title else f"- Slide {slide.get('n')}")
    lines.append("")
    lines.append("This is an outline. office_read returns every paragraph and every cell, with nothing left out.")
    return "\n".join(lines)


def _md_cell(value: object) -> str:
    text = "" if value is None else str(value)
    return text.replace("\\", "\\\\").replace("|", "\\|").replace("\r\n", "\n").replace("\r", "\n").replace("\n", "<br>")


def _point(size: object) -> str:
    if isinstance(size, float):
        return f"{size:g}pt"
    return f"{size}pt"


def _paragraph_line(paragraph: dict) -> str:
    bits = [str(paragraph.get("style") or "Normal")]
    if paragraph.get("font"):
        bits.append(str(paragraph["font"]))
    if paragraph.get("size_pt") is not None:
        bits.append(_point(paragraph["size_pt"]))
    if paragraph.get("bold"):
        bits.append("bold")
    if paragraph.get("italic"):
        bits.append("italic")
    return f"{paragraph.get('n')}. {' · '.join(bits)} – {_md_cell(paragraph.get('text') or '')}"


def _grid_markdown(rows: list) -> list[str]:
    if not rows:
        return ["_empty_"]
    normalized = []
    width = 1
    for row in rows:
        values = list(row) if isinstance(row, (list, tuple)) else [row]
        normalized.append(values)
        width = max(width, len(values))
    header = ["row"] + [str(index) for index in range(1, width + 1)]
    lines = [
        "| " + " | ".join(header) + " |",
        "| " + " | ".join("---" for _ in header) + " |",
    ]
    for number, values in enumerate(normalized, start=1):
        cells = [_md_cell(values[index]) if index < len(values) else "" for index in range(width)]
        lines.append("| " + " | ".join([str(number), *cells]) + " |")
    return lines


def _sheet_markdown(cells: list) -> tuple[list[str], list[str]]:
    used: dict[int, dict[int, object]] = {}
    columns: set[int] = set()
    notes = []
    for cell in cells or []:
        if not isinstance(cell, dict):
            continue
        row, col = int(cell.get("row") or 0), int(cell.get("col") or 0)
        if row < 1 or col < 1:
            continue
        used.setdefault(row, {})[col] = cell.get("value")
        columns.add(col)
        marks = []
        if cell.get("bold"):
            marks.append("bold")
        if cell.get("italic"):
            marks.append("italic")
        font = cell.get("font")
        size = cell.get("size_pt")
        if font and font != "Calibri":
            marks.append(str(font))
        if size not in (None, 11, 11.0):
            marks.append(_point(size))
        if marks:
            notes.append(f"r{row}c{col} {', '.join(marks)}")
    ordered = sorted(columns)
    if not ordered:
        return ["_empty_"], notes
    header = ["row"] + [str(col) for col in ordered]
    lines = [
        "| " + " | ".join(header) + " |",
        "| " + " | ".join("---" for _ in header) + " |",
    ]
    for row in sorted(used):
        values = [_md_cell(used[row].get(col, "")) for col in ordered]
        lines.append("| " + " | ".join([str(row), *values]) + " |")
    return lines, notes


class OfficeInspectTool(_OfficeTool):
    @property
    def tool_id(self) -> str:
        return "office_inspect"

    @property
    def name(self) -> str:
        return "office_inspect"

    @property
    def display_phrase(self) -> str:
        return "documents inspected"

    @property
    def description(self) -> str:
        return (
            "Describe a Microsoft Office or LibreOffice document before reading or editing it. "
            "The reply is a short outline, not the document: word count, paragraph count, table count, "
            "and the shape of each table (rows by columns, fonts, sizes). "
            "Styles, sheet sizes, and slide titles are included. Cell text is not. "
            "Paragraphs are the paragraphs outside tables. "
            "Call office_read for the whole document. office_read does not shorten anything. "
            "Use this instead of read_file. Suffix: docx, xlsx, pptx, odt, ods, odp."
        )

    @property
    def input_schema(self) -> Dict[str, Any]:
        return {
            "type": "object",
            "properties": {"path": {"type": "string", "description": "Workspace path of the document."}},
            "required": ["path"],
        }

    async def execute(self, arguments: Dict[str, Any], context: Any) -> ToolResult:
        call_id = str(arguments.get("tool_call_id") or "")
        engine = _engine()

        def run():
            doc = engine.inspect_document(_safe(arguments.get("path"), context, must_exist=True))
            if isinstance(doc, dict):
                doc = dict(doc)
                doc["path"] = arguments.get("path") or doc.get("path")
            return doc

        return self._run(call_id, run)


class OfficeReadTool(_OfficeTool):
    @property
    def tool_id(self) -> str:
        return "office_read"

    @property
    def name(self) -> str:
        return "office_read"

    @property
    def display_phrase(self) -> str:
        return "documents read"

    @property
    def description(self) -> str:
        return (
            "Read a Word, Excel, PowerPoint, or LibreOffice document in full. The reply is markdown. "
            "Nothing is shortened: every paragraph, every table cell, and every sheet cell comes back. "
            "A table is a grid. Column numbers are the header, and the first cell of each row is the row number office_edit expects. "
            "Paragraph styles and fonts are listed with the paragraphs. Spreadsheet sheets use the same markdown table. "
            "Not plain text, so read_file will refuse it. path is inside the workspace. "
            "Do not pass a limit. A limit is ignored, because cutting the document is the wrong way to read it."
        )

    @property
    def input_schema(self) -> Dict[str, Any]:
        return {
            "type": "object",
            "properties": {
                "path": {"type": "string"},
            },
            "required": ["path"],
        }

    async def execute(self, arguments: Dict[str, Any], context: Any) -> ToolResult:
        call_id = str(arguments.get("tool_call_id") or "")
        engine = _engine()
        def run():
            target = _safe(arguments.get("path"), context, must_exist=True)
            doc = engine.read_document(target)
            if isinstance(doc, dict):
                doc = dict(doc)
                doc["path"] = arguments.get("path") or _virtual_path_for_return(target, context)
            return doc

        return self._run(call_id, run)


class OfficeCreateTool(_OfficeTool):
    @property
    def tool_id(self) -> str:
        return "office_create"

    @property
    def name(self) -> str:
        return "office_create"

    @property
    def display_phrase(self) -> str:
        return "documents created"

    @property
    def description(self) -> str:
        return (
            "Create a Microsoft Office or LibreOffice document. create_file cannot do this. "
            "The suffix chooses the package: docx and odt are documents, xlsx and ods are spreadsheets, pptx and odp are slides. "
            "Fails if the file already exists. paragraphs: text, style (Normal, Heading1, Heading2, Title), font, size_pt, bold, italic. "
            "tables: cells as a grid of rows, or as [row, col, text] lists. Rows and columns start at 1. "
            "sheets: one object per sheet, and the name is a field of that same object. "
            "cells is a list of [row, col, text]. Row and column start at 1. "
            "Example: {\"name\": \"Report\", \"cells\": [[1, 1, \"File\"], [1, 2, \"Indicator\"]]}. "
            "A cell may also be {\"row\": 1, \"col\": 1, \"value\": \"File\", \"bold\": true}. "
            "Do not put the sheet name in a second object. "
            "slides: {title, body, font, size_pt}, or a string that is the title."
        )

    @property
    def input_schema(self) -> Dict[str, Any]:
        return {
            "type": "object",
            "properties": {
                "path": {"type": "string"},
                "title": {"type": "string"},
                "paragraphs": {"type": "array"},
                "tables": {"type": "array", "description": "Each item has cells, a grid of strings or {text, font, size_pt, bold}."},
                "sheets": {"type": "array"},
                "slides": {"type": "array"},
            },
            "required": ["path"],
        }

    async def execute(self, arguments: Dict[str, Any], context: Any) -> ToolResult:
        call_id = str(arguments.get("tool_call_id") or "")
        engine = _engine()
        spec = {key: arguments.get(key) for key in ("title", "paragraphs", "tables", "sheets", "slides") if key in arguments}

        def run():
            doc = engine.create_document(_safe(arguments.get("path"), context, must_exist=False), spec)
            if isinstance(doc, dict):
                doc = dict(doc)
                doc["path"] = arguments.get("path") or doc.get("path")
            return doc

        return self._run(call_id, run)


class OfficeEditTool(_OfficeTool):
    @property
    def tool_id(self) -> str:
        return "office_edit"

    @property
    def name(self) -> str:
        return "office_edit"

    @property
    def display_phrase(self) -> str:
        return "documents edited"

    @property
    def description(self) -> str:
        return (
            "Edit an existing Office document in place. Do not use edit_file or write_file. "
            "ops is a list. Documents (docx, odt): set_paragraph {index, text, style, font, size_pt, bold, italic}, "
            "add_paragraph, set_style {index, style}, set_cell {table, row, col, text, font, size_pt, bold}, add_table {cells}. "
            "Spreadsheets (xlsx, ods): set_sheet_cell {sheet, row, col, value}, "
            "set_cells {sheet, cells}, or append_row {sheet, values}. "
            "A value that starts with = is a formula, not text. "
            "Booleans stay booleans. Other sheets and the workbook styles are left as they are. "
            "Slides (pptx, odp): set_slide {number, title, body, font, size_pt}, add_slide. "
            "Every index, table, row, and column starts at 1. Paragraph indexes skip text inside tables."
        )

    @property
    def input_schema(self) -> Dict[str, Any]:
        return {
            "type": "object",
            "properties": {
                "path": {"type": "string"},
                "ops": {"type": "array", "description": "Edit operations. See the tool description."},
                "op": {"type": "string", "description": "One operation name, if you are not passing ops."},
            },
            "required": ["path"],
        }

    async def execute(self, arguments: Dict[str, Any], context: Any) -> ToolResult:
        call_id = str(arguments.get("tool_call_id") or "")
        engine = _engine()
        ops = arguments.get("ops")
        if not ops and arguments.get("op"):
            ops = [{key: value for key, value in arguments.items() if key not in {"path", "tool_call_id"}}]

        def run():
            doc = engine.edit_document(_safe(arguments.get("path"), context, must_exist=True), list(ops or []))
            if isinstance(doc, dict):
                doc = dict(doc)
                doc["path"] = arguments.get("path") or doc.get("path")
            return doc

        return self._run(call_id, run)


def create_tools(plugin):
    return [
        OfficeInspectTool(plugin),
        OfficeReadTool(plugin),
        OfficeCreateTool(plugin),
        OfficeEditTool(plugin),
    ]
