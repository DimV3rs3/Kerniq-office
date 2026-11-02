"""Writing a spreadsheet must not rebuild the workbook around the edited cell."""

from __future__ import annotations

import importlib.util
import zipfile
from pathlib import Path


def _engine():
    path = Path(__file__).resolve().parents[1] / "engine.py"
    spec = importlib.util.spec_from_file_location("kerniq_office_engine_write", path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def _cell(doc, sheet, row, col):
    for item in doc["sheets"]:
        if item["name"] != sheet:
            continue
        for cell in item["cells"]:
            if cell["row"] == row and cell["col"] == col:
                return cell["value"]
    return None


def test_xlsx_keeps_the_other_sheet_and_its_styles(tmp_path):
    engine = _engine()
    book = tmp_path / "book.xlsx"
    engine.create_document(book, {"sheets": [
        {"name": "Data", "cells": [
            {"row": 1, "col": 1, "value": "Name"},
            {"row": 1, "col": 2, "value": 5},
            {"row": 1, "col": 3, "value": True},
            {"row": 2, "col": 2, "value": "=B1+1"},
        ]},
        {"name": "Other", "cells": [[1, 1, "keep"]]},
    ]})
    with zipfile.ZipFile(book) as archive:
        before_sheet = archive.read("xl/worksheets/sheet2.xml")
        before_styles = archive.read("xl/styles.xml")
    edited = engine.edit_document(book, [
        {"op": "set_sheet_cell", "sheet": "Data", "row": 3, "col": 1, "value": "more"},
        {"op": "append_row", "sheet": "Data", "values": ["tail", 9]},
    ])
    assert _cell(edited, "Data", 1, 1) == "Name"
    assert _cell(edited, "Data", 1, 2) == 5
    assert _cell(edited, "Data", 1, 3) is True
    assert _cell(edited, "Data", 2, 2) == "=B1+1"
    assert _cell(edited, "Data", 3, 1) == "more"
    assert _cell(edited, "Data", 4, 1) == "tail"
    assert _cell(edited, "Data", 4, 2) == 9
    assert _cell(edited, "Other", 1, 1) == "keep"
    with zipfile.ZipFile(book) as archive:
        assert archive.read("xl/worksheets/sheet2.xml") == before_sheet
        assert archive.read("xl/styles.xml") == before_styles


def test_ods_appends_a_row_without_dropping_the_first(tmp_path):
    engine = _engine()
    book = tmp_path / "book.ods"
    engine.create_document(book, {"sheets": [{"name": "Data", "cells": [[1, 1, "q"], [1, 2, 3]]}]})
    edited = engine.edit_document(book, [{"op": "append_row", "sheet": "Data", "values": ["next", "=B1+1"]}])
    assert _cell(edited, "Data", 1, 1) == "q"
    assert _cell(edited, "Data", 1, 2) == 3
    assert _cell(edited, "Data", 2, 1) == "next"
    assert str(_cell(edited, "Data", 2, 2)).startswith("=")
