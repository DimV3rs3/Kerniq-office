# Changelog

## 0.1.2

- Excel edits no longer rebuild `xl/styles.xml`. A cell change keeps the other sheets, their formulas, and the existing stylesheet. A new font is appended only when the edit asks for one.
- Cell values can be text, numbers, booleans, empty, or a formula (`=B1+1`). Formulas are stored as formulas, not as plain text.
- `office_edit` accepts `set_cells` and `append_row` on `.xlsx` and `.ods`, in addition to `set_sheet_cell`.
- Reading a workbook returns formula text (with the leading `=`) and real booleans.

## 0.1.1

- First public Office plugin: create, read, inspect, and edit Word, Excel, PowerPoint, and LibreOffice documents.
