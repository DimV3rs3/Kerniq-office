# Office

Kerniq plugin. This folder is the whole plugin. It is not part of the Kerniq install. Publish this folder as its own repository. Kerniq loads it only after it is installed into `plugins/`.

The model gets `office_inspect`, `office_read`, `office_create`, and `office_edit`. `create_file`, `write_file`, `edit_file`, and `read_file` refuse these suffixes, because a document is not plain text.

Supported files:

| Suffix | Opens in |
|--------|----------|
| `.docx`, `.xlsx`, `.pptx` | Microsoft Office and LibreOffice |
| `.odt`, `.ods`, `.odp` | LibreOffice, and current Microsoft Office |

`office_inspect` is an outline: how many words, how many paragraphs (text outside tables), how many tables, and the shape of each table (rows by columns, fonts, point sizes), plus styles, sheet sizes, and slide titles. It does not include cell text, so the model can see the document before it reads it.

`office_read` returns the whole document as markdown. It does not stop at a character count, a row count, or a limit. A limit passed by the model is ignored. A table is a full grid. The header row is the column numbers, and the first cell of each body row is the row number. Those are the same numbers `office_edit` uses. A font that lives on the paragraph style is reported too, not only a font set on the run. Excel shared strings are read as text.

A Word or LibreOffice text table is as wide as the text area, and its columns share that width, so cell text wraps instead of the table being squeezed off the page. Spreadsheet columns are widened to the text in them, so Excel and LibreOffice do not clip the values. A new presentation includes a slide master, a layout, and a theme, which PowerPoint and LibreOffice expect.

`office_create` accepts the shapes a model actually sends. A sheet cell may be `{row, col, value}` or `[row, col, value]`, and a sheet may also be a grid of rows or a map whose keys are `A1`, `1,1`, or a column number. A string after a sheet object is that sheet's name, and so is an object that only has a name. Cells that are only broken fragments are refused, so a bad call does not leave a file behind. The same `[row, col, text]` form works for a text table. A slide may be `{title, body}` or just the title.

`office_edit` changes a document that already exists. Rows, columns, tables, paragraphs, and slides all start at 1.

Legacy `.doc`, `.xls`, and `.ppt` binaries are refused. Save them as OOXML or ODF first.
