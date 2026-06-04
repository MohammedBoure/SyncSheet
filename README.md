# PyExcel Lite

PyExcel Lite is a PySide6 desktop spreadsheet built as a practical test for the provided virtual environment.

## Features

- Excel-style grid with row and column headers.
- Excel-like tabbed ribbon with Home, Insert, Formulas, Data, View, and Network command pages.
- Professional top panel under the ribbon with icon-led sections for selection, stats, network collaboration, formulas, algorithms, and charts.
- Large-table performance improvements with batched paste/clear operations, partial redraws, and cached formula results.
- Undo and redo for cell edits, paste operations, clears, and generated formulas.
- Bar, line, and pie charts from the selected cells.
- Realtime LAN collaboration with Host, Join, and Leave controls so multiple users can edit the same workbook together.
- TCP socket sync for workbook snapshots, cell edits, paste/clear batches, sheet add/rename/delete, and row/column insert/delete operations.
- Formula bar with live display and raw formula editing.
- Formula library for applying arithmetic, statistical, logical, lookup, and custom cell algorithms to the active cell or selected range.
- Basic formulas: `SUM`, `AVERAGE`, `MIN`, `MAX`, `COUNT`, `COUNTA`, `IF`, `ROUND`, `ABS`, `SQRT`, `POWER`, `CONCAT`, `LEN`, `LEFT`, `RIGHT`, `UPPER`, `LOWER`, `AND`, `OR`, `NOT`, `TODAY`, and `NOW`.
- Expanded Excel-like formulas: criteria functions such as `SUMIF`, lookup functions such as `VLOOKUP`, statistics, trigonometry, date functions, text functions, `IFERROR`, percentage expressions such as `50%`, and `&` text concatenation.
- Formula separators can be commas or semicolons, for example `=IF(A1=10; "ok"; "no")`.
- Cell references and ranges such as `=A1+B1` and `=SUM(A1:A10)`.
- Cross-sheet references such as `=SUM(Sheet2!A1:A10)` and `=SUM('Data Sheet'!A1:A10)`.
- Zoom in, zoom out, reset zoom, `Ctrl` + mouse wheel zooming, and shortcuts `Ctrl++`, `Ctrl+-`, and `Ctrl+0`.
- Multiple worksheets with add, rename, and delete actions.
- Formatting toolbar for font, size, bold, italic, underline, alignment, number format, text color, and fill color.
- XLSX open/save using `openpyxl`.
- CSV export for the active sheet.
- Clipboard copy, cut, paste, delete, and quick selection statistics.

## Formula Coverage

- Math: `SUM`, `PRODUCT`, `ROUND`, `ROUNDUP`, `ROUNDDOWN`, `INT`, `TRUNC`, `CEILING`, `FLOOR`, `MOD`, `ABS`, `SQRT`, `POWER`, `EXP`, `LN`, `LOG`, `LOG10`, `PI`, `RAND`, `RANDBETWEEN`.
- Statistics: `AVERAGE`, `MIN`, `MAX`, `MEDIAN`, `MODE`, `COUNT`, `COUNTA`, `STDEV.S`, `STDEV.P`, `VAR.S`, `VAR.P`, `PERCENTILE.INC`, `QUARTILE.INC`, `RANK.EQ`, `CORREL`, `COVARIANCE.P`, `COVARIANCE.S`.
- Criteria: `SUMIF`, `SUMIFS`, `COUNTIF`, `AVERAGEIF`.
- Logic: `IF`, `IFS`, `IFERROR`, `AND`, `OR`, `NOT`, `TRUE`, `FALSE`.
- Text: `CONCAT`, `CONCATENATE`, `TEXTJOIN`, `LEN`, `LEFT`, `RIGHT`, `MID`, `UPPER`, `LOWER`, `TRIM`, `PROPER`, `FIND`, `SEARCH`, `SUBSTITUTE`, `REPLACE`, `REPT`.
- Date: `TODAY`, `NOW`, `DATE`, `DATEVALUE`, `YEAR`, `MONTH`, `DAY`, `DAYS`, `WEEKDAY`, `HOUR`, `MINUTE`, `SECOND`.
- Lookup: `INDEX`, `MATCH`, `VLOOKUP`, `HLOOKUP`, `CHOOSE`.

## Run

Use the venv requested for this build:

```powershell
D:\git\GoldShop\venv\Scripts\python.exe run_excel.py
```

## Realtime Collaboration

On the server computer, choose `Network > Host` and keep the selected port, usually `8765`.
The Network panel shows the local address clients should use, for example `192.168.1.20:8765`.
On each client computer on the same LAN, choose `Network > Join` and enter that server address.

The host sends the current workbook as a snapshot when a user joins, then all users receive live TCP socket updates for cell edits, pasted ranges, sheet changes, and row or column structure changes.
Connected sockets use TCP keepalive, blocking reads after the initial connect timeout, app-level ping/pong heartbeats, and automatic client retry when the server connection drops.

## Test

```powershell
D:\git\GoldShop\venv\Scripts\python.exe -m unittest discover -s tests
```
