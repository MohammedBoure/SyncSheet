# PyExcel Lite

PyExcel Lite is a PySide6 desktop spreadsheet built as a practical test for the provided virtual environment.

## Features

- Excel-style grid with row and column headers.
- Excel-like top ribbon with modern icons for file, sheet, edit, formatting, and zoom commands.
- Formula bar with live display and raw formula editing.
- Basic formulas: `SUM`, `AVERAGE`, `MIN`, `MAX`, `COUNT`, `COUNTA`, `IF`, `ROUND`, `ABS`, `SQRT`, `POWER`, `CONCAT`, `LEN`, `LEFT`, `RIGHT`, `UPPER`, `LOWER`, `AND`, `OR`, `NOT`, `TODAY`, and `NOW`.
- Expanded Excel-like formulas: criteria functions such as `SUMIF`, lookup functions such as `VLOOKUP`, statistics, trigonometry, date functions, text functions, `IFERROR`, percentage expressions such as `50%`, and `&` text concatenation.
- Formula separators can be commas or semicolons, for example `=IF(A1=10; "ok"; "no")`.
- Cell references and ranges such as `=A1+B1` and `=SUM(A1:A10)`.
- Zoom in, zoom out, reset zoom, `Ctrl` + mouse wheel zooming, and shortcuts `Ctrl++`, `Ctrl+-`, and `Ctrl+0`.
- Multiple worksheets with add, rename, and delete actions.
- Formatting toolbar for font, size, bold, italic, underline, alignment, number format, text color, and fill color.
- XLSX open/save using `openpyxl`.
- CSV export for the active sheet.
- Clipboard copy, cut, paste, delete, and quick selection statistics.

## Formula Coverage

- Math: `SUM`, `PRODUCT`, `ROUND`, `ROUNDUP`, `ROUNDDOWN`, `INT`, `TRUNC`, `CEILING`, `FLOOR`, `MOD`, `ABS`, `SQRT`, `POWER`, `EXP`, `LN`, `LOG`, `LOG10`, `PI`, `RAND`, `RANDBETWEEN`.
- Statistics: `AVERAGE`, `MIN`, `MAX`, `MEDIAN`, `MODE`, `COUNT`, `COUNTA`, `STDEV.S`, `STDEV.P`, `VAR.S`, `VAR.P`.
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

## Test

```powershell
D:\git\GoldShop\venv\Scripts\python.exe -m unittest discover -s tests
```
