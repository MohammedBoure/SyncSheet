# PyExcel Lite

PyExcel Lite is a PySide6 desktop spreadsheet built as a practical test for the provided virtual environment.

## Features

- Excel-style grid with row and column headers.
- Formula bar with live display and raw formula editing.
- Basic formulas: `SUM`, `AVERAGE`, `MIN`, `MAX`, `COUNT`, `COUNTA`, `IF`, `ROUND`, `ABS`, `SQRT`, `POWER`, `CONCAT`, `LEN`, `LEFT`, `RIGHT`, `UPPER`, `LOWER`, `AND`, `OR`, `NOT`, `TODAY`, and `NOW`.
- Cell references and ranges such as `=A1+B1` and `=SUM(A1:A10)`.
- Multiple worksheets with add, rename, and delete actions.
- Formatting toolbar for font, size, bold, italic, underline, alignment, number format, text color, and fill color.
- XLSX open/save using `openpyxl`.
- CSV export for the active sheet.
- Clipboard copy, cut, paste, delete, and quick selection statistics.

## Run

Use the venv requested for this build:

```powershell
D:\git\GoldShop\venv\Scripts\python.exe run_excel.py
```

## Test

```powershell
D:\git\GoldShop\venv\Scripts\python.exe -m unittest discover -s tests
```
