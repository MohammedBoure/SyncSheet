# PyExcel Lite

PyExcel Lite is a PySide6 desktop spreadsheet built as a practical test for the provided virtual environment.

## Features

- Excel-style grid with row and column headers.
- Excel-like tabbed ribbon with Home, Project, Insert, Formulas, Data, View, and Network command pages.
- Slim Excel-like formula bar under the ribbon, leaving more room for the worksheet.
- Right task pane with icon-led sections for selection, stats, project files, network collaboration, formulas, algorithms, and charts.
- Organized UI theme rules in `pyexcel_lite/theme.py` so interface styling is separated from spreadsheet logic.
- Project workspace support for opening a whole folder tree as one project, including nested folders and multiple workbook or CSV files.
- The last opened project is remembered and reopened automatically on the next launch.
- Large-table performance improvements with batched paste/clear operations, partial redraws, and cached formula results.
- Undo and redo for cell edits, paste operations, clears, and generated formulas.
- Bar, line, and pie charts from the selected cells, with full-screen dialog viewing.
- Realtime LAN collaboration with Host, Join, and Leave controls so multiple users can edit the same workbook together.
- Project-level collaboration keeps each project workbook separate, so users can work in different files while edits sync in the background.
- Network startup settings that can launch the app in manual mode, connect directly to a shared server, or make the current program host a server automatically.
- Join and Host dialogs can save automatic network startup, so the next launch connects or hosts immediately.
- Standalone shared-workbook and project server for LAN client/server use, with JSON state persistence.
- TCP socket sync for workbook snapshots, cell edits, paste/clear batches, sheet add/rename/delete, and row/column insert/delete operations.
- Formula bar with live display and raw formula editing.
- Right-click context menu for selected cells with copy, cut, paste, clear, row/column, formula, and chart operations.
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

The host sends the current workbook and any synchronized project workbooks as snapshots when a user joins, then all users receive live TCP socket updates for cell edits, pasted ranges, sheet changes, and row or column structure changes.
Connected sockets use TCP keepalive, blocking reads after the initial connect timeout, app-level ping/pong heartbeats, and automatic client retry when the server connection drops.

Use `Network > Startup` to choose what happens when the program opens:

- `Manual network controls`: start offline and use Host or Join manually.
- `Connect to shared server`: connect to a fixed LAN IP and port immediately.
- `Run this program as server`: start hosting immediately from this program.

You can also enable automatic startup directly from `Network > Join` or `Network > Host` by checking the automatic option in that dialog. The program saves the selected IP and port, then uses them on the next launch without asking again.

For a dedicated shared server on a LAN computer:

```powershell
D:\git\GoldShop\venv\Scripts\python.exe run_collaboration_server.py --host 0.0.0.0 --port 8765
```

Then set each client to `Network > Startup > Connect to shared server` and enter the server IP and port.
The standalone server keeps the shared workbook in `shared_workbook_state.json` by default, so the latest central state can survive a server restart.

## Project Workspaces

Use `Project > Open Project` to choose a parent folder. The program scans nested folders as one workspace and shows the project tree in the right Project task pane.
The last opened project folder is saved in `pyexcel_lite_settings.json`; if it still exists, PyExcel Lite reopens it automatically the next time the program starts.
XLSX and CSV files can be opened directly from the project tree; opening a project spreadsheet sends a full workbook snapshot to connected teammates.

Use `Project > Share Project` while connected to a host or shared server to synchronize the project structure with the team.
Each openable project file is tracked by its relative path. If one user edits `reports/budget.xlsx` while another user is viewing a different file, the update is cached in the background and appears when that second user opens `reports/budget.xlsx`.
If a client opens a project spreadsheet that is not cached locally yet, it requests that workbook from the host or shared server and opens it when the snapshot arrives.
The shared server stores the active workbook, synchronized project workbooks, and the latest project snapshot, so new clients receive the same workspace context when they join.

