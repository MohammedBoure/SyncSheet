"""In-memory workbook data structures."""

from __future__ import annotations

from dataclasses import dataclass, field, replace
from typing import Any


@dataclass
class CellStyle:
    font_family: str = "Segoe UI"
    font_size: int = 10
    bold: bool = False
    italic: bool = False
    underline: bool = False
    text_color: str = "#111827"
    fill_color: str = "#ffffff"
    horizontal: str = "general"
    vertical: str = "center"
    number_format: str = "General"
    wrap_text: bool = False

    def updated(self, **changes: Any) -> "CellStyle":
        return replace(self, **changes)


@dataclass
class CellData:
    value: Any = ""
    style: CellStyle = field(default_factory=CellStyle)

    @property
    def is_empty(self) -> bool:
        return self.value in ("", None)


@dataclass
class WorksheetData:
    name: str
    row_count: int = 200
    column_count: int = 52
    cells: dict[tuple[int, int], CellData] = field(default_factory=dict)
    column_widths: dict[int, int] = field(default_factory=dict)
    row_heights: dict[int, int] = field(default_factory=dict)

    def ensure_size(self, row: int, column: int) -> None:
        if row >= self.row_count:
            self.row_count = row + 1
        if column >= self.column_count:
            self.column_count = column + 1

    def get_cell(self, row: int, column: int) -> CellData:
        self.ensure_size(row, column)
        key = (row, column)
        if key not in self.cells:
            self.cells[key] = CellData()
        return self.cells[key]

    def raw_value(self, row: int, column: int) -> Any:
        if row < 0 or column < 0:
            return ""
        cell = self.cells.get((row, column))
        return "" if cell is None else cell.value

    def set_value(self, row: int, column: int, value: Any) -> None:
        self.ensure_size(row, column)
        cell = self.get_cell(row, column)
        cell.value = "" if value is None else value
        if cell.is_empty and cell.style == CellStyle():
            self.cells.pop((row, column), None)

    def set_style(self, row: int, column: int, **changes: Any) -> None:
        self.ensure_size(row, column)
        cell = self.get_cell(row, column)
        cell.style = cell.style.updated(**changes)

    def clear(self, rows: list[int], columns: list[int]) -> None:
        for row in rows:
            for column in columns:
                cell = self.cells.get((row, column))
                if cell is not None:
                    cell.value = ""

    def iter_used_cells(self):
        for (row, column), cell in sorted(self.cells.items()):
            if not cell.is_empty or cell.style != CellStyle():
                yield row, column, cell

    def insert_rows(self, start: int, count: int = 1) -> None:
        moved: dict[tuple[int, int], CellData] = {}
        for (row, column), cell in self.cells.items():
            target_row = row + count if row >= start else row
            moved[(target_row, column)] = cell
        self.cells = moved
        self.row_count += count

    def remove_rows(self, start: int, count: int = 1) -> None:
        moved: dict[tuple[int, int], CellData] = {}
        end = start + count
        for (row, column), cell in self.cells.items():
            if start <= row < end:
                continue
            target_row = row - count if row >= end else row
            moved[(target_row, column)] = cell
        self.cells = moved
        self.row_count = max(1, self.row_count - count)

    def insert_columns(self, start: int, count: int = 1) -> None:
        moved: dict[tuple[int, int], CellData] = {}
        for (row, column), cell in self.cells.items():
            target_column = column + count if column >= start else column
            moved[(row, target_column)] = cell
        self.cells = moved
        self.column_count += count

    def remove_columns(self, start: int, count: int = 1) -> None:
        moved: dict[tuple[int, int], CellData] = {}
        end = start + count
        for (row, column), cell in self.cells.items():
            if start <= column < end:
                continue
            target_column = column - count if column >= end else column
            moved[(row, target_column)] = cell
        self.cells = moved
        self.column_count = max(1, self.column_count - count)


@dataclass
class WorkbookData:
    sheets: list[WorksheetData] = field(default_factory=list)
    active_sheet_index: int = 0
    path: str | None = None

    def __post_init__(self) -> None:
        if not self.sheets:
            self.sheets.append(WorksheetData(name="Sheet1"))

    @property
    def active_sheet(self) -> WorksheetData:
        return self.sheets[self.active_sheet_index]

    def sheet_names(self) -> list[str]:
        return [sheet.name for sheet in self.sheets]

    def unique_sheet_name(self, base: str = "Sheet") -> str:
        names = set(self.sheet_names())
        if base not in names:
            return base
        index = 2
        while f"{base}{index}" in names:
            index += 1
        return f"{base}{index}"

    def add_sheet(self, name: str | None = None) -> WorksheetData:
        sheet = WorksheetData(name=name or self.unique_sheet_name())
        self.sheets.append(sheet)
        self.active_sheet_index = len(self.sheets) - 1
        return sheet

    def remove_sheet(self, index: int) -> None:
        if len(self.sheets) == 1:
            return
        self.sheets.pop(index)
        self.active_sheet_index = min(self.active_sheet_index, len(self.sheets) - 1)

    def sheet_by_name(self, name: str) -> WorksheetData | None:
        lowered = name.lower()
        for sheet in self.sheets:
            if sheet.name.lower() == lowered:
                return sheet
        return None
