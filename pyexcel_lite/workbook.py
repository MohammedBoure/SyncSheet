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
    formula_cells: set[tuple[int, int]] = field(default_factory=set)
    column_widths: dict[int, int] = field(default_factory=dict)
    row_heights: dict[int, int] = field(default_factory=dict)
    revision: int = 0

    def bump_revision(self) -> None:
        self.revision += 1

    def track_formula_cell(self, row: int, column: int, value: Any) -> None:
        key = (row, column)
        if isinstance(value, str) and value.startswith("="):
            self.formula_cells.add(key)
        else:
            self.formula_cells.discard(key)

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

    def set_value(self, row: int, column: int, value: Any, *, touch: bool = True) -> bool:
        self.ensure_size(row, column)
        next_value = "" if value is None else value
        existing = self.cells.get((row, column))
        if existing is None and next_value == "":
            return False
        cell = existing or self.get_cell(row, column)
        if cell.value == next_value:
            return False
        cell.value = next_value
        if cell.is_empty and cell.style == CellStyle():
            self.cells.pop((row, column), None)
        self.track_formula_cell(row, column, next_value)
        if touch:
            self.bump_revision()
        return True

    def set_style(self, row: int, column: int, **changes: Any) -> None:
        self.ensure_size(row, column)
        cell = self.get_cell(row, column)
        cell.style = cell.style.updated(**changes)

    def clear(self, rows: list[int], columns: list[int]) -> None:
        changed = False
        for row in rows:
            for column in columns:
                cell = self.cells.get((row, column))
                if cell is not None:
                    changed = changed or cell.value != ""
                    cell.value = ""
                    self.formula_cells.discard((row, column))
        if changed:
            self.bump_revision()

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
        self.formula_cells = {(row + count if row >= start else row, column) for row, column in self.formula_cells}
        self.row_count += count
        self.bump_revision()

    def remove_rows(self, start: int, count: int = 1) -> None:
        moved: dict[tuple[int, int], CellData] = {}
        end = start + count
        for (row, column), cell in self.cells.items():
            if start <= row < end:
                continue
            target_row = row - count if row >= end else row
            moved[(target_row, column)] = cell
        self.cells = moved
        self.formula_cells = {
            (row - count if row >= end else row, column)
            for row, column in self.formula_cells
            if not start <= row < end
        }
        self.row_count = max(1, self.row_count - count)
        self.bump_revision()

    def insert_columns(self, start: int, count: int = 1) -> None:
        moved: dict[tuple[int, int], CellData] = {}
        for (row, column), cell in self.cells.items():
            target_column = column + count if column >= start else column
            moved[(row, target_column)] = cell
        self.cells = moved
        self.formula_cells = {(row, column + count if column >= start else column) for row, column in self.formula_cells}
        self.column_count += count
        self.bump_revision()

    def remove_columns(self, start: int, count: int = 1) -> None:
        moved: dict[tuple[int, int], CellData] = {}
        end = start + count
        for (row, column), cell in self.cells.items():
            if start <= column < end:
                continue
            target_column = column - count if column >= end else column
            moved[(row, target_column)] = cell
        self.cells = moved
        self.formula_cells = {
            (row, column - count if column >= end else column)
            for row, column in self.formula_cells
            if not start <= column < end
        }
        self.column_count = max(1, self.column_count - count)
        self.bump_revision()


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
