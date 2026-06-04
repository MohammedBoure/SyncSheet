"""Helpers for Excel-style cell addresses."""

from __future__ import annotations

import re
from dataclasses import dataclass

CELL_RE = re.compile(r"^\$?([A-Za-z]{1,4})\$?([1-9][0-9]*)$")


@dataclass(frozen=True)
class CellRef:
    row: int
    column: int

    @property
    def address(self) -> str:
        return index_to_column_name(self.column) + str(self.row + 1)


def index_to_column_name(index: int) -> str:
    if index < 0:
        raise ValueError("Column index must be positive")
    result = []
    value = index + 1
    while value:
        value, remainder = divmod(value - 1, 26)
        result.append(chr(65 + remainder))
    return "".join(reversed(result))


def column_name_to_index(name: str) -> int:
    if not name or not name.isalpha():
        raise ValueError(f"Invalid column name: {name!r}")
    value = 0
    for char in name.upper():
        value = value * 26 + (ord(char) - 64)
    return value - 1


def parse_cell_ref(address: str) -> CellRef:
    match = CELL_RE.match(address.strip())
    if not match:
        raise ValueError(f"Invalid cell reference: {address!r}")
    column_name, row_text = match.groups()
    return CellRef(row=int(row_text) - 1, column=column_name_to_index(column_name))


def normalize_cell_ref(address: str) -> str:
    ref = parse_cell_ref(address.replace("$", ""))
    return ref.address


def iter_range(start: str, end: str):
    first = parse_cell_ref(start)
    last = parse_cell_ref(end)
    row_from, row_to = sorted((first.row, last.row))
    col_from, col_to = sorted((first.column, last.column))
    for row in range(row_from, row_to + 1):
        for column in range(col_from, col_to + 1):
            yield CellRef(row=row, column=column)
