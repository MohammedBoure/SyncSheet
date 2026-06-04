"""Qt table model for worksheet data."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Callable

from PySide6.QtCore import QAbstractTableModel, QModelIndex, QItemSelectionRange, Qt
from PySide6.QtGui import QBrush, QColor, QFont

from .cell_address import index_to_column_name
from .formula import FormulaEvaluator
from .workbook import CellStyle, WorksheetData

DEFAULT_STYLE = CellStyle()


@dataclass
class CellChangeCommand:
    model: "WorksheetTableModel"
    changes: list[tuple[int, int, object, object]]

    def undo(self) -> None:
        self.model.set_values(
            [(row, column, old_value) for row, column, old_value, _new_value in self.changes],
            refresh_dependents=True,
            record_undo=False,
        )

    def redo(self) -> None:
        self.model.set_values(
            [(row, column, new_value) for row, column, _old_value, new_value in self.changes],
            refresh_dependents=True,
            record_undo=False,
        )


class WorksheetTableModel(QAbstractTableModel):
    def __init__(self, sheet: WorksheetData, evaluator: FormulaEvaluator):
        super().__init__()
        self.sheet = sheet
        self.evaluator = evaluator
        self.zoom_factor = 1.0
        self.undo_stack: list[CellChangeCommand] = []
        self.redo_stack: list[CellChangeCommand] = []
        self.history_limit = 200
        self.history_changed: Callable[[], None] | None = None
        self.values_changed: Callable[[WorksheetData, list[tuple[int, int, object]]], None] | None = None

    def rowCount(self, parent: QModelIndex = QModelIndex()) -> int:
        return 0 if parent.isValid() else self.sheet.row_count

    def columnCount(self, parent: QModelIndex = QModelIndex()) -> int:
        return 0 if parent.isValid() else self.sheet.column_count

    def data(self, index: QModelIndex, role: int = Qt.DisplayRole):
        if not index.isValid():
            return None
        row = index.row()
        column = index.column()
        cell = self.sheet.cells.get((row, column))
        raw_value = "" if cell is None else cell.value
        style = DEFAULT_STYLE if cell is None else cell.style
        if role == Qt.DisplayRole:
            if raw_value == "":
                return ""
            return self.evaluator.display(raw_value, self.sheet, row, column)
        if role == Qt.EditRole:
            return raw_value
        if role == Qt.FontRole:
            if cell is None or style == DEFAULT_STYLE:
                return None
            return self._font_from_style(style)
        if role == Qt.ForegroundRole:
            if cell is None or style.text_color == DEFAULT_STYLE.text_color:
                return None
            return QBrush(QColor(style.text_color))
        if role == Qt.BackgroundRole:
            if cell is None or style.fill_color == DEFAULT_STYLE.fill_color:
                return None
            return QBrush(QColor(style.fill_color))
        if role == Qt.TextAlignmentRole:
            if cell is None or style.horizontal == DEFAULT_STYLE.horizontal:
                return None
            return self._alignment_from_style(style)
        return None

    def setData(
        self,
        index: QModelIndex,
        value,
        role: int = Qt.EditRole,
        refresh_dependents: bool = True,
        record_undo: bool = True,
        notify_change: bool = True,
    ) -> bool:
        if role != Qt.EditRole or not index.isValid():
            return False
        old_value = self.sheet.raw_value(index.row(), index.column())
        next_value = "" if value is None else str(value)
        changed = self.sheet.set_value(index.row(), index.column(), next_value)
        if not changed:
            return True
        if record_undo:
            self._push_undo([(index.row(), index.column(), old_value, next_value)])
        self.evaluator.invalidate_sheet()
        self.dataChanged.emit(index, index, [Qt.DisplayRole, Qt.EditRole])
        if refresh_dependents:
            self.refresh_formulas()
        if notify_change:
            self._notify_values_changed([(index.row(), index.column(), next_value)])
        return True

    def set_values(
        self,
        values: list[tuple[int, int, object]],
        refresh_dependents: bool = True,
        record_undo: bool = True,
        notify_change: bool = True,
    ) -> None:
        if not values:
            return
        max_row = max(row for row, _column, _value in values)
        max_column = max(column for _row, column, _value in values)
        self._ensure_model_size(max_row, max_column)
        top = min(row for row, _column, _value in values)
        bottom = max(row for row, _column, _value in values)
        left = min(column for _row, column, _value in values)
        right = max(column for _row, column, _value in values)
        changed = False
        changes: list[tuple[int, int, object, object]] = []
        for row, column, value in values:
            old_value = self.sheet.raw_value(row, column)
            next_value = "" if value is None else str(value)
            if old_value != next_value:
                changes.append((row, column, old_value, next_value))
            changed = self.sheet.set_value(row, column, next_value, touch=False) or changed
        if not changed:
            return
        self.sheet.bump_revision()
        if record_undo and changes:
            self._push_undo(changes)
        self.evaluator.invalidate_sheet()
        self.dataChanged.emit(self.index(top, left), self.index(bottom, right), [Qt.DisplayRole, Qt.EditRole])
        if refresh_dependents:
            self.refresh_formulas()
        if notify_change and changes:
            self._notify_values_changed([(row, column, new_value) for row, column, _old_value, new_value in changes])

    def clear_indexes(self, indexes: list[QModelIndex], refresh_dependents: bool = True) -> None:
        values = [(index.row(), index.column(), "") for index in indexes if index.isValid()]
        self.set_values(values, refresh_dependents=refresh_dependents)

    def clear_ranges(self, ranges: list[QItemSelectionRange], refresh_dependents: bool = True) -> None:
        normalized = self._normalized_ranges(ranges)
        if not normalized:
            return
        values = [
            (row, column, "")
            for row, column in self._used_positions_in_ranges(normalized)
            if self.sheet.raw_value(row, column) not in ("", None)
        ]
        self.set_values(values, refresh_dependents=refresh_dependents)

    def can_undo(self) -> bool:
        return bool(self.undo_stack)

    def can_redo(self) -> bool:
        return bool(self.redo_stack)

    def undo(self) -> None:
        if not self.undo_stack:
            return
        command = self.undo_stack.pop()
        command.undo()
        self.redo_stack.append(command)
        self._notify_history_changed()

    def redo(self) -> None:
        if not self.redo_stack:
            return
        command = self.redo_stack.pop()
        command.redo()
        self.undo_stack.append(command)
        self._notify_history_changed()

    def _push_undo(self, changes: list[tuple[int, int, object, object]]) -> None:
        self.undo_stack.append(CellChangeCommand(self, changes))
        if len(self.undo_stack) > self.history_limit:
            self.undo_stack = self.undo_stack[-self.history_limit :]
        self.redo_stack.clear()
        self._notify_history_changed()

    def _notify_history_changed(self) -> None:
        if self.history_changed:
            self.history_changed()

    def _notify_values_changed(self, values: list[tuple[int, int, object]]) -> None:
        if self.values_changed:
            self.values_changed(self.sheet, values)

    def flags(self, index: QModelIndex) -> Qt.ItemFlags:
        if not index.isValid():
            return Qt.NoItemFlags
        return Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsEditable

    def headerData(self, section: int, orientation: Qt.Orientation, role: int = Qt.DisplayRole):
        if role != Qt.DisplayRole:
            return None
        if orientation == Qt.Horizontal:
            return index_to_column_name(section)
        return str(section + 1)

    def set_style_for_indexes(self, indexes: list[QModelIndex], **changes) -> None:
        if not indexes:
            return
        top = min(index.row() for index in indexes)
        bottom = max(index.row() for index in indexes)
        left = min(index.column() for index in indexes)
        right = max(index.column() for index in indexes)
        for index in indexes:
            self.sheet.set_style(index.row(), index.column(), **changes)
        self.dataChanged.emit(self.index(top, left), self.index(bottom, right), [Qt.FontRole, Qt.ForegroundRole, Qt.BackgroundRole, Qt.TextAlignmentRole])

    def set_style_for_ranges(self, ranges: list[QItemSelectionRange], **changes) -> None:
        normalized = self._normalized_ranges(ranges)
        if not normalized:
            return
        roles = [Qt.FontRole, Qt.ForegroundRole, Qt.BackgroundRole, Qt.TextAlignmentRole]
        for top, left, bottom, right in normalized:
            for row in range(top, bottom + 1):
                for column in range(left, right + 1):
                    self.sheet.set_style(row, column, **changes)
            self.dataChanged.emit(self.index(top, left), self.index(bottom, right), roles)

    def set_zoom_factor(self, factor: float) -> None:
        if abs(self.zoom_factor - factor) < 0.001:
            return
        self.zoom_factor = factor

    def insert_rows(self, start: int, count: int = 1) -> None:
        self.beginInsertRows(QModelIndex(), start, start + count - 1)
        self.sheet.insert_rows(start, count)
        self.endInsertRows()
        self.evaluator.invalidate_sheet()

    def remove_rows(self, start: int, count: int = 1) -> None:
        self.beginRemoveRows(QModelIndex(), start, start + count - 1)
        self.sheet.remove_rows(start, count)
        self.endRemoveRows()
        self.evaluator.invalidate_sheet()

    def insert_columns(self, start: int, count: int = 1) -> None:
        self.beginInsertColumns(QModelIndex(), start, start + count - 1)
        self.sheet.insert_columns(start, count)
        self.endInsertColumns()
        self.evaluator.invalidate_sheet()

    def remove_columns(self, start: int, count: int = 1) -> None:
        self.beginRemoveColumns(QModelIndex(), start, start + count - 1)
        self.sheet.remove_columns(start, count)
        self.endRemoveColumns()
        self.evaluator.invalidate_sheet()

    def refresh_all(self) -> None:
        if self.rowCount() and self.columnCount():
            self.dataChanged.emit(self.index(0, 0), self.index(self.rowCount() - 1, self.columnCount() - 1), [Qt.DisplayRole])

    def refresh_formulas(self) -> None:
        formula_positions = list(self.sheet.formula_cells)
        if not formula_positions:
            return
        if len(formula_positions) <= 150:
            for row, column in formula_positions:
                index = self.index(row, column)
                self.dataChanged.emit(index, index, [Qt.DisplayRole])
            return
        top = min(row for row, _column in formula_positions)
        bottom = max(row for row, _column in formula_positions)
        left = min(column for _row, column in formula_positions)
        right = max(column for _row, column in formula_positions)
        self.dataChanged.emit(self.index(top, left), self.index(bottom, right), [Qt.DisplayRole])

    def _ensure_model_size(self, row: int, column: int) -> None:
        if row >= self.sheet.row_count:
            old_count = self.sheet.row_count
            self.beginInsertRows(QModelIndex(), old_count, row)
            self.sheet.row_count = row + 1
            self.endInsertRows()
        if column >= self.sheet.column_count:
            old_count = self.sheet.column_count
            self.beginInsertColumns(QModelIndex(), old_count, column)
            self.sheet.column_count = column + 1
            self.endInsertColumns()

    def _normalized_ranges(self, ranges: list[QItemSelectionRange]) -> list[tuple[int, int, int, int]]:
        normalized = []
        for item in ranges:
            top = max(0, item.top())
            bottom = max(top, item.bottom())
            left = max(0, item.left())
            right = max(left, item.right())
            normalized.append((top, left, bottom, right))
        return normalized

    def _used_positions_in_ranges(self, ranges: list[tuple[int, int, int, int]]):
        for row, column in sorted(self.sheet.cells):
            if any(top <= row <= bottom and left <= column <= right for top, left, bottom, right in ranges):
                yield row, column

    def _font_from_style(self, style: CellStyle) -> QFont:
        font = QFont(style.font_family, max(1, round(style.font_size * self.zoom_factor)))
        font.setBold(style.bold)
        font.setItalic(style.italic)
        font.setUnderline(style.underline)
        return font

    def _alignment_from_style(self, style: CellStyle) -> Qt.AlignmentFlag:
        horizontal = {
            "general": Qt.AlignVCenter | Qt.AlignLeft,
            "left": Qt.AlignVCenter | Qt.AlignLeft,
            "center": Qt.AlignVCenter | Qt.AlignHCenter,
            "right": Qt.AlignVCenter | Qt.AlignRight,
        }.get(style.horizontal, Qt.AlignVCenter | Qt.AlignLeft)
        return horizontal
