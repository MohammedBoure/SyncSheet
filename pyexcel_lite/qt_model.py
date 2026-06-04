"""Qt table model for worksheet data."""

from __future__ import annotations

from PySide6.QtCore import QAbstractTableModel, QModelIndex, Qt
from PySide6.QtGui import QBrush, QColor, QFont

from .cell_address import index_to_column_name
from .formula import FormulaEvaluator
from .workbook import CellStyle, WorksheetData

DEFAULT_STYLE = CellStyle()


class WorksheetTableModel(QAbstractTableModel):
    def __init__(self, sheet: WorksheetData, evaluator: FormulaEvaluator):
        super().__init__()
        self.sheet = sheet
        self.evaluator = evaluator
        self.zoom_factor = 1.0

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

    def setData(self, index: QModelIndex, value, role: int = Qt.EditRole, refresh_dependents: bool = True) -> bool:
        if role != Qt.EditRole or not index.isValid():
            return False
        changed = self.sheet.set_value(index.row(), index.column(), "" if value is None else str(value))
        if not changed:
            return True
        self.evaluator.invalidate_sheet(self.sheet)
        self.dataChanged.emit(index, index, [Qt.DisplayRole, Qt.EditRole])
        if refresh_dependents:
            self.refresh_formulas()
        return True

    def set_values(self, values: list[tuple[int, int, object]], refresh_dependents: bool = True) -> None:
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
        for row, column, value in values:
            changed = self.sheet.set_value(row, column, "" if value is None else str(value), touch=False) or changed
        if not changed:
            return
        self.sheet.bump_revision()
        self.evaluator.invalidate_sheet(self.sheet)
        self.dataChanged.emit(self.index(top, left), self.index(bottom, right), [Qt.DisplayRole, Qt.EditRole])
        if refresh_dependents:
            self.refresh_formulas()

    def clear_indexes(self, indexes: list[QModelIndex], refresh_dependents: bool = True) -> None:
        values = [(index.row(), index.column(), "") for index in indexes if index.isValid()]
        self.set_values(values, refresh_dependents=refresh_dependents)

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

    def set_zoom_factor(self, factor: float) -> None:
        if abs(self.zoom_factor - factor) < 0.001:
            return
        self.zoom_factor = factor
        if self.rowCount() and self.columnCount():
            self.dataChanged.emit(self.index(0, 0), self.index(self.rowCount() - 1, self.columnCount() - 1), [Qt.FontRole, Qt.DisplayRole])

    def insert_rows(self, start: int, count: int = 1) -> None:
        self.beginInsertRows(QModelIndex(), start, start + count - 1)
        self.sheet.insert_rows(start, count)
        self.endInsertRows()
        self.evaluator.invalidate_sheet(self.sheet)

    def remove_rows(self, start: int, count: int = 1) -> None:
        self.beginRemoveRows(QModelIndex(), start, start + count - 1)
        self.sheet.remove_rows(start, count)
        self.endRemoveRows()
        self.evaluator.invalidate_sheet(self.sheet)

    def insert_columns(self, start: int, count: int = 1) -> None:
        self.beginInsertColumns(QModelIndex(), start, start + count - 1)
        self.sheet.insert_columns(start, count)
        self.endInsertColumns()
        self.evaluator.invalidate_sheet(self.sheet)

    def remove_columns(self, start: int, count: int = 1) -> None:
        self.beginRemoveColumns(QModelIndex(), start, start + count - 1)
        self.sheet.remove_columns(start, count)
        self.endRemoveColumns()
        self.evaluator.invalidate_sheet(self.sheet)

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
