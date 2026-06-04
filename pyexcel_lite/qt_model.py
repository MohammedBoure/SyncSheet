"""Qt table model for worksheet data."""

from __future__ import annotations

from PySide6.QtCore import QAbstractTableModel, QModelIndex, Qt
from PySide6.QtGui import QBrush, QColor, QFont

from .cell_address import index_to_column_name
from .formula import FormulaEvaluator
from .workbook import CellStyle, WorksheetData


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
        cell = self.sheet.get_cell(row, column)
        if role == Qt.DisplayRole:
            return self.evaluator.display(cell.value, self.sheet, row, column)
        if role == Qt.EditRole:
            return cell.value
        if role == Qt.FontRole:
            return self._font_from_style(cell.style)
        if role == Qt.ForegroundRole:
            return QBrush(QColor(cell.style.text_color))
        if role == Qt.BackgroundRole:
            return QBrush(QColor(cell.style.fill_color))
        if role == Qt.TextAlignmentRole:
            return self._alignment_from_style(cell.style)
        return None

    def setData(self, index: QModelIndex, value, role: int = Qt.EditRole) -> bool:
        if role != Qt.EditRole or not index.isValid():
            return False
        self.sheet.set_value(index.row(), index.column(), "" if value is None else str(value))
        self.dataChanged.emit(index, index, [Qt.DisplayRole, Qt.EditRole])
        self.refresh_all()
        return True

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
        self.zoom_factor = factor
        if self.rowCount() and self.columnCount():
            self.dataChanged.emit(self.index(0, 0), self.index(self.rowCount() - 1, self.columnCount() - 1), [Qt.FontRole, Qt.DisplayRole])

    def insert_rows(self, start: int, count: int = 1) -> None:
        self.beginInsertRows(QModelIndex(), start, start + count - 1)
        self.sheet.insert_rows(start, count)
        self.endInsertRows()

    def remove_rows(self, start: int, count: int = 1) -> None:
        self.beginRemoveRows(QModelIndex(), start, start + count - 1)
        self.sheet.remove_rows(start, count)
        self.endRemoveRows()

    def insert_columns(self, start: int, count: int = 1) -> None:
        self.beginInsertColumns(QModelIndex(), start, start + count - 1)
        self.sheet.insert_columns(start, count)
        self.endInsertColumns()

    def remove_columns(self, start: int, count: int = 1) -> None:
        self.beginRemoveColumns(QModelIndex(), start, start + count - 1)
        self.sheet.remove_columns(start, count)
        self.endRemoveColumns()

    def refresh_all(self) -> None:
        if self.rowCount() and self.columnCount():
            self.dataChanged.emit(self.index(0, 0), self.index(self.rowCount() - 1, self.columnCount() - 1), [Qt.DisplayRole])

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
