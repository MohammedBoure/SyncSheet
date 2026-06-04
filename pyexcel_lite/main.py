"""PySide6 desktop spreadsheet application."""

from __future__ import annotations

import sys
from pathlib import Path

from PySide6.QtCore import QItemSelection, QModelIndex, Qt
from PySide6.QtGui import QAction, QColor, QKeySequence, QShortcut
from PySide6.QtWidgets import (
    QApplication,
    QColorDialog,
    QComboBox,
    QFileDialog,
    QFontComboBox,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QSpinBox,
    QSplitter,
    QStatusBar,
    QTableView,
    QTabWidget,
    QToolBar,
    QToolButton,
    QVBoxLayout,
    QWidget,
)

from .cell_address import index_to_column_name
from .formula import FormulaEvaluator, to_number
from .io_xlsx import export_csv, load_xlsx, save_xlsx
from .qt_model import WorksheetTableModel
from .workbook import WorkbookData, WorksheetData


class SpreadsheetView(QTableView):
    def __init__(self, window: "SpreadsheetWindow"):
        super().__init__()
        self.window = window
        self.setAlternatingRowColors(True)
        self.setSortingEnabled(False)
        self.setSelectionMode(QTableView.ExtendedSelection)
        self.setSelectionBehavior(QTableView.SelectItems)
        self.horizontalHeader().setDefaultSectionSize(95)
        self.verticalHeader().setDefaultSectionSize(26)
        self.setCornerButtonEnabled(True)
        self.setWordWrap(False)
        QShortcut(QKeySequence.Copy, self, self.copy_selection)
        QShortcut(QKeySequence.Cut, self, self.cut_selection)
        QShortcut(QKeySequence.Paste, self, self.paste_selection)
        QShortcut(QKeySequence.Delete, self, self.clear_selection)

    def copy_selection(self) -> None:
        indexes = self.selectedIndexes()
        if not indexes:
            return
        rows = sorted(set(index.row() for index in indexes))
        columns = sorted(set(index.column() for index in indexes))
        model = self.model()
        lines = []
        for row in rows:
            values = []
            for column in columns:
                values.append(str(model.data(model.index(row, column), Qt.EditRole) or ""))
            lines.append("\t".join(values))
        QApplication.clipboard().setText("\n".join(lines))

    def cut_selection(self) -> None:
        self.copy_selection()
        self.clear_selection()

    def paste_selection(self) -> None:
        model = self.model()
        current = self.currentIndex()
        if not current.isValid():
            current = model.index(0, 0)
        text = QApplication.clipboard().text()
        if not text:
            return
        for row_offset, line in enumerate(text.splitlines()):
            for column_offset, value in enumerate(line.split("\t")):
                target = model.index(current.row() + row_offset, current.column() + column_offset)
                model.setData(target, value, Qt.EditRole)
        model.refresh_all()

    def clear_selection(self) -> None:
        indexes = self.selectedIndexes()
        if not indexes:
            return
        model = self.model()
        for index in indexes:
            model.setData(index, "", Qt.EditRole)
        model.refresh_all()


class SpreadsheetWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.workbook = WorkbookData()
        self.evaluator = FormulaEvaluator(self.workbook)
        self.models: list[WorksheetTableModel] = []
        self.current_path: Path | None = None
        self.setWindowTitle("PyExcel Lite")
        self.resize(1280, 760)
        self._build_actions()
        self._build_ui()
        self.load_workbook(self.workbook)

    @property
    def current_view(self) -> SpreadsheetView:
        return self.tabs.currentWidget()

    @property
    def current_model(self) -> WorksheetTableModel:
        return self.current_view.model()

    @property
    def current_sheet(self) -> WorksheetData:
        return self.current_model.sheet

    def _build_actions(self) -> None:
        self.new_action = QAction("New", self, shortcut=QKeySequence.New, triggered=self.new_file)
        self.open_action = QAction("Open", self, shortcut=QKeySequence.Open, triggered=self.open_file)
        self.save_action = QAction("Save", self, shortcut=QKeySequence.Save, triggered=self.save_file)
        self.save_as_action = QAction("Save As", self, shortcut=QKeySequence.SaveAs, triggered=self.save_file_as)
        self.export_csv_action = QAction("Export CSV", self, triggered=self.export_current_csv)
        self.add_sheet_action = QAction("Add Sheet", self, triggered=self.add_sheet)
        self.rename_sheet_action = QAction("Rename Sheet", self, triggered=self.rename_sheet)
        self.delete_sheet_action = QAction("Delete Sheet", self, triggered=self.delete_sheet)
        self.insert_row_action = QAction("Insert Row", self, triggered=self.insert_row)
        self.delete_row_action = QAction("Delete Row", self, triggered=self.delete_row)
        self.insert_column_action = QAction("Insert Column", self, triggered=self.insert_column)
        self.delete_column_action = QAction("Delete Column", self, triggered=self.delete_column)
        self.clear_action = QAction("Clear", self, shortcut=QKeySequence.Delete, triggered=self.clear_cells)
        self.about_action = QAction("About", self, triggered=self.about)

    def _build_ui(self) -> None:
        self._build_menu()
        self._build_toolbars()

        self.tabs = QTabWidget()
        self.tabs.setTabsClosable(False)
        self.tabs.setMovable(True)
        self.tabs.currentChanged.connect(self.on_tab_changed)

        self.inspector = self._build_inspector()
        splitter = QSplitter()
        splitter.addWidget(self.tabs)
        splitter.addWidget(self.inspector)
        splitter.setStretchFactor(0, 5)
        splitter.setStretchFactor(1, 1)
        self.setCentralWidget(splitter)

        self.setStatusBar(QStatusBar())
        self.statusBar().showMessage("Ready")

    def _build_menu(self) -> None:
        file_menu = self.menuBar().addMenu("File")
        file_menu.addActions([self.new_action, self.open_action, self.save_action, self.save_as_action, self.export_csv_action])
        sheet_menu = self.menuBar().addMenu("Sheet")
        sheet_menu.addActions([self.add_sheet_action, self.rename_sheet_action, self.delete_sheet_action])
        edit_menu = self.menuBar().addMenu("Edit")
        edit_menu.addActions([self.insert_row_action, self.delete_row_action, self.insert_column_action, self.delete_column_action, self.clear_action])
        help_menu = self.menuBar().addMenu("Help")
        help_menu.addAction(self.about_action)

    def _build_toolbars(self) -> None:
        file_bar = QToolBar("File")
        self.addToolBar(file_bar)
        file_bar.addActions([self.new_action, self.open_action, self.save_action])
        file_bar.addSeparator()
        file_bar.addActions([self.add_sheet_action, self.export_csv_action])

        formula_bar = QToolBar("Formula")
        self.addToolBarBreak()
        self.addToolBar(formula_bar)
        self.name_box = QLineEdit()
        self.name_box.setReadOnly(True)
        self.name_box.setFixedWidth(90)
        self.formula_input = QLineEdit()
        self.formula_input.setPlaceholderText("Type a value or formula, for example =SUM(A1:A5)")
        self.formula_input.returnPressed.connect(self.commit_formula_bar)
        formula_bar.addWidget(QLabel("Cell"))
        formula_bar.addWidget(self.name_box)
        formula_bar.addWidget(QLabel("fx"))
        formula_bar.addWidget(self.formula_input)

        format_bar = QToolBar("Format")
        self.addToolBarBreak()
        self.addToolBar(format_bar)
        self.font_box = QFontComboBox()
        self.font_box.currentFontChanged.connect(lambda font: self.apply_style(font_family=font.family()))
        self.size_box = QSpinBox()
        self.size_box.setRange(6, 48)
        self.size_box.setValue(10)
        self.size_box.valueChanged.connect(lambda value: self.apply_style(font_size=value))
        self.bold_button = self._format_button("B", "Bold", lambda checked: self.apply_style(bold=checked))
        self.italic_button = self._format_button("I", "Italic", lambda checked: self.apply_style(italic=checked))
        self.underline_button = self._format_button("U", "Underline", lambda checked: self.apply_style(underline=checked))
        self.align_box = QComboBox()
        self.align_box.addItems(["general", "left", "center", "right"])
        self.align_box.currentTextChanged.connect(lambda value: self.apply_style(horizontal=value))
        self.number_format_box = QComboBox()
        self.number_format_box.addItems(["General", "0", "0.00", "#,##0", "#,##0.00", "0%", "$#,##0.00", "yyyy-mm-dd"])
        self.number_format_box.currentTextChanged.connect(lambda value: self.apply_style(number_format=value))
        self.text_color_button = QToolButton()
        self.text_color_button.setText("Text")
        self.text_color_button.clicked.connect(lambda: self.pick_color("text_color"))
        self.fill_color_button = QToolButton()
        self.fill_color_button.setText("Fill")
        self.fill_color_button.clicked.connect(lambda: self.pick_color("fill_color"))
        for widget in [self.font_box, self.size_box, self.bold_button, self.italic_button, self.underline_button, QLabel("Align"), self.align_box, QLabel("Format"), self.number_format_box, self.text_color_button, self.fill_color_button]:
            format_bar.addWidget(widget)

    def _format_button(self, text: str, tooltip: str, callback) -> QToolButton:
        button = QToolButton()
        button.setText(text)
        button.setCheckable(True)
        button.setToolTip(tooltip)
        button.toggled.connect(callback)
        return button

    def _build_inspector(self) -> QWidget:
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.addWidget(QLabel("Selection"))
        self.selection_label = QLabel("No selection")
        self.selection_label.setWordWrap(True)
        layout.addWidget(self.selection_label)
        layout.addWidget(QLabel("Quick stats"))
        self.stats_label = QLabel("Sum: 0\nAverage: 0\nCount: 0")
        self.stats_label.setWordWrap(True)
        layout.addWidget(self.stats_label)
        layout.addStretch(1)
        return widget

    def load_workbook(self, workbook: WorkbookData) -> None:
        self.workbook = workbook
        self.evaluator = FormulaEvaluator(self.workbook)
        self.models.clear()
        self.tabs.clear()
        for sheet in workbook.sheets:
            model = WorksheetTableModel(sheet, self.evaluator)
            view = SpreadsheetView(self)
            view.setModel(model)
            view.selectionModel().selectionChanged.connect(self.on_selection_changed)
            self.models.append(model)
            self.tabs.addTab(view, sheet.name)
        self.tabs.setCurrentIndex(workbook.active_sheet_index)
        self.current_path = Path(workbook.path) if workbook.path else None
        self.update_window_title()

    def new_file(self) -> None:
        self.load_workbook(WorkbookData())
        self.statusBar().showMessage("New workbook created")

    def open_file(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "Open workbook", "", "Excel workbooks (*.xlsx)")
        if not path:
            return
        try:
            self.load_workbook(load_xlsx(path))
            self.statusBar().showMessage(f"Opened {path}")
        except Exception as exc:
            QMessageBox.critical(self, "Open failed", str(exc))

    def save_file(self) -> None:
        if self.current_path is None:
            self.save_file_as()
            return
        self._save_to_path(self.current_path)

    def save_file_as(self) -> None:
        path, _ = QFileDialog.getSaveFileName(self, "Save workbook", "", "Excel workbooks (*.xlsx)")
        if not path:
            return
        if not path.lower().endswith(".xlsx"):
            path += ".xlsx"
        self._save_to_path(Path(path))

    def _save_to_path(self, path: Path) -> None:
        try:
            self.workbook.active_sheet_index = self.tabs.currentIndex()
            save_xlsx(self.workbook, path)
            self.current_path = path
            self.update_window_title()
            self.statusBar().showMessage(f"Saved {path}")
        except Exception as exc:
            QMessageBox.critical(self, "Save failed", str(exc))

    def export_current_csv(self) -> None:
        path, _ = QFileDialog.getSaveFileName(self, "Export current sheet", f"{self.current_sheet.name}.csv", "CSV files (*.csv)")
        if not path:
            return
        try:
            export_csv(self.current_sheet, path)
            self.statusBar().showMessage(f"Exported {path}")
        except Exception as exc:
            QMessageBox.critical(self, "Export failed", str(exc))

    def add_sheet(self) -> None:
        sheet = self.workbook.add_sheet()
        model = WorksheetTableModel(sheet, self.evaluator)
        view = SpreadsheetView(self)
        view.setModel(model)
        view.selectionModel().selectionChanged.connect(self.on_selection_changed)
        self.models.append(model)
        self.tabs.addTab(view, sheet.name)
        self.tabs.setCurrentWidget(view)

    def rename_sheet(self) -> None:
        text, ok = QInputDialog.getText(self, "Rename sheet", "Sheet name", text=self.current_sheet.name)
        if ok and text.strip():
            self.current_sheet.name = text.strip()
            self.tabs.setTabText(self.tabs.currentIndex(), self.current_sheet.name)

    def delete_sheet(self) -> None:
        if len(self.workbook.sheets) == 1:
            QMessageBox.information(self, "Delete sheet", "A workbook must keep at least one sheet.")
            return
        index = self.tabs.currentIndex()
        self.workbook.remove_sheet(index)
        self.tabs.removeTab(index)
        self.models.pop(index)

    def insert_row(self) -> None:
        row = self.current_view.currentIndex().row()
        self.current_model.insert_rows(max(row, 0), 1)

    def delete_row(self) -> None:
        row = self.current_view.currentIndex().row()
        if row >= 0:
            self.current_model.remove_rows(row, 1)

    def insert_column(self) -> None:
        column = self.current_view.currentIndex().column()
        self.current_model.insert_columns(max(column, 0), 1)

    def delete_column(self) -> None:
        column = self.current_view.currentIndex().column()
        if column >= 0:
            self.current_model.remove_columns(column, 1)

    def clear_cells(self) -> None:
        self.current_view.clear_selection()

    def commit_formula_bar(self) -> None:
        index = self.current_view.currentIndex()
        if not index.isValid():
            return
        self.current_model.setData(index, self.formula_input.text(), Qt.EditRole)
        self.current_model.refresh_all()

    def apply_style(self, **changes) -> None:
        if not hasattr(self, "tabs") or not self.tabs.count():
            return
        indexes = self.current_view.selectedIndexes()
        if not indexes and self.current_view.currentIndex().isValid():
            indexes = [self.current_view.currentIndex()]
        self.current_model.set_style_for_indexes(indexes, **changes)

    def pick_color(self, style_key: str) -> None:
        color = QColorDialog.getColor(QColor("#111827" if style_key == "text_color" else "#ffffff"), self)
        if color.isValid():
            self.apply_style(**{style_key: color.name()})

    def on_tab_changed(self, index: int) -> None:
        if index >= 0:
            self.workbook.active_sheet_index = index
            self.update_formula_bar()
            self.update_window_title()

    def on_selection_changed(self, selected: QItemSelection, _deselected: QItemSelection) -> None:
        self.update_formula_bar()
        self.update_selection_stats()

    def update_formula_bar(self) -> None:
        if not self.tabs.count():
            return
        index = self.current_view.currentIndex()
        if not index.isValid():
            self.name_box.clear()
            self.formula_input.clear()
            return
        address = index_to_column_name(index.column()) + str(index.row() + 1)
        self.name_box.setText(address)
        self.formula_input.setText(str(self.current_model.data(index, Qt.EditRole) or ""))

    def update_selection_stats(self) -> None:
        indexes = self.current_view.selectedIndexes()
        if not indexes:
            self.selection_label.setText("No selection")
            self.stats_label.setText("Sum: 0\nAverage: 0\nCount: 0")
            return
        rows = [index.row() for index in indexes]
        columns = [index.column() for index in indexes]
        first = index_to_column_name(min(columns)) + str(min(rows) + 1)
        last = index_to_column_name(max(columns)) + str(max(rows) + 1)
        self.selection_label.setText(f"{first}:{last}\nCells: {len(indexes)}")
        numbers = []
        for index in indexes:
            value = self.evaluator.evaluate_cell(self.current_sheet, index.row(), index.column())
            try:
                numbers.append(to_number(value))
            except Exception:
                pass
        total = sum(numbers)
        average = total / len(numbers) if numbers else 0
        self.stats_label.setText(f"Sum: {total:g}\nAverage: {average:g}\nCount: {len(numbers)}")

    def update_window_title(self) -> None:
        name = self.current_path.name if self.current_path else "Untitled"
        self.setWindowTitle(f"{name} - PyExcel Lite")

    def about(self) -> None:
        QMessageBox.about(
            self,
            "About PyExcel Lite",
            "PyExcel Lite\n\nA PySide6 spreadsheet sample with formulas, formatting, multiple sheets, XLSX save/load, CSV export, and clipboard editing.",
        )


def main() -> int:
    app = QApplication(sys.argv)
    app.setApplicationName("PyExcel Lite")
    window = SpreadsheetWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
