"""PySide6 desktop spreadsheet application."""

from __future__ import annotations

import sys
from pathlib import Path

from PySide6.QtCore import QItemSelection, QModelIndex, QSize, Qt
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

from .chart import ChartPoint, ChartWidget
from .cell_address import index_to_column_name
from .formula import FormulaEvaluator, to_number
from .icons import app_icon
from .io_xlsx import export_csv, load_xlsx, save_xlsx
from .network import (
    DEFAULT_PORT,
    CollaborationClient,
    CollaborationEndpoint,
    CollaborationServer,
    cell_update_message,
    local_join_addresses,
    sheet_message,
    structure_message,
    workbook_from_payload,
    workbook_to_payload,
)
from .qt_model import WorksheetTableModel
from .workbook import WorkbookData, WorksheetData

FORMULA_LIBRARY = {
    "Arithmetic": [
        ("Sum selected", "SUM({range})", "SUM"),
        ("Product selected", "PRODUCT({range})", "PRODUCT"),
        ("Rounded sum", "ROUND(SUM({range}), 2)", "ROUND"),
        ("Percent of total", "{cell}/SUM({range})", "PERCENT"),
    ],
    "Statistics": [
        ("Average selected", "AVERAGE({range})", "AVERAGE"),
        ("Median selected", "MEDIAN({range})", "MEDIAN"),
        ("Minimum selected", "MIN({range})", "MIN"),
        ("Maximum selected", "MAX({range})", "MAX"),
        ("Std deviation", "STDEV.S({range})", "STDEV.S"),
        ("Variance", "VAR.S({range})", "VAR.S"),
        ("Percentile 90", "PERCENTILE.INC({range}, 0.9)", "PERCENTILE"),
        ("Quartile 3", "QUARTILE.INC({range}, 3)", "QUARTILE"),
    ],
    "Logic": [
        ("Positive check", 'IF({cell}>0, "OK", "Check")', "IF"),
        ("Error fallback", "IFERROR({cell}, 0)", "IFERROR"),
        ("Blank check", 'IF(ISBLANK({cell}), "Empty", {cell})', "ISBLANK"),
    ],
    "Lookup": [
        ("Index first", "INDEX({range}, 1, 1)", "INDEX"),
        ("Rank current", "RANK.EQ({cell}, {range}, 0)", "RANK"),
        ("Choose example", 'CHOOSE(1, "A", "B", "C")', "CHOOSE"),
    ],
    "Algorithms": [
        ("Double current", "{cell}*2", "x2"),
        ("Weighted score", "ROUND({cell}*0.7+SUM({range})*0.3, 2)", "WEIGHT"),
        ("Normalize in range", "IFERROR(({cell}-MIN({range}))/(MAX({range})-MIN({range})), 0)", "NORMALIZE"),
        ("Row calculation", "A{row}*2", "ROW"),
    ],
}


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

    def wheelEvent(self, event) -> None:
        if event.modifiers() & Qt.ControlModifier:
            if event.angleDelta().y() > 0:
                self.window.zoom_in()
            else:
                self.window.zoom_out()
            event.accept()
            return
        super().wheelEvent(event)

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
        values = []
        for row_offset, line in enumerate(text.splitlines()):
            for column_offset, value in enumerate(line.split("\t")):
                values.append((current.row() + row_offset, current.column() + column_offset, value))
        model.set_values(values, refresh_dependents=True)

    def clear_selection(self) -> None:
        indexes = self.selectedIndexes()
        if not indexes:
            return
        self.model().clear_indexes(indexes, refresh_dependents=True)


class SpreadsheetWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.workbook = WorkbookData()
        self.evaluator = FormulaEvaluator(self.workbook)
        self.models: list[WorksheetTableModel] = []
        self.current_path: Path | None = None
        self.zoom_percent = 100
        self.max_stats_cells = 5000
        self.collaboration: CollaborationEndpoint | None = None
        self.collaboration_role = ""
        self.collaboration_status = "Offline"
        self.collaboration_clients = 0
        self.applying_remote_update = False
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
        self.undo_action = QAction("Undo", self, shortcut=QKeySequence.Undo, triggered=self.undo)
        self.redo_action = QAction("Redo", self, shortcut=QKeySequence.Redo, triggered=self.redo)
        self.chart_action = QAction("Chart", self, triggered=self.create_chart_from_selection)
        self.host_network_action = QAction("Host", self, triggered=self.host_collaboration)
        self.join_network_action = QAction("Join", self, triggered=self.join_collaboration)
        self.leave_network_action = QAction("Leave", self, triggered=self.leave_collaboration)
        self.leave_network_action.setEnabled(False)
        self.zoom_in_action = QAction("Zoom In", self, triggered=self.zoom_in)
        self.zoom_in_action.setShortcuts([QKeySequence("Ctrl++"), QKeySequence("Ctrl+=")])
        self.zoom_out_action = QAction("Zoom Out", self, shortcut=QKeySequence("Ctrl+-"), triggered=self.zoom_out)
        self.zoom_reset_action = QAction("Reset Zoom", self, shortcut=QKeySequence("Ctrl+0"), triggered=self.reset_zoom)
        self.about_action = QAction("About", self, triggered=self.about)
        self._decorate_actions()

    def _decorate_actions(self) -> None:
        action_icons = {
            self.new_action: ("new", "Create a new workbook"),
            self.open_action: ("open", "Open an Excel workbook"),
            self.save_action: ("save", "Save the current workbook"),
            self.save_as_action: ("save_as", "Save this workbook with a new name"),
            self.export_csv_action: ("csv", "Export the active sheet as CSV"),
            self.add_sheet_action: ("sheet_add", "Add a new worksheet"),
            self.rename_sheet_action: ("sheet_rename", "Rename the active worksheet"),
            self.delete_sheet_action: ("sheet_delete", "Delete the active worksheet"),
            self.insert_row_action: ("row_insert", "Insert a row above the current row"),
            self.delete_row_action: ("row_delete", "Delete the current row"),
            self.insert_column_action: ("column_insert", "Insert a column before the current column"),
            self.delete_column_action: ("column_delete", "Delete the current column"),
            self.clear_action: ("clear", "Clear selected cells"),
            self.undo_action: ("undo", "Undo the last cell edit"),
            self.redo_action: ("redo", "Redo the last undone edit"),
            self.chart_action: ("chart", "Create a chart from the selected cells"),
            self.host_network_action: ("network_host", "Host a realtime collaboration session"),
            self.join_network_action: ("network_join", "Join a realtime collaboration session"),
            self.leave_network_action: ("network_leave", "Leave the collaboration session"),
            self.zoom_in_action: ("zoom_in", "Zoom in"),
            self.zoom_out_action: ("zoom_out", "Zoom out"),
            self.zoom_reset_action: ("zoom_reset", "Reset zoom"),
        }
        for action, (icon_name, tooltip) in action_icons.items():
            action.setIcon(app_icon(icon_name))
            action.setToolTip(tooltip)
            action.setStatusTip(tooltip)

    def _build_ui(self) -> None:
        self._build_menu()
        self._build_toolbars()
        self._apply_ribbon_style()

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
        edit_menu.addActions([self.undo_action, self.redo_action])
        edit_menu.addSeparator()
        edit_menu.addActions([self.insert_row_action, self.delete_row_action, self.insert_column_action, self.delete_column_action, self.clear_action])
        view_menu = self.menuBar().addMenu("View")
        view_menu.addActions([self.zoom_in_action, self.zoom_out_action, self.zoom_reset_action])
        insert_menu = self.menuBar().addMenu("Insert")
        insert_menu.addAction(self.chart_action)
        network_menu = self.menuBar().addMenu("Network")
        network_menu.addActions([self.host_network_action, self.join_network_action, self.leave_network_action])
        help_menu = self.menuBar().addMenu("Help")
        help_menu.addAction(self.about_action)

    def _build_toolbars(self) -> None:
        file_bar = QToolBar("File")
        self._configure_ribbon_toolbar(file_bar)
        self.addToolBar(file_bar)
        file_bar.addActions([self.new_action, self.open_action, self.save_action, self.save_as_action])
        file_bar.addSeparator()
        file_bar.addActions([self.undo_action, self.redo_action])
        file_bar.addSeparator()
        file_bar.addActions([self.add_sheet_action, self.export_csv_action])
        file_bar.addSeparator()
        file_bar.addAction(self.chart_action)
        file_bar.addSeparator()
        file_bar.addActions([self.host_network_action, self.join_network_action, self.leave_network_action])
        file_bar.addSeparator()
        file_bar.addActions([self.zoom_out_action, self.zoom_reset_action, self.zoom_in_action])
        self.zoom_box = QSpinBox()
        self.zoom_box.setRange(40, 220)
        self.zoom_box.setSingleStep(10)
        self.zoom_box.setSuffix("%")
        self.zoom_box.setValue(self.zoom_percent)
        self.zoom_box.valueChanged.connect(self.set_zoom_percent)
        file_bar.addWidget(QLabel("Zoom"))
        file_bar.addWidget(self.zoom_box)

        formula_bar = QToolBar("Formula")
        self._configure_ribbon_toolbar(formula_bar, text_under_icon=False)
        self.addToolBarBreak()
        self.addToolBar(formula_bar)
        self.name_box = QLineEdit()
        self.name_box.setReadOnly(True)
        self.name_box.setFixedWidth(90)
        self.formula_input = QLineEdit()
        self.formula_input.setPlaceholderText("Type a value or formula, for example =SUM(A1:A5)")
        self.formula_input.returnPressed.connect(self.commit_formula_bar)
        formula_icon = QToolButton()
        formula_icon.setIcon(app_icon("formula"))
        formula_icon.setAutoRaise(True)
        formula_icon.setToolTip("Formula bar")
        formula_bar.addWidget(QLabel("Cell"))
        formula_bar.addWidget(self.name_box)
        formula_bar.addWidget(formula_icon)
        formula_bar.addWidget(self.formula_input)

        format_bar = QToolBar("Format")
        self._configure_ribbon_toolbar(format_bar)
        self.addToolBarBreak()
        self.addToolBar(format_bar)
        self.font_box = QFontComboBox()
        self.font_box.currentFontChanged.connect(lambda font: self.apply_style(font_family=font.family()))
        self.size_box = QSpinBox()
        self.size_box.setRange(6, 48)
        self.size_box.setValue(10)
        self.size_box.valueChanged.connect(lambda value: self.apply_style(font_size=value))
        self.bold_button = self._format_button("B", "bold", "Bold", lambda checked: self.apply_style(bold=checked))
        self.italic_button = self._format_button("I", "italic", "Italic", lambda checked: self.apply_style(italic=checked))
        self.underline_button = self._format_button("U", "underline", "Underline", lambda checked: self.apply_style(underline=checked))
        self.align_box = QComboBox()
        self.align_box.addItems(["general", "left", "center", "right"])
        self.align_box.currentTextChanged.connect(lambda value: self.apply_style(horizontal=value))
        self.number_format_box = QComboBox()
        self.number_format_box.addItems(["General", "0", "0.00", "#,##0", "#,##0.00", "0%", "$#,##0.00", "yyyy-mm-dd"])
        self.number_format_box.currentTextChanged.connect(lambda value: self.apply_style(number_format=value))
        self.text_color_button = QToolButton()
        self.text_color_button.setText("Text")
        self.text_color_button.setIcon(app_icon("text_color"))
        self.text_color_button.setToolButtonStyle(Qt.ToolButtonTextUnderIcon)
        self.text_color_button.setToolTip("Text color")
        self.text_color_button.setAutoRaise(True)
        self.text_color_button.clicked.connect(lambda: self.pick_color("text_color"))
        self.fill_color_button = QToolButton()
        self.fill_color_button.setText("Fill")
        self.fill_color_button.setIcon(app_icon("fill_color"))
        self.fill_color_button.setToolButtonStyle(Qt.ToolButtonTextUnderIcon)
        self.fill_color_button.setToolTip("Fill color")
        self.fill_color_button.setAutoRaise(True)
        self.fill_color_button.clicked.connect(lambda: self.pick_color("fill_color"))
        for control in [self.font_box, self.size_box, self.bold_button, self.italic_button, self.underline_button, QLabel("Align"), self.align_box, QLabel("Format"), self.number_format_box, self.text_color_button, self.fill_color_button]:
            format_bar.addWidget(control)
        format_bar.addSeparator()
        format_bar.addActions([self.insert_row_action, self.delete_row_action, self.insert_column_action, self.delete_column_action, self.clear_action])

    def _configure_ribbon_toolbar(self, toolbar: QToolBar, text_under_icon: bool = True) -> None:
        toolbar.setMovable(False)
        toolbar.setFloatable(False)
        toolbar.setIconSize(QSize(24, 24))
        toolbar.setToolButtonStyle(Qt.ToolButtonTextUnderIcon if text_under_icon else Qt.ToolButtonIconOnly)
        toolbar.setAllowedAreas(Qt.TopToolBarArea)

    def _apply_ribbon_style(self) -> None:
        self.setStyleSheet(
            """
            QToolBar {
                background: #f8fafc;
                border: 0;
                border-bottom: 1px solid #d0d7de;
                spacing: 5px;
                padding: 5px 7px;
            }
            QToolBar::separator {
                background: #d8dee4;
                width: 1px;
                margin: 5px 7px;
            }
            QToolBar QToolButton {
                border: 1px solid transparent;
                border-radius: 7px;
                color: #1f2937;
                min-width: 48px;
                padding: 4px 6px;
            }
            QToolBar QToolButton:hover {
                background: #eef7f1;
                border-color: #a8d5ba;
            }
            QToolBar QToolButton:pressed,
            QToolBar QToolButton:checked {
                background: #dff3e8;
                border-color: #107c41;
            }
            QToolBar QComboBox,
            QToolBar QSpinBox,
            QToolBar QLineEdit,
            QToolBar QFontComboBox {
                border: 1px solid #cfd7df;
                border-radius: 6px;
                padding: 4px 7px;
                background: #ffffff;
            }
            QMenuBar {
                background: #ffffff;
                border-bottom: 1px solid #e5e7eb;
            }
            QStatusBar {
                background: #f8fafc;
            }
            """
        )

    def _format_button(self, text: str, icon_name: str, tooltip: str, callback) -> QToolButton:
        button = QToolButton()
        button.setText(text)
        button.setIcon(app_icon(icon_name))
        button.setToolButtonStyle(Qt.ToolButtonTextUnderIcon)
        button.setCheckable(True)
        button.setToolTip(tooltip)
        button.setAutoRaise(True)
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
        layout.addWidget(QLabel("Network"))
        self.network_status_label = QLabel("Offline")
        self.network_status_label.setWordWrap(True)
        layout.addWidget(self.network_status_label)
        layout.addWidget(QLabel("Formula Library"))
        self.formula_category_box = QComboBox()
        self.formula_category_box.addItems(FORMULA_LIBRARY.keys())
        self.formula_category_box.currentTextChanged.connect(self.reload_formula_templates)
        self.formula_template_box = QComboBox()
        self.formula_template_box.currentIndexChanged.connect(self.update_formula_preview)
        self.formula_preview = QLineEdit()
        self.formula_preview.setReadOnly(True)
        formula_buttons = QHBoxLayout()
        self.apply_template_button = QToolButton()
        self.apply_template_button.setText("Apply")
        self.apply_template_button.setIcon(app_icon("formula"))
        self.apply_template_button.setToolButtonStyle(Qt.ToolButtonTextUnderIcon)
        self.apply_template_button.clicked.connect(self.apply_formula_template)
        self.insert_template_button = QToolButton()
        self.insert_template_button.setText("Insert")
        self.insert_template_button.setIcon(app_icon("formula"))
        self.insert_template_button.setToolButtonStyle(Qt.ToolButtonTextUnderIcon)
        self.insert_template_button.clicked.connect(self.insert_formula_template)
        formula_buttons.addWidget(self.apply_template_button)
        formula_buttons.addWidget(self.insert_template_button)
        self.custom_algorithm_input = QLineEdit()
        self.custom_algorithm_input.setPlaceholderText("=A{row}*2")
        custom_buttons = QHBoxLayout()
        self.apply_custom_button = QToolButton()
        self.apply_custom_button.setText("Apply")
        self.apply_custom_button.setIcon(app_icon("formula"))
        self.apply_custom_button.setToolButtonStyle(Qt.ToolButtonTextUnderIcon)
        self.apply_custom_button.clicked.connect(self.apply_custom_algorithm)
        self.fill_custom_button = QToolButton()
        self.fill_custom_button.setText("Fill")
        self.fill_custom_button.setIcon(app_icon("formula"))
        self.fill_custom_button.setToolButtonStyle(Qt.ToolButtonTextUnderIcon)
        self.fill_custom_button.clicked.connect(self.fill_selection_with_custom_algorithm)
        custom_buttons.addWidget(self.apply_custom_button)
        custom_buttons.addWidget(self.fill_custom_button)
        for control in [self.formula_category_box, self.formula_template_box, self.formula_preview]:
            layout.addWidget(control)
        layout.addLayout(formula_buttons)
        layout.addWidget(QLabel("Cell Algorithm"))
        layout.addWidget(self.custom_algorithm_input)
        layout.addLayout(custom_buttons)
        self.reload_formula_templates()
        layout.addWidget(QLabel("Charts"))
        self.chart_type_box = QComboBox()
        self.chart_type_box.addItems(["Bar", "Line", "Pie"])
        self.chart_title_input = QLineEdit()
        self.chart_title_input.setPlaceholderText("Chart title")
        self.create_chart_button = QToolButton()
        self.create_chart_button.setText("Create")
        self.create_chart_button.setIcon(app_icon("chart"))
        self.create_chart_button.setToolButtonStyle(Qt.ToolButtonTextUnderIcon)
        self.create_chart_button.clicked.connect(self.create_chart_from_selection)
        self.chart_widget = ChartWidget()
        layout.addWidget(self.chart_type_box)
        layout.addWidget(self.chart_title_input)
        layout.addWidget(self.create_chart_button)
        layout.addWidget(self.chart_widget)
        layout.addStretch(1)
        return widget

    def reload_formula_templates(self, _category: str | None = None) -> None:
        if not hasattr(self, "formula_template_box"):
            return
        category = self.formula_category_box.currentText() or next(iter(FORMULA_LIBRARY))
        self.formula_template_box.blockSignals(True)
        self.formula_template_box.clear()
        for label, expression, _short_name in FORMULA_LIBRARY[category]:
            self.formula_template_box.addItem(label, expression)
        self.formula_template_box.blockSignals(False)
        self.update_formula_preview()

    def update_formula_preview(self) -> None:
        if not hasattr(self, "formula_preview"):
            return
        template = self.formula_template_box.currentData()
        self.formula_preview.setText(self.build_formula_text(template or ""))

    def selected_range_reference(self) -> str:
        if not hasattr(self, "tabs") or not self.tabs.count():
            return "A1"
        selection = self.current_view.selectionModel().selection()
        if selection.isEmpty():
            index = self.current_view.currentIndex()
            if index.isValid():
                return self.cell_address(index.row(), index.column())
            return "A1"
        ranges = list(selection)
        min_row = min(item.top() for item in ranges)
        max_row = max(item.bottom() for item in ranges)
        min_column = min(item.left() for item in ranges)
        max_column = max(item.right() for item in ranges)
        first = self.cell_address(min_row, min_column)
        last = self.cell_address(max_row, max_column)
        return first if first == last else f"{first}:{last}"

    def current_cell_address(self) -> str:
        if not hasattr(self, "tabs") or not self.tabs.count():
            return "A1"
        index = self.current_view.currentIndex()
        if not index.isValid():
            return "A1"
        return self.cell_address(index.row(), index.column())

    def cell_address(self, row: int, column: int) -> str:
        return index_to_column_name(column) + str(row + 1)

    def formula_context(self, row: int | None = None, column: int | None = None) -> dict[str, str | int]:
        if row is None or column is None:
            index = self.current_view.currentIndex() if hasattr(self, "tabs") and self.tabs.count() else QModelIndex()
            row = index.row() if index.isValid() else 0
            column = index.column() if index.isValid() else 0
        return {
            "cell": self.cell_address(row, column),
            "range": self.selected_range_reference(),
            "row": row + 1,
            "column": index_to_column_name(column),
        }

    def build_formula_text(self, template: str, row: int | None = None, column: int | None = None) -> str:
        if not template:
            return ""
        try:
            expression = template.format(**self.formula_context(row, column))
        except (KeyError, ValueError):
            expression = template
        return expression if expression.startswith("=") else f"={expression}"

    def apply_formula_template(self) -> None:
        formula = self.build_formula_text(self.formula_template_box.currentData() or "")
        self.apply_formula_to_current_cell(formula)

    def insert_formula_template(self) -> None:
        formula = self.build_formula_text(self.formula_template_box.currentData() or "")
        self.formula_input.setText(formula)
        self.formula_input.setFocus()

    def apply_custom_algorithm(self) -> None:
        formula = self.build_formula_text(self.custom_algorithm_input.text().strip())
        self.apply_formula_to_current_cell(formula)

    def fill_selection_with_custom_algorithm(self) -> None:
        template = self.custom_algorithm_input.text().strip()
        if not template or not hasattr(self, "tabs") or not self.tabs.count():
            return
        selection = self.current_view.selectionModel().selection()
        if selection.isEmpty():
            self.apply_custom_algorithm()
            return
        values = []
        for item in list(selection):
            for row in range(item.top(), item.bottom() + 1):
                for column in range(item.left(), item.right() + 1):
                    values.append((row, column, self.build_formula_text(template, row, column)))
        self.current_model.set_values(values, refresh_dependents=True)
        self.update_formula_bar()
        self.update_selection_stats()

    def apply_formula_to_current_cell(self, formula: str) -> None:
        if not formula:
            return
        index = self.current_view.currentIndex()
        if not index.isValid():
            index = self.current_model.index(0, 0)
            self.current_view.setCurrentIndex(index)
        self.current_model.setData(index, formula, Qt.EditRole)
        self.formula_input.setText(formula)
        self.update_selection_stats()

    def create_chart_from_selection(self) -> None:
        points = self.chart_points_from_selection()
        title = self.chart_title_input.text().strip() if hasattr(self, "chart_title_input") else ""
        if not title:
            title = self.selected_range_reference()
        chart_type = self.chart_type_box.currentText() if hasattr(self, "chart_type_box") else "Bar"
        self.chart_widget.set_chart(points, chart_type, title)
        if points:
            self.statusBar().showMessage(f"Chart created with {len(points)} points")
        else:
            self.statusBar().showMessage("No numeric data found for chart")

    def chart_points_from_selection(self) -> list[ChartPoint]:
        if not hasattr(self, "tabs") or not self.tabs.count():
            return []
        selection = self.current_view.selectionModel().selection()
        if selection.isEmpty():
            index = self.current_view.currentIndex()
            if not index.isValid():
                return []
            selection = QItemSelection(index, index)
        ranges = list(selection)
        min_row = min(item.top() for item in ranges)
        max_row = max(item.bottom() for item in ranges)
        min_column = min(item.left() for item in ranges)
        max_column = max(item.right() for item in ranges)
        if max_column > min_column:
            points = self._paired_chart_points(min_row, max_row, min_column, min_column + 1)
            if points:
                return points
        return self._flat_numeric_chart_points(ranges)

    def _paired_chart_points(self, top: int, bottom: int, label_column: int, value_column: int) -> list[ChartPoint]:
        points: list[ChartPoint] = []
        for row in range(top, bottom + 1):
            label = str(self.current_sheet.raw_value(row, label_column) or self.cell_address(row, label_column))
            try:
                value = to_number(self.evaluator.evaluate_cell(self.current_sheet, row, value_column))
            except Exception:
                continue
            points.append(ChartPoint(label, value))
        return points

    def _flat_numeric_chart_points(self, ranges: list[QItemSelection]) -> list[ChartPoint]:
        points: list[ChartPoint] = []
        for item in ranges:
            for row in range(item.top(), item.bottom() + 1):
                for column in range(item.left(), item.right() + 1):
                    try:
                        value = to_number(self.evaluator.evaluate_cell(self.current_sheet, row, column))
                    except Exception:
                        continue
                    points.append(ChartPoint(self.cell_address(row, column), value))
                    if len(points) >= 200:
                        return points
        return points

    def _create_sheet_view(self, sheet: WorksheetData) -> SpreadsheetView:
        model = WorksheetTableModel(sheet, self.evaluator)
        model.history_changed = self.update_undo_redo_actions
        model.values_changed = self.on_local_values_changed
        view = SpreadsheetView(self)
        view.setModel(model)
        self.apply_zoom_to_view(view, model)
        view.selectionModel().selectionChanged.connect(self.on_selection_changed)
        return view

    def _insert_sheet_view(self, index: int, sheet: WorksheetData) -> None:
        view = self._create_sheet_view(sheet)
        model = view.model()
        index = max(0, min(index, len(self.models)))
        self.models.insert(index, model)
        self.tabs.insertTab(index, view, sheet.name)

    def on_local_values_changed(self, sheet: WorksheetData, values: list[tuple[int, int, object]]) -> None:
        self.refresh_workbook_formulas()
        if self.applying_remote_update or self.collaboration is None:
            return
        try:
            sheet_index = self.workbook.sheets.index(sheet)
        except ValueError:
            return
        self.send_collaboration_message(cell_update_message(sheet_index, sheet.name, values))

    def refresh_workbook_formulas(self) -> None:
        for model in self.models:
            model.refresh_formulas()

    def send_collaboration_message(self, message: dict) -> None:
        if self.applying_remote_update or self.collaboration is None or not self.collaboration.running:
            return
        try:
            self.collaboration.send(message)
        except OSError as exc:
            self.show_collaboration_error(f"Send failed: {exc}")

    def send_collaboration_snapshot(self) -> None:
        if self.applying_remote_update or not isinstance(self.collaboration, CollaborationServer):
            return
        self.send_collaboration_message({"type": "snapshot", "workbook": workbook_to_payload(self.workbook)})

    def host_collaboration(self) -> None:
        if self.collaboration is not None:
            self.leave_collaboration()
        port, ok = QInputDialog.getInt(self, "Host collaboration", "Port", DEFAULT_PORT, 1, 65535)
        if not ok:
            return
        server = CollaborationServer(port=port, snapshot_provider=lambda: workbook_to_payload(self.workbook))
        self.attach_collaboration(server, "Host")
        try:
            server.start()
        except OSError as exc:
            server.stop()
            self.detach_collaboration()
            QMessageBox.critical(self, "Network failed", str(exc))

    def join_collaboration(self) -> None:
        if self.collaboration is not None:
            self.leave_collaboration()
        target, ok = QInputDialog.getText(self, "Join collaboration", "Server IP:port", text=f"127.0.0.1:{DEFAULT_PORT}")
        if not ok or not target.strip():
            return
        try:
            host, port = self.parse_collaboration_target(target.strip())
        except ValueError as exc:
            QMessageBox.warning(self, "Join collaboration", str(exc))
            return
        client = CollaborationClient(host, port)
        self.attach_collaboration(client, "Client")
        client.start()

    def leave_collaboration(self) -> None:
        if self.collaboration is None:
            return
        self.collaboration.stop()
        self.detach_collaboration()
        self.statusBar().showMessage("Left collaboration session")

    def attach_collaboration(self, endpoint: CollaborationEndpoint, role: str) -> None:
        self.collaboration = endpoint
        self.collaboration_role = role
        self.collaboration_clients = 0
        endpoint.message_received.connect(self.on_collaboration_message)
        endpoint.status_changed.connect(self.update_collaboration_status)
        endpoint.error_occurred.connect(self.show_collaboration_error)
        endpoint.client_count_changed.connect(self.update_collaboration_client_count)
        self.update_collaboration_status("Starting")
        self.update_collaboration_actions()

    def detach_collaboration(self) -> None:
        endpoint = self.collaboration
        if endpoint is not None:
            for signal, slot in (
                (endpoint.message_received, self.on_collaboration_message),
                (endpoint.status_changed, self.update_collaboration_status),
                (endpoint.error_occurred, self.show_collaboration_error),
                (endpoint.client_count_changed, self.update_collaboration_client_count),
            ):
                try:
                    signal.disconnect(slot)
                except (RuntimeError, TypeError):
                    pass
        self.collaboration = None
        self.collaboration_role = ""
        self.collaboration_clients = 0
        self.update_collaboration_status("Offline")
        self.update_collaboration_actions()

    def parse_collaboration_target(self, text: str) -> tuple[str, int]:
        if ":" not in text:
            return text, DEFAULT_PORT
        host, port_text = text.rsplit(":", 1)
        host = host.strip()
        if not host:
            raise ValueError("Host is required.")
        try:
            port = int(port_text)
        except ValueError as exc:
            raise ValueError("Port must be a number.") from exc
        if not 1 <= port <= 65535:
            raise ValueError("Port must be between 1 and 65535.")
        return host, port

    def update_collaboration_status(self, status: str) -> None:
        self.collaboration_status = status
        if hasattr(self, "network_status_label"):
            if self.collaboration_role == "Host":
                addresses = ", ".join(local_join_addresses(self.collaboration.port)) if isinstance(self.collaboration, CollaborationServer) else ""
                text = f"Server\n{status}\nAddress: {addresses}\nClients: {self.collaboration_clients}"
            elif self.collaboration_role == "Client":
                text = f"Client\n{status}"
            else:
                text = "Offline"
            self.network_status_label.setText(text)
        if self.statusBar():
            self.statusBar().showMessage(status)

    def update_collaboration_client_count(self, count: int) -> None:
        self.collaboration_clients = count
        self.update_collaboration_status(self.collaboration_status)

    def update_collaboration_actions(self) -> None:
        online = self.collaboration is not None
        self.host_network_action.setEnabled(not online)
        self.join_network_action.setEnabled(not online)
        self.leave_network_action.setEnabled(online)

    def show_collaboration_error(self, message: str) -> None:
        self.update_collaboration_status(message)
        if self.statusBar():
            self.statusBar().showMessage(message)

    def on_collaboration_message(self, message: dict) -> None:
        message_type = message.get("type")
        try:
            if message_type == "snapshot":
                self.apply_remote_snapshot(message)
            elif message_type == "cell_update":
                self.apply_remote_cell_update(message)
            elif message_type == "sheet_add":
                self.apply_remote_sheet_add(message)
            elif message_type == "sheet_rename":
                self.apply_remote_sheet_rename(message)
            elif message_type == "sheet_delete":
                self.apply_remote_sheet_delete(message)
            elif message_type in {"insert_rows", "remove_rows", "insert_columns", "remove_columns"}:
                self.apply_remote_structure_update(message)
        except Exception as exc:
            self.show_collaboration_error(f"Network update failed: {exc}")

    def apply_remote_snapshot(self, message: dict) -> None:
        payload = message.get("workbook", {})
        self.applying_remote_update = True
        try:
            self.load_workbook(workbook_from_payload(payload))
        finally:
            self.applying_remote_update = False
        self.update_formula_bar()
        self.update_selection_stats()
        self.statusBar().showMessage("Workbook synchronized")

    def apply_remote_cell_update(self, message: dict) -> None:
        model = self.remote_model(message, create=True)
        if model is None:
            return
        values = [
            (int(item.get("row", 0) or 0), int(item.get("column", 0) or 0), item.get("value", ""))
            for item in message.get("values", [])
        ]
        self.applying_remote_update = True
        try:
            model.set_values(values, refresh_dependents=True, record_undo=False, notify_change=False)
        finally:
            self.applying_remote_update = False
        self.refresh_workbook_formulas()
        if model is self.current_model:
            self.update_formula_bar()
            self.update_selection_stats()
        self.statusBar().showMessage(f"Remote update: {model.sheet.name}")

    def apply_remote_sheet_add(self, message: dict) -> None:
        name = str(message.get("sheet_name") or self.workbook.unique_sheet_name())
        if self.workbook.sheet_by_name(name) is not None:
            return
        index = max(0, min(int(message.get("sheet_index", len(self.workbook.sheets)) or 0), len(self.workbook.sheets)))
        sheet = WorksheetData(name=name)
        self.applying_remote_update = True
        try:
            self.workbook.sheets.insert(index, sheet)
            self._insert_sheet_view(index, sheet)
        finally:
            self.applying_remote_update = False
        self.update_undo_redo_actions()
        self.statusBar().showMessage(f"Remote sheet added: {name}")

    def apply_remote_sheet_rename(self, message: dict) -> None:
        model = self.remote_model(message)
        if model is None:
            return
        name = str(message.get("sheet_name") or model.sheet.name)
        model.sheet.name = name
        self.tabs.setTabText(self.models.index(model), name)
        self.statusBar().showMessage(f"Remote sheet renamed: {name}")

    def apply_remote_sheet_delete(self, message: dict) -> None:
        if len(self.workbook.sheets) == 1:
            return
        model = self.remote_model(message)
        if model is None:
            return
        index = self.models.index(model)
        self.applying_remote_update = True
        try:
            self.workbook.remove_sheet(index)
            self.tabs.removeTab(index)
            self.models.pop(index)
        finally:
            self.applying_remote_update = False
        self.update_undo_redo_actions()
        self.statusBar().showMessage("Remote sheet deleted")

    def apply_remote_structure_update(self, message: dict) -> None:
        model = self.remote_model(message)
        if model is None:
            return
        start = int(message.get("start", 0) or 0)
        count = max(1, int(message.get("count", 1) or 1))
        self.applying_remote_update = True
        try:
            message_type = message.get("type")
            if message_type == "insert_rows":
                model.insert_rows(start, count)
            elif message_type == "remove_rows":
                model.remove_rows(start, count)
            elif message_type == "insert_columns":
                model.insert_columns(start, count)
            elif message_type == "remove_columns":
                model.remove_columns(start, count)
        finally:
            self.applying_remote_update = False
        self.refresh_workbook_formulas()
        self.statusBar().showMessage(f"Remote structure update: {model.sheet.name}")

    def remote_model(self, message: dict, create: bool = False) -> WorksheetTableModel | None:
        raw_index = message.get("sheet_index", -1)
        try:
            sheet_index = int(raw_index)
        except (TypeError, ValueError):
            sheet_index = -1
        if 0 <= sheet_index < len(self.models):
            return self.models[sheet_index]
        sheet_name = str(message.get("sheet_name") or "")
        sheet = self.workbook.sheet_by_name(sheet_name) if sheet_name else None
        if sheet is not None:
            return self.models[self.workbook.sheets.index(sheet)]
        if not create:
            return None
        sheet = WorksheetData(name=sheet_name or self.workbook.unique_sheet_name())
        self.workbook.sheets.append(sheet)
        self._insert_sheet_view(len(self.models), sheet)
        return self.models[-1]

    def load_workbook(self, workbook: WorkbookData) -> None:
        self.workbook = workbook
        self.evaluator = FormulaEvaluator(self.workbook)
        self.models.clear()
        self.tabs.clear()
        for sheet in workbook.sheets:
            self._insert_sheet_view(len(self.models), sheet)
        self.tabs.setCurrentIndex(workbook.active_sheet_index)
        self.current_path = Path(workbook.path) if workbook.path else None
        self.update_window_title()
        self.update_undo_redo_actions()

    def new_file(self) -> None:
        self.load_workbook(WorkbookData())
        self.send_collaboration_snapshot()
        self.statusBar().showMessage("New workbook created")

    def open_file(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "Open workbook", "", "Excel workbooks (*.xlsx)")
        if not path:
            return
        try:
            self.load_workbook(load_xlsx(path))
            self.send_collaboration_snapshot()
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
        self._insert_sheet_view(len(self.models), sheet)
        self.tabs.setCurrentIndex(len(self.models) - 1)
        self.update_undo_redo_actions()
        self.send_collaboration_message(sheet_message("sheet_add", len(self.workbook.sheets) - 1, sheet.name))

    def rename_sheet(self) -> None:
        text, ok = QInputDialog.getText(self, "Rename sheet", "Sheet name", text=self.current_sheet.name)
        if ok and text.strip():
            self.current_sheet.name = text.strip()
            self.tabs.setTabText(self.tabs.currentIndex(), self.current_sheet.name)
            self.send_collaboration_message(sheet_message("sheet_rename", self.tabs.currentIndex(), self.current_sheet.name))

    def delete_sheet(self) -> None:
        if len(self.workbook.sheets) == 1:
            QMessageBox.information(self, "Delete sheet", "A workbook must keep at least one sheet.")
            return
        index = self.tabs.currentIndex()
        sheet_name = self.current_sheet.name
        self.workbook.remove_sheet(index)
        self.tabs.removeTab(index)
        self.models.pop(index)
        self.update_undo_redo_actions()
        self.send_collaboration_message(sheet_message("sheet_delete", index, sheet_name))

    def undo(self) -> None:
        if not self.tabs.count():
            return
        self.current_model.undo()
        self.update_formula_bar()
        self.update_selection_stats()

    def redo(self) -> None:
        if not self.tabs.count():
            return
        self.current_model.redo()
        self.update_formula_bar()
        self.update_selection_stats()

    def update_undo_redo_actions(self) -> None:
        if not hasattr(self, "undo_action") or not hasattr(self, "tabs") or not self.tabs.count():
            return
        self.undo_action.setEnabled(self.current_model.can_undo())
        self.redo_action.setEnabled(self.current_model.can_redo())

    def insert_row(self) -> None:
        row = self.current_view.currentIndex().row()
        start = max(row, 0)
        sheet_index = self.tabs.currentIndex()
        sheet_name = self.current_sheet.name
        self.current_model.insert_rows(start, 1)
        self.refresh_workbook_formulas()
        self.send_collaboration_message(structure_message("insert_rows", sheet_index, sheet_name, start, 1))

    def delete_row(self) -> None:
        row = self.current_view.currentIndex().row()
        if row >= 0:
            sheet_index = self.tabs.currentIndex()
            sheet_name = self.current_sheet.name
            self.current_model.remove_rows(row, 1)
            self.refresh_workbook_formulas()
            self.send_collaboration_message(structure_message("remove_rows", sheet_index, sheet_name, row, 1))

    def insert_column(self) -> None:
        column = self.current_view.currentIndex().column()
        start = max(column, 0)
        sheet_index = self.tabs.currentIndex()
        sheet_name = self.current_sheet.name
        self.current_model.insert_columns(start, 1)
        self.refresh_workbook_formulas()
        self.send_collaboration_message(structure_message("insert_columns", sheet_index, sheet_name, start, 1))

    def delete_column(self) -> None:
        column = self.current_view.currentIndex().column()
        if column >= 0:
            sheet_index = self.tabs.currentIndex()
            sheet_name = self.current_sheet.name
            self.current_model.remove_columns(column, 1)
            self.refresh_workbook_formulas()
            self.send_collaboration_message(structure_message("remove_columns", sheet_index, sheet_name, column, 1))

    def clear_cells(self) -> None:
        self.current_view.clear_selection()

    def zoom_in(self) -> None:
        self.set_zoom_percent(min(220, self.zoom_percent + 10))

    def zoom_out(self) -> None:
        self.set_zoom_percent(max(40, self.zoom_percent - 10))

    def reset_zoom(self) -> None:
        self.set_zoom_percent(100)

    def set_zoom_percent(self, percent: int) -> None:
        next_percent = max(40, min(220, int(percent)))
        if next_percent == self.zoom_percent:
            return
        self.zoom_percent = next_percent
        if hasattr(self, "zoom_box") and self.zoom_box.value() != self.zoom_percent:
            self.zoom_box.blockSignals(True)
            self.zoom_box.setValue(self.zoom_percent)
            self.zoom_box.blockSignals(False)
        if hasattr(self, "tabs"):
            for index in range(self.tabs.count()):
                view = self.tabs.widget(index)
                model = view.model()
                self.apply_zoom_to_view(view, model)
        if self.statusBar():
            self.statusBar().showMessage(f"Zoom {self.zoom_percent}%")

    def apply_zoom_to_view(self, view: SpreadsheetView, model: WorksheetTableModel) -> None:
        factor = self.zoom_percent / 100
        model.set_zoom_factor(factor)
        view.horizontalHeader().setDefaultSectionSize(max(45, round(95 * factor)))
        view.verticalHeader().setDefaultSectionSize(max(18, round(26 * factor)))
        view.setStyleSheet(f"QTableView {{ font-size: {max(7, round(10 * factor))}pt; }}")

    def commit_formula_bar(self) -> None:
        index = self.current_view.currentIndex()
        if not index.isValid():
            return
        self.current_model.setData(index, self.formula_input.text(), Qt.EditRole)

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
            self.update_formula_preview()
            self.update_undo_redo_actions()
            self.update_window_title()

    def on_selection_changed(self, selected: QItemSelection, _deselected: QItemSelection) -> None:
        self.update_formula_bar()
        self.update_selection_stats()
        self.update_formula_preview()

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
        selection = self.current_view.selectionModel().selection()
        if selection.isEmpty():
            self.selection_label.setText("No selection")
            self.stats_label.setText("Sum: 0\nAverage: 0\nCount: 0")
            return
        ranges = list(selection)
        min_row = min(item.top() for item in ranges)
        max_row = max(item.bottom() for item in ranges)
        min_column = min(item.left() for item in ranges)
        max_column = max(item.right() for item in ranges)
        cell_count = sum((item.bottom() - item.top() + 1) * (item.right() - item.left() + 1) for item in ranges)
        first = index_to_column_name(min_column) + str(min_row + 1)
        last = index_to_column_name(max_column) + str(max_row + 1)
        self.selection_label.setText(f"{first}:{last}\nCells: {cell_count}")
        numbers = []
        scanned_count = 0
        for item in ranges:
            for row in range(item.top(), item.bottom() + 1):
                for column in range(item.left(), item.right() + 1):
                    if scanned_count >= self.max_stats_cells:
                        break
                    scanned_count += 1
                    try:
                        value = self.evaluator.evaluate_cell(self.current_sheet, row, column)
                        numbers.append(to_number(value))
                    except Exception:
                        pass
                if scanned_count >= self.max_stats_cells:
                    break
            if scanned_count >= self.max_stats_cells:
                break
        total = sum(numbers)
        average = total / len(numbers) if numbers else 0
        suffix = f"\nScanned: {scanned_count} of {cell_count}" if cell_count > scanned_count else ""
        self.stats_label.setText(f"Sum: {total:g}\nAverage: {average:g}\nCount: {len(numbers)}{suffix}")

    def update_window_title(self) -> None:
        name = self.current_path.name if self.current_path else "Untitled"
        self.setWindowTitle(f"{name} - PyExcel Lite")

    def about(self) -> None:
        QMessageBox.about(
            self,
            "About PyExcel Lite",
            "PyExcel Lite\n\nA PySide6 spreadsheet sample with formulas, formatting, multiple sheets, XLSX save/load, CSV export, clipboard editing, and realtime LAN collaboration.",
        )

    def closeEvent(self, event) -> None:
        if self.collaboration is not None:
            self.collaboration.stop()
        super().closeEvent(event)


def main() -> int:
    app = QApplication(sys.argv)
    app.setApplicationName("PyExcel Lite")
    window = SpreadsheetWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
