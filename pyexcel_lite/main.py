"""PySide6 desktop spreadsheet application."""

from __future__ import annotations

import sys
from pathlib import Path, PurePosixPath

from PySide6.QtCore import QItemSelection, QModelIndex, Qt
from PySide6.QtGui import QAction, QColor, QKeySequence, QShortcut
from PySide6.QtWidgets import (
    QApplication,
    QColorDialog,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFontComboBox,
    QFormLayout,
    QFrame,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QScrollArea,
    QSpinBox,
    QStatusBar,
    QTableView,
    QTabWidget,
    QToolButton,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

from .chart import ChartPoint, ChartWidget
from .cell_address import index_to_column_name
from .formula import FormulaEvaluator, to_number
from .icons import app_icon
from .io_xlsx import export_csv, load_csv, load_xlsx, save_xlsx
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
from .project import (
    ProjectData,
    project_from_payload,
    project_snapshot_message,
    project_to_payload,
    scan_project_folder,
)
from .qt_model import WorksheetTableModel
from .settings import StartupSettings, load_startup_settings, save_startup_settings
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
        self.project = ProjectData()
        self.models: list[WorksheetTableModel] = []
        self.current_path: Path | None = None
        self.zoom_percent = 100
        self.max_stats_cells = 5000
        self.collaboration: CollaborationEndpoint | None = None
        self.collaboration_role = ""
        self.collaboration_status = "Offline"
        self.collaboration_clients = 0
        self.applying_remote_update = False
        self.startup_settings = load_startup_settings()
        self.setWindowTitle("PyExcel Lite")
        self.resize(1280, 760)
        self._build_actions()
        self._build_ui()
        self.load_workbook(self.workbook)
        self.apply_startup_settings()

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
        self.open_project_action = QAction("Open Project", self, triggered=self.open_project)
        self.refresh_project_action = QAction("Refresh Project", self, triggered=self.refresh_project)
        self.open_project_file_action = QAction("Open Item", self, triggered=self.open_selected_project_file)
        self.share_project_action = QAction("Share Project", self, triggered=self.send_project_snapshot)
        self.close_project_action = QAction("Close Project", self, triggered=self.close_project)
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
        self.startup_network_action = QAction("Startup", self, triggered=self.open_startup_settings)
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
            self.open_project_action: ("project", "Open a folder tree as one project"),
            self.refresh_project_action: ("refresh", "Refresh the current project folder"),
            self.open_project_file_action: ("open", "Open the selected project spreadsheet"),
            self.share_project_action: ("network_host", "Share the project snapshot with collaborators"),
            self.close_project_action: ("clear", "Close the current project"),
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
            self.startup_network_action: ("settings", "Choose automatic network startup mode"),
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
        self._apply_ribbon_style()

        self.ribbon = self._build_ribbon()
        self.inspector = self._build_inspector()
        self.tabs = QTabWidget()
        self.tabs.setTabsClosable(False)
        self.tabs.setMovable(True)
        self.tabs.currentChanged.connect(self.on_tab_changed)

        central = QWidget()
        central.setObjectName("mainSurface")
        layout = QVBoxLayout(central)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        layout.addWidget(self.ribbon)
        layout.addWidget(self.inspector)
        layout.addWidget(self.tabs, 1)
        self.setCentralWidget(central)

        self.setStatusBar(QStatusBar())
        self.statusBar().showMessage("Ready")

    def _build_menu(self) -> None:
        file_menu = self.menuBar().addMenu("File")
        file_menu.addActions([self.new_action, self.open_action, self.save_action, self.save_as_action, self.export_csv_action])
        project_menu = self.menuBar().addMenu("Project")
        project_menu.addActions(
            [
                self.open_project_action,
                self.refresh_project_action,
                self.open_project_file_action,
                self.share_project_action,
                self.close_project_action,
            ]
        )
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
        network_menu.addSeparator()
        network_menu.addAction(self.startup_network_action)
        help_menu = self.menuBar().addMenu("Help")
        help_menu.addAction(self.about_action)

    def _build_ribbon(self) -> QTabWidget:
        ribbon = QTabWidget()
        ribbon.setObjectName("excelRibbon")
        ribbon.setDocumentMode(True)
        ribbon.setTabPosition(QTabWidget.North)
        ribbon.setMaximumHeight(172)
        ribbon.addTab(self._home_ribbon_page(), "Home")
        ribbon.addTab(self._project_ribbon_page(), "Project")
        ribbon.addTab(self._insert_ribbon_page(), "Insert")
        ribbon.addTab(self._formulas_ribbon_page(), "Formulas")
        ribbon.addTab(self._data_ribbon_page(), "Data")
        ribbon.addTab(self._view_ribbon_page(), "View")
        ribbon.addTab(self._network_ribbon_page(), "Network")
        return ribbon

    def _ribbon_page(self) -> tuple[QWidget, QHBoxLayout]:
        page = QWidget()
        page.setObjectName("ribbonPage")
        layout = QHBoxLayout(page)
        layout.setContentsMargins(8, 7, 8, 7)
        layout.setSpacing(8)
        return page, layout

    def _ribbon_group(self, parent_layout: QHBoxLayout, title: str) -> QHBoxLayout:
        frame = QFrame()
        frame.setObjectName("ribbonGroup")
        frame_layout = QVBoxLayout(frame)
        frame_layout.setContentsMargins(8, 5, 8, 4)
        frame_layout.setSpacing(4)
        controls = QHBoxLayout()
        controls.setContentsMargins(0, 0, 0, 0)
        controls.setSpacing(5)
        label = QLabel(title)
        label.setObjectName("ribbonGroupTitle")
        label.setAlignment(Qt.AlignCenter)
        frame_layout.addLayout(controls)
        frame_layout.addWidget(label)
        parent_layout.addWidget(frame)
        return controls

    def _ribbon_action_button(self, action: QAction, style: Qt.ToolButtonStyle = Qt.ToolButtonTextUnderIcon) -> QToolButton:
        button = QToolButton()
        button.setObjectName("ribbonButton")
        button.setDefaultAction(action)
        button.setToolButtonStyle(style)
        button.setAutoRaise(False)
        return button

    def _ribbon_callback_button(self, text: str, icon_name: str, callback) -> QToolButton:
        button = QToolButton()
        button.setObjectName("ribbonButton")
        button.setText(text)
        button.setIcon(app_icon(icon_name))
        button.setToolButtonStyle(Qt.ToolButtonTextUnderIcon)
        button.setAutoRaise(False)
        button.clicked.connect(callback)
        return button

    def _ribbon_small_label(self, text: str) -> QLabel:
        label = QLabel(text)
        label.setObjectName("ribbonInlineLabel")
        return label

    def _home_ribbon_page(self) -> QWidget:
        page, layout = self._ribbon_page()
        workbook = self._ribbon_group(layout, "Workbook")
        for action in (self.new_action, self.open_action, self.save_action, self.save_as_action):
            workbook.addWidget(self._ribbon_action_button(action))

        edit = self._ribbon_group(layout, "Edit")
        for action in (self.undo_action, self.redo_action, self.clear_action):
            edit.addWidget(self._ribbon_action_button(action))

        font = self._ribbon_group(layout, "Font")
        self.font_box = QFontComboBox()
        self.font_box.setMaximumWidth(170)
        self.font_box.currentFontChanged.connect(lambda font: self.apply_style(font_family=font.family()))
        self.size_box = QSpinBox()
        self.size_box.setRange(6, 48)
        self.size_box.setValue(10)
        self.size_box.setMaximumWidth(64)
        self.size_box.valueChanged.connect(lambda value: self.apply_style(font_size=value))
        self.bold_button = self._format_button("B", "bold", "Bold", lambda checked: self.apply_style(bold=checked))
        self.italic_button = self._format_button("I", "italic", "Italic", lambda checked: self.apply_style(italic=checked))
        self.underline_button = self._format_button("U", "underline", "Underline", lambda checked: self.apply_style(underline=checked))
        for control in (self.font_box, self.size_box, self.bold_button, self.italic_button, self.underline_button):
            font.addWidget(control)

        color = self._ribbon_group(layout, "Color")
        self.text_color_button = self._ribbon_callback_button("Text", "text_color", lambda: self.pick_color("text_color"))
        self.fill_color_button = self._ribbon_callback_button("Fill", "fill_color", lambda: self.pick_color("fill_color"))
        color.addWidget(self.text_color_button)
        color.addWidget(self.fill_color_button)

        alignment = self._ribbon_group(layout, "Alignment")
        self.align_box = QComboBox()
        self.align_box.addItems(["general", "left", "center", "right"])
        self.align_box.currentTextChanged.connect(lambda value: self.apply_style(horizontal=value))
        alignment.addWidget(self._ribbon_small_label("Align"))
        alignment.addWidget(self.align_box)

        layout.addStretch(1)
        return page

    def _project_ribbon_page(self) -> QWidget:
        page, layout = self._ribbon_page()
        project = self._ribbon_group(layout, "Project")
        for action in (self.open_project_action, self.refresh_project_action, self.open_project_file_action, self.close_project_action):
            project.addWidget(self._ribbon_action_button(action))

        team = self._ribbon_group(layout, "Team")
        team.addWidget(self._ribbon_action_button(self.share_project_action))
        team.addWidget(self._ribbon_action_button(self.host_network_action))
        team.addWidget(self._ribbon_action_button(self.join_network_action))
        layout.addStretch(1)
        return page

    def _insert_ribbon_page(self) -> QWidget:
        page, layout = self._ribbon_page()
        sheets = self._ribbon_group(layout, "Sheets")
        for action in (self.add_sheet_action, self.rename_sheet_action, self.delete_sheet_action):
            sheets.addWidget(self._ribbon_action_button(action))

        rows_columns = self._ribbon_group(layout, "Rows & Columns")
        for action in (self.insert_row_action, self.delete_row_action, self.insert_column_action, self.delete_column_action):
            rows_columns.addWidget(self._ribbon_action_button(action))

        charts = self._ribbon_group(layout, "Charts")
        charts.addWidget(self._ribbon_action_button(self.chart_action))

        layout.addStretch(1)
        return page

    def _formulas_ribbon_page(self) -> QWidget:
        page, layout = self._ribbon_page()
        formula_bar = self._ribbon_group(layout, "Formula Bar")
        self.name_box = QLineEdit()
        self.name_box.setReadOnly(True)
        self.name_box.setFixedWidth(90)
        self.formula_input = QLineEdit()
        self.formula_input.setPlaceholderText("Type a value or formula, for example =SUM(A1:A5)")
        self.formula_input.returnPressed.connect(self.commit_formula_bar)
        self.formula_input.setMinimumWidth(420)
        formula_bar.addWidget(self._ribbon_small_label("Cell"))
        formula_bar.addWidget(self.name_box)
        formula_icon = QToolButton()
        formula_icon.setObjectName("ribbonButton")
        formula_icon.setIcon(app_icon("formula"))
        formula_icon.setAutoRaise(False)
        formula_icon.setToolTip("Formula bar")
        formula_bar.addWidget(formula_icon)
        formula_bar.addWidget(self.formula_input)

        library = self._ribbon_group(layout, "Library")
        library.addWidget(self._ribbon_callback_button("Apply", "formula", self.apply_formula_template))
        library.addWidget(self._ribbon_callback_button("Insert", "formula", self.insert_formula_template))
        library.addWidget(self._ribbon_callback_button("Fill", "formula", self.fill_selection_with_custom_algorithm))
        layout.addStretch(1)
        return page

    def _data_ribbon_page(self) -> QWidget:
        page, layout = self._ribbon_page()
        import_export = self._ribbon_group(layout, "Import & Export")
        import_export.addWidget(self._ribbon_action_button(self.open_action))
        import_export.addWidget(self._ribbon_action_button(self.export_csv_action))

        number = self._ribbon_group(layout, "Number")
        self.number_format_box = QComboBox()
        self.number_format_box.addItems(["General", "0", "0.00", "#,##0", "#,##0.00", "0%", "$#,##0.00", "yyyy-mm-dd"])
        self.number_format_box.currentTextChanged.connect(lambda value: self.apply_style(number_format=value))
        number.addWidget(self._ribbon_small_label("Format"))
        number.addWidget(self.number_format_box)

        layout.addStretch(1)
        return page

    def _view_ribbon_page(self) -> QWidget:
        page, layout = self._ribbon_page()
        zoom = self._ribbon_group(layout, "Zoom")
        for action in (self.zoom_out_action, self.zoom_reset_action, self.zoom_in_action):
            zoom.addWidget(self._ribbon_action_button(action))
        self.zoom_box = QSpinBox()
        self.zoom_box.setRange(40, 220)
        self.zoom_box.setSingleStep(10)
        self.zoom_box.setSuffix("%")
        self.zoom_box.setValue(self.zoom_percent)
        self.zoom_box.valueChanged.connect(self.set_zoom_percent)
        zoom.addWidget(self._ribbon_small_label("Value"))
        zoom.addWidget(self.zoom_box)
        layout.addStretch(1)
        return page

    def _network_ribbon_page(self) -> QWidget:
        page, layout = self._ribbon_page()
        session = self._ribbon_group(layout, "Session")
        for action in (self.host_network_action, self.join_network_action, self.leave_network_action):
            session.addWidget(self._ribbon_action_button(action))
        startup = self._ribbon_group(layout, "Startup")
        startup.addWidget(self._ribbon_action_button(self.startup_network_action))
        layout.addStretch(1)
        return page

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
            QTabWidget#excelRibbon {
                background: #ffffff;
                border: 0;
            }
            QTabWidget#excelRibbon::pane {
                background: #f8fafc;
                border: 0;
                border-top: 1px solid #d0d7de;
                border-bottom: 1px solid #cfd7df;
            }
            QTabWidget#excelRibbon QTabBar::tab {
                background: #ffffff;
                color: #1f2937;
                padding: 7px 17px;
                border: 0;
                border-right: 1px solid #e5e7eb;
            }
            QTabWidget#excelRibbon QTabBar::tab:selected {
                color: #107c41;
                background: #f8fafc;
                border-bottom: 2px solid #107c41;
                font-weight: 600;
            }
            QWidget#ribbonPage {
                background: #f8fafc;
            }
            QFrame#ribbonGroup {
                background: #ffffff;
                border: 1px solid #d8dee4;
                border-radius: 7px;
            }
            QLabel#ribbonGroupTitle {
                color: #6b7280;
                font-size: 10px;
            }
            QLabel#ribbonInlineLabel {
                color: #4b5563;
                font-size: 11px;
            }
            QToolButton#ribbonButton {
                border: 1px solid transparent;
                border-radius: 6px;
                color: #1f2937;
                min-width: 44px;
                padding: 5px 7px;
                background: #ffffff;
            }
            QToolButton#ribbonButton:hover {
                background: #eef7f1;
                border-color: #a8d5ba;
            }
            QToolButton#ribbonButton:pressed,
            QToolButton#ribbonButton:checked {
                background: #dff3e8;
                border-color: #107c41;
            }
            QFrame#ribbonGroup QComboBox,
            QFrame#ribbonGroup QSpinBox,
            QFrame#ribbonGroup QLineEdit,
            QFrame#ribbonGroup QFontComboBox {
                border: 1px solid #cfd7df;
                border-radius: 6px;
                padding: 5px 7px;
                background: #ffffff;
                color: #111827;
            }
            QMenuBar {
                background: #ffffff;
                border-bottom: 1px solid #e5e7eb;
            }
            QStatusBar {
                background: #f8fafc;
            }
            QScrollArea#inspectorPanel {
                background: #f6f8fa;
                border: 0;
                border-bottom: 1px solid #d0d7de;
            }
            QWidget#inspectorContent {
                background: #f6f8fa;
            }
            QFrame#inspectorSection {
                background: #ffffff;
                border: 1px solid #d8dee4;
                border-radius: 8px;
            }
            QLabel#inspectorSectionTitle {
                color: #111827;
                font-weight: 600;
                font-size: 12px;
            }
            QLabel#inspectorValue {
                background: #f8fafc;
                border: 1px solid #e5e7eb;
                border-radius: 6px;
                color: #374151;
                padding: 7px;
            }
            QFrame#inspectorSection QComboBox,
            QFrame#inspectorSection QLineEdit {
                border: 1px solid #cfd7df;
                border-radius: 6px;
                padding: 6px 8px;
                background: #ffffff;
                color: #111827;
            }
            QToolButton#sidebarButton {
                border: 1px solid #cfd7df;
                border-radius: 6px;
                background: #ffffff;
                color: #1f2937;
                padding: 6px 8px;
            }
            QToolButton#sidebarButton:hover {
                background: #eef7f1;
                border-color: #a8d5ba;
            }
            QToolButton#sidebarButton:pressed {
                background: #dff3e8;
                border-color: #107c41;
            }
            QTreeWidget#projectTree {
                background: #ffffff;
                border: 1px solid #e5e7eb;
                border-radius: 6px;
                color: #1f2937;
                padding: 3px;
            }
            QTreeWidget#projectTree::item {
                min-height: 20px;
                padding: 2px 4px;
            }
            QTreeWidget#projectTree::item:selected {
                background: #dff3e8;
                color: #0b5f32;
            }
            """
        )

    def _format_button(self, text: str, icon_name: str, tooltip: str, callback) -> QToolButton:
        button = QToolButton()
        button.setObjectName("ribbonButton")
        button.setText(text)
        button.setIcon(app_icon(icon_name))
        button.setToolButtonStyle(Qt.ToolButtonTextUnderIcon)
        button.setCheckable(True)
        button.setToolTip(tooltip)
        button.setAutoRaise(False)
        button.toggled.connect(callback)
        return button

    def _inspector_section(self, title: str, icon_name: str) -> tuple[QFrame, QVBoxLayout]:
        frame = QFrame()
        frame.setObjectName("inspectorSection")
        outer = QVBoxLayout(frame)
        outer.setContentsMargins(12, 10, 12, 12)
        outer.setSpacing(8)

        header = QHBoxLayout()
        header.setContentsMargins(0, 0, 0, 0)
        header.setSpacing(7)
        icon = QLabel()
        icon.setObjectName("inspectorSectionIcon")
        icon.setPixmap(app_icon(icon_name, 20).pixmap(20, 20))
        icon.setFixedSize(22, 22)
        label = QLabel(title)
        label.setObjectName("inspectorSectionTitle")
        header.addWidget(icon)
        header.addWidget(label)
        header.addStretch(1)
        outer.addLayout(header)

        body = QVBoxLayout()
        body.setContentsMargins(0, 0, 0, 0)
        body.setSpacing(7)
        outer.addLayout(body)
        return frame, body

    def _sidebar_button(self, text: str, icon_name: str, callback) -> QToolButton:
        button = QToolButton()
        button.setObjectName("sidebarButton")
        button.setText(text)
        button.setIcon(app_icon(icon_name))
        button.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)
        button.setAutoRaise(False)
        button.clicked.connect(callback)
        return button

    def _sidebar_action_button(self, action: QAction) -> QToolButton:
        button = QToolButton()
        button.setObjectName("sidebarButton")
        button.setDefaultAction(action)
        button.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)
        button.setAutoRaise(False)
        return button

    def _build_inspector(self) -> QWidget:
        scroll = QScrollArea()
        scroll.setObjectName("inspectorPanel")
        scroll.setWidgetResizable(False)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setFixedHeight(258)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)

        widget = QWidget()
        widget.setObjectName("inspectorContent")
        widget.setMinimumWidth(2100)
        layout = QHBoxLayout(widget)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)

        selection_section, selection_layout = self._inspector_section("Selection", "selection")
        selection_section.setFixedWidth(210)
        self.selection_label = QLabel("No selection")
        self.selection_label.setObjectName("inspectorValue")
        self.selection_label.setWordWrap(True)
        selection_layout.addWidget(self.selection_label)
        layout.addWidget(selection_section)

        stats_section, stats_layout = self._inspector_section("Quick Stats", "stats")
        stats_section.setFixedWidth(210)
        self.stats_label = QLabel("Sum: 0\nAverage: 0\nCount: 0")
        self.stats_label.setObjectName("inspectorValue")
        self.stats_label.setWordWrap(True)
        stats_layout.addWidget(self.stats_label)
        layout.addWidget(stats_section)

        network_section, network_layout = self._inspector_section("Network", "network_host")
        network_section.setFixedWidth(300)
        self.network_status_label = QLabel("Offline")
        self.network_status_label.setObjectName("inspectorValue")
        self.network_status_label.setWordWrap(True)
        network_buttons = QHBoxLayout()
        network_buttons.setContentsMargins(0, 0, 0, 0)
        network_buttons.setSpacing(6)
        for action in (self.host_network_action, self.join_network_action, self.leave_network_action, self.startup_network_action):
            network_buttons.addWidget(self._sidebar_action_button(action))
        network_layout.addWidget(self.network_status_label)
        network_layout.addLayout(network_buttons)
        layout.addWidget(network_section)

        project_section, project_layout = self._inspector_section("Project", "project")
        project_section.setFixedWidth(380)
        self.project_summary_label = QLabel("No project")
        self.project_summary_label.setObjectName("inspectorValue")
        self.project_summary_label.setWordWrap(True)
        self.project_tree = QTreeWidget()
        self.project_tree.setObjectName("projectTree")
        self.project_tree.setHeaderHidden(True)
        self.project_tree.setFixedHeight(92)
        self.project_tree.itemSelectionChanged.connect(self.update_project_actions)
        self.project_tree.itemDoubleClicked.connect(lambda _item, _column: self.open_selected_project_file())
        project_buttons = QHBoxLayout()
        project_buttons.setContentsMargins(0, 0, 0, 0)
        project_buttons.setSpacing(6)
        for action in (self.open_project_action, self.refresh_project_action, self.open_project_file_action, self.share_project_action):
            project_buttons.addWidget(self._sidebar_action_button(action))
        project_layout.addWidget(self.project_summary_label)
        project_layout.addWidget(self.project_tree)
        project_layout.addLayout(project_buttons)
        layout.addWidget(project_section)

        formula_section, formula_layout = self._inspector_section("Formula Library", "formula")
        formula_section.setFixedWidth(300)
        self.formula_category_box = QComboBox()
        self.formula_category_box.addItems(FORMULA_LIBRARY.keys())
        self.formula_category_box.currentTextChanged.connect(self.reload_formula_templates)
        self.formula_template_box = QComboBox()
        self.formula_template_box.currentIndexChanged.connect(self.update_formula_preview)
        self.formula_preview = QLineEdit()
        self.formula_preview.setReadOnly(True)
        formula_buttons = QHBoxLayout()
        formula_buttons.setContentsMargins(0, 0, 0, 0)
        formula_buttons.setSpacing(6)
        self.apply_template_button = QToolButton()
        self.apply_template_button.setObjectName("sidebarButton")
        self.apply_template_button.setText("Apply")
        self.apply_template_button.setIcon(app_icon("formula"))
        self.apply_template_button.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)
        self.apply_template_button.clicked.connect(self.apply_formula_template)
        self.insert_template_button = QToolButton()
        self.insert_template_button.setObjectName("sidebarButton")
        self.insert_template_button.setText("Insert")
        self.insert_template_button.setIcon(app_icon("formula"))
        self.insert_template_button.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)
        self.insert_template_button.clicked.connect(self.insert_formula_template)
        formula_buttons.addWidget(self.apply_template_button)
        formula_buttons.addWidget(self.insert_template_button)
        for control in [self.formula_category_box, self.formula_template_box, self.formula_preview]:
            formula_layout.addWidget(control)
        formula_layout.addLayout(formula_buttons)
        layout.addWidget(formula_section)

        algorithm_section, algorithm_layout = self._inspector_section("Cell Algorithm", "formula")
        algorithm_section.setFixedWidth(250)
        self.custom_algorithm_input = QLineEdit()
        self.custom_algorithm_input.setPlaceholderText("=A{row}*2")
        custom_buttons = QHBoxLayout()
        custom_buttons.setContentsMargins(0, 0, 0, 0)
        custom_buttons.setSpacing(6)
        self.apply_custom_button = QToolButton()
        self.apply_custom_button.setObjectName("sidebarButton")
        self.apply_custom_button.setText("Apply")
        self.apply_custom_button.setIcon(app_icon("formula"))
        self.apply_custom_button.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)
        self.apply_custom_button.clicked.connect(self.apply_custom_algorithm)
        self.fill_custom_button = QToolButton()
        self.fill_custom_button.setObjectName("sidebarButton")
        self.fill_custom_button.setText("Fill")
        self.fill_custom_button.setIcon(app_icon("formula"))
        self.fill_custom_button.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)
        self.fill_custom_button.clicked.connect(self.fill_selection_with_custom_algorithm)
        custom_buttons.addWidget(self.apply_custom_button)
        custom_buttons.addWidget(self.fill_custom_button)
        algorithm_layout.addWidget(self.custom_algorithm_input)
        algorithm_layout.addLayout(custom_buttons)
        layout.addWidget(algorithm_section)

        self.reload_formula_templates()

        chart_section, chart_layout = self._inspector_section("Charts", "chart")
        chart_section.setFixedWidth(390)
        self.chart_type_box = QComboBox()
        self.chart_type_box.addItems(["Bar", "Line", "Pie"])
        self.chart_title_input = QLineEdit()
        self.chart_title_input.setPlaceholderText("Chart title")
        self.create_chart_button = QToolButton()
        self.create_chart_button.setObjectName("sidebarButton")
        self.create_chart_button.setText("Create")
        self.create_chart_button.setIcon(app_icon("chart"))
        self.create_chart_button.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)
        self.create_chart_button.clicked.connect(self.create_chart_from_selection)
        self.chart_widget = ChartWidget()
        self.chart_widget.setMinimumHeight(112)
        self.chart_widget.setMaximumHeight(112)
        chart_layout.addWidget(self.chart_type_box)
        chart_layout.addWidget(self.chart_title_input)
        chart_layout.addWidget(self.create_chart_button)
        chart_layout.addWidget(self.chart_widget)
        layout.addWidget(chart_section)

        layout.addStretch(1)
        scroll.setWidget(widget)
        self.update_project_panel()
        return scroll

    def open_project(self) -> None:
        path = QFileDialog.getExistingDirectory(self, "Open project folder")
        if not path:
            return
        self.load_project(Path(path))
        self.send_project_snapshot()

    def load_project(self, path: Path) -> None:
        try:
            self.project = scan_project_folder(path)
        except OSError as exc:
            QMessageBox.critical(self, "Open project failed", str(exc))
            return
        self.update_project_panel()
        self.statusBar().showMessage(f"Project opened: {self.project.name}")

    def refresh_project(self) -> None:
        if not self.project.root_path or self.project.remote:
            self.statusBar().showMessage("No local project folder to refresh")
            return
        self.load_project(Path(self.project.root_path))
        self.send_project_snapshot()

    def close_project(self) -> None:
        self.project = ProjectData()
        self.update_project_panel()
        self.send_project_snapshot()
        self.statusBar().showMessage("Project closed")

    def send_project_snapshot(self) -> None:
        if self.applying_remote_update or self.collaboration is None or not self.collaboration.running:
            return
        self.send_collaboration_message(project_snapshot_message(self.project))

    def update_project_panel(self) -> None:
        if not hasattr(self, "project_tree"):
            return
        self.project_tree.clear()
        if not self.project.is_open:
            self.project_summary_label.setText("No project")
            self.update_project_actions()
            return

        location = "Remote team project" if self.project.remote else self.project.root_path or "Local project"
        self.project_summary_label.setText(
            f"{self.project.name}\n{self.project.total_entries} items\n{self.project.openable_count} spreadsheets\n{location}"
        )
        root_item = QTreeWidgetItem([self.project.name])
        root_item.setIcon(0, app_icon("project"))
        root_item.setData(0, Qt.UserRole, "")
        root_item.setData(0, Qt.UserRole + 1, "folder")
        self.project_tree.addTopLevelItem(root_item)
        folder_items: dict[str, QTreeWidgetItem] = {"": root_item}

        for folder in self.project.folders:
            folder_items[folder] = self._ensure_project_folder_item(folder, folder_items)
        for file_item in self.project.files:
            parent_path = PurePosixPath(file_item.relative_path).parent.as_posix()
            if parent_path == ".":
                parent_path = ""
            parent_item = self._ensure_project_folder_item(parent_path, folder_items)
            item = QTreeWidgetItem([file_item.name])
            icon_name = "csv" if file_item.kind == "csv" else "sheet_add" if file_item.openable else "open"
            item.setIcon(0, app_icon(icon_name))
            item.setData(0, Qt.UserRole, file_item.relative_path)
            item.setData(0, Qt.UserRole + 1, "file")
            parent_item.addChild(item)
        self.project_tree.expandToDepth(1)
        self.update_project_actions()

    def _ensure_project_folder_item(
        self, folder_path: str, folder_items: dict[str, QTreeWidgetItem]
    ) -> QTreeWidgetItem:
        folder_path = "" if folder_path == "." else folder_path
        if folder_path in folder_items:
            return folder_items[folder_path]
        parent_path = PurePosixPath(folder_path).parent.as_posix()
        if parent_path == ".":
            parent_path = ""
        parent_item = self._ensure_project_folder_item(parent_path, folder_items)
        item = QTreeWidgetItem([PurePosixPath(folder_path).name])
        item.setIcon(0, app_icon("project"))
        item.setData(0, Qt.UserRole, folder_path)
        item.setData(0, Qt.UserRole + 1, "folder")
        parent_item.addChild(item)
        folder_items[folder_path] = item
        return item

    def selected_project_file(self):
        if not hasattr(self, "project_tree"):
            return None
        item = self.project_tree.currentItem()
        if item is None or item.data(0, Qt.UserRole + 1) != "file":
            return None
        return self.project.file_by_relative_path(str(item.data(0, Qt.UserRole) or ""))

    def open_selected_project_file(self) -> None:
        project_file = self.selected_project_file()
        if project_file is None:
            self.statusBar().showMessage("Select a spreadsheet file from the project")
            return
        if not project_file.openable:
            self.statusBar().showMessage("Only XLSX and CSV project files can be opened in the spreadsheet")
            return
        path = self.project.absolute_path_for(project_file)
        if path is None:
            QMessageBox.information(self, "Project file", "This project snapshot came from the team server. Open the matching local project folder to load files from disk.")
            return
        try:
            workbook = load_xlsx(path) if project_file.extension == ".xlsx" else load_csv(path)
            self.load_workbook(workbook)
            self.send_collaboration_snapshot()
            self.statusBar().showMessage(f"Opened project file: {project_file.relative_path}")
        except Exception as exc:
            QMessageBox.critical(self, "Open project file failed", str(exc))

    def update_project_actions(self) -> None:
        if not hasattr(self, "open_project_file_action"):
            return
        local_project = self.project.is_open and not self.project.remote
        selected_file = self.selected_project_file() if hasattr(self, "project_tree") else None
        self.refresh_project_action.setEnabled(bool(local_project and self.project.root_path))
        self.close_project_action.setEnabled(self.project.is_open)
        self.open_project_file_action.setEnabled(bool(selected_file and selected_file.openable and local_project))
        self.share_project_action.setEnabled(self.collaboration is not None and self.collaboration.running)

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
        if self.applying_remote_update or self.collaboration is None or not self.collaboration.running:
            return
        self.send_collaboration_message(self.build_collaboration_snapshot_message())

    def build_collaboration_snapshot_message(self) -> dict:
        return {
            "type": "snapshot",
            "workbook": workbook_to_payload(self.workbook),
            "project": project_to_payload(self.project),
        }

    def apply_startup_settings(self, *, restart: bool = False) -> None:
        settings = self.startup_settings.normalized()
        self.startup_settings = settings
        if settings.startup_mode == "manual":
            if restart and self.collaboration is not None:
                self.leave_collaboration()
            return
        if self.collaboration is not None:
            if not restart:
                return
            self.leave_collaboration()
        if settings.startup_mode == "local_server":
            self.start_host_collaboration(settings.local_server_port, notify_user=False)
        elif settings.startup_mode == "shared_client":
            self.start_client_collaboration(settings.shared_server_host, settings.shared_server_port)

    def open_startup_settings(self) -> None:
        dialog = QDialog(self)
        dialog.setWindowTitle("Network startup settings")
        layout = QFormLayout(dialog)

        mode_box = QComboBox()
        mode_items = [
            ("Manual network controls", "manual"),
            ("Connect to shared server", "shared_client"),
            ("Run this program as server", "local_server"),
        ]
        for label, value in mode_items:
            mode_box.addItem(label, value)
        current_mode = self.startup_settings.normalized().startup_mode
        mode_index = mode_box.findData(current_mode)
        mode_box.setCurrentIndex(max(0, mode_index))

        host_input = QLineEdit(self.startup_settings.shared_server_host)
        shared_port_box = QSpinBox()
        shared_port_box.setRange(1, 65535)
        shared_port_box.setValue(self.startup_settings.shared_server_port)
        local_port_box = QSpinBox()
        local_port_box.setRange(1, 65535)
        local_port_box.setValue(self.startup_settings.local_server_port)

        def update_fields() -> None:
            mode = mode_box.currentData()
            host_input.setEnabled(mode == "shared_client")
            shared_port_box.setEnabled(mode == "shared_client")
            local_port_box.setEnabled(mode == "local_server")

        mode_box.currentIndexChanged.connect(update_fields)
        update_fields()

        layout.addRow("Startup mode", mode_box)
        layout.addRow("Shared server IP", host_input)
        layout.addRow("Shared server port", shared_port_box)
        layout.addRow("Local server port", local_port_box)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Save | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addRow(buttons)

        if dialog.exec() != QDialog.Accepted:
            return
        self.startup_settings = StartupSettings(
            startup_mode=str(mode_box.currentData()),
            shared_server_host=host_input.text(),
            shared_server_port=shared_port_box.value(),
            local_server_port=local_port_box.value(),
        ).normalized()
        save_startup_settings(self.startup_settings)
        self.statusBar().showMessage("Network startup settings saved")
        self.apply_startup_settings(restart=True)

    def host_collaboration(self) -> None:
        port, ok = QInputDialog.getInt(self, "Host collaboration", "Port", DEFAULT_PORT, 1, 65535)
        if not ok:
            return
        self.start_host_collaboration(port)

    def start_host_collaboration(self, port: int, *, notify_user: bool = True) -> bool:
        if self.collaboration is not None:
            self.leave_collaboration()
        server = CollaborationServer(
            port=port,
            snapshot_message_provider=self.build_collaboration_snapshot_message,
        )
        self.attach_collaboration(server, "Host")
        try:
            server.start()
        except OSError as exc:
            server.stop()
            self.detach_collaboration()
            if notify_user:
                QMessageBox.critical(self, "Network failed", str(exc))
            else:
                self.show_collaboration_error(f"Server startup failed: {exc}")
            return False
        return True

    def join_collaboration(self) -> None:
        target, ok = QInputDialog.getText(self, "Join collaboration", "Server IP:port", text=f"127.0.0.1:{DEFAULT_PORT}")
        if not ok or not target.strip():
            return
        try:
            host, port = self.parse_collaboration_target(target.strip())
        except ValueError as exc:
            QMessageBox.warning(self, "Join collaboration", str(exc))
            return
        self.start_client_collaboration(host, port)

    def start_client_collaboration(self, host: str, port: int) -> bool:
        if self.collaboration is not None:
            self.leave_collaboration()
        client = CollaborationClient(host, port)
        self.attach_collaboration(client, "Client")
        client.start()
        return True

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
        self.update_project_actions()

    def show_collaboration_error(self, message: str) -> None:
        self.update_collaboration_status(message)
        if self.statusBar():
            self.statusBar().showMessage(message)

    def on_collaboration_message(self, message: dict) -> None:
        message_type = message.get("type")
        try:
            if message_type == "snapshot":
                self.apply_remote_snapshot(message)
            elif message_type == "project_snapshot":
                self.apply_remote_project_snapshot(message)
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
            if "project" in message:
                self.apply_remote_project_snapshot(message)
        finally:
            self.applying_remote_update = False
        self.update_formula_bar()
        self.update_selection_stats()
        self.statusBar().showMessage("Workbook synchronized")

    def apply_remote_project_snapshot(self, message: dict) -> None:
        self.project = project_from_payload(message.get("project"), remote=True)
        self.update_project_panel()
        self.statusBar().showMessage("Project synchronized")

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
