"""Application stylesheet for the Excel-like PySide interface."""

APP_STYLESHEET = """
QMainWindow,
QWidget#mainSurface {
    background: #ffffff;
    color: #111827;
}
QMenuBar {
    background: #ffffff;
    border-bottom: 1px solid #d0d7de;
    padding: 0;
}
QMenuBar::item {
    padding: 4px 9px;
    background: transparent;
}
QMenuBar::item:selected {
    background: #e7f3ea;
    color: #0b5f32;
}
QTabWidget#excelRibbon {
    background: #ffffff;
    border: 0;
}
QTabWidget#excelRibbon::pane {
    background: #f5f6f7;
    border: 0;
    border-top: 1px solid #d0d7de;
    border-bottom: 1px solid #c8d0d8;
}
QTabWidget#excelRibbon QTabBar::tab {
    background: #ffffff;
    color: #202124;
    padding: 7px 18px;
    border: 0;
    border-right: 1px solid #e5e7eb;
}
QTabWidget#excelRibbon QTabBar::tab:selected {
    color: #107c41;
    background: #f5f6f7;
    border-bottom: 2px solid #107c41;
    font-weight: 600;
}
QWidget#ribbonPage {
    background: #f5f6f7;
}
QFrame#ribbonGroup {
    background: #ffffff;
    border: 1px solid #d8dee4;
    border-radius: 4px;
}
QLabel#ribbonGroupTitle {
    color: #5f6b76;
    font-size: 10px;
}
QLabel#ribbonInlineLabel {
    color: #4b5563;
    font-size: 11px;
}
QToolButton#ribbonButton {
    border: 1px solid transparent;
    border-radius: 4px;
    color: #202124;
    min-width: 42px;
    padding: 4px 6px;
    background: #ffffff;
}
QToolButton#ribbonButton:hover {
    background: #e7f3ea;
    border-color: #9ed0b4;
}
QToolButton#ribbonButton:pressed,
QToolButton#ribbonButton:checked {
    background: #d8eadf;
    border-color: #107c41;
}
QFrame#ribbonGroup QComboBox,
QFrame#ribbonGroup QSpinBox,
QFrame#ribbonGroup QLineEdit,
QFrame#ribbonGroup QFontComboBox {
    border: 1px solid #c8d0d8;
    border-radius: 3px;
    padding: 4px 6px;
    background: #ffffff;
    color: #111827;
}
QFrame#formulaBar {
    background: #ffffff;
    border-bottom: 1px solid #d0d7de;
}
QLineEdit#nameBox,
QLineEdit#formulaInput {
    border: 1px solid #c8d0d8;
    border-radius: 2px;
    padding: 5px 7px;
    background: #ffffff;
    color: #111827;
}
QToolButton#formulaBarIcon {
    border: 0;
    min-width: 26px;
    padding: 3px;
}
QSplitter#sheetWorkspace::handle {
    background: #d0d7de;
    width: 1px;
}
QTabWidget::pane {
    border: 0;
}
QTableView {
    background: #ffffff;
    alternate-background-color: #f7f7f7;
    gridline-color: #d9d9d9;
    selection-background-color: #d8eadf;
    selection-color: #111827;
}
QHeaderView::section {
    background: #f3f3f3;
    color: #111827;
    border: 0;
    border-right: 1px solid #d0d7de;
    border-bottom: 1px solid #d0d7de;
    padding: 4px;
}
QScrollArea#inspectorPanel {
    background: #f6f8fa;
    border: 0;
    border-left: 1px solid #d0d7de;
}
QWidget#inspectorContent {
    background: #f6f8fa;
}
QFrame#inspectorSection {
    background: #ffffff;
    border: 1px solid #d8dee4;
    border-radius: 4px;
}
QLabel#inspectorSectionTitle {
    color: #111827;
    font-weight: 600;
    font-size: 12px;
}
QLabel#inspectorValue {
    background: #f8fafc;
    border: 1px solid #e5e7eb;
    border-radius: 3px;
    color: #374151;
    padding: 7px;
}
QFrame#inspectorSection QComboBox,
QFrame#inspectorSection QLineEdit {
    border: 1px solid #c8d0d8;
    border-radius: 3px;
    padding: 6px 8px;
    background: #ffffff;
    color: #111827;
}
QToolButton#sidebarButton {
    border: 1px solid #c8d0d8;
    border-radius: 3px;
    background: #ffffff;
    color: #202124;
    padding: 6px 8px;
}
QToolButton#sidebarButton:hover {
    background: #e7f3ea;
    border-color: #9ed0b4;
}
QToolButton#sidebarButton:pressed {
    background: #d8eadf;
    border-color: #107c41;
}
QTreeWidget#projectTree {
    background: #ffffff;
    border: 1px solid #e5e7eb;
    border-radius: 3px;
    color: #1f2937;
    padding: 3px;
}
QTreeWidget#projectTree::item {
    min-height: 20px;
    padding: 2px 4px;
}
QTreeWidget#projectTree::item:selected {
    background: #d8eadf;
    color: #0b5f32;
}
QMenu#cellContextMenu,
QMenu#cellContextFormulasMenu {
    background: #ffffff;
    border: 1px solid #d0d7de;
    padding: 5px;
    color: #1f2937;
}
QMenu#cellContextMenu::item,
QMenu#cellContextFormulasMenu::item {
    padding: 7px 30px 7px 24px;
    border-radius: 3px;
}
QMenu#cellContextMenu::item:selected,
QMenu#cellContextFormulasMenu::item:selected {
    background: #d8eadf;
    color: #0b5f32;
}
QMenu#cellContextMenu::separator,
QMenu#cellContextFormulasMenu::separator {
    height: 1px;
    background: #e5e7eb;
    margin: 5px 4px;
}
QStatusBar {
    background: #f6f8fa;
    border-top: 1px solid #d0d7de;
}
"""
