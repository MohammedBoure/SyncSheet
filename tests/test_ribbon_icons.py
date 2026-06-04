import os
import unittest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication, QFrame, QLabel, QScrollArea, QSplitter, QTabWidget, QToolButton

from pyexcel_lite.main import SpreadsheetWindow


class RibbonIconTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def test_primary_actions_and_format_buttons_have_icons(self):
        window = SpreadsheetWindow()
        try:
            actions = [
                window.new_action,
                window.open_action,
                window.save_action,
                window.save_as_action,
                window.export_csv_action,
                window.open_project_action,
                window.refresh_project_action,
                window.open_project_file_action,
                window.share_project_action,
                window.close_project_action,
                window.add_sheet_action,
                window.insert_row_action,
                window.delete_row_action,
                window.insert_column_action,
                window.delete_column_action,
                window.clear_action,
                window.undo_action,
                window.redo_action,
                window.chart_action,
                window.host_network_action,
                window.join_network_action,
                window.leave_network_action,
                window.zoom_in_action,
                window.zoom_out_action,
                window.zoom_reset_action,
            ]
            self.assertTrue(all(not action.icon().isNull() for action in actions))
            self.assertFalse(window.bold_button.icon().isNull())
            self.assertFalse(window.italic_button.icon().isNull())
            self.assertFalse(window.underline_button.icon().isNull())
            self.assertFalse(window.text_color_button.icon().isNull())
            self.assertFalse(window.fill_color_button.icon().isNull())
        finally:
            window.close()

    def test_top_ribbon_uses_excel_like_tabs(self):
        window = SpreadsheetWindow()
        try:
            self.assertIsInstance(window.ribbon, QTabWidget)
            self.assertEqual(window.ribbon.objectName(), "excelRibbon")
            self.assertLessEqual(window.ribbon.maximumHeight(), 142)
            self.assertEqual(
                [window.ribbon.tabText(index) for index in range(window.ribbon.count())],
                ["Home", "Project", "Insert", "Formulas", "Data", "View", "Network"],
            )
            ribbon_buttons = window.ribbon.findChildren(QToolButton, "ribbonButton")
            self.assertGreaterEqual(len(ribbon_buttons), 20)
            self.assertTrue(all(not button.icon().isNull() for button in ribbon_buttons if button.defaultAction() or button.text()))
        finally:
            window.close()

    def test_layout_uses_formula_bar_and_right_task_pane(self):
        window = SpreadsheetWindow()
        try:
            central_layout = window.centralWidget().layout()
            self.assertEqual(central_layout.itemAt(0).widget().objectName(), "excelRibbon")
            self.assertEqual(central_layout.itemAt(1).widget().objectName(), "formulaBar")

            workspace = central_layout.itemAt(2).widget()
            self.assertIsInstance(workspace, QSplitter)
            self.assertEqual(workspace.objectName(), "sheetWorkspace")
            self.assertIs(workspace.widget(0), window.tabs)
            self.assertIs(workspace.widget(1), window.inspector)

            self.assertIsInstance(window.formula_bar, QFrame)
            self.assertEqual(window.formula_bar.objectName(), "formulaBar")
            self.assertEqual(window.name_box.objectName(), "nameBox")
            self.assertEqual(window.formula_input.objectName(), "formulaInput")
        finally:
            window.close()

    def test_task_pane_keeps_section_cards_and_icons(self):
        window = SpreadsheetWindow()
        try:
            self.assertIsInstance(window.inspector, QScrollArea)
            self.assertEqual(window.inspector.objectName(), "inspectorPanel")
            self.assertTrue(window.inspector.widgetResizable())
            self.assertLessEqual(window.inspector.maximumWidth(), 380)
            self.assertEqual(window.inspector.horizontalScrollBarPolicy(), Qt.ScrollBarAlwaysOff)
            self.assertEqual(window.inspector.verticalScrollBarPolicy(), Qt.ScrollBarAsNeeded)

            sections = window.inspector.findChildren(QFrame, "inspectorSection")
            titles = [label.text() for label in window.inspector.findChildren(QLabel, "inspectorSectionTitle")]
            section_icons = window.inspector.findChildren(QLabel, "inspectorSectionIcon")
            sidebar_buttons = window.inspector.findChildren(QToolButton, "sidebarButton")

            self.assertGreaterEqual(len(sections), 7)
            self.assertTrue({"Selection", "Quick Stats", "Network", "Project", "Formula Library", "Cell Algorithm", "Charts"}.issubset(set(titles)))
            self.assertGreaterEqual(len(section_icons), 7)
            self.assertTrue(all(label.pixmap() is not None and not label.pixmap().isNull() for label in section_icons))
            self.assertGreaterEqual(len(sidebar_buttons), 8)
            self.assertTrue(all(not button.icon().isNull() for button in sidebar_buttons))
        finally:
            window.close()


if __name__ == "__main__":
    unittest.main()
