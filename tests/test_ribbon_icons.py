import os
import unittest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication, QFrame, QLabel, QScrollArea, QToolBar, QToolButton

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

    def test_top_toolbars_use_excel_like_icon_layout(self):
        window = SpreadsheetWindow()
        try:
            toolbars = window.findChildren(QToolBar)
            icon_toolbars = [toolbar for toolbar in toolbars if toolbar.windowTitle() in {"File", "Format"}]
            self.assertGreaterEqual(len(icon_toolbars), 2)
            for toolbar in icon_toolbars:
                self.assertEqual(toolbar.toolButtonStyle(), Qt.ToolButtonTextUnderIcon)
        finally:
            window.close()

    def test_sidebar_uses_section_cards_and_icons(self):
        window = SpreadsheetWindow()
        try:
            self.assertIsInstance(window.inspector, QScrollArea)
            self.assertEqual(window.inspector.objectName(), "inspectorPanel")

            sections = window.inspector.findChildren(QFrame, "inspectorSection")
            titles = [label.text() for label in window.inspector.findChildren(QLabel, "inspectorSectionTitle")]
            section_icons = window.inspector.findChildren(QLabel, "inspectorSectionIcon")
            sidebar_buttons = window.inspector.findChildren(QToolButton, "sidebarButton")

            self.assertGreaterEqual(len(sections), 6)
            self.assertTrue({"Selection", "Quick Stats", "Network", "Formula Library", "Cell Algorithm", "Charts"}.issubset(set(titles)))
            self.assertGreaterEqual(len(section_icons), 6)
            self.assertTrue(all(label.pixmap() is not None and not label.pixmap().isNull() for label in section_icons))
            self.assertGreaterEqual(len(sidebar_buttons), 8)
            self.assertTrue(all(not button.icon().isNull() for button in sidebar_buttons))
        finally:
            window.close()


if __name__ == "__main__":
    unittest.main()
