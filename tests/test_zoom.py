import os
import unittest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication

from pyexcel_lite.main import SpreadsheetWindow


class ZoomTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def test_zoom_changes_model_and_grid_sizes(self):
        window = SpreadsheetWindow()
        try:
            base_column_width = window.current_view.horizontalHeader().defaultSectionSize()
            window.zoom_in()
            self.assertEqual(window.zoom_percent, 110)
            self.assertGreater(window.current_view.horizontalHeader().defaultSectionSize(), base_column_width)
            self.assertEqual(window.current_model.zoom_factor, 1.1)

            window.reset_zoom()
            self.assertEqual(window.zoom_percent, 100)
            self.assertEqual(window.current_model.zoom_factor, 1.0)
        finally:
            window.close()


if __name__ == "__main__":
    unittest.main()
