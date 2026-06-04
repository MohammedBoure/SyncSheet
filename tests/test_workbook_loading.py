import os
import tempfile
import time
import unittest
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication

from pyexcel_lite.io_xlsx import save_xlsx
from pyexcel_lite.main import SpreadsheetWindow
from pyexcel_lite.workbook import WorkbookData


class WorkbookLoadingTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def wait_until(self, condition, timeout: float = 5.0) -> None:
        deadline = time.monotonic() + timeout
        while time.monotonic() < deadline:
            self.app.processEvents()
            if condition():
                return
            time.sleep(0.01)
        self.fail("condition was not reached before timeout")

    def test_workbook_loader_opens_file_without_blocking_window(self):
        workbook = WorkbookData()
        workbook.active_sheet.set_value(0, 0, "loaded")
        window = SpreadsheetWindow()
        try:
            with tempfile.TemporaryDirectory() as folder:
                path = Path(folder) / "load-me.xlsx"
                save_xlsx(workbook, path)

                window.start_workbook_load(path, send_snapshot=False)

                self.assertIsNotNone(window.workbook_load_progress)
                self.wait_until(lambda: window.workbook_load_thread is None)
                self.assertEqual(window.current_sheet.raw_value(0, 0), "loaded")
                self.assertIsNone(window.workbook_load_progress)
        finally:
            window.close()


if __name__ == "__main__":
    unittest.main()
