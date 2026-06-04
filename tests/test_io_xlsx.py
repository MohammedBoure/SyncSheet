import tempfile
import unittest
from pathlib import Path

from pyexcel_lite.io_xlsx import load_xlsx, save_xlsx
from pyexcel_lite.workbook import WorkbookData


class XlsxIoTest(unittest.TestCase):
    def test_save_and_load_values_and_formulas(self):
        workbook = WorkbookData()
        sheet = workbook.active_sheet
        sheet.set_value(0, 0, "12")
        sheet.set_value(0, 1, "=A1*2")
        sheet.set_style(0, 1, bold=True, fill_color="#fff2cc")

        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "sample.xlsx"
            save_xlsx(workbook, path)
            loaded = load_xlsx(path)

        loaded_sheet = loaded.active_sheet
        self.assertEqual(loaded_sheet.raw_value(0, 0), "12")
        self.assertEqual(loaded_sheet.raw_value(0, 1), "=A1*2")
        self.assertTrue(loaded_sheet.get_cell(0, 1).style.bold)


if __name__ == "__main__":
    unittest.main()
