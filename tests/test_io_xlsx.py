import tempfile
import unittest
from pathlib import Path

from pyexcel_lite.io_xlsx import load_csv, load_xlsx, save_xlsx
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

    def test_load_csv_as_workbook(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "sample.csv"
            path.write_text("Name,Value\nGold,24\n", encoding="utf-8-sig")

            loaded = load_csv(path)

        self.assertEqual(loaded.active_sheet.name, "sample")
        self.assertEqual(loaded.active_sheet.raw_value(0, 0), "Name")
        self.assertEqual(loaded.active_sheet.raw_value(1, 1), "24")


if __name__ == "__main__":
    unittest.main()
