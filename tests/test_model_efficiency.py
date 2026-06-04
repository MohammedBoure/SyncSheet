import os
import unittest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication

from pyexcel_lite.formula import FormulaEvaluator
from pyexcel_lite.qt_model import WorksheetTableModel
from pyexcel_lite.workbook import WorkbookData


class ModelEfficiencyTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def setUp(self):
        self.workbook = WorkbookData()
        self.sheet = self.workbook.active_sheet
        self.model = WorksheetTableModel(self.sheet, FormulaEvaluator(self.workbook))

    def test_reading_blank_cells_does_not_create_cell_objects(self):
        for row in range(50):
            for column in range(20):
                index = self.model.index(row, column)
                self.assertEqual(self.model.data(index, Qt.DisplayRole), "")
                self.assertEqual(self.model.data(index, Qt.EditRole), "")
        self.assertEqual(self.sheet.cells, {})

    def test_batch_set_values_writes_cells_once(self):
        self.model.set_values([(0, 0, "1"), (0, 1, "2"), (1, 0, "=SUM(A1:B1)")])

        self.assertEqual(len(self.sheet.cells), 3)
        self.assertEqual(self.model.data(self.model.index(1, 0), Qt.DisplayRole), "3")


if __name__ == "__main__":
    unittest.main()
