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

    def test_batch_set_values_bumps_revision_once_and_expands_model(self):
        self.model.set_values([(250, 60, "42"), (251, 60, "43"), (252, 60, "=SUM(BI251:BI252)")])

        self.assertEqual(self.sheet.revision, 1)
        self.assertGreaterEqual(self.model.rowCount(), 253)
        self.assertGreaterEqual(self.model.columnCount(), 61)
        self.assertEqual(self.model.data(self.model.index(252, 60), Qt.DisplayRole), "85")

    def test_formula_results_are_cached_until_sheet_changes(self):
        self.model.set_values([(0, 0, "10"), (0, 1, "20"), (0, 2, "=SUM(A1:B1)")])
        formula_index = self.model.index(0, 2)

        self.assertEqual(self.model.data(formula_index, Qt.DisplayRole), "30")
        self.assertIn((id(self.sheet), self.sheet.revision, 0, 2), self.model.evaluator._formula_cache)

        self.model.setData(self.model.index(0, 0), "15")
        self.assertEqual(self.model.data(formula_index, Qt.DisplayRole), "35")
        self.assertIn((id(self.sheet), self.sheet.revision, 0, 2), self.model.evaluator._formula_cache)

    def test_formula_index_tracks_formula_cells_without_scanning_all_cells(self):
        self.model.set_values([(0, 0, "=SUM(B1:C1)"), (0, 1, "1"), (0, 2, "2")])
        self.assertEqual(self.sheet.formula_cells, {(0, 0)})

        self.model.setData(self.model.index(0, 0), "")
        self.assertEqual(self.sheet.formula_cells, set())

        self.model.set_values([(1, 0, "=SUM(B2:C2)")])
        self.model.insert_rows(0)
        self.assertEqual(self.sheet.formula_cells, {(2, 0)})


if __name__ == "__main__":
    unittest.main()
