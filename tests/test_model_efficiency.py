import os
import unittest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import QItemSelection, Qt
from PySide6.QtWidgets import QApplication

from pyexcel_lite.formula import FormulaEvaluator
from pyexcel_lite.main import SpreadsheetWindow
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

    def test_clear_large_range_only_touches_used_cells(self):
        self.model.set_values([(0, 0, "keep"), (20, 20, "clear"), (40, 40, "clear")])
        selection = QItemSelection(self.model.index(10, 10), self.model.index(120, 50))

        self.model.clear_ranges(list(selection))

        self.assertEqual(self.sheet.raw_value(0, 0), "keep")
        self.assertEqual(self.sheet.raw_value(20, 20), "")
        self.assertEqual(self.sheet.raw_value(40, 40), "")
        self.assertEqual(list(self.sheet.cells), [(0, 0)])

    def test_zoom_factor_update_does_not_emit_full_model_refresh(self):
        emissions = []
        self.model.dataChanged.connect(
            lambda top_left, bottom_right, roles: emissions.append((top_left.row(), bottom_right.row(), list(roles)))
        )

        self.model.set_zoom_factor(1.25)

        self.assertEqual(self.model.zoom_factor, 1.25)
        self.assertEqual(emissions, [])

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

    def test_local_edit_does_not_refresh_current_sheet_formulas_twice(self):
        window = SpreadsheetWindow()
        try:
            model = window.current_model
            model.set_values([(0, 0, "10"), (0, 1, "=A1*2")], notify_change=False)
            refresh_count = 0
            original_refresh = model.refresh_formulas

            def counted_refresh():
                nonlocal refresh_count
                refresh_count += 1
                return original_refresh()

            model.refresh_formulas = counted_refresh

            model.setData(model.index(0, 0), "15")

            self.assertEqual(refresh_count, 1)
            self.assertEqual(model.data(model.index(0, 1), Qt.DisplayRole), "30")
        finally:
            window.close()


if __name__ == "__main__":
    unittest.main()
