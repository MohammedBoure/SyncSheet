import os
import unittest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import QItemSelection, QItemSelectionModel, Qt
from PySide6.QtWidgets import QApplication

from pyexcel_lite.main import SpreadsheetWindow


class FormulaLibraryTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def test_template_applies_selected_range_to_current_cell(self):
        window = SpreadsheetWindow()
        try:
            model = window.current_model
            model.set_values([(0, 0, "10"), (1, 0, "20"), (2, 0, "30")])
            selection = QItemSelection(model.index(0, 0), model.index(2, 0))
            window.current_view.selectionModel().select(selection, QItemSelectionModel.ClearAndSelect)
            window.current_view.selectionModel().setCurrentIndex(model.index(0, 1), QItemSelectionModel.NoUpdate)

            window.formula_category_box.setCurrentText("Statistics")
            window.formula_template_box.setCurrentIndex(window.formula_template_box.findText("Average selected"))
            window.apply_formula_template()

            self.assertEqual(model.data(model.index(0, 1), Qt.EditRole), "=AVERAGE(A1:A3)")
            self.assertEqual(model.data(model.index(0, 1), Qt.DisplayRole), "20")
        finally:
            window.close()

    def test_custom_algorithm_can_fill_each_selected_cell(self):
        window = SpreadsheetWindow()
        try:
            model = window.current_model
            model.set_values([(0, 0, "1"), (1, 0, "2"), (2, 0, "3")])
            selection = QItemSelection(model.index(0, 1), model.index(2, 1))
            window.current_view.selectionModel().select(selection, QItemSelectionModel.ClearAndSelect)
            window.custom_algorithm_input.setText("=A{row}*2")

            window.fill_selection_with_custom_algorithm()

            self.assertEqual(model.data(model.index(0, 1), Qt.EditRole), "=A1*2")
            self.assertEqual(model.data(model.index(1, 1), Qt.EditRole), "=A2*2")
            self.assertEqual(model.data(model.index(2, 1), Qt.DisplayRole), "6")
        finally:
            window.close()


if __name__ == "__main__":
    unittest.main()
