import os
import unittest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import QItemSelection, QItemSelectionModel, Qt
from PySide6.QtWidgets import QApplication

from pyexcel_lite.main import SpreadsheetWindow


class UndoChartCrossSheetTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def test_undo_and_redo_restore_cell_values(self):
        window = SpreadsheetWindow()
        try:
            model = window.current_model
            model.set_values([(0, 0, "10"), (0, 1, "20")])

            self.assertEqual(model.data(model.index(0, 0), Qt.EditRole), "10")
            self.assertTrue(model.can_undo())

            window.undo()
            self.assertEqual(model.data(model.index(0, 0), Qt.EditRole), "")
            self.assertEqual(model.data(model.index(0, 1), Qt.EditRole), "")
            self.assertTrue(model.can_redo())

            window.redo()
            self.assertEqual(model.data(model.index(0, 0), Qt.EditRole), "10")
            self.assertEqual(model.data(model.index(0, 1), Qt.EditRole), "20")
        finally:
            window.close()

    def test_chart_uses_selected_label_and_value_columns(self):
        window = SpreadsheetWindow()
        try:
            model = window.current_model
            model.set_values([(0, 0, "Gold"), (0, 1, "120"), (1, 0, "Silver"), (1, 1, "80")])
            selection = QItemSelection(model.index(0, 0), model.index(1, 1))
            window.current_view.selectionModel().select(selection, QItemSelectionModel.ClearAndSelect)
            window.chart_type_box.setCurrentText("Bar")
            window.chart_title_input.setText("Metals")

            window.create_chart_from_selection()

            self.assertEqual(window.chart_widget.title, "Metals")
            self.assertEqual(window.chart_widget.chart_type, "Bar")
            self.assertEqual([point.label for point in window.chart_widget.points], ["Gold", "Silver"])
            self.assertEqual([point.value for point in window.chart_widget.points], [120.0, 80.0])
            self.app.processEvents()
            self.assertIsNotNone(window.chart_dialog)
            self.assertTrue(window.chart_dialog.isVisible())
            self.assertEqual(window.chart_dialog_widget.title, "Metals")
            self.assertEqual(window.chart_dialog_widget.chart_type, "Bar")
            self.assertEqual([point.label for point in window.chart_dialog_widget.points], ["Gold", "Silver"])
        finally:
            if window.chart_dialog is not None:
                window.chart_dialog.close()
            window.close()

    def test_cross_sheet_formula_updates_after_source_sheet_change(self):
        window = SpreadsheetWindow()
        try:
            target = window.workbook.active_sheet
            source = window.workbook.add_sheet("Source")
            source.set_value(0, 0, "10")
            target.set_value(0, 0, "=Source!A1*2")

            self.assertEqual(window.evaluator.evaluate_cell(target, 0, 0), 20.0)
            source.set_value(0, 0, "15")
            self.assertEqual(window.evaluator.evaluate_cell(target, 0, 0), 30.0)
        finally:
            window.close()


if __name__ == "__main__":
    unittest.main()
