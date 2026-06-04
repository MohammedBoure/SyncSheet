import os
import unittest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import QItemSelection, QItemSelectionModel
from PySide6.QtWidgets import QApplication

from pyexcel_lite.main import SpreadsheetWindow


class CellContextMenuTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def test_context_menu_contains_selection_operations(self):
        window = SpreadsheetWindow()
        try:
            model = window.current_model
            selection = QItemSelection(model.index(0, 0), model.index(1, 1))
            window.current_view.selectionModel().select(selection, QItemSelectionModel.ClearAndSelect)
            QApplication.clipboard().setText("pasted")

            menu = window.current_view.build_context_menu()
            action_texts = [action.text() for action in menu.actions() if not action.isSeparator()]
            formula_menu = next(action.menu() for action in menu.actions() if action.text() == "Formulas")
            formula_texts = [action.text() for action in formula_menu.actions()]

            self.assertIn("Copy", action_texts)
            self.assertIn("Cut", action_texts)
            self.assertIn("Paste", action_texts)
            self.assertIn("Clear Selection", action_texts)
            self.assertIn("Insert Row", action_texts)
            self.assertIn("Delete Column", action_texts)
            self.assertIn("Create Chart From Selection", action_texts)
            self.assertIn("Apply Selected Template", formula_texts)
            self.assertTrue(all(not action.icon().isNull() for action in menu.actions() if not action.isSeparator()))
        finally:
            window.close()

    def test_right_click_selection_preparation_keeps_or_moves_selection(self):
        window = SpreadsheetWindow()
        try:
            model = window.current_model
            selection = QItemSelection(model.index(0, 0), model.index(1, 1))
            window.current_view.selectionModel().select(selection, QItemSelectionModel.ClearAndSelect)

            window.current_view.prepare_context_selection(model.index(1, 1))

            self.assertEqual(len(window.current_view.selectedIndexes()), 4)

            window.current_view.prepare_context_selection(model.index(4, 4))

            selected = window.current_view.selectedIndexes()
            self.assertEqual(len(selected), 1)
            self.assertEqual((selected[0].row(), selected[0].column()), (4, 4))
        finally:
            window.close()


if __name__ == "__main__":
    unittest.main()
