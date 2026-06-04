import unittest

from pyexcel_lite.formula import FormulaCycleError, FormulaEvaluator
from pyexcel_lite.workbook import WorkbookData


class FormulaEvaluatorTest(unittest.TestCase):
    def setUp(self):
        self.workbook = WorkbookData()
        self.sheet = self.workbook.active_sheet
        self.evaluator = FormulaEvaluator(self.workbook)

    def test_arithmetic_and_cell_refs(self):
        self.sheet.set_value(0, 0, "10")
        self.sheet.set_value(0, 1, "5")
        self.sheet.set_value(0, 2, "=A1+B1*2")
        self.assertEqual(self.evaluator.evaluate_cell(self.sheet, 0, 2), 20.0)

    def test_sum_range(self):
        for row, value in enumerate([1, 2, 3, 4]):
            self.sheet.set_value(row, 0, value)
        self.sheet.set_value(0, 1, "=SUM(A1:A4)")
        self.assertEqual(self.evaluator.evaluate_cell(self.sheet, 0, 1), 10.0)

    def test_average_and_count(self):
        self.sheet.set_value(0, 0, "3")
        self.sheet.set_value(1, 0, "9")
        self.sheet.set_value(2, 0, "text")
        self.sheet.set_value(0, 1, "=AVERAGE(A1:A2)")
        self.sheet.set_value(1, 1, "=COUNT(A1:A3)")
        self.sheet.set_value(2, 1, "=COUNTA(A1:A3)")
        self.assertEqual(self.evaluator.evaluate_cell(self.sheet, 0, 1), 6.0)
        self.assertEqual(self.evaluator.evaluate_cell(self.sheet, 1, 1), 2)
        self.assertEqual(self.evaluator.evaluate_cell(self.sheet, 2, 1), 3)

    def test_if_and_text_functions(self):
        self.sheet.set_value(0, 0, "mohamed")
        self.sheet.set_value(0, 1, '=IF(1<2, UPPER(A1), "no")')
        self.assertEqual(self.evaluator.evaluate_cell(self.sheet, 0, 1), "MOHAMED")

    def test_excel_style_equals_comparison(self):
        self.sheet.set_value(0, 0, "10")
        self.sheet.set_value(0, 1, '=IF(A1=10, "ok", "no")')
        self.assertEqual(self.evaluator.evaluate_cell(self.sheet, 0, 1), "ok")

    def test_cycle_detection(self):
        self.sheet.set_value(0, 0, "=B1")
        self.sheet.set_value(0, 1, "=A1")
        with self.assertRaises(FormulaCycleError):
            self.evaluator.evaluate_cell(self.sheet, 0, 0)


if __name__ == "__main__":
    unittest.main()
