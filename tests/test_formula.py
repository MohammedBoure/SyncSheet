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

    def test_sum_range_ignores_text_cells_like_excel(self):
        self.sheet.set_value(0, 0, "10")
        self.sheet.set_value(1, 0, "djoher")
        self.sheet.set_value(2, 0, "5")
        self.sheet.set_value(0, 1, "=SUM(A1:A3)")
        self.assertEqual(self.evaluator.evaluate_cell(self.sheet, 0, 1), 15.0)

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

    def test_if_is_lazy_and_iferror_catches_math_errors(self):
        self.sheet.set_value(0, 0, '=IF(FALSE, 1/0, "safe")')
        self.sheet.set_value(0, 1, '=IFERROR(1/0, "fallback")')
        self.assertEqual(self.evaluator.evaluate_cell(self.sheet, 0, 0), "safe")
        self.assertEqual(self.evaluator.evaluate_cell(self.sheet, 0, 1), "fallback")

    def test_excel_style_equals_comparison(self):
        self.sheet.set_value(0, 0, "10")
        self.sheet.set_value(0, 1, '=IF(A1=10, "ok", "no")')
        self.assertEqual(self.evaluator.evaluate_cell(self.sheet, 0, 1), "ok")

    def test_semicolon_argument_separator(self):
        self.sheet.set_value(0, 0, "10")
        self.sheet.set_value(0, 1, '=IF(A1=10; "ok"; "no")')
        self.assertEqual(self.evaluator.evaluate_cell(self.sheet, 0, 1), "ok")

    def test_more_excel_math_and_text_operators(self):
        self.sheet.set_value(0, 0, "10")
        self.sheet.set_value(1, 0, "20")
        self.sheet.set_value(2, 0, "30")
        self.sheet.set_value(0, 1, "=PRODUCT(A1:A3)")
        self.sheet.set_value(1, 1, "=MEDIAN(A1:A3)")
        self.sheet.set_value(2, 1, '="Total: " & SUM(A1:A3)')
        self.sheet.set_value(3, 1, "=50%*200")
        self.assertEqual(self.evaluator.evaluate_cell(self.sheet, 0, 1), 6000.0)
        self.assertEqual(self.evaluator.evaluate_cell(self.sheet, 1, 1), 20)
        self.assertEqual(self.evaluator.evaluate_cell(self.sheet, 2, 1), "Total: 60")
        self.assertEqual(self.evaluator.evaluate_cell(self.sheet, 3, 1), 100.0)

    def test_criteria_and_lookup_functions(self):
        self.sheet.set_value(0, 0, "Ali")
        self.sheet.set_value(0, 1, "100")
        self.sheet.set_value(1, 0, "Sara")
        self.sheet.set_value(1, 1, "250")
        self.sheet.set_value(2, 0, "Ali")
        self.sheet.set_value(2, 1, "50")
        self.sheet.set_value(0, 2, '=SUMIF(A1:A3, "Ali", B1:B3)')
        self.sheet.set_value(1, 2, '=COUNTIF(B1:B3, ">100")')
        self.sheet.set_value(2, 2, '=VLOOKUP("Sara", A1:B3, 2, FALSE)')
        self.assertEqual(self.evaluator.evaluate_cell(self.sheet, 0, 2), 150.0)
        self.assertEqual(self.evaluator.evaluate_cell(self.sheet, 1, 2), 1)
        self.assertEqual(self.evaluator.evaluate_cell(self.sheet, 2, 2), 250)

    def test_advanced_statistical_functions(self):
        for row, value in enumerate([10, 20, 30, 40, 50]):
            self.sheet.set_value(row, 0, value)
            self.sheet.set_value(row, 1, value * 2)
        self.sheet.set_value(0, 2, "=PERCENTILE.INC(A1:A5, 0.5)")
        self.sheet.set_value(1, 2, "=QUARTILE.INC(A1:A5, 3)")
        self.sheet.set_value(2, 2, "=RANK.EQ(40, A1:A5, 0)")
        self.sheet.set_value(3, 2, "=CORREL(A1:A5, B1:B5)")
        self.sheet.set_value(4, 2, "=COVARIANCE.P(A1:A5, B1:B5)")
        self.assertEqual(self.evaluator.evaluate_cell(self.sheet, 0, 2), 30)
        self.assertEqual(self.evaluator.evaluate_cell(self.sheet, 1, 2), 40)
        self.assertEqual(self.evaluator.evaluate_cell(self.sheet, 2, 2), 2)
        self.assertAlmostEqual(self.evaluator.evaluate_cell(self.sheet, 3, 2), 1.0)
        self.assertEqual(self.evaluator.evaluate_cell(self.sheet, 4, 2), 400.0)

    def test_cycle_detection(self):
        self.sheet.set_value(0, 0, "=B1")
        self.sheet.set_value(0, 1, "=A1")
        with self.assertRaises(FormulaCycleError):
            self.evaluator.evaluate_cell(self.sheet, 0, 0)


if __name__ == "__main__":
    unittest.main()
