import unittest
import coverage
import sheets
import decimal

class SmokeTest(unittest.TestCase):
    def test_smoke_test(self):
        self.assertIsNotNone(sheets.__version__)
        # print(f'Using sheets engine version {sheets.__version__}')

        wb = sheets.Workbook()
        (index, name) = wb.new_sheet()
        self.assertEqual(name, "Sheet1")
        self.assertEqual(0, index)

        # Should print:  New spreadsheet "Sheet1" at index 0
        # print(f'New spreadsheet "{name}" at index {index}')

        wb.set_cell_contents(name, 'a1', '12')
        wb.set_cell_contents(name, 'b1', '34')
        wb.set_cell_contents(name, 'c1', '=a1+b1')

        # value should be a decimal.Decimal('46')
        value = wb.get_cell_value(name, 'c1')
        self.assertEqual(value, decimal.Decimal('46'))

        # Should print:  c1 = 46
        self.assertEqual(value, 46)

        wb.set_cell_contents(name, 'd3', '=nonexistent!b4')

        # value should be a CellError with type BAD_REFERENCE
        value = wb.get_cell_value(name, 'd3')
        self.assertIsInstance(value, sheets.CellError)
        self.assertEqual(value.get_type(), sheets.CellErrorType.BAD_REFERENCE)

        # Cells can be set to error values as well
        wb.set_cell_contents(name, 'e1', '#div/0!')
        wb.set_cell_contents(name, 'e2', '=e1+5')
        value = wb.get_cell_value(name, 'e2')
        self.assertIsInstance(value, sheets.CellError)
        self.assertEqual(value.get_type(), sheets.CellErrorType.DIVIDE_BY_ZERO)

if __name__ == "__main__":
    cov = coverage.Coverage()
    cov.start()
    unittest.main()
    cov.stop()
    cov.save()
    cov.html_report()