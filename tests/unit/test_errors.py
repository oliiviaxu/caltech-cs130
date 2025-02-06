import unittest
import coverage
import sheets

class ErrorTests(unittest.TestCase):
    def test_error_propagation(self):
        wb = sheets.Workbook()
        wb.new_sheet()

        wb.set_cell_contents('Sheet1', 'A1', '=1 + #ref!')
        self.assertIsInstance(wb.get_cell_value('Sheet1', 'A1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1').get_type(), sheets.CellErrorType.BAD_REFERENCE)

        wb.set_cell_contents('Sheet1', 'A1', '=1 + 1/0')
        self.assertIsInstance(wb.get_cell_value('Sheet1', 'A1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1').get_type(), sheets.CellErrorType.DIVIDE_BY_ZERO)

        wb.set_cell_contents('Sheet1', 'A1', '=1 + A2')
        wb.set_cell_contents('Sheet1', 'A2', '=1 + "mystring"')
        self.assertIsInstance(wb.get_cell_value('Sheet1', 'A1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1').get_type(), sheets.CellErrorType.TYPE_ERROR)

        # now with multiply instead of addition
        wb.set_cell_contents('Sheet1', 'A1', '=1 * #ref!')
        self.assertIsInstance(wb.get_cell_value('Sheet1', 'A1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1').get_type(), sheets.CellErrorType.BAD_REFERENCE)

        wb.set_cell_contents('Sheet1', 'A1', '=1 * (1/0)')
        self.assertIsInstance(wb.get_cell_value('Sheet1', 'A1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1').get_type(), sheets.CellErrorType.DIVIDE_BY_ZERO)

        wb.set_cell_contents('Sheet1', 'A1', '=1 * A2')
        wb.set_cell_contents('Sheet1', 'A2', '=1 + "mystring"')
        self.assertIsInstance(wb.get_cell_value('Sheet1', 'A1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1').get_type(), sheets.CellErrorType.TYPE_ERROR)

        # now with unary
        # now with multiply instead of addition
        wb.set_cell_contents('Sheet1', 'A1', '=1 * -#ref!')
        self.assertIsInstance(wb.get_cell_value('Sheet1', 'A1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1').get_type(), sheets.CellErrorType.BAD_REFERENCE)

        wb.set_cell_contents('Sheet1', 'A1', '=1 * (-(1/0))')
        self.assertIsInstance(wb.get_cell_value('Sheet1', 'A1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1').get_type(), sheets.CellErrorType.DIVIDE_BY_ZERO)

        wb.set_cell_contents('Sheet1', 'A1', '=1 * -A2')
        wb.set_cell_contents('Sheet1', 'A2', '=1 + -"mystring"')
        self.assertIsInstance(wb.get_cell_value('Sheet1', 'A1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1').get_type(), sheets.CellErrorType.TYPE_ERROR)

        # now test propagation of circref
        wb.set_cell_contents('Sheet1', 'A1', '=B1')
        wb.set_cell_contents('Sheet1', 'B1', '=A1')
        wb.set_cell_contents('Sheet1', 'C1', '=Sheet1!a1')
        self.assertIsInstance(wb.get_cell_value('Sheet1', 'A1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1').get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)
        self.assertIsInstance(wb.get_cell_value('Sheet1', 'B1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'B1').get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)
        self.assertIsInstance(wb.get_cell_value('Sheet1', 'C1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'C1').get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)

    def test_error_priority(self):
        wb = sheets.Workbook()
        wb.new_sheet()

        # parse error comes before circref or any other error
        wb.set_cell_contents('Sheet1', 'A1', '=B1')
        wb.set_cell_contents('Sheet1', 'B1', '=C1')
        wb.set_cell_contents('Sheet1', 'C1', '=BB/0')
        self.assertIsInstance(wb.get_cell_value('Sheet1', 'A1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1').get_type(), sheets.CellErrorType.PARSE_ERROR)
        self.assertIsInstance(wb.get_cell_value('Sheet1', 'B1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'B1').get_type(), sheets.CellErrorType.PARSE_ERROR)
        self.assertIsInstance(wb.get_cell_value('Sheet1', 'C1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'C1').get_type(), sheets.CellErrorType.PARSE_ERROR)
        
        # circref comes before other errors (assuming no parse error)
        wb.set_cell_contents('Sheet1', 'A1', '=B1')
        wb.set_cell_contents('Sheet1', 'B1', '=C1')
        wb.set_cell_contents('Sheet1', 'C1', '=B1/0')
        self.assertIsInstance(wb.get_cell_value('Sheet1', 'A1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1').get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)
        self.assertIsInstance(wb.get_cell_value('Sheet1', 'B1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'B1').get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)
        self.assertIsInstance(wb.get_cell_value('Sheet1', 'C1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'C1').get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)
    def test_cell_error(self):
        wb = sheets.Workbook()
        wb.new_sheet()

        # test divide by zero
        wb.set_cell_contents('Sheet1', 'A1', '=1/0')
        self.assertIsInstance(wb.get_cell_value('Sheet1', 'A1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1').get_type(), sheets.CellErrorType.DIVIDE_BY_ZERO)

        # test parse error
        wb.set_cell_contents('Sheet1', 'A1', '=A5+++46a4')
        self.assertIsInstance(wb.get_cell_value('Sheet1', 'A1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1').get_type(), sheets.CellErrorType.PARSE_ERROR)

        # test bad reference (sheet does not exist)
        wb.set_cell_contents('Sheet1', 'A1', '=1 + Sheet2!A1')
        self.assertIsInstance(wb.get_cell_value('Sheet1', 'A1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1').get_type(), sheets.CellErrorType.BAD_REFERENCE)

        # test bad reference (sheet was deleted)
        wb.new_sheet('Sheet2')
        wb.set_cell_contents('Sheet1', 'A2', '=1 + Sheet2!A1')
        self.assertEqual(wb.get_cell_value('Sheet1', 'A2'), 1)
        wb.del_sheet('Sheet2')
        self.assertIsInstance(wb.get_cell_value('Sheet1', 'A2'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A2').get_type(), sheets.CellErrorType.BAD_REFERENCE)

        # test bad reference (reference exceeds ZZZZ9999)
        wb.set_cell_contents('Sheet1', 'A1', '=1 + Sheet1!ZZZZZ99')
        self.assertIsInstance(wb.get_cell_value('Sheet1', 'A1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1').get_type(), sheets.CellErrorType.BAD_REFERENCE)

        # test circular reference
        wb.set_cell_contents('Sheet1', 'A1', '=B1')
        wb.set_cell_contents('Sheet1', 'B1', '=A1')
        self.assertIsInstance(wb.get_cell_value('Sheet1', 'A1'), sheets.CellError)
        self.assertIsInstance(wb.get_cell_value('Sheet1', 'B1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1').get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)
        self.assertEqual(wb.get_cell_value('Sheet1', 'B1').get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)

        # test type error
        wb.set_cell_contents('Sheet1', 'A1', '\'mystring')
        wb.set_cell_contents('Sheet1', 'A2', '=1 + A1')
        self.assertIsInstance(wb.get_cell_value('Sheet1', 'A2'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A2').get_type(), sheets.CellErrorType.TYPE_ERROR)

        wb.set_cell_contents('Sheet1', 'A1', '=1 + "test"')
        self.assertIsInstance(wb.get_cell_value('Sheet1', 'A1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1').get_type(), sheets.CellErrorType.TYPE_ERROR)

        wb.set_cell_contents('Sheet1', 'A1', '\'mystring')
        wb.set_cell_contents('Sheet1', 'A2', '=1 * A1')
        self.assertIsInstance(wb.get_cell_value('Sheet1', 'A2'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A2').get_type(), sheets.CellErrorType.TYPE_ERROR)

        wb.set_cell_contents('Sheet1', 'A1', '=1 * "test"')
        self.assertIsInstance(wb.get_cell_value('Sheet1', 'A1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1').get_type(), sheets.CellErrorType.TYPE_ERROR)

        wb.set_cell_contents('Sheet1', 'A1', '=1 * +"test"')
        self.assertIsInstance(wb.get_cell_value('Sheet1', 'A1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1').get_type(), sheets.CellErrorType.TYPE_ERROR)

        # test setting content to string representation of error
        wb.set_cell_contents('Sheet1', 'A1', '=#REF!')
        self.assertIsInstance(wb.get_cell_value('Sheet1', 'A1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1').get_type(), sheets.CellErrorType.BAD_REFERENCE)

        wb.set_cell_contents('Sheet1', 'A1', '=#ref!')
        self.assertIsInstance(wb.get_cell_value('Sheet1', 'A1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1').get_type(), sheets.CellErrorType.BAD_REFERENCE)

        wb.set_cell_contents('Sheet1', 'A2', '#div/0!')
        wb.set_cell_contents('Sheet1', 'A1', '=1 + A2')
        self.assertIsInstance(wb.get_cell_value('Sheet1', 'A2'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A2').get_type(), sheets.CellErrorType.DIVIDE_BY_ZERO)
        self.assertIsInstance(wb.get_cell_value('Sheet1', 'A1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1').get_type(), sheets.CellErrorType.DIVIDE_BY_ZERO)

if __name__ == "__main__":
    cov = coverage.Coverage()
    cov.start()
    unittest.main()
    cov.stop()
    cov.save()
    cov.html_report()