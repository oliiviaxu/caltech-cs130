# python -m unittest discover -s tests
import unittest
import coverage
import sheets
import os
import lark
from sheets.interpreter import FormulaEvaluator
import decimal

current_dir = os.path.dirname(os.path.abspath(__file__))
lark_path = os.path.join(current_dir, "../sheets/formulas.lark")

class BasicTests(unittest.TestCase):
    def test_new_sheet(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.new_sheet('Lorem ipsum')
        wb.new_sheet('Sheet3')
        wb.new_sheet('sheet2')
        wb.new_sheet('Sheet5')
        wb.new_sheet()
        self.assertEqual(wb.list_sheets(), ['Sheet1', 'Lorem ipsum', 'Sheet3', 'sheet2', 'Sheet5', 'Sheet4'])

        invalid_sheet_names = ['', ' Sheet', '~', 'Lorem ipsum', 'Sheet\' name', 'Sheet \" name']
        for sheet_name in invalid_sheet_names:
            with self.assertRaises(ValueError):
                wb.new_sheet(sheet_name)
    
    def test_del_sheet(self):
        """
        Test the deletion of a sheet from the workbook.
        Arrange: Create a workbook and add a sheet with some cells.
        Act: Delete the sheet.
        Assert: Verify that the sheet is removed and references are updated.
        """
        # Arrange
        wb = sheets.Workbook()
        wb.new_sheet('Sheet1')
        wb.new_sheet('Sheet2')

        wb.set_cell_contents('Sheet1', 'A1', '2')
        wb.set_cell_contents('Sheet1', 'A2', '=1 + A1')
        wb.set_cell_contents('Sheet2', 'A1', '=1+Sheet1!A1')
        
        # Act
        wb.del_sheet('Sheet2')
        
        # Assert
        self.assertEqual(wb.list_sheets(), ['Sheet1'])
        self.assertEqual(len(wb.sheets['sheet1'].cells[0][0].ingoing), 1)
        self.assertEqual(wb.sheets['sheet1'].cells[0][0].ingoing[0].location, 'A2')
    
    def test_spreadsheet_cells(self):
        wb = sheets.Workbook()
        wb.new_sheet()

        # test setting to number
        wb.set_cell_contents('Sheet1', 'A1', '4')
        self.assertEqual(wb.get_cell_contents('Sheet1', 'A1'), '4')
        self.assertIsInstance(wb.get_cell_value('Sheet1', 'A1'), decimal.Decimal)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1'), decimal.Decimal(4))

        # test setting to number, removing trailing zeros
        wb.set_cell_contents('Sheet1', 'A1', '4.0000')
        self.assertEqual(wb.get_cell_contents('Sheet1', 'A1'), '4.0000')
        self.assertIsInstance(wb.get_cell_value('Sheet1', 'A1'), decimal.Decimal)
        self.assertEqual(str(wb.get_cell_value('Sheet1', 'A1')), '4')

        # test setting to Infinity
        wb.set_cell_contents('Sheet1', 'A1', 'inf')
        self.assertEqual(wb.get_cell_contents('Sheet1', 'A1'), 'inf')
        self.assertIsInstance(wb.get_cell_value('Sheet1', 'A1'), str)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1'), 'inf')

        # test setting to NaN
        wb.set_cell_contents('Sheet1', 'A1', 'NaN')
        self.assertEqual(wb.get_cell_contents('Sheet1', 'A1'), 'NaN')
        self.assertIsInstance(wb.get_cell_value('Sheet1', 'A1'), str)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1'), 'NaN')

        # test setting to string, no whitespace
        wb.set_cell_contents('Sheet1', 'A1', '\'string')
        self.assertEqual(wb.get_cell_contents('Sheet1', 'A1'), '\'string')
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1'), 'string')

        # test setting to string, handles whitespace correctly
        wb.set_cell_contents('Sheet1', 'A1', '\'  my string  ')
        self.assertEqual(wb.get_cell_contents('Sheet1', 'A1'), '\'  my string')
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1'), '  my string')

        # test setting to literal, handles whitespace correctly
        wb.set_cell_contents('Sheet1', 'A1', '  my string  ')
        self.assertEqual(wb.get_cell_contents('Sheet1', 'A1'), 'my string')
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1'), 'my string')

        # test setting to empty string
        wb.set_cell_contents('Sheet1', 'A1', '')
        self.assertEqual(wb.get_cell_contents('Sheet1', 'A1'), None)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1'), None)
    
    def test_spreadsheet_extent(self):
        wb = sheets.Workbook()
        wb.new_sheet()

        # test the extent of spreadsheet
        wb.set_cell_contents('Sheet1', 'AA26', 'test')
        self.assertEqual(wb.get_sheet_extent('Sheet1'), (27, 26))
        wb.set_cell_contents('Sheet1', 'C27', 'test')
        self.assertEqual(wb.get_sheet_extent('Sheet1'), (27, 27)) # test adding rows
        wb.set_cell_contents('Sheet1', 'AB4', '  test  ')
        self.assertEqual(wb.get_sheet_extent('Sheet1'), (28, 27)) # test adding columns

        # test KeyError for missing sheet and ValueError for bad location
        with self.assertRaises(KeyError):
            wb.set_cell_contents('Sheet2', 'D5', 'test')
        with self.assertRaises(ValueError):
            wb.set_cell_contents('Sheet1', 'D5D5', 'test')
    
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
    
    def test_unset_cells(self):
        wb = sheets.Workbook()
        wb.new_sheet()

        # addition/subtraction
        wb.set_cell_contents('Sheet1', 'A1', '=1 - D1')
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1'), 1)

        # multiplication/division
        wb.set_cell_contents('Sheet1', 'A1', '=1 * D1')
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1'), 0)
        
        # TODO: is division supposed to cause DIVIDE_BY_ZERO?
        wb.set_cell_contents('Sheet1', 'A1', '=1 / D1')
        self.assertIsInstance(wb.get_cell_value('Sheet1', 'A1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1').get_type(), sheets.CellErrorType.DIVIDE_BY_ZERO)

        # unary
        # TODO: this failed acceptance tests, but passes here?
        wb.set_cell_contents('Sheet1', 'A1', '=-D1')
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1'), 0)

        # concat
        wb.set_cell_contents('Sheet1', 'A1', '="test" & D1')
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1'), 'test')

        # parentheses
        wb.set_cell_contents('Sheet1', 'A1', '=(D1)')
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1'), 0)
    
    def test_concat(self):
        wb = sheets.Workbook()
        wb.new_sheet()

        # Test that trailing zeros are removed when concatenating a number and a string
        wb.set_cell_contents('Sheet1', 'A1', '=4.00000 & "test"')
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1'), '4test')

        wb.set_cell_contents('Sheet1', 'A2', '4.00000')
        wb.set_cell_contents('Sheet1', 'A1', '="test" & A2')
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1'), 'test4')

    def test_implicit_conversion(self):
        wb = sheets.Workbook()
        wb.new_sheet()

        # test with None -> 0
        wb.set_cell_contents('Sheet1', 'A1', '=A2 + 4')
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1'), 4)
        
        # test with None -> empty string
        wb.set_cell_contents('Sheet1', 'A1', '=A2 & "test"')
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1'), "test")

        # test with string -> number for +, *, unary
        wb.set_cell_contents('Sheet1', 'A1', '=1 + "   0.04   "')
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1'), decimal.Decimal('1.04'))

        wb.set_cell_contents('Sheet1', 'A1', '=2 * "   0.04   "')
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1'), decimal.Decimal('0.08'))

        wb.set_cell_contents('Sheet1', 'A1', '=4 * -"  .01  "')
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1'), decimal.Decimal('-0.04'))

        # test with number -> string for &
        wb.set_cell_contents('Sheet1', 'A1', '="test " & A2')
        wb.set_cell_contents('Sheet1', 'A2', '0.4')
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1'), "test 0.4")

    def test_formula_evaluation(self):
        wb = sheets.Workbook()
        wb.new_sheet()

        wb.set_cell_contents('Sheet1', 'A1', '=1 + 2 * 3')
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1'), 7)

        wb.set_cell_contents('Sheet1', 'A1', '="aba" & "cadabra"')
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1'), 'abacadabra')

        wb.set_cell_contents('Sheet1', 'A1', '=3 * -4 - 5')
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1'), -17)

        wb.set_cell_contents('Sheet1', 'A1', '=-A2')
        wb.set_cell_contents('Sheet1', 'A2', '-4')
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1'), 4)
    
    def test_cell_reference(self):
        wb = sheets.Workbook()
        wb.new_sheet()

        wb.set_cell_contents('Sheet1', 'D2', '2')
        wb.set_cell_contents('Sheet1', 'D3', '=1 + D2')

        # print(wb.sheets['sheet1'].cells[row_idx][col_idx].outgoing[0].location)

        self.assertEqual(wb.sheets['sheet1'].cells[2][3].outgoing[0].location, 'D2')
        self.assertEqual(wb.get_cell_value('Sheet1', 'D3'), 3)

        wb.set_cell_contents('Sheet1', 'A1', '=1 + Sheet1!A2')
        wb.set_cell_contents('Sheet1', 'A2', '4')
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1'), 5)

        # test reference to another sheet, plus capitalization differences
        wb.new_sheet()
        wb.set_cell_contents('Sheet2', 'A1', 'asdf')
        wb.set_cell_contents('Sheet1', 'A1', '="test" & shEEt2!A1')
        self.assertEqual(wb.get_cell_value('sheet1', 'a1'), "testasdf")

        # test changing cell's references
        wb.set_cell_contents('Sheet1', 'A1', '=A2 + 4')
        wb.set_cell_contents('Sheet1', 'B1', '=1 + A3')
        wb.set_cell_contents('Sheet1', 'A2', '2')
        wb.set_cell_contents('Sheet1', 'A3', '1')
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1'), 6)
        self.assertEqual(len(wb.get_cell('Sheet1', 'A2').ingoing), 1)
        self.assertEqual(len(wb.get_cell('Sheet1', 'A3').ingoing), 1)
        wb.set_cell_contents('Sheet1', 'A1', '=A3 + 4')
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1'), 5)
        self.assertEqual(len(wb.get_cell('Sheet1', 'A2').ingoing), 0)
        self.assertEqual(len(wb.get_cell('Sheet1', 'A3').ingoing), 2)
    
    def test_ingoing_outgoing(self):
        wb = sheets.Workbook()
        wb.new_sheet()

        wb.set_cell_contents('Sheet1', 'D1', '2')
        wb.set_cell_contents('Sheet1', 'D2', '=D1+3')
        wb.set_cell_contents('Sheet1', 'D3', '=D1+4')

        d_1 = wb.sheets['sheet1'].cells[0][3]
        d_2 = wb.sheets['sheet1'].cells[1][3]
        d_3 = wb.sheets['sheet1'].cells[2][3]

        # def print_lists(cell):
        #     print(f'########## PRINTING {cell.location} LIST ############')
        #     print('Ingoing: ')
        #     for c in cell.ingoing:
        #         print(f'{c.location}')
            
        #     print('Outgoing: ')
        #     for c in cell.outgoing:
        #         print(f'{c.location}')
        
        # print_lists(d_1)
        # print_lists(d_2)
        # print_lists(d_3)
        
        wb.set_cell_contents('Sheet1', 'D2', '=1')

        d_1 = wb.sheets['sheet1'].cells[0][3]
        d_2 = wb.sheets['sheet1'].cells[1][3]
        d_3 = wb.sheets['sheet1'].cells[2][3]

        # print_lists(d_1)
        # print_lists(d_2)
        # print_lists(d_3)
    
    def test_cycle_detection(self):
        wb = sheets.Workbook()
        wb.new_sheet()

        wb.set_cell_contents('Sheet1', 'A1', '=B1')
        wb.set_cell_contents('Sheet1', 'B1', '=A1')

        cell = wb.sheets['sheet1'].cells[0][0]
        self.assertEqual(wb.detect_cycle(cell), True)

        wb.set_cell_contents('Sheet1', 'B1', '=C1')
        self.assertEqual(wb.detect_cycle(cell), False)
    
    def test_interpreter(self):
        wb = sheets.Workbook()
        wb.new_sheet()

        parser = lark.Lark.open(lark_path, start='formula')
        tree_1 = parser.parse('=1 + D3')
        ref_info = wb.get_cell_ref_info(tree_1, 'sheet1')

        ev = FormulaEvaluator('sheet1', ref_info)

        self.assertEqual(ev.visit(tree_1), 1)

        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', '=-3')
        wb.set_cell_contents('sheet1', 'A2', '=+A1')
        self.assertEqual(wb.get_cell_value('sheet1', 'A2'), -3)

        wb.set_cell_contents('sheet1', 'A1', '=1/0')
        wb.set_cell_contents('sheet1', 'A2', '=A1+3')
        wb.set_cell_contents('sheet1', 'A3', '=A1/3')
        wb.set_cell_contents('sheet1', 'A4', '=1/2')
        wb.set_cell_contents('sheet1', 'A5', '=A1&A1')

        self.assertIsInstance(wb.get_cell_value('sheet1', 'A1'), sheets.CellError)
        self.assertIsInstance(wb.get_cell_value('sheet1', 'A2'), sheets.CellError)
        self.assertIsInstance(wb.get_cell_value('sheet1', 'A3'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('sheet1', 'A4'), 1/2)
        self.assertIsInstance(wb.get_cell_value('sheet1', 'A5'), sheets.CellError)

        # print(tree.pretty())
        # print(ev.visit(tree))
    
    def test_automatic_updates(self):
        # TODO: test automatic updates
        wb = sheets.Workbook()
        wb.new_sheet()

        wb.set_cell_contents('Sheet1', 'A1', '5')
        wb.set_cell_contents('Sheet1', 'A2', '=A1 + 2')
        wb.set_cell_contents('Sheet1', 'A3', '=5 * A1')   
        self.assertEqual(wb.get_cell_value('Sheet1', 'A2'), 7)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A3'), 25)

        wb.set_cell_contents('Sheet1', 'A1', '7')
        self.assertEqual(wb.get_cell_value('Sheet1', 'A2'), 9)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A3'), 35)

    def test_order_evaluation(self):
        wb = sheets.Workbook()
        wb.new_sheet()

        wb.set_cell_contents('Sheet1', 'A1', '=B1+C1')
        wb.set_cell_contents('Sheet1', 'B1', '=C1')
        wb.set_cell_contents('Sheet1', 'C1', '1')
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1'), 2)

        wb.set_cell_contents('Sheet1', 'C1', '2')
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1'), 4)

        # diamond pattern
        wb.new_sheet()
        wb.set_cell_contents('Sheet2', 'A1', '=B1 + D1')
        wb.set_cell_contents('Sheet2', 'B1', '=C1 + 2')
        wb.set_cell_contents('Sheet2', 'D1', '=C1')
        wb.set_cell_contents('Sheet2', 'C1', '4')
        self.assertEqual(wb.get_cell_value('Sheet2', 'C1'), 4)
        self.assertEqual(wb.get_cell_value('Sheet2', 'D1'), 4)
        self.assertEqual(wb.get_cell_value('Sheet2', 'B1'), 6)
        self.assertEqual(wb.get_cell_value('Sheet2', 'A1'), 10)

        wb.set_cell_contents('Sheet2', 'C1', '5')
        self.assertEqual(wb.get_cell_value('Sheet2', 'C1'), 5)
        self.assertEqual(wb.get_cell_value('Sheet2', 'D1'), 5)
        self.assertEqual(wb.get_cell_value('Sheet2', 'B1'), 7)
        self.assertEqual(wb.get_cell_value('Sheet2', 'A1'), 12)

        # more expansive topological order evaluation test
        wb.new_sheet()
        wb.set_cell_contents('Sheet3', 'A1', '=B1 + C1 + D1')
        wb.set_cell_contents('Sheet3', 'B1', '=E1 * 4')
        wb.set_cell_contents('Sheet3', 'C1', '=E1 - 1')
        wb.set_cell_contents('Sheet3', 'D1', '=E1 / 2')
        wb.set_cell_contents('Sheet3', 'E1', '=4')
        self.assertEqual(wb.get_cell_value('Sheet3', 'E1'), 4)
        self.assertEqual(wb.get_cell_value('Sheet3', 'D1'), 2)
        self.assertEqual(wb.get_cell_value('Sheet3', 'C1'), 3)
        self.assertEqual(wb.get_cell_value('Sheet3', 'B1'), 16)
        self.assertEqual(wb.get_cell_value('Sheet3', 'A1'), 21)

        wb.set_cell_contents('Sheet3', 'E1', '2')
        self.assertEqual(wb.get_cell_value('Sheet3', 'E1'), 2)
        self.assertEqual(wb.get_cell_value('Sheet3', 'D1'), 1)
        self.assertEqual(wb.get_cell_value('Sheet3', 'C1'), 1)
        self.assertEqual(wb.get_cell_value('Sheet3', 'B1'), 8)
        self.assertEqual(wb.get_cell_value('Sheet3', 'A1'), 10)
    
    def test_update_tree(self):
        wb = sheets.Workbook()
        wb.new_sheet()

        wb.set_cell_contents('Sheet1', 'A1', '=B1 + C1')
        wb.set_cell_contents('Sheet1', 'B1', '=C1')
        wb.set_cell_contents('Sheet1', 'C1', '=1')
        wb.set_cell_contents('Sheet1', 'D1', '=A1 + E1')
        wb.set_cell_contents('Sheet1', 'E1', '=F1')
        wb.set_cell_contents('Sheet1', 'F1', '=D1')
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1'), 2)
        self.assertEqual(wb.get_cell_value('Sheet1', 'B1'), 1)
        self.assertEqual(wb.get_cell_value('Sheet1', 'C1'), 1)
        self.assertIsInstance(wb.get_cell_value('Sheet1', 'D1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'D1').get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)
        self.assertIsInstance(wb.get_cell_value('Sheet1', 'E1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'E1').get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)
        self.assertIsInstance(wb.get_cell_value('Sheet1', 'F1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'F1').get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)

        wb.set_cell_contents('Sheet1', 'C1', '2')
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1'), 4)
        self.assertEqual(wb.get_cell_value('Sheet1', 'B1'), 2)
        self.assertEqual(wb.get_cell_value('Sheet1', 'C1'), 2)
        self.assertIsInstance(wb.get_cell_value('Sheet1', 'D1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'D1').get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)
        self.assertIsInstance(wb.get_cell_value('Sheet1', 'E1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'E1').get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)
        self.assertIsInstance(wb.get_cell_value('Sheet1', 'F1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'F1').get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)
        
        wb.set_cell_contents('Sheet1', 'A1', '3')
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1'), 3)
        self.assertEqual(wb.get_cell_value('Sheet1', 'B1'), 2)
        self.assertEqual(wb.get_cell_value('Sheet1', 'C1'), 2)
        self.assertIsInstance(wb.get_cell_value('Sheet1', 'D1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'D1').get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)
        self.assertIsInstance(wb.get_cell_value('Sheet1', 'E1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'E1').get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)
        self.assertIsInstance(wb.get_cell_value('Sheet1', 'F1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'F1').get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)

        wb.set_cell_contents('Sheet1', 'A1', '=B1 + C1')
        wb.set_cell_contents('Sheet1', 'D1', '=A1')
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1'), 4)
        self.assertEqual(wb.get_cell_value('Sheet1', 'B1'), 2)
        self.assertEqual(wb.get_cell_value('Sheet1', 'C1'), 2)
        self.assertEqual(wb.get_cell_value('Sheet1', 'D1'), 4)
        self.assertEqual(wb.get_cell_value('Sheet1', 'E1'), 4)
        self.assertEqual(wb.get_cell_value('Sheet1', 'F1'), 4)
    
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

# TODO: change ingoing and outgoing to sets (not arrays)
# TODO: edge cases for deleting sheets - Sheet1!A1 references Sheet2!A1, then delete Sheet2. 
# what happens to Sheet1!A1 - contents? value?
# and what if we create another sheet called Sheet1?

# TODO: more extensive tests for cell references, automatic updating, sheet deletion, cycles
# TODO: tests for unset cells
