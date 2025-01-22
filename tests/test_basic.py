# python -m unittest discover -s tests
import unittest
import sheets
import os
import sheets.Sheet as Sheet
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

        with self.assertRaises(ValueError):
            wb.new_sheet('')
        with self.assertRaises(ValueError):
            wb.new_sheet(' Sheet')
        with self.assertRaises(ValueError):
            wb.new_sheet('~')
        with self.assertRaises(ValueError):
            wb.new_sheet('Lorem ipsum')
    
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
        wb.del_sheet('Sheet2')
        self.assertIsInstance(wb.get_cell_value('Sheet1', 'A2'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A2').get_type(), sheets.CellErrorType.BAD_REFERENCE)

        # TODO: is a cell reference invalid if it is out of the extent of the sheet?
        # # test bad reference (reference out of extent of sheet)
        # wb.new_sheet()
        # wb.set_cell_contents('Sheet1', 'A1', '=1 + Sheet2!ZZ45')
        # self.assertIsInstance(wb.get_cell_value('Sheet1', 'A1'), sheets.CellError)
        # self.assertEqual(wb.get_cell_value('Sheet1', 'A1').get_type(), sheets.CellErrorType.BAD_REFERENCE)

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

        # test setting content to string representation of error
        wb.set_cell_contents('Sheet1', 'A1', '=#REF!')
        self.assertIsInstance(wb.get_cell_value('Sheet1', 'A1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1').get_type(), sheets.CellErrorType.BAD_REFERENCE)

        wb.set_cell_contents('Sheet1', 'A1', '=#ref!')
        self.assertIsInstance(wb.get_cell_value('Sheet1', 'A1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1').get_type(), sheets.CellErrorType.BAD_REFERENCE)

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

    def test_error_priority(self):
        # test error prioritization
        pass
    
    def test_implicit_conversion(self):
        # TODO: test implicit type conversion
        # Test with None -> 0, None -> empty string
        # test with number -> string for &
        # test with string -> number for +, *
        # 
        pass

    def test_formula_evaluation(self):
        wb = sheets.Workbook()
        wb.new_sheet()

        wb.set_cell_contents('Sheet1', 'A1', '=1 + 2 * 3')
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1'), 7)

        wb.set_cell_contents('Sheet1', 'A1', '="aba" & "cadabra"')
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1'), 'abacadabra')
    
    def test_cell_reference(self):
        wb = sheets.Workbook()
        wb.new_sheet()

        wb.set_cell_contents('Sheet1', 'D2', '2')
        wb.set_cell_contents('Sheet1', 'D3', '=1 + D2')

        col_idx, row_idx = Sheet.split_cell_ref('D3')
        # print(wb.sheets['sheet1'].cells[row_idx][col_idx].outgoing[0].location)

        self.assertEqual(wb.sheets['sheet1'].cells[row_idx][col_idx].outgoing[0].location, 'D2')
        self.assertEqual(wb.get_cell_value('Sheet1', 'D3'), 3)

        wb.set_cell_contents('Sheet1', 'A1', '=1 + Sheet1!A2')
        wb.set_cell_contents('Sheet1', 'A2', '4')
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1'), 5)

        # test reference to another sheet, plus capitalization differences
        wb.new_sheet()
        wb.set_cell_contents('Sheet2', 'A1', 'asdf')
        wb.set_cell_contents('Sheet1', 'A1', '="test" & shEEt2!A1')
        self.assertEqual(wb.get_cell_value('sheet1', 'a1'), "testasdf")
    
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
        tree = parser.parse('=1 + D3')
        ref_info = wb.get_cell_ref_info(tree, 'sheet1')

        ev = FormulaEvaluator('sheet1', ref_info)
        # print(tree.pretty())
        # print(ev.visit(tree))
    
    def test_automatic_updates(self):
        # TODO: test automatic updates
        pass

if __name__ == "__main__":
    unittest.main()




# TODO: edge case
# cell A1 references cell AAA45, which is error at first
# then populate AAA45, cell A1 should fix itself

# TODO: change ingoing and outgoing to sets
# TODO: for arithmetic on empty cell - do we treat empty cell as 0?

# TODO: edge cases for deleting sheets - Sheet1!A1 references Sheet2!A1, then delete Sheet2. 
# what happens to Sheet1!A1 - contents? value?
# and what if we create another sheet called Sheet1?

# TODO: add new sheet - in excel, if you add sheet1, add sheet2, delete sheet2, then new_sheet gives sheet3.
# is this desired for our engine also?
# TODO: test unary operator