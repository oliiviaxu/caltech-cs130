import unittest
import coverage
import sheets
import os
import lark
from sheets.interpreter import FormulaEvaluator
import decimal
import json

class SpreadsheetTests(unittest.TestCase):
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
        self.assertEqual(len(wb.graph.ingoing_get('sheet1', 'a1')), 1)
        self.assertEqual(wb.graph.ingoing_get('sheet1', 'a1')[0], ('sheet1', 'a2'))

        wb2 = sheets.Workbook()
        wb2.new_sheet()
        wb2.new_sheet()

        wb2.del_sheet('Sheet2')
        self.assertEqual(wb2.list_sheets(), ['Sheet1'])

        with self.assertRaises(KeyError):
            wb2.del_sheet('Sheet4')

        # edge case
        wb3 = sheets.Workbook()
        wb3.new_sheet()
        wb3.new_sheet()
        wb3.new_sheet()
        wb3.set_cell_contents('Sheet1', 'A1', '=1 + Sheet2!A1')
        wb3.set_cell_contents('Sheet2', 'A1', '=1 + Sheet3!A1')
        wb3.del_sheet('Sheet2')
    
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

        self.assertEqual(wb.get_sheet_extent('Sheet1'), (0, 0))

        # test the extent of spreadsheet
        wb.set_cell_contents('Sheet1', 'AA20', 'test')
        self.assertEqual(wb.get_sheet_extent('Sheet1'), (27, 20))
        wb.set_cell_contents('Sheet1', 'C24', 'test')
        self.assertEqual(wb.get_sheet_extent('Sheet1'), (27, 24)) # test adding rows
        wb.set_cell_contents('Sheet1', 'AB4', '  test  ')
        self.assertEqual(wb.get_sheet_extent('Sheet1'), (28, 24)) # test adding columns

        # test KeyError for missing sheet and ValueError for bad location
        with self.assertRaises(KeyError):
            wb.set_cell_contents('Sheet2', 'D5', 'test')
        with self.assertRaises(ValueError):
            wb.set_cell_contents('Sheet1', 'D5D5', 'test')

        wb.new_sheet()

        # If we have a cell A1 which is set to the value of D4, then the extent
        # of the sheet is (1, 1) if D4 = None even though it is part of the
        # formula for A1.
        wb.set_cell_contents('Sheet2', 'A1', '=1 + D4')
        self.assertEqual(wb.get_cell_value('Sheet2', 'A1'), 1)
        self.assertEqual(wb.get_sheet_extent('Sheet2'), (1, 1))

        wb.set_cell_contents('Sheet2', 'D4', '2')
        self.assertEqual(wb.get_cell_value('Sheet2', 'A1'), 3)
        self.assertEqual(wb.get_sheet_extent('Sheet2'), (4, 4))

        # A sheet's extent should shrink as the maximal cell's contents are cleared.
        wb.set_cell_contents('Sheet1', 'AB4', None)
        self.assertEqual(wb.get_sheet_extent('Sheet1'), (27, 24))

        wb.set_cell_contents('Sheet1', 'C24', None)
        self.assertEqual(wb.get_sheet_extent('Sheet1'), (27, 20))

        # A sheet's extent should shrink to 0 if all cell contents are cleared.
        wb.set_cell_contents('Sheet1', 'AA20', None)
        self.assertEqual(wb.get_sheet_extent('Sheet1'), (0, 0))

        wb.set_cell_contents('Sheet2', 'A1', None)
        wb.set_cell_contents('Sheet2', 'D4', None)
        self.assertEqual(wb.get_sheet_extent('Sheet2'), (0, 0))
    
    def test_move_sheet(self):
        wb = sheets.Workbook()

        wb.new_sheet()
        wb.new_sheet('MySheet')
        wb.new_sheet()

        wb.move_sheet('Sheet1', 2)
        self.assertEqual(wb.list_sheets(), ['MySheet', 'Sheet2', 'Sheet1'])

        wb.move_sheet('Sheet2', 0)
        self.assertEqual(wb.list_sheets(), ['Sheet2', 'MySheet', 'Sheet1'])

        with self.assertRaises(KeyError):
            wb.move_sheet('Sheet4', 0)
        
        with self.assertRaises(IndexError):
            wb.move_sheet('Sheet2', 5)
    
    def test_copy_sheet(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.new_sheet()
        wb.new_sheet()

        wb.set_cell_contents('Sheet1', 'A1', '5')
        wb.set_cell_contents('Sheet1', 'B2', '=A1+3')

        index, sheet_name = wb.copy_sheet('Sheet1')

        with self.assertRaises(KeyError):
            wb.copy_sheet('Sheet4')
        
        self.assertEqual(index, 3)
        self.assertEqual(sheet_name, 'Sheet1_1')
        self.assertEqual(wb.get_cell_contents(sheet_name, 'A1'), '5')
        self.assertEqual(wb.get_cell_contents(sheet_name, 'B2'), '=A1+3')
        self.assertEqual(wb.get_cell_value(sheet_name, 'A1'), decimal.Decimal(5))
        self.assertEqual(wb.get_cell_value(sheet_name, 'B2'), decimal.Decimal(8))

        # check edge case
        wb.set_cell_contents('Sheet2', 'A1', '=2 + Sheet1_2!A1')
        self.assertIsInstance(wb.get_cell_value('Sheet2', 'A1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet2', 'A1').get_type(), sheets.CellErrorType.BAD_REFERENCE)

        wb.copy_sheet('Sheet1')
        self.assertEqual(wb.get_cell_value('Sheet2', 'A1'), decimal.Decimal('7'))
    
    def test_sheet_with_quotes(self):
        wb = sheets.Workbook()
        wb.new_sheet('Sheet 1')

        wb.new_sheet('Sheet2')
        wb.set_cell_contents('Sheet2', 'A1', "=1 + 'Sheet 1'!A1")
        self.assertEqual(wb.get_cell_value('Sheet2', 'A1'), 1)

        wb.set_cell_contents('Sheet2', 'A5', '4')
        wb.set_cell_contents('Sheet2', 'A2', "='Sheet2'!A5 + 5")
        self.assertEqual(wb.get_cell_value('Sheet2', 'A2'), decimal.Decimal('9'))
        self.assertEqual(wb.graph.ingoing_get('Sheet2', 'A5'), [('sheet2', 'a2')])
    
    def test_rename_sheet_1(self):
        # TODO: check empty list issue again
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.new_sheet()
        
        # basic test
        wb.set_cell_contents('Sheet2', 'A1', '=4 + Sheet1!A1')
        wb.set_cell_contents('Sheet1', 'B1', '=1 + Sheet2!B1')
        wb.rename_sheet('Sheet1', 'SheetBla')
        self.assertEqual(wb.graph.outgoing, {'sheet2': {'a1': [('sheetbla', 'a1')]}, 'sheetbla': {'b1': [('sheet2', 'b1')], 'a1': []}})
        self.assertEqual(wb.graph.ingoing, {'sheet2': {'b1': [('sheetbla', 'b1')]}, 'sheetbla': {'a1': [('sheet2', 'a1')]}})

        self.assertEqual(wb.get_cell_contents('Sheet2', 'A1'), '=4 + SheetBla!A1')
        self.assertEqual(wb.get_cell_contents('SheetBla', 'B1'), '=1 + Sheet2!B1')

        self.assertEqual(wb.get_cell('SheetBla', 'A1').sheet_name, 'SheetBla')
        self.assertEqual(wb.get_cell('SheetBla', 'B1').sheet_name, 'SheetBla')
    
    def test_rename_sheet_2(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('Sheet1', 'A1', '=A2')
        wb.rename_sheet('Sheet1', 'blah')
        self.assertEqual(wb.graph.outgoing, {'blah': {'a1': [('blah', 'a2')], 'a2': []}})
        self.assertEqual(wb.graph.ingoing, {'blah': {'a2': [('blah', 'a1')]}})

        self.assertEqual(wb.get_cell('blah', 'A1').sheet_name, 'blah')
    
    def test_rename_sheet_3(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.new_sheet()
        wb.new_sheet()
        wb.set_cell_contents('Sheet1', 'A1', '=Sheet2!A1 + Sheet3!A1')
        wb.set_cell_contents('Sheet2', 'B1', '=Sheet1!B1 + Sheet3!B1')
        wb.set_cell_contents('Sheet1', 'C1', '=Sheet2!C1 + D1')
        wb.rename_sheet('Sheet1', 'blah')
        self.assertEqual(wb.list_sheets(), ['blah', 'Sheet2', 'Sheet3'])
        self.assertEqual('sheet1' not in wb.graph.ingoing and 'sheet1' not in wb.graph.outgoing, True)
        self.assertEqual(sorted(wb.graph.outgoing_get('blah', 'a1')), sorted([('sheet2', 'a1'), ('sheet3', 'a1')]))
        self.assertEqual(sorted(wb.graph.outgoing_get('blah', 'c1')), sorted([('sheet2', 'c1'), ('blah', 'd1')]))
        self.assertEqual(wb.graph.ingoing_get('blah', 'b1'), [('sheet2', 'b1')])
        self.assertEqual(wb.graph.ingoing_get('blah', 'd1'), [('blah', 'c1')])
        self.assertEqual(sorted(wb.graph.outgoing_get('Sheet2', 'B1')), sorted([('sheet3', 'b1'), ('blah', 'b1')]))
        self.assertEqual(sorted(wb.graph.ingoing_get('Sheet2', 'A1')), sorted([('blah', 'a1')]))
        self.assertEqual(sorted(wb.graph.ingoing_get('Sheet3', 'A1')), sorted([('blah', 'a1')]))

    def test_rename_sheet(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.new_sheet()
        
        # basic test
        wb.set_cell_contents('Sheet2', 'A1', '=4 + Sheet1!A1')
        wb.rename_sheet('Sheet1', 'SheetBla')

        self.assertEqual(wb.list_sheets(), ['SheetBla', 'Sheet2'])
        self.assertEqual(wb.get_cell_contents('Sheet2', 'A1'), '=4 + SheetBla!A1')

        # test that Sheet1 dependency graph is changed
        self.assertEqual('sheet1' not in wb.graph.ingoing and 'sheet1' not in wb.graph.outgoing, True)
        self.assertEqual(wb.graph.outgoing_get('Sheet2', 'A1'), [('sheetbla', 'a1')])

        # test multiplication, addition, unary, parethesis kept
        wb.set_cell_contents('Sheet2', 'A1', '=(SheetBla!A1 * 4.0 / (SheetBla!A2 + 1))')
        wb.set_cell_contents('Sheet2', 'A2', '=-SheetBla!A1')
        wb.rename_sheet('SheetBla', 'Sheet1')

        self.assertEqual(wb.list_sheets(), ['Sheet1', 'Sheet2'])
        self.assertEqual(wb.get_cell_contents('Sheet2', 'A1'), '=(Sheet1!A1 * 4.0 / (Sheet1!A2 + 1))')
        self.assertEqual(wb.get_cell_contents('Sheet2', 'A2'), '=-Sheet1!A1')

        # test sheet names with quotes are handled correctly
        wb.set_cell_contents('Sheet2', 'A1', "='Sheet1'!A5 + 'Sheet2'!A6")
        wb.rename_sheet('Sheet1', 'SheetBla')
        self.assertEqual(wb.get_cell_contents('Sheet2', 'A1'), "=SheetBla!A5 + Sheet2!A6")

        wb.new_sheet('Sheet 3')
        wb.set_cell_contents('Sheet2', 'A1', "='SheetBla'!A5 + 'Sheet 3'!A6")
        wb.rename_sheet('SheetBla', 'Sheet1')
        self.assertEqual(wb.get_cell_contents('Sheet2', 'A1'), "=Sheet1!A5 + 'Sheet 3'!A6")

        # more sheet names with quotes
        wb.set_cell_contents('Sheet2', 'A1', "='Sheet1'!A5 + 'Sheet2'!A6")
        wb.rename_sheet('Sheet1', 'Sheet Bla')
        self.assertEqual(wb.get_cell_contents('Sheet2', 'A1'), "='Sheet Bla'!A5 + Sheet2!A6")

        wb.set_cell_contents('Sheet2', 'A1', "='Sheet Bla'!A5 + 'Sheet 3'!A6")
        wb.rename_sheet('Sheet Bla', 'Sheet1')
        self.assertEqual(wb.get_cell_contents('Sheet2', 'A1'), "=Sheet1!A5 + 'Sheet 3'!A6")
        
        # Formula-updates are only performed on formulas affected by the sheet-rename operation, unparseable formulas untouched
        wb2 = sheets.Workbook()
        wb2.new_sheet()
        wb2.new_sheet()

        wb2.set_cell_contents('Sheet2', 'A1', "=A4 + Sheet5!A7 + 'B1'")
        wb2.set_cell_contents('Sheet2', 'A2', '34')
        wb2.set_cell_contents('Sheet2', 'A3', "=Sheet1!A1 + 4")
        wb2.set_cell_contents('Sheet2', 'A4', "=Sheet1!A4 + Sheet1!A") # unparseable
        wb2.rename_sheet('Sheet1', 'SheetBla')

        self.assertEqual(wb2.get_cell_contents('Sheet2', 'A1'), "=A4 + Sheet5!A7 + 'B1'")
        self.assertEqual(wb2.get_cell_contents('Sheet2', 'A2'), "34")
        self.assertEqual(wb2.get_cell_contents('Sheet2', 'A3'), "=SheetBla!A1 + 4")
        self.assertEqual(wb2.get_cell_contents('Sheet2', 'A4'), "=Sheet1!A4 + Sheet1!A") # unchanged
        
        # KeyError, ValueError
        with self.assertRaises(KeyError):
            wb2.rename_sheet('Sheet7', 'Sheet15')
        
        wb2.new_sheet('Lorem ipsum')
        invalid_sheet_names = ['', ' Sheet', '~', 'Lorem ipsum', 'Sheet\' name', 'Sheet \" name']
        for sheet_name in invalid_sheet_names:
            with self.assertRaises(ValueError):
                wb2.rename_sheet('Sheet2', sheet_name)
    
    def test_rename_sheet_update(self):
        # Renaming a sheet may cause some invalid cell-reference to become valid. 
        # Make sure that all cell updates are performed correctly, even in these unusual cases.
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.new_sheet()

        wb.set_cell_contents('Sheet1', 'A1', '=SheetBla!A1')
        wb.set_cell_contents('Sheet2', 'A1', '5')
        wb.rename_sheet('Sheet2', 'SheetBla')

        self.assertEqual(wb.get_cell_value('Sheet1', 'A1'), 5)
    
    def test_move_cells(self):
        # KeyError, ValueError
        wb_0 = sheets.Workbook()
        with self.assertRaises(KeyError):
            wb_0.move_cells('Sheet7', 'A1', 'B1', 'C1')
        
        wb_0.new_sheet()
        with self.assertRaises(ValueError):
            wb_0.move_cells('Sheet1', 'A1', 'B1', 'ZZZZ998')

        # very basic test case
        wb = sheets.Workbook()
        wb.new_sheet()

        wb.set_cell_contents("Sheet1", "A1", "1")
        wb.set_cell_contents("Sheet1", "A2", "=D1")
        wb.set_cell_contents("Sheet1", "B1", "=C1")
        wb.set_cell_contents("Sheet1", "B2", "3")

        wb.move_cells("Sheet1", "A1", "B2", "C1", None)

        cell_d1 = wb.get_cell("sheet1", "D1")
        cell_c2 = wb.get_cell("sheet1", "C2")
        self.assertEqual(cell_d1.contents, "=E1" )
        self.assertEqual(cell_c2.contents, "=F1" )
        
        # test for to_sheet not being None
        # TODO: not sure
        wb.new_sheet()
        wb.set_cell_contents("Sheet1", "A1", "1")
        wb.set_cell_contents("Sheet1", "A2", "=D1")
        wb.set_cell_contents("Sheet1", "B1", "=C1")
        wb.set_cell_contents("Sheet1", "B2", "3")
        wb.move_cells("Sheet1", "A1", "B2", "C1", "Sheet2")

        self.assertEqual(wb.get_cell_contents("Sheet2", "C1"), "1")
        self.assertEqual(wb.get_cell_contents("Sheet2", "C2"), "=F1")
        self.assertEqual(wb.get_cell_contents("Sheet2", "D1"), "=E1")
        self.assertEqual(wb.get_cell_contents("Sheet2", "D2"), "3")

        # test for overlapping
        wb_2 = sheets.Workbook()
        wb_2.new_sheet()

        # Set up overlapping data
        wb_2.set_cell_contents("Sheet1", "A1", "1")
        wb_2.set_cell_contents("Sheet1", "A2", "2")
        wb_2.set_cell_contents("Sheet1", "B1", "3")
        wb_2.set_cell_contents("Sheet1", "B2", "4")

        # Move cells A1-B2 to B1 (overlapping move within the same sheet)
        wb_2.move_cells("Sheet1", "A1", "B2", "B1")

        # Check if the cells were moved correctly (accounting for overlap)
        self.assertEqual(wb_2.get_cell_contents("Sheet1", "A1"), None)
        self.assertEqual(wb_2.get_cell_contents("Sheet1", "A2"), None)
        self.assertEqual(wb_2.get_cell_contents("Sheet1", "B1"), "1")
        self.assertEqual(wb_2.get_cell_contents("Sheet1", "B2"), "2")
        self.assertEqual(wb_2.get_cell_contents("Sheet1", "C1"), "3")
        self.assertEqual(wb_2.get_cell_contents("Sheet1", "C2"), "4")


if __name__ == "__main__":
    cov = coverage.Coverage()
    cov.start()
    unittest.main()
    cov.stop()
    cov.save()
    cov.html_report()