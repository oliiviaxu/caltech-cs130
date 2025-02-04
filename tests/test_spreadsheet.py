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

if __name__ == "__main__":
    cov = coverage.Coverage()
    cov.start()
    unittest.main()
    cov.stop()
    cov.save()
    cov.html_report()