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

        # other basic test
        wb2 = sheets.Workbook()
        wb2.new_sheet()
        wb2.set_cell_contents('Sheet1', 'A1', '=Sheet1!A2')
        wb2.rename_sheet('Sheet1', 'Sheet8')
        self.assertEqual(wb2.get_cell_contents('Sheet8', 'A1'), '=Sheet8!A2')
        
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

        wb.new_sheet()
        wb.set_cell_contents("Sheet2", "A1", "1")
        wb.set_cell_contents("Sheet2", "B1", "=A1")

        wb.move_cells("Sheet2", "A1", "B1", "C3", None)

        cell_c3 = wb.get_cell("sheet2", "C3")
        cell_d3 = wb.get_cell("sheet2", "D3")
        self.assertEqual(cell_c3.contents, "1")
        self.assertEqual(cell_d3.contents, "=C3" )        
        
        # test for to_sheet not being None
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

        # test absolute
        wb_3 = sheets.Workbook()
        wb_3.new_sheet()
        wb_3.set_cell_contents("Sheet1", "A1", "1")
        wb_3.set_cell_contents("Sheet1", "B1", "=$A$1")

        wb_3.move_cells("Sheet1", "A1", "B1", "C1")
        self.assertEqual(wb_3.get_cell_contents("Sheet1", "D1"), "=$A$1")
        self.assertEqual(wb_3.get_cell_value("Sheet1", "D1"), decimal.Decimal('0'))

        # test mixed reference whrere column is not modified
        wb_3.new_sheet()
        wb_3.set_cell_contents("Sheet2", "A1", "1")
        wb_3.set_cell_contents("Sheet2", "B2", "=$A1")
        
        wb_3.move_cells("Sheet2", "A1", "B2", "B3")
        self.assertEqual(wb_3.get_cell_contents("Sheet2", "C4"), "=$A3")

        # test mixed reference whrere row is not modified
        wb_3.new_sheet()
        wb_3.set_cell_contents("Sheet3", "A1", "1")
        wb_3.set_cell_contents("Sheet3", "B2", "=A$1")
        
        wb_3.move_cells("Sheet3", "A1", "B2", "B3")
        self.assertEqual(wb_3.get_cell_contents("Sheet3", "C4"), "=B$1")

        # test moving backwards (-delta_x and -delta_y)
        wb_4 = sheets.Workbook()
        wb_4.new_sheet()
        wb_4.set_cell_contents("Sheet1", "C1", "=$D1")
        wb_4.set_cell_contents("Sheet1", "D1", "=1")

        wb_4.move_cells("Sheet1", "C1", "D1", "A2")
        self.assertEqual(wb_4.get_cell_contents("Sheet1", "A2"), "=$D2")

        # test using other corners
        wb_4.new_sheet()
        wb_4.set_cell_contents("Sheet2", "A1", "=1")
        wb_4.set_cell_contents("Sheet2", "B3", "=A1")

        wb_4.move_cells("Sheet2", "B1", "A3", "C1")
        self.assertEqual(wb_4.get_cell_contents("Sheet2", "D3"), "=C1")

        # move only one cell
        wb_4.set_cell_contents("Sheet2", "F1", "=1")
        wb_4.set_cell_contents("Sheet2", "G1", "=F1")

        wb_4.move_cells("Sheet2", "G1", "G1", "H1")
        self.assertEqual(wb_4.get_cell_contents("Sheet2", "H1"), "=G1")

        # test #REF
        wb_5 = sheets.Workbook()
        wb_5.new_sheet()
        wb_5.set_cell_contents("Sheet1", "A1", "=2.2")
        wb_5.set_cell_contents("Sheet1", "A2", "=4.5")
        wb_5.set_cell_contents("Sheet1", "B1", "=5.3")
        wb_5.set_cell_contents("Sheet1", "B2", "=3.1")
        wb_5.set_cell_contents("Sheet1", "C1", "=A1*B1")
        wb_5.set_cell_contents("Sheet1", "C2", "=A2*B2")

        wb_5.move_cells("Sheet1", "C1", "C2", "B1")
        self.assertEqual(wb_5.get_cell_contents('Sheet1', 'B1'), "=#REF! * A1")
        self.assertEqual(wb_5.get_cell_contents('Sheet1', 'B2'), "=#REF! * A2")    

        wb = sheets.Workbook()
        wb.new_sheet()

        wb.set_cell_contents('sheet1', 'A1', '=B1 + B2 + B3 + B4')    

        wb.move_cells("Sheet1", "A1", "A1", "A9999")
        self.assertEqual(wb.get_cell_contents('Sheet1', 'A9999'), "=B9999 + #REF! + #REF! + #REF!")


    def test_copy_cells(self):
        wb_0 = sheets.Workbook()
        with self.assertRaises(KeyError):
            wb_0.move_cells('Sheet1', 'A1', 'B1', 'C1')
        
        wb_0.new_sheet()
        with self.assertRaises(ValueError):
            wb_0.move_cells('Sheet1', 'A1', 'B1', 'ZZZZ9998')
        
        wb = sheets.Workbook()
        wb.new_sheet()

        wb.set_cell_contents("Sheet1", "A1", "1")
        wb.set_cell_contents("Sheet1", "A2", "=D1")
        wb.set_cell_contents("Sheet1", "B1", "=C1")
        wb.set_cell_contents("Sheet1", "B2", "3")
        wb.set_cell_contents("Sheet1", "D1", "=2")
        wb.set_cell_contents("Sheet1", "C1", "=2")

        wb.copy_cells("Sheet1", "A1", "B2", "B1", None)

        self.assertEqual(wb.get_cell_value("Sheet1", "A1"), decimal.Decimal('1'))
        self.assertEqual(wb.get_cell_value("Sheet1", "A2"), decimal.Decimal('2'))

        self.assertEqual(wb.get_cell_contents("Sheet1", "B2"), "=E1")
        self.assertEqual(wb.get_cell_contents("Sheet1", "B1"), "1")
        self.assertEqual(wb.get_cell_contents("Sheet1", "B2"), "=E1")
        self.assertEqual(wb.get_cell_contents("Sheet1", "C1"), "=D1")
        self.assertEqual(wb.get_cell_contents("Sheet1", "C2"), "3")

        # TEST #REF!
        wb.new_sheet()
        wb.set_cell_contents("Sheet2", "B3", "1")
        wb.set_cell_contents("Sheet2", "C3", "=C1")
        wb.set_cell_contents("Sheet2", "B4", "=D1")
        wb.set_cell_contents("Sheet2", "C4", "3")
        wb.set_cell_contents("Sheet2", "C1", "1")
        wb.set_cell_contents("Sheet2", "D1", "2")

        wb.copy_cells("Sheet2", "B3", "C4", "A2", None)

        self.assertEqual(wb.get_cell_contents("Sheet2", "A2"), "1")
        self.assertEqual(wb.get_cell_contents("Sheet2", "A3"), '=#REF!')
        self.assertEqual(wb.get_cell_contents("Sheet2", "B2"), '=#REF!')
        self.assertEqual(wb.get_cell_contents("Sheet2", "B3"), "3")

        self.assertEqual(wb.get_cell_value("Sheet2", "C3"), decimal.Decimal('1'))
        self.assertEqual(wb.get_cell_value("Sheet2", "B4"), decimal.Decimal('2'))
        self.assertEqual(wb.get_cell_value("Sheet2", "C4"), decimal.Decimal('3'))

        # test for overlapping
        wb_2 = sheets.Workbook()
        wb_2.new_sheet()

        # Set up overlapping data
        wb_2.set_cell_contents("Sheet1", "A1", "1")
        wb_2.set_cell_contents("Sheet1", "A2", "2")
        wb_2.set_cell_contents("Sheet1", "B1", "3")
        wb_2.set_cell_contents("Sheet1", "B2", "4")

        # Move cells A1-B2 to B1 (overlapping move within the same sheet)
        wb_2.copy_cells("Sheet1", "A1", "B2", "B1")

        # Check if the cells were moved correctly (accounting for overlap)
        self.assertEqual(wb_2.get_cell_value("Sheet1", "A1"), decimal.Decimal('1'))
        self.assertEqual(wb_2.get_cell_value("Sheet1", "A2"), decimal.Decimal('2'))
        self.assertEqual(wb_2.get_cell_contents("Sheet1", "B1"), "1")
        self.assertEqual(wb_2.get_cell_contents("Sheet1", "B2"), "2")
        self.assertEqual(wb_2.get_cell_contents("Sheet1", "C1"), "3")
        self.assertEqual(wb_2.get_cell_contents("Sheet1", "C2"), "4")

        # test absolute
        wb_3 = sheets.Workbook()
        wb_3.new_sheet()
        wb_3.set_cell_contents("Sheet1", "A1", "1")
        wb_3.set_cell_contents("Sheet1", "B1", "=$A$1")

        wb_3.copy_cells("Sheet1", "A1", "B1", "C1")
        self.assertEqual(wb_3.get_cell_contents("Sheet1", "D1"), "=$A$1")
        self.assertEqual(wb_3.get_cell_value("Sheet1", "D1"), decimal.Decimal('1'))
        self.assertEqual(wb_3.get_cell_value("Sheet1", "A1"), decimal.Decimal('1'))

        # test mixed reference where column is absolute
        wb_3.new_sheet()
        wb_3.set_cell_contents('Sheet2', 'A1', '=$B4')
        wb_3.set_cell_contents('Sheet2', 'B4', '1')

        wb_3.copy_cells('Sheet2', 'A1', 'A1', 'D7')
        wb_3.set_cell_contents('Sheet2', 'B10', '4')
        self.assertEqual(wb_3.get_cell_contents('Sheet2', 'D7'), '=$B10')
        self.assertEqual(wb_3.get_cell_value('Sheet2', 'D7'), decimal.Decimal('4'))

        # test mixed reference where row is absolute
        wb_3.new_sheet()
        wb_3.set_cell_contents('Sheet3', 'A1', '=C$4')
        wb_3.set_cell_contents('Sheet3', 'C4', '1')

        wb_3.copy_cells('Sheet3', 'A1', 'A1', 'D7')
        wb_3.set_cell_contents('Sheet3', 'F4', '4')
        self.assertEqual(wb_3.get_cell_contents('Sheet3', 'D7'), '=F$4')
        self.assertEqual(wb_3.get_cell_value('Sheet3', 'D7'), decimal.Decimal('4'))

        # test moving backwards (-delta_x and -delta_y)
        wb_4 = sheets.Workbook()
        wb_4.new_sheet()
        wb_4.set_cell_contents("Sheet1", "C1", "=$D1")
        wb_4.set_cell_contents("Sheet1", "D1", "=1")

        wb_4.copy_cells("Sheet1", "C1", "D1", "A2")
        self.assertEqual(wb_4.get_cell_contents("Sheet1", "A2"), "=$D2")
        # self.assertEqual(wb_3.get_cell_value("Sheet1", "A2"), decimal.Decimal('1')) # TODO: idk how this supp. to work

        # test using other corners
        wb_4.new_sheet()
        wb_4.set_cell_contents("Sheet2", "A1", "=1")
        wb_4.set_cell_contents("Sheet2", "B3", "=A1")

        wb_4.copy_cells("Sheet2", "B1", "A3", "C1")
        self.assertEqual(wb_4.get_cell_contents("Sheet2", "D3"), "=C1")

        # test #REF (given test case)
        wb_5 = sheets.Workbook()
        wb_5.new_sheet()
        wb_5.set_cell_contents("Sheet1", "A1", "=2.2")
        wb_5.set_cell_contents("Sheet1", "A2", "=4.5")
        wb_5.set_cell_contents("Sheet1", "B1", "=5.3")
        wb_5.set_cell_contents("Sheet1", "B2", "=3.1")
        wb_5.set_cell_contents("Sheet1", "C1", "=A1*B1")
        wb_5.set_cell_contents("Sheet1", "C2", "=A2*B2")

        wb_5.move_cells("Sheet1", "C1", "C2", "B1")
        self.assertEqual(wb_5.get_cell_contents('Sheet1', 'B1'), "=#REF! * A1")
        self.assertEqual(wb_5.get_cell_contents('Sheet1', 'B2'), "=#REF! * A2")

        # Simple, given test case
        wb_6 = sheets.Workbook()
        wb_6.new_sheet()
        wb_6.set_cell_contents("Sheet1", "A1", "'123")
        wb_6.set_cell_contents("Sheet1", "B1", "5.3")
        wb_6.set_cell_contents("Sheet1", "C1", "=A1*B1")

        wb_6.copy_cells("Sheet1", "A1", "C1", "A2")
        self.assertEqual(wb_6.get_cell_contents('Sheet1', 'A2'), "'123")
        self.assertEqual(wb_6.get_cell_contents('Sheet1', 'B2'), "5.3")
        self.assertEqual(wb_6.get_cell_contents('Sheet1', 'C2'), "=A2 * B2")

        wb_6.new_sheet()
        wb_6.set_cell_contents("Sheet2", "A1", "'123")
        wb_6.set_cell_contents("Sheet2", "B1", "5.3")
        wb_6.set_cell_contents("Sheet2", "C1", "=A1*B1")

        wb_6.copy_cells("Sheet2", "A1", "C1", "B2")
        self.assertEqual(wb_6.get_cell_contents('Sheet2', 'B2'), "'123")
        self.assertEqual(wb_6.get_cell_contents('Sheet2', 'C2'), "5.3")
        self.assertEqual(wb_6.get_cell_contents('Sheet2', 'D2'), "=B2 * C2")

        wb = sheets.Workbook()
        wb.new_sheet()

        wb.set_cell_contents('sheet1', 'A1', '=B1 + B2 + B3 + B4')    

        wb.copy_cells("Sheet1", "A1", "A1", "A9999")
        self.assertEqual(wb.get_cell_contents('Sheet1', 'A9999'), "=B9999 + #REF! + #REF! + #REF!")

    def test_move_cells_edge(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.new_sheet()

        # to another sheet
        wb.set_cell_contents('Sheet1', 'A1', '4')
        wb.set_cell_contents('Sheet1', 'A2', '5')
        wb.set_cell_contents('Sheet1', 'B1', '=D4 / E8') # D4: (+2, +3), E8: (+3, +7)
        wb.set_cell_contents('Sheet1', 'B2', '=A1 - A2') # A1: (-1, -1), A2: (-1, 0)
        wb.move_cells('Sheet1', 'A1', 'B2', 'D6', 'Sheet2')

        self.assertEqual(wb.get_sheet_extent('Sheet1'), (0, 0))

        self.assertEqual(wb.get_cell_contents('Sheet2', 'D6'), '4')
        self.assertEqual(wb.get_cell_contents('Sheet2', 'D7'), '5')
        self.assertEqual(wb.get_cell_contents('Sheet2', 'E6'), '=G9 / H13')
        self.assertEqual(wb.get_cell_contents('Sheet2', 'E7'), '=D6 - D7')

        # to another sheet that doesn't exist
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('Sheet1', 'A1', '1')
        wb.set_cell_contents('Sheet1', 'B1', '1')

        with self.assertRaises(KeyError):
            wb.move_cells('Sheet1', 'A1', 'B1', 'A1', 'Sheet2')

        # different corners (same case as above)
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.new_sheet()
        wb.new_sheet()
        wb.set_cell_contents('Sheet1', 'A1', '4')
        wb.set_cell_contents('Sheet1', 'B1', '=D4 / E8') # D4: (+2, +3), E8: (+3, +7)
        wb.set_cell_contents('Sheet1', 'B2', '=A1 - A2') # A1: (-1, -1), A2: (-1, 0)
        wb.move_cells('Sheet1', 'A2', 'B1', 'D6', 'Sheet3')

        self.assertEqual(wb.get_sheet_extent('Sheet1'), (0, 0))

        self.assertEqual(wb.get_cell_contents('Sheet3', 'D6'), '4')
        self.assertEqual(wb.get_cell_contents('Sheet3', 'D7'), None)
        self.assertEqual(wb.get_cell_contents('Sheet3', 'E6'), '=G9 / H13')
        self.assertEqual(wb.get_cell_contents('Sheet3', 'E7'), '=D6 - D7')

        # test None cells
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.new_sheet()
        wb.new_sheet()
        wb.set_cell_contents('Sheet1', 'B1', '=D4 / E8') # D4: (+2, +3), E8: (+3, +7)
        wb.set_cell_contents('Sheet1', 'B2', '=A1 - A2') # A1: (-1, -1), A2: (-1, 0)
        wb.move_cells('Sheet1', 'A2', 'B1', 'D6', 'Sheet3')

        self.assertEqual(wb.get_sheet_extent('Sheet1'), (0, 0))

        self.assertEqual(wb.get_cell_contents('Sheet3', 'D6'), None)
        self.assertEqual(wb.get_cell_contents('Sheet3', 'D7'), None)
        self.assertEqual(wb.get_cell_contents('Sheet3', 'E6'), '=G9 / H13')
        self.assertEqual(wb.get_cell_contents('Sheet3', 'E7'), '=D6 - D7')

        # overlap
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.new_sheet()
        wb.set_cell_contents('Sheet1', 'A1', '4')
        wb.set_cell_contents('Sheet1', 'A2', '5')
        wb.set_cell_contents('Sheet1', 'B1', '=D4 / E8') # D4: (+2, +3), E8: (+3, +7)
        wb.set_cell_contents('Sheet1', 'B2', '=A1 - A2') # A1: (-1, -1), A2: (-1, 0)
        wb.move_cells('Sheet1', 'A2', 'B1', 'B1')

        self.assertEqual(wb.get_cell_contents('Sheet1', 'B1'), '4')
        self.assertEqual(wb.get_cell_contents('Sheet1', 'B2'), '5')
        self.assertEqual(wb.get_cell_contents('Sheet1', 'C1'), '=E4 / F8')
        self.assertEqual(wb.get_cell_contents('Sheet1', 'C2'), '=B1 - B2')

        # value error for moving cell outside of bounds
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.new_sheet()
        wb.set_cell_contents('Sheet1', 'A1', '4')
        wb.set_cell_contents('Sheet1', 'ZZ50', '=A1 - A2') # A1: (-1, -1), A2: (-1, 0)
        with self.assertRaises(ValueError):
            wb.move_cells('Sheet1', 'A1', 'ZZZ100', 'A9990')
        with self.assertRaises(ValueError):
            wb.move_cells('Sheet1', 'A1', 'ZZZ100', 'ZZZA1')

        # references have sheet name
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.new_sheet()
        wb.new_sheet('Sheet 3')
        wb.set_cell_contents('Sheet1', 'A1', '4')
        wb.set_cell_contents('Sheet1', 'A2', '5')
        wb.set_cell_contents('Sheet1', 'B1', '=Sheet2!D4 / E8') # D4: (+2, +3), E8: (+3, +7)
        wb.set_cell_contents('Sheet1', 'B2', '=A1 - \'Sheet 3\'!A2') # A1: (-1, -1), A2: (-1, 0)
        wb.move_cells('Sheet1', 'A1', 'B2', 'D6')

        self.assertEqual(wb.get_cell_contents('Sheet1', 'D6'), '4')
        self.assertEqual(wb.get_cell_contents('Sheet1', 'D7'), '5')
        self.assertEqual(wb.get_cell_contents('Sheet1', 'E6'), '=Sheet2!G9 / H13')
        self.assertEqual(wb.get_cell_contents('Sheet1', 'E7'), '=D6 - \'Sheet 3\'!D7')

        # #REF! in negative direction
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('Sheet1', 'D4', '=A4')
        wb.set_cell_contents('Sheet1', 'D5', '=D1')
        wb.move_cells('Sheet1', 'D5', 'D4', 'A1')

        self.assertEqual(wb.get_cell_contents('Sheet1', 'A1'), '=#REF!')
        self.assertEqual(wb.get_cell_contents('Sheet1', 'A2'), '=#REF!')

        # #REF! in too large direction
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('Sheet1', 'A1', '=AAA1')
        wb.move_cells('Sheet1', 'A1', 'A1', 'ZZZA1')

        self.assertEqual(wb.get_cell_contents('Sheet1', 'ZZZA1'), '=#REF!')

        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('Sheet1', 'A1', '=A1000')
        wb.move_cells('Sheet1', 'A1', 'A1', 'A9990')

        self.assertEqual(wb.get_cell_contents('Sheet1', 'A9990'), '=#REF!')

        # absolute references across sheets
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.new_sheet()
        wb.set_cell_contents('Sheet1', 'A1', '=$BB$100')
        wb.set_cell_contents('Sheet1', 'BB100', '=1')

        wb.move_cells('Sheet1', 'A1', 'A1', 'A1', 'Sheet2')
        self.assertEqual(wb.get_cell_contents('Sheet2', 'A1'), '=$BB$100')
        self.assertEqual(wb.get_cell_value('Sheet2', 'A1'), decimal.Decimal('0'))

        # mixed references
        wb = sheets.Workbook()
        wb.new_sheet()

        wb.set_cell_contents('Sheet1', 'A1', '=$B1 + B$2')
        wb.move_cells('Sheet1', 'A1', 'A1', 'C2')

        self.assertEqual(wb.get_cell_contents('Sheet1', 'C2'), '=$B2 + D$2')

        # error propagation
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('Sheet1', 'A1', '=B1')
        wb.set_cell_contents('Sheet1', 'B1', '=A1')
        
        wb.move_cells('Sheet1', 'A1', 'B1', 'C1')
        self.assertIsInstance(wb.get_cell_value('Sheet1', 'C1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'C1').get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)

        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('Sheet1', 'A1', '=1')
        wb.set_cell_contents('Sheet1', 'B1', '=A99999')
        
        wb.move_cells('Sheet1', 'A1', 'B1', 'C1')
        self.assertIsInstance(wb.get_cell_value('Sheet1', 'D1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'D1').get_type(), sheets.CellErrorType.BAD_REFERENCE)

        # complete overlap
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('Sheet1', 'A1', '=1')
        wb.set_cell_contents('Sheet1', 'B1', '=$A2')
        
        wb.move_cells('Sheet1', 'A1', 'B1', 'A1')
        self.assertEqual(wb.get_cell_contents('Sheet1', 'A1'), '=1')
        self.assertEqual(wb.get_cell_contents('Sheet1', 'B1'), '=$A2')

    def test_copy_cells_edge(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.new_sheet()

        # to another sheet
        wb.set_cell_contents('Sheet1', 'A1', '4')
        wb.set_cell_contents('Sheet1', 'A2', '5')
        wb.set_cell_contents('Sheet1', 'B1', '=D4 / E8') # D4: (+2, +3), E8: (+3, +7)
        wb.set_cell_contents('Sheet1', 'B2', '=A1 - A2') # A1: (-1, -1), A2: (-1, 0)
        wb.copy_cells('Sheet1', 'A1', 'B2', 'D6', 'Sheet2')

        self.assertEqual(wb.get_cell_contents('Sheet2', 'D6'), '4')
        self.assertEqual(wb.get_cell_contents('Sheet2', 'D7'), '5')
        self.assertEqual(wb.get_cell_contents('Sheet2', 'E6'), '=G9 / H13')
        self.assertEqual(wb.get_cell_contents('Sheet2', 'E7'), '=D6 - D7')

        # sheet that doesn't exist
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('Sheet1', 'A1', '1')
        wb.set_cell_contents('Sheet1', 'B1', '1')

        with self.assertRaises(KeyError):
            wb.move_cells('Sheet1', 'A1', 'B1', 'A1', 'Sheet2')

        # different corners (same case as above)
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.new_sheet()
        wb.new_sheet()

        wb.set_cell_contents('Sheet1', 'A1', '4')
        wb.set_cell_contents('Sheet1', 'B1', '=D4 / E8') # D4: (+2, +3), E8: (+3, +7)
        wb.set_cell_contents('Sheet1', 'B2', '=A1 - A2') # A1: (-1, -1), A2: (-1, 0)
        wb.copy_cells('Sheet1', 'A2', 'B1', 'D6', 'Sheet3')

        self.assertEqual(wb.get_cell_contents('Sheet3', 'D6'), '4')
        self.assertEqual(wb.get_cell_contents('Sheet3', 'D7'), None)
        self.assertEqual(wb.get_cell_contents('Sheet3', 'E6'), '=G9 / H13')
        self.assertEqual(wb.get_cell_contents('Sheet3', 'E7'), '=D6 - D7')

        # test None cells
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.new_sheet()
        wb.new_sheet()
        wb.set_cell_contents('Sheet1', 'B1', '=D4 / E8') # D4: (+2, +3), E8: (+3, +7)
        wb.set_cell_contents('Sheet1', 'B2', '=A1 - A2') # A1: (-1, -1), A2: (-1, 0)
        wb.copy_cells('Sheet1', 'A2', 'B1', 'D6', 'Sheet3')

        self.assertEqual(wb.get_cell_contents('Sheet3', 'D6'), None)
        self.assertEqual(wb.get_cell_contents('Sheet3', 'D7'), None)
        self.assertEqual(wb.get_cell_contents('Sheet3', 'E6'), '=G9 / H13')
        self.assertEqual(wb.get_cell_contents('Sheet3', 'E7'), '=D6 - D7')

        # overlap
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.new_sheet()
        wb.set_cell_contents('Sheet1', 'A1', '4')
        wb.set_cell_contents('Sheet1', 'A2', '5')
        wb.set_cell_contents('Sheet1', 'B1', '=D4 / E8') # D4: (+2, +3), E8: (+3, +7)
        wb.set_cell_contents('Sheet1', 'B2', '=A1 - A2') # A1: (-1, -1), A2: (-1, 0)
        wb.copy_cells('Sheet1', 'A2', 'B1', 'B1')

        self.assertEqual(wb.get_cell_contents('Sheet1', 'A1'), '4')
        self.assertEqual(wb.get_cell_contents('Sheet1', 'A2'), '5')
        self.assertEqual(wb.get_cell_contents('Sheet1', 'B1'), '4')
        self.assertEqual(wb.get_cell_contents('Sheet1', 'B2'), '5')
        self.assertEqual(wb.get_cell_contents('Sheet1', 'C1'), '=E4 / F8')
        self.assertEqual(wb.get_cell_contents('Sheet1', 'C2'), '=B1 - B2')

        # value error for moving cell outside of bounds
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.new_sheet()
        wb.set_cell_contents('Sheet1', 'A1', '4')
        wb.set_cell_contents('Sheet1', 'ZZ50', '=A1 - A2') # A1: (-1, -1), A2: (-1, 0)
        with self.assertRaises(ValueError):
            wb.copy_cells('Sheet1', 'A1', 'ZZZ100', 'A9990')
        with self.assertRaises(ValueError):
            wb.copy_cells('Sheet1', 'A1', 'ZZZ100', 'ZZZA1')

        # references have sheet name
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.new_sheet()
        wb.new_sheet('Sheet 3')
        wb.set_cell_contents('Sheet1', 'A1', '4')
        wb.set_cell_contents('Sheet1', 'A2', '5')
        wb.set_cell_contents('Sheet1', 'B1', '=Sheet2!D4 / E8') # D4: (+2, +3), E8: (+3, +7)
        wb.set_cell_contents('Sheet1', 'B2', '=A1 - \'Sheet 3\'!A2') # A1: (-1, -1), A2: (-1, 0)
        wb.copy_cells('Sheet1', 'A1', 'B2', 'D6')

        self.assertEqual(wb.get_cell_contents('Sheet1', 'D6'), '4')
        self.assertEqual(wb.get_cell_contents('Sheet1', 'D7'), '5')
        self.assertEqual(wb.get_cell_contents('Sheet1', 'E6'), '=Sheet2!G9 / H13')
        self.assertEqual(wb.get_cell_contents('Sheet1', 'E7'), '=D6 - \'Sheet 3\'!D7')

        # #REF! in negative direction
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('Sheet1', 'D4', '=A4')
        wb.set_cell_contents('Sheet1', 'D5', '=D1')
        wb.copy_cells('Sheet1', 'D5', 'D4', 'A1')

        self.assertEqual(wb.get_cell_contents('Sheet1', 'A1'), '=#REF!')
        self.assertEqual(wb.get_cell_contents('Sheet1', 'A2'), '=#REF!')

        # #REF! in too large direction
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('Sheet1', 'A1', '=AAA1')
        wb.copy_cells('Sheet1', 'A1', 'A1', 'ZZZA1')

        self.assertEqual(wb.get_cell_contents('Sheet1', 'ZZZA1'), '=#REF!')

        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('Sheet1', 'A1', '=A1000')
        wb.copy_cells('Sheet1', 'A1', 'A1', 'A9990')

        self.assertEqual(wb.get_cell_contents('Sheet1', 'A9990'), '=#REF!')

        # absolute references across sheets
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.new_sheet()
        wb.set_cell_contents('Sheet1', 'A1', '=$BB$100')
        wb.set_cell_contents('Sheet1', 'BB100', '=1')

        wb.move_cells('Sheet1', 'A1', 'A1', 'A1', 'Sheet2')
        self.assertEqual(wb.get_cell_contents('Sheet2', 'A1'), '=$BB$100')
        self.assertEqual(wb.get_cell_value('Sheet2', 'A1'), decimal.Decimal('0'))

        # mixed references
        wb = sheets.Workbook()
        wb.new_sheet()

        wb.set_cell_contents('Sheet1', 'A1', '=$B1 + B$2')
        wb.move_cells('Sheet1', 'A1', 'A1', 'C2')

        self.assertEqual(wb.get_cell_contents('Sheet1', 'C2'), '=$B2 + D$2')

        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('Sheet1', 'A1', '=B1')
        wb.set_cell_contents('Sheet1', 'B1', '=A1')
        
        wb.move_cells('Sheet1', 'A1', 'B1', 'C1')
        self.assertIsInstance(wb.get_cell_value('Sheet1', 'C1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'C1').get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)

        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('Sheet1', 'A1', '=1')
        wb.set_cell_contents('Sheet1', 'B1', '=A99999')
        
        wb.move_cells('Sheet1', 'A1', 'B1', 'C1')
        self.assertIsInstance(wb.get_cell_value('Sheet1', 'D1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'D1').get_type(), sheets.CellErrorType.BAD_REFERENCE)

        # complete overlap
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('Sheet1', 'A1', '=1')
        wb.set_cell_contents('Sheet1', 'B1', '=$A2')
        
        wb.move_cells('Sheet1', 'A1', 'B1', 'A1')
        self.assertEqual(wb.get_cell_contents('Sheet1', 'A1'), '=1')
        self.assertEqual(wb.get_cell_contents('Sheet1', 'B1'), '=$A2')

        # copying empty cells
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('Sheet1', 'D1', '=1')
        wb.set_cell_contents('Sheet1', 'E1', '=2')
        wb.set_cell_contents('Sheet1', 'D2', '=3')
        wb.set_cell_contents('Sheet1', 'E2', '=4')

        wb.move_cells('Sheet1', 'A1', 'B2', 'D1')
        self.assertEqual(wb.get_cell_contents('Sheet1', 'D1'), None)
        self.assertEqual(wb.get_cell_contents('Sheet1', 'E1'), None)
        self.assertEqual(wb.get_cell_contents('Sheet1', 'D2'), None)
        self.assertEqual(wb.get_cell_contents('Sheet1', 'E2'), None)
    
    def test_sort(self):
        wb = sheets.Workbook()
        wb.new_sheet()

        # basic test
        wb.set_cell_contents('Sheet1', 'A1', 'Alice')
        wb.set_cell_contents('Sheet1', 'A2', 'Bob')
        wb.set_cell_contents('Sheet1', 'A3', 'Charlie')
        # wb.set_cell_contents('Sheet1', 'D1', '=25')

        wb.set_cell_contents('Sheet1', 'B1', '=25')
        wb.set_cell_contents('Sheet1', 'B2', '=30')
        wb.set_cell_contents('Sheet1', 'B3', '=25')

        wb.set_cell_contents('Sheet1', 'C1', 'Engineer')
        wb.set_cell_contents('Sheet1', 'C2', 'Designer')
        wb.set_cell_contents('Sheet1', 'C3', 'Manager')

        wb.sort_region('Sheet1', 'A1', 'C3', [2, -1])

        self.assertEqual(wb.get_cell_value('sheet1', 'A1'), 'Charlie')
        self.assertEqual(wb.get_cell_value('sheet1', 'A2'), 'Alice')
        self.assertEqual(wb.get_cell_value('sheet1', 'A3'), 'Bob')

        self.assertEqual(wb.get_cell_value('sheet1', 'B1'), decimal.Decimal('25'))
        self.assertEqual(wb.get_cell_value('sheet1', 'B2'), decimal.Decimal('25'))
        self.assertEqual(wb.get_cell_value('sheet1', 'B3'), decimal.Decimal('30'))

        self.assertEqual(wb.get_cell_value('sheet1', 'C1'), 'Manager')
        self.assertEqual(wb.get_cell_value('sheet1', 'C2'), 'Engineer')
        self.assertEqual(wb.get_cell_value('sheet1', 'C3'), 'Designer')

        wb = sheets.Workbook()
        wb.new_sheet()

        # basic test
        wb.set_cell_contents('Sheet1', 'A1', 'Alice')
        wb.set_cell_contents('Sheet1', 'A2', 'Bob')
        wb.set_cell_contents('Sheet1', 'A3', 'Charlie')
        wb.set_cell_contents('Sheet1', 'D1', '=25')
        wb.set_cell_contents('Sheet1', 'D2', '=30')
        wb.set_cell_contents('Sheet1', 'D3', '=25')

        wb.set_cell_contents('Sheet1', 'B1', '=$D$1')
        wb.set_cell_contents('Sheet1', 'B2', '=$D$2')
        wb.set_cell_contents('Sheet1', 'B3', '=$D$3')

        wb.set_cell_contents('Sheet1', 'C1', 'Engineer')
        wb.set_cell_contents('Sheet1', 'C2', 'Designer')
        wb.set_cell_contents('Sheet1', 'C3', 'Manager')

        wb.sort_region('Sheet1', 'A1', 'C3', [2, -1])

        self.assertEqual(wb.get_cell_value('sheet1', 'A1'), 'Charlie')
        self.assertEqual(wb.get_cell_value('sheet1', 'A2'), 'Alice')
        self.assertEqual(wb.get_cell_value('sheet1', 'A3'), 'Bob')

        self.assertEqual(wb.get_cell_value('sheet1', 'B1'), decimal.Decimal('25'))
        self.assertEqual(wb.get_cell_value('sheet1', 'B2'), decimal.Decimal('25'))
        self.assertEqual(wb.get_cell_value('sheet1', 'B3'), decimal.Decimal('30'))

        self.assertEqual(wb.get_cell_value('sheet1', 'C1'), 'Manager')
        self.assertEqual(wb.get_cell_value('sheet1', 'C2'), 'Engineer')
        self.assertEqual(wb.get_cell_value('sheet1', 'C3'), 'Designer')

        # test reference outside sort region
        wb = sheets.Workbook()
        wb.new_sheet()

        wb.set_cell_contents('Sheet1', 'A1', 'Alice')
        wb.set_cell_contents('Sheet1', 'A2', 'Bob')
        wb.set_cell_contents('Sheet1', 'A3', 'Charlie')
        wb.set_cell_contents('Sheet1', 'D1', '=25')

        wb.set_cell_contents('Sheet1', 'B1', '=D1')
        wb.set_cell_contents('Sheet1', 'B2', '=30')
        wb.set_cell_contents('Sheet1', 'B3', '=25')

        wb.set_cell_contents('Sheet1', 'C1', 'Engineer')
        wb.set_cell_contents('Sheet1', 'C2', 'Designer')
        wb.set_cell_contents('Sheet1', 'C3', 'Manager')

        wb.sort_region('Sheet1', 'A1', 'C3', [2, -1])

        self.assertEqual(wb.get_cell_value('sheet1', 'A1'), 'Charlie')
        self.assertEqual(wb.get_cell_value('sheet1', 'A2'), 'Alice')
        self.assertEqual(wb.get_cell_value('sheet1', 'A3'), 'Bob')

        self.assertEqual(wb.get_cell_value('sheet1', 'B1'), decimal.Decimal('25'))
        self.assertEqual(wb.get_cell_value('sheet1', 'B2'), decimal.Decimal('25'))
        self.assertEqual(wb.get_cell_value('sheet1', 'B3'), decimal.Decimal('30'))

        self.assertEqual(wb.get_cell_value('sheet1', 'C1'), 'Manager')
        self.assertEqual(wb.get_cell_value('sheet1', 'C2'), 'Engineer')
        self.assertEqual(wb.get_cell_value('sheet1', 'C3'), 'Designer')

        # test invalid sort_cols and sheet names
        invalid_sheet_names = ['', ' Sheet', '~', 'Lorem ipsum', 'Sheet\' name', 'Sheet \" name']
        for sheet_name in invalid_sheet_names:
            with self.assertRaises(KeyError):
                wb.sort_region(sheet_name, 'A1', 'C3', [2, -1])

        with self.assertRaises(ValueError):
            wb.sort_region('Sheet1', 'A1', 'C3', [0, 0, 0, 0])
        
        with self.assertRaises(ValueError):
            wb.sort_region('Sheet1', 'A1', 'C3', [2, 1, 2])
        
        with self.assertRaises(ValueError):
            wb.sort_region('Sheet1', 'A1', 'C3', [2, 4])
        
        with self.assertRaises(ValueError):
            wb.sort_region('Sheet1', 'A1', 'C3', [2, -2])

        with self.assertRaises(ValueError):
            wb.sort_region('Sheet1', 'A1', 'C3', [2, -500])
        
        with self.assertRaises(ValueError):
            wb.sort_region('Sheet1', 'A1', 'C3', [1.5])
        
        # test invalid start and end location
        with self.assertRaises(ValueError):
            wb.sort_region('Sheet1', 'ZZZZZ9999', 'ZZZZZ10000', [1.5])

        # test for sorting region that isn't a corner
        # TODO: this is failing
        wb = sheets.Workbook()
        wb.new_sheet()

        wb.set_cell_contents('Sheet1', 'A1', 'Alice')
        wb.set_cell_contents('Sheet1', 'A2', 'Bob')
        wb.set_cell_contents('Sheet1', 'A3', 'Charlie')
        wb.set_cell_contents('Sheet1', 'D1', '=25')

        wb.set_cell_contents('Sheet1', 'B1', '=D1')
        wb.set_cell_contents('Sheet1', 'B2', '=30')
        wb.set_cell_contents('Sheet1', 'B3', '=25')

        wb.set_cell_contents('Sheet1', 'C1', 'Engineer')
        wb.set_cell_contents('Sheet1', 'C2', 'Designer')
        wb.set_cell_contents('Sheet1', 'C3', 'Manager')

        wb.sort_region('Sheet1', 'A1', 'D5', [2, -1])

        self.assertEqual(wb.get_cell_value('sheet1', 'A1'), 'Charlie')
        self.assertEqual(wb.get_cell_value('sheet1', 'A2'), 'Alice')
        self.assertEqual(wb.get_cell_value('sheet1', 'A3'), 'Bob')

        self.assertEqual(wb.get_cell_value('sheet1', 'B1'), decimal.Decimal('25'))
        self.assertEqual(wb.get_cell_value('sheet1', 'B2'), decimal.Decimal('25'))
        self.assertEqual(wb.get_cell_value('sheet1', 'B3'), decimal.Decimal('30'))

        self.assertEqual(wb.get_cell_value('sheet1', 'C1'), 'Manager')
        self.assertEqual(wb.get_cell_value('sheet1', 'C2'), 'Engineer')
        self.assertEqual(wb.get_cell_value('sheet1', 'C3'), 'Designer')
        
        # another basic test, for references changing within the region
        wb = sheets.Workbook()
        wb.new_sheet()

        wb.set_cell_contents('Sheet1', 'A1', 'Alice')
        wb.set_cell_contents('Sheet1', 'A2', 'Bob')
        wb.set_cell_contents('Sheet1', 'A3', 'Charlie')
        wb.set_cell_contents('Sheet1', 'D1', '=1/0')

        wb.set_cell_contents('Sheet1', 'B1', '=D1')
        wb.set_cell_contents('Sheet1', 'B2', '=30')
        wb.set_cell_contents('Sheet1', 'B3', '=25')

        wb.set_cell_contents('Sheet1', 'C1', 'Engineer')
        wb.set_cell_contents('Sheet1', 'C2', 'Designer')
        wb.set_cell_contents('Sheet1', 'C3', 'Manager')

        wb.sort_region('Sheet1', 'A1', 'D5', [2, -1])

        self.assertEqual(wb.get_cell_value('sheet1', 'A1'), 'Charlie')
        self.assertEqual(wb.get_cell_value('sheet1', 'A2'), 'Bob')
        self.assertEqual(wb.get_cell_value('sheet1', 'A3'), 'Alice')

        self.assertEqual(wb.get_cell_value('sheet1', 'B1'), decimal.Decimal('25'))
        self.assertEqual(wb.get_cell_value('sheet1', 'B2'), decimal.Decimal('30'))
        self.assertIsInstance(wb.get_cell_value('Sheet1', 'B3'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'B3').get_type(), sheets.CellErrorType.DIVIDE_BY_ZERO)

        self.assertEqual(wb.get_cell_value('sheet1', 'C1'), 'Manager')
        self.assertEqual(wb.get_cell_value('sheet1', 'C2'), 'Designer')
        self.assertEqual(wb.get_cell_value('sheet1', 'C3'), 'Engineer')


if __name__ == "__main__":
    cov = coverage.Coverage()
    cov.start()
    unittest.main()
    cov.stop()
    cov.save()
    cov.html_report()