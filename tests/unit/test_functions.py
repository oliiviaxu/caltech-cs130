# python -m unittest tests.unit.test_functions.FunctionsTests

import unittest
import coverage
import sheets
import os
import lark
from sheets.interpreter import FormulaEvaluator
import decimal
import json
import contextlib

current_dir = os.path.dirname(os.path.abspath(__file__))
lark_path = os.path.join(current_dir, '../../sheets/formulas.lark')

class FunctionsTests(unittest.TestCase):

    def test_and_function(self):
        wb = sheets.Workbook()
        wb.new_sheet()

        parser = lark.Lark.open(lark_path, start='formula')
        ev = FormulaEvaluator('sheet1', None, func_directory=wb.func_directory)

        # Basic Cases
        tree_1 = parser.parse('=AND(1, 0)')
        ev.ref_info = wb.get_cell_ref_info(tree_1, 'sheet1')
        self.assertEqual(ev.visit(tree_1), False)

        tree_2 = parser.parse('=AND(TRUE, TRUE)')
        ev.ref_info = wb.get_cell_ref_info(tree_2, 'sheet1')
        self.assertEqual(ev.visit(tree_2), True)

        # Wrong Number of Arguments
        tree_3 = parser.parse('=AND()')
        ev.ref_info = wb.get_cell_ref_info(tree_3, 'sheet1')
        self.assertIsInstance(ev.visit(tree_3).val, sheets.CellError)

        # Argument Type Errors
        tree_4 = parser.parse('=AND("text", TRUE)')
        ev.ref_info = wb.get_cell_ref_info(tree_4, 'sheet1')
        self.assertIsInstance(ev.visit(tree_4).val, sheets.CellError)

        tree_5 = parser.parse('=AND(0, 10)')
        ev.ref_info = wb.get_cell_ref_info(tree_5, 'sheet1')
        self.assertEqual(ev.visit(tree_5), False)

        tree_5b = parser.parse('=AND(0, 3.14)')
        ev.ref_info = wb.get_cell_ref_info(tree_5b, 'sheet1')
        self.assertEqual(ev.visit(tree_5b), False)

        # Case Insensitivity
        tree_6 = parser.parse('=and(1, 1)')
        ev.ref_info = wb.get_cell_ref_info(tree_6, 'sheet1')
        self.assertEqual(ev.visit(tree_6), True)

        # Nested AND
        tree_7 = parser.parse('=and(AND(1, 0), 1)')
        ev.ref_info = wb.get_cell_ref_info(tree_7, 'sheet1')
        self.assertEqual(ev.visit(tree_7), False)

        tree_8 = parser.parse('=AND(AND(TRUE, "fire"), FALSE)')
        ev.ref_info = wb.get_cell_ref_info(tree_8, 'sheet1')
        self.assertIsInstance(ev.visit(tree_8).val, sheets.CellError)

        # Extra space
        tree_9 = parser.parse('=AND(      TRUE      ,  FALSE )')
        ev.ref_info = wb.get_cell_ref_info(tree_9, 'sheet1')
        self.assertEqual(ev.visit(tree_9), False)

    def test_or_function(self):
        wb = sheets.Workbook()
        wb.new_sheet()

        parser = lark.Lark.open(lark_path, start='formula')
        ev = FormulaEvaluator('sheet1', None, func_directory=wb.func_directory)

        # Basic Cases
        tree_1 = parser.parse('=OR(1, 0)')
        ev.ref_info = wb.get_cell_ref_info(tree_1, 'sheet1')
        self.assertEqual(ev.visit(tree_1), True)

        tree_2 = parser.parse('=OR(False, FALSE)')
        ev.ref_info = wb.get_cell_ref_info(tree_2, 'sheet1')
        self.assertEqual(ev.visit(tree_2), False)

        # Wrong Number of Arguments
        tree_3 = parser.parse('=OR()')
        ev.ref_info = wb.get_cell_ref_info(tree_3, 'sheet1')
        self.assertIsInstance(ev.visit(tree_3).val, sheets.CellError)

        # Argument Type Errors
        tree_4 = parser.parse('=OR("text", TRUE)')
        ev.ref_info = wb.get_cell_ref_info(tree_4, 'sheet1')
        self.assertIsInstance(ev.visit(tree_4).val, sheets.CellError)

        tree_5 = parser.parse('=OR(0, 10)')
        ev.ref_info = wb.get_cell_ref_info(tree_5, 'sheet1')
        self.assertEqual(ev.visit(tree_5), True)

        tree_5b = parser.parse('=OR(0, 3.14)')
        ev.ref_info = wb.get_cell_ref_info(tree_5b, 'sheet1')
        self.assertEqual(ev.visit(tree_5b), True)

        # Case Insensitivity
        tree_6 = parser.parse('=or(0, 0)')
        ev.ref_info = wb.get_cell_ref_info(tree_6, 'sheet1')
        self.assertEqual(ev.visit(tree_6), False)

        # Nested OR
        tree_7 = parser.parse('=or(OR(1, 0), 0)')
        ev.ref_info = wb.get_cell_ref_info(tree_7, 'sheet1')
        self.assertEqual(ev.visit(tree_7), True)

        tree_8 = parser.parse('=OR(OR(FALSE, "fire"), TRUE)')
        ev.ref_info = wb.get_cell_ref_info(tree_8, 'sheet1')
        self.assertIsInstance(ev.visit(tree_8).val, sheets.CellError)

        # Extra space
        tree_9 = parser.parse('=OR(      TRUE      ,  FALSE )')
        ev.ref_info = wb.get_cell_ref_info(tree_9, 'sheet1')
        self.assertEqual(ev.visit(tree_9), True)

        # Mixed valid/invalid, error propagation
        tree_10 = parser.parse('=OR(TRUE, "text", FALSE)')
        ev.ref_info = wb.get_cell_ref_info(tree_10, 'sheet1')
        self.assertIsInstance(ev.visit(tree_10).val, sheets.CellError)

    def test_not_function(self):
        wb = sheets.Workbook()
        wb.new_sheet()

        parser = lark.Lark.open(lark_path, start='formula')
        ev = FormulaEvaluator('sheet1', None, func_directory=wb.func_directory)

        tree_1 = parser.parse('=NOT(TRUE)')
        ev.ref_info = wb.get_cell_ref_info(tree_1, 'sheet1')
        self.assertEqual(ev.visit(tree_1), False)

        tree_2 = parser.parse('=NOT(FALSE)')
        ev.ref_info = wb.get_cell_ref_info(tree_2, 'sheet1')
        self.assertEqual(ev.visit(tree_2), True)

        # Wrong Number of Arguments
        tree_3 = parser.parse('=NOT()')
        ev.ref_info = wb.get_cell_ref_info(tree_3, 'sheet1')
        self.assertIsInstance(ev.visit(tree_3).val, sheets.CellError)

        tree_4 = parser.parse('=NOT(0)')
        ev.ref_info = wb.get_cell_ref_info(tree_4, 'sheet1')
        self.assertEqual(ev.visit(tree_4), True)

        tree_5 = parser.parse('=NOT(1)')
        ev.ref_info = wb.get_cell_ref_info(tree_5, 'sheet1')
        self.assertEqual(ev.visit(tree_5), False)

        tree_6 = parser.parse('=NOT("hi")')
        ev.ref_info = wb.get_cell_ref_info(tree_6, 'sheet1')
        self.assertIsInstance(ev.visit(tree_6).val, sheets.CellError)

        tree_7 = parser.parse('=NOT(100)')
        ev.ref_info = wb.get_cell_ref_info(tree_7, 'sheet1')
        self.assertEqual(ev.visit(tree_7), False)

        tree_8 = parser.parse('=NOT( true )')
        ev.ref_info = wb.get_cell_ref_info(tree_8, 'sheet1')
        self.assertEqual(ev.visit(tree_8), False)

        tree_9 = parser.parse('=not(false)')
        ev.ref_info = wb.get_cell_ref_info(tree_9, 'sheet1')
        self.assertEqual(ev.visit(tree_9), True)
    
    def test_xor_function(self):
        wb = sheets.Workbook()
        wb.new_sheet()

        parser = lark.Lark.open(lark_path, start='formula')
        tree_1 = parser.parse('=xor(true, false)')
        tree_2 = parser.parse('=xor(False, true)')
        tree_3 = parser.parse('=xor(true, false, false)')
        tree_4 = parser.parse('=xor(true, true, false)')

        ref_info = wb.get_cell_ref_info(tree_1, 'sheet1')

        ev = FormulaEvaluator('sheet1', ref_info, func_directory=wb.func_directory)
        # print(ev.visit(tree_1))
        self.assertEqual(ev.visit(tree_1), True)

        ref_info = wb.get_cell_ref_info(tree_2, 'sheet1')
        self.assertEqual(ev.visit(tree_2), True)

        ref_info = wb.get_cell_ref_info(tree_3, 'sheet1')
        self.assertEqual(ev.visit(tree_3), True)

        ref_info = wb.get_cell_ref_info(tree_4, 'sheet1')
        self.assertEqual(ev.visit(tree_4), False)

        tree_5 = parser.parse('=XOR()')
        ref_info = wb.get_cell_ref_info(tree_5, 'sheet1')
        self.assertIsInstance(ev.visit(tree_5).val, sheets.CellError)
    
    def test_exact_function(self):
        wb = sheets.Workbook()
        wb.new_sheet()

        parser = lark.Lark.open(lark_path, start='formula')
        tree_1 = parser.parse('=EXACT("hi", "hi")')
        ref_info = wb.get_cell_ref_info(tree_1, 'sheet1')

        ev = FormulaEvaluator('sheet1', ref_info, func_directory=wb.func_directory)

        # print(ev.visit(tree_1))
        self.assertEqual(ev.visit(tree_1), True)

        tree_2 = parser.parse('=EXACT("hi", "bye")')
        ref_info = wb.get_cell_ref_info(tree_2, 'sheet1')

        self.assertEqual(ev.visit(tree_2), False)

        # wrong number of arguments
        tree_3 = parser.parse('=EXACT()')
        ref_info = wb.get_cell_ref_info(tree_3, 'sheet1')

        self.assertIsInstance(ev.visit(tree_3).val, sheets.CellError)
    
    def test_if_function(self):
        wb = sheets.Workbook()
        wb.new_sheet()

        parser = lark.Lark.open(lark_path, start='formula')
        tree_1 = parser.parse('=IF(1==1, "yes", "no")')
        ref_info = wb.get_cell_ref_info(tree_1, 'sheet1')

        ev = FormulaEvaluator('sheet1', ref_info, func_directory=wb.func_directory)

        self.assertEqual(ev.visit(tree_1), "yes")

        tree_2 = parser.parse('=IF("blue">"BLUE", "yes", "no")')
        ref_info = wb.get_cell_ref_info(tree_2, 'sheet1')
        self.assertEqual(ev.visit(tree_2), "no")

        # wrong number of arguments
        tree_3 = parser.parse('=IF()')
        ref_info = wb.get_cell_ref_info(tree_3, 'sheet1')

        self.assertIsInstance(ev.visit(tree_3).val, sheets.CellError)
    
    def test_iferror_function(self):
        # TODO
        # wb = sheets.Workbook()
        # wb.new_sheet()

        # parser = lark.Lark.open(lark_path, start='formula')
        # tree_1 = parser.parse('=IFERROR(1/0, 5)')
        # ref_info = wb.get_cell_ref_info(tree_1, 'sheet1')

        # ev = FormulaEvaluator('sheet1', ref_info, func_directory=wb.func_directory)

        # # print(ev.visit(tree_1))
        # self.assertEqual(ev.visit(tree_1), decimal.Decimal('5'))

        # tree_2 = parser.parse('=IFERROR(1+1)')
        # ref_info = wb.get_cell_ref_info(tree_1, 'sheet1')

        # # print(ev.visit(tree_1))
        # self.assertEqual(ev.visit(tree_2), decimal.Decimal('2'))
        pass
    
    def test_choose_function(self):
        # TODO: not finished yet
        wb = sheets.Workbook()
        wb.new_sheet()

        parser = lark.Lark.open(lark_path, start='formula')
        tree_1 = parser.parse('=choose(1, 0, 1, 2, 3)')
        ref_info = wb.get_cell_ref_info(tree_1, 'sheet1')

        ev = FormulaEvaluator('sheet1', ref_info, func_directory=wb.func_directory)

        self.assertEqual(ev.visit(tree_1), decimal.Decimal('0'))

        # TODO

        # wrong number of arguments
        tree_3 = parser.parse('=CHOOSE()')
        ref_info = wb.get_cell_ref_info(tree_3, 'sheet1')

        self.assertIsInstance(ev.visit(tree_3).val, sheets.CellError)
    
    def test_isblank_function(self):
        # TODO
        wb = sheets.Workbook()
        wb.new_sheet()

        parser = lark.Lark.open(lark_path, start='formula')
        tree_1 = parser.parse('=ISBLANK()')

        ref_info = wb.get_cell_ref_info(tree_1, 'sheet1')
        
        ev = FormulaEvaluator('sheet1', ref_info, func_directory=wb.func_directory)
        
        self.assertEqual(ev.visit(tree_1), True)

        # TODO: more 

        # wrong number of arguments
        tree_3 = parser.parse('=ISBLANK(1, 2, 3)')
        ref_info = wb.get_cell_ref_info(tree_3, 'sheet1')

        self.assertIsInstance(ev.visit(tree_3).val, sheets.CellError)

    def test_iserror_function(self):
        pass
    
    def test_version_function(self):
        wb = sheets.Workbook()
        wb.new_sheet()

        parser = lark.Lark.open(lark_path, start='formula')
        tree_1 = parser.parse('=VERSION()')

        ref_info = wb.get_cell_ref_info(tree_1, 'sheet1')
        
        ev = FormulaEvaluator('sheet1', ref_info, func_directory=wb.func_directory)
        
        self.assertEqual(ev.visit(tree_1), "2.0") # TODO change this

        tree_3 = parser.parse('=VERSION(1, 2)')
        ref_info = wb.get_cell_ref_info(tree_3, 'sheet1')

        self.assertIsInstance(ev.visit(tree_3).val, sheets.CellError)

    def test_indirect_function(self):
        # TODO 
        pass

