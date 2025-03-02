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
        tree_1 = parser.parse('=AND(1, 0)')
        ref_info = wb.get_cell_ref_info(tree_1, 'sheet1')

        ev = FormulaEvaluator('sheet1', ref_info, func_directory=wb.func_directory)

        # print(ev.visit(tree_1))
        self.assertEqual(ev.visit(tree_1), False)

        tree_2 = parser.parse('=AND(TrUE, TRUE)')
        ref_info = wb.get_cell_ref_info(tree_2, 'sheet1')

        self.assertEqual(ev.visit(tree_2), True)

    def test_or_function(self):
        wb = sheets.Workbook()
        wb.new_sheet()

        parser = lark.Lark.open(lark_path, start='formula')
        tree_1 = parser.parse('=OR(1, 0)')
        ref_info = wb.get_cell_ref_info(tree_1, 'sheet1')

        ev = FormulaEvaluator('sheet1', ref_info, func_directory=wb.func_directory)

        # print(ev.visit(tree_1))
        self.assertEqual(ev.visit(tree_1), True)

        tree_2 = parser.parse('=OR(False, FALSE)')
        ref_info = wb.get_cell_ref_info(tree_2, 'sheet1')

        self.assertEqual(ev.visit(tree_2), False)
    
    def test_not_function(self):
        wb = sheets.Workbook()
        wb.new_sheet()

        parser = lark.Lark.open(lark_path, start='formula')
        tree_1 = parser.parse('=NOT(TRUE)')
        ref_info = wb.get_cell_ref_info(tree_1, 'sheet1')

        ev = FormulaEvaluator('sheet1', ref_info, func_directory=wb.func_directory)

        # print(ev.visit(tree_1))
        self.assertEqual(ev.visit(tree_1), False)

        tree_2 = parser.parse('=NOT(0)')
        ref_info = wb.get_cell_ref_info(tree_2, 'sheet1')

        self.assertEqual(ev.visit(tree_2), True)
    
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
    
    def test_if_function(self):
        wb = sheets.Workbook()
        wb.new_sheet()

        parser = lark.Lark.open(lark_path, start='formula')
        tree_1 = parser.parse('=IF(1==1, "yes", "no")')
        ref_info = wb.get_cell_ref_info(tree_1, 'sheet1')

        ev = FormulaEvaluator('sheet1', ref_info, func_directory=wb.func_directory)

        print(ev.visit(tree_1))

