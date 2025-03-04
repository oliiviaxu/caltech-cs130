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
        self.assertEqual(ev.visit(tree_1).val, False)

        tree_2 = parser.parse('=AND(TRUE, TRUE)')
        ev.ref_info = wb.get_cell_ref_info(tree_2, 'sheet1')
        self.assertEqual(ev.visit(tree_2).val, True)

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
        self.assertEqual(ev.visit(tree_5).val, False)

        tree_5b = parser.parse('=AND(0, 3.14)')
        ev.ref_info = wb.get_cell_ref_info(tree_5b, 'sheet1')
        self.assertEqual(ev.visit(tree_5b).val, False)

        # Case Insensitivity
        tree_6 = parser.parse('=and(1, 1)')
        ev.ref_info = wb.get_cell_ref_info(tree_6, 'sheet1')
        self.assertEqual(ev.visit(tree_6).val, True)

        # Nested AND
        tree_7 = parser.parse('=and(AND(1, 0), 1)')
        ev.ref_info = wb.get_cell_ref_info(tree_7, 'sheet1')
        self.assertEqual(ev.visit(tree_7).val, False)

        tree_8 = parser.parse('=AND(AND(TRUE, "fire"), FALSE)')
        ev.ref_info = wb.get_cell_ref_info(tree_8, 'sheet1')
        self.assertIsInstance(ev.visit(tree_8).val, sheets.CellError)

        # Extra space
        tree_9 = parser.parse('=AND(      TRUE      ,  FALSE )')
        ev.ref_info = wb.get_cell_ref_info(tree_9, 'sheet1')
        self.assertEqual(ev.visit(tree_9).val, False)

        # SET_CELL_CONTENTS tests
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "=AND(1, 1)")
        self.assertEqual(wb.get_cell_value('sheet1', 'A1'), True)
        
        wb.set_cell_contents("sheet1", "A2", "=AND(AND(true, false), 1)")
        self.assertEqual(wb.get_cell_value('sheet1', 'A2'), False)

        # check cell refs
        wb.set_cell_contents("sheet1", "A3", "=AND(A1, A2)")
        self.assertEqual(wb.get_cell_value('sheet1', 'A3'), False)

        # cell error propagation
        wb.set_cell_contents("sheet1", "A4", "=1/0")
        self.assertIsInstance(wb.get_cell_value('sheet1', 'A4'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A4').get_type(), sheets.CellErrorType.DIVIDE_BY_ZERO)

        wb.set_cell_contents("sheet1", "A5", "=AND(A4, True)")
        self.assertIsInstance(wb.get_cell_value('sheet1', 'A5'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A5').get_type(), sheets.CellErrorType.DIVIDE_BY_ZERO)

        wb.set_cell_contents("sheet1", "C1", "=C2")
        wb.set_cell_contents("sheet1", "C2", "=C1")
        self.assertIsInstance(wb.get_cell_value('sheet1', 'C1'), sheets.CellError)
        self.assertIsInstance(wb.get_cell_value('sheet1', 'C1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'C1').get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)
        self.assertEqual(wb.get_cell_value('Sheet1', 'C2').get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)
        wb.set_cell_contents("sheet1", "C3", "=AND(C1, false)")
        self.assertIsInstance(wb.get_cell_value('sheet1', 'C3'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'C3').get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)

        # invalid number of arguments
        wb.set_cell_contents("sheet1", "A6", "=AND()")
        self.assertIsInstance(wb.get_cell_value('sheet1', 'A6'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A6').get_type(), sheets.CellErrorType.TYPE_ERROR)

        wb.set_cell_contents("sheet1", "A7", "=AND(1, 1, 1, 1, 1, 1, 1, 1, 1, 0)")
        self.assertEqual(wb.get_cell_value('sheet1', 'A3'), False)

        wb.set_cell_contents("sheet1", "A8", "=AND(10, 10)")
        self.assertEqual(wb.get_cell_value('sheet1', 'A8'), True)
        
        # TODO: CHECK
        wb.set_cell_contents("sheet1", "A9", "=AND(\"text\", \"text\")")
        self.assertIsInstance(wb.get_cell_value('sheet1', 'A9'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A9').get_type(), sheets.CellErrorType.TYPE_ERROR)

        wb.set_cell_contents("sheet1", "B1", "=AND(TRUE, TRUE)")  # TRUE
        wb.set_cell_contents("sheet1", "B2", "=AND(FALSE, TRUE)")  # FALSE
        wb.set_cell_contents("sheet1", "B3", "=AND(B1, B2)")  # AND(TRUE, FALSE) = FALSE
        self.assertEqual(wb.get_cell_value('sheet1', 'B3'), False)
        
        # Empty Cells
        wb.set_cell_contents("sheet1", "B4", "")
        wb.set_cell_contents("sheet1", "A11", "=AND(B4, TRUE)")
        self.assertEqual(wb.get_cell_value('sheet1', 'A11'), False)

        wb.set_cell_contents("sheet1", "B5", None)
        wb.set_cell_contents("sheet1", "A11", "=AND(B5, TRUE)")
        self.assertEqual(wb.get_cell_value('sheet1', 'A11'), False)


    def test_or_function(self):
        wb = sheets.Workbook()
        wb.new_sheet()

        parser = lark.Lark.open(lark_path, start='formula')
        ev = FormulaEvaluator('sheet1', None, func_directory=wb.func_directory)

        # Basic Cases
        tree_1 = parser.parse('=OR(1, 0)')
        ev.ref_info = wb.get_cell_ref_info(tree_1, 'sheet1')
        self.assertEqual(ev.visit(tree_1).val, True)

        tree_2 = parser.parse('=OR(False, FALSE)')
        ev.ref_info = wb.get_cell_ref_info(tree_2, 'sheet1')
        self.assertEqual(ev.visit(tree_2).val, False)

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
        self.assertEqual(ev.visit(tree_5).val, True)

        tree_5b = parser.parse('=OR(0, 3.14)')
        ev.ref_info = wb.get_cell_ref_info(tree_5b, 'sheet1')
        self.assertEqual(ev.visit(tree_5b).val, True)

        # Case Insensitivity
        tree_6 = parser.parse('=or(0, 0)')
        ev.ref_info = wb.get_cell_ref_info(tree_6, 'sheet1')
        self.assertEqual(ev.visit(tree_6).val, False)

        # Nested OR
        tree_7 = parser.parse('=or(OR(1, 0), 0)')
        ev.ref_info = wb.get_cell_ref_info(tree_7, 'sheet1')
        self.assertEqual(ev.visit(tree_7).val, True)

        tree_8 = parser.parse('=OR(OR(FALSE, "fire"), TRUE)')
        ev.ref_info = wb.get_cell_ref_info(tree_8, 'sheet1')
        self.assertIsInstance(ev.visit(tree_8).val, sheets.CellError)

        # Extra space
        tree_9 = parser.parse('=OR(      TRUE      ,  FALSE )')
        ev.ref_info = wb.get_cell_ref_info(tree_9, 'sheet1')
        self.assertEqual(ev.visit(tree_9).val, True)

        # Mixed valid/invalid, error propagation
        tree_10 = parser.parse('=OR(TRUE, "text", FALSE)')
        ev.ref_info = wb.get_cell_ref_info(tree_10, 'sheet1')
        self.assertIsInstance(ev.visit(tree_10).val, sheets.CellError)

        # TODO: SET_CELL_CONTENTS tests
        wb.set_cell_contents("sheet1", "A1", "=OR(TRUE, FALSE)")
        self.assertEqual(wb.get_cell_value('sheet1', 'A1'), True)

        wb.set_cell_contents("sheet1", "A2", "=OR(0, 0, 0, 0, 0)")
        self.assertEqual(wb.get_cell_value('sheet1', 'A2'), False)

        wb.set_cell_contents("sheet1", "A3", "=OR(1, 0)")
        self.assertEqual(wb.get_cell_value('sheet1', 'A3'), True)

        wb.set_cell_contents("sheet1", "A4", "=OR(0, 10)")
        self.assertEqual(wb.get_cell_value('sheet1', 'A4'), True)

        # Case insensitivity
        wb.set_cell_contents("sheet1", "A5", "=or(FALSE, TRUE)")
        self.assertEqual(wb.get_cell_value('sheet1', 'A5'), True)

        # Nested ORs
        wb.set_cell_contents("sheet1", "A6", "=OR(OR(1, 0), 0)")
        self.assertEqual(wb.get_cell_value('sheet1', 'A6'), True)

        # cell refs
        wb.set_cell_contents("sheet1", "A7", "=OR(0, 0)") # F
        wb.set_cell_contents("sheet1", "A8", "=OR(1, 0)") # T
        wb.set_cell_contents("sheet1", "A9", "=OR(1, 1)") # T
        wb.set_cell_contents("sheet1", "A10", "=OR(A7, A8, A9)") # T
        self.assertEqual(wb.get_cell_value('sheet1', 'A10'), True)

        # invalid num args
        wb.set_cell_contents("sheet1", "A11", "=OR()")
        self.assertIsInstance(wb.get_cell_value('sheet1', 'A11'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A11').get_type(), sheets.CellErrorType.TYPE_ERROR)

        # error propagation
        wb.set_cell_contents("sheet1", "B1", "=OR(TRUE, \"text\")") 
        self.assertIsInstance(wb.get_cell_value('sheet1', 'B1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'B1').get_type(), sheets.CellErrorType.TYPE_ERROR)

        wb.set_cell_contents("sheet1", "C1", "=C2")
        wb.set_cell_contents("sheet1", "C2", "=C1")
        self.assertIsInstance(wb.get_cell_value('sheet1', 'C1'), sheets.CellError)
        self.assertIsInstance(wb.get_cell_value('sheet1', 'C1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'C1').get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)
        self.assertEqual(wb.get_cell_value('Sheet1', 'C2').get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)
        wb.set_cell_contents("sheet1", "C3", "=OR(C1, false)")
        self.assertIsInstance(wb.get_cell_value('sheet1', 'C3'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'C3').get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)

        # Empty Cells
        wb.set_cell_contents("sheet1", "B4", "")
        wb.set_cell_contents("sheet1", "A12", "=Or(B4, TRUE)")
        self.assertEqual(wb.get_cell_value('sheet1', 'A12'), True)

        wb.set_cell_contents("sheet1", "B5", None)
        wb.set_cell_contents("sheet1", "A12", "=or(B5, TRUE)")
        self.assertEqual(wb.get_cell_value('sheet1', 'A12'), True)

    def test_not_function(self):
        wb = sheets.Workbook()
        wb.new_sheet()

        parser = lark.Lark.open(lark_path, start='formula')
        ev = FormulaEvaluator('sheet1', None, func_directory=wb.func_directory)

        tree_1 = parser.parse('=NOT(TRUE)')
        ev.ref_info = wb.get_cell_ref_info(tree_1, 'sheet1')
        self.assertEqual(ev.visit(tree_1).val, False)

        tree_2 = parser.parse('=NOT(FALSE)')
        ev.ref_info = wb.get_cell_ref_info(tree_2, 'sheet1')
        self.assertEqual(ev.visit(tree_2).val, True)

        # Wrong Number of Arguments
        tree_3 = parser.parse('=NOT()')
        ev.ref_info = wb.get_cell_ref_info(tree_3, 'sheet1')
        self.assertIsInstance(ev.visit(tree_3).val, sheets.CellError)

        tree_4 = parser.parse('=NOT(0)')
        ev.ref_info = wb.get_cell_ref_info(tree_4, 'sheet1')
        self.assertEqual(ev.visit(tree_4).val, True)

        tree_5 = parser.parse('=NOT(1)')
        ev.ref_info = wb.get_cell_ref_info(tree_5, 'sheet1')
        self.assertEqual(ev.visit(tree_5).val, False)

        tree_6 = parser.parse('=NOT("hi")')
        ev.ref_info = wb.get_cell_ref_info(tree_6, 'sheet1')
        self.assertIsInstance(ev.visit(tree_6).val, sheets.CellError)

        tree_7 = parser.parse('=NOT(100)')
        ev.ref_info = wb.get_cell_ref_info(tree_7, 'sheet1')
        self.assertEqual(ev.visit(tree_7).val, False)

        tree_8 = parser.parse('=NOT( true )')
        ev.ref_info = wb.get_cell_ref_info(tree_8, 'sheet1')
        self.assertEqual(ev.visit(tree_8).val, False)

        tree_9 = parser.parse('=not(false)')
        ev.ref_info = wb.get_cell_ref_info(tree_9, 'sheet1')
        self.assertEqual(ev.visit(tree_9).val, True)

        # TODO: set_cell_contents tests
        wb.new_sheet()
        wb.set_cell_contents("sheet1", "A1", "=NOT(1)")
        self.assertEqual(wb.get_cell_value('sheet1', 'A1'), False)
        
        wb.set_cell_contents("sheet1", "A2", "=NOT(NOT(true))")
        self.assertEqual(wb.get_cell_value('sheet1', 'A2'), True)

        # check cell refs
        wb.set_cell_contents("sheet1", "A3", "=NOT(A1)")
        self.assertEqual(wb.get_cell_value('sheet1', 'A3'), True)

        # cell error propagation
        wb.set_cell_contents("sheet1", "A4", "=1/0")
        self.assertIsInstance(wb.get_cell_value('sheet1', 'A4'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A4').get_type(), sheets.CellErrorType.DIVIDE_BY_ZERO)

        wb.set_cell_contents("sheet1", "A5", "=NOT(A4)")
        self.assertIsInstance(wb.get_cell_value('sheet1', 'A5'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A5').get_type(), sheets.CellErrorType.DIVIDE_BY_ZERO)

        wb.set_cell_contents("sheet1", "C1", "=C2")
        wb.set_cell_contents("sheet1", "C2", "=C1")
        self.assertIsInstance(wb.get_cell_value('sheet1', 'C1'), sheets.CellError)
        self.assertIsInstance(wb.get_cell_value('sheet1', 'C1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'C1').get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)
        self.assertEqual(wb.get_cell_value('Sheet1', 'C2').get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)
        wb.set_cell_contents("sheet1", "C3", "=NOT(C1)")
        self.assertIsInstance(wb.get_cell_value('sheet1', 'C3'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'C3').get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)

        # invalid number of arguments
        wb.set_cell_contents("sheet1", "A6", "=NOT()")
        self.assertIsInstance(wb.get_cell_value('sheet1', 'A6'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A6').get_type(), sheets.CellErrorType.TYPE_ERROR)

        wb.set_cell_contents("sheet1", "A8", "=NOT(10)")
        self.assertEqual(wb.get_cell_value('sheet1', 'A8'), False)
        
        # TODO: CHECK
        wb.set_cell_contents("sheet1", "A9", "=NOT(\"text\")")
        self.assertIsInstance(wb.get_cell_value('sheet1', 'A9'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A9').get_type(), sheets.CellErrorType.TYPE_ERROR)

        wb.set_cell_contents("sheet1", "B1", "=NOT(faLSE)")  # TRUE
        wb.set_cell_contents("sheet1", "B3", "=NOT(B1)")
        self.assertEqual(wb.get_cell_value('sheet1', 'B3'), False)

        wb.set_cell_contents("sheet1", "B4", "=NOT(1000000000000)")
        self.assertEqual(wb.get_cell_value('sheet1', 'B4'), False)

        # Empty Cells
        wb.set_cell_contents("sheet1", "B5", "")
        wb.set_cell_contents("sheet1", "A12", "=NOT(B5)")
        self.assertEqual(wb.get_cell_value('sheet1', 'A12'), True)

        wb.set_cell_contents("sheet1", "B6", None)
        wb.set_cell_contents("sheet1", "A12", "=not(B6)")
        self.assertEqual(wb.get_cell_value('sheet1', 'A12'), True)
    
    def test_xor_function(self):
        wb = sheets.Workbook()
        wb.new_sheet()

        parser = lark.Lark.open(lark_path, start='formula')
        ev = FormulaEvaluator('sheet1', None, func_directory=wb.func_directory)

        tree_1 = parser.parse('=XOR(TRUE, TRUE)')
        ev.ref_info = wb.get_cell_ref_info(tree_1, 'sheet1')
        self.assertEqual(ev.visit(tree_1).val, False)

        tree_2 = parser.parse('=XOR(TRUE, FALSE)')
        ev.ref_info = wb.get_cell_ref_info(tree_2, 'sheet1')
        self.assertEqual(ev.visit(tree_2).val, True)

        tree_3 = parser.parse('=XOR(1, 0)')
        ev.ref_info = wb.get_cell_ref_info(tree_3, 'sheet1')
        self.assertEqual(ev.visit(tree_3).val, True)

        tree_4 = parser.parse('=XOR(0, 0)')
        ev.ref_info = wb.get_cell_ref_info(tree_4, 'sheet1')
        self.assertEqual(ev.visit(tree_4).val, False)

        tree_5 = parser.parse('=XOR()')
        ev.ref_info = wb.get_cell_ref_info(tree_5, 'sheet1')
        self.assertIsInstance(ev.visit(tree_5).val, sheets.CellError)

        tree_6 = parser.parse('=XOR("text", TRUE)')
        ev.ref_info = wb.get_cell_ref_info(tree_6, 'sheet1')
        self.assertIsInstance(ev.visit(tree_6).val, sheets.CellError)

        tree_7 = parser.parse('=XOR(3.14, FALSE)')
        ev.ref_info = wb.get_cell_ref_info(tree_7, 'sheet1')
        self.assertEqual(ev.visit(tree_7).val, True)

        tree_8 = parser.parse('=XOR(1, 0, 1, 0, 1)')
        ev.ref_info = wb.get_cell_ref_info(tree_8, 'sheet1')
        self.assertEqual(ev.visit(tree_8).val, True)

        tree_9 = parser.parse('=XOR(true        , FALSE)')
        ev.ref_info = wb.get_cell_ref_info(tree_9, 'sheet1')
        self.assertEqual(ev.visit(tree_9).val, True)

        tree_10 = parser.parse('=xor(XOR(TRUE, FALSE), TRUE)')
        ev.ref_info = wb.get_cell_ref_info(tree_10, 'sheet1')
        self.assertEqual(ev.visit(tree_10).val, False)

        # TODO: set_cell_contents tests
        wb.set_cell_contents("sheet1", "A1", "=XOR(TRUE, TRUE)")
        self.assertEqual(wb.get_cell_value('sheet1', 'A1'), False)

        wb.set_cell_contents("sheet1", "A2", "=XOR(TRUE, FALSE)")
        self.assertEqual(wb.get_cell_value('sheet1', 'A2'), True)

        wb.set_cell_contents("sheet1", "A3", "=XOR(1, 0)")
        self.assertEqual(wb.get_cell_value('sheet1', 'A3'), True)

        wb.set_cell_contents("sheet1", "A4", "=XOR(0, 0)")
        self.assertEqual(wb.get_cell_value('sheet1', 'A4'), False)

        # Error cases
        wb.set_cell_contents("sheet1", "A5", "=XOR()")
        self.assertIsInstance(wb.get_cell_value('sheet1', 'A5'), sheets.CellError)

        wb.set_cell_contents("sheet1", "A6", "=XOR(\"text\", TRUE)")
        self.assertIsInstance(wb.get_cell_value('sheet1', 'A6'), sheets.CellError)

        # Numeric coercion
        wb.set_cell_contents("sheet1", "A7", "=XOR(3.14, FALSE)")
        self.assertEqual(wb.get_cell_value('sheet1', 'A7'), True)

        # Multiple arguments
        wb.set_cell_contents("sheet1", "A8", "=XOR(1, 0, 1, 0, 1)")
        self.assertEqual(wb.get_cell_value('sheet1', 'A8'), True)

        # Extra space and case insensitivity
        wb.set_cell_contents("sheet1", "A9", "=XOR(true        , FALSE)")
        self.assertEqual(wb.get_cell_value('sheet1', 'A9'), True)

        # Nested XOR
        wb.set_cell_contents("sheet1", "A10", "=xor(XOR(TRUE, FALSE), TRUE)")
        self.assertEqual(wb.get_cell_value('sheet1', 'A10'), False)

        # Cell References
        wb.set_cell_contents("sheet1", "B1", "TRUE")
        wb.set_cell_contents("sheet1", "B2", "FALSE")
        wb.set_cell_contents("sheet1", "A11", "=XOR(B1, B2)")
        self.assertEqual(wb.get_cell_value('sheet1', 'A11'), True)

        # error propagation
        wb.set_cell_contents("sheet1", "C1", "=C2")
        wb.set_cell_contents("sheet1", "C2", "=C1")
        self.assertIsInstance(wb.get_cell_value('sheet1', 'C1'), sheets.CellError)
        self.assertIsInstance(wb.get_cell_value('sheet1', 'C1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'C1').get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)
        self.assertEqual(wb.get_cell_value('Sheet1', 'C2').get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)
        wb.set_cell_contents("sheet1", "C3", "=XOR(C1, C2)")
        self.assertIsInstance(wb.get_cell_value('sheet1', 'C3'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'C3').get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)

        # Null/Empty Values
        wb.set_cell_contents("sheet1", "B3", "")
        wb.set_cell_contents("sheet1", "A12", "=XOR(B3, TRUE)")
        self.assertEqual(wb.get_cell_value('sheet1', 'A12'), True)

        wb.set_cell_contents("sheet1", "B4", None)
        wb.set_cell_contents("sheet1", "A13", "=XOR(B4, TRUE)")
        self.assertEqual(wb.get_cell_value('sheet1', 'A13'), True)

        wb.set_cell_contents("sheet1", "B7", '1000000000000')
        wb.set_cell_contents("sheet1", "A17", "=XOR(B7, FALSE)")
        self.assertEqual(wb.get_cell_value('sheet1', 'A17'), True)

        # Incorrect Argument Types
        wb.set_cell_contents("sheet1", "E1", "test")
        wb.set_cell_contents("sheet1", "A19", "=XOR(E1, TRUE)")
        self.assertIsInstance(wb.get_cell_value('sheet1', 'A19'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('sheet1', 'A19').get_type(), sheets.CellErrorType.TYPE_ERROR)

    def test_exact_function(self):
        wb = sheets.Workbook()
        wb.new_sheet()

        parser = lark.Lark.open(lark_path, start='formula')
        ev = FormulaEvaluator('sheet1', None, func_directory=wb.func_directory)

        tree_1 = parser.parse('=EXACT("hi", "hi")')
        ev.ref_info = wb.get_cell_ref_info(tree_1, 'sheet1')
        # print(ev.visit(tree_1))
        self.assertEqual(ev.visit(tree_1).val, True)

        tree_2 = parser.parse('=EXACT("hi", "bye")')
        ev.ref_info = wb.get_cell_ref_info(tree_2, 'sheet1')
        self.assertEqual(ev.visit(tree_2).val, False)

        # wrong number of arguments
        tree_3 = parser.parse('=EXACT()')
        ev.ref_info = wb.get_cell_ref_info(tree_3, 'sheet1')
        self.assertIsInstance(ev.visit(tree_3).val, sheets.CellError)

        tree_4 = parser.parse('=EXACT("", "")')
        ev.ref_info = wb.get_cell_ref_info(tree_4, 'sheet1')
        self.assertEqual(ev.visit(tree_4).val, True)

        tree_5 = parser.parse('=EXACT(123, 456)')
        ev.ref_info = wb.get_cell_ref_info(tree_5, 'sheet1')
        self.assertEqual(ev.visit(tree_5).val, False)

        tree_6 = parser.parse('=EXACT(EXACT("hi", "hi"), "Hi")')
        ev.ref_info = wb.get_cell_ref_info(tree_6, 'sheet1')
        self.assertEqual(ev.visit(tree_6).val, False)

        # TODO: set_cell_contents tests
        wb.set_cell_contents("sheet1", "A1", '=EXACT("hello", "hello")')
        self.assertEqual(wb.get_cell_value('sheet1', 'A1'), True)

        wb.set_cell_contents("sheet1", "A2", '=EXACT("hello", "Hello")')
        self.assertEqual(wb.get_cell_value('sheet1', 'A2'), False)

        # string conversion
        wb.set_cell_contents("sheet1", "A3", '=EXACT("123", 123)')
        self.assertEqual(wb.get_cell_value('sheet1', 'A3'), True)

        wb.set_cell_contents("sheet1", "E1", '=EXACT("1.5", 1.5)')
        self.assertEqual(wb.get_cell_value('sheet1', 'A3'), True)

        wb.set_cell_contents("sheet1", "A5", '=EXACT("  test", "test  ")')
        self.assertEqual(wb.get_cell_value('sheet1', 'A5'), False)

        wb.set_cell_contents("sheet1", "B1", "test")
        wb.set_cell_contents("sheet1", "B2", "test")
        wb.set_cell_contents("sheet1", "A6", "=EXACT(B1,B2)")
        self.assertEqual(wb.get_cell_value('sheet1', 'A6'), True)

        # error propagation
        wb.set_cell_contents("sheet1", "A7", "=EXACT()")
        self.assertIsInstance(wb.get_cell_value('sheet1', 'A7'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A7').get_type(), sheets.CellErrorType.TYPE_ERROR)

        wb.set_cell_contents("sheet1", "C1", "=1/0")
        wb.set_cell_contents("sheet1", "A10", "=EXACT(C1, \"test\")")
        self.assertIsInstance(wb.get_cell_value('sheet1', 'A10'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A10').get_type(), sheets.CellErrorType.DIVIDE_BY_ZERO)

        wb.set_cell_contents("sheet1", "C2", "=C3")
        wb.set_cell_contents("sheet1", "C3", "=C2")
        wb.set_cell_contents("sheet1", "A11", "=EXACT(C2, C3)")
        self.assertIsInstance(wb.get_cell_value('sheet1', 'A11'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A11').get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)

        # empty values
        # TODO: check
        wb.set_cell_contents("sheet1", "D1", "")
        wb.set_cell_contents("sheet1", "D2", None)

        wb.set_cell_contents("sheet1", "A12", "=EXACT(D1, D2)") 
        self.assertEqual(wb.get_cell_value('sheet1', 'A12'), True)

    def test_if_function(self):
        wb = sheets.Workbook()
        wb.new_sheet()

        parser = lark.Lark.open(lark_path, start='formula')
        ev = FormulaEvaluator('sheet1', None, func_directory=wb.func_directory)

        tree_1 = parser.parse('=IF(1==1, "yes", "no")')
        ev.ref_info = wb.get_cell_ref_info(tree_1, 'sheet1')
        self.assertEqual(ev.visit(tree_1).val, "yes")

        tree_2 = parser.parse('=IF("blue">"BLUE", "yes", "no")')
        ev.ref_info = wb.get_cell_ref_info(tree_2, 'sheet1')
        self.assertEqual(ev.visit(tree_2).val, "no")

        # wrong number of arguments
        tree_3 = parser.parse('=IF()')
        ev.ref_info = wb.get_cell_ref_info(tree_3, 'sheet1')
        self.assertIsInstance(ev.visit(tree_3).val, sheets.CellError)

        # error propagation
        tree_4 = parser.parse('=IF(1/0, 1, 2)')
        ev.ref_info = wb.get_cell_ref_info(tree_4, 'sheet1')
        self.assertIsInstance(ev.visit(tree_4).val, sheets.CellError)

        # TODO: set_cell_contents tests
        wb.set_cell_contents('sheet1', 'A1', '10')
        wb.set_cell_contents('sheet1', 'A2', '5')
        wb.set_cell_contents('sheet1', 'A3', '=IF(A1 > B1, "A1 greater", "B1 greater")')
        self.assertEqual(wb.get_cell_value('sheet1', 'A3'), "A1 greater")

        wb.set_cell_contents('sheet1', 'A4', '=IF(1>0, "true", "false")')
        self.assertEqual(wb.get_cell_value('sheet1', 'A4'), "true")

        # cell refs
        wb.set_cell_contents('sheet1', 'A5', '=10')
        wb.set_cell_contents('sheet1', 'A6', '=IF(A1==A5, "A1 and A5 are equal", "A1 and A5 are not equal")')

        self.assertEqual(wb.get_cell_value('sheet1', 'A6'), "A1 and A5 are equal")

        # error propagation
        wb.set_cell_contents('sheet1', 'A7', '=IF(TRUE, 1/0, 1)')
        self.assertIsInstance(wb.get_cell_value('sheet1', 'A7'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A7').get_type(), sheets.CellErrorType.DIVIDE_BY_ZERO)

        wb.set_cell_contents('sheet1', 'A8', '=IF(FALSE, 1, 1/0)')
        self.assertIsInstance(wb.get_cell_value('sheet1', 'A8'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A8').get_type(), sheets.CellErrorType.DIVIDE_BY_ZERO)

        wb.set_cell_contents('sheet1', 'B2', '=1/0')
        wb.set_cell_contents('sheet1', 'A9', '=IF(A1>B2, "yes", "no")')
        self.assertIsInstance(wb.get_cell_value('sheet1', 'A9'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A9').get_type(), sheets.CellErrorType.DIVIDE_BY_ZERO)

        # nested
        wb.set_cell_contents('sheet1', 'A10', '=IF(1=1, IF(1/0, 1, 2), 3)')
        self.assertIsInstance(wb.get_cell_value('sheet1', 'A10'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A10').get_type(), sheets.CellErrorType.DIVIDE_BY_ZERO)

        wb.set_cell_contents('sheet1', 'A11', '=IF(A1>5, IF(A1>10, "high", "medium"), "low")') # recall A1 = 10
        self.assertEqual(wb.get_cell_value('sheet1', 'A11'), "medium")
        wb.set_cell_contents('sheet1', 'A1', '15')
        self.assertEqual(wb.get_cell_value('sheet1', 'A11'), "high")

        # empty values
        wb.set_cell_contents('sheet1', 'A14', '=IF(""="", "empty", "not empty")')
        self.assertEqual(wb.get_cell_value('sheet1', 'A14'), "empty")

    
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
        # self.assertEqual(ev.visit(tree_2).val, decimal.Decimal('2'))
        pass
    
    def test_choose_function(self):
        # TODO: not finished yet
        wb = sheets.Workbook()
        wb.new_sheet()

        parser = lark.Lark.open(lark_path, start='formula')
        ev = FormulaEvaluator('sheet1', None, func_directory=wb.func_directory)

        tree_1 = parser.parse('=choose(1, 0, 1, 2, 3)')
        ev.ref_info = wb.get_cell_ref_info(tree_1, 'sheet1')
        self.assertEqual(ev.visit(tree_1).val, decimal.Decimal('0'))

        tree_2 = parser.parse('=CHOOSE(2, "a", "b", "c")')
        ev.ref_info = wb.get_cell_ref_info(tree_2, 'sheet1')
        self.assertEqual(ev.visit(tree_2).val, "b")

        # wrong number of arguments
        tree_3 = parser.parse('=CHOOSE()')
        ev.ref_info = wb.get_cell_ref_info(tree_3, 'sheet1')
        self.assertIsInstance(ev.visit(tree_3).val, sheets.CellError)

        # out of bounds
        tree_4 = parser.parse('=CHOOSE(0, "a", "b", "c")')
        ev.ref_info = wb.get_cell_ref_info(tree_4, 'sheet1')
        self.assertIsInstance(ev.visit(tree_3).val, sheets.CellError)

        tree_5 = parser.parse('=CHOOSE(4, "a", "b", "c")')
        ev.ref_info = wb.get_cell_ref_info(tree_5, 'sheet1')
        self.assertIsInstance(ev.visit(tree_5).val, sheets.CellError)

        tree_6 = parser.parse('=CHOOSE(-1, "a", "b", "c")')
        ev.ref_info = wb.get_cell_ref_info(tree_6, 'sheet1')
        self.assertIsInstance(ev.visit(tree_6).val, sheets.CellError)

        tree_7 = parser.parse('=CHOOSE(1.5, "a", "b", "c")')
        ev.ref_info = wb.get_cell_ref_info(tree_7, 'sheet1')
        self.assertIsInstance(ev.visit(tree_7).val, sheets.CellError)

        tree_8 = parser.parse('=choose("text", "a", "b", "c")')
        ev.ref_info = wb.get_cell_ref_info(tree_8, 'sheet1')
        self.assertIsInstance(ev.visit(tree_8).val, sheets.CellError)

        # TODO: not sure if this is the case
        tree_9 = parser.parse('=CHOOSE(1, 1/0, 2, 3)')
        ev.ref_info = wb.get_cell_ref_info(tree_9, 'sheet1')
        # print(ev.visit(tree_9).val)
        self.assertIsInstance(ev.visit(tree_8).val, sheets.CellError)

        # set cell contents tests
    
    def test_isblank_function(self):
        wb = sheets.Workbook()
        wb.new_sheet()

        parser = lark.Lark.open(lark_path, start='formula')
        ev = FormulaEvaluator('sheet1', None, func_directory=wb.func_directory)

        tree_1 = parser.parse('=ISBLANK()')
        ev.ref_info = wb.get_cell_ref_info(tree_1, 'sheet1')        
        self.assertIsInstance(ev.visit(tree_1).val, sheets.CellError)

        tree_2 = parser.parse('=ISBLANK("A")')
        ev.ref_info = wb.get_cell_ref_info(tree_2, 'sheet1')        
        self.assertEqual(ev.visit(tree_2).val, False)

        # wrong number of arguments
        tree_3 = parser.parse('=ISBLANK(1, 2, 3)')
        ev.ref_info = wb.get_cell_ref_info(tree_3, 'sheet1')
        self.assertIsInstance(ev.visit(tree_3).val, sheets.CellError)

        # TODO: set cell contents tests

    def test_iserror_function(self):
        # TODO

        pass
    
    def test_version_function(self):
        # TODO
        wb = sheets.Workbook()
        wb.new_sheet()

        parser = lark.Lark.open(lark_path, start='formula')
        tree_1 = parser.parse('=VERSION()')

        ref_info = wb.get_cell_ref_info(tree_1, 'sheet1')
        
        ev = FormulaEvaluator('sheet1', ref_info, func_directory=wb.func_directory)
        
        self.assertEqual(ev.visit(tree_1), "2.0")

        tree_3 = parser.parse('=VERSION(1, 2)')
        ref_info = wb.get_cell_ref_info(tree_3, 'sheet1')

        self.assertIsInstance(ev.visit(tree_3).val, sheets.CellError)

    def test_indirect_function(self):
        # TODO 
        pass

