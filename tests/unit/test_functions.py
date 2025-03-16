# python -m unittest tests.unit.test_functions.FunctionsTests

import unittest
import coverage
import sheets
import os
import lark
import sheets.Cell
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
        ev = FormulaEvaluator('sheet1', wb, func_directory=wb.func_directory)

        # Basic Cases
        tree_1 = parser.parse('=AND(1, 0)')
        self.assertEqual(ev.visit(tree_1).val, False)

        tree_2 = parser.parse('=AND(TRUE, TRUE)')
        self.assertEqual(ev.visit(tree_2).val, True)

        # Wrong Number of Arguments
        tree_3 = parser.parse('=AND()')
        self.assertIsInstance(ev.visit(tree_3).val, sheets.CellError)

        # Argument Type Errors
        tree_4 = parser.parse('=AND("text", TRUE)')
        self.assertIsInstance(ev.visit(tree_4).val, sheets.CellError)

        tree_5 = parser.parse('=AND(0, 10)')
        self.assertEqual(ev.visit(tree_5).val, False)

        tree_5b = parser.parse('=AND(0, 3.14)')
        self.assertEqual(ev.visit(tree_5b).val, False)

        # Case Insensitivity
        tree_6 = parser.parse('=and(1, 1)')
        self.assertEqual(ev.visit(tree_6).val, True)

        # Nested AND
        tree_7 = parser.parse('=and(AND(1, 0), 1)')
        self.assertEqual(ev.visit(tree_7).val, False)

        tree_8 = parser.parse('=AND(AND(TRUE, "fire"), FALSE)')
        self.assertIsInstance(ev.visit(tree_8).val, sheets.CellError)

        # Extra space
        tree_9 = parser.parse('=AND(      TRUE      ,  FALSE )')
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
        ev = FormulaEvaluator('sheet1', wb, func_directory=wb.func_directory)

        # Basic Cases
        tree_1 = parser.parse('=OR(1, 0)')
        self.assertEqual(ev.visit(tree_1).val, True)

        tree_2 = parser.parse('=OR(False, FALSE)')
        self.assertEqual(ev.visit(tree_2).val, False)

        # Wrong Number of Arguments
        tree_3 = parser.parse('=OR()')
        self.assertIsInstance(ev.visit(tree_3).val, sheets.CellError)

        # Argument Type Errors
        tree_4 = parser.parse('=OR("text", TRUE)')
        self.assertIsInstance(ev.visit(tree_4).val, sheets.CellError)

        tree_5 = parser.parse('=OR(0, 10)')
        self.assertEqual(ev.visit(tree_5).val, True)

        tree_5b = parser.parse('=OR(0, 3.14)')
        self.assertEqual(ev.visit(tree_5b).val, True)

        # Case Insensitivity
        tree_6 = parser.parse('=or(0, 0)')
        self.assertEqual(ev.visit(tree_6).val, False)

        # Nested OR
        tree_7 = parser.parse('=or(OR(1, 0), 0)')
        self.assertEqual(ev.visit(tree_7).val, True)

        tree_8 = parser.parse('=OR(OR(FALSE, "fire"), TRUE)')
        self.assertIsInstance(ev.visit(tree_8).val, sheets.CellError)

        # Extra space
        tree_9 = parser.parse('=OR(      TRUE      ,  FALSE )')
        self.assertEqual(ev.visit(tree_9).val, True)

        # Mixed valid/invalid, error propagation
        tree_10 = parser.parse('=OR(TRUE, "text", FALSE)')
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
        ev = FormulaEvaluator('sheet1', wb, func_directory=wb.func_directory)

        tree_1 = parser.parse('=NOT(TRUE)')
        self.assertEqual(ev.visit(tree_1).val, False)

        tree_2 = parser.parse('=NOT(FALSE)')
        self.assertEqual(ev.visit(tree_2).val, True)

        # Wrong Number of Arguments
        tree_3 = parser.parse('=NOT()')
        self.assertIsInstance(ev.visit(tree_3).val, sheets.CellError)

        tree_4 = parser.parse('=NOT(0)')
        self.assertEqual(ev.visit(tree_4).val, True)

        tree_5 = parser.parse('=NOT(1)')
        self.assertEqual(ev.visit(tree_5).val, False)

        tree_6 = parser.parse('=NOT("hi")')
        self.assertIsInstance(ev.visit(tree_6).val, sheets.CellError)

        tree_7 = parser.parse('=NOT(100)')
        self.assertEqual(ev.visit(tree_7).val, False)

        tree_8 = parser.parse('=NOT( true )')
        self.assertEqual(ev.visit(tree_8).val, False)

        tree_9 = parser.parse('=not(false)')
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
        ev = FormulaEvaluator('sheet1', wb, func_directory=wb.func_directory)

        tree_1 = parser.parse('=XOR(TRUE, TRUE)')
        self.assertEqual(ev.visit(tree_1).val, False)

        tree_2 = parser.parse('=XOR(TRUE, FALSE)')
        self.assertEqual(ev.visit(tree_2).val, True)

        tree_3 = parser.parse('=XOR(1, 0)')
        self.assertEqual(ev.visit(tree_3).val, True)

        tree_4 = parser.parse('=XOR(0, 0)')
        self.assertEqual(ev.visit(tree_4).val, False)

        tree_5 = parser.parse('=XOR()')
        self.assertIsInstance(ev.visit(tree_5).val, sheets.CellError)

        tree_6 = parser.parse('=XOR("text", TRUE)')
        self.assertIsInstance(ev.visit(tree_6).val, sheets.CellError)

        tree_7 = parser.parse('=XOR(3.14, FALSE)')
        self.assertEqual(ev.visit(tree_7).val, True)

        tree_8 = parser.parse('=XOR(1, 0, 1, 0, 1)')
        self.assertEqual(ev.visit(tree_8).val, True)

        tree_9 = parser.parse('=XOR(true        , FALSE)')
        self.assertEqual(ev.visit(tree_9).val, True)

        tree_10 = parser.parse('=xor(XOR(TRUE, FALSE), TRUE)')
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
        ev = FormulaEvaluator('sheet1', wb, func_directory=wb.func_directory)

        tree_1 = parser.parse('=EXACT("hi", "hi")')
        # print(ev.visit(tree_1))
        self.assertEqual(ev.visit(tree_1).val, True)

        tree_2 = parser.parse('=EXACT("hi", "bye")')
        self.assertEqual(ev.visit(tree_2).val, False)

        # wrong number of arguments
        tree_3 = parser.parse('=EXACT()')
        self.assertIsInstance(ev.visit(tree_3).val, sheets.CellError)

        tree_4 = parser.parse('=EXACT("", "")')
        self.assertEqual(ev.visit(tree_4).val, True)

        tree_5 = parser.parse('=EXACT(123, 456)')
        self.assertEqual(ev.visit(tree_5).val, False)

        tree_6 = parser.parse('=EXACT(EXACT("hi", "hi"), "Hi")')
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
        ev = FormulaEvaluator('sheet1', wb, func_directory=wb.func_directory)

        tree_1 = parser.parse('=IF(1==1, "yes", "no")')
        self.assertEqual(ev.visit(tree_1).val, "yes")

        tree_2 = parser.parse('=IF("blue">"BLUE", "yes", "no")')
        self.assertEqual(ev.visit(tree_2).val, "no")

        # wrong number of arguments
        tree_3 = parser.parse('=IF()')
        self.assertIsInstance(ev.visit(tree_3).val, sheets.CellError)        

        # error propagation
        tree_4 = parser.parse('=IF(1/0, 1, 2)')
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

        wb.set_cell_contents("sheet1", "C2", "=C3")
        wb.set_cell_contents("sheet1", "C3", "=C2")
        wb.set_cell_contents("sheet1", "A11", "=IF(1==1, C2)")
        self.assertIsInstance(wb.get_cell_value('sheet1', 'A11'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A11').get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)

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

        wb.set_cell_contents('sheet1', 'A15', '=IF(1=2, "yes", "")')
        self.assertEqual(wb.get_cell_value('sheet1', 'A15'), "")

        # only 2 arguments 
        wb.set_cell_contents('sheet1', 'A16', '=IF(1=1, "yes")')
        self.assertEqual(wb.get_cell_value('sheet1', 'A16'), "yes")
        
        wb.set_cell_contents('Sheet1', 'A1', '=IF(False, 5)')
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1'), False)

    
    def test_iferror_function(self):
        wb = sheets.Workbook()
        wb.new_sheet()

        parser = lark.Lark.open(lark_path, start='formula')
        ev = FormulaEvaluator('sheet1', wb, func_directory=wb.func_directory)

        tree_1 = parser.parse('=IFERROR(1/0, 5)')
        # print(ev.visit(tree_1))
        self.assertEqual(ev.visit(tree_1).val, decimal.Decimal('5'))

        tree_2 = parser.parse('=IFERROR(1+1)')
        # print(ev.visit(tree_2))
        self.assertEqual(ev.visit(tree_2).val, decimal.Decimal('2'))

        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', '=B1+')
        wb.set_cell_contents('sheet1', 'B1', '=IFERROR(A1, "yes")')
        self.assertIsInstance(wb.get_cell_value('sheet1', 'A1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1').get_type(), sheets.CellErrorType.PARSE_ERROR)
        self.assertEqual(wb.get_cell_value('sheet1', 'B1'), "yes")

        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', '=IFERROR(B1, "B1 is an error")')
        wb.set_cell_contents('sheet1', 'B1', '=IFERROR(A1, "A1 is an error")')
        self.assertIsInstance(wb.get_cell_value('sheet1', 'A1'), sheets.CellError)
        self.assertIsInstance(wb.get_cell_value('sheet1', 'B1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1').get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)
        self.assertEqual(wb.get_cell_value('Sheet1', 'B1').get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)

        # spec example 2
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', '=B1')
        wb.set_cell_contents('sheet1', 'B1', '=A1')
        wb.set_cell_contents('sheet1', 'C1', '=IFERROR(B1, "B1 is an error")')
        self.assertIsInstance(wb.get_cell_value('sheet1', 'A1'), sheets.CellError)
        self.assertIsInstance(wb.get_cell_value('sheet1', 'B1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1').get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)
        self.assertEqual(wb.get_cell_value('Sheet1', 'B1').get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)
        self.assertEqual(wb.get_cell_value('sheet1', 'C1'), 'B1 is an error')   
        
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', '=IFERROR(B1, "B1 is an error")')
        wb.set_cell_contents('sheet1', 'B1', '=IFERROR(A1, "A1 is an error")')
        wb.set_cell_contents('sheet1', 'C1', '=IFERROR(B1, "B1 is an error")')
        self.assertIsInstance(wb.get_cell_value('sheet1', 'A1'), sheets.CellError)
        self.assertIsInstance(wb.get_cell_value('sheet1', 'B1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1').get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)
        self.assertEqual(wb.get_cell_value('Sheet1', 'B1').get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)
        self.assertEqual(wb.get_cell_value('sheet1', 'C1'), 'B1 is an error')

        # TODO: check
        # wb = sheets.Workbook()
        # wb.new_sheet()

        # wb.set_cell_contents('sheet1', 'A1', '=IFERROR(A2, B1)')
        # wb.set_cell_contents('sheet1', 'B1', '=A1')
        # wb.set_cell_contents('sheet1', 'C1', '=5')
        # wb.set_cell_contents('sheet1', 'A2', 'False')

        # self.assertEqual(wb.get_cell_value('sheet1', 'A1'), decimal.Decimal('5'))

        # TODO: check
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', '=IFERROR(B1, "B1 is an error")')
        wb.set_cell_contents('sheet1', 'B1', '=IFERROR(C1, "C1 is an error")')
        wb.set_cell_contents('sheet1', 'C1', '=IFERROR(A1, "A1 is an error")')
        self.assertIsInstance(wb.get_cell_value('sheet1', 'A1'), sheets.CellError)
        self.assertIsInstance(wb.get_cell_value('sheet1', 'B1'), sheets.CellError)
        self.assertIsInstance(wb.get_cell_value('sheet1', 'C1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1').get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)
        self.assertEqual(wb.get_cell_value('Sheet1', 'B1').get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)
        self.assertEqual(wb.get_cell_value('Sheet1', 'C1').get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)
        
        # wrong number of arguments
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', '=IFERROR()')
        self.assertIsInstance(wb.get_cell_value('sheet1', 'A1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1').get_type(), sheets.CellErrorType.TYPE_ERROR)

        # nested
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', '=IFERROR(IFERROR(A1, "A1 is an error"), B1)')
        wb.set_cell_contents('sheet1', 'B1', '=1')
        self.assertIsInstance(wb.get_cell_value('sheet1', 'A1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1').get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)

        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', '=IFERROR(1/0, B1)')
        wb.set_cell_contents('sheet1', 'B1', 'yay')
        self.assertEqual(wb.get_cell_value('sheet1', 'A1'), "yay")

        
    def test_choose_function(self):
        # TODO: not finished yet
        wb = sheets.Workbook()
        wb.new_sheet()

        parser = lark.Lark.open(lark_path, start='formula')
        ev = FormulaEvaluator('sheet1', wb, func_directory=wb.func_directory)

        tree_1 = parser.parse('=choose(1, 0, 1, 2, 3)')
        self.assertEqual(ev.visit(tree_1).val, decimal.Decimal('0'))

        tree_2 = parser.parse('=CHOOSE(2, "a", "b", "c")')
        self.assertEqual(ev.visit(tree_2).val, "b")

        # wrong number of arguments
        tree_3 = parser.parse('=CHOOSE()')
        self.assertIsInstance(ev.visit(tree_3).val, sheets.CellError)

        # out of bounds
        tree_4 = parser.parse('=CHOOSE(0, "a", "b", "c")')
        self.assertIsInstance(ev.visit(tree_3).val, sheets.CellError)

        tree_5 = parser.parse('=CHOOSE(4, "a", "b", "c")')
        self.assertIsInstance(ev.visit(tree_5).val, sheets.CellError)

        tree_6 = parser.parse('=CHOOSE(-1, "a", "b", "c")')
        self.assertIsInstance(ev.visit(tree_6).val, sheets.CellError)

        tree_7 = parser.parse('=CHOOSE(1.5, "a", "b", "c")')
        self.assertIsInstance(ev.visit(tree_7).val, sheets.CellError)

        tree_8 = parser.parse('=choose("text", "a", "b", "c")')
        self.assertIsInstance(ev.visit(tree_8).val, sheets.CellError)

        # TODO: not sure if this is the case
        tree_9 = parser.parse('=CHOOSE(1, 1/0, 2, 3)')
        # print(ev.visit(tree_9).val)
        self.assertIsInstance(ev.visit(tree_8).val, sheets.CellError)

        # set cell contents tests
        wb.set_cell_contents('sheet1', 'A1', '1')
        wb.set_cell_contents('sheet1', 'A2', '=CHOOSE(A1, "apple", "banana", "cherry")')
        self.assertEqual(wb.get_cell_value('sheet1', 'A2'), "apple")

        wb.set_cell_contents('sheet1', 'A1', '2')
        self.assertEqual(wb.get_cell_value('sheet1', 'A2'), "banana")

        wb.set_cell_contents('sheet1', 'B1', '3')
        wb.set_cell_contents('sheet1', 'B2', '=CHOOSE(B1, 10, 20, 30, 40)')
        self.assertEqual(wb.get_cell_value('sheet1', 'B2'), decimal.Decimal('30'))

        wb.set_cell_contents('sheet1', 'C1', '0')
        wb.set_cell_contents('sheet1', 'C2', '=CHOOSE(C1, "a", "b")')
        self.assertIsInstance(wb.get_cell_value('sheet1', 'C2'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'C2').get_type(), sheets.CellErrorType.TYPE_ERROR)

        wb.set_cell_contents('sheet1', 'D1', '5')
        wb.set_cell_contents('sheet1', 'D2', '=CHOOSE(D1, "a", "b")')
        self.assertIsInstance(wb.get_cell_value('sheet1', 'D2'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'D2').get_type(), sheets.CellErrorType.TYPE_ERROR)

        wb.set_cell_contents('sheet1', 'E1', '1.5')
        wb.set_cell_contents('sheet1', 'E2', '=CHOOSE(E1, "a", "b")')
        self.assertIsInstance(wb.get_cell_value('sheet1', 'E2'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'E2').get_type(), sheets.CellErrorType.TYPE_ERROR)

        wb.set_cell_contents('sheet1', 'F1', '"text"')
        wb.set_cell_contents('sheet1', 'F2', '=CHOOSE(F1, "a", "b")')
        self.assertIsInstance(wb.get_cell_value('sheet1', 'F2'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'F2').get_type(), sheets.CellErrorType.TYPE_ERROR)

        # test error propagation
        wb.set_cell_contents("sheet1", "C2", "=C3")
        wb.set_cell_contents("sheet1", "C3", "=C2")
        wb.set_cell_contents("sheet1", "A11", "=CHOOSE(1, C2, C3)")
        self.assertIsInstance(wb.get_cell_value('sheet1', 'A11'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A11').get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)

        # nested
        wb.set_cell_contents('sheet1', 'D1', '1')
        wb.set_cell_contents('sheet1', 'D2', '2')
        wb.set_cell_contents('sheet1', 'D3', '=CHOOSE(D1, CHOOSE(D2, "a", "b", "c"), "d", "e")') # CHOOSE(D2, "a", "b", "c") = b
        self.assertEqual(wb.get_cell_value('sheet1', 'D3'), "b")

    
    def test_isblank_function(self):
        wb = sheets.Workbook()
        wb.new_sheet()

        parser = lark.Lark.open(lark_path, start='formula')
        ev = FormulaEvaluator('sheet1', None, func_directory=wb.func_directory)

        tree_1 = parser.parse('=ISBLANK()')
        self.assertIsInstance(ev.visit(tree_1).val, sheets.CellError)

        tree_2 = parser.parse('=ISBLANK("A")')
        self.assertEqual(ev.visit(tree_2).val, False)

        # wrong number of arguments
        tree_3 = parser.parse('=ISBLANK(1, 2, 3)')
        self.assertIsInstance(ev.visit(tree_3).val, sheets.CellError)

        # set cell contents tests
        wb.set_cell_contents('sheet1', 'B1', '=ISBLANK(A1)')
        self.assertEqual(wb.get_cell_value('sheet1', 'B1'), True)

        wb.set_cell_contents('sheet1', 'B2', '=ISBLANK("")')
        self.assertEqual(wb.get_cell_value('sheet1', 'B2'), False)

        wb.set_cell_contents('sheet1', 'B3', '=ISBLANK(FALSE)')
        self.assertEqual(wb.get_cell_value('sheet1', 'B3'), False)

        wb.set_cell_contents('sheet1', 'B4', '=ISBLANK(0)')
        self.assertEqual(wb.get_cell_value('sheet1', 'B4'), False)

        # error propagation
        wb.set_cell_contents('sheet1', 'B5', '=ISBLANK(1, 2, 3, 4)')
        self.assertIsInstance(wb.get_cell_value('sheet1', 'B5'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'B5').get_type(), sheets.CellErrorType.TYPE_ERROR)

        wb.set_cell_contents('sheet1', 'B6', '=1/0')
        wb.set_cell_contents('sheet1', 'B7', '=ISBLANK(B6)')
        self.assertIsInstance(wb.get_cell_value('sheet1', 'B7'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'B7').get_type(), sheets.CellErrorType.DIVIDE_BY_ZERO)

        wb.set_cell_contents("sheet1", "C2", "=C3")
        wb.set_cell_contents("sheet1", "C3", "=C2")
        wb.set_cell_contents("sheet1", "A11", "=ISBLANK(C2)")
        self.assertIsInstance(wb.get_cell_value('sheet1', 'A11'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A11').get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)
        
        # nested
        wb.set_cell_contents('sheet1', 'D1', '')
        wb.set_cell_contents('sheet1', 'D2', '=ISBLANK(D1)')
        wb.set_cell_contents('sheet1', 'D3', '=ISBLANK(D2)')  # ISBLANK(TRUE)
        self.assertEqual(wb.get_cell_value('sheet1', 'D3'), False)


    def test_iserror_function(self):
        # from spec
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', '=B1+')
        wb.set_cell_contents('sheet1', 'B1', '=ISERROR(A1)')
        self.assertIsInstance(wb.get_cell_value('sheet1', 'A1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1').get_type(), sheets.CellErrorType.PARSE_ERROR)
        self.assertEqual(wb.get_cell_value('sheet1', 'B1'), True)

        # spec example 2
        wb.set_cell_contents('sheet1', 'A1', '=B1')
        wb.set_cell_contents('sheet1', 'B1', '=A1')
        wb.set_cell_contents('sheet1', 'C1', '=ISERROR(B1)')
        self.assertIsInstance(wb.get_cell_value('sheet1', 'A1'), sheets.CellError)
        self.assertIsInstance(wb.get_cell_value('sheet1', 'B1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1').get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)
        self.assertEqual(wb.get_cell_value('Sheet1', 'B1').get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)
        self.assertEqual(wb.get_cell_value('sheet1', 'C1'), True)

        # This is failing
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', '=ISERROR(B1)')
        wb.set_cell_contents('sheet1', 'B1', '=ISERROR(A1)')
        wb.set_cell_contents('sheet1', 'C1', '=ISERROR(B1)')

        self.assertIsInstance(wb.get_cell_value('sheet1', 'A1'), sheets.CellError)
        self.assertIsInstance(wb.get_cell_value('sheet1', 'B1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1').get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)
        self.assertEqual(wb.get_cell_value('Sheet1', 'B1').get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)
        self.assertEqual(wb.get_cell_value('sheet1', 'C1'), True)

        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', '=ISERROR(B1)')
        wb.set_cell_contents('sheet1', 'B1', '=ISERROR(C1)')
        wb.set_cell_contents('sheet1', 'C1', '=ISERROR(A1)')
        self.assertIsInstance(wb.get_cell_value('sheet1', 'A1'), sheets.CellError)
        self.assertIsInstance(wb.get_cell_value('sheet1', 'B1'), sheets.CellError)
        self.assertIsInstance(wb.get_cell_value('sheet1', 'C1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1').get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)
        self.assertEqual(wb.get_cell_value('Sheet1', 'B1').get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)
        self.assertEqual(wb.get_cell_value('Sheet1', 'C1').get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)
        
        # wrong number of arguments
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', '=ISERROR()')
        self.assertIsInstance(wb.get_cell_value('sheet1', 'A1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1').get_type(), sheets.CellErrorType.TYPE_ERROR)

        # nested
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', '=ISERROR(IFERROR(A1), B1)')
        wb.set_cell_contents('sheet1', 'B1', '=1')
        self.assertIsInstance(wb.get_cell_value('sheet1', 'A1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1').get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)

        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', '=ISERROR(1/0)')
        self.assertEqual(wb.get_cell_value('sheet1', 'A1'), True)

        # not errors
        wb = sheets.Workbook()
        wb.new_sheet()

        wb.set_cell_contents('sheet1', 'A1', '=ISERROR(10)')
        wb.set_cell_contents('sheet1', 'A2', '=ISERROR(3.14)')
        wb.set_cell_contents('sheet1', 'A3', '=ISERROR("Hello")')
        wb.set_cell_contents('sheet1', 'A4', '=ISERROR(TRUE)')
        wb.set_cell_contents('sheet1', 'A5', '=ISERROR(A1)')


        self.assertEqual(wb.get_cell_value('sheet1', 'A1'), False)
        self.assertEqual(wb.get_cell_value('sheet1', 'A2'), False)
        self.assertEqual(wb.get_cell_value('sheet1', 'A3'), False)
        self.assertEqual(wb.get_cell_value('sheet1', 'A4'), False)
        self.assertEqual(wb.get_cell_value('sheet1', 'A5'), False)

        # more basic tests
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'B1', '=1/0')
        wb.set_cell_contents('sheet1', 'A1', '=ISERROR(B1)')

        self.assertEqual(wb.get_cell_value('sheet1', 'A1'), True)

        wb.set_cell_contents('sheet1', 'B1', '=1')
        wb.set_cell_contents('sheet1', 'A1', '=ISERROR(B1)')

        self.assertEqual(wb.get_cell_value('sheet1', 'A1'), False)


    def test_version_function(self):
        wb = sheets.Workbook()
        wb.new_sheet()

        parser = lark.Lark.open(lark_path, start='formula')
        tree_1 = parser.parse('=VERSION()')

        
        ev = FormulaEvaluator('sheet1', wb, func_directory=wb.func_directory)
        
        self.assertEqual(ev.visit(tree_1).val, sheets.__version__)

        tree_3 = parser.parse('=VERSION(1, 2)')

        self.assertIsInstance(ev.visit(tree_3).val, sheets.CellError)

        # set cell contents tests
        wb.set_cell_contents('sheet1', 'A1', '=version()')
        self.assertEqual(wb.get_cell_value('sheet1', 'A1'), sheets.__version__)

    def test_indirect_function(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.new_sheet('Sheet 4')

        # basic
        wb.set_cell_contents('sheet1', 'A1', '\'C1')
        wb.set_cell_contents('Sheet1', 'C1', '4')
        wb.set_cell_contents('sheet1', 'B1', '=INDIRECT(A1)')
        self.assertEqual(wb.get_cell_value('Sheet1', 'B1'), decimal.Decimal('4'))
        
        # from the spec
        wb.set_cell_contents('sheet1', 'A1', '=B1')
        wb.set_cell_contents('sheet1', 'B1', '=INDIRECT("A1")')
        self.assertIsInstance(wb.get_cell_value('sheet1', 'A1'), sheets.CellError)
        self.assertIsInstance(wb.get_cell_value('sheet1', 'B1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1').get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)
        self.assertEqual(wb.get_cell_value('Sheet1', 'B1').get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)

        # with quotes
        wb.set_cell_contents('Sheet1', 'A1', "=INDIRECT(\"'Sheet 4'!B1\")")
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1'), 0)

        wb = sheets.Workbook()
        wb.new_sheet()

        # from spec
        wb.set_cell_contents('sheet1', 'A1', '=B1')
        wb.set_cell_contents('sheet1', 'B1', '=INDIRECT(A1)')

        self.assertIsInstance(wb.get_cell_value('sheet1', 'A1'), sheets.CellError)
        self.assertIsInstance(wb.get_cell_value('sheet1', 'B1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1').get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)
        self.assertEqual(wb.get_cell_value('Sheet1', 'B1').get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)

        # from the spec
        wb.set_cell_contents('sheet1', 'A1', '=B1')
        wb.set_cell_contents('sheet1', 'B1', '=INDIRECT("Sheet1!A1")')

        self.assertIsInstance(wb.get_cell_value('sheet1', 'A1'), sheets.CellError)
        self.assertIsInstance(wb.get_cell_value('sheet1', 'B1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1').get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)
        self.assertEqual(wb.get_cell_value('Sheet1', 'B1').get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)

        # more complicated arguments 
        wb = sheets.Workbook()
        wb.new_sheet()

        # from the spec
        wb.set_cell_contents('sheet1', 'A1', '=INDIRECT("Sheet" & 1 & "!B1")')
        wb.set_cell_contents('sheet1', 'B1', '1')

        self.assertEqual(wb.get_cell_value('sheet1', 'A1'), decimal.Decimal('1'))

        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', '=1')
        wb.set_cell_contents('sheet1', 'B1', '=INDIRECT("$A$1")')

        # self.move_cells('sheet1', 'A1', 'A1', 'B1')
        self.assertEqual(wb.get_cell_value('Sheet1', 'B1'), decimal.Decimal('1'))

        wb = sheets.Workbook()
        wb.new_sheet()
        wb.new_sheet('Another Sheet')
        wb.set_cell_contents('sheet1', 'A1', '=INDIRECT("\'another sheet\'!B1")')
        wb.set_cell_contents('another sheet', 'B1', '5')
        self.assertEqual(wb.get_cell_value('sheet1', 'A1'), decimal.Decimal('5'))
        
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.new_sheet('She-et1')
        wb.set_cell_contents('sheet1', 'A1', '=INDIRECT("\'she-et1\'!B1")')
        wb.set_cell_contents('she-et1', 'B1', '5')
        self.assertEqual(wb.get_cell_value('sheet1', 'A1'), decimal.Decimal('5'))
       
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', '=INDIRECT("A" & "A2")')
        wb.set_cell_contents('sheet1', 'AA2', '5')
        self.assertEqual(wb.get_cell_value('sheet1', 'A1'), decimal.Decimal('5'))
        
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', '=INDIRECT("ZZZZ99999")')
        self.assertIsInstance(wb.get_cell_value('sheet1', 'A1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1').get_type(), sheets.CellErrorType.BAD_REFERENCE)

        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', '=1/0')
        wb.set_cell_contents('sheet1', 'B1', '=INDIRECT("A1")')

        self.assertIsInstance(wb.get_cell_value('sheet1', 'B1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'B1').get_type(), sheets.CellErrorType.DIVIDE_BY_ZERO)

        # various kinds of bad inputs
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', '=INDIRECT("123")')
        self.assertIsInstance(wb.get_cell_value('sheet1', 'A1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1').get_type(), sheets.CellErrorType.BAD_REFERENCE)

        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', '=INDIRECT("A!1")')
        self.assertIsInstance(wb.get_cell_value('sheet1', 'A1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1').get_type(), sheets.CellErrorType.BAD_REFERENCE)

        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', '=INDIRECT(TRUE)')
        self.assertIsInstance(wb.get_cell_value('sheet1', 'A1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1').get_type(), sheets.CellErrorType.BAD_REFERENCE)

        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', '=INDIRECT("")')
        self.assertIsInstance(wb.get_cell_value('sheet1', 'A1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1').get_type(), sheets.CellErrorType.BAD_REFERENCE)

        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', '=INDIRECT("A" & "XYZ")')
        self.assertIsInstance(wb.get_cell_value('sheet1', 'A1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1').get_type(), sheets.CellErrorType.BAD_REFERENCE)

        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('sheet1', 'A1', '=INDIRECT("A" & "2")')
        wb.set_cell_contents('sheet1', 'A2', '10')
        self.assertEqual(wb.get_cell_value('sheet1', 'A1'), decimal.Decimal('10'))

        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('Sheet1', 'A1', '=INDIRECT("$A" & 4)')
        wb.set_cell_contents('Sheet1', 'A4', '5')
        self.assertEqual(wb.get_cell_value('sheet1', 'A1'), decimal.Decimal('5'))

        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('Sheet1', 'A1', '=INDIRECT("A" & "1")')
        self.assertIsInstance(wb.get_cell_value('sheet1', 'A1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('sheet1', 'A1').get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)

        wb = sheets.Workbook()
        wb.new_sheet()
        wb.new_sheet('Sheet 2')
        wb.set_cell_contents('Sheet1', 'A1', '=INDIRECT("Sheet 2!A1")')
        self.assertIsInstance(wb.get_cell_value('sheet1', 'A1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('sheet1', 'A1').get_type(), sheets.CellErrorType.BAD_REFERENCE)

        # TODO: referencing a non existent sheet 
        # invalid_sheet_names = ['', ' Sheet', '~', 'Lorem ipsum', 'Sheet\' name', 'Sheet \" name']

        # wb = sheets.Workbook()
        # wb.new_sheet('Sheet1')
        # for sheet_name in invalid_sheet_names:
        #     wb.set_cell_contents('Sheet1', 'A1', f'=INDIRECT("{sheet_name}!B1")')
        #     self.assertIsInstance(wb.get_cell_value('Sheet1', 'A1'), sheets.CellError)
        #     self.assertEqual(wb.get_cell_value('Sheet1', 'A1').get_type(), sheets.CellErrorType.BAD_REFERENCE)

    def test_move_function(self):
        wb = sheets.Workbook()
        wb.new_sheet()

        wb.set_cell_contents('Sheet1', 'A1', '=VERSION()')
        wb.set_cell_contents('Sheet1', 'B1', '=INDIRECT(A1)')
        wb.set_cell_contents('Sheet1', 'C1', '=IF(True, 1, 1)')
        wb.move_cells('Sheet1', 'A1', 'C1', 'A2')
        
        self.assertEqual(wb.get_cell_contents('Sheet1', 'A2'), '=VERSION()')
        self.assertEqual(wb.get_cell_contents('Sheet1', 'B2'), '=INDIRECT(A2)')
        self.assertEqual(wb.get_cell_contents('Sheet1', 'C2'), '=IF(True, 1, 1)')

    def test_general(self):
        wb = sheets.Workbook()
        wb.new_sheet()

        wb.set_cell_contents('Sheet1', 'A1', '=or(AnD(Z1 > 5, B1 < 2), ANd(C1 < 6, D1 = 14))')
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1'), False)

        wb.set_cell_contents('Sheet1', 'A1', '=OR(AND(Z1 = 0, B1 = 0), AND(C1 < 6, D1 = 14))')
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1'), True)

        # TODO: 
        wb.set_cell_contents('Sheet1', 'A1', '=IFERROR(ISERROR(1/0), FALSE)')
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1'), True)

        wb.set_cell_contents('Sheet1', 'A1', '=ISERROR(IFERROR(1/0, FALSE))')
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1'), False)

    def test_conditionals(self):
        wb = sheets.Workbook()
        wb.new_sheet()

        wb.set_cell_contents('Sheet1', 'A1', '=IF(A2, B1, C1)')
        wb.set_cell_contents('Sheet1', 'A2', 'True')
        wb.set_cell_contents('Sheet1', 'B1', '=A1')
        wb.set_cell_contents('Sheet1', 'C1', '5')
        self.assertIsInstance(wb.get_cell_value('Sheet1', 'A1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1').get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)

        # from spec
        wb.set_cell_contents('Sheet1', 'A1', '=IF(A2, B1, C1)')
        wb.set_cell_contents('Sheet1', 'A2', 'False')
        wb.set_cell_contents('Sheet1', 'B1', '=A1')
        wb.set_cell_contents('Sheet1', 'C1', '5')

        self.assertEqual(wb.get_cell_value('Sheet1', 'A1'), 5)
        self.assertEqual(wb.get_cell_value('Sheet1', 'B1'), 5)
        self.assertEqual(wb.get_cell_value('Sheet1', 'C1'), 5)

        # iferror
        wb.set_cell_contents('Sheet1', 'A1', '=IFERROR(1 + 1/0, B1)')
        wb.set_cell_contents('Sheet1', 'B1', '=A1')

        self.assertIsInstance(wb.get_cell_value('Sheet1', 'A1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1').get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)

        wb.set_cell_contents('Sheet1', 'A1', '=IFERROR(1, B1)')
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1'), 1)

        # choose
        wb.set_cell_contents('Sheet1', 'A1', '=CHOOSE(2, B1, 3)')
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1'), 3)

        wb.set_cell_contents('Sheet1', 'A1', '=CHOOSE(1, B1, 3)')
        self.assertIsInstance(wb.get_cell_value('Sheet1', 'A1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1').get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)



    def test_alt(self):
        wb = sheets.Workbook()
        wb.new_sheet()

        # wb.set_cell_contents('Sheet1', 'A1', '=IF()')
        wb.set_cell_contents('Sheet1', 'A1', '=IF(True, 5)')
        # wb.set_cell_contents('Sheet1', 'A1', '=IF(True, 5)')

if __name__ == "__main__":
    cov = coverage.Coverage()
    cov.start()
    unittest.main()
    cov.stop()
    cov.save()
    cov.html_report()