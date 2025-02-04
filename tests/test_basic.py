# python -m unittest discover -s tests
import unittest
import coverage
import sheets
import os
import lark
from sheets.interpreter import FormulaEvaluator
from sheets.transformer import SheetNameExtractor
import decimal
import json
# import timeit
import contextlib
from io import StringIO

current_dir = os.path.dirname(os.path.abspath(__file__))
lark_path = os.path.join(current_dir, '../sheets/formulas.lark')

class BasicTests(unittest.TestCase):

    # def time_operation(self, stmt, setup="pass", number=100, repeat=5):
    #     """Helper function to time an operation and return average time."""
    #     times = timeit.repeat(stmt=stmt, setup=setup, number=number, repeat=repeat)
    #     average_time = sum(times) / len(times)
    #     print(f"Times: {times}")
    #     print(f"Average time: {average_time}")
    #     return average_time

    def test_bad_reference_edge(self):
        wb = sheets.Workbook()
        wb.new_sheet()

        wb.set_cell_contents('Sheet1', 'A1', '=Sheet2!A1')
        wb.new_sheet()
        self.assertEqual(wb.get_sheet_extent('Sheet2'), (0, 0))
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1'), 0)

        wb.del_sheet('Sheet2')
        self.assertIsInstance(wb.get_cell_value('Sheet1', 'A1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1').get_type(), sheets.CellErrorType.BAD_REFERENCE)

        wb.new_sheet()
        self.assertEqual(wb.get_sheet_extent('Sheet2'), (0, 0))
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1'), 0)
    
    def test_unset_cells(self):
        wb = sheets.Workbook()
        wb.new_sheet()

        # addition/subtraction
        wb.set_cell_contents('Sheet1', 'A1', '=1 - D1')
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1'), 1)

        # multiplication/division
        wb.set_cell_contents('Sheet1', 'A1', '=1 * D1')
        self.assertEqual(str(wb.get_cell_value('Sheet1', 'A1')), '0')
        
        wb.set_cell_contents('Sheet1', 'A1', '=1 / D1')
        self.assertIsInstance(wb.get_cell_value('Sheet1', 'A1'), sheets.CellError)
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1').get_type(), sheets.CellErrorType.DIVIDE_BY_ZERO)

        # unary
        wb.set_cell_contents('Sheet1', 'A1', '=-D1')
        self.assertEqual(str(wb.get_cell_value('Sheet1', 'A1')), '0')

        # concat
        wb.set_cell_contents('Sheet1', 'A1', '="test" & D1')
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1'), 'test')

        # parentheses
        wb.set_cell_contents('Sheet1', 'A1', '=(D1)')
        self.assertEqual(str(wb.get_cell_value('Sheet1', 'A1')), '0')
    
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

        wb.set_cell_contents('Sheet1', 'A1', '=    4  *    A2  ')
        wb.set_cell_contents('Sheet1', 'A2', '=     4.5     ')
        self.assertEqual(str(wb.get_cell_value('Sheet1', 'A1')), '18')

        wb.set_cell_contents('Sheet1', 'A1', '=B1 * C1')
        wb.set_cell_contents('Sheet1', 'B1', '=0.5')
        wb.set_cell_contents('Sheet1', 'C1', '=0.2')
        self.assertEqual(str(wb.get_cell_value('Sheet1', 'A1')), '0.1')

        wb.set_cell_contents('Sheet1', 'A1', '=1.5 * 2')
        self.assertEqual(str(wb.get_cell_value('Sheet1', 'A1')), '3')

        wb.set_cell_contents('Sheet1', 'A1', '=-4.0000')
        self.assertEqual(str(wb.get_cell_value('Sheet1', 'A1')), '-4')

        wb.set_cell_contents('Sheet1', 'A1', '1')
        wb.set_cell_contents('Sheet1', 'B1', '2')
        wb.set_cell_contents('Sheet1', 'C1', '=A1 & B1')
        self.assertEqual(wb.get_cell_value('Sheet1', 'C1'), '12')

    def test_cell_reference(self):
        wb = sheets.Workbook()
        wb.new_sheet()

        wb.set_cell_contents('Sheet1', 'D2', '2')
        wb.set_cell_contents('Sheet1', 'D3', '=1 + D2')

        # print(wb.sheets['sheet1'].cells[row_idx][col_idx].outgoing[0].location)

        # self.assertEqual(wb.sheets['sheet1'].cells[2][3].outgoing[0].location, 'D2')
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
        self.assertEqual(len(wb.graph.ingoing_get('Sheet1', 'A2')), 1)
        self.assertEqual(len(wb.graph.ingoing_get('Sheet1', 'A3')), 1)
        wb.set_cell_contents('Sheet1', 'A1', '=A3 + 4')
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1'), 5)
        self.assertEqual(len(wb.graph.ingoing_get('Sheet1', 'A2')), 0)
        self.assertEqual(len(wb.graph.ingoing_get('Sheet1', 'A3')), 2)
    
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
    
    def test_transformer(self):
        parser = lark.Lark.open(lark_path, start='formula')
        tree_1 = parser.parse('=1 + 1')

        sne = SheetNameExtractor('Sheet1', 'SheetBla')
        self.assertEqual(sne.transform(tree_1), '1 + 1')
        # print(sne.transform(tree_1))

        tree_2 = parser.parse('=1 + Sheet1!A1')
        self.assertEqual(sne.transform(tree_2), '1 + SheetBla!A1')

        tree_3 = parser.parse('=1*A5')
        self.assertEqual(sne.transform(tree_3), '1 * A5')

        tree_4 = parser.parse('=-Sheet1!A1')
        self.assertEqual(sne.transform(tree_4), '-SheetBla!A1')

        tree_5 = parser.parse('=(((((Sheet1!B1)))))')
        self.assertEqual(sne.transform(tree_5), '(((((SheetBla!B1)))))')

    def test_automatic_updates(self):
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

        # test case with topsort
        wb.new_sheet('Sheet4')
        wb.set_cell_contents('Sheet4', 'A1', '=B1+C1')
        wb.set_cell_contents('Sheet4', 'B1', '=C1 + E1 + F1')
        wb.set_cell_contents('Sheet4', 'C1', '1')
        self.assertEqual(wb.get_cell_value('Sheet4', 'A1'), 2)

        wb.set_cell_contents('Sheet4', 'C1', '2')
        self.assertEqual(wb.get_cell_value('Sheet4', 'A1'), 4)
    
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
    
    def test_notify_basic(self):
        # basic test
        wb = sheets.Workbook()

        def on_cells_changed(workbook, changed_cells):
            print(f'Cell(s) changed: {changed_cells}')
        wb.notify_cells_changed(on_cells_changed)

        temp_stdout = StringIO()
        with contextlib.redirect_stdout(temp_stdout):
            wb.new_sheet()
            wb.set_cell_contents("Sheet1", "A1", "'123")
            wb.set_cell_contents("Sheet1", "C1", "=A1+B1")
            wb.set_cell_contents("Sheet1", "B1", "5.3")
        output = temp_stdout.getvalue()
        self.assertEqual(output, (
            "Cell(s) changed: [('sheet1', 'a1')]\n"
            "Cell(s) changed: [('sheet1', 'c1')]\n"
            "Cell(s) changed: [('sheet1', 'b1'), ('sheet1', 'c1')]\n"
        ))

        # exceptions raised by notificaiton function should not affect anything
        wb2 = sheets.Workbook()

        def on_cells_changed_error(workbook, changed_cells):
            print(f'Cell(s) changed: {changed_cells[len(changed_cells)]}')
        wb2.notify_cells_changed(on_cells_changed_error)

        temp_stdout = StringIO()
        with contextlib.redirect_stdout(temp_stdout):
            wb2.new_sheet()
            wb2.set_cell_contents("Sheet1", "A1", "'123")
            wb2.set_cell_contents("Sheet1", "C1", "=A1+B1")
            wb2.set_cell_contents("Sheet1", "B1", "5.3")
        output = temp_stdout.getvalue()
        self.assertEqual(output, '')
        self.assertEqual(wb2.get_cell_contents('Sheet1', 'A1'), "'123")
        self.assertEqual(wb2.get_cell_contents('Sheet1', 'C1'), "=A1+B1")
        self.assertEqual(wb2.get_cell_contents('Sheet1', 'B1'), "5.3")

        # test all of these features:
        # multiple notification functions (including some that fail)
        # cells that don't change don't get notified
        wb3 = sheets.Workbook()

        def count_cells_changed(workbook, changed_cells):
            print(f'Number of Cell(s) changed: {len(changed_cells)}')
        wb3.notify_cells_changed(count_cells_changed)
        wb3.notify_cells_changed(on_cells_changed_error)
        wb3.notify_cells_changed(on_cells_changed)

        temp_stdout = StringIO()
        with contextlib.redirect_stdout(temp_stdout):
            wb3.new_sheet()
            wb3.set_cell_contents("Sheet1", "A1", "=B1")
            wb3.set_cell_contents("Sheet1", "B1", "=A1+C1")
            wb3.set_cell_contents("Sheet1", "C1", "5.3")
        output = temp_stdout.getvalue()
        self.assertEqual(output, (
            "Number of Cell(s) changed: 1\n"
            "Cell(s) changed: [('sheet1', 'a1')]\n"
            "Number of Cell(s) changed: 2\n"
            "Cell(s) changed: [('sheet1', 'b1'), ('sheet1', 'a1')]\n"
            "Number of Cell(s) changed: 1\n"
            "Cell(s) changed: [('sheet1', 'c1')]\n"
        ))

    def test_notify_sheet_operations(self):
        # test that adding sheet causes notification
        wb = sheets.Workbook()
        def on_cells_changed(workbook, changed_cells):
            print(f'Cell(s) changed: {changed_cells}')
        wb.notify_cells_changed(on_cells_changed)

        temp_stdout = StringIO()
        with contextlib.redirect_stdout(temp_stdout):
            wb.new_sheet()
            wb.set_cell_contents("Sheet1", "A1", "=1 + Sheet2!A1")
            wb.new_sheet()
        output = temp_stdout.getvalue()
        self.assertEqual(output, (
            "Cell(s) changed: [('sheet1', 'a1')]\n"
            "Cell(s) changed: [('sheet1', 'a1')]\n"
        ))

        # test deleting sheet causes notification
        temp_stdout = StringIO()
        with contextlib.redirect_stdout(temp_stdout):
            wb.set_cell_contents('Sheet2', 'A5', '4.5')
            wb.del_sheet('Sheet2')
        output = temp_stdout.getvalue()
        self.assertEqual(output, (
            "Cell(s) changed: [('sheet2', 'a5')]\n"
            "Cell(s) changed: [('sheet1', 'a1')]\n"
        ))

        # copying causes notification:
        # Sheet1 -> Sheet2 -> Sheet1_1
        temp_stdout = StringIO()
        with contextlib.redirect_stdout(temp_stdout):
            wb.new_sheet('Sheet2')
            wb.set_cell_contents("Sheet2", "A1", "=1 + Sheet1_1!A1")
            self.assertEqual(wb.get_cell_value('Sheet2', 'A1').get_type(), sheets.CellErrorType.BAD_REFERENCE)
            self.assertEqual(wb.get_cell_value('Sheet1', 'A1').get_type(), sheets.CellErrorType.BAD_REFERENCE)
            wb.copy_sheet('Sheet1')
            self.assertEqual(wb.get_cell_value('Sheet2', 'A1').get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)
            self.assertEqual(wb.get_cell_value('Sheet1', 'A1').get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)
        output = temp_stdout.getvalue()
        self.assertEqual(output, (
            "Cell(s) changed: [('sheet1', 'a1')]\n" # from new_sheet
            "Cell(s) changed: [('sheet2', 'a1'), ('sheet1', 'a1')]\n" # from set_cell_contents
            "Cell(s) changed: [('sheet2', 'a1'), ('sheet1', 'a1')]\n" # from the new_sheet call inside copy_sheet
            "Cell(s) changed: [('sheet1_1', 'a1'), ('sheet2', 'a1'), ('sheet1', 'a1')]\n" # from set_cell_contents call inside copy_sheet
        ))
        # TODO: verify that functionality above is correct (should copy sheet cause two cell notifications?)

        # TODO: renaming causes notification

if __name__ == "__main__":
    cov = coverage.Coverage()
    cov.start()
    unittest.main()
    cov.stop()
    cov.save()
    cov.html_report()

# TODO: Note that adding sheets, renaming sheets, copying sheets and deleting sheets can all cause cell-values to be updated.
    # adding sheets, deleting sheets implementing - but not copying/renaming sheets yet
    # test case - wb.set_cell_contents('Sheet1', 'A1', '=Sheet1_1!A2')
    # then wb.copy_sheet('Sheet1')
# TODO: quoted sheet name tests