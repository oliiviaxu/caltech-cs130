# python -m unittest discover -s tests
import unittest
import coverage
import sheets
import os
import lark
import decimal
import json
import contextlib
import cProfile
import random

class CycleDetectionTests(unittest.TestCase):

    @staticmethod
    def make_large_cycle(wb, num_cells_in_cycle):
        for i in range(1, num_cells_in_cycle):
            wb.set_cell_contents('Sheet1', f'A{i}', f'=A{i + 1}')

        wb.set_cell_contents('Sheet1', f'A{num_cells_in_cycle}', '=A1')

    def test_large_cycle(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        num_cells_in_cycle = 100
        
        CycleDetectionTests.make_large_cycle(wb, num_cells_in_cycle)

        cell_a1 = wb.sheets['sheet1'].get_cell('A1')
        self.assertEqual(wb.detect_cycle(cell_a1), True)        
    
    def test_small_cycles(self):

        num_small_cycles = 100

        # small cycles with 2 cells
        for _ in range(num_small_cycles):
            wb = sheets.Workbook()
            sheet_num_1, sheet_num_2 = random.randint(0, 10), random.randint(0, 10)
            sheet_name_1, sheet_name_2 = f'sheet{sheet_num_1}', f'sheet{sheet_num_2}'

            wb.new_sheet(sheet_name_1)
            if sheet_name_1 != sheet_name_2:
                wb.new_sheet(sheet_name_2)

            wb.set_cell_contents(sheet_name_1, 'A1', f'={sheet_name_2}!A2')
            wb.set_cell_contents(sheet_name_2, 'A2', f'={sheet_name_1}!A1')

            cell = wb.sheets[sheet_name_1.lower()].get_cell('A1')
            self.assertEqual(wb.detect_cycle(cell), True)
        
        # small cycles with 3 cells
        for _ in range(num_small_cycles):
            wb = sheets.Workbook()
            sheet_num_1, sheet_num_2, sheet_num_3 = random.randint(0, 5), random.randint(0, 5), random.randint(0, 5)
            sheet_name_1, sheet_name_2, sheet_name_3 = f'sheet{sheet_num_1}', f'sheet{sheet_num_2}', f'sheet{sheet_num_3}'

            wb.new_sheet(sheet_name_1)
            if sheet_name_1!= sheet_name_2:
                wb.new_sheet(sheet_name_2)
            if sheet_name_1!= sheet_name_3 and sheet_name_2!= sheet_name_3:
                wb.new_sheet(sheet_name_3)

            wb.set_cell_contents(sheet_name_1, 'A1', f'={sheet_name_2}!A2')
            wb.set_cell_contents(sheet_name_2, 'A2', f'={sheet_name_3}!A3')
            wb.set_cell_contents(sheet_name_3, 'A3', f'={sheet_name_1}!A1')

            cell = wb.sheets[sheet_name_1.lower()].get_cell('A1')
            self.assertEqual(wb.detect_cycle(cell), True)


    def test_cell_in_multiple_cycles(self):
        wb = sheets.Workbook()
        wb.new_sheet()

        wb.set_cell_contents('Sheet1', 'A1', '=B1')
        wb.set_cell_contents('Sheet1', 'B1', '=C1')
        wb.set_cell_contents('Sheet1', 'C1', '=A1')

        wb.set_cell_contents('Sheet1', 'D1', '=E1')
        wb.set_cell_contents('Sheet1', 'E1', '=A1')

        wb.set_cell_contents('Sheet1', 'F1', '=A1')

        cell_a1 = wb.sheets['sheet1'].get_cell('A1')
        self.assertEqual(wb.detect_cycle(cell_a1), True)
    
    def test_make_break_cycle(self):
        # gemini output
        wb = sheets.Workbook()
        wb.new_sheet()

        wb.set_cell_contents('Sheet1', 'A1', '=B1')
        wb.set_cell_contents('Sheet1', 'B1', '1')
        cell_a1 = wb.sheets['sheet1'].get_cell('A1')
        self.assertEqual(wb.detect_cycle(cell_a1), False)

        wb.set_cell_contents('Sheet1', 'B1', '=A1')
        self.assertEqual(wb.detect_cycle(cell_a1), True)

        wb.set_cell_contents('Sheet1', 'B1', '1')
        self.assertEqual(wb.detect_cycle(cell_a1), False)

        wb.set_cell_contents('Sheet1', 'B1', '=C1')
        wb.set_cell_contents('Sheet1', 'C1', '=A1')
        self.assertEqual(wb.detect_cycle(cell_a1), True)

        wb.set_cell_contents('Sheet1', 'A1', '=1')
        self.assertEqual(wb.detect_cycle(cell_a1), False)

        # END of gemini output

        # make a large cycle andthen break
        wb_2 = sheets.Workbook()
        wb_2.new_sheet()
        num_cells_in_cycle = 100
        CycleDetectionTests.make_large_cycle(wb_2, num_cells_in_cycle)
        wb_2.set_cell_contents('Sheet1', f'A{num_cells_in_cycle}', '=1')
        cell_a1 = wb_2.sheets['sheet1'].get_cell('A1')
        self.assertEqual(wb_2.detect_cycle(cell_a1), False) 

if __name__ == "__main__":
    cov = coverage.Coverage()
    cov.start()
    unittest.main()
    cov.stop()
    cov.save()
    cov.html_report()