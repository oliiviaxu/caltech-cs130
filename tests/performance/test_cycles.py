# python -m unittest tests.performance.test_cycles

import os
import unittest
import coverage
import sheets
import cProfile
from pstats import Stats

from .testStructures import create_large_cycle, create_small_cycles, create_chain

current_dir = os.path.dirname(os.path.abspath(__file__))
dir = os.path.join(current_dir, 'cProfile_output/')

class CycleDetectionTests(unittest.TestCase):

    def setUp(self):
        self.pr = cProfile.Profile()
        self.pr.enable()
        print(f"\n<<<--- Starting Test: {self._testMethodName}")
    
    def tearDown(self):
        self.pr.disable()

        class_name = self.__class__.__name__
        test_name = self._testMethodName
        dir_path = os.path.join(current_dir, 'cProfile_output', class_name)
        os.makedirs(dir_path, exist_ok=True)

        test_name = self._testMethodName
        file_path = os.path.join(dir_path, f'{test_name}_profile.txt')

        with open(file_path, 'w') as f:
            p = Stats(self.pr, stream=f)
            p.strip_dirs()
            p.sort_stats('cumtime')
            p.print_stats()

        print(f"--->>> Ending Test: {self._testMethodName}\n")
    
    def test_large_cycle(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        num_cells_in_cycle = 1000
        
        create_large_cycle(wb, num_cells_in_cycle)

        cell_a1 = wb.sheets['sheet1'].get_cell('A1')
        self.assertEqual(wb.detect_cycle(cell_a1), True)
    
    def test_small_cycles(self):
        wb = sheets.Workbook()
        num_cycles = 1000
        _, sheet_name = wb.new_sheet()
        create_small_cycles(wb, sheet_name, num_cycles)

        for i in range(1, num_cycles + 1):
            self.assertIsInstance(wb.get_cell_value(sheet_name, f'A{i}'), sheets.CellError)
            self.assertEqual(wb.get_cell_value(sheet_name, f'A{i}').get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)

            self.assertIsInstance(wb.get_cell_value(sheet_name, f'B{i}'), sheets.CellError)
            self.assertEqual(wb.get_cell_value(sheet_name, f'B{i}').get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)

    def test_cell_in_multiple_cycles(self):
        wb = sheets.Workbook()
        _, sheet_name = wb.new_sheet()
        num_cycles = 1000

        cell_in_multi_cycles = 'C1'

        create_small_cycles(wb, sheet_name, num_cycles)
        
        for i in range(1, num_cycles + 1):
            wb.set_cell_contents(sheet_name, cell_in_multi_cycles, f'=B{i}')

        cell_c1 = wb.sheets[sheet_name.lower()].get_cell(cell_in_multi_cycles)
        self.assertEqual(wb.detect_cycle(cell_c1), True)
    
    def test_make_break_cycle(self):
        # crete small cycle
        wb = sheets.Workbook()
        num_cycles = 1000
        _, sheet_name = wb.new_sheet()
        create_small_cycles(wb, sheet_name, num_cycles)

        # break
        for i in range(1, num_cycles):
            wb.set_cell_contents('Sheet1', f'A{i}', '1')
            cell_a1 = wb.sheets['sheet1'].get_cell(f'A{i}')
            self.assertEqual(wb.detect_cycle(cell_a1), False)
            self.assertEqual(wb.detect_cycle(cell_a1), False)
        
        # create large cycle
        wb_2 = sheets.Workbook()
        wb_2.new_sheet()
        num_cells_in_cycle = 1000
        create_large_cycle(wb_2, num_cells_in_cycle)
        
        # break
        wb_2.set_cell_contents('Sheet1', f'A{num_cells_in_cycle}', '1')
        cell_a1 = wb_2.sheets['sheet1'].get_cell('A1')
        self.assertEqual(wb.detect_cycle(cell_a1), False)
        self.assertEqual(wb.detect_cycle(cell_a1), False)

        # gemini
        wb_3 = sheets.Workbook()
        wb_3.new_sheet()

        wb_3.set_cell_contents('Sheet1', 'A1', '=B1')
        wb_3.set_cell_contents('Sheet1', 'B1', '1')
        cell_a1 = wb_3.sheets['sheet1'].get_cell('A1')
        self.assertEqual(wb_3.detect_cycle(cell_a1), False)

        wb_3.set_cell_contents('Sheet1', 'B1', '=A1')
        self.assertEqual(wb_3.detect_cycle(cell_a1), True)

        wb_3.set_cell_contents('Sheet1', 'B1', '1')
        self.assertEqual(wb_3.detect_cycle(cell_a1), False)

        wb_3.set_cell_contents('Sheet1', 'B1', '=C1')
        wb_3.set_cell_contents('Sheet1', 'C1', '=A1')
        self.assertEqual(wb_3.detect_cycle(cell_a1), True)

        wb_3.set_cell_contents('Sheet1', 'A1', '=1')
        self.assertEqual(wb_3.detect_cycle(cell_a1), False)
    
    def test_self_ref(self):
        wb = sheets.Workbook()
        _, sheet_name = wb.new_sheet()
        num_cells = 1000
        create_chain(wb, sheet_name, num_cells, '12')
        
        wb.set_cell_contents(sheet_name, f'A{num_cells}', f'=A{num_cells}')
        for i in range(1, num_cells + 1):
            cell_value = wb.get_cell_value(sheet_name, f'A{i}')
            self.assertIsInstance(cell_value, sheets.CellError)
            self.assertEqual(cell_value.get_type(), sheets.CellErrorType.CIRCULAR_REFERENCE)

if __name__ == "__main__":
    cov = coverage.Coverage()
    cov.start()

    # unittest.main()
    unittest.main(module='test_cycles_i')

    cov.stop()
    cov.save()
    cov.html_report()