# TODO: general performance tests, like loading a workbook

import os
import unittest
import coverage
import sheets
import cProfile
from pstats import Stats
from .testStructures import create_large_cycle, create_small_cycles, create_chain_2

current_dir = os.path.dirname(os.path.abspath(__file__))
dir = os.path.join(current_dir, 'cProfile_output/')

class GeneralPerformanceTests(unittest.TestCase):
    
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
    
    def rename_stress_test(self):
        wb = sheets.Workbook()
        _, sn_1 = wb.new_sheet()
        _, sn_2 = wb.new_sheet()
        num_cells = 1000

        create_chain_2(wb, sn_1, sn_2, num_cells, '12')

        for i in range(1, num_cells):
            self.assertEqual(wb.graph.outgoing_get(sn_1, f'A{i}')[0], (sn_2.lower(), f'a{i}'))

        wb.rename_sheet(sn_2, 'SheetBla')
        wb.set_cell_contents(sn_1, f'A{num_cells}', '0')

        for i in range(1, num_cells):
            self.assertEqual(wb.graph.outgoing_get(sn_1, f'A{i}')[0], ('sheetbla', f'a{i}'))
            self.assertEqual(wb.get_cell_value('sheetbla', f'A{i}'), 0)