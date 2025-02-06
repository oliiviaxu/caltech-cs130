# python -m unittest tests.performance.test_updates
import os
import unittest
import coverage
import decimal
import sheets
import cProfile
from pstats import Stats

from .testStructures import create_large_cycle, create_small_cycles, create_chain, create_web

current_dir = os.path.dirname(os.path.abspath(__file__))
dir = os.path.join(current_dir, 'cProfile_output/')

class CellUpdateTests(unittest.TestCase):

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
            
    def test_long_chain(self):
        wb = sheets.Workbook()
        _, sheet_name = wb.new_sheet()
        num_cells = 1000

        create_chain(wb, sheet_name, num_cells)
        # self.assertEqual(wb.get_cell_value(sheet_name, f'A{num_cells - 1}'), decimal.Decimal('1'))

        wb.set_cell_contents(sheet_name, f'A{num_cells}', '=0')

        for i in range(1, num_cells):
            self.assertEqual(wb.get_cell_value(sheet_name, f'A{i}'), decimal.Decimal('0'))
        
    def many_refs_helper(self, wb, sheet_name, num_cells, val):
        for i in range(1, num_cells):
            self.assertEqual(wb.get_cell_value(sheet_name, f'B{i}'), decimal.Decimal(int(val) + 1))

    def test_web(self):
        wb = sheets.Workbook()
        num_cells = 1000
        _, sheet_name = wb.new_sheet()

        create_web(wb, sheet_name, num_cells)

        # check that the 'B' cells are indeed A + 1
        self.many_refs_helper(wb, sheet_name, num_cells, '1')

        wb.set_cell_contents(sheet_name, 'A1', '0')
        self.many_refs_helper(wb, sheet_name, num_cells, '0')

    def test_rename_updates(self):
        # create a chain andthen rename the sheet
        wb = sheets.Workbook()
        num_cells = 1000
        _, sheet_name = wb.new_sheet()

        create_chain(wb, sheet_name, num_cells)

        wb.rename_sheet(sheet_name, 'SheetBla')
        for i in range(1, num_cells):
            outgoing_lst = wb.graph.outgoing_get('SheetBla', f'A{i}')
            self.assertEqual(outgoing_lst[0], ('sheetbla', f'a{i+1}'))