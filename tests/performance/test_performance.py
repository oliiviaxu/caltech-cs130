import os
import unittest
import coverage
import decimal
import sheets
import cProfile
from pstats import Stats
from .testStructures import create_large_cycle, create_small_cycles, create_chain_2

current_dir = os.path.dirname(os.path.abspath(__file__))
dir = os.path.join(current_dir, 'cProfile_output/')

num_iterations = 100

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
    
    def test_rename(self):
        wb = sheets.Workbook()
        _, sn_1 = wb.new_sheet()
        _, sn_2 = wb.new_sheet()
        num_cells = num_iterations

        create_chain_2(wb, sn_1, sn_2, num_cells)

        # for i in range(1, num_cells):
        #     self.assertEqual(wb.graph.outgoing_get(sn_1, f'A{i}')[0], (sn_2.lower(), f'a{i}'))

        wb.rename_sheet(sn_2, 'SheetBla')
        for i in range(1, num_cells):
            self.assertEqual(wb.graph.outgoing_get(sn_1, f'A{i}')[0], ('sheetbla', f'a{i}'))

        wb.set_cell_contents(sn_1, f'A{num_cells}', '0')
        for i in range(1, num_cells):
            self.assertEqual(wb.get_cell_value(sn_1, f'A{i}'), 0)
            self.assertEqual(wb.get_cell_value('sheetbla', f'A{i}'), 0)
    
    def test_fibonacci(self):
        wb = sheets.Workbook()
        wb.new_sheet()

        wb.set_cell_contents("Sheet1", "A1", "1")
        wb.set_cell_contents("Sheet1", "A2", "1")

        num_numbers = 1000
        for i in range(3, num_numbers):
            wb.set_cell_contents("Sheet1", f"A{i}", f"=A{i-1}+A{i-2}")

        a = 2
        b = 1
        wb.set_cell_contents('Sheet1', 'A1', str(a))
        for i in range(3, num_numbers):
            self.assertEqual(wb.get_cell_value('Sheet1', f'A{i}'), decimal.Decimal(a + b))
            a, b = b, (a + b)
    
    @staticmethod
    def index_to_col(col_index):
        col = ""
        while col_index >= 0:
            col = chr(65 + col_index % 26) + col
            col_index = col_index // 26 - 1
        return col

    def test_pascals_triangle(self):
        wb = sheets.Workbook()
        wb.new_sheet()

        num_rows = 50
        wb.set_cell_contents("Sheet1", "A1", "1")
        for i in range(1, num_rows):
            wb.set_cell_contents("Sheet1", f"A{i + 1}", "1")
            wb.set_cell_contents("Sheet1", f"{GeneralPerformanceTests.index_to_col(i)}{i + 1}", "1")
            for j in range(1, i):
                wb.set_cell_contents("Sheet1", f"{GeneralPerformanceTests.index_to_col(j)}{i + 1}",
                                     f"={GeneralPerformanceTests.index_to_col(j - 1)}{i}+{GeneralPerformanceTests.index_to_col(j)}{i}")

        # Calculate and verify row sums
        for i in range(num_rows):
            sum = 0
            for j in range(i + 1):
                sum += wb.get_cell_value('Sheet1', f'{GeneralPerformanceTests.index_to_col(j)}{i + 1}')
            self.assertEqual(sum, 2 ** i)
    
    def test_copy_cells(self):
        # test copying an entire sheet to a new sheet
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.new_sheet()

        num_cells = 1000
        last_cell = GeneralPerformanceTests.index_to_col(num_cells - 1) + str(num_cells - 1)

        for i in range(1, num_cells):
            wb.set_cell_contents('Sheet1', f"A{i}", "1")
                
        wb.move_cells('Sheet1', 'A1', last_cell, 'A1', 'Sheet2')
        
        for i in range(1, num_cells):
            self.assertEqual(wb.get_cell_value('Sheet1', f"A{i}"), None)
            self.assertEqual(wb.get_cell_value('Sheet2', f"A{i}"), "1")

        self.assertEqual(wb.get_cell_value('Sheet2', last_cell), "1")
    
    def test_copy_cells(self):
        # test copying an entire sheet to a new sheet
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.new_sheet()

        num_cells = 1000
        last_cell = GeneralPerformanceTests.index_to_col(num_cells - 1) + str(num_cells - 1)

        for i in range(1, num_cells):
            wb.set_cell_contents('Sheet1', f"A{i}", "1")
        
        wb.copy_cells('Sheet1', 'A1', last_cell, 'A1', 'Sheet2')
        
        for i in range(1, num_cells):
            self.assertEqual(wb.get_cell_value('Sheet1', f"A{i}"), "1")
            self.assertEqual(wb.get_cell_value('Sheet2', f"A{i}"), "1")

        self.assertEqual(wb.get_cell_value('Sheet2', last_cell), "1")