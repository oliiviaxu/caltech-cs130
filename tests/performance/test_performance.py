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

num_iterations = 10

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
    
    # gemini did this
    def test_fibonacci(self):
        wb = sheets.Workbook()
        wb.new_sheet()  # Default sheet name is "Sheet1"

        # Set up Fibonacci sequence (up to A1000)
        wb.set_cell_contents("Sheet1", "A1", "1")  # Set A1 initially
        wb.set_cell_contents("Sheet1", "A2", "1")
        for i in range(3, 11):
            wb.set_cell_contents("Sheet1", f"A{i}", f"=A{i-1}+A{i-2}")

        # Trigger updates by setting A1 last
        wb.set_cell_contents("Sheet1", "A1", "1")

        # self.assertEqual(wb.get_cell_value("Sheet1", "A10"), 55)
        # self.assertEqual(wb.get_cell_value("Sheet1", "A20"), 6765)
        # self.assertEqual(wb.get_cell_value("Sheet1", "A30"), 832040)

        # Assertions to check Fibonacci values up to A10
        self.assertEqual(wb.get_cell_value("Sheet1", "A1"), 1)
        self.assertEqual(wb.get_cell_value("Sheet1", "A2"), 1)
        self.assertEqual(wb.get_cell_value("Sheet1", "A3"), 2)
        self.assertEqual(wb.get_cell_value("Sheet1", "A4"), 3)
        self.assertEqual(wb.get_cell_value("Sheet1", "A5"), 5)
        self.assertEqual(wb.get_cell_value("Sheet1", "A6"), 8)
        self.assertEqual(wb.get_cell_value("Sheet1", "A7"), 13)
        self.assertEqual(wb.get_cell_value("Sheet1", "A8"), 21)
        self.assertEqual(wb.get_cell_value("Sheet1", "A9"), 34)
        self.assertEqual(wb.get_cell_value("Sheet1", "A10"), 55)

    def test_pascals_triangle(self):
        wb = sheets.Workbook()
        wb.new_sheet()  # Default sheet name is "Sheet1"

        # Set up Pascal's Triangle
        wb.set_cell_contents("Sheet1", "A1", "1")
        for i in range(2, 11):  # Create up to 10 rows
            wb.set_cell_contents("Sheet1", f"A{i}", "1")
            wb.set_cell_contents("Sheet1", f"{chr(ord('A') + i - 1)}{i}", "1")
            for j in range(2, i):
                wb.set_cell_contents("Sheet1", f"{chr(ord('A') + j - 1)}{i}",
                                     f"={chr(ord('A') + j - 2)}{i - 1}+{chr(ord('A') + j - 1)}{i - 1}")

        # Calculate and verify row sums
        for i in range(1, 11):
            row_sum_formula = "+".join([f"{chr(ord('A') + j - 1)}{i}" for j in range(1, i + 1)])
            wb.set_cell_contents("Sheet1", f"B{i}", f"={row_sum_formula}")
            expected_sum = 2**(i - 1)
            self.assertEqual(wb.get_cell_value("Sheet1", f"B{i}"), expected_sum)