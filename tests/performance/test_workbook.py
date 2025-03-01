import os
import unittest
import coverage
import decimal
import sheets
import cProfile
from pstats import Stats
from .testStructures import create_chain, create_chain_2, create_web

current_dir = os.path.dirname(os.path.abspath(__file__))
dir = os.path.join(current_dir, 'cProfile_output/')

num_iterations = 100

class WorkbookTests(unittest.TestCase):
    @staticmethod
    def index_to_col(col_index):
        col = ""
        while col_index >= 0:
            col = chr(65 + col_index % 26) + col
            col_index = col_index // 26 - 1
        return col
    
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

    def test_load_wb(self):
        file_path = os.path.join(current_dir, 'fibonacci.json')
        num_loads = 1
        for _ in range(num_loads):
            with open(file_path, "r") as file:
                _ = sheets.Workbook.load_workbook(file)
    
    def test_rename(self):
        wb = sheets.Workbook()
        _, curr_sheet_name = wb.new_sheet('TestSheet')

        chain_size = 10
        create_chain(wb, curr_sheet_name, chain_size)

        num_renames = num_iterations
        for i in range(num_renames):
            new_sheet_name = f'Sheet{i}'
            wb.rename_sheet(curr_sheet_name, new_sheet_name)
            curr_sheet_name = new_sheet_name
            for j in range(1, chain_size):
                self.assertEqual(wb.get_cell_contents(curr_sheet_name, f'A{j}'), f'=A{j+1}')
    
    def test_rename_2(self):
        wb = sheets.Workbook()
        _, sn_1 = wb.new_sheet('ConstantSheet')
        _, sn_2 = wb.new_sheet('VariableSheet')

        chain_size = 10
        create_chain_2(wb, sn_1, sn_2, chain_size)

        num_renames = num_iterations
        for i in range(num_renames):
            new_sheet_name = f'Sheet{i}'
            wb.rename_sheet(sn_2, new_sheet_name)
            sn_2 = new_sheet_name
            for j in range(1, chain_size):
                self.assertEqual(wb.get_cell_contents(sn_1, f'A{j}'), f'={sn_2}!A{j}')
                self.assertEqual(wb.get_cell_contents(sn_2, f'A{j}'), f'={sn_1}!A{j + 1}')

    def test_copy_sheet(self):
        wb = sheets.Workbook()
        _, sn = wb.new_sheet('TestSheet')

        chain_size = 10
        create_chain(wb, sn, chain_size)

        num_copies = num_iterations
        for i in range(1, num_copies + 1):
            wb.copy_sheet(sn)
            new_sn = f'{sn}_{i}'
            for j in range(1, chain_size):
                self.assertEqual(wb.get_cell_contents(new_sn, f'A{j}'), f'=A{j+1}')
    
    def test_copy_sheet_2(self):
        wb = sheets.Workbook()
        _, sn_1 = wb.new_sheet('ConstantSheet')
        _, sn_2 = wb.new_sheet('TestSheet')

        chain_size = 10
        create_chain_2(wb, sn_1, sn_2, chain_size)

        num_copies = num_iterations
        for i in range(1, num_copies):
            wb.copy_sheet(sn_2)
            new_sn = f'{sn_2}_{i}'
            for j in range(1, chain_size):
                self.assertEqual(wb.get_cell_contents(new_sn, f'A{j}'), f'={sn_1}!A{j + 1}')

    def test_move_cells(self):
        wb = sheets.Workbook()
        _, sn = wb.new_sheet()

        web_size = 10
        create_web(wb, sn, web_size)

        num_moves = num_iterations
        for _ in range(num_moves):
            wb.move_cells(sn, 'A1', f'B{web_size}', 'C1')
            wb.move_cells(sn, 'C1', f'D{web_size}', 'A1')

    def test_copy_cells(self):
        wb = sheets.Workbook()
        _, sn = wb.new_sheet()

        web_size = 10
        create_web(wb, sn, web_size)

        num_moves = num_iterations
        for i in range(num_moves):
            column = WorkbookTests.index_to_col(i * 2 + 2)
            wb.copy_cells(sn, 'A1', f'B{web_size}', f'{column}1')