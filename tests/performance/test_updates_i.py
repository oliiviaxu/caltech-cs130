# TODO:Alternately, you might also have updates where each cell is referenced by many other cells, 
# perhaps with much shallower chains, but still with large amounts of cell updates. 

import unittest
import sheets
import decimal
from typing import Optional
from .testStructures import create_web
import cProfile

class CellUpdateTests(unittest.TestCase):
            
    def test_long_chain(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        num_cells = 50
        CellUpdateTests.create_dependency_structure(wb, 'Sheet1', 1, num_cells)

        wb.set_cell_contents('sheet1', f'A{num_cells}', '20') # this value has to match that in line 37 TODO: fix
        for i in range(1, num_cells + 1):
            expected_value = 20 + (num_cells - i)
            self.assertEqual(wb.get_cell_value('sheet1', f'A{i}'), decimal.Decimal(expected_value))
        
    def many_refs_helper(self, wb, sheet_name, num_cells, val):
        wb.set_cell_contents(sheet_name, 'A1', val)
        for i in range(1, num_cells):
            self.assertEqual(wb.get_cell_value(sheet_name, f'B{i}'), decimal.Decimal(int(val) + 1))

    def test_web(self):
        wb = sheets.Workbook()
        num_cells = 50
        src_cell = 'A1'
        _, sheet_name = wb.new_sheet(sheet_name)
        create_web(sheet_name, num_cells)

        self.many_refs_helper(wb, sheet_name, num_cells, '12')

        self.many_refs_helper(wb, sheet_name, num_cells, f'{5000*5000}')
    
    def test_rename(self):
        pass