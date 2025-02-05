import unittest
import enum
import sheets
import decimal
import random
from typing import Optional, List, Tuple

class StructureType(enum.Enum):
    """Enum representing different types of dependency structures."""
    CHAIN = 1
    WEB = 2
    # POLYGON = 3
    # TODO: add more structure types

class CellUpdateTests(unittest.TestCase):

    @staticmethod
    def create_dependency_structure(wb, sheet_name, structure_type: int,
                                    num_cells: Optional[int] = None):
        if structure_type == StructureType.CHAIN.value:
            cell_val = 12
            for i in range(1, num_cells):
                wb.set_cell_contents(sheet_name, f'A{i}', f'=A{i+1} + {1}')
            # Set the last cell to a number
            wb.set_cell_contents(sheet_name, f'A{num_cells}', str(cell_val))
        
        elif structure_type == StructureType.WEB.value:
            # many cells depend on this cell
            src_cell, src_val = 'A1', '12'
            wb.set_cell_contents(sheet_name, src_cell, src_val)

            for i in range(1, num_cells):
                wb.set_cell_contents(sheet_name, f'{chr(65 + (i % 26))}1', f'={src_cell} + 1')

    def test_long_chain(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        num_cells = 50
        CellUpdateTests.create_dependency_structure(wb, 'Sheet1', 1, num_cells)

        wb.set_cell_contents('sheet1', f'A{num_cells}', '20') # this value has to match that in line 37
        print(wb.sheets)
        for i in range(1, num_cells + 1):
            expected_value = 20 + (num_cells - i)
            self.assertEqual(wb.get_cell_value('sheet1', f'A{i}'), decimal.Decimal(expected_value))

    def test_many_refs(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        num_cells = 50
        CellUpdateTests.create_dependency_structure(wb, 'Sheet1', 2, num_cells)

        wb.set_cell_contents('Sheet1', 'A1', '=1')
        for i in range(1, num_cells):
            cell_val = wb.get_cell_value('Sheet1', f'{chr(65 + (i % 26))}1')
            self.assertEqual(decimal.Decimal('2'), cell_val)
            
        # TODO:Alternately, you might also have updates where each cell is referenced by many other cells, 
        # perhaps with much shallower chains, but still with large amounts of cell updates. 

# TODO: Can you come up with a general way of constructing such tests, and use it to exercise 
# several different scenarios?

# TODO: How do you trigger many updates with only one cell-change, so that you are maximally 
# exercising code inside your library, rather than code in the test?