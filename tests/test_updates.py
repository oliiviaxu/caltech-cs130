import unittest
import enum
import sheets
import random
from typing import Optional, List, Tuple

class StructureType(enum.Enum):
    """Enum representing different types of dependency structures."""
    CHAIN = 1
    WEB = 2
    POLYGON = 3
    # TODO: add more structure types

class CellUpdateTests(unittest.TestCase):

    @staticmethod
    def create_dependency_structure(wb, structure_type: int, num_cells: Optional[int] = None):
        if structure_type == StructureType.CHAIN:
            wb.new_sheet()
            for i in range(1, num_cells):
                rand = random.randint(0, 10)
                wb.set_cell_contents('Sheet1', f'A{i}', f'=A{i+1} + {rand}')
            wb.set_cell_contents('Sheet1', f'A{num_cells}', '10')  # Set the last cell to a number
        
        # elif structure_type == StructureType.WEB:
        #     pass

    def test_long_chain(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        chain_length = 50
        CellUpdateTests.create_dependency_structure(wb, 1, chain_length)
        print(wb.get_cell('sheet1', 'A50'))

        wb.set_cell_contents('sheet1', f'A{chain_length}', '20')  # Change the last cell
        print(wb.sheets)
        for i in range(1, chain_length + 1):
            expected_value = 20 + (chain_length - i)
            self.assertEqual(wb.get_cell_value('sheet1', f'A{i}'), expected_value)

    def test_many_refs(self):
        wb = sheets.Workbook()
        num_cells = 10
        CellUpdateTests.create_dependency_structure(wb, 2, num_cells)
        # TODO:Alternately, you might also have updates where each cell is referenced by many other cells, 
        # perhaps with much shallower chains, but still with large amounts of cell updates. 

# TODO: Can you come up with a general way of constructing such tests, and use it to exercise 
# several different scenarios?

# TODO: How do you trigger many updates with only one cell-change, so that you are maximally 
# exercising code inside your library, rather than code in the test?