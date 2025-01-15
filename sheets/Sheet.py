from .Cell import *
from typing import List, Optional, Tuple, Any
import re

class Sheet:
    def __init__(self, sheet_name=''):
        self.sheet_name = sheet_name
        self.num_rows = 1
        self.num_cols = 0
        self.cells = [[]]
    
    def str_to_index(self, column: str) -> int:
        column = column.upper()
        index = 0
        for i in range(len(column)):
            index += 26 ** i * (ord(column[len(column) - 1 - i]) - ord('A') + 1)
        return index - 1
    
    def split_cell_ref(self, location: str) -> Tuple[int, int]:
        i = 0
        while i < len(location) and location[i].isalpha():
            i += 1
        col, row = location[:i], location[i:]
        return self.str_to_index(col), int(row) - 1
    
    def resize_sheet(self, new_num_rows, new_num_cols) -> None:
        # resizes sheet and updates num_rows and num_cols fields accordingly

        # Add rows
        for _ in range(new_num_rows - self.num_rows):
            self.cells.append([None] * self.num_cols)

        # Add columns
        for row in self.cells:
            row.extend([None] * (new_num_cols - self.num_cols))

        self.num_rows = new_num_rows
        self.num_cols = new_num_cols

    def set_cell_contents(self, location: str, contents: Optional[str]) -> None:

        col_idx, row_idx = self.split_cell_ref(location)

        updated_num_rows = max(self.num_rows, row_idx + 1)
        updated_num_cols = max(self.num_cols, col_idx + 1)

        self.resize_sheet(updated_num_rows, updated_num_cols)

        self.cells[row_idx][col_idx] = Cell(contents) # TODO: maybe change this
    
    def get_cell_contents(self, location: str) -> Optional[str]:

        # if specified location is beyond extent of sheet, raises a ValueError 

        col_idx, row_idx = self.split_cell_ref(location)

        if col_idx >= self.num_cols or row_idx >= self.num_rows:
            raise ValueError('Location is beyond current extent of sheet.')
        
        return self.cells[row_idx][col_idx].contents