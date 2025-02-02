from .Cell import *
from typing import List, Optional, Tuple, Any
import re
from .interpreter import FormulaEvaluator

class Sheet:
    def __init__(self, sheet_name=''):
        self.sheet_name = sheet_name
        self.num_rows = 0
        self.num_cols = 0
        self.cells = []
    
    def get_cell(self, location):
        col_idx, row_idx = Sheet.split_cell_ref(location)
        return self.cells[row_idx][col_idx]

    def to_sheet_coords(col_index, row_index):
        """
        Converts 0-indexed coordinates (col_index, row_index) to spreadsheet coordinates.
        """
        col = ""
        while col_index >= 0:
            col = chr(65 + col_index % 26) + col
            col_index = col_index // 26 - 1

        row = str(row_index + 1)
        return col + row
    
    def str_to_index(column: str) -> int:
        column = column.lower()
        index = 0
        for i in range(len(column)):
            index += 26 ** i * (ord(column[len(column) - 1 - i]) - ord('a') + 1)
        return index - 1
    
    def split_cell_ref(location: str) -> Tuple[int, int]:
        i = 0
        while i < len(location) and location[i].isalpha():
            i += 1
        col, row = location[:i], location[i:]
        return Sheet.str_to_index(col), int(row) - 1
    
    def resize_sheet(self, new_num_rows, new_num_cols) -> None:
        # resizes sheet and updates num_rows and num_cols fields accordingly

        # Add rows
        for row_idx in range(self.num_rows, new_num_rows):
            row = []
            for col_idx in range(self.num_cols):
                location = Sheet.to_sheet_coords(col_idx, row_idx)
                cell = Cell(self.sheet_name, location, None)
                row.append(cell)
            self.cells.append(row)

        # Add columns
        for row_idx, row in enumerate(self.cells):
            for col_idx in range(self.num_cols, new_num_cols):
                location = Sheet.to_sheet_coords(col_idx, row_idx)
                cell = Cell(self.sheet_name, location, None)
                row.append(cell)

        self.num_rows = new_num_rows
        self.num_cols = new_num_cols
    
    def resize(self, location):
        col_idx, row_idx = Sheet.split_cell_ref(location)

        updated_num_rows = max(self.num_rows, row_idx + 1)
        updated_num_cols = max(self.num_cols, col_idx + 1)

        self.resize_sheet(updated_num_rows, updated_num_cols)
    
    def get_cell_contents(self, location: str) -> Optional[str]:
        self.resize(location)
        return self.get_cell(location).contents