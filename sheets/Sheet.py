from .Cell import *
from typing import List, Optional, Tuple, Any
import re
from .interpreter import FormulaEvaluator

class Sheet:
    def __init__(self, sheet_name=''):
        self.sheet_name = sheet_name
        self.num_rows = 25
        self.num_cols = 25
        self.cells = []
        for i in range(self.num_rows):
            row = []
            for j in range(self.num_cols):
                location = Sheet.to_sheet_coords(j, i)
                cell = Cell(sheet_name, location, None)
                row.append(cell)
            self.cells.append(row)
        # self.ev = FormulaEvaluator(sheet_name)

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
        column = column.upper()
        index = 0
        for i in range(len(column)):
            index += 26 ** i * (ord(column[len(column) - 1 - i]) - ord('A') + 1)
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

    def set_cell_contents(self, sheet_name, location, contents, outgoing) -> None:

        col_idx, row_idx = Sheet.split_cell_ref(location)

        updated_num_rows = max(self.num_rows, row_idx + 1)
        updated_num_cols = max(self.num_cols, col_idx + 1)

        self.resize_sheet(updated_num_rows, updated_num_cols)

        cell = self.cells[row_idx][col_idx]

        for referenced_cell in outgoing:
            referenced_cell.ingoing.append(cell)   

        cell.outgoing = outgoing
        cell.contents = contents
    
    def get_cell_contents(self, location: str) -> Optional[str]:

        # if specified location is beyond extent of sheet, raises a ValueError 

        col_idx, row_idx = Sheet.split_cell_ref(location)

        if col_idx >= self.num_cols or row_idx >= self.num_rows:
            raise ValueError('Location is beyond current extent of sheet.')
        
        return self.cells[row_idx][col_idx].contents