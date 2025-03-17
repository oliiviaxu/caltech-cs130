from .Cell import Cell
from typing import Optional, Tuple

class Sheet:
    def __init__(self, sheet_name=''):
        self.sheet_name = sheet_name
        self.num_rows = 0
        self.num_cols = 0
        self.cells = []
    
    def get_cell(self, location):
        if (self.out_of_bounds(location)):
            return None
        col_idx, row_idx = Sheet.split_cell_ref(location)
        return self.cells[row_idx][col_idx]
    
    @staticmethod
    def index_to_col(col_index):
        col = ""
        while col_index >= 0:
            col = chr(65 + col_index % 26) + col
            col_index = col_index // 26 - 1
        return col

    def to_sheet_coords(col_index, row_index):
        """
        Converts 0-indexed coordinates (col_index, row_index) to spreadsheet coordinates.
        """
        col = Sheet.index_to_col(col_index)
        row = str(row_index + 1)
        return col + row
    
    def str_to_index(column: str) -> int:
        column = column.lower()
        index = 0
        for i in range(len(column)):
            index += 26 ** i * (ord(column[len(column) - 1 - i]) - ord('a') + 1)
        return index - 1

    def is_col_mixed_ref(location: str) -> bool:
        return location[0] == '$'

    def is_row_mixed_ref(location: str) -> bool:
        return '$' in location[1:]
    
    def split_cell_ref(location: str) -> Tuple[int, int]:
        i = 0
        # check if column has $
        col_start = 0
        if location[0] == '$':
            i += 1
            col_start = 1
        while i < len(location) and location[i].isalpha():
            i += 1
        col = location[col_start:i]
        # check if row has $
        if location[i] == '$':
            row = location[i + 1:]
        else:
            row = location[i:]
        return Sheet.str_to_index(col), int(row) - 1

    def out_of_bounds(self, location):
        col_idx, row_idx = Sheet.split_cell_ref(location)
        if (col_idx >= self.num_cols or row_idx >= self.num_rows):
            return True
        return False
    
    def resize_sheet(self, new_num_rows, new_num_cols) -> None:
        # resizes sheet and updates num_rows and num_cols fields accordingly

        # Add rows
        for row_idx in range(self.num_rows, new_num_rows):
            row = []
            for col_idx in range(self.num_cols):
                location = Sheet.to_sheet_coords(col_idx, row_idx)
                cell = Cell(location, None)
                row.append(cell)
            self.cells.append(row)

        # Add columns
        for row_idx, row in enumerate(self.cells):
            for col_idx in range(self.num_cols, new_num_cols):
                location = Sheet.to_sheet_coords(col_idx, row_idx)
                cell = Cell(location, None)
                row.append(cell)

        self.num_rows = new_num_rows
        self.num_cols = new_num_cols
    
    def resize(self, location):
        col_idx, row_idx = Sheet.split_cell_ref(location)

        updated_num_rows = max(self.num_rows, row_idx + 1)
        updated_num_cols = max(self.num_cols, col_idx + 1)

        self.resize_sheet(updated_num_rows, updated_num_cols)
    
    def get_cell_contents(self, location: str) -> Optional[str]:
        if (self.out_of_bounds(location)):
            return None
        return self.get_cell(location).contents
    
    def empty_row(self, row_idx):
        for col_idx in range(self.num_cols):
            if (self.cells[row_idx][col_idx].contents is not None):
                return False
        return True
    
    def empty_col(self, col_idx):
        for row_idx in range(self.num_rows):
            if (self.cells[row_idx][col_idx].contents is not None):
                return False
        return True
    
    def delete_row(self):
        self.cells = self.cells[:-1]

    def delete_col(self):
        for row_idx in range(self.num_rows):
            self.cells[row_idx] = self.cells[row_idx][:-1]
    
    def check_shrink(self, location):
        col_idx, row_idx = Sheet.split_cell_ref(location)
        if (col_idx == self.num_cols - 1):
            # try to reduce columns
            while self.num_cols > 0:
                if (self.empty_col(self.num_cols - 1)):
                    self.num_cols -= 1
                    self.delete_col()
                else:
                    break
        if (row_idx == self.num_rows - 1):
            # try to reduce rows
            while self.num_rows > 0:
                if (self.empty_row(self.num_rows - 1)):
                    self.num_rows -= 1
                    self.delete_row()
                else:
                    break