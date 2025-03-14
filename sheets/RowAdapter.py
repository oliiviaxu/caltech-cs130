import decimal
from functools import total_ordering
from .Cell import Cell
from .CellError import CellError, CellErrorType
from .CellValue import CellValue

@total_ordering
class RowAdapter:
    def __init__(self, row_idx, row_data, sort_cols):
        self.row_idx = row_idx
        self.row_data = row_data
        self.sort_cols = sort_cols

    def get_sort_key(self):
        key = []
        for col in self.sort_cols:
            index = abs(col) - 1

            cell = self.row_data[index]
            value = self._get_cell_sort_value(cell)
            
            reverse_flag = col < 0
            key.append((value, reverse_flag))

        return tuple(key)

    def _get_cell_sort_value(self, cell):
        
        if cell is None or cell.value is None:
            return (2, None) 
        
        value = cell.value.val

        if isinstance(value, CellError):
            return (1, value.get_type().value)
        else:
            return (0, value)

    def __eq__(self, other):
        return self.get_sort_key() == other.get_sort_key()

    def __lt__(self, other):
        self_key = self.get_sort_key()
        other_key = other.get_sort_key()

        for (self_value, self_reverse), (other_value, other_reverse) in zip(self_key, other_key):
            if self_value == other_value:
                continue

            # Sorting direction handling
            if self_reverse or other_reverse:
                return self_value > other_value
            else:
                return self_value < other_value
        
        return False