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
        # establish heirarchy
        
        if cell is None or cell.value is None:
            return (5, None)

        value = cell.value.val
        
        if isinstance(value, bool):
            return (1, value)
        
        if isinstance(value, str):
            return (2, value.lower())
        
        if isinstance(value, (int, float, decimal.Decimal)):
            return (3, decimal.Decimal(value))
        
        if isinstance(value, CellError):
            return (4, value.get_type().value)

        return (6, str(value))

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