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

            cell = self.row_data[index] if index < len(self.row_data) else None

            val_1, val_2 = self._get_cell_sort_value(cell)

            if col < 0:  # Handle descending order
                if isinstance(val_2, (decimal.Decimal)):
                    val_2 = -val_2
                elif isinstance(val_2, str):
                    val_2 = "".join(chr(255 - ord(c)) for c in val_2) # reverse lexicographically

            key.append((val_1, val_2))

        return tuple(key)

    def _get_cell_sort_value(self, cell):
        """
        1. Blank
        2. Cell error
        3. Numeric 
        4. Text
        5. Bool
        6. Other
        """
        if cell is None or cell.value is None:
            return (0, decimal.Decimal("-Infinity"))

        value = cell.value.val

        if value is None:
            return (0, decimal.Decimal("-Infinity"))  # Blank cells sort first

        if isinstance(value, CellError):
            return (1, value.get_type().value)

        if isinstance(value, decimal.Decimal):
            return (2, decimal.Decimal(value))
        
        if isinstance(value, str):
            return (3, value.lower())

        if isinstance(value, bool):
            return (4, int(value))

        return (5, str(value))

    def __eq__(self, other):
        return self.get_sort_key() == other.get_sort_key()

    def __lt__(self, other):
        self_key = self.get_sort_key()
        other_key = other.get_sort_key()
        return self_key < other_key