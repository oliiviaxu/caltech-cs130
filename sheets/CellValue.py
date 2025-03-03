import math
import decimal
from .CellError import CellError, CellErrorType

class CellValue:
    def __init__(self, val):
        # literal, number, string, formula, boolean, CellError, None
        self.val = val

    @staticmethod
    def is_number(s):
        try:
            float(s)

            if (float(s) == float("inf") or float(s) == float("-inf")):
                return False
            if (math.isnan(float(s))):
                return False

            return True
        except ValueError:
            return False

    @staticmethod
    def strip_trailing_zeros(contents):
        if ('.' in contents):
            contents = contents.rstrip('0').rstrip('.')
        return contents
    
    def to_string(self):
        if self.val is None:
            self.val = ''
        elif self.is_cell_error():
            return
        elif isinstance(self.val, bool):
            if self.val == True:
                self.val = 'TRUE'
            else:
                self.val = 'FALSE'
        elif CellValue.is_number(self.val):
            self.val = CellValue.strip_trailing_zeros(str(self.val))
        else:
            self.val = str(self.val)
            
    def to_number(self):
        if self.val is None:
            self.val = decimal.Decimal('0')
        elif self.is_cell_error():
            return
        else:
            if isinstance(self.val, bool):
                self.val = float(self.val)
            self.val = str(self.val)
            if CellValue.is_number(self.val):
                self.val = decimal.Decimal(CellValue.strip_trailing_zeros(self.val))
            else:
                self.val = CellError(CellErrorType.TYPE_ERROR, f'Invalid type for {self.val}')

    def is_cell_error(self):
        if isinstance(self.val, CellError):
            return True
        return False