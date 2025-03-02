import math

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

    def strip_trailing_zeros(contents):
        if ('.' in contents):
            contents = contents.rstrip('0').rstrip('.')
        return contents