import os
import math

current_dir = os.path.dirname(os.path.abspath(__file__))
lark_path = os.path.join(current_dir, "formulas.lark")

class Cell:
    def __init__(self, sheet_name = '', location=str, contents=''):
        self.sheet_name = sheet_name.lower()
        self.location = location.lower()
        self.contents = contents
        self.value = None
        self.tree = None
        self.parse_error = False
    
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