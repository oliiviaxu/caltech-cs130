import os
import math
from .CellValue import CellValue

current_dir = os.path.dirname(os.path.abspath(__file__))
lark_path = os.path.join(current_dir, "formulas.lark")

class Cell:
    def __init__(self, location=str, contents=''):
        # self.sheet_name = sheet_name.lower()
        self.location = location.lower()
        self.contents = contents
        self.value = CellValue(None)
        self.tree = None
        self.parse_error = False
        self.in_cycle = False