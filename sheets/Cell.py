from typing import Any
import lark
from .interpreter import FormulaEvaluator
import decimal 
import os

# Get the directory containing Cell.py
current_dir = os.path.dirname(os.path.abspath(__file__))

# Construct the absolute path to formulas.lark
lark_path = os.path.join(current_dir, "formulas.lark")

class Cell:
    def __init__(self, contents=''):
        self.contents = contents
        self.value = ''
        self.outgoing = []
        self.ingoing = []
    
    def is_number(self, s):
        try:
            float(s)
            return True
        except ValueError:
            return False
    
    def get_cell_value(self) -> Any:
        contents = self.contents.strip() # remove whitespace
        if contents.startswith('='):
            parser = lark.Lark.open(lark_path, start='formula')
            ev = FormulaEvaluator()
            tree = parser.parse(self.contents)
            self.value = ev.visit(tree)
        elif contents.startswith("'"):
            self.value = contents[1:]                        
        else:
            if self.is_number(contents):
                self.value = decimal.Decimal(contents)
            else:
                assert False
        return self.value