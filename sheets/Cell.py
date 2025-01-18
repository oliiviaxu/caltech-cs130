from typing import Any
import lark
from interpreter import FormulaEvaluator
import decimal 

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
            # formula evaluator
            pass
        elif contents.startswith("'"):
            self.value = contents[1:]                        
        else:
            if self.is_number(contents):
                self.value = decimal.Decimal(contents)
            else:
                assert False
        return self.value