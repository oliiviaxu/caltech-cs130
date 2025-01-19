from typing import Any
import lark
import decimal 
import os

current_dir = os.path.dirname(os.path.abspath(__file__))
lark_path = os.path.join(current_dir, "formulas.lark")

class Cell:
    def __init__(self, contents=''):
        self.contents = contents
        self.value = ''
        self.outgoing = []
        self.ingoing = []
    
    def is_number(s):
        try:
            float(s)
            return True
        except ValueError:
            return False
    
    def get_cell_value(self, ev) -> Any:
        contents = self.contents
        if (contents is None):
            return ""
        
        contents = contents.strip()
        if contents.startswith('='):
            parser = lark.Lark.open(lark_path, start='formula')
            tree = parser.parse(self.contents)
            self.value = ev.visit(tree)
        elif contents.startswith("'"):
            self.value = contents[1:]                        
        else:
            if Cell.is_number(contents):
                self.value = decimal.Decimal(contents)
            else:
                self.value = contents
        return self.value