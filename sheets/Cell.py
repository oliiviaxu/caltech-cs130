from typing import Any
import lark
import decimal 
import os

current_dir = os.path.dirname(os.path.abspath(__file__))
lark_path = os.path.join(current_dir, "formulas.lark")

class Cell:
    def __init__(self, sheet_name = '', location=str, contents=''):
        self.sheet_name = sheet_name
        self.location = location
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