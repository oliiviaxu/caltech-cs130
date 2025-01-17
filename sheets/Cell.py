from typing import Any
import lark
import decimal 

class Cell:
    def __init__(self, contents=''):
        self.cell_type = None
        self.contents = contents
        self.value = ''
        self.outgoing = []
        self.ingoing = []
    
    def get_cell_value(self) -> Any:
        pass