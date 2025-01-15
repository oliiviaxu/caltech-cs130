from .Cell import *

class Sheet:
    def __init__(self, sheet_name=''):
        self.sheet_name = sheet_name
        self.cells = [] # TODO: fix later
        self.num_rows = 0
        self.num_cols = 0