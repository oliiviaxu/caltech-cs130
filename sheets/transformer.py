from lark import Lark, Transformer
import decimal
from .Cell import *
from .CellError import CellError, CellErrorType
import re

class SheetNameExtractor(lark.visitors.Transformer):
    def sheet_name_needs_quotes(sheet_name):
        pattern = r"^[A-Za-z_][A-Za-z0-9_]*$"
        return not bool(re.fullmatch(pattern, sheet_name))

    def __init__(self, sheet_name, new_sheet_name):
        self.sheet_name = sheet_name
        self.new_sheet_name = new_sheet_name
    
    def mul_expr(self, tree):
        return str(tree[0]) + ' ' + str(tree[1]) + ' ' + str(tree[2]) 

    def add_expr(self, tree):
        return str(tree[0]) + ' ' + str(tree[1]) +  ' ' + str(tree[2]) 
    
    def unary_op(self, tree):
        return str(tree[0]) + str(tree[1])

    def concat_expr(self, tree):
        return str(tree[0]) + ' & ' + str(tree[1])

    def error(self, tree):
        return str(tree[0])

    def parens(self, tree):
        return '(' + str(tree[0]) + ')'
    
    def number(self, tree):
        return str(tree[0])

    def string(self, tree):
        return str(tree[0])
    
    def cell(self, tree):
        # processes a parse tree node 
        if len(tree) == 1:
            return str(tree[0])
        if len(tree) == 2:
            curr_name = str(tree[0])
            # check if has quotes.
            # if it has quotes and doesn't need, then remove
            # if it has quotes and needs, then keep
            if (len(curr_name) > 2 and curr_name[0] == '\'' and curr_name[-1] == '\''):
                if (not SheetNameExtractor.sheet_name_needs_quotes(curr_name[1:-1])):
                    curr_name = curr_name[1:-1]
                    if curr_name == self.sheet_name:
                        curr_name = self.new_sheet_name
                        if SheetNameExtractor.sheet_name_needs_quotes(curr_name):
                            curr_name = "'" + curr_name + "'"
                else:
                    # it needs quotes
                    if curr_name[1:-1] == self.sheet_name:
                        curr_name = self.new_sheet_name
                        if SheetNameExtractor.sheet_name_needs_quotes(curr_name):
                            curr_name = "'" + curr_name + "'"
            else:
                if curr_name == self.sheet_name:
                    curr_name = self.new_sheet_name
                    if SheetNameExtractor.sheet_name_needs_quotes(curr_name):
                        curr_name = "'" + curr_name + "'"
            return curr_name + '!' + str(tree[1])
        else:
            assert False, 'Invalid formula. Format must be in ZZZZ9999.'