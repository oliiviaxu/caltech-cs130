from lark import Lark, Transformer
import decimal
from .Cell import *
from .CellError import CellError, CellErrorType

class SheetNameExtractor(lark.visitors.Transformer):

    def __init__(self, old_name, new_name):
        self.old_name = old_name
        self.new_name = new_name
    
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
            if curr_name == self.old_name:
                # change this node 
                new_str = self.new_name + '!' + str(tree[1])
                return new_str
        else:
            assert False, 'Invalid formula. Format must be in ZZZZ9999.'