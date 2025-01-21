import decimal
from .CellError import CellError, CellErrorType
import lark
from typing import Tuple
from .Cell import *
from lark.visitors import visit_children_decor

class FormulaEvaluator(lark.visitors.Interpreter):
    def __init__(self, sheet_name, ref_info):
        self.sheet_name = sheet_name
        self.ref_info = ref_info

    error_dict = {
        "#ERROR!": 1,
        "#CIRCREF!": 2,
        "#REF!": 3,
        "#NAME?": 4,
        "#VALUE!": 5,
        "#DIV/0!": 6
    }
    
    def change_type(val_1, val_2) -> Any:
        # helper function for implicit type conversion needed in add_expr and mul_expr

        if val_1 is None or type(val_1) == CellError:
            raise TypeError(f'Invalid type: {val_1}')
        if val_2 is None or type(val_2) == CellError:
            raise TypeError(f'Invalid type: {val_2}')
        if type(val_1) == str:
            if Cell.is_number(val_1):
                val_1 = decimal.Decimal(val_1)
            else:
                raise TypeError(f'Invalid type: {val_1}')
        if type(val_2) == str:
            if Cell.is_number(val_2):
                val_2 = decimal.Decimal(val_2)
            else:
                raise TypeError(f'Invalid type: {val_2}')
        
        return val_1, val_2
        
    @visit_children_decor
    def add_expr(self, values):
        orig_val_1, operator, orig_val_2 = values[0], values[1], values[2]

        try:
            val_1, val_2 = FormulaEvaluator.change_type(orig_val_1, orig_val_2)
        except TypeError:
            return CellError(CellErrorType.TYPE_ERROR, 'Cannot perform addition/subtraction. Invalid types')
        
        if operator == '+':
            return val_1 + val_2
        elif operator == '-':
            return val_1 - val_2
        else:
            assert False, f'Unexpected operation: {values[1]}'
    
    @visit_children_decor
    def mul_expr(self, values):

        orig_val_1, operator, orig_val_2 = values[0], values[1], values[2]

        try:
            val_1, val_2 = FormulaEvaluator.change_type(orig_val_1, orig_val_2)
        except TypeError:
            return CellError(CellErrorType.TYPE_ERROR, 'Cannot perform multiplication/division. Invalid types')

        if operator == '*':
            return val_1 * val_2
        elif operator == '/':
            if (val_2 == 0):
                return CellError(CellErrorType.DIVIDE_BY_ZERO, 'Cannot divide by zero')
            return val_1 / val_2
        else:
            assert False, f'Unexpected operation: {operator}'
    
    @visit_children_decor
    def unary_op(self, values):
        operator, val_1 = values[0], values[1]

        if val_1 is None or type(val_1) == CellError:
            return val_1
        if type(val_1) == str:
            if Cell.is_number(val_1):
                val_1 = decimal.Decimal(val_1)
            else:
                return CellError(CellErrorType.TYPE_ERROR, 'Invalid type.')
        
        if operator == '+':
            return abs(val_1)
        elif operator == '-':
            return -abs(val_1)
        else:
            assert False, f'Unexpected operation: {operator}'

    @visit_children_decor
    def concat_expr(self, values):
        assert len(values) == 2, 'Unexpected number of args'

        val_1, val_2 = values[0], values[1]
        if val_1 is None or type(val_1) == CellError:
            return val_1
        elif val_2 is None or type(val_2) == CellError:
            return val_2
        else:
            return str(val_1) + str(val_2)

    def error(self, tree):
        error_value = FormulaEvaluator.error_dict[self.children[0]] # TODO: fix
        return error_value

    def parens(self, tree):
        values = self.visit_children(tree)
        assert len(values) == 1, f'Unexpected tree {tree.pretty()}'
        return values[0]
    
    def number(self, tree):
        # called when run into number node 
        return decimal.Decimal(tree.children[0])

    def string(self, tree):
        # called when run into a string node
        return tree.children[0].value[1:-1]
    
    def cell(self, tree):
        # first parse the value into sheet (if given) and location
        reference = tree.children[0].value.upper()
        assert reference in self.ref_info, f'Could not find cell information for reference {reference}'
        return self.ref_info[reference]