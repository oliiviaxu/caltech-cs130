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
        "#error!": CellErrorType.PARSE_ERROR,
        "#circref!": CellErrorType.CIRCULAR_REFERENCE,
        "#ref!": CellErrorType.BAD_REFERENCE,
        "#name?": CellErrorType.BAD_NAME,
        "#value!": CellErrorType.TYPE_ERROR,
        "#div/0!": CellErrorType.DIVIDE_BY_ZERO
    }
    
    def change_type(val_1, val_2) -> Any:
        # helper function for implicit type conversion needed in add_expr and mul_expr
        if val_1 is None:
            val_1 = 0
        if val_2 is None:
            val_2 = 0

        if isinstance(val_1, CellError) or isinstance(val_2, CellError):
            return val_1, val_2

        def convert(val):
            if isinstance(val, str):
                if Cell.is_number(val):
                    return decimal.Decimal(val)
                else:
                    return CellError(CellErrorType.TYPE_ERROR, f'Invalid type for {val}')
            return val

        res_1 = convert(val_1)
        res_2 = convert(val_2)

        return res_1, res_2
        
    @visit_children_decor
    def add_expr(self, values):
            val_1, operator, val_2 = values[0], values[1], values[2]

            res_1, res_2 = FormulaEvaluator.change_type(val_1, val_2)
            if isinstance(res_1, CellError):
                return res_1
            elif isinstance(res_2, CellError):
                return res_2
            else:
                if operator == '+':
                    return res_1 + res_2
                elif operator == '-':
                    return res_1 - res_2
                else:
                    assert False, f'Unexpected operation: {operator}'
        
    @visit_children_decor
    def mul_expr(self, values):

        val_1, operator, val_2 = values[0], values[1], values[2]
        res_1, res_2 = FormulaEvaluator.change_type(val_1, val_2)
        
        if isinstance(res_1, CellError):
            return res_1
        elif isinstance(res_2, CellError):
            return res_2
        else:
            if operator == '*':
                return res_1 * res_2
            elif operator == '/':
                if (res_2 == 0):
                    return CellError(CellErrorType.DIVIDE_BY_ZERO, 'Cannot divide by zero')
                return res_1 / res_2
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
        error_value = CellError(FormulaEvaluator.error_dict[tree.children[0].lower()], 'String representation of error given')
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
        if (len(tree.children) == 1):
            reference = tree.children[0].value.lower()
            if reference not in self.ref_info:
                return CellError(CellErrorType.BAD_REFERENCE, f'Could not find cell information for reference {reference}')
            return self.ref_info[reference]
        elif (len(tree.children) == 2):
            sheet = tree.children[0].value.lower()
            location = tree.children[1].value.lower()
            reference = sheet + '!' + location
            if reference not in self.ref_info:
                return CellError(CellErrorType.BAD_REFERENCE, f'Could not find cell information for reference {reference}')
            return self.ref_info[reference]
        else:
            assert False, 'Length of tree for cell is not one or two'