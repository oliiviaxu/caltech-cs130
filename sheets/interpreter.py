import decimal
from .CellError import CellError, CellErrorType
from .CellValue import CellValue
import lark
from .Cell import Cell 
from lark.visitors import visit_children_decor
from typing import Any

decimal.getcontext().prec = 500

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
            val_1 = decimal.Decimal('0')
        if val_2 is None:
            val_2 = decimal.Decimal('0')

        if isinstance(val_1, CellError) or isinstance(val_2, CellError):
            return val_1, val_2

        def convert(val):
            if isinstance(val, str):
                if CellValue.is_number(val):
                    return decimal.Decimal(CellValue.strip_trailing_zeros(val))
                else:
                    return CellError(CellErrorType.TYPE_ERROR, f'Invalid type for {val}')
            elif isinstance(val, decimal.Decimal):
                return decimal.Decimal(CellValue.strip_trailing_zeros(str(val)))
            return val

        res_1 = convert(val_1)
        res_2 = convert(val_2)

        return res_1, res_2
    
    def change_type_concat(val_1, val_2):
        if val_1 is None:
            val_1 = ''
        if val_2 is None:
            val_2 = ''
        
        if CellValue.is_number(val_1):
            val_1 = CellValue.strip_trailing_zeros(str(val_1))
        if CellValue.is_number(val_2):
            val_2 = CellValue.strip_trailing_zeros(str(val_2))
        return val_1, val_2
        
    @visit_children_decor
    def add_expr(self, values):
        val_1, operator, val_2 = values[0], values[1], values[2]
        val_1.to_number()
        val_2.to_number()

        if val_1.is_cell_error():
            return val_1
        elif val_2.is_cell_error():
            return val_2
        else:
            if operator == '+':
                return CellValue(decimal.Decimal(CellValue.strip_trailing_zeros(str(val_1.val + val_2.val))))
            elif operator == '-':
                return CellValue(decimal.Decimal(CellValue.strip_trailing_zeros(str(val_1.val - val_2.val))))
            else:
                raise AssertionError(f'Unexpected operation: {operator}')
        
    @visit_children_decor
    def mul_expr(self, values):
        val_1, operator, val_2 = values[0], values[1], values[2]
        val_1.to_number()
        val_2.to_number()
        
        if val_1.is_cell_error():
            return val_1
        elif val_2.is_cell_error():
            return val_2
        else:
            if operator == '*':
                return CellValue(decimal.Decimal(CellValue.strip_trailing_zeros(str(val_1.val * val_2.val))))
            elif operator == '/':
                if (val_2.val == 0):
                    return CellValue(CellError(CellErrorType.DIVIDE_BY_ZERO, 'Cannot divide by zero'))
                return CellValue(decimal.Decimal(CellValue.strip_trailing_zeros(str(val_1.val / val_2.val))))
            else:
                raise AssertionError(f'Unexpected operation: {operator}')
    
    @visit_children_decor
    def unary_op(self, values):
        operator, val = values[0], values[1]
        val.to_number()

        if val.is_cell_error():
            return val
        
        if operator == '+':
            val.val = decimal.Decimal(CellValue.strip_trailing_zeros(str(val.val)))
            return val
        elif operator == '-':
            val.val = -decimal.Decimal(CellValue.strip_trailing_zeros(str(val.val)))
            return val
        else:
            raise AssertionError(f'Unexpected operation: {operator}')

    @visit_children_decor
    def concat_expr(self, values):
        assert len(values) == 2, 'Unexpected number of args'

        val_1, val_2 = values[0], values[1]
        val_1.to_string()
        val_2.to_string()

        if val_1.is_cell_error():
            return val_1
        if val_2.is_cell_error():
            return val_1

        output = CellValue(val_1.val + val_2.val)
        output.to_string()
        return output

    def error(self, tree):
        error_value = CellError(FormulaEvaluator.error_dict[tree.children[0].lower()], 'String representation of error given')
        return CellValue(error_value)

    def parens(self, tree):
        values = self.visit_children(tree)
        assert len(values) == 1, f'Unexpected tree {tree.pretty()}'
        if values[0] is None or values[0].val is None:
            return CellValue(decimal.Decimal('0'))
        else:
            return values[0]
    
    def number(self, tree):
        # called when run into number node 
        return CellValue(decimal.Decimal(tree.children[0]))

    def string(self, tree):
        # called when run into a string node
        return CellValue(tree.children[0].value[1:-1])
    
    def cell(self, tree):
        # first parse the value into sheet (if given) and location
        if (len(tree.children) == 1):
            reference = tree.children[0].value.lower().replace('$', '')
            if reference not in self.ref_info:
                return CellValue(CellError(CellErrorType.BAD_REFERENCE, f'Could not find cell information for reference {reference}'))
            return CellValue(self.ref_info[reference])
        elif (len(tree.children) == 2):
            sheet = tree.children[0].value.lower()
            location = tree.children[1].value.lower().replace('$', '')
            reference = sheet + '!' + location
            if reference not in self.ref_info:
                return CellValue(CellError(CellErrorType.BAD_REFERENCE, f'Could not find cell information for reference {reference}'))
            return CellValue(self.ref_info[reference])
        else:
            raise AssertionError('Length of tree for cell is not one or two')