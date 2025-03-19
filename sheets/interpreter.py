import decimal
from .CellError import CellError, CellErrorType
from .CellValue import CellValue
from .Sheet import Sheet
import lark
from .Cell import Cell 
from lark.visitors import visit_children_decor
from typing import Any
import re

decimal.getcontext().prec = 500

def is_valid_location(location: str) -> bool:
    # Checks if a given location string is a valid spreadsheet cell location.
    pattern = r'^[A-Za-z]{1,4}[1-9][0-9]{0,3}$'
    return bool(re.match(pattern, location))

class FormulaEvaluator(lark.visitors.Interpreter):
    def __init__(self, sheet_name, workbook, func_directory):
        self.sheet_name = sheet_name
        self.workbook = workbook
        self.func_directory = func_directory
        self.refs = set()

    error_dict = {
        "#error!": CellErrorType.PARSE_ERROR,
        "#circref!": CellErrorType.CIRCULAR_REFERENCE,
        "#ref!": CellErrorType.BAD_REFERENCE,
        "#name?": CellErrorType.BAD_NAME,
        "#value!": CellErrorType.TYPE_ERROR,
        "#div/0!": CellErrorType.DIVIDE_BY_ZERO
    }
    
    @visit_children_decor
    def compare_expr(self, values):
        val_1, operator, val_2 = values[0], values[1], values[2]

        if val_1.is_cell_error():
            return val_1
        elif val_2.is_cell_error():
            return val_2
        
        if isinstance(val_1.val, decimal.Decimal) and isinstance(val_2.val, decimal.Decimal):
            # check if both numbers
            # do nothing
            pass
        elif isinstance(val_1.val, str) and isinstance(val_2.val, str):
            # check if both strings
            val_1.val = val_1.val.lower()
            val_2.val = val_2.val.lower()
        elif isinstance(val_1.val, bool) and isinstance(val_2.val, bool):
            # check if both booleans
            # do nothing
            pass
        elif val_1.val is None and val_2.val is None:
            # check for both being empty cell
            val_1.val = 1
            val_2.val = 1
        elif val_1.val is None:
            # check for val_1 being empty cell
            if isinstance(val_2.val, decimal.Decimal):
                val_1.val = 0
            elif isinstance(val_2.val, str):
                val_1.val = ''
            elif isinstance(val_2.val, bool):
                val_1.val = False
        elif val_2.val is None:
            # check for val_2 being empty cell
            if isinstance(val_1.val, decimal.Decimal):
                val_2.val = 0
            elif isinstance(val_1.val, str):
                val_2.val = ''
            elif isinstance(val_1.val, bool):
                val_2.val = False
        else:
            # handle different types
            if isinstance(val_1.val, decimal.Decimal):
                val_1.val = 0
            elif isinstance(val_1.val, str):
                val_1.val = 1
            elif isinstance(val_1.val, bool):
                val_1.val = 2

            if isinstance(val_2.val, decimal.Decimal):
                val_2.val = 0
            elif isinstance(val_2.val, str):
                val_2.val = 1
            elif isinstance(val_2.val, bool):
                val_2.val = 2

        if operator == '=' or operator == '==':
            return CellValue(val_1.val == val_2.val)
        elif operator == '<>' or operator == '!=':
            return CellValue(val_1.val != val_2.val)
        elif operator == '<':
            return CellValue(val_1.val < val_2.val)
        elif operator == '>':
            return CellValue(val_1.val > val_2.val)
        elif operator == '<=':
            return CellValue(val_1.val <= val_2.val)
        elif operator == '>=':
            return CellValue(val_1.val >= val_2.val)
        else:
            raise AssertionError('Compare operator is unidentifiable.')
        
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
            return val_2

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
    
    def boolean(self, tree):
        # called when run into a boolean
        s = tree.children[0].lower()
        if s == 'true':
            return CellValue(True)
        else:
            return CellValue(False)
        
    def cell_range(self, tree):
        curr_tree = tree.children[0]
        args = curr_tree.children
        if len(args) == 3:
            sheet_name, start_location, end_location = args[0].lower(), args[1], args[2]
        else:
            sheet_name = self.sheet_name.lower()
            start_location, end_location = args[0], args[1]
        
        if (not is_valid_location(start_location)) \
        or (not is_valid_location(end_location)):
            raise ValueError('Spreadsheet cell location is invalid. ZZZZ9999 is the bottom-right-most cell.')
        
        start_col, start_row = Sheet.split_cell_ref(start_location)
        end_col, end_row = Sheet.split_cell_ref(end_location)

        top_left_col = min(start_col, end_col)
        top_left_row = min(start_row, end_row)
        bottom_right_col = max(start_col, end_col)
        bottom_right_row = max(start_row, end_row)

        m = bottom_right_col - top_left_col + 1
        n = bottom_right_row - top_left_row + 1
        
        orig_cells = [[0 for _ in range(m)] for _ in range(n)]
        for i in range(n):
            source_row = top_left_row + i
            for j in range(m):
                source_col = top_left_col + j
                orig_loc = Sheet.to_sheet_coords(source_col, source_row)
                self.refs.add((self.sheet_name.lower(), orig_loc))

                cell = self.workbook.get_cell(sheet_name, orig_loc)
                orig_cells[i][j] = cell        
    
        return orig_cells

    def function(self, tree):
        # extract the sheet name and the arguments
        # this checks if there are arguments - if there are arguments, tree.children[0] is a Tree
        if isinstance(tree.children[0], lark.Tree):
            curr_tree = tree.children[0]
            function_name = curr_tree.children[0].upper()
            arg_tree = curr_tree.children[1]
        else:
            # otherwise, there are no arguments
            function_name = tree.children[0].upper()
            arg_tree = None

        if function_name not in self.func_directory:
            return CellValue(CellError(CellErrorType.BAD_NAME, f"Unknown function: {function_name}"))
        return self.func_directory[function_name](arg_tree, self)

    def cell(self, tree):
        # first parse the value into sheet (if given) and location
        if (len(tree.children) == 1):
            location = tree.children[0].value.lower().replace('$', '')
            if is_valid_location(location):
                self.refs.add((self.sheet_name.lower(), location))
            try:
                output = self.workbook.get_cell_value(self.sheet_name, location)
                return CellValue(output)
            except (ValueError, KeyError) as e:
                output = CellError(CellErrorType.BAD_REFERENCE, 'Bad reference', e)
                return CellValue(output)
        elif (len(tree.children) == 2):
            sheet = tree.children[0].value.lower()
            location = tree.children[1].value.lower().replace('$', '')
            # strip quotes
            if (len(sheet) > 2 and sheet[0] == '\'' and sheet[-1] == '\''):
                sheet = sheet[1:-1]
            if is_valid_location(location):
                self.refs.add((sheet, location))
            try:
                output = self.workbook.get_cell_value(sheet, location)
                return CellValue(output)
            except (ValueError, KeyError) as e:
                output = CellError(CellErrorType.BAD_REFERENCE, 'Bad reference', e)
                return CellValue(output)
        else:
            raise AssertionError('Length of tree for cell is not one or two')