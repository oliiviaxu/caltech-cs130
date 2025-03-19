from collections import defaultdict
import sheets
from .CellValue import CellValue
from .CellError import CellError, CellErrorType
import decimal
import re

def sheet_name_needs_quotes(sheet_name):
    pattern = r"^[A-Za-z_][A-Za-z0-9_]*$"
    return not bool(re.fullmatch(pattern, sheet_name))

def visit_all(arg_tree, ev):
    if arg_tree is None:
        return []
    args = ev.visit(arg_tree)
    if not isinstance(args, list):
        args = [args]
    return args

def get_first_arg(arg_tree, ev, valid_arg_len):
    if arg_tree is None:
        return None, True
    elif len(arg_tree.children) not in valid_arg_len:
        return None, True
    else:
        arg_one = ev.visit(arg_tree.children[0])
        return arg_one, False

def get_ith_arg(arg_tree, ev, i):
    if i < 0 or i >= len(arg_tree.children):
        return None
    return ev.visit(arg_tree.children[i])

# TODO: Propagate cell errors
# BOOLEAN FUNCTIONS
def and_function(arg_tree, ev):
    """Returns TRUE if all arguments are TRUE. All arguments are converted to Boolean values."""
    args = visit_all(arg_tree, ev)
    if len(args) == 0:
        return CellValue(CellError(CellErrorType.TYPE_ERROR, "Expected at least 1 arguments, but got 0 arguments."))
    
    converted = []
    for cell_val in args:
        if not isinstance(cell_val.val, bool):
            cell_val.to_bool()

            if isinstance(cell_val.val, sheets.CellError):
                return cell_val
            
        converted.append(cell_val.val)
    return CellValue(all(arg for arg in converted))

def or_function(arg_tree, ev):
    """Returns TRUE if any argument is TRUE. All arguments are converted to Boolean values."""
    args = visit_all(arg_tree, ev)
    if len(args) == 0:
        return CellValue(CellError(CellErrorType.TYPE_ERROR, "Expected at least 1 arguments, but got 0 arguments."))
    
    converted = []
    for cell_val in args:
        if not isinstance(cell_val.val, bool):
            cell_val.to_bool()

            if isinstance(cell_val.val, sheets.CellError):
                return cell_val
            
        converted.append(cell_val.val)
    return CellValue(any(arg for arg in converted))

def not_function(arg_tree, ev):
    """Returns the logical negation of the argument. The argument is converted to a Boolean value."""
    args = visit_all(arg_tree, ev)
    if len(args) != 1:
        return CellValue(CellError(CellErrorType.TYPE_ERROR, f"Expected exactly 1 argument, but got {len(args)} arguments."))
    
    arg = args[0]
    if not isinstance(arg.val, bool):
        arg.to_bool()

        if isinstance(arg.val, sheets.CellError):
            return arg
        
    return CellValue(not arg.val)

def xor_function(arg_tree, ev):
    """Returns TRUE if an odd number of arguments are TRUE. All arguments are converted to Boolean values."""
    args = visit_all(arg_tree, ev)
    if len(args) == 0:
        return CellValue(CellError(CellErrorType.TYPE_ERROR, "Expected at least 1 arguments, but got 0 arguments."))
    converted = []
    for cell_val in args:
        if not isinstance(cell_val.val, bool):
            cell_val.to_bool()

            if isinstance(cell_val.val, sheets.CellError):
                return cell_val
            
        converted.append(cell_val.val)
    return CellValue(sum(arg for arg in converted) % 2 == 1)

# STRING-MATCH FUNCTIONS
def exact_function(arg_tree, ev):
    """Returns TRUE if the two strings are identical. Case-sensitive. The arguments are converted to string values."""
    args = visit_all(arg_tree, ev)
    if len(args) != 2:
        return CellValue(CellError(CellErrorType.TYPE_ERROR, f"Expected 2 arguments, but got {len(args)} arguments."))      
    converted = []
    for cell_val in args:
        if not isinstance(cell_val.val, str):
            cell_val.to_string()

            if isinstance(cell_val.val, sheets.CellError):
                return cell_val
            
        converted.append(cell_val.val)
    return CellValue(converted[0] == converted[1])

# CONDITIONAL FUNCTIONS
def if_function(arg_tree, ev):
    # arguments: condition, true_value, false_value=None 
    """Returns `true_value` if `condition` is TRUE, otherwise `false_value`. The condition is converted to a Boolean value."""
    arg_one, not_enough_args = get_first_arg(arg_tree, ev, [2, 3])
    if not_enough_args:
        return CellValue(CellError(CellErrorType.TYPE_ERROR, f"Expected 2 or 3 arguments."))      
    
    if not isinstance(arg_one.val, bool):    
        arg_one.to_bool()

        if isinstance(arg_one.val, sheets.CellError):
            return arg_one
    
    if arg_one.val:
        value_1 = get_ith_arg(arg_tree, ev, 1)
        return value_1
    else:
        value_2 = get_ith_arg(arg_tree, ev, 2)
        if value_2 is None:
            value_2 = CellValue(False)
        return value_2
    
def iferror_function(arg_tree, ev):
    """Returns `value` if it is not an error, otherwise `value_if_error`."""
    arg_one, not_enough_args = get_first_arg(arg_tree, ev, [1, 2])
    if not_enough_args:
        return CellValue(CellError(CellErrorType.TYPE_ERROR, f"Expected 1 or 2 arguments."))     
    
    if not isinstance(arg_one.val, CellError):
        return arg_one
    else:
        value_2 = get_ith_arg(arg_tree, ev, 1)
        if value_2 is None:
            value_2 = CellValue("")
        return value_2

def choose_function(arg_tree, ev):
    """Returns the `index`-th argument (1-based indexing). The index is converted to a number."""
    if arg_tree is None or len(arg_tree.children) < 2:
        return CellValue(CellError(CellErrorType.TYPE_ERROR, f"Expected atleast 2 arguments."))
    
    arg_one = ev.visit(arg_tree.children[0])
    if not isinstance(arg_one.val, decimal.Decimal):    
        arg_one.to_number()

        if isinstance(arg_one.val, sheets.CellError):
            return arg_one
    
    if (arg_one.val != int(arg_one.val)) or (arg_one.val <= 0) or (arg_one.val > len(arg_tree.children)-1):
        return CellValue(CellError(CellErrorType.TYPE_ERROR, f"Invalid Index."))  
    
    index = int(arg_one.val)
    return ev.visit(arg_tree.children[index])

# INFORMATIONAL FUNCTIONS
def isblank_function(arg_tree, ev):
    """
    Returns TRUE if empty cell value.
    evaluates to TRUE if its input is an empty-cell value, or FALSE otherwise. 
    This function always takes exactly one argument. Note specifically that ISBLANK("") evaluates to FALSE, 
    as do ISBLANK(FALSE) and ISBLANK(0).
    """
    args = visit_all(arg_tree, ev)
    if len(args) != 1:
        return CellValue(CellError(CellErrorType.TYPE_ERROR, f"Expected exactly 1 arguments, but got {len(args)} arguments."))

    arg = args[0]
    # if isinstance(arg.val, sheets.CellError) and arg.val.get_type() == sheets.CellErrorType.CIRCULAR_REFERENCE:
    #     return arg
        
    if arg.val is None:
        return CellValue(True)
    else:
        return CellValue(False)

def iserror_function(arg_tree, ev):
    """Returns TRUE if the value is an error."""
    args = visit_all(arg_tree, ev)
    if len(args) != 1:
        return CellValue(CellError(CellErrorType.TYPE_ERROR, f"Expected exactly 1 arguments, but got {len(args)} arguments."))
    
    arg = args[0]
    if isinstance(arg.val, sheets.CellError):
        return CellValue(True)
    else:
        return CellValue(False)

def version_function(arg_tree, ev):
    """Returns the version of the spreadsheet library."""
    args = visit_all(arg_tree, ev)
    if len(args) != 0:
        return CellValue(CellError(CellErrorType.TYPE_ERROR, f"Expected no arguments, but got {len(args)} arguments."))
    return CellValue(sheets.__version__)

def indirect_function(workbook):
    """Returns a function that parses its string argument as a cell-reference and 
    returns the value of the specified cell. 
    
    This function always takes exactly one argument. The argument is converted to a string.

    If the argument string cannot be parsed as a cell reference for any reason, or if the 
    argument can be parsed as a cell reference, but the cell-reference is invalid for some reason 
    (other than creating a circular reference in the workbook), this function returns a BAD_REFERENCE error."""

    def indirect(arg_tree, ev):
        args = visit_all(arg_tree, ev)
        if len(args) != 1:
            return CellValue(CellError(CellErrorType.TYPE_ERROR, f"Expected exactly 1 arguments, but got {len(args)} arguments."))
        
        arg = args[0]

        if not isinstance(arg.val, str):
            arg.to_string()

            if isinstance(arg.val, sheets.CellError):
                return arg
        
        # attempt to parse as a cell reference
        split_ref = arg.val.split('!')
        # print(split_ref)

        ref_sheet_name, ref_location = "", ""
        if len(split_ref) == 1:
            ref_sheet_name = ev.sheet_name.lower()
            ref_location = split_ref[0].lower().replace('$', '')
        else:
            ref_sheet_name = split_ref[0].lower()
            if (len(ref_sheet_name) > 2 and ref_sheet_name[0] == '\'' and ref_sheet_name[-1] == '\''):
                ref_sheet_name = ref_sheet_name[1:-1]
            elif sheet_name_needs_quotes(ref_sheet_name):
                return CellValue(CellError(CellErrorType.BAD_REFERENCE, f"Invalid reference: {str(ref_sheet_name)}"))
            ref_location = split_ref[1].lower().replace('$', '')
        
        if not sheets.Workbook.is_valid_location(ref_location):
            return CellValue(CellError(CellErrorType.BAD_REFERENCE, f"Failed to parse {arg.val} as a cell reference."))
        
        try:
            output = CellValue(workbook.get_cell_value(ref_sheet_name, ref_location))
            ev.refs.add((ref_sheet_name, ref_location))
            return output
        except (KeyError, ValueError) as e:
            return CellValue(CellError(CellErrorType.BAD_REFERENCE, f"Invalid reference: {str(e)}"))
    
    return indirect

# EXTRA CREDIT
def min_function(arg_tree, ev):
    """
    MIN(value1, ...) returns the minimum value over the set of inputs. 
    Arguments may include cell-range references as well as normal expressions; 
    values from the cell-range are also considered by the function. All non-empty 
    inputs are converted to numbers; if any input cannot be converted to a number 
    then the function returns a TYPE_ERROR. Only non-empty cells should be considered; 
    empty cells should be ignored. 
    """
    args = visit_all(arg_tree, ev)
    if len(args) == 0:
        return CellValue(CellError(CellErrorType.TYPE_ERROR, "Expected at least 1 arguments, but got 0 arguments."))

    converted = []
    for cell_val in args:
        if cell_val is None:
            continue
        
        if isinstance(cell_val, list):
            m, n = len(cell_val), len(cell_val[0])
            for i in range(m):
                for j in range(n):
                    if cell_val[i][j]:
                        curr = cell_val[i][j].value

                        if curr.val is None:
                            continue

                        if not isinstance(curr.val, decimal.Decimal):
                            curr.to_number()

                            if isinstance(curr.val, sheets.CellError):
                                return curr

                        converted.append(curr.val)

        elif cell_val.val is None:
            continue
        else:
            if not isinstance(cell_val.val, decimal.Decimal):
                cell_val.to_number()

                if isinstance(cell_val.val, sheets.CellError):
                    return cell_val
            
            converted.append(cell_val.val)

    # print('converted', converted)
    if not converted:
        return CellValue(decimal.Decimal('0'))

    return CellValue(min(converted))

def max_function(arg_tree, ev):
    args = visit_all(arg_tree, ev)
    if len(args) == 0:
        return CellValue(CellError(CellErrorType.TYPE_ERROR, "Expected at least 1 arguments, but got 0 arguments."))

    converted = []
    for cell_val in args:
        if cell_val is None:
            continue
        
        if isinstance(cell_val, list):
            m, n = len(cell_val), len(cell_val[0])
            for i in range(m):
                for j in range(n):
                    if cell_val[i][j]:
                        curr = cell_val[i][j].value

                        if curr.val is None:
                            continue

                        if isinstance(curr.val, sheets.CellError):
                            return curr

                        if not isinstance(curr.val, decimal.Decimal):
                            curr.to_number()

                            if isinstance(curr.val, sheets.CellError):
                                return curr

                        converted.append(curr.val)
        elif cell_val.val is None:
            continue
        else:
            if not isinstance(cell_val.val, decimal.Decimal):
                cell_val.to_number()

                if isinstance(cell_val.val, sheets.CellError):
                    return cell_val
            
            converted.append(cell_val.val)

    # print('converted', converted)
    if not converted:
        return CellValue(decimal.Decimal('0'))

    return CellValue(max(converted))

def sum_function(arg_tree, ev):
    """
    returns the sum of all inputs. Arguments may include cell-range references as well 
    as normal expressions; values from the cell-range are added into the sum.
    All non-empty inputs are converted to numbers; if any input cannot be converted 
    to a number then the function returns a TYPE_ERROR. If the functions inputs only 
    include empty cells then the functions result is 0. This function requires at least 1 argument.
    """
    args = visit_all(arg_tree, ev)
    if len(args) == 0:
        return CellValue(CellError(CellErrorType.TYPE_ERROR, "Expected at least 1 arguments, but got 0 arguments."))
    
    converted = []
    for cell_val in args:
        if cell_val is None:
            continue
        
        # cell range arguments
        if isinstance(cell_val, list):
            m, n = len(cell_val), len(cell_val[0])
            for i in range(m):
                for j in range(n):
                    if cell_val[i][j]:
                        curr = cell_val[i][j].value

                        if curr.val is None:
                            continue

                        if not isinstance(curr.val, decimal.Decimal):
                            curr.to_number()

                            if isinstance(curr.val, sheets.CellError):
                                return curr

                        converted.append(curr.val)
        elif cell_val.val is None:
            continue
        else:
            if not isinstance(cell_val.val, decimal.Decimal):
                cell_val.to_number()

                if isinstance(cell_val.val, sheets.CellError):
                    return cell_val
            
            converted.append(cell_val.val)

    # print('converted', converted)
    if not converted:
        return CellValue(decimal.Decimal('0'))

    return CellValue(sum(converted))

def average_function(arg_tree, ev):
    args = visit_all(arg_tree, ev)
    if len(args) == 0:
        return CellValue(CellError(CellErrorType.TYPE_ERROR, "Expected at least 1 arguments, but got 0 arguments."))
    
    converted = []
    for cell_val in args:
        if cell_val is None:
            continue
        
        # cell range arguments
        if isinstance(cell_val, list):
            m, n = len(cell_val), len(cell_val[0])
            for i in range(m):
                for j in range(n):
                    if cell_val[i][j]:
                        curr = cell_val[i][j].value

                        if curr.val is None:
                            continue

                        if not isinstance(curr.val, decimal.Decimal):
                            curr.to_number()

                            if isinstance(curr.val, sheets.CellError):
                                return curr

                        converted.append(curr.val)
        elif cell_val.val is None:
            continue
        else:
            if not isinstance(cell_val.val, decimal.Decimal):
                cell_val.to_number()

                if isinstance(cell_val.val, sheets.CellError):
                    return cell_val
            
            converted.append(cell_val.val)

    # print('converted', converted)
    if not converted:
        return CellValue(CellError(CellErrorType.DIVIDE_BY_ZERO, "All arguments are None."))
    
    total = sum(converted)
    count = len(converted)
    average = total / count

    return CellValue(average)

def hlookup_function(arg_tree, ev):
    """
    HLOOKUP(key, range, index) searches horizontally through a range of cells. 
    The function searches through the first (i.e. topmost) row in range, looking 
    for the first column that contains key in the search row (exact match, both 
    type and value). If such a column is found, the cell in the index-th row of 
    the found column. The index is 1-based; an index of 1 refers to the search row. 
    If no column is found, the function returns a TYPE_ERROR.
    """
    args = visit_all(arg_tree, ev)
    if len(args) != 3:
        return CellValue(CellError(CellErrorType.TYPE_ERROR, "HLOOKUP requires exactly 3 arguments."))
    
    key, cell_range, index = args

    if not int(index.val) == index.val:
        return CellValue(CellError(CellErrorType.TYPE_ERROR, "Index must be an integer."))
    if index.val < 1:
        return CellValue(CellError(CellErrorType.TYPE_ERROR, "Index must be at least 1."))
    
    index = int(index.val)
    
    if not isinstance(cell_range, list):
        return CellValue(CellError(CellErrorType.TYPE_ERROR, "Range must be a cell range."))
    
    first_row = cell_range[0]
    column_found = False

    for col_idx, cell in enumerate(first_row):

        if cell.value.val == key.val:
            if index > len(cell_range):
                return CellValue(CellError(CellErrorType.TYPE_ERROR, "Index is out of range."))
            column_found = True
            output_val = cell_range[index - 1][col_idx].value.val
            return CellValue(output_val)

    if not column_found:
        return CellValue(CellError(CellErrorType.TYPE_ERROR, "No such column found."))
        

def vlookup_function(arg_tree, ev):
    """
    VLOOKUP(key, range, index) searches vertically through a range of cells.
    The function searches through the first (i.e. leftmost) column in range,
    looking for the first row that contains key in the search column (exact
    match, both type and value). If such a row is found, the cell in the
    index-th column of the found row is returned. The index is 1-based; an
    index of 1 refers to the search column. If no row is found, the function
    returns a TYPE_ERROR.
    """
    args = visit_all(arg_tree, ev)
    if len(args) != 3:
        return CellValue(CellError(CellErrorType.TYPE_ERROR, "VLOOKUP requires exactly 3 arguments."))

    key, cell_range, index = args

    # Validate the index
    if not int(index.val) == index.val:
        return CellValue(CellError(CellErrorType.TYPE_ERROR, "Index must be an integer."))
    if index.val < 1:
        return CellValue(CellError(CellErrorType.TYPE_ERROR, "Index must be at least 1."))

    index = int(index.val)

    if not isinstance(cell_range, list):
        return CellValue(CellError(CellErrorType.TYPE_ERROR, "Range must be a cell range."))

    row_found = False
    for row_idx, row in enumerate(cell_range):

        if row[0].value.val == key.val:
            if index > len(row):
                return CellValue(CellError(CellErrorType.TYPE_ERROR, "Index is out of range."))
            row_found = True

            output_val = row[index - 1].value.val
            return CellValue(output_val)

    if not row_found:
        return CellValue(CellError(CellErrorType.TYPE_ERROR, "Key not found in the search column."))

def create_function_directory(workbook):
    BUILTIN_SPREADSHEET_FUNCTIONS = {
        # BOOLEAN FUNCTIONS
        "AND": and_function, 
        "OR": or_function,
        "NOT": not_function,
        "XOR": xor_function,

        # STRING-MATCH FUNCTIONS
        "EXACT": exact_function,

        # CONDITIONAL FUNCTIONS
        "IF": if_function,     
        "IFERROR": iferror_function, 
        "CHOOSE": choose_function,

        # INFORMATIONAL FUNCTIONS
        "ISBLANK": isblank_function,
        "ISERROR": iserror_function,
        "VERSION": version_function,

        # INDIRECT (is workbook-specific)
        "INDIRECT": indirect_function(workbook),

        # EXTRA CREDIT 
        "MIN": min_function,
        "MAX": max_function,
        "SUM": sum_function,
        "AVERAGE": average_function,
        "HLOOKUP": hlookup_function,
        "VLOOKUP": vlookup_function,
    }
    return BUILTIN_SPREADSHEET_FUNCTIONS