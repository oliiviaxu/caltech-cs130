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
    if isinstance(arg.val, sheets.CellError):
        return arg
        
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
    }
    return BUILTIN_SPREADSHEET_FUNCTIONS