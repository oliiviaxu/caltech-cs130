from collections import defaultdict
import sheets
from .CellValue import CellValue
from .CellError import CellError, CellErrorType
import decimal

# TODO: Propagate cell errors
# BOOLEAN FUNCTIONS
def and_function(args):
    """Returns TRUE if all arguments are TRUE. All arguments are converted to Boolean values."""
    if not args:
        return CellValue(CellError(CellErrorType.TYPE_ERROR, "Expected at least 1 arguments, but got 0 arguments."))
    
    converted = []
    for cell_val in args:
        if not isinstance(cell_val.val, bool):
            cell_val.to_bool()

            if isinstance(cell_val.val, sheets.CellError):
                return cell_val
            
        converted.append(cell_val.val)
    return CellValue(all(arg for arg in converted))

def or_function(args):
    """Returns TRUE if any argument is TRUE. All arguments are converted to Boolean values."""
    if not args:
        return CellValue(CellError(CellErrorType.TYPE_ERROR, "Expected at least 1 arguments, but got 0 arguments."))
    converted = []
    for cell_val in args:
        if not isinstance(cell_val.val, bool):
            cell_val.to_bool()

            if isinstance(cell_val.val, sheets.CellError):
                return cell_val
            
        converted.append(cell_val.val)
    return CellValue(any(arg for arg in converted))

def not_function(args):
    """Returns the logical negation of the argument. The argument is converted to a Boolean value."""
    if len(args) != 1:
        return CellValue(CellError(CellErrorType.TYPE_ERROR, f"Expected exactly 1 argument, but got {len(args)} arguments."))
    arg = args[0]
    if not isinstance(arg.val, bool):
        arg.to_bool()

        if isinstance(arg.val, sheets.CellError):
            return arg
        
    return CellValue(not arg.val)

def xor_function(args):
    """Returns TRUE if an odd number of arguments are TRUE. All arguments are converted to Boolean values."""
    if not args:
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
def exact_function(args):
    """Returns TRUE if the two strings are identical. Case-sensitive. The arguments are converted to string values."""
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
def if_function(args):
    # arguments: condition, true_value, false_value=None 
    """Returns `true_value` if `condition` is TRUE, otherwise `false_value`. The condition is converted to a Boolean value."""
    if len(args) != 2 and len(args) != 3:
        return CellValue(CellError(CellErrorType.TYPE_ERROR, f"Expected 2 or 3 arguments, but got {len(args)} arguments."))      
    
    if not isinstance(args[0].val, bool):    
        args[0].to_bool()

        if isinstance(args[0].val, sheets.CellError):
            return args[0]
    
    condition, true_value, false_value = args[0], args[1].val, None
    if len(args) == 3:
        false_value = args[2].val
    else:
        false_value = None

    if condition.val:
        return CellValue(true_value)
    else:
        return CellValue(false_value)
    
def iferror_function(args):
    # TODO: 
    # arguments: value, value_if_error=""
    """Returns `value` if it is not an error, otherwise `value_if_error`."""
    # TODO: check
    # if len(args) != 1 and len(args) != 2:
    #     return CellValue(CellError(CellErrorType.TYPE_ERROR, f"Expected 1 or 2 arguments, but got {len(args)} arguments."))      
    
    # value, value_if_error = None, ""
    # if len(args) == 1:
    #     value = args[0]
    # else:
    #     value = args[0]
    #     value_if_error = args[1]
    
    pass

def choose_function(args):
    """Returns the `index`-th argument (1-based indexing). The index is converted to a number."""
    if len(args) < 2:
        return CellValue(CellError(CellErrorType.TYPE_ERROR, f"Expected atleast 2 arguments, but got {len(args)} arguments."))  

    if not isinstance(args[0].val, decimal.Decimal):    
        args[0].to_number()

        if isinstance(args[0].val, sheets.CellError):
            return args[0]

    remaining_args = args[1:]
    
    if (args[0].val != int(args[0].val)) or (args[0].val <= 0) or (args[0].val > len(remaining_args)):
        return CellValue(CellError(CellErrorType.TYPE_ERROR, f"Invalid Index."))  
    
    index = int(args[0].val) - 1
    return CellValue(remaining_args[index].val)

# INFORMATIONAL FUNCTIONS
def isblank_function(args):
    """
    Returns TRUE if empty cell value.
    evaluates to TRUE if its input is an empty-cell value, or FALSE otherwise. 
    This function always takes exactly one argument. Note specifically that ISBLANK("") evaluates to FALSE, 
    as do ISBLANK(FALSE) and ISBLANK(0).
    """
    if len(args) != 1:
        return CellValue(CellError(CellErrorType.TYPE_ERROR, f"Expected exactly 1 arguments, but got {len(args)} arguments."))

    arg = args[0]
    if isinstance(arg.val, sheets.CellError):
        return arg
        
    if arg.val is None:
        return CellValue(True)
    else:
        return CellValue(False)

def iserror_function(args):
    """Returns TRUE if the value is an error."""
    if len(args) != 1:
        return CellValue(CellError(CellErrorType.TYPE_ERROR, f"Expected exactly 1 arguments, but got {len(args)} arguments."))
    
    arg = args[0]
    if isinstance(arg.val, sheets.CellError):
        return CellValue(True)
    else:
        return CellValue(False)

def version_function(args):
    """Returns the version of the spreadsheet library."""
    if len(args) != 0:
        return CellValue(CellError(CellErrorType.TYPE_ERROR, f"Expected no arguments, but got {len(args)} arguments."))
    return sheets.__version__

def indirect_function(workbook):
    """Returns a function that parses its string argument as a cell-reference and returns the value of the specified cell."""
    def indirect(args):
        return workbook.get_cell_value(str(args))
    pass

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