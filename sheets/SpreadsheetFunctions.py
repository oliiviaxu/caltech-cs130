from collections import defaultdict
import sheets
from .CellValue import CellValue
from .CellError import CellError, CellErrorType
# TODO: Propagate cell errors
# BOOLEAN FUNCTIONS
def and_function(args):
    """Returns TRUE if all arguments are TRUE. All arguments are converted to Boolean values."""
    if not args:
        return CellValue(CellError(CellErrorType.TYPE_ERROR, "Expected at least 1 arguments, but got 0 arguments."))
    
    converted = []
    for cell_val in args:
        if not isinstance(cell_val, bool):
            cell_val.to_bool()

            # TODO: not sure if this is the case
            if isinstance(cell_val.val, sheets.CellError):
                return CellValue(CellError(CellErrorType.TYPE_ERROR, f"Could not convert to boolean type."))
            
            converted.append(cell_val.val)
        else:
            converted.append(cell_val)
    return all(arg for arg in converted)

def or_function(args):
    """Returns TRUE if any argument is TRUE. All arguments are converted to Boolean values."""
    if not args:
        return CellValue(CellError(CellErrorType.TYPE_ERROR, "Expected at least 1 arguments, but got 0 arguments."))
    converted = []
    for cell_val in args:
        if not isinstance(cell_val, bool):
            cell_val.to_bool()

            # TODO: not sure if this is the case
            if isinstance(cell_val.val, sheets.CellError):
                return CellValue(CellError(CellErrorType.TYPE_ERROR, f"Could not convert to boolean type."))
            
            converted.append(cell_val.val)
        else:
            converted.append(cell_val)
    return any(arg for arg in converted)

def not_function(args):
    """Returns the logical negation of the argument. The argument is converted to a Boolean value."""
    if len(args) != 1:
        return CellValue(CellError(CellErrorType.TYPE_ERROR, f"Expected exactly 1 argument, but got {len(args)} arguments."))
    arg = args[0]
    if not isinstance(arg, bool):
        arg.to_bool()

        if isinstance(arg.val, sheets.CellError):
                return CellValue(CellError(CellErrorType.TYPE_ERROR, f"Could not convert to boolean type."))
        
        return not arg.val
    else:
        return not arg

def xor_function(args):
    """Returns TRUE if an odd number of arguments are TRUE. All arguments are converted to Boolean values."""
    if not args:
        return CellValue(CellError(CellErrorType.TYPE_ERROR, "Expected at least 1 arguments, but got 0 arguments."))
    converted = []
    for cell_val in args:
        cell_val.to_bool()
        converted.append(cell_val.val)
    return sum(arg for arg in converted) % 2 == 1

# STRING-MATCH FUNCTIONS
def exact_function(args):
    """Returns TRUE if the two strings are identical. Case-sensitive. The arguments are converted to string values."""
    if len(args) != 2:
        return CellValue(CellError(CellErrorType.TYPE_ERROR, f"Expected 2 arguments, but got {len(args)} arguments."))      
    converted = []
    for arg in args:
        arg.to_string()
        converted.append(arg.val)
    return converted[0] == converted[1]

# CONDITIONAL FUNCTIONS
def if_function(args):
    # arguments: condition, true_value, false_value=None 
    """Returns `true_value` if `condition` is TRUE, otherwise `false_value`. The condition is converted to a Boolean value."""
    if len(args) != 2 and len(args) != 3:
        return CellValue(CellError(CellErrorType.TYPE_ERROR, f"Expected 2 or 3 arguments, but got {len(args)} arguments."))      
    args[0].to_bool()

    condition, true_value, false_value = args[0], args[1], args[2]
    if condition.val:
        return true_value.val
    else:
        return false_value.val
    
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

    args[0].to_number()
    index = int(args[0].val) - 1
    
    remaining_args = args[1:]
    return remaining_args[index].val

# INFORMATIONAL FUNCTIONS
def isblank_function(args):
    # TODO
    """Returns TRUE if the value is blank or None."""
    if len(args) != 0:
        return CellValue(CellError(CellErrorType.TYPE_ERROR, f"Expected exactly 1 arguments, but got {len(args)} arguments."))

    if not args:
        return True
    else:
        return False

def iserror_function(args):
    """Returns TRUE if the value is an error."""
    if len(args) != 1:
        return CellValue(CellError(CellErrorType.TYPE_ERROR, f"Expected exactly 1 arguments, but got {len(args)} arguments."))
    
    # TODO
    pass

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