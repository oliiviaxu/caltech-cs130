from collections import defaultdict
import sheets
from .CellValue import CellValue

# BOOLEAN FUNCTIONS
def and_function(*args):
    """Returns TRUE if all arguments are TRUE. All arguments are converted to Boolean values."""
    converted = []
    for cell_val in args[0]:
        cell_val.to_bool()
        converted.append(cell_val.val)
    return all(arg for arg in converted)

def or_function(*args):
    """Returns TRUE if any argument is TRUE. All arguments are converted to Boolean values."""
    converted = []
    for cell_val in args[0]:
        cell_val.to_bool()
        converted.append(cell_val.val)
    return any(arg for arg in converted)

def not_function(arg):
    """Returns the logical negation of the argument. The argument is converted to a Boolean value."""
    arg.to_bool()
    return not arg.val

def xor_function(*args):
    """Returns TRUE if an odd number of arguments are TRUE. All arguments are converted to Boolean values."""
    converted = []
    for cell_val in args[0]:
        cell_val.to_bool()
        converted.append(cell_val.val)
    return sum(arg for arg in converted) % 2 == 1

# STRING-MATCH FUNCTIONS
def exact_function(*args):
    """Returns TRUE if the two strings are identical. Case-sensitive. The arguments are converted to string values."""
    converted = []
    for cell_val in args[0]:
        cell_val.to_string()
        converted.append(cell_val.val)
    return converted[0] == converted[1]

# CONDITIONAL FUNCTIONS
def if_function(*args):
    # arguments: condition, true_value, false_value=None 
    """Returns `true_value` if `condition` is TRUE, otherwise `false_value`. The condition is converted to a Boolean value."""
    condition = args[0]
    print(args)
    pass

def iferror_function(value, value_if_error=""):
    """Returns `value` if it is not an error, otherwise `value_if_error`."""
    return value if not isinstance(value, Exception) else value_if_error

def choose_function(index, *args):
    """Returns the `index`-th argument (1-based indexing). The index is converted to a number."""
    try:
        index = int(index)
        if 1 <= index <= len(args):
            return args[index - 1]
        else:
            raise TypeError("Index out of range")
    except (ValueError, TypeError):
        raise TypeError("Index must be a valid integer")

# INFORMATIONAL FUNCTIONS
def isblank_function(value):
    """Returns TRUE if the value is blank or None."""
    return value is None or value == ""

def iserror_function(value):
    """Returns TRUE if the value is an error."""
    return isinstance(value, Exception)

def version_function():
    """Returns the version of the spreadsheet library."""
    return sheets.__version__

def indirect_function(workbook):
    """Returns a function that parses its string argument as a cell-reference and returns the value of the specified cell."""
    def indirect(cell_ref):
        return workbook.get_cell_value(str(cell_ref))
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