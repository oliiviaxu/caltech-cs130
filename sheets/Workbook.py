from .Sheet import *
from .Cell import *
from .visitor import CellRefFinder
from typing import List, Optional, Tuple, Any
import os
import lark

current_dir = os.path.dirname(os.path.abspath(__file__))
lark_path = os.path.join(current_dir, "formulas.lark")

class Workbook:
    # A workbook containing zero or more named spreadsheets.
    #
    # Any and all operations on a workbook that may affect calculated cell
    # values should cause the workbook's contents to be updated properly.

    def __init__(self):
        self.sheets = {} # lowercase keys  
        pass

    def num_sheets(self) -> int:
        return len(self.sheets.keys())

    def list_sheets(self) -> List[str]:
        # Return a list of the spreadsheet names in the workbook, with the
        # capitalization specified at creation, and in the order that the sheets
        # appear within the workbook.
        #
        # In this project, the sheet names appear in the order that the user
        # created them; later, when the user is able to move and copy sheets,
        # the ordering of the sheets in this function's result will also reflect
        # such operations.
        #
        # A user should be able to mutate the return-value without affecting the
        # workbook's internal state.
        lst = []
        for _, sheet in self.sheets.items():
            lst.append(sheet.sheet_name)
        return lst # preserves case 

    def new_sheet(self, sheet_name: Optional[str] = None) -> Tuple[int, str]:
        # Add a new sheet to the workbook.  If the sheet name is specified, it
        # must be unique.  If the sheet name is None, a unique sheet name is
        # generated.  "Uniqueness" is determined in a case-insensitive manner,
        # but the case specified for the sheet name is preserved.
        #
        # The function returns a tuple with two elements:
        # (0-based index of sheet in workbook, sheet name).  This allows the
        # function to report the sheet's name when it is auto-generated.
        #
        # If the spreadsheet name is an empty string (not None), or it is
        # otherwise invalid, a ValueError is raised.
        
        sheet_names_lower = [sheet_name.lower() for sheet_name in self.list_sheets()]

        if (sheet_name == None):
            num = 1
            while True:
                new_name = 'Sheet' + str(num)
                if (new_name.lower() not in sheet_names_lower):
                    sheet_name = new_name
                    break
                num += 1
        else:
            # According to spec, User-specified spreadsheet names can be comprised 
            # of letters, numbers, spaces, and these punctuation characters: .?!,:;!@#$%^&*()-_. 
            # Note specifically that all quote marks are excluded from names. Spreadsheet names 
            # cannot start or end with whitespace characters, and they cannot be an empty string.
            sheet_name = sheet_name.replace('\'', '').replace('\"', '')
            if (sheet_name == '' or sheet_name[0] in [' ', '\t', '\n'] or sheet_name[len(sheet_name)-1] in [' ', '\t', '\n']):
                raise ValueError('Spreadsheet names cannot start or end with whitespace characters, and they cannot be an empty string.')
            
            for char in sheet_name:
                if (char != ' ' and not char.isalnum() and char not in '.?!,:;!@#$%^&*()-_'):
                    raise ValueError('Spreadsheet name can be comprised of letters, numbers, spaces, and these punctuation characters: .?!,:;!@#$%^&*()-_')
        
            if (sheet_name.lower() in sheet_names_lower):
                raise ValueError('Spreadsheet names must be unique.')

        self.sheets[sheet_name.lower()] = Sheet(sheet_name)
        return len(self.sheets.keys()) - 1, sheet_name

    def del_sheet(self, sheet_name: str) -> None:
        # Delete the spreadsheet with the specified name.
        #
        # The sheet name match is case-insensitive; the text must match but the
        # case does not have to.
        #
        # If the specified sheet name is not found, a KeyError is raised.

        # TODO: put off until finish dependency graph implementation
        pass

    def get_sheet_extent(self, sheet_name: str) -> Tuple[int, int]:
        # Return a tuple (num-cols, num-rows) indicating the current extent of
        # the specified spreadsheet.
        #
        # The sheet name match is case-insensitive; the text must match but the
        # case does not have to.
        #
        # If the specified sheet name is not found, a KeyError is raised.
        if sheet_name.lower() not in self.sheets.keys():
            raise KeyError('Sheet not found.')
        
        sheet = self.sheets[sheet_name.lower()]
        num_cols, num_rows = sheet.num_cols, sheet.num_rows
        return num_cols, num_rows

    def is_valid_location(self, location: str) -> bool:
        # Checks if a given location string is a valid spreadsheet cell location.
        pattern = r'^[A-Za-z]{1,4}[1-9][0-9]{0,3}$'
        return bool(re.match(pattern, location))

    def set_cell_contents(self, sheet_name: str, location: str,
                          contents: Optional[str]) -> None:
        # Set the contents of the specified cell on the specified sheet.
        #
        # The sheet name match is case-insensitive; the text must match but the
        # case does not have to.  Additionally, the cell location can be
        # specified in any case.
        #
        # If the specified sheet name is not found, a KeyError is raised.
        # If the cell location is invalid, a ValueError is raised.
        #
        # A cell may be set to "empty" by specifying a contents of None.
        #
        # Leading and trailing whitespace are removed from the contents before
        # storing them in the cell.  Storing a zero-length string "" (or a
        # string composed entirely of whitespace) is equivalent to setting the
        # cell contents to None.
        #
        # If the cell contents appear to be a formula, and the formula is
        # invalid for some reason, this method does not raise an exception;
        # rather, the cell's value will be a CellError object indicating the
        # nature of the issue.

        if sheet_name.lower() not in self.sheets.keys():
            raise KeyError('Sheet not found.')
        
        if not self.is_valid_location(location):
            raise ValueError('Spreadsheet cell location is invalid. ZZZZ9999 is the bottom-right-most cell.') 
        
        if (contents is not None):
            contents = contents.strip()
        self.sheets[sheet_name.lower()].set_cell_contents(location, contents)

    def get_cell_contents(self, sheet_name: str, location: str) -> Optional[str]:
        # Return the contents of the specified cell on the specified sheet.
        #
        # The sheet name match is case-insensitive; the text must match but the
        # case does not have to.  Additionally, the cell location can be
        # specified in any case.
        #
        # If the specified sheet name is not found, a KeyError is raised.
        # If the cell location is invalid, a ValueError is raised.
        #
        # Any string returned by this function will not have leading or trailing
        # whitespace, as this whitespace will have been stripped off by the
        # set_cell_contents() function.
        #
        # This method will never return a zero-length string; instead, empty
        # cells are indicated by a value of None.

        if sheet_name.lower() not in self.sheets.keys():
            raise KeyError('Sheet not found.')
        
        if not self.is_valid_location(location):
            raise ValueError('Spreadsheet cell location is invalid. ZZZZ9999 is the bottom-right-most cell.')
        
        return self.sheets[sheet_name.lower()].get_cell_contents(location)

    def get_cell_value(self, sheet_name: str, location: str) -> Any:
        # Return the evaluated value of the specified cell on the specified
        # sheet.
        #
        # The sheet name match is case-insensitive; the text must match but the
        # case does not have to.  Additionally, the cell location can be
        # specified in any case.
        #
        # If the specified sheet name is not found, a KeyError is raised.
        # If the cell location is invalid, a ValueError is raised.
        #
        # The value of empty cells is None.  Non-empty cells may contain a
        # value of str, decimal.Decimal, or CellError.
        #
        # Decimal values will not have trailing zeros to the right of any
        # decimal place, and will not include a decimal place if the value is a
        # whole number.  For example, this function would not return
        # Decimal('1.000'); rather it would return Decimal('1').
        if sheet_name.lower() not in self.sheets.keys():
            raise KeyError('Sheet not found.')
        
        if not self.is_valid_location(location):
            raise ValueError('Spreadsheet cell location is invalid. ZZZZ9999 is the bottom-right-most cell.') 
        
        sheet = self.sheets[sheet_name.lower()]
        col_idx, row_idx = Sheet.split_cell_ref(location)

        if col_idx >= sheet.num_cols or row_idx >= sheet.num_rows:
            raise ValueError('Location is beyond current extent of sheet.')
        
        cell = sheet.cells[row_idx][col_idx]

        ### This code below is copied from Cell.py (mostly)
        contents = cell.contents
        if (contents is None):
            return ""
        
        contents = contents.strip()
        if contents.startswith('='):
            # parse formula into tree
            parser = lark.Lark.open(lark_path, start='formula')
            tree = parser.parse(cell.contents)

            # obtain reference info from tree with visitor
            ref_info = self.get_cell_ref_info(tree, sheet_name)

            # feed references and sheet name into interpreter
            ev = FormulaEvaluator(sheet_name, ref_info)
            cell.value = ev.visit(tree)
        elif contents.startswith("'"):
            cell.value = contents[1:]
        else:
            if Cell.is_number(contents):
                cell.value = decimal.Decimal(contents)
            else:
                cell.value = contents
        return cell.value
    
    def get_cell_ref_info(self, tree, sheet_name):
        """
        Given a parsed formula tree and a sheet name, this finds the cell
        references in the tree, and stores their value into a map. Used by
        the FormulaEvaluator in get_cell_value.
        """
        finder = CellRefFinder()
        finder.visit(tree)

        info = {}
        for ref in finder.refs:
            # parse ref if necessary
            if ('!' in ref):
                curr_sheet_name = ref[:ref.index('!')]
                curr_location = ref[ref.index('!') + 1:]
            else:
                curr_sheet_name = sheet_name
                curr_location = ref
            info[ref] = self.get_cell_value(curr_sheet_name, curr_location)
        return info