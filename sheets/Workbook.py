from .Sheet import *
from .Cell import *
from .CellError import CellError, CellErrorType
from .visitor import CellRefFinder
from typing import List, Optional, Tuple, Any, Set
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

        sheet_to_delete = self.sheets[sheet_name.lower()]
        cells = sheet_to_delete.cells
        
        del self.sheets[sheet_name.lower()]

        for row_idx in range(sheet_to_delete.num_rows):
            for col_idx in range(sheet_to_delete.num_cols):
                curr_cell = cells[row_idx][col_idx]

                for outgoing_cell in curr_cell.outgoing:
                    outgoing_cell.ingoing.remove(curr_cell)
                
                # error propagation
                visited = set()
                for ingoing_cell in curr_cell.ingoing:
                    self.update_cell(ingoing_cell, visited)

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

    def is_valid_location(location: str) -> bool:
        # Checks if a given location string is a valid spreadsheet cell location.
        pattern = r'^[A-Za-z]{1,4}[1-9][0-9]{0,3}$'
        return bool(re.match(pattern, location))
    
    def get_cell(self, sheet_name, location):
        if sheet_name.lower() not in self.sheets.keys():
            raise KeyError('Sheet not found.')
        
        if not Workbook.is_valid_location(location):
            raise ValueError('Spreadsheet cell location is invalid. ZZZZ9999 is the bottom-right-most cell.')
        
        sheet = self.sheets[sheet_name.lower()]
        sheet.resize(location)

        return sheet.get_cell(location)
    
    def update_cell(self, cell, visited):
        if (cell in visited):
            return
        visited.add(cell)
        self.evaluate_cell(cell)
        for ingoing_cell in cell.ingoing:
            self.update_cell(ingoing_cell, visited)
    
    def evaluate_cell(self, cell):
        contents = cell.contents
        if (contents is None):
            cell.value = None
            return
        
        if contents.startswith('='):
            # parse formula into tree
            parser = lark.Lark.open(lark_path, start='formula')
            parse_error = False
            try:
                tree = parser.parse(cell.contents)
            except:
                parse_error = True

            if parse_error:
                cell.value = CellError(CellErrorType.PARSE_ERROR, 'Failed to parse formula')
            else:
                if self.detect_cycle(cell):
                    cell.value = CellError(CellErrorType.CIRCULAR_REFERENCE, 'Circular reference found')
                else:
                    # obtain reference info from tree with visitor
                    ref_info = self.get_cell_ref_info(tree, cell.sheet_name)

                    # feed references and sheet name into interpreter
                    ev = FormulaEvaluator(cell.sheet_name, ref_info)
                    cell.value = ev.visit(tree)
        elif contents.startswith("'"):
            cell.value = contents[1:]
        else:
            if Cell.is_number(contents):
                contents = Cell.strip_trailing_zeros(contents)
                cell.value = decimal.Decimal(contents)
            elif contents.lower() in FormulaEvaluator.error_dict:
                cell.value = CellError(FormulaEvaluator.error_dict[contents.lower()], 'String representation')
            else:
                cell.value = contents

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
        
        if not Workbook.is_valid_location(location):
            raise ValueError('Spreadsheet cell location is invalid. ZZZZ9999 is the bottom-right-most cell.') 
        
        # remove original outgoing cells' ingoing & outgoing lists before setting new content
        curr_sheet = self.sheets[sheet_name.lower()]
        curr_sheet.resize(location)
        curr_cell = curr_sheet.get_cell(location)
        orig_outgoing = curr_cell.outgoing

        for orig_ref_cell in orig_outgoing:
            orig_ref_cell.ingoing.remove(curr_cell)

        outgoing = []
        if contents == '' or contents == None:
            contents = None
        else:
            contents = contents.strip()

        # Only need to change cell.outgoing if a formula is used in the cell
        if contents is not None and contents.startswith('='):
            # parse formula into tree
            parser = lark.Lark.open(lark_path, start='formula')
            parse_error = False
            try:
                tree = parser.parse(contents)
            except:
                parse_error = True

            if not parse_error:
                # Obtain references
                finder = CellRefFinder()
                finder.visit(tree) 
                
                for ref in finder.refs:
                    if '!' in ref:
                        # get the referenced cells
                        split_ref = ref.split('!')
                        ref_sheet_name = split_ref[0]
                        ref_location = split_ref[1]
                    else:
                        ref_sheet_name = sheet_name
                        ref_location = ref
                    
                    if (ref_sheet_name.lower() in self.sheets.keys() and Workbook.is_valid_location(ref_location)):
                        self.sheets[ref_sheet_name.lower()].resize(ref_location)
                        referenced_cell = self.sheets[ref_sheet_name.lower()].get_cell(ref_location)
                        outgoing.append(referenced_cell)
            
        for referenced_cell in outgoing:
            referenced_cell.ingoing.append(curr_cell)   

        curr_cell.outgoing = outgoing
        curr_cell.contents = contents

        ### Update the value field of the cell
        visited = set()
        self.update_cell(curr_cell, visited)

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
        
        if not Workbook.is_valid_location(location):
            raise ValueError('Spreadsheet cell location is invalid. ZZZZ9999 is the bottom-right-most cell.')
        
        return self.sheets[sheet_name.lower()].get_cell_contents(location)
    
    def dfs(self, node: Cell, visited: Set[str]) -> bool:
        """
        Perform DFS to detect cycles.
        :param node: The current cell node.
        :param visited: A set of visited cell locations.
        :return: True if a cycle is detected, False otherwise.
        """
        has_cycle = False
        for ref in node.outgoing:
            ref_id = ref.sheet_name + '!' + ref.location
            if (ref_id in visited):
                return True
            visited.add(ref_id)
            has_cycle = has_cycle or self.dfs(ref, visited)
            visited.remove(ref_id)
        return has_cycle

    def detect_cycle(self, src: Cell) -> bool:
        """
        Detect cycles starting from the source cell.
        :param src: The source cell node.
        :return: True if a cycle is detected, False otherwise.
        """

        visited = set()
        id = src.sheet_name + '!' + src.location
        visited.add(id)
        return self.dfs(src, visited)

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
        
        if not Workbook.is_valid_location(location):
            raise ValueError('Spreadsheet cell location is invalid. ZZZZ9999 is the bottom-right-most cell.') 
        
        sheet = self.sheets[sheet_name.lower()]
        sheet.resize(location)
        cell = sheet.get_cell(location)
        return cell.value
    
    def get_cell_ref_info(self, tree, sheet_name):
        """
        Given a parsed formula tree and a sheet name, this finds the cell
        references in the tree, and stores their value into a map. Used by
        the FormulaEvaluator in get_cell_value.
        """
        finder = CellRefFinder()
        finder.visit(tree)

        ref_info = {}
        for ref in finder.refs:
            # parse ref if necessary
            if ('!' in ref):
                curr_sheet_name = ref[:ref.index('!')]
                curr_location = ref[ref.index('!') + 1:]
            else:
                curr_sheet_name = sheet_name
                curr_location = ref

            try:
                ref_info[ref.lower()] = self.get_cell_value(curr_sheet_name, curr_location)
            except (ValueError, KeyError) as e:
                ref_info[ref.lower()] = CellError(CellErrorType.BAD_REFERENCE, 'Bad reference', e)
        return ref_info