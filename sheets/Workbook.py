from __future__ import annotations
from .Sheet import Sheet
from .Cell import Cell
from .CellError import CellError, CellErrorType
from .visitor import CellRefFinder
from collections import OrderedDict, deque
from typing import List, Optional, Tuple, Any, Callable, Iterable, TextIO
import os
import lark
import json
from .DependencyGraph import DependencyGraph
from .transformer import SheetNameExtractor
from .interpreter import FormulaEvaluator
import decimal
import re

current_dir = os.path.dirname(os.path.abspath(__file__))
lark_path = os.path.join(current_dir, "formulas.lark")
parser = lark.Lark.open(lark_path, start='formula')

class Workbook:
    # A workbook containing zero or more named spreadsheets.
    #
    # Any and all operations on a workbook that may affect calculated cell
    # values should cause the workbook's contents to be updated properly.

    def __init__(self):
        self.graph = DependencyGraph()
        self.sheets = OrderedDict() # lowercase keys
        self.notify_functions = []
        self.is_deleting = False

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

        if sheet_name is None:
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
            if (sheet_name == '' or '\'' in sheet_name or '\"' in sheet_name or sheet_name[0] in [' ', '\t', '\n'] or sheet_name[len(sheet_name)-1] in [' ', '\t', '\n']):
                raise ValueError('Spreadsheet names cannot start or end with whitespace characters, and they cannot be an empty string.')
            
            for char in sheet_name:
                if (char != ' ' and not char.isalnum() and char not in '.?!,:;!@#$%^&*()-_'):
                    raise ValueError('Spreadsheet name can be comprised of letters, numbers, spaces, and these punctuation characters: .?!,:;!@#$%^&*()-_')
        
            if (sheet_name.lower() in sheet_names_lower):
                raise ValueError('Spreadsheet names must be unique.')

        self.sheets[sheet_name.lower()] = Sheet(sheet_name)
        if sheet_name.lower() in self.graph.ingoing:
            for loc in self.graph.ingoing[sheet_name.lower()]:
                self.set_cell_contents(sheet_name, loc, None)
        self.graph.add_sheet(sheet_name.lower())
        return len(self.sheets.keys()) - 1, sheet_name

    def del_sheet(self, sheet_name: str) -> None:
        # Delete the spreadsheet with the specified name.
        #
        # The sheet name match is case-insensitive; the text must match but the
        # case does not have to.
        #
        # If the specified sheet name is not found, a KeyError is raised.

        sheet_name = sheet_name.lower()
        if (sheet_name not in self.sheets):
            raise KeyError(f'{sheet_name} not found, cannot delete.')

        sheet_graph_outgoing = self.graph.outgoing[sheet_name]
        for loc, outgoing_arr in sheet_graph_outgoing.items():
            for outgoing_sn, outgoing_loc in outgoing_arr:
                self.graph.ingoing_remove(outgoing_sn, outgoing_loc, sheet_name, loc)
        
        self.is_deleting = True
        for loc in self.graph.ingoing[sheet_name]:
            self.set_cell_contents(sheet_name, loc, '#ref!')
        
        del self.graph.outgoing[sheet_name]
        del self.sheets[sheet_name]
        self.is_deleting = False

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
        if (sheet.out_of_bounds(location)):
            return None
        return sheet.get_cell(location)
    
    def handle_update_tree(self, cell):
        pending_notifications = []

        out_degree = {}
        self.calculate_out_degree(cell, set(), out_degree)

        visited = {key: False for key in out_degree} # cells "connected" to src
        
        visited[cell] = True
        queue = [cell]
        
        while len(queue):
            curr_cell = queue.pop(0)

            prev_value = self.get_cell_value(curr_cell.sheet_name, curr_cell.location)
            self.evaluate_cell(curr_cell)
            new_value = self.get_cell_value(curr_cell.sheet_name, curr_cell.location)
            if (prev_value != new_value):
                pending_notifications.append((curr_cell.sheet_name, curr_cell.location))

            visited[curr_cell] = True
            for sn, loc in self.graph.ingoing_get(curr_cell.sheet_name, curr_cell.location):
                ingoing_cell = self.get_cell(sn, loc)
                if visited[ingoing_cell]:
                    continue
                out_degree[ingoing_cell] -= 1
                if (out_degree[ingoing_cell] == 0):
                    queue.append(ingoing_cell)
        
        for c in visited:
            if not visited[c]:
                self.evaluate_cell(c)
        
        return pending_notifications
    
    # def calculate_out_degree(self, cell, visited, out_degree):
    #     if (cell in visited):
    #         return
    #     visited.add(cell)
    #     for sn, loc in self.graph.ingoing_get(cell.sheet_name, cell.location):
    #         ingoing_cell = self.get_cell(sn, loc)
    #         out_degree[ingoing_cell] = out_degree.get(ingoing_cell, 0) + 1
    #         self.calculate_out_degree(ingoing_cell, visited, out_degree)
    
    def calculate_out_degree(self, cell, visited, out_degree):

        stack = deque([(cell, set())])

        while stack:
            curr_cell, visited_in_path = stack.pop()

            if curr_cell in visited_in_path:
                continue

            visited_in_path.add(curr_cell)
            visited.add(curr_cell)

            for sn, loc in self.graph.ingoing_get(curr_cell.sheet_name, curr_cell.location):
                ingoing_cell = self.get_cell(sn, loc)
                out_degree[ingoing_cell] = out_degree.get(ingoing_cell, 0) + 1
                stack.append((ingoing_cell, set(visited_in_path)))
    
    def evaluate_cell(self, cell):
        contents = cell.contents
        if (contents is None):
            cell.value = None
            return
        
        if contents.startswith('='):
            tree = cell.tree
            if cell.parse_error:
                cell.value = CellError(CellErrorType.PARSE_ERROR, 'Failed to parse formula')
            else:
                if self.detect_cycle(cell):
                    cell.value = CellError(CellErrorType.CIRCULAR_REFERENCE, 'Circular reference found')
                else:
                    # obtain reference info from tree with visitor
                    ref_info = self.get_cell_ref_info(tree, cell.sheet_name)

                    # feed references and sheet name into interpreter
                    ev = FormulaEvaluator(cell.sheet_name, ref_info)
                    visit_value = ev.visit(tree)
                    if (visit_value is None):
                        cell.value = decimal.Decimal('0')
                    else:
                        cell.value = visit_value
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

        if not self.is_deleting:
            orig_outgoing = self.graph.outgoing_get(sheet_name, location)
            for sn, loc in orig_outgoing:
                self.graph.ingoing_remove(sn, loc, sheet_name, location)

        outgoing = []
        if contents is not None:
            contents = contents.strip()
        if contents == '':
            contents = None
        
        # Only need to change cell.outgoing if a formula is used in the cell
        if contents is None:
            curr_cell.contents = contents
            curr_sheet.check_shrink(location)
        elif contents.startswith('='):
            # parse formula into tree
            parse_error = False
            if contents == curr_cell.contents:
                if curr_cell.parse_error:
                    parse_error = True
                else:
                    tree = curr_cell.tree
            else:
                try:
                    tree = parser.parse(contents)
                    curr_cell.tree = tree
                    curr_cell.parse_error = False
                except lark.exceptions.LarkError:
                    parse_error = True
                    curr_cell.parse_error = True

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
                    
                    if (len(ref_sheet_name) > 2 and ref_sheet_name[0] == "'" and ref_sheet_name[-1] == "'"):
                        ref_sheet_name = ref_sheet_name[1:-1]
                    
                    if (Workbook.is_valid_location(ref_location)):
                        outgoing.append((ref_sheet_name.lower(), ref_location.lower()))
        
        curr_cell.contents = contents
        for sn, loc in outgoing:
            self.graph.ingoing_add(sn, loc, sheet_name, location)

        self.graph.outgoing_set(sheet_name, location, outgoing)

        ### Update the value field of the cell
        pending_notifications = self.handle_update_tree(curr_cell)
        if (len(pending_notifications) > 0):
            if (self.is_deleting):
                pending_notifications = pending_notifications[1:]
            for notify_function in self.notify_functions:
                try:
                    notify_function(self, pending_notifications)
                except Exception:
                    pass

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
    
    def has_cycle(self, src) -> bool:
        """
        Perform DFS to detect cycles.
        :param node: The current cell node.
        :param visited: A set of visited cell locations.
        :return: True if a cycle is detected, False otherwise.
        """
        stack = deque([(src.sheet_name, src.location, set())])

        while stack:
            sheet_name, location, visited_in_path = stack.pop()
            ref_id = sheet_name + '!' + location

            if ref_id in visited_in_path:
                return True

            visited_in_path.add(ref_id)

            outgoing = self.graph.outgoing_get(sheet_name, location)
            for next_sheet, next_loc in outgoing:
                stack.append((next_sheet, next_loc, set(visited_in_path)))

        return False

    def detect_cycle(self, src: Cell) -> bool:
        """
        Detect cycles starting from the source cell.
        :param src: The source cell node.
        :return: True if a cycle is detected, False otherwise.
        """
        return self.has_cycle(src)

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
        cell = sheet.get_cell(location)
        if (cell is None):
            return None
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
                if (len(curr_sheet_name) > 2 and curr_sheet_name[0] == '\'' and curr_sheet_name[-1] == '\''):
                    curr_sheet_name = curr_sheet_name[1:-1]
                curr_location = ref[ref.index('!') + 1:]
            else:
                curr_sheet_name = sheet_name
                curr_location = ref

            try:
                ref_info[ref.lower()] = self.get_cell_value(curr_sheet_name, curr_location)
            except (ValueError, KeyError) as e:
                ref_info[ref.lower()] = CellError(CellErrorType.BAD_REFERENCE, 'Bad reference', e)
        return ref_info

    @staticmethod
    def load_workbook(fp: TextIO) -> Workbook:
        # This is a static method (not an instance method) to load a workbook
        # from a text file or file-like object in JSON format, and return the
        # new Workbook instance.  Note that the _caller_ of this function is
        # expected to have opened the file; this function merely reads the file.
        #
        # If the contents of the input cannot be parsed by the Python json
        # module then a json.JSONDecodeError should be raised by the method.
        # (Just let the json module's exceptions propagate through.)  Similarly,
        # if an IO read error occurs (unlikely but possible), let any raised
        # exception propagate through.
        #
        # If any expected value in the input JSON is missing (e.g. a sheet
        # object doesn't have the "cell-contents" key), raise a KeyError with
        # a suitably descriptive message.
        #
        # If any expected value in the input JSON is not of the proper type
        # (e.g. an object instead of a list, or a number instead of a string),
        # raise a TypeError with a suitably descriptive message.
        try:
            json_data = json.load(fp)
        except json.JSONDecodeError as e:
            raise e
        except IOError as e:
            raise e
        
        if (not isinstance(json_data, dict)):
            raise TypeError('JSON must be dictionary.')
        if ('sheets' not in json_data):
            raise KeyError('JSON is missing sheets key.')
        if (not isinstance(json_data['sheets'], list)):
            raise TypeError('Value corresponding to sheets key must be list.')
        
        sheets_data = json_data['sheets']
        wb = Workbook()
        for sheet_data in sheets_data:
            if ('name' not in sheet_data or 'cell-contents' not in sheet_data):
                raise KeyError('Sheet is missing necessary key(s) (must have name and cell-contents)')
            if (not isinstance(sheet_data['name'], str)):
                raise TypeError('Sheet name must be string.')
            if (not isinstance(sheet_data['cell-contents'], dict)):
                raise TypeError('Sheet cell contents must be dictionary.')
            
            wb.new_sheet(sheet_data['name'])
            for cell_location in sheet_data['cell-contents']:
                cell_contents = sheet_data['cell-contents'][cell_location]
                if (not isinstance(cell_location, str) or not isinstance(cell_contents, str)):
                    raise TypeError('Cell data is not strings.')
                wb.set_cell_contents(sheet_data['name'], cell_location, cell_contents)
        return wb

    def save_workbook(self, fp: TextIO) -> None:
        # TODO 
        # Instance method (not a static/class method) to save a workbook to a
        # text file or file-like object in JSON format.  Note that the _caller_
        # of this function is expected to have opened the file; this function
        # merely writes the file.
        #
        # If an IO write error occurs (unlikely but possible), let any raised
        # exception propagate through.
        try:
            sheet_list = []
            for sheet in self.sheets.values():
                cell_contents = {}
                for row_idx in range(sheet.num_rows):
                    for col_idx in range(sheet.num_cols):
                        curr_cell = sheet.cells[row_idx][col_idx]
                        loc = Sheet.to_sheet_coords(col_idx, row_idx).upper()
                        if curr_cell and curr_cell.contents is not None:
                            cell_contents[loc] = curr_cell.contents
                sheet_data = {"name": sheet.sheet_name, "cell-contents": cell_contents}
                sheet_list.append(sheet_data)
            workbook_data = {"sheets": sheet_list}
            json.dump(workbook_data, fp, indent=4)
        except Exception as e:
            raise e

    def notify_cells_changed(self,
            notify_function: Callable[[Workbook, Iterable[Tuple[str, str]]], None]) -> None:
        # Request that all changes to cell values in the workbook are reported
        # to the specified notify_function.  The values passed to the notify
        # function are the workbook, and an iterable of 2-tuples of strings,
        # of the form ([sheet name], [cell location]).  The notify_function is
        # expected not to return any value; any return-value will be ignored.
        #
        # Multiple notification functions may be registered on the workbook;
        # functions will be called in the order that they are registered.
        #
        # A given notification function may be registered more than once; it
        # will receive each notification as many times as it was registered.
        #
        # If the notify_function raises an exception while handling a
        # notification, this will not affect workbook calculation updates or
        # calls to other notification functions.
        #
        # A notification function is expected to not mutate the workbook or
        # iterable that it is passed to it.  If a notification function violates
        # this requirement, the behavior is undefined.
        self.notify_functions.append(notify_function)

    def update_cell_sn(self, sheet_name):
        sheet = self.sheets[sheet_name.lower()]
        for row_idx in range(sheet.num_rows):
            for col_idx in range(sheet.num_cols):
                curr_cell = sheet.cells[row_idx][col_idx]
                curr_cell.sheet_name = sheet_name

    def rename_sheet(self, sheet_name: str, new_sheet_name: str) -> None:
        # Rename the specified sheet to the new sheet name.  Additionally, all
        # cell formulas that referenced the original sheet name are updated to
        # reference the new sheet name (using the same case as the new sheet
        # name, and single-quotes iff [if and only if] necessary).
        #
        # The sheet_name match is case-insensitive; the text must match but the
        # case does not have to.
        #
        # As with new_sheet(), the case of the new_sheet_name is preserved by
        # the workbook.
        #
        # If the sheet_name is not found, a KeyError is raised.
        #
        # If the new_sheet_name is an empty string or is otherwise invalid, a
        # ValueError is raised.
        if (sheet_name.lower() not in self.sheets):
            raise KeyError(f'Sheet name {sheet_name} not found.')
        if (new_sheet_name == '' or '\'' in new_sheet_name or '\"' in new_sheet_name or new_sheet_name[0] in [' ', '\t', '\n'] or new_sheet_name[-1] in [' ', '\t', '\n']):
            raise ValueError('Spreadsheet names cannot start or end with whitespace characters, and they cannot be an empty string.')
        
        for char in new_sheet_name:
            if (char != ' ' and not char.isalnum() and char not in '.?!,:;!@#$%^&*()-_'):
                raise ValueError('Spreadsheet name can be comprised of letters, numbers, spaces, and these punctuation characters: .?!,:;!@#$%^&*()-_')
    
        if (new_sheet_name.lower() in self.sheets):
            raise ValueError('Spreadsheet names must be unique.')

        sne = SheetNameExtractor(sheet_name, new_sheet_name)
        
        sheet_name = sheet_name.lower()
        
        # copying old_sheet_name info in self.graph.ingoings, outgoings over to new_sheet_name
        self.graph.outgoing[new_sheet_name.lower()] = self.graph.outgoing[sheet_name]
        if (new_sheet_name.lower() in self.graph.ingoing):
            self.graph.ingoing[new_sheet_name.lower()] = {**self.graph.ingoing[new_sheet_name.lower()], **self.graph.ingoing[sheet_name]}
        else:
            self.graph.ingoing[new_sheet_name.lower()] = self.graph.ingoing[sheet_name]
        
        # resolve internal errors (when the references are within current sheet_name)
        sheet_ingoings = self.graph.ingoing[new_sheet_name.lower()]
        for loc in sheet_ingoings:
            cell_ingoings = sheet_ingoings[loc]
            for i in range(len(cell_ingoings)):
                sn, new_loc = cell_ingoings[i]
                if (sn == sheet_name):
                    cell_ingoings[i] = (new_sheet_name.lower(), new_loc)

        sheet_outgoings = self.graph.outgoing[new_sheet_name.lower()]
        for loc in sheet_outgoings:
            cell_outgoings = sheet_outgoings[loc]
            for i in range(len(cell_outgoings)):
                sn, new_loc = cell_outgoings[i]
                if (sn == sheet_name):
                    cell_outgoings[i] = (new_sheet_name.lower(), new_loc)

        # changing the sheet to be changed's ingoing cells' outgoing lists, 
        # which contain the sheet to be changed

        # this changes the contents, as well as the outgoing of the ingoings
        for loc in sheet_ingoings:
            cell_ingoings = sheet_ingoings[loc].copy()
            for sn, loc2 in cell_ingoings:
                if sn == new_sheet_name:
                    continue

                # update outgoing of this cell
                if self.get_cell_contents(sn, loc2).startswith('='):
                    cell = self.get_cell(sn, loc2)
                    if not cell.parse_error:
                        new_formula = sne.transform(cell.tree)
                        self.set_cell_contents(sn, loc2, '=' + new_formula)

        # changes the ingoing of the outgoings
        for loc in sheet_outgoings:
            cell_outgoings = sheet_outgoings[loc]
            for sn, loc2 in cell_outgoings:
                # access ingoing of this cell and modify
                ingoing_lst = self.graph.ingoing_get(sn, loc2)
                for i, (ingoing_sn, ingoing_loc) in enumerate(ingoing_lst):
                      if ingoing_sn == sheet_name:
                          ingoing_lst[i] = (new_sheet_name.lower(), ingoing_loc)
                                                                                                                                                                                    
        # update sheet name in self.sheets, self.graph.ingoing, self.graph.outgoing
        old_index = list(self.sheets.keys()).index(sheet_name.lower())
        self.sheets[new_sheet_name.lower()] = self.sheets.pop(sheet_name.lower()) # sheet object
        self.sheets[new_sheet_name.lower()].sheet_name = new_sheet_name
        self.move_sheet(new_sheet_name.lower(), old_index)
        
        self.update_cell_sn(new_sheet_name)     

        self.graph.ingoing.pop(sheet_name)
        self.graph.outgoing.pop(sheet_name)

        for loc in self.graph.ingoing[new_sheet_name.lower()]:
            self.set_cell_contents(new_sheet_name, loc, self.get_cell_contents(new_sheet_name, loc))

    def move_sheet(self, sheet_name: str, index: int) -> None:
        # Move the specified sheet to the specified index in the workbook's
        # ordered sequence of sheets.  The index can range from 0 to
        # workbook.num_sheets() - 1.  The index is interpreted as if the
        # specified sheet were removed from the list of sheets, and then
        # re-inserted at the specified index.
        #
        # The sheet name match is case-insensitive; the text must match but the
        # case does not have to.
        #
        # If the specified sheet name is not found, a KeyError is raised.
        #
        # If the index is outside the valid range, an IndexError is raised.
        
        if sheet_name.lower() not in self.sheets.keys():
            raise KeyError('Sheet not found.')
        
        if not (0 <= index < self.num_sheets()):
            raise IndexError('Index out of range.')

        sheets_list = list(self.sheets.items())

        sheet_to_move = sheets_list.pop(list(self.sheets.keys()).index((sheet_name.lower())))

        sheets_list.insert(index, sheet_to_move)

        self.sheets = OrderedDict(sheets_list)

    def copy_sheet(self, sheet_name: str) -> Tuple[int, str]:
        # Make a copy of the specified sheet, storing the copy at the end of the
        # workbook's sequence of sheets.  The copy's name is generated by
        # appending "_1", "_2", ... to the original sheet's name (preserving the
        # original sheet name's case), incrementing the number until a unique
        # name is found.  As usual, "uniqueness" is determined in a
        # case-insensitive manner.
        #
        # The sheet name match is case-insensitive; the text must match but the
        # case does not have to.
        #
        # The copy should be added to the end of the sequence of sheets in the
        # workbook.  Like new_sheet(), this function returns a tuple with two
        # elements:  (0-based index of copy in workbook, copy sheet name).  This
        # allows the function to report the new sheet's name and index in the
        # sequence of sheets.
        #
        # If the specified sheet name is not found, a KeyError is raised.
        if sheet_name.lower() not in self.sheets.keys():
            raise KeyError('Sheet not found.')
        
        sheet_names_lower = [sheet_name.lower() for sheet_name in self.list_sheets()]
        num = 1
        new_name = ""
        while True:
            new_name = sheet_name + '_' + str(num)
            if (new_name.lower() not in sheet_names_lower):
                break
            num += 1
        
        self.new_sheet(new_name)

        sheet_to_copy = self.sheets[sheet_name.lower()]
        copy = self.sheets[new_name.lower()]
        copy.resize_sheet(sheet_to_copy.num_rows, sheet_to_copy.num_cols)

        for row_idx in range(sheet_to_copy.num_rows):
            for col_idx in range(sheet_to_copy.num_cols):
                curr_contents = sheet_to_copy.cells[row_idx][col_idx].contents
                loc = Sheet.to_sheet_coords(col_idx, row_idx)
                self.set_cell_contents(new_name, loc, curr_contents)

        return (len(self.sheets.keys()) - 1, new_name)