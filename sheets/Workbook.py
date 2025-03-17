from __future__ import annotations
from .Sheet import Sheet
from .Cell import Cell
from .CellError import CellError, CellErrorType
from .CellValue import CellValue
from collections import OrderedDict, deque
from typing import List, Optional, Tuple, Any, Callable, Iterable, TextIO
import os
import lark
import json
from .DependencyGraph import DependencyGraph
from .transformer import SheetNameExtractor, FormulaUpdater
from .interpreter import FormulaEvaluator
from .SpreadsheetFunctions import create_function_directory
from .RowAdapter import RowAdapter
import decimal
import re
import copy

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
        self.notify_info = {}
        self.is_deleting = False
        self.is_renaming = False
        self.func_directory = create_function_directory(self)
        self.renaming_info = {}

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
    
    def handle_update_tree(self, cell_tuple):
        pending_notifications = []
        sheet_name, location = cell_tuple
        sheet_name = sheet_name.lower()
        location = location.lower()
        out_degree = self.calculate_out_degree((sheet_name, location))

        visited = {key: False for key in out_degree} # cells "connected" to src
        
        visited[(sheet_name, location)] = True
        queue = [(sheet_name, location)]
        first = True
        
        while len(queue):
            sheet_name, location = queue.pop(0)
            sheet_name = sheet_name.lower()
            location = location.lower()

            prev_value = self.get_cell_value(sheet_name, location)
            self.evaluate_cell((sheet_name, location), first)
            first = False
            new_value = self.get_cell_value(sheet_name, location)
            if (prev_value != new_value):
                if not (isinstance(prev_value, CellError) and isinstance(new_value, CellError) and prev_value.get_type() == new_value.get_type()):
                    pending_notifications.append((sheet_name, location))
                    if self.is_renaming:
                        if (sheet_name, location) not in self.notify_info:
                            self.notify_info[(sheet_name, location)] = prev_value

            visited[(sheet_name, location)] = True
            for sn, loc in self.graph.ingoing_get(sheet_name, location):
                # ingoing_cell = self.get_cell(sn, loc)
                if visited[(sn, loc)]:
                    continue
                out_degree[(sn, loc)] -= 1
                if (out_degree[(sn, loc)] == 0):
                    queue.append((sn, loc))
        
        for sn, loc in visited:
            if not visited[(sn, loc)]:
                prev_value = self.get_cell_value(sn, loc)
                self.evaluate_cell((sn, loc))
                new_value = self.get_cell_value(sn, loc)
                if (prev_value != new_value):
                    if not (isinstance(prev_value, CellError) and isinstance(new_value, CellError) and prev_value.get_type() == new_value.get_type()):
                        pending_notifications.append((sn, loc))
                        if self.is_renaming:
                            if (sn, loc) not in self.notify_info:
                                self.notify_info[(sn, loc)] = prev_value
        
        return pending_notifications
    
    def calculate_out_degree(self, cell_tup):
        sheet_name, location = cell_tup
        stack = [(sheet_name, location)]
        visited = set()
        out_degree = {}

        while stack:
            sheet_name, location = stack.pop()
            if (sheet_name, location) in visited:
                continue
            visited.add((sheet_name, location))
            for sn, loc in self.graph.ingoing_get(sheet_name, location):
                # ingoing_cell = self.get_cell(sn, loc)
                out_degree[(sn, loc)] = out_degree.get((sn, loc), 0) + 1
                stack.append((sn, loc))

        return out_degree
    
    def evaluate_cell(self, cell_tup, first=False):
        sheet_name, location = cell_tup
        sheet_name = sheet_name.lower()
        location = location.lower()
        cell = self.get_cell(sheet_name, location)
        if cell is None:
            return
        contents = cell.contents
        if (contents is None):
            cell.value = CellValue(None)
            return
        
        if contents.startswith('='):
            tree = cell.tree
            if cell.parse_error:
                cell.value = CellValue(CellError(CellErrorType.PARSE_ERROR, 'Failed to parse formula'))
            else:
                orig_outgoing = self.graph.outgoing_get(sheet_name, location)
                for sn, loc in orig_outgoing:
                    self.graph.ingoing_remove(sn, loc, sheet_name, location)

                # feed references and sheet name into interpreter
                ev = FormulaEvaluator(sheet_name, self, self.func_directory)
                visit_value = ev.visit(tree)
                if (visit_value is None or visit_value.val is None):
                    cell.value = CellValue(decimal.Decimal('0'))
                else:
                    cell.value = visit_value

                # update graph
                # if first:
                outgoing = list(ev.refs)
                for sn, loc in outgoing:
                    self.graph.ingoing_add(sn, loc, sheet_name, location)

                if len(outgoing):
                    self.graph.outgoing_set(sheet_name, location, outgoing)

                # detect cycle
                if first and self.detect_cycle((sheet_name, location)):
                    cell.value = CellValue(CellError(CellErrorType.CIRCULAR_REFERENCE, 'Circular reference found'))
                elif cell.in_cycle:
                    cell.value = CellValue(CellError(CellErrorType.CIRCULAR_REFERENCE, 'Circular reference found'))
                
        elif contents.startswith("'"):
            cell.value = CellValue(contents[1:])
        else:
            if CellValue.is_number(contents):
                contents = CellValue.strip_trailing_zeros(contents)
                cell.value = CellValue(decimal.Decimal(contents))
            elif contents.lower() in FormulaEvaluator.error_dict:
                cell.value = CellValue(CellError(FormulaEvaluator.error_dict[contents.lower()], 'String representation'))
            elif contents.lower() == 'true':
                cell.value = CellValue(True)
            elif contents.lower() == 'false':
                cell.value = CellValue(False)
            else:
                cell.value = CellValue(contents)

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

        nodes = set()
        self.find_nodes(sheet_name, location, nodes)
        for (sn, loc) in nodes:
            cell = self.get_cell(sn, loc)
            if cell is not None:
                cell.in_cycle = False

        if not self.is_deleting:
            orig_outgoing = self.graph.outgoing_get(sheet_name, location)
            for sn, loc in orig_outgoing:
                self.graph.ingoing_remove(sn, loc, sheet_name, location)

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
            if contents == curr_cell.contents:
                if not curr_cell.parse_error:
                    tree = curr_cell.tree
            else:
                try:
                    tree = parser.parse(contents)
                    curr_cell.tree = tree
                    curr_cell.parse_error = False
                except lark.exceptions.LarkError:
                    curr_cell.parse_error = True
        
        curr_cell.contents = contents
        self.graph.outgoing_reset(sheet_name, location)

        pending_notifications = []
        prev_value = self.get_cell_value(sheet_name, location)
        self.evaluate_cell((sheet_name, location), True)
        new_value = self.get_cell_value(sheet_name, location)
        if (prev_value != new_value):
            if not (isinstance(prev_value, CellError) and isinstance(new_value, CellError) and prev_value.get_type() == new_value.get_type()):
                pending_notifications.append((sheet_name, location))
                if self.is_renaming:
                    if (sheet_name, location) not in self.notify_info:
                        self.notify_info[(sheet_name, location)] = prev_value
        
        ### Update the value field of the cell
        pending_notifications = pending_notifications + self.handle_update_tree((sheet_name, location))
        if (not self.is_renaming and len(pending_notifications) > 0):
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

    def detect_cycle(self, cell_tup) -> bool:
        sheet_name, location = cell_tup
        sheet_name = sheet_name.lower()
        location = location.lower()
        nodes = set()
        self.find_nodes(sheet_name, location, nodes)

        current_id = 0
        ids = {node: -1 for node in nodes}
        low = {node: 0 for node in nodes}
        on_stack = {node: False for node in nodes}
        stack_scc = []
        is_cycle = {node: False for node in nodes}

        # pre-fetch adjacency to simplify lookups
        adjacency = {}
        for (sn, loc) in nodes:
            adjacency[(sn, loc)] = self.graph.ingoing_get(sn, loc)

        def iterative_dfs(start_node):
            nonlocal current_id, ids, low, on_stack, stack_scc, is_cycle, adjacency
            call_stack = [(start_node, 0, None)]

            while call_stack:
                node, child_idx, parent = call_stack.pop()
                sn, loc = node

                if ids[node] == -1:
                    current_id += 1
                    ids[node] = current_id
                    low[node] = current_id
                    on_stack[node] = True
                    stack_scc.append(node)

                if child_idx < len(adjacency[node]):
                    #pPut this frame back, but increment the child's index
                    call_stack.append((node, child_idx + 1, parent))

                    # process the next child
                    child = adjacency[node][child_idx]
                    if ids[child] == -1:
                        call_stack.append((child, 0, node))
                    elif on_stack[child]:
                        low[node] = min(low[node], ids[child])
                else:
                    # update parent
                    if parent is not None:
                        low[parent] = min(low[parent], low[node])

                    # found scc start node
                    if ids[node] == low[node]:
                        scc = []
                        while True:
                            top_node = stack_scc.pop()
                            on_stack[top_node] = False
                            scc.append(top_node)
                            low[top_node] = ids[node]
                            if top_node == node:
                                break

                        # if SCC size > 1, or a single node with a self-loop, mark as cycle
                        if (len(scc) > 1) or (len(scc) == 1 and node in adjacency[node]):
                            for comp_node in scc:
                                is_cycle[comp_node] = True

        for node in nodes:
            if ids[node] == -1:
                iterative_dfs(node)

        for node in nodes:
            cell = self.get_cell(node[0], node[1])
            cell.in_cycle = is_cycle[node]
        return is_cycle[(sheet_name, location)]

    def find_nodes(self, sn, loc, nodes):
        stack = [(sn, loc)]
        while len(stack):
            sn, loc = stack.pop()
            ref_id = (sn, loc)
            if ref_id in nodes:
                continue
            nodes.add(ref_id)
            ingoing = self.graph.ingoing_get(sn, loc)
            for next_sheet, next_loc in ingoing:
                stack.append((next_sheet, next_loc))

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
        if (cell is None or cell.value is None):
            return None
        return cell.value.val

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
        
        self.is_renaming = True

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
                if sn == new_sheet_name.lower():
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

        self.graph.ingoing.pop(sheet_name)
        self.graph.outgoing.pop(sheet_name)

        for loc in sheet_ingoings:
            cell_ingoings = sheet_ingoings[loc].copy()
            for sn, loc2 in cell_ingoings:
                if sn == new_sheet_name.lower():
                    cell = self.get_cell(sn, loc2)
                    if not cell.parse_error:
                        new_formula = sne.transform(cell.tree)
                        self.set_cell_contents(sn, loc2, '=' + new_formula)
        
        for loc in self.graph.ingoing[new_sheet_name.lower()]:
            self.set_cell_contents(new_sheet_name, loc, self.get_cell_contents(new_sheet_name, loc))

        self.is_renaming = False
        if len(self.notify_info):
            notifications = []
            for (sn, loc), v in self.notify_info.items():
                if sn in self.sheets:
                    cell = self.get_cell(sn, loc)
                    if cell is not None and cell.value.val != v:
                        notifications.append((sn, loc))
            for notify_function in self.notify_functions:
                try:
                    notify_function(self, notifications)
                except Exception:
                    pass
            self.notify_info = {}

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
        self.sheets[new_name.lower()] = copy.deepcopy(sheet_to_copy)
        self.sheets[new_name.lower()].sheet_name = new_name
        
        outgoings = self.graph.outgoing[sheet_name.lower()]
        for loc in outgoings:
            self.set_cell_contents(new_name, loc, self.get_cell_contents(sheet_name, loc))

        new_ingoings = self.graph.ingoing[new_name.lower()]
        for loc in new_ingoings:
            self.set_cell_contents(new_name, loc, self.get_cell_contents(sheet_name, loc))

        return (len(self.sheets.keys()) - 1, new_name)
    
    def transfer_cells(self, sheet_name: str, start_location: str,
            end_location: str, to_location: str, move: int, to_sheet: Optional[str] = None) -> None:
        if sheet_name.lower() not in self.sheets.keys() or \
        (to_sheet and (to_sheet.lower() not in self.sheets.keys())):
            raise KeyError('Sheet not found.')
        
        if (not Workbook.is_valid_location(start_location)) \
        or (not Workbook.is_valid_location(end_location)) \
        or (not Workbook.is_valid_location(to_location)):
            raise ValueError('Spreadsheet cell location is invalid. ZZZZ9999 is the bottom-right-most cell.')

        start_col, start_row = Sheet.split_cell_ref(start_location)
        end_col, end_row = Sheet.split_cell_ref(end_location)

        top_left_corner = (min(start_col, end_col), min(start_row, end_row)) # (col_idx, row_idx)
        bottom_right_corner = (max(start_col, end_col), max(start_row, end_row)) # (col_idx, row_idx)

        to_loc_col, to_loc_row = Sheet.split_cell_ref(to_location) # (col_idx, row_idx)

        new_bottom_right_col = to_loc_col + abs(end_col - start_col)
        new_bottom_right_row = to_loc_row + abs(end_row - start_row)

        # Check if the new bottom-right corner is valid
        if not Workbook.is_valid_location(Sheet.to_sheet_coords(new_bottom_right_col, new_bottom_right_row)):
            raise ValueError("Target area extends beyond the valid spreadsheet area.")

        m = bottom_right_corner[0] - top_left_corner[0] + 1
        n = bottom_right_corner[1] - top_left_corner[1] + 1

        contents_grid = [[0 for _ in range(m)] for _ in range(n)] # holds the updated contents

        # find delta_x and delta_y 
        delta_x = to_loc_col - top_left_corner[0] # change in column 
        delta_y = to_loc_row - top_left_corner[1] # change in row
        updater = FormulaUpdater(delta_x, delta_y) # column, row

        for i in range(n):
            for j in range(m):
                source_col = top_left_corner[0] + j
                source_row = top_left_corner[1] + i

                orig_loc = Sheet.to_sheet_coords(source_col, source_row)

                cell = self.get_cell(sheet_name, orig_loc)
                if cell:
                    if (cell.contents and cell.contents.startswith('=')):
                        new_formula = updater.transform(cell.tree)
                        contents_grid[i][j] = '=' + new_formula
                    else:
                        contents_grid[i][j] = cell.contents
                else:
                    contents_grid[i][j] = None # I'm not sure, if there wasn't a cell there, we're copying it over, right? 

                if move:
                    self.set_cell_contents(sheet_name, orig_loc, None)

        for i in range(to_loc_row, to_loc_row + n):
            for j in range(to_loc_col, to_loc_col + m):
                grid_i, grid_j = i - to_loc_row, j - to_loc_col
                updated_contents = contents_grid[grid_i][grid_j]

                loc = Sheet.to_sheet_coords(j, i)
                
                if to_sheet:
                    self.set_cell_contents(to_sheet, loc, updated_contents)
                else:
                    self.set_cell_contents(sheet_name, loc, updated_contents)

    
    def move_cells(self, sheet_name: str, start_location: str,
            end_location: str, to_location: str, to_sheet: Optional[str] = None) -> None:
        # Move cells from one location to another, possibly moving them to
        # another sheet.  All formulas in the area being moved will also have
        # all relative and mixed cell-references updated by the relative
        # distance each formula is being copied.
        #
        # Cells in the source area (that are not also in the target area) will
        # become empty due to the move operation.
        #
        # The start_location and end_location specify the corners of an area of
        # cells in the sheet to be moved.  The to_location specifies the
        # top-left corner of the target area to move the cells to.
        #
        # Both corners are included in the area being moved; for example,
        # copying cells A1-A3 to B1 would be done by passing
        # start_location="A1", end_location="A3", and to_location="B1".
        #
        # The start_location value does not necessarily have to be the top left
        # corner of the area to move, nor does the end_location value have to be
        # the bottom right corner of the area; they are simply two corners of
        # the area to move.
        #
        # This function works correctly even when the destination area overlaps
        # the source area.
        #
        # The sheet name matches are case-insensitive; the text must match but
        # the case does not have to.
        #
        # If to_sheet is None then the cells are being moved to another
        # location within the source sheet.
        #
        # If any specified sheet name is not found, a KeyError is raised.
        # If any cell location is invalid, a ValueError is raised.
        #
        # If the target area would extend outside the valid area of the
        # spreadsheet (i.e. beyond cell ZZZZ9999), a ValueError is raised, and
        # no changes are made to the spreadsheet.
        #
        # If a formula being moved contains a relative or mixed cell-reference
        # that will become invalid after updating the cell-reference, then the
        # cell-reference is replaced with a #REF! error-literal in the formula.
        self.transfer_cells(sheet_name, start_location, end_location, to_location, True, to_sheet)
        

    def copy_cells(self, sheet_name: str, start_location: str,
            end_location: str, to_location: str, to_sheet: Optional[str] = None) -> None:
        # Copy cells from one location to another, possibly copying them to
        # another sheet.  All formulas in the area being copied will also have
        # all relative and mixed cell-references updated by the relative
        # distance each formula is being copied.
        #
        # Cells in the source area (that are not also in the target area) are
        # left unchanged by the copy operation.
        #
        # The start_location and end_location specify the corners of an area of
        # cells in the sheet to be copied.  The to_location specifies the
        # top-left corner of the target area to copy the cells to.
        #
        # Both corners are included in the area being copied; for example,
        # copying cells A1-A3 to B1 would be done by passing
        # start_location="A1", end_location="A3", and to_location="B1".
        #
        # The start_location value does not necessarily have to be the top left
        # corner of the area to copy, nor does the end_location value have to be
        # the bottom right corner of the area; they are simply two corners of
        # the area to copy.
        #
        # This function works correctly even when the destination area overlaps
        # the source area.
        #
        # The sheet name matches are case-insensitive; the text must match but
        # the case does not have to.
        #
        # If to_sheet is None then the cells are being copied to another
        # location within the source sheet.
        #
        # If any specified sheet name is not found, a KeyError is raised.
        # If any cell location is invalid, a ValueError is raised.
        #
        # If the target area would extend outside the valid area of the
        # spreadsheet (i.e. beyond cell ZZZZ9999), a ValueError is raised, and
        # no changes are made to the spreadsheet.
        #
        # If a formula being copied contains a relative or mixed cell-reference
        # that will become invalid after updating the cell-reference, then the
        # cell-reference is replaced with a #REF! error-literal in the formula.
        self.transfer_cells(sheet_name, start_location, end_location, to_location, False, to_sheet)


    def sort_region(self, sheet_name: str, start_location: str, end_location: str, sort_cols: List[int]):
        # Sort the specified region of a spreadsheet with a stable sort, using
        # the specified columns for the comparison.
        #
        # The sheet name match is case-insensitive; the text must match but the
        # case does not have to.
        #
        # The start_location and end_location specify the corners of an area of
        # cells in the sheet to be sorted.  Both corners are included in the
        # area being sorted; for example, sorting the region including cells B3
        # to J12 would be done by specifying start_location="B3" and
        # end_location="J12".
        #
        # The start_location value does not necessarily have to be the top left
        # corner of the area to sort, nor does the end_location value have to be
        # the bottom right corner of the area; they are simply two corners of
        # the area to sort.
        #
        # The sort_cols argument specifies one or more columns to sort on.  Each
        # element in the list is the one-based index of a column in the region,
        # with 1 being the leftmost column in the region.  A column's index in
        # this list may be positive to sort in ascending order, or negative to
        # sort in descending order.  For example, to sort the region B3..J12 on
        # the first two columns, but with the second column in descending order,
        # one would specify sort_cols=[1, -2].
        #
        # The sorting implementation is a stable sort:  if two rows compare as
        # "equal" based on the sorting columns, then they will appear in the
        # final result in the same order as they are at the start.
        #
        # If multiple columns are specified, the behavior is as one would
        # expect:  the rows are ordered on the first column indicated in
        # sort_cols; when multiple rows have the same value for the first
        # column, they are then ordered on the second column indicated in
        # sort_cols; and so forth.
        #
        # No column may be specified twice in sort_cols; e.g. [1, 2, 1] or
        # [2, -2] are both invalid specifications.
        #
        # The sort_cols list may not be empty.  No index may be 0, or refer
        # beyond the right side of the region to be sorted.
        #
        # If the specified sheet name is not found, a KeyError is raised.
        # If any cell location is invalid, a ValueError is raised.
        # If the sort_cols list is invalid in any way, a ValueError is raised.
        if sheet_name.lower() not in self.sheets.keys():
            raise KeyError('Sheet not found.')
        
        if (not Workbook.is_valid_location(start_location)) \
        or (not Workbook.is_valid_location(end_location)):
            raise ValueError('Spreadsheet cell location is invalid. ZZZZ9999 is the bottom-right-most cell.')
        
        if len(sort_cols) == 0:
            raise ValueError('Column list cannot be empty.')
        
        start_col, start_row = Sheet.split_cell_ref(start_location)
        end_col, end_row = Sheet.split_cell_ref(end_location)

        top_left_col = min(start_col, end_col)
        top_left_row = min(start_row, end_row)
        bottom_right_col = max(start_col, end_col)
        bottom_right_row = max(start_row, end_row)

        seen = set()
        for index in sort_cols:
            if index == 0:
                raise ValueError('Column index cannot be 0.')
            if abs(index) in seen:
                raise ValueError('Column specified more than once in column list.')
            elif top_left_col + abs(index) - 1 > bottom_right_col:
                raise ValueError('Column specified is beyond the right side of the region to be sorted.')
            elif not isinstance(index, int):
                raise ValueError('Column index must be an integer.') # TODO: is this necessary?
            else:
                seen.add(index)
        
        # get row data
        m = bottom_right_col - top_left_col + 1
        n = bottom_right_row - top_left_row + 1

        adapters = []
        orig_cells = [[0 for _ in range(m)] for _ in range(n)]
        for i in range(n):
            row_data = []
            source_row = top_left_row + i

            for j in range(m):
                source_col = top_left_col + j

                orig_loc = Sheet.to_sheet_coords(source_col, source_row)

                cell = self.get_cell(sheet_name, orig_loc)
                orig_cells[i][j] = cell
                
                if cell:
                    # row_data.append(self.get_cell_value(sheet_name, orig_loc))
                    row_data.append(cell)
                else:
                    row_data.append(None)
            row_adapter = RowAdapter(source_row, row_data, sort_cols)
            adapters.append(row_adapter)

        sorted_adapters = sorted(adapters)

        contents_grid = [[0 for _ in range(m)] for _ in range(n)]

        for i in range(n):
            source_row = top_left_row + i
            target_row = sorted_adapters[i].row_idx

            delta_y = source_row - target_row
            sort_region = (top_left_col, top_left_row, bottom_right_col, bottom_right_row)
            updater = FormulaUpdater(0, delta_y, sort_region)

            for j in range(m):
                source_col = top_left_col + j 

                target_cell = orig_cells[target_row - top_left_row][j]

                if target_cell:

                    if target_cell.contents and target_cell.contents.startswith('='):
                        new_formula = updater.transform(target_cell.tree)
                        if new_formula: # is within region
                            contents_grid[i][j] = '=' + new_formula
                        else:
                            contents_grid[i][j] = target_cell.contents
                    else:
                        # Update the cell with non-formula contents
                        contents_grid[i][j] = target_cell.contents
                else:
                #     # Clear the cell if it's empty
                    contents_grid[i][j] = None
        
        # print(contents_grid)
                
        for i in range(n):
            source_row = top_left_row + i
            for j in range(m):
                updated_contents = contents_grid[i][j]
                source_col = top_left_col + j 

                target_loc = Sheet.to_sheet_coords(source_col, source_row)

                self.set_cell_contents(sheet_name, target_loc, updated_contents)
