import lark
import re
from .Sheet import Sheet

class SheetNameExtractor(lark.visitors.Transformer):

    def __init__(self, sheet_name, new_sheet_name):
        self.sheet_name = sheet_name
        self.new_sheet_name = new_sheet_name

    @staticmethod
    def sheet_name_needs_quotes(sheet_name):
        pattern = r"^[A-Za-z_][A-Za-z0-9_]*$"
        return not bool(re.fullmatch(pattern, sheet_name))
    
    def mul_expr(self, tree):
        return str(tree[0]) + ' ' + str(tree[1]) + ' ' + str(tree[2]) 

    def add_expr(self, tree):
        return str(tree[0]) + ' ' + str(tree[1]) +  ' ' + str(tree[2]) 
    
    def unary_op(self, tree):
        return str(tree[0]) + str(tree[1])

    def concat_expr(self, tree):
        return str(tree[0]) + ' & ' + str(tree[1])

    def error(self, tree):
        return str(tree[0])

    def parens(self, tree):
        return '(' + str(tree[0]) + ')'
    
    def number(self, tree):
        return str(tree[0])

    def string(self, tree):
        return str(tree[0])
    
    def args(self, tree):
        return tree
    
    def function(self, tree):
        return str(tree[0].children[0]) + '(' + ', '.join(tree[0].children[1]) + ')'
    
    def cell(self, tree):
        # processes a parse tree node 
        if len(tree) == 1:
            return str(tree[0])
        if len(tree) == 2:
            curr_name = str(tree[0])
            # check if has quotes.
            # if it has quotes and doesn't need, then remove
            # if it has quotes and needs, then keep
            if (len(curr_name) > 2 and curr_name[0] == '\'' and curr_name[-1] == '\''):
                if (not SheetNameExtractor.sheet_name_needs_quotes(curr_name[1:-1])):
                    curr_name = curr_name[1:-1]
                    if curr_name.lower() == self.sheet_name.lower():
                        curr_name = self.new_sheet_name
                        if SheetNameExtractor.sheet_name_needs_quotes(curr_name):
                            curr_name = "'" + curr_name + "'"
                else:
                    # it needs quotes
                    if curr_name[1:-1].lower() == self.sheet_name.lower():
                        curr_name = self.new_sheet_name
                        if SheetNameExtractor.sheet_name_needs_quotes(curr_name):
                            curr_name = "'" + curr_name + "'"
            else:
                if curr_name.lower() == self.sheet_name.lower():
                    curr_name = self.new_sheet_name
                    if SheetNameExtractor.sheet_name_needs_quotes(curr_name):
                        curr_name = "'" + curr_name + "'"
            return curr_name + '!' + str(tree[1])
        else:
            raise AssertionError('Invalid formula. Format must be in ZZZZ9999.')

class FormulaUpdater(lark.visitors.Transformer):

    def __init__(self, delta_x, delta_y, sort_region=None):
        self.delta_x = delta_x
        self.delta_y = delta_y
        self.sort_region = sort_region
    
    def mul_expr(self, tree):
        return str(tree[0]) + ' ' + str(tree[1]) + ' ' + str(tree[2]) 

    def add_expr(self, tree):
        return str(tree[0]) + ' ' + str(tree[1]) +  ' ' + str(tree[2]) 
    
    def unary_op(self, tree):
        return str(tree[0]) + str(tree[1])

    def concat_expr(self, tree):
        return str(tree[0]) + ' & ' + str(tree[1])

    def error(self, tree):
        return str(tree[0])

    def parens(self, tree):
        return '(' + str(tree[0]) + ')'
    
    def number(self, tree):
        return str(tree[0])

    def string(self, tree):
        return str(tree[0])
    
    def boolean(self, tree):
        return str(tree[0])
    
    def args(self, tree):
        return tree
    
    def function(self, tree):
        return str(tree[0].children[0]) + '(' + ', '.join(tree[0].children[1]) + ')'
    
    def is_within_region(self, col_idx, row_idx):
        top_left_col, top_left_row, bottom_right_col, bottom_right_row = self.sort_region
        return (top_left_col <= col_idx <= bottom_right_col) and (top_left_row <= row_idx <= bottom_right_row)


    def update_contents(self, orig_location):
        if Sheet.is_col_mixed_ref(orig_location) and Sheet.is_row_mixed_ref(orig_location):
            return orig_location
        
        # if delta_y is outside the region, then dont transform
        
        col_idx, row_idx = Sheet.split_cell_ref(orig_location)
        max_col, max_row = Sheet.str_to_index('ZZZZ'), 9999

        within_region = True
        if self.sort_region:
            within_region = self.is_within_region(col_idx, row_idx)

        if within_region:
            if Sheet.is_row_mixed_ref(orig_location):
                new_col_idx = col_idx + self.delta_x
                if new_col_idx < 0 or new_col_idx >= max_col:
                    return '#REF!'

                new_col = Sheet.index_to_col(new_col_idx)
                return new_col + '$' + str(row_idx + 1)
            elif Sheet.is_col_mixed_ref(orig_location):
                new_row_idx = row_idx + self.delta_y
                if new_row_idx < 0 or new_row_idx >= max_row:
                    return '#REF!'
                
                new_row = new_row_idx + 1
                return '$' + Sheet.index_to_col(col_idx) + str(new_row)
            else:
                new_col_idx = col_idx + self.delta_x
                new_row_idx = row_idx + self.delta_y

                if new_col_idx < 0 or new_col_idx >= max_col or new_row_idx < 0 or new_row_idx >= max_row:
                    return '#REF!'

                return Sheet.to_sheet_coords(new_col_idx, new_row_idx)
        else:
            return ""

    def cell(self, tree):
        # processes a parse tree node 
        if len(tree) == 1: # cell ref with no sheetname, just location
            return self.update_contents(str(tree[0]))
        if len(tree) == 2: # cell ref with sheetname 
            curr_name = str(tree[0])
            formula = self.update_contents(str(tree[1]))
            return curr_name + '!' + formula
        else:
            raise AssertionError('Invalid formula. Format must be in ZZZZ9999.')