import decimal
import lark
from lark.visitors import visit_children_decor

class FormulaEvaluator(lark.visitors.Interpreter):
    def __init__(self, sheet_name, ref_info):
        self.sheet_name = sheet_name
        self.ref_info = ref_info

    error_dict = {
        "#ERROR!": 1,
        "#CIRCREF!": 2,
        "#REF!": 3,
        "#NAME?": 4,
        "#VALUE!": 5,
        "#DIV/0!": 6
    }
    
    @visit_children_decor
    def add_expr(self, values):
        if values[1] == '+':
            return values[0] + values[2]
        elif values[1] == '-':
            return values[0] - values[2]
        else:
            assert False, f'Unexpected operation: {values[1]}'
    
    @visit_children_decor
    def mul_expr(self, values):
        if values[1] == '*':
            return values[0] * values[2]
        elif values[1] == '/':
            return values[0] / values[2]
        else:
            assert False, f'Unexpected operation: {values[1]}'
    
    @visit_children_decor
    def unary_op(self, values):
        if values[0] == '+':
            return abs(values[1])
        elif values[0] == '-':
            return -abs(values[1])
        else:
            assert False, f'Unexpected operation: {values[0]}'

    @visit_children_decor
    def concat_expr(self, values):
        assert len(values) == 2, 'Unexpected number of args'
        return values[0] + values[1]

    def error(self, tree):
        error_value = FormulaEvaluator.error_dict[self.children[0]]
        return error_value

    def parens(self, tree):
        values = self.visit_children(tree)
        assert len(values) == 1, f'Unexpected tree {tree.pretty()}'
        return values[0]
    
    def number(self, tree):
        # called when run into number node 
        return decimal.Decimal(tree.children[0])

    def string(self, tree):
        # called when run into a string node
        return tree.children[0].value[1:-1]
    
    def cell(self, tree):
        # first parse the value into sheet (if given) and location
        reference = tree.children[0].value.upper()
        assert reference in self.ref_info, f'Could not find cell information for reference {reference}'
        return self.ref_info[reference]