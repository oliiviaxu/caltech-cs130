import decimal
import lark
# from lark.visitors import visit_children_decor

class FormulaEvaluator(lark.visitors.Interpreter):
    
    # @visit_children_decor
    def add_expr(self, values):
        if values[1] == '+':
            return values[0] + values[2]
        elif values[1] == '-':
            return values[0] - values[2]
        else:
            assert False, f'Unexpected operation: {values[1]}'

    def parens(self, tree):
        values = self.visit_children(tree)
        assert len(values) == 1, f'Unexpected tree {tree.pretty()}'
    
    def number(self, tree):
        # called when run into number node 
        return decimal.Decimal(tree.children)

    def string(self, tree):
        # called when run into a string node
        return tree.children[0].value[1:-1]
    
