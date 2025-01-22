import lark
from typing import Optional

class CellRefFinder(lark.Visitor):

    def __init__(self, sheet_name: Optional[str] = None):
        self.sheet_name = sheet_name
        self.refs = set()

    def cell(self, tree):
        # visitor method processes a parse tree node 
        if len(tree.children) == 1:
            self.refs.add(str(tree.children[0]).lower())
        elif len(tree.children) == 2:
            self.refs.add(str(tree.children[0]) + '!' + str(tree.children[1]).lower())
        else:
            assert False, 'Invalid formula. Format must be in ZZZZ9999.'