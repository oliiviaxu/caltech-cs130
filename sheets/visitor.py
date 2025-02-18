import lark
from typing import Optional

class CellRefFinder(lark.Visitor):

    def __init__(self, sheet_name: Optional[str] = None):
        self.sheet_name = sheet_name
        self.refs = set()

    def cell(self, tree):
        # visitor method processes a parse tree node 
        if len(tree.children) == 1:
            location = str(tree.children[0]).lower().replace('$', '')
            self.refs.add(location)
        elif len(tree.children) == 2:
            sheet_name = str(tree.children[0])
            location = str(tree.children[1]).lower().replace('$', '')
            self.refs.add(sheet_name + '!' + location)
        else:
            raise AssertionError('Invalid formula. Format must be in ZZZZ9999.')