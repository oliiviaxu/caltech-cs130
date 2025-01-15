import lark

class CellRefFinder(lark.Visitor):
    # TODO: normalize all the cell references to the same case, convert them all to uppercase

    def __init__(self, sheet_name: str):
        self.sheet_name = sheet_name
        self.refs = set()

    def cell(self, tree):
        # print(tree.children)

        if len(tree.children) == 1:
            self.refs.add(str(tree.children[0]))
        elif len(tree.children) == 2:
            self.refs.add('!'.join(tree.children))
        else:
            assert False, 'input error message' # TODO