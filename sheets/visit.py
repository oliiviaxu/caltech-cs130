import lark

class CellRefFiner(lark.Visitor):
    def __init__(self, sheet_name: str):
        self.refs = set()
    
    def cell(self, tree):
        # print(tree.children)

        if len(tree.children) == 1:
            self.refs.add(str(tree.children[0]))
        elif len(tree.children) == 2:
            self.refs.add('!'.join(tree.children))
        else:
            assert False, 'input error message' # TODO