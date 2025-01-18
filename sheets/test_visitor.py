import lark
from visitor import CellRefFinder

def main():
    f = CellRefFinder()
    parser = lark.Lark.open('formulas.lark', start='formula')
    tree = parser.parse('=a4*3 + Sheet1!a1')
    f.visit(tree)
    print(f.refs)

if __name__ == "__main__":
    main()