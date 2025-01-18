import lark
from interpreter import FormulaEvaluator
import sheet

def main():
    parser = lark.Lark.open('formulas.lark', start='formula')
    ev = FormulaEvaluator()
    tree = parser.parse('=1 + D3')
    print(tree.pretty())
    print(ev.visit(tree))

if __name__ == "__main__":
    main()