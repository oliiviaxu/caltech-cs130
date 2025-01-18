import lark
from interpreter import FormulaEvaluator

def main():
    parser = lark.Lark.open('formulas.lark', start='formula')
    ev = FormulaEvaluator()
    tree = parser.parse('=(3)')
    ev.visit(tree)

if __name__ == "__main__":
    main()