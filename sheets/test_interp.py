import lark
from sheets.interpretor import FormulaEvaluator

def main():
    ev = FormulaEvaluator()
    parser = lark.Lark.open('formulas.lark', start='formula')
    tree = parser.parse('=3')
    ev.visit(tree)

if __name__ == "__main__":
    main()