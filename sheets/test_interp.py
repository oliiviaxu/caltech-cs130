import lark
from interpreter import FormulaEvaluator

def main():
    parser = lark.Lark.open('formulas.lark', start='formula')
    ev = FormulaEvaluator()
    tree = parser.parse('="aba" & "cadabra"')
    print(tree.pretty())
    print(ev.visit(tree))

if __name__ == "__main__":
    main()