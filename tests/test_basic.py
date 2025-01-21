# python -m unittest discover -s tests
import unittest
import sheets
import os
import sheets.Sheet as Sheet
import lark

current_dir = os.path.dirname(os.path.abspath(__file__))
lark_path = os.path.join(current_dir, "../sheets/formulas.lark")

class BasicTests(unittest.TestCase):
    def test_new_sheet(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.new_sheet('Lorem ipsum')
        wb.new_sheet('Sheet3')
        wb.new_sheet('sheet2')
        wb.new_sheet('Sheet5')
        wb.new_sheet()
        self.assertEqual(wb.list_sheets(), ['Sheet1', 'Lorem ipsum', 'Sheet3', 'sheet2', 'Sheet5', 'Sheet4'])

        with self.assertRaises(ValueError):
            wb.new_sheet('')
        with self.assertRaises(ValueError):
            wb.new_sheet(' Sheet')
        with self.assertRaises(ValueError):
            wb.new_sheet('~')
    
    def test_set_cell_contents(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.set_cell_contents('Sheet1', 'AA26', 'test')
        self.assertEqual(wb.get_sheet_extent('Sheet1'), (27, 26))
        wb.set_cell_contents('Sheet1', 'C27', 'test')
        self.assertEqual(wb.get_sheet_extent('Sheet1'), (27, 27)) # test adding rows
        wb.set_cell_contents('Sheet1', 'AB4', '  test  ')
        self.assertEqual(wb.get_sheet_extent('Sheet1'), (28, 27)) # test adding columns

        with self.assertRaises(KeyError):
            wb.set_cell_contents('Sheet2', 'D5', 'test')
        with self.assertRaises(ValueError):
            wb.set_cell_contents('Sheet1', 'D5D5', 'test')
        
        self.assertEqual(wb.get_cell_contents('Sheet1', 'AA26'), 'test') # test basic get
        self.assertEqual(wb.get_cell_contents('Sheet1', 'AB4'), 'test') # test that trimming whitespace worked

        wb.set_cell_contents('Sheet1', 'A1', '\'string')
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1'), 'string')

    def test_formula_evaluation(self):
        wb = sheets.Workbook()
        wb.new_sheet()

        wb.set_cell_contents('Sheet1', 'A1', '=1 + 2 * 3')
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1'), 7)

        wb.set_cell_contents('Sheet1', 'A1', '="aba" & "cadabra"')
        self.assertEqual(wb.get_cell_value('Sheet1', 'A1'), 'abacadabra')
    
    def test_cell_reference(self):
        wb = sheets.Workbook()
        wb.new_sheet()

        wb.set_cell_contents('Sheet1', 'D2', '2')
        wb.set_cell_contents('Sheet1', 'D3', '=1 + D2')

        col_idx, row_idx = Sheet.split_cell_ref('D3')
        # print(wb.sheets['sheet1'].cells[row_idx][col_idx].outgoing[0].location)

        self.assertEqual(wb.sheets['sheet1'].cells[row_idx][col_idx].outgoing[0].location, 'D2')

        self.assertEqual(wb.get_cell_value('Sheet1', 'D3'), 3)
    
    def test_ingoing_outgoing(self):
        wb = sheets.Workbook()
        wb.new_sheet()

        wb.set_cell_contents('Sheet1', 'D1', '2')
        wb.set_cell_contents('Sheet1', 'D2', '=D1+3')
        wb.set_cell_contents('Sheet1', 'D3', '=D1+4')

        d_1 = wb.sheets['sheet1'].cells[0][3]
        d_2 = wb.sheets['sheet1'].cells[1][3]
        d_3 = wb.sheets['sheet1'].cells[2][3]

        # def print_lists(cell):
        #     print(f'########## PRINTING {cell.location} LIST ############')
        #     print('Ingoing: ')
        #     for c in cell.ingoing:
        #         print(f'{c.location}')
            
        #     print('Outgoing: ')
        #     for c in cell.outgoing:
        #         print(f'{c.location}')
        
        # print_lists(d_1)
        # print_lists(d_2)
        # print_lists(d_3)
        
        wb.set_cell_contents('Sheet1', 'D2', '=1')

        d_1 = wb.sheets['sheet1'].cells[0][3]
        d_2 = wb.sheets['sheet1'].cells[1][3]
        d_3 = wb.sheets['sheet1'].cells[2][3]

        # print_lists(d_1)
        # print_lists(d_2)
        # print_lists(d_3)
    
    def test_random(self):
        wb = sheets.Workbook()
        wb.new_sheet()

        wb.set_cell_contents('Sheet1', 'A1', '4')
        # print(wb.get_cell_value('Sheet1', 'A1'), wb.get_cell_value('Sheet1', 'A2'))

        wb.set_cell_contents('Sheet1', 'A1', '0')
        # print(wb.get_cell_value('Sheet1', 'A1'), wb.sheets['sheet1'].cells[1][0].value)

if __name__ == "__main__":
    unittest.main()