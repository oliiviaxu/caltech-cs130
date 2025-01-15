# python -m unittest discover -s tests
import unittest
import sheets

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
        wb.set_cell_contents('Sheet1', 'D5', 'test')
        self.assertEqual(wb.get_sheet_extent('Sheet1'), (4, 5))
        wb.set_cell_contents('Sheet1', 'C6', 'test')
        self.assertEqual(wb.get_sheet_extent('Sheet1'), (4, 6)) # test adding rows
        wb.set_cell_contents('Sheet1', 'E4', '  test  ')
        self.assertEqual(wb.get_sheet_extent('Sheet1'), (5, 6)) # test adding columns

        with self.assertRaises(KeyError):
            wb.set_cell_contents('Sheet2', 'D5', 'test')
        with self.assertRaises(ValueError):
            wb.set_cell_contents('Sheet1', 'D5D5', 'test')
        
        self.assertEqual(wb.get_cell_contents('Sheet1', 'D5'), 'test') # test basic get
        self.assertEqual(wb.get_cell_contents('Sheet1', 'E4'), 'test') # test that trimming whitespace worked

if __name__ == "__main__":
    unittest.main()