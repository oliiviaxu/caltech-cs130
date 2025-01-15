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
        

if __name__ == "__main__":
    unittest.main()