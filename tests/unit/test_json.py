import unittest
import coverage
import sheets
import os
import decimal
import json

current_dir = os.path.dirname(os.path.abspath(__file__))
json_dir = os.path.join(current_dir, 'json_workbooks/')

class JsonTests(unittest.TestCase):
    def test_load_workbook(self):
        file_path = os.path.join(json_dir, 'simple.json')
        with open(file_path, "r") as file:
            wb = sheets.Workbook.load_workbook(file)
            self.assertEqual(wb.get_sheet_extent('Sheet1'), (4, 1))
            self.assertEqual(wb.get_cell_contents('Sheet1', 'A1'), '\'123')
            self.assertEqual(wb.get_cell_contents('Sheet1', 'B1'), '5.3')
            self.assertEqual(wb.get_cell_contents('Sheet1', 'C1'), '=A1*B1')
            self.assertEqual(wb.get_cell_value('Sheet1', 'C1'), decimal.Decimal('651.9'))
            self.assertEqual(wb.get_cell_contents('Sheet1', 'D1'), '=\"double\" & \" quote\"')
            self.assertEqual(wb.get_cell_value('Sheet1', 'D1'), 'double quote')

            self.assertEqual(wb.get_sheet_extent('Sheet2'), (3, 1))
            self.assertEqual(wb.get_cell_contents('Sheet2', 'A1'), '4.3')
            self.assertEqual(wb.get_cell_contents('Sheet2', 'B1'), '\'testing')
            self.assertEqual(wb.get_cell_contents('Sheet2', 'C1'), '=A1&B1')
            self.assertEqual(wb.get_cell_value('Sheet2', 'C1'), '4.3testing')
        
        # check with txt
        file_path = os.path.join(json_dir, 'simple.txt')
        with open(file_path, "r") as file:
            wb = sheets.Workbook.load_workbook(file)
            self.assertEqual(wb.get_sheet_extent('Sheet1'), (4, 1))
            self.assertEqual(wb.get_cell_contents('Sheet1', 'A1'), '\'123')
            self.assertEqual(wb.get_cell_contents('Sheet1', 'B1'), '5.3')
            self.assertEqual(wb.get_cell_contents('Sheet1', 'C1'), '=A1*B1')
            self.assertEqual(wb.get_cell_value('Sheet1', 'C1'), decimal.Decimal('651.9'))
            self.assertEqual(wb.get_cell_contents('Sheet1', 'D1'), '=\"double\" & \" quote\"')
            self.assertEqual(wb.get_cell_value('Sheet1', 'D1'), 'double quote')

            self.assertEqual(wb.get_sheet_extent('Sheet2'), (3, 1))
            self.assertEqual(wb.get_cell_contents('Sheet2', 'A1'), '4.3')
            self.assertEqual(wb.get_cell_contents('Sheet2', 'B1'), '\'testing')
            self.assertEqual(wb.get_cell_contents('Sheet2', 'C1'), '=A1&B1')
            self.assertEqual(wb.get_cell_value('Sheet2', 'C1'), '4.3testing')
        
        # test that appropriate errors are raised
        type_error_files = ['fail_list.json', 'fail_sheets_not_list.json', 'fail_name_not_str.json', 'fail_contents_not_dict.json', 'fail_cell_data_not_str.json']
        key_error_files = ['fail_missing_sheets.json', 'fail_missing_name.json']
        for type_error_file in type_error_files:
            file_path = os.path.join(json_dir, type_error_file)
            with open(file_path, "r") as file:
                with self.assertRaises(TypeError):
                    sheets.Workbook.load_workbook(file)

        for key_error_file in key_error_files:
            file_path = os.path.join(json_dir, key_error_file)
            with open(file_path, "r") as file:
                with self.assertRaises(KeyError):
                    sheets.Workbook.load_workbook(file)

    def test_save_with_string_formulas(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        wb.new_sheet('MySheet')

        wb.set_cell_contents('Sheet1', 'A1', '=B1 + C1')
        wb.set_cell_contents('Sheet1', 'B1', '=\'3')
        wb.set_cell_contents('MySheet', 'A1', '=Sheet1!A1 + 2')
        wb.set_cell_contents('MySheet', 'B1', '="testing" & " double"')

        file_path = os.path.join(json_dir, 'writing.json')

        with open(file_path, "w") as f:
            wb.save_workbook(f)

        with open(file_path, "r") as file:
            d = json.load(file)
            sheet_dic = d["sheets"]
            self.assertEqual(sheet_dic[0]["name"], "Sheet1")
            self.assertEqual(sheet_dic[0]["cell-contents"]["A1"], "=B1 + C1")
            self.assertEqual(sheet_dic[0]["cell-contents"]["B1"], "=\'3")

            self.assertEqual(sheet_dic[1]["name"], "MySheet")
            self.assertEqual(sheet_dic[1]["cell-contents"]["A1"], "=Sheet1!A1 + 2")
            self.assertEqual(sheet_dic[1]["cell-contents"]["B1"], "=\"testing\" & \" double\"")
        
        empty_wb = sheets.Workbook()
        with open(file_path, "w") as f:
            empty_wb.save_workbook(f)
        
        with open(file_path, "r") as file:
            d = json.load(file)
            self.assertEqual(d["sheets"], [])
        
        empty_cell_wb = sheets.Workbook()
        empty_cell_wb.new_sheet()
        with open(file_path, "w") as f:
            empty_cell_wb.save_workbook(f)
        
        with open(file_path, "r") as file:
            d = json.load(file)
            self.assertEqual(len(d["sheets"]), 1)
            self.assertEqual(d["sheets"][0]["name"], "Sheet1")
            self.assertEqual(d["sheets"][0]["cell-contents"], {})

if __name__ == "__main__":
    cov = coverage.Coverage()
    cov.start()
    unittest.main()
    cov.stop()
    cov.save()
    cov.html_report()