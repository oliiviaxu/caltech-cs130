import unittest
import coverage
import sheets
import os
import lark
import decimal
import json
import contextlib
import cProfile

class CycleDetectionTests(unittest.TestCase):

    def test_large_cycle(self):
        wb = sheets.Workbook()
        wb.new_sheet()
        
        for i in range(1, 100):
            wb.set_cell_contents('Sheet1', f'A{i}', f'=A{i + 1}')

        wb.set_cell_contents('Sheet1', 'A100', '=A1')
        cell = wb.sheets['sheet1'].cells[0][0]
        self.assertEqual(wb.detect_cycle(cell), True)        

# TODO: With cycle-detection, you might try constructing large cycles that 
# contain many cells, or many small cycles, each containing a small number of cells. 

# TODO: You could try constructing another test where one cell is a part of many different cycles. 

# TODO: Your tests could repeatedly make and then break cycles, to really exercise the cycle-detection code.

if __name__ == "__main__":
    cov = coverage.Coverage()
    cov.start()
    unittest.main()
    cov.stop()
    cov.save()
    cov.html_report()