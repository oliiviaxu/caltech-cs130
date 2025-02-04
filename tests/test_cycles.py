import unittest
import coverage
import sheets
import os
import lark
import decimal
import json
import contextlib
import cProfile

# TODO: With cycle-detection, you might try constructing large cycles that 
# contain many cells, or many small cycles, each containing a small number of cells. 

# TODO: You could try constructing another test where one cell is a part of many different cycles. 

# TODO: Your tests could repeatedly make and then break cycles, to really exercise the cycle-detection code.