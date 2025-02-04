import unittest
import coverage
import sheets
import os
import lark
import decimal
import json
import contextlib
import cProfile

# TODO: You might create a test that requires updates to propagate through 
# long chains of cell references, where each cell depends only on one other cell. 

# TODO:Alternately, you might also have updates where each cell is referenced by many other cells, 
# perhaps with much shallower chains, but still with large amounts of cell updates. 
# Can you come up with a general way of constructing such tests, and use it to exercise 
# several different scenarios?

# TODO: How do you trigger many updates with only one cell-change, so that you are maximally 
# exercising code inside your library, rather than code in the test?