"""
Helper functions for creating dependencies between cells based on common structures.
"""

def create_chain(wb, sheet_name, num_cells_in_cycle):
    """
    Create a chain of cell dependencies in the specified sheet.
    Each cell depends on the next cell in the chain.
    :param wb: The workbook object.
    :param sheet_name: The name of the sheet.
    :param num_cells_in_cycle: The number of cells in the chain.
    :param last_cell_val: The value of the last cell in the chain.
    """
    for i in range(1, num_cells_in_cycle):
        wb.set_cell_contents(sheet_name, f'A{i}', f'=A{i+1}')
    # Set the last cell to a number
    wb.set_cell_contents(sheet_name, f'A{num_cells_in_cycle}', '1')

def create_chain_2(wb, sn_1, sn_2, num_cells_in_cycle):
    """
    Create a chain of cell dependencies between 2 sheets.
    :param wb: The workbook object.
    :param sheet_name: The name of the sheet.
    :param num_cells_in_cycle: The number of cells in the chain.
    :param last_cell_val: The value of the last cell in the chain.
    """
    for i in range(1, num_cells_in_cycle):
        wb.set_cell_contents(sn_1, f'A{i}', f'={sn_2}!A{i}')
        wb.set_cell_contents(sn_2, f'A{i}', f'={sn_1}!A{i + 1}')

    wb.set_cell_contents(sn_1, f'A{num_cells_in_cycle}', '1')

def create_web(wb, sheet_name, num_cells):
    """
    Create a web of cell dependencies in the specified sheet.
    Each cell in column B depends on the first cell in column A.
    :param wb: The workbook object.
    :param sheet_name: The name of the sheet.
    :param num_cells: The number of cells in the web.
    """
    wb.set_cell_contents(sheet_name, 'A1', '1')

    for i in range(1, num_cells):
        wb.set_cell_contents(sheet_name, f'B{i}', f'=A1 + 1')

def create_large_cycle(wb, num_cells_in_cycle):
    """
    Create a large cycle of cell dependencies in the specified sheet.
    Each cell depends on the next cell in the cycle, forming a loop.
    :param wb: The workbook object.
    :param num_cells_in_cycle: The number of cells in the cycle.
    """
    for i in range(1, num_cells_in_cycle):
        wb.set_cell_contents('Sheet1', f'A{i}', f'=A{i + 1}')
    wb.set_cell_contents('Sheet1', f'A{num_cells_in_cycle}', '=A1')

def create_small_cycles(wb, sheet_name, num_cycles):
    """
    Create multiple small cycles of cell dependencies in the specified sheet.
    Each cycle consists of two cells referencing each other.
    :param wb: The workbook object.
    :param sheet_name: The name of the sheet.
    :param num_cycles: The number of small cycles to create.
    """
    for i in range(1, num_cycles + 1):
        wb.set_cell_contents(sheet_name, f'A{i}', f'=B{i}')
        wb.set_cell_contents(sheet_name, f'B{i}', f'=A{i}')