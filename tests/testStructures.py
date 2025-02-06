# collection of helper functions for creating dependencies between cells
# based on common structures

def create_chain(wb, sheet_name, num_cells_in_cycle, last_cell_val):
    for i in range(1, num_cells_in_cycle):
        wb.set_cell_contents(sheet_name, f'A{i}', f'=A{i+1} + {1}')
    # Set the last cell to a number
    wb.set_cell_contents(sheet_name, f'A{num_cells_in_cycle}', str(last_cell_val))

def create_web(wb, sheet_name, num_cells):
    wb.set_cell_contents(sheet_name, 'A1', '1')

    for i in range(1, num_cells):
        wb.set_cell_contents(sheet_name, f'B{i}', f'=A1 + 1')

def create_large_cycle(wb, num_cells_in_cycle):
    for i in range(1, num_cells_in_cycle):
        wb.set_cell_contents('Sheet1', f'A{i}', f'=A{i + 1}')
    wb.set_cell_contents('Sheet1', f'A{num_cells_in_cycle}', '=A1')

def create_small_cycles(wb, sheet_name, num_cycles):
    # this creates a cycle of 2 cells
    for i in range(1, num_cycles + 1):
        wb.set_cell_contents(sheet_name, f'A{i}', f'=B{i}')
        wb.set_cell_contents(sheet_name, f'B{i}', f'=A{i}')