# A function to find the next empty cell
def nextEmptyCell(ws, row, column):
    next_cell = ws.cell(row=row, column=column)
    while next_cell.value is not None:
        next_cell = ws.cell(row=row, column=next_cell.column+1)
    
    return next_cell

def sumRow(ws):

    row_values = [cell.value if cell.value is not None else 0 for cell in ws[2]]
    row_sum = sum(row_values)
      
    return row_sum