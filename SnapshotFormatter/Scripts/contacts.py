from snapshot_structures import *
from colors import Color, pattern_fill


# Highlight missing Cell and email info for legal guardians and address info in Contact Info tab
school_addresses = ['1100 e 111', '1100 east 111', '9900 blank']

def highlight_missing_contact(book):
    contac = book['Contact Info']
    # Check that legal guardians have cell or email on file
    for row in contac.iter_rows(min_row=2, max_row=contac.max_row, min_col=7, max_col=7):
        for cell in row:
            if cell.value == 'Yes':
                if cell.offset(column=3).value is None and cell.offset(column=4).value is None:
                    pattern_fill(cell.offset(column=3), color=Color.YELLOW)
                    pattern_fill(cell.offset(column=4), color=Color.YELLOW)
                    
    # Highlight students with no contacts
    for row in contac.iter_rows(min_row=2, max_row=contac.max_row, min_col=5, max_col=5):
        for cell in row:
            if cell.value is None:
                for x in range(7):
                    pattern_fill(cell.offset(column=x), color=Color.YELLOW)
    
    # Check that parents & guardians aren't missing legal guardianship info       
    for row in contac.iter_rows(min_row=2, max_row=contac.max_row, min_col=6, max_col=6):
        for cell in row:
            if cell.value == 'Parent' or cell.value == 'Parent / Guardian' or cell.value == 'Other Guardian':
                if cell.offset(column=1).value is None:
                    pattern_fill(cell.offset(column=1), color=Color.YELLOW)
               
    # Check that students have a legal guardian on listed
    for row in contac.iter_rows(min_row=2, max_row=contac.max_row, min_col=8, max_col=8):
        for cell in row:
            if cell.value == 'No':
                pattern_fill(cell, color=Color.YELLOW)
    
    # Highlight when student address is the same as school's address
    for row in contac.iter_rows(min_row=2, max_row=contac.max_row, min_col=12, max_col=12):
        for cell in row:
            if cell.value is None:
                pattern_fill(cell, color=Color.YELLOW)
            else:
                for address in school_addresses:
                    if address in cell.value.lower():
                        pattern_fill(cell, color=Color.ORANGE)

                    