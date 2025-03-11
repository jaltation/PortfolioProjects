from snapshot_structures import *
from colors import Color, pattern_fill

# Highlight non-matches in Attributes Check sheet
def attributes_check(book):
    atr = book['Attributes Check']
    # Highlight if Attribute doesn't match district
    for row in atr.iter_rows(min_row=2, max_row=atr.max_row, min_col=7, max_col=13):
        for cell in row:
            if cell.value != 'Match' and 'District: Unknown' not in cell.value:
                cm_status = atr['E' + str(cell.row)].value
                if cm_status == 'Case Managed':
                    pattern_fill(cell, color=Color.YELLOW)
    
    # Highlight if gender or race are listed as Unknown in CISDM
    for row in atr.iter_rows(min_row=2, max_row=atr.max_row, min_col=14, max_col=15):
        for cell in row:
            if 'CISDM: Unknown' in cell.value:
                if 'District: Decline to State' not in cell.value:
                    cm_status = atr['E' + str(cell.row)].value
                    if cm_status == 'Case Managed':
                        pattern_fill(cell, color=Color.YELLOW)
    
    # Highlight if Ethnicity is listed as unknown
    for row in atr.iter_rows(min_row=2, max_row=atr.max_row, min_col=16, max_col=16):
        for cell in row:
            if 'District: Hispanic | CISDM: Unknown' in cell.value or 'District: Hispanic | CISDM: None Listed' in cell.value or 'District: Hispanic | CISDM: None Listed, Unknown' in cell.value:
                cm_status = atr['E' + str(cell.row)].value
                if cm_status == 'Case Managed':
                    pattern_fill(cell, color=Color.YELLOW)
                