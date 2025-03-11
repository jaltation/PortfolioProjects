from snapshot_structures import *
from colors import Color, pattern_fill

    
# Highlight Tier 2 supports missing TSG in Activity
def tier2_missing_activity(book):
    tr2 = book['Tier II']
#     # Highlight Sports/Recreation/Clubs Tier 2 supports missing TSG in Activity
#     for row in tr2.iter_rows(min_row=2, max_row=tr2.max_row, min_col=6, max_col=6):
#         for cell in row:
#             if cell.value == 'Recreation/Sports/Clubs' and cell.offset(column=1).value is None:
#                 pattern_fill(cell, color=Color.ORANGE)
#                 pattern_fill(cell.offset(column=1), color=Color.ORANGE)
#                 pattern_fill(cell.offset(column=2), color=Color.ORANGE)
                
    # Highlight Tier 2 supports that have no activity but a tsg keyword in the notes
    for row in tr2.iter_rows(min_row=2, max_row=tr2.max_row, min_col=8, max_col=8):
        for cell in row:
            for word in tier2_kw:
                if word in str(cell.value).lower() and cell.offset(column=-1).value is None:
                    pattern_fill(cell, color=Color.YELLOW)
                    pattern_fill(cell.offset(column=-1), color=Color.YELLOW)
                    pattern_fill(cell.offset(column=-2), color=Color.YELLOW)
                    
    # Highlight SEL Tier 2 supports missing TSG in Activity
    for row in tr2.iter_rows(min_row=2, max_row=tr2.max_row, min_col=6, max_col=6):
        for cell in row:
            if '(SEL)' in cell.value or 'Relationship Skills' in cell.value:
                pattern_fill(cell, color=Color.YELLOW)
                pattern_fill(cell.offset(column=1), color=Color.YELLOW)
                pattern_fill(cell.offset(column=2), color=Color.YELLOW)

def tier2_missing_notes(book):
    tr2 = book['Tier II']
    for row in tr2.iter_rows(min_row=2, max_row=tr2.max_row, min_col=8, max_col=8):
        for cell in row:
            if cell.value is None:
                pattern_fill(cell, color=Color.RED)              
