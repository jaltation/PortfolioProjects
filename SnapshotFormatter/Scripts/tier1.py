from snapshot_structures import *
from colors import Color, pattern_fill


###Tier 1 Highlights                    
# Highlight parent attendance of 0 when category is Family Engagement in Tier 1 sheet
def highlight_zero_parents(worksheet):                    
    sprtcat = tr1_col_num['Support Category']
    for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=sprtcat, max_col=sprtcat):
        for cell in row:
            if (cell.value == 'Family Engagement' or cell.offset(column=3).value == 'Family Engagement') and (cell.offset(column=7).value is None):
                pattern_fill(cell.offset(column=7), color=Color.YELLOW)
                
# Highlight Tier 1 supports w/ no notes Tier 1 sheet 
def highlight_tier1_notes(worksheet):
    notes_num = tr1_col_num['Notes']
    for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=notes_num, max_col=notes_num):
        for cell in row:
            if cell.value is None:
                pattern_fill(cell, color=Color.RED)
                
# Highlight Tier 1 supports missing TSG in Activity
def tier1_missing_activity(tr1):               
    # Highlight Tier 1 supports that have no activity but a tsg keyword in the notes
    notes = tr1_col_num['Notes']
    activity_offset = tr1_col_num['Activity'] - tr1_col_num['Notes']
    for row in tr1.iter_rows(min_row=2, max_row=tr1.max_row, min_col=notes, max_col=notes):
        for cell in row:
            for word in tier2_kw:
                if word in str(cell.value).lower() and cell.offset(column=activity_offset).value is None:
                    pattern_fill(cell.offset(column=activity_offset), color=Color.YELLOW)
                    