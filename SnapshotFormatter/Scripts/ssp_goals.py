from snapshot_structures import *
from colors import Color, pattern_fill

# Highlight students with fewer than 2 targets in SSP Goals sheet
def ssp_less_than_2_goals(book):
    gol = book['SSP Goals']
    for row in gol.iter_rows(min_row=2, max_row=gol.max_row, min_col=10, max_col=10):
        for cell in row:
            if cell.value < 2:
                if cell.offset(column=-5).value != 'Non-Case Managed':
                    pattern_fill(cell, color=Color.YELLOW)
                    
# Highlights students with no SEL goals
def ssp_no_SEL_goals(book):
    gol = book['SSP Goals']
    for row in gol.iter_rows(min_row=2, max_row=gol.max_row, min_col=11, max_col=11):
        for cell in row:
            if cell.value == 'No':
                if cell.offset(column=-6).value != 'Non-Case Managed':
                    pattern_fill(cell, color=Color.YELLOW)
                    
# Highlight students with 0 as the baseline for Other (SEL) Metric in SSP Goals sheet
def ssp_zero_baseline(book):
    gol = book['SSP Goals']
    for row in gol.iter_rows(min_row=2, max_row=gol.max_row, min_col=7, max_col=7):
        for cell in row:
            if cell.value == 'Other (SEL)' and cell.offset(column=1).value == '0':
                pattern_fill(cell.offset(column=1), color=Color.YELLOW)
    
# Highlight goals outside CISLA specs SSP Goals sheet
def ssp_highlight_banned_goals(book):
    allowed_goals = ['Attendance Rate (%)', 'GPA', 'Overall SEAD assessment score', 'Other (SEL)', 'Reading Level']
    gol = book['SSP Goals']
    for row in gol.iter_rows(min_row=2, max_row=gol.max_row, min_col=7, max_col=7):
        for cell in row:
            if cell.value not in allowed_goals:
                pattern_fill(cell, color=Color.YELLOW)
