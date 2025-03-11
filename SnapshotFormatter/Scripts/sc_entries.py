from snapshot_structures import months_end, today
from colors import Color, pattern_fill

# create exceptions for schools that weren't allowed on campus August or September
exceptions = ['ABC Elementary School']

# Highlight months with no SC entries in Site Coordination Entries sheet
def highlight_sc_entries(worksheet):  
    for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=3, max_col=5):
        for cell in row:
            row_school_name = worksheet['A' + str(cell.row)].value
            sce_month_name = worksheet['B' + str(cell.row)].value
            row_date = months_end[sce_month_name]
            #Check row month against current month
            if row_date <= today:
                if row_school_name in exceptions:
                    if sce_month_name == 'August' or sce_month_name == 'September' or sce_month_name == 'October':
                        pass
#                     elif row_school_name == 'Grape Street Elementary School' and sce_month_name == 'November':
#                         pass
                    else:
                        if cell.value == 0:
                            pattern_fill(cell, color=Color.YELLOW)
                else:
                    if cell.value == 0:
                        pattern_fill(cell, color=Color.YELLOW)