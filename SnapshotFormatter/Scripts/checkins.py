from snapshot_structures import *
#from snapshot_structures import months_beg, today, cie_col_num
from colors import Color, pattern_fill

###Check-ins Highlights
# Highlight months with no check-ins after enrollment date but before exit date in Check-ins sheet
def highlight_no_checkins(worksheet):
    # Create list of column letters used by Month columns
    month_letters = []
    for col in worksheet.iter_cols(min_row=1, max_row=1, min_col=8, max_col=worksheet.max_column):
        month_letters.append(col[0].coordinate[0])
    # Remove December & June from list
    month_letters.remove('L')
    month_letters.remove('R')

    for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=6, max_col=6):
        for cell in row:
            if cell.offset(column=-1).value == 'Case Managed':
                enroll_start = cell.value
                enroll_end = cell.offset(column=1).value
                # Loop through month columns using their cell letters
                for letter in month_letters:
                    header = worksheet[letter + '1']
                    c_offset = worksheet[letter + str(cell.row)]
                    col_date = months_beg[header.value]
                    # skip august for high schools
                    if ('High' in cell.offset(column=-5).value or 'Complex' in cell.offset(column=-5).value) and header.value == 'August':
                        pass
                    # Check column date against today
                    elif col_date <= today:
                        try:
                            if c_offset.value == 0 and (enroll_start < months_end[header.value]) and (enroll_end > months_beg[header.value]):
                                pattern_fill(c_offset, color=Color.YELLOW)
                        except:
                            print(header, c_offset, col_date)
                            print(col_date <= today, c_offset.value)
                            print(cell.coordinate, cell.value, 'Error in Check-ins tab for ' + cell.offset(column=-5).value)

def highlight_exits(worksheet):
    # Highlight exited students in grey               
    for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=7, max_col=7):
        for cell in row:
            if cell.value != datetime(9999, 12, 31, 0, 0):
                pattern_fill(cell, color=Color.GREY)
    # Highlight students that may need to be exited (last 3 months flagged)
    
    # set last month equal to the month before current
    this_month = datetime.now().strftime("%B")
    current_month_col = cie_col_num[this_month]
    student = cie_col_num['Student']
    
    for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=student, max_col=student):
        column_diff = current_month_col - student
        for cell in row:
            # Account for December not getting highlights
            if this_month == 'January':
                # Check if current month is highlighted
                current_month = cell.offset(column=column_diff)
                if 'FFCE00' in str(current_month.fill) or '00FFFF00' in str(current_month.fill):
                    #Check if previous 2 months are highlighted / Dec == 0
                    prev_month1, prev_month2 = cell.offset(column=column_diff-1), cell.offset(column=column_diff-2)
                    if (prev_month1.value == 0) and ('FFCE00' in str(prev_month2.fill) or '00FFFF00' in str(prev_month2.fill)):
                        pattern_fill(cell, color=Color.YELLOW)
            elif this_month == 'February':
                # Check if current month is highlighted
                current_month = cell.offset(column=column_diff)
                if 'FFCE00' in str(current_month.fill) or '00FFFF00' in str(current_month.fill):
                    #Check if previous 2 months are highlighted
                    prev_month1, prev_month2 = cell.offset(column=column_diff-1), cell.offset(column=column_diff-2)
                    if (prev_month2.value == 0) and ('FFCE00' in str(prev_month1.fill) or '00FFFF00' in str(prev_month1.fill)):
                        pattern_fill(cell, color=Color.YELLOW)
            else:
                # Check if current month is highlighted
                current_month = cell.offset(column=column_diff)
                if 'FFCE00' in str(current_month.fill) or '00FFFF00' in str(current_month.fill):
                    #Check if previous 2 months are highlighted / Dec == 0
                    prev_month1, prev_month2 = cell.offset(column=column_diff-1), cell.offset(column=column_diff-2)
                    if ('FFCE00' in str(prev_month1.fill) or '00FFFF00' in str(prev_month1.fill)) and ('FFCE00' in str(prev_month2.fill) or '00FFFF00' in str(prev_month2.fill)):
                        pattern_fill(cell, color=Color.YELLOW)
