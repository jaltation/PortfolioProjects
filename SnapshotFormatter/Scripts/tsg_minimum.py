from snapshot_structures import *
from colors import Color, pattern_fill
from pandas import DataFrame
from itertools import islice

# May not be necessary later in the year               
def highlight_initiative(book, all_school=False, school="123 Middle School"):
    cld = book['Caseload Details']
    ini_col = cld_col_num['Initiatives']
    # Find number of CM students on caseload
    data = cld.values
    cols = next(data)[1:]
    data = list(data)
    idx = [r[0] for r in data]
    data = (islice(r, 1, None) for r in data)
    cld_df = DataFrame(data, index=idx, columns=cols)
    try:
        cm_student_tally = cld_df['Case Managed Status'].value_counts()[0]
    except:
        cm_student_tally = 0
#     print("------CM Tally:", cm_student_tally)
    
    # Highlight Initiative column if empty
    sel_tally = 0
    #school_name = cld['A2'].value
    school_name = school
    school_tsg_list = school_tsg[school_name]
    
    for row in cld.iter_rows(min_row=2, max_row=cld.max_row, min_col=ini_col, max_col=ini_col):
        for cell in row:
            # Highlight tsgs not in initiatives matrix
            if cell.value is not None:
                student_tsg_list = list(str(cell.value).split('; '))
                # Remove non-groups that use group initiative classification
                removal = ['Backpack Club', 'Good Grades Push']
                for activity in removal:
                    if activity in student_tsg_list:
                        student_tsg_list.remove(activity)
                # Code for All-School file
                if all_school is True:
                    school_name = cld['A' + str(cell.row)].value
                    school_tsg_list = school_tsg[school_name]
                    for tsg in student_tsg_list:
                        if tsg not in school_tsg_list:
                            pattern_fill(cell, color=Color.ORANGE)
                # Code for individual school files
                else:
                    for tsg in student_tsg_list:
                        if tsg not in school_tsg_list:
                            pattern_fill(cell, color=Color.ORANGE)
                # Count the number of students in a SEL TSG
                student_sel_tally = 0
                for tsg in student_tsg_list:
                    if 'SEL' in tsg:
                        student_sel_tally += 1
                if student_sel_tally > 0:
                    sel_tally += 1
                                    
    if cm_student_tally != 0:
        sel_ratio = sel_tally/cm_student_tally
    else:
        sel_ratio = 0
    if sel_ratio < 0.195:
        for row in cld.iter_rows(min_row=2, max_row=cld.max_row, min_col=ini_col, max_col=ini_col):
            for cell in row:
                if (cell.value is None) or ('SEL' not in cell.value):
                    pattern_fill(cell, color=Color.YELLOW)
                    
# hihghlight tsg activities that aren't on the school's tsg matrix              
def highlight_activity(book, all_school=False, school="123 Middle School"):
    #worksheets = [book['Tier 1'], book['Tier II']]
    worksheets = [book['Tier 1']]
    school_name = school
    school_tsg_list = school_tsg[school_name]
    
    for worksheet in worksheets:
        for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=7, max_col=7):
            for cell in row:
                # Highlight tsgs not in initiatives matrix
                if cell.value is not None:
                    student_tsg = cell.value
                    # Code for All-School file
                    if all_school is True:
                        school_name = worksheet['A' + str(cell.row)].value
                        school_tsg_list = school_tsg[school_name]
                        if student_tsg not in school_tsg_list:
                            pattern_fill(cell, color=Color.ORANGE)
                    # Code for individual school files
                    else:
                        if student_tsg not in school_tsg_list:
                            pattern_fill(cell, color=Color.ORANGE)