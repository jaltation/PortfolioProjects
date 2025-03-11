### Functions for tallying flags in Snapshots
import datetime
from numpy import unique
from snapshot_structures import cld_col_num, tr1_col_num, cie_col_num, school_tsg


## Tally functions for the caseload details sheet

# Tally students that need to be switched to non-cm
def tally_wrong_id(worksheet):
    id_col = cld_col_num['Student ID']
    wrong_id_tally = 0
    for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=id_col, max_col=id_col):
        for cell in row:
            if 'E21836' in str(cell.fill):
                wrong_id_tally += 1
    return wrong_id_tally

# Tally students that need to be switched to non-cm
def tally_cm_switch(worksheet):
    cm_col = cld_col_num['Case Managed Status']
    cm_switch_tally = 0
    for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=cm_col, max_col=cm_col):
        for cell in row:
            if 'FFCE00' in str(cell.fill) or '00FFFF00' in str(cell.fill):
                cm_switch_tally += 1
    return cm_switch_tally

# Tally missing parental consent forms
def tally_missing_consents(worksheet):
    parent = cld_col_num['Parent Consent + Supports Within Consent Date?']
    missing_consent = 0
    for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=parent, max_col=parent):
        for cell in row:
            if 'FFCE00' in str(cell.fill) or '00FFFF00' in str(cell.fill):
                missing_consent += 1
    return missing_consent

# Tally monthly check-in flags
def tally_missing_checkins(worksheet):
    checkin = cld_col_num['At least 1 check-in per month?']
    missing_checkin = 0
    for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=checkin, max_col=checkin):
        for cell in row:
            if 'FFCE00' in str(cell.fill) or '00FFFF00' in str(cell.fill):
                missing_checkin += 1
    return missing_checkin
    
    
# Tally caseload total
def tally_caseload_total(cld_df):
    cm_only  = cld_df[cld_df['Case Managed Status'] == 'Case Managed']
    cm_only = cm_only[(cm_only['Enrollment Exit Date'] == datetime.datetime(9999, 12, 31, 0, 0)) | 
                      ((cm_only['Student Needs Assessment Entered?'] == 'Yes') & (cm_only['Student Support Plan Entered?'] == 'Yes'))]
    try:
        total_cm_students = cm_only['Case Managed Status'].value_counts()[0]
    except:
        total_cm_students = 0
#     print(cld_df.head(2))
#     print(total_cm_students)
    return total_cm_students
    
# Find percent of students with SEL TSG
def sel_tsg_students(worksheet):    
    ini = cld_col_num['Initiatives']
    sel_tsg_students = 0
    tsg_students = 0
    student_total = worksheet.max_row - 1

    for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=ini, max_col=ini):
        for cell in row:
            # Tally students in an SEL TSG
            if cell.value is not None:
                if cell.value != 'Backpack Club' and cell.value != 'Good Grades Push':
                    tsg_students += 1
                if 'SEL' in cell.value:
                    sel_tsg_students += 1
    
    
    if student_total == 0:
        all_tsg_percent = 0
    else:
        all_tsg_percent = round((tsg_students / student_total * 100), 2)
    if sel_tsg_students == 0:
        sel_initiative_percent = 0
    else:
        sel_initiative_percent = round((sel_tsg_students / student_total * 100), 2)
        
    return sel_tsg_students, sel_initiative_percent, all_tsg_percent
        
# Find TSGs with no students
def list_unstarted_tsgs(worksheet, school_name):
    started_tsgs = []
    if school_name != 'All School':
        ini = cld_col_num['Initiatives']
        for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=ini, max_col=ini):
            for cell in row:
                if cell.value is not None:
                    student_enrolled_initiatives_list = cell.value.split("; ")
                    for tsg in student_enrolled_initiatives_list:
                        if tsg not in started_tsgs:
                            started_tsgs.append(tsg)

        list_of_unstarted_tsgs = [x for x in school_tsg[school_name] if x not in started_tsgs]
    else:
        list_of_unstarted_tsgs = []
    return list_of_unstarted_tsgs
        
# Tally initiatives not in school's listing
def tally_wrong_tsg(worksheet):
    wrong_init = 0
    ini = cld_col_num['Initiatives']
    for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=ini, max_col=ini):
        for cell in row:
            if 'F68E1E' in str(cell.fill):
                wrong_init += 1
    return wrong_init
                
# Tally SEL counts less than 1
def tally_missing_sel(worksheet):
    sel_col = cld_col_num['SEAD/DESSA']
    sel_tally = 0
    for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=sel_col, max_col=sel_col):
        for cell in row:
            if 'FFCE00' in str(cell.fill) or '00FFFF00' in str(cell.fill):
                sel_tally += 1
    return sel_tally
                
# Tally missing SNA
def tally_missing_sna(worksheet):
    sna_check = cld_col_num['Student Needs Assessment Entered?']
    sna_tally = 0
    for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=sna_check, max_col=sna_check):
        for cell in row:
            if 'FFCE00' in str(cell.fill) or '00FFFF00' in str(cell.fill):
                # don't count towards tally if student should be switched to non-cm
                if 'FFCE00' not in str(cell.offset(column=-7).fill) and '00FFFF00' not in str(cell.offset(column=-7).fill):
                    sna_tally += 1
    return sna_tally
                
# Tally missing SSP
def tally_missing_ssp(worksheet):
    ssp_check = cld_col_num['Student Support Plan Entered?']
    ssp_tally = 0
    for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=ssp_check, max_col=ssp_check):
        for cell in row:
            if 'FFCE00' in str(cell.fill) or '00FFFF00' in str(cell.fill):
                # don't count towards tally if student should be switched to non-cm
                if 'FFCE00' not in str(cell.offset(column=-8).fill) and '00FFFF00' not in str(cell.offset(column=-8).fill):
                    ssp_tally += 1
#     print(f"Missing SSP Tally: {ssp_tally}")
    return ssp_tally
                
# Tally missing TSGs, no more than one per student
def tally_missing_tsg(worksheet):
    missing_tsg = 0
    coordinate_list = []
    first_tsg = cld_col_num['At least 1 check-in per month?'] + 1
    for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=first_tsg, max_col=worksheet.max_column):
        for cell in row:
            if 'FFCE00' in str(cell.fill) or '00FFFF00' in str(cell.fill):
                coordinate_list.append(cell.coordinate[1:])
    missing_tsg = len(unique(coordinate_list))
    return missing_tsg




## Tally functions for Check-ins tab

# Tally students with 0 check-ins for a month
def tally_missing_monthly_checkin(worksheet):
    august = cie_col_num['August']
    # set last full month equal to the month before current
    last_month = datetime.datetime.now().strftime("%B")
    last_full_month = cie_col_num[last_month] - 1
    coordinate_list = []
    for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=august, max_col=last_full_month):
        for cell in row:
            if 'FFCE00' in str(cell.fill) or '00FFFF00' in str(cell.fill):
                coordinate_list.append(cell.coordinate[1:])
    missing_monthly_checkin = len(unique(coordinate_list))
    return missing_monthly_checkin

def exit_check(worksheet):
    exit_check_tally = 0
    exited_check_list = []
    student = cie_col_num['Student']
    for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=student, max_col=student):
        for cell in row:
            if 'FFCE00' in str(cell.fill) or '00FFFF00' in str(cell.fill):
                exit_check_tally += 1
                # Add student name to list
                exited_check_list.append(cell.value)

    return exit_check_tally, exited_check_list
    


    
## Tally functions for sc entries sheet

# Tally missing SC Entries
def tally_missing_sce(worksheet):
    missing_sc = 0
    for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=3, max_col=5):
        for cell in row:
            if 'FFCE00' in str(cell.fill) or '00FFFF00' in str(cell.fill):
                missing_sc += 1
    return missing_sc




## Tally functions for tier 1 sheet

# Tally missing parents served in tier 1 sheet
def tally_missing_parent_numbers(worksheet):
    parent_num = tr1_col_num['# of Parents / Guardians Served']
    missing_par = 0
    for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=parent_num, max_col=parent_num):
        for cell in row:
            if 'FFCE00' in str(cell.fill) or '00FFFF00' in str(cell.fill):
                missing_par += 1    
    return missing_par
                
# Tally missing tier 1 support notes
def tally_missing_tr1_notes(worksheet):
    notes_num = tr1_col_num['Notes']
    missing_notes = 0
    for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=notes_num, max_col=notes_num):
        for cell in row:
            if 'E21836' in str(cell.fill):
                missing_notes += 1
    return missing_notes
                
# Tally number of family & student tier 1 supports
def tally_tr1_supports(worksheet):
    category_col = tr1_col_num['Support Category']
    tr1_student_supports = 0
    tr1_family_supports = 0
    for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=category_col, max_col=category_col):
        for cell in row:
            if cell.value == 'Family Engagement' or cell.offset(column=3).value == 'Family Engagement':
                tr1_family_supports += 1
            else:
                tr1_student_supports += 1
    return tr1_student_supports, tr1_family_supports

    
def tally_tr1_activities(worksheet):
    activity_col = tr1_col_num['Activity']
    missing_tr1_activities = 0
    for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=activity_col, max_col=activity_col):
        for cell in row:
            if 'FFCE00' in str(cell.fill) or '00FFFF00' in str(cell.fill):
                missing_tr1_activities += 1
    return missing_tr1_activities

def tally_tr1_incorrect_activities(worksheet):
    activity_col = tr1_col_num['Activity']
    incorrect_tr1_activities = 0
    for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=activity_col, max_col=activity_col):
        for cell in row:
            if 'F68E1E' in str(cell.fill):
                incorrect_tr1_activities += 1
    return incorrect_tr1_activities

# Tally number ASSP Goal Tier 1 supports (not implemented)
def tally_assp_goals(worksheet):
    category_col = tr1_col_num['Support Category']
    if school_name != 'All School':
        for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=category_col, max_col=category_col):
            for cell in row:
                for goal in school_goals[school_name]:
                    if goal in cell.value or goal in cell.offset(column=1).value:
                        school_goals[school_name][goal] += 1
                        
def tally_assp_goals_from_notes(worksheet):
    notes_col = tr1_col_num['Notes']
    if school_name != 'All School':
        for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=notes_col, max_col=notes_col):
            for cell in row:
                if cell.value is not None:
                    for goal in school_goals_from_notes[school_name]:
                        for keyword in goal_kw[goal]:
                            if keyword.lower() in cell.value.lower():
                                school_goals_from_notes[school_name][goal] += 1
                                

                                

## Tally functions for Tier 2 sheet

# Tally missing activity in tier 2 tab
def tally_missing_activity(worksheet):
    missing_activities = 0
    for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=7, max_col=7):
        for cell in row:
            if 'FFCE00' in str(cell.fill) or '00FFFF00' in str(cell.fill):
                missing_activities += 1
    return missing_activities

def tally_incorrect_activity(worksheet):
    incorrect_activities = 0
    for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=7, max_col=7):
        for cell in row:
            if 'F68E1E' in str(cell.fill):
                incorrect_activities += 1
    return incorrect_activities

# tally missing notes in tier 2 tab
def tally_missing_t2_notes(worksheet):
    missing_t2_notes = 0
    for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=8, max_col=8):
        for cell in row:
            if 'E21836' in str(cell.fill):
                missing_t2_notes += 1
    return missing_t2_notes




## Tally functions for SSP Goals sheet

# Tally baselines of zero in ssp goals tab
def tally_baselines_zero(worksheet):
    baselines = 0
    for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=8, max_col=8):
        for cell in row:
            if 'FFCE00' in str(cell.fill) or '00FFFF00' in str(cell.fill):
                baselines += 1
    return baselines
    
# Tally targets less than 2 in ssp goals tab
def tally_few_targets(worksheet):
    low_targets = 0
    goal_target_list = []
    for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=10, max_col=10):
        for cell in row:
            if 'FFCE00' in str(cell.fill) or '00FFFF00' in str(cell.fill):
                goal_target_list.append(cell.offset(column=-8).value)
    low_targets = len(unique(goal_target_list))
    return low_targets
                    
# Tally missing SEL goal ssp goals tab
def tally_no_sel_goal(worksheet):
    no_sel_goal = 0
    no_sel_list = []
    for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=11, max_col=11):
        for cell in row:
            if 'FFCE00' in str(cell.fill) or '00FFFF00' in str(cell.fill):
                no_sel_list.append(cell.offset(column=-9).value)
    no_sel_goal = len(unique(no_sel_list))
    return no_sel_goal




## Tally functions for attributes sheet

# Tally mismatched attributes
def tally_mismatched_attributes(worksheet):
    mismatches = 0
    for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=7, max_col=16):
        student_mismatches = 0
        for cell in row:
            if 'FFCE00' in str(cell.fill) or '00FFFF00' in str(cell.fill):
                student_mismatches += 1
        if student_mismatches > 0:
            mismatches += 1
    return mismatches




## Tally functions for Contacts sheet

# Tally student contacts missing info in contact info tab
def tally_missing_contact(worksheet):
    contact_coordinate_list = []
    for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=5, max_col=11):
        for cell in row:
            if 'FFCE00' in str(cell.fill) or '00FFFF00' in str(cell.fill):
                if worksheet['E' + str(cell.row)].value is None:
                    contact_name = worksheet['B' + str(cell.row)].value
                else:
                    contact_name = worksheet['E' + str(cell.row)].value
                contact_coordinate_list.append(contact_name)
    missing_contact = len(unique(contact_coordinate_list))     
    return missing_contact
                
        
# Tally students with school address in contact info tab
def tally_school_as_address(worksheet):
    students_w_school_address_list = []
    for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=12, max_col=12):
        for cell in row:
            if 'F68E1E' in str(cell.fill):
                students_w_school_address_list.append(cell.offset(column=-10).value)
    school_addresses = len(unique(students_w_school_address_list))
    return school_addresses


