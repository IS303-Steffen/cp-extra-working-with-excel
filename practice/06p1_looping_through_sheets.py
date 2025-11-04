from helper_functions import clear_screen
clear_screen() # This just clears the terminal window each time we run code

# ===============================
# LOOPING THROUGH SHEETS PRACTICE
# ===============================

# 1. PRACTICE LOOPING
'''
Import the mock_grades file.

PART 1:
    Change the grade of anyone with a "C-" to be an "F" on the active sheet.
    Save the results as a new workbook called "example_06.xlsx"

PART 2:
    If you can successfully do that, then alter your code to change "C-" to "F"
    on ALL of the sheets in mock_grades. Save the results to
    "example_06.xlsx"

'''
import openpyxl
external_wb = openpyxl.load_workbook(r"mock_grades.xlsx")    #load the workbook
for ws_obj in external_wb.worksheets:   #to make this work for all sheets, not just the active one
    for row in ws_obj.iter_rows(min_col=2, max_col=2):  #loop through the active worksheet, focusing only on the rows with grades
        if row[0].value == "C-":
            row[0].value = "F"        #change the value from C- to F in all of the applicable rows in that column
external_wb.save("example_06.xlsx")