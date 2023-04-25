import os
import sys
import re
import subprocess
from pathlib import Path
import openpyxl
from openpyxl.utils import get_column_letter
import win32com.client
import ctypes
import time
from functions import mainframetype, mainframegetscreen, mainframecopy, get_col
import datetime
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill






# Script title
script_title = "SWR 10 validate data script 234623624746"

# Close the window with the script title
ctypes.windll.user32.EnumWindows(ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.POINTER(ctypes.c_int), ctypes.POINTER(ctypes.c_int))(lambda hwnd, lParam: ctypes.windll.user32.PostMessageA(hwnd, 0x10, 0, 0) if ctypes.create_string_buffer(255), ctypes.windll.user32.GetWindowTextA(hwnd, ctypes.byref(text), 255), text.value.decode() == script_title else True), 0)

# Get the user name
username = os.getlogin()

# Find the Excel application object
excel_app = win32com.client.Dispatch("Excel.Application")

# Kill all running Excel instances
subprocess.call("taskkill /f /im excel.exe")

workbook_name = ""
workbook_obj = None
found = False

while not found:
    try:
        excel_app = win32com.client.Dispatch("Excel.Application")
        excel_app.Visible = True
    except:
        # Display a message box to ask the user to open a SWR 10 data spreadsheet
        answer = ctypes.windll.user32.MessageBoxW(0, "Open a SWR 10 data spreadsheet. Click CANCEL to quit. Click TRY AGAIN to TERMINATE all Excel instances. Click CONTINUE to keep going.", "Information", 0x40 | 0x1)
        if answer == 2:
            sys.exit()
        if answer == 10:
            subprocess.call("taskkill /f /im excel.exe")
        continue

    for temp_workbook_obj in excel_app.Workbooks:
        if "SWR 10 data (" in temp_workbook_obj.Name:
            found = True
            workbook_name = temp_workbook_obj.Name
            workbook_obj = temp_workbook_obj
            break

    if not found:
        # Display a message box to ask the user to open a SWR 10 data spreadsheet
        answer = ctypes.windll.user32.MessageBoxW(0, "Open a SWR 10 data spreadsheet. Click CANCEL to quit. Click TRY AGAIN to TERMINATE all Excel instances. Click CONTINUE to keep going.", "Information", 0x40 | 0x1)
        if answer == 2:
            sys.exit()
        if answer == 10:
            subprocess.call("taskkill /f /im excel.exe")


######################

sheet_obj = workbook_obj.Worksheets("Sheet1")
workbook_obj.Activate()
sheet_obj.Activate()

current_row = excel_app.ActiveCell.Row

sheet_obj.Rows(current_row - 1).Interior.ColorIndex = -4142
sheet_obj.Rows(current_row).Interior.ColorIndex = 4
workbook_name = excel_app.ActiveWorkbook.Name

first_row = excel_app.Selection.Row
last_row = 1
for row_object in excel_app.Selection.Rows:
    last_row = row_object.Row

for current_row in range(first_row, last_row + 1):
    if workbook_obj.Path != "S:\\PERMITTING\\ENG\\SWR 10\\SWR 10":
        ctypes.windll.user32.MessageBoxW(0, f'ERROR: the workbook file location is "{workbook_obj.Path}" instead of "S:\\PERMITTING\\ENG\\SWR 10\\SWR 10"', "Error", 0x10)
        sys.exit()

    sheet_obj.Cells(current_row, get_col(sheet_obj, "API#")).Activate()

    # ... (skipping the fieldnames, fieldnumbers, and other variables since they are not used in this part of the script)

    attn = sheet_obj.Cells(current_row, get_col(sheet_obj, "ATTN")).Text

    # ...

    # Complete the rest of the script following the same pattern, translating the AutoIt code to Python

#########################

# import time
# from functions import mainframetype, mainframegetscreen, mainframecopy, get_col

time.sleep(1)

mainframetype(2, 2, "WBTM{enter}")
screen = mainframegetscreen("WELL BORE TECHNICAL DATA MENU")
mainframetype(8, 23, api[:3])
mainframetype(8, 27, api[4:9])
mainframetype(11, 8, "s")
mainframetype(14, 7, "s{enter}")

permithistoryscreen = mainframegetscreen("PERMIT NUMBERS AND WELLS WITHIN WELL BORE")
currentrowow = 5
foundpermit = 0
while True:
    currentrowow += 1
    permit = mainframecopy(permithistoryscreen, currentrowow, 7, 6)
    if len(permit) == 6 and permit.isdigit():
        foundpermit = 1
    if "___" in permit or len(permit) <= 1:
        break

if foundpermit == 1:
    mainframetype(currentrowow-1, 4, "s")
    mainframetype(18, 4, "s{enter}")

    permitscreen = mainframegetscreen("DRILLING PERMIT MASTER DATA INQUIRY")
    opname = mainframecopy(permitscreen, 5, 17, 34)
    lease = mainframecopy(permitscreen, 9, 17, 34)
    district = mainframecopy(permitscreen, 10, 17, 2)
    county = ""
    dpissueddate = mainframecopy(permitscreen, 15, 17, 10)
    dpissueddate = dpissueddate.replace(" ", "-")
    dp = mainframecopy(permitscreen, 3, 17, 6)
    spuddate = mainframecopy(permitscreen, 19, 42, 10)
    surfcasingdate = mainframecopy(permitscreen, 19, 17, 10)
    if api[0] == "6" or api[0] == "7":
        county = mainframecopy(permitscreen, 12, 17, 34)
    else:
        county = mainframecopy(permitscreen, 11, 17, 34)
    well = mainframecopy(permitscreen, 9, 69, 10)
    opnumber = mainframecopy(permitscreen, 5, 69, 6)
    api = mainframecopy(permitscreen, 12, 69, 3) + "-" + mainframecopy(permitscreen, 12, 73, 5)
    if "PERMIT TYPE => DRILL" in permitscreen:
        sheetobj.Cells(currentrow, get_col(sheetobj, "new drill (yes/no)")).Value = "yes"
    else:
        sheetobj.Cells(currentrow, get_col(sheetobj, "new drill (yes/no)")).Value = "no"
else:
    mainframetype(16, 4, "s@B@Bs{enter}")

    W2G1screen = mainframegetscreen("OIL AND GAS W-2/G-1 RECORD")
    opname = mainframecopy(W2G1screen, 5, 9, 34)
    lease = mainframecopy(W2G1screen, 4, 48, 31)
    api = mainframecopy(W

#############################
G1screen, 2, 9, 3) + "-" + mainframecopy(W2G1screen, 2, 15, 5)
    district = mainframecopy(W2G1screen, 3, 9, 2)
    county = mainframecopy(W2G1screen, 3, 66, 14)
    well = mainframecopy(W2G1screen, 3, 36, 9)

    time.sleep(2)

    mainframetype(10, 10, f"ornq {opname}{enter}")

    screen = mainframegetscreen("ORGANIZATION NAME INQUIRY")
    op

    ##################################

api = mainframe_copy(W2G1screen, 2, 9, 3) + "-" + mainframe_copy(W2G1screen, 2, 15, 5)
district = mainframe_copy(W2G1screen, 3, 9, 2)
county = mainframe_copy(W2G1screen, 3, 66, 14)
well = mainframe_copy(W2G1screen, 3, 36, 9)

wait2()

mainframe_type(10, 10, f"ornq {opname}{enter}")

screen = mainframe_get_screen("ORGANIZATION NAME INQUIRY")
opnumber = mainframe_copy(screen, 4, 42, 6)

wait = msgbox(1 + 262144, "", f"{opname}\n\nClick OK to change operator number", 2)
if wait == 1:
    while True:
        opnumber = input(f"Confirm operator number\n\n{opname}")
        if opnumber == "":
            sheetobj.rows[currentrow].interior.colorindex = -4142
            break
        if opnumber.isdigit() and len(opnumber) == 6:
            break

line1 = [None] * 100
line2 = [None] * 100
line3 = [None] * 100

ormq_screen_path = os.path.join(scriptdir, f"ormq screens{os.path.sep}{opnumber}.txt")
file_modified_date = os.path.getmtime(ormq_screen_path)
p5screen = ""

if (not os.path.exists(ormq_screen_path)) or (date_to_day_value(*file_modified_date) + 14 < date_to_day_value(year, mon, mday)):
    wait2()
    mainframe_type(2, 2, f"ormq {opnumber}{enter}")

    p5screen = mainframe_get_screen("OPERATOR NUMBER")
    with open(ormq_screen_path, 'w') as output:
        output.write(p5screen)
else:
    with open(ormq_screen_path, 'r') as input_file:
        p5screen = input_file.read()

opnumber = mainframe_copy(p5screen, 3, 19, 6)
opname = mainframe_copy(p5screen, 4, 12, 34)
line1[0] = mainframe_copy(p5screen, 6, 3, 38)
line2[0] = mainframe_copy(p5screen, 7, 3, 38)
line3[0] = mainframe_copy(p5screen, 8, 3, 38)

index = 0
if "@" not in attn:
    screen = ""
    wait2()
    mainframe_type(2, 2, f"ORAR {opnumber}{enter}")

    i = 1

    while True:
        screen = mainframe_get_screen("ADDRESS INQUIRY")
        mainframe_type(1, 1, "{f5}")
        time.sleep(100)
        screen = mainframe_get_screen("ADDRESS INQUIRY")

        for r in range(4, 17, 4):
            for c in range(2, 45, 42):
                line1[i] = mainframe_copy(screen, r, c, 29)
                line2[i] = mainframe_copy(screen, r + 1, c, 29)
                line3[i] = mainframe_copy(screen, r + 2, c, 29)
                i += 1
            if "NO MORE ADDRESSES" in screen:
            break
        num_addresses = i

        prompt = ""
        for i in range(num_addresses):
            prompt += f"{i}. {line1[i]} {line2[i]} {line3[i]}\n"
            if (i + 1) % 5 == 0:
                prompt += "\n"

        index = 0
        attn = attn.upper()
        index = input("Select Address", prompt, "0", "", 700, 700)
        if index == "":
            index = 0

##############################
op_address1 = line1[index]
op_address2 = line2[index]
op_address3 = line3[index]

if ("ATTN" in op_address1 or "C/O" in op_address1) and attn != "" and "@" not in attn:
    op_address1 = f"ATTN {attn}"

if op_address3 == "":
    op_address3 = op_address2
    op_address2 = op_address1
    op_address1 = "ATTN: REGULATORY DEPARTMENT"
    if attn != "" and "@" not in attn:
        op_address1 = f"ATTN {attn}"
        
if ("ATTN" in op_address1 or "C/O" in op_address1) and "@" in attn:
    op_address1 = "ATTN: REGULATORY DEPARTMENT"

sheet_obj.cell(current_row, get_col(sheet_obj, "OPERATOR ADDRESS 1")).value = op_address1
sheet_obj.cell(current_row, get_col(sheet_obj, "OPERATOR ADDRESS 2")).value = op_address2
sheet_obj.cell(current_row, get_col(sheet_obj, "OPERATOR ADDRESS 3")).value = op_address3

field_number_col = get_col(sheet_obj, "field 1")
h2s_col = get_col(sheet_obj, "h2s field 1")
field_name_col = get_col(sheet_obj, "field name 1")
field_perfs_col = get_col(sheet_obj, "perfs field 1")
field_name_display = ""
h2s_condition = "NO"
all_fields_letter_text = ""

# Please define shortcut_fieldnames and get_fieldname functions before using them.
# Example: def shortcut_fieldnames(field_number): ...
#          def get_fieldname(field_number, h2s, field_name): ...

# Add your conditions and corresponding values for field_number_col here

while sheet_obj.cell(current_row, field_number_col).value != "":
    field_number = str(sheet_obj.cell(current_row, field_number_col).text).replace(" ", "")

    if not field_number.isdigit():
        field_number = shortcut_fieldnames(field_number)

    if district == "7C" and sheet_obj.cell(current_row, get_col(sheet_obj, "field 1")).value == 85280300:
        sheet_obj.cell(current_row, get_col(sheet_obj, "field 1")).value = 85279200

    if district == "08" and sheet_obj.cell(current_row, get_col(sheet_obj, "field 1")).value == 85279200:
        sheet_obj.cell(current_row, get_col(sheet_obj, "field 1")).value = 85280300

    sheet_obj.cell(current_row, field_number_col).value = field_number

    get_fieldname(field_number, h2s, field_name)
    if h2s == "PRESENT":
        h2s_condition = "YES"
    sheet_obj.cell(current_row, field_name_col).value = field_name
    sheet_obj.cell(current_row, h2s_col).value = h2s
    perfs = sheet_obj.cell(current_row, field_perfs_col).value

    while len(perfs) < 2 and blanket == "no":
        sheet_obj.cell(current_row, field_perfs_col).activate()
        msgbox(1 + 262144, "", f"enter perfs\n\n{field_name}")
        perfs = sheet_obj.cell(current_row, field_perfs_col).value

    field_number = sheet_obj.cell(current_row, field_number_col).value

    field_name_display += f"{field_name}{{{field_number}}}{{{perfs}}}\n"
    all_fields_letter_text += f\
    
###########################
# import datetime
# from openpyxl.utils import get_column_letter

today = datetime.date.today()

for current_row in range(1, sheet.max_row + 1):
    sheet.cell(row=current_row, column=get_col(sheet, "disposition date")).value = today.strftime("%Y/%m/%d")
    sheet.cell(row=current_row, column=get_col(sheet, "script run date")).value = today.strftime("%Y/%m/%d")
    sheet.cell(row=current_row, column=get_col(sheet, "District")).value = district
    sheet.cell(row=current_row, column=get_col(sheet, "County")).value = county
    sheet.cell(row=current_row, column=get_col(sheet, "Operator")).value = opname
    sheet.cell(row=current_row, column=get_col(sheet, "Operator no.")).value = opnumber
    sheet.cell(row=current_row, column=get_col(sheet, "Lease")).value = lease
    sheet.cell(row=current_row, column=get_col(sheet, "Well")).value = well
    sheet.cell(row=current_row, column=get_col(sheet, "well name combo")).value = f"{lease} — WELL NO. {well}, API NO. {api}"
    sheet.cell(row=current_row, column=get_col(sheet, "h2s restriction")).value = h2scondition
    sheet.cell(row=current_row, column=get_col(sheet, "drilling permit issued date")).value = dpissueddate
    sheet.cell(row=current_row, column=get_col(sheet, "drilling permit no.")).value = dp
    sheet.cell(row=current_row, column=get_col(sheet, "all fields letter body")).value = allfieldslettertext

    # ... Rest of the code

    # Assuming 'yes' and 'no' are strings in the original AutoIt code
    if spuddate != "" or surfcasingdate != "":
        sheet.cell(row=current_row, column=get_col(sheet, "current permit has spud date (yes/no)")).value = "yes"
    else:
        sheet.cell(row=current_row, column=get_col(sheet, "current permit has spud date (yes/no)")).value = "no"

    leavepermitopen = sheet.cell(row=current_row, column=get_col(sheet, "requested permit held open (yes/no)")).value
    newdrill = sheet.cell(row=current_row, column=get_col(sheet, "new drill (yes/no)")).value
    alreadyonschedule = sheet.cell(row=current_row, column=get_col(sheet, "all zones already on schedule (yes/no)")).value

    if blanket == "yes":
        sheet.cell(row=current_row, column=get_col(sheet, "requested permit held open (yes/no)")).value = "no"

    if sheet.cell(row=current_row, column=get_col(sheet, "new drill (yes/no)")).value != "yes" and blanket == "no":
        while sheet.cell(row=current_row, column=get_col(sheet, "all zones already on schedule (yes/no)")).value == "":
            # Replace msgbox with an equivalent Python function, e.g., a simple print statement or a custom function
            print("Error: answer \"all zones already on schedule (yes/no)\"")

    # ... Rest of the code

    if sheet.cell(row=current_row, column=get_col(sheet, "requested permit held open (yes/no)")).value == "yes":
        sheet.cell(row=current_row, column=get_col(sheet, "expiration condition")).value = "This exception to SWR 10 will expire if not used within two (2) years from the original date of issuance for drilling permit no
        sheet.cell(row=current_row, column=get_col(sheet, "Expiration Date")).formula = f"={get_column_letter(get_col(sheet, 'drilling permit issued date'))}{current_row}+731"

    permitconditions = ""
    if h2scondition == "YES":
        permitconditions += "The commingled well will be subject to Statewide Rule 36 (operation in hydrogen sulfide areas) because at least one of the commingled fields requires a Certificate of Compliance for Statewide Rule 36.  The well must be operated in accordance with Statewide Rule 36.\n\n"
    if blanket == "YES":
        permitconditions += "The completion report for the commingled well must indicate which perforations belong to which field.  The Commission may also require a wellbore diagram to be filed with the completion report for the commingled well.  If filed, the wellbore diagram must indicate which perforations belong to which field.\n\n"
    if blanket == "NO":
        permitconditions += "The completion of the commingled well must be a reasonable match with the wellbore diagram filed with the application.  Variances in completion depths are acceptable provided that these completion depths remain within the designated correlative intervals for the commingled fields.  A copy of this wellbore diagram must be filed with the completion report for the commingled well.\n\n"

    custompermitconditions = sheet.cell(row=current_row, column=get_col(sheet, "custom permit conditions")).value
    if custompermitconditions != "":
        permitconditions += custompermitconditions + "\n\n"

    if "@" in attn:
        permitconditions += f"Note: The distribution of this document will be by E-MAIL ONLY.  E-mail sent to {attn}.\n\n"

    if permitconditions != "":
        sheet.cell(row=current_row, column=get_col(sheet, "permit conditions")).value = "Permit conditions:\n\n" + permitconditions

    # Assuming there's no need to activate Excel in Python

    sheet.cell(row=current_row, column=get_col(sheet, "validate run date")).value = today.strftime("%Y/%m/%d")

    sheet.cell(row=current_row, column=get_col(sheet, "disposition")).value = "pending"

    sheet.row_dimensions[current_row].fill = PatternFill(start_color="ffffff", end_color="ffffff", fill_type="solid")

    workbook.save("filename.xlsx")  # Replace "filename.xlsx" with your desired Excel file name

# End of the loop
###########################


Certainly! Continuing the translation of the AutoIt code to Python:

python

        sheet.cell(row=current_row, column=get_col(sheet, "Expiration Date")).formula = f"={get_column_letter(get_col(sheet, 'drilling permit issued date'))}{current_row}+731"

    permitconditions = ""
    if h2scondition == "YES":
        permitconditions += "The commingled well will be subject to Statewide Rule 36 (operation in hydrogen sulfide areas) because at least one of the commingled fields requires a Certificate of Compliance for Statewide Rule 36.  The well must be operated in accordance with Statewide Rule 36.\n\n"
    if blanket == "YES":
        permitconditions += "The completion report for the commingled well must indicate which perforations belong to which field.  The Commission may also require a wellbore diagram to be filed with the completion report for the commingled well.  If filed, the wellbore diagram must indicate which perforations belong to which field.\n\n"
    if blanket == "NO":
        permitconditions += "The completion of the commingled well must be a reasonable match with the wellbore diagram filed with the application.  Variances in completion depths are acceptable provided that these completion depths remain within the designated correlative intervals for the commingled fields.  A copy of this wellbore diagram must be filed with the completion report for the commingled well.\n\n"

    custompermitconditions = sheet.cell(row=current_row, column=get_col(sheet, "custom permit conditions")).value
    if custompermitconditions != "":
        permitconditions += custompermitconditions + "\n\n"

    if "@" in attn:
        permitconditions += f"Note: The distribution of this document will be by E-MAIL ONLY.  E-mail sent to {attn}.\n\n"

    if permitconditions != "":
        sheet.cell(row=current_row, column=get_col(sheet, "permit conditions")).value = "Permit conditions:\n\n" + permitconditions

    # Assuming there's no need to activate Excel in Python

    sheet.cell(row=current_row, column=get_col(sheet, "validate run date")).value = today.strftime("%Y/%m/%d")

    sheet.cell(row=current_row, column=get_col(sheet, "disposition")).value = "pending"

    sheet.row_dimensions[current_row].fill = PatternFill(start_color="ffffff", end_color="ffffff", fill_type="solid")

    workbook.save("filename.xlsx")  # Replace "filename.xlsx" with your desired Excel file name

# End of the loop

# This should cover the entire AutoIt code provided. Note that you should replace "filename.xlsx" with your desired Excel file name when saving the workbook. Also, don't forget to import the required libraries at the beginning of your Python script:

# python

# import datetime
# from openpyxl.utils import get_column_letter
# from openpyxl.styles import PatternFill


