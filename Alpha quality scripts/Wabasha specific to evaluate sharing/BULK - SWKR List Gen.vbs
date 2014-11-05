'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ACTIONS - SWKR list gen"
start_time = timer

'LOADING ROUTINE FUNCTIONS----------------------------------------------------------------------------------------------------
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\DHS-MAXIS-Scripts\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'CONNECTS TO MAXIS
EMConnect ""

'DIALOG TO DETERMINE WHERE TO GO IN MAXIS TO GET THE INFO
BeginDialog SWKR_list_generator_dialog, 0, 0, 161, 87, "SWKR generator"
  DropListBox 75, 5, 65, 15, "REPT/ACTV"+chr(9)+"REPT/REVS"+chr(9)+"REPT/REVW", REPT_panel
  EditBox 60, 25, 20, 15, footer_month
  EditBox 135, 25, 20, 15, footer_year
  EditBox 105, 45, 35, 15, X179_number
  ButtonGroup ButtonPressed
    OkButton 25, 65, 50, 15
    CancelButton 85, 65, 50, 15
  Text 20, 10, 50, 10, "Create list from:"
  Text 10, 30, 45, 10, "Footer month:"
  Text 90, 30, 40, 10, "Footer year:"
  Text 20, 50, 85, 10, "X179### (if not yourself):"
EndDialog

'DISPLAYS DIALOG
Dialog SWKR_list_generator_dialog
If buttonpressed = 0 then stopscript

'CHECKS FOR PASSWORD PROMPT/MAXIS STATUS
transmit
EMReadScreen MAXIS_check, 5, 1, 39
If MAXIS_check <> "MAXIS" and MAXIS_check <> "AXIS " then script_end_procedure("MAXIS is not found on this screen. Are you passworded out? The script will close. Navigate to MAXIS and try again. You may need to restart BlueZone.")

'NAVIGATES BACK TO SELF TO FORCE THE FOOTER MONTH, THEN NAVIGATES TO THE SELECTED SCREEN
back_to_self
EMWriteScreen "________", 18, 43
call navigate_to_screen("rept", right(REPT_panel, 4))
If right(REPT_panel, 4) = "REVS" then
  current_month_plus_one = datepart("m", dateadd("m", 1, date))
  If len(current_month_plus_one) = 1 then current_month_plus_one = "0" & current_month_plus_one
  current_month_plus_one_year = datepart("yyyy", dateadd("m", 1, date))
  current_month_plus_one_year = right(current_month_plus_one_year, 2)
  EMWriteScreen current_month_plus_one, 20, 43
  EMWriteScreen current_month_plus_one_year, 20, 46
  transmit
  EMWriteScreen footer_month, 20, 55
  EMWriteScreen footer_year, 20, 58
  transmit
  footer_month = current_month_plus_one
  footer_year = current_month_plus_one_year
End if

'CHECKS TO MAKE SURE WE'VE MOVED PAST SELF MENU. IF WE HAVEN'T, THE SCRIPT WILL STOP. AN ERROR MESSAGE SHOULD DISPLAY ON THE BOTTOM OF THE MENU.
EMReadScreen SELF_check, 4, 2, 50
If SELF_check = "SELF" then script_end_procedure("Can't get past SELF menu. Check error message and try again!")

'IF THE X179 NUMBER WASN'T ENTERED, IT WILL SELECT THE DEFAULT. THE FOLLOWING IF...THEN WILL ONLY BE USED IF A NUMBER WAS ENTERED.
If X179_number <> "" then
  If REPT_panel = "REPT/ACTV" then 'THE REPT PANEL HAS THE X179 NUMBER IN DIFFERENT COLUMNS. THIS WILL DETERMINE THE CORRECT COLUMN FOR THE X179 NUMBER TO GO
    X179_col = 13
  Else
    X179_col = 6
  End if  
  EMReadScreen default_X179_number, 3, 21, X179_col 'CHECKING THE CURRENT X179 NUMBER. IF IT DOESN'T NEED TO CHANGE IT WON'T. OTHERWISE, THE SCRIPT WILL INPUT THE CORRECT NUMBER.
  If ucase(X179_number) <> default_X179_number then
    EMWriteScreen worker_county_code & X179_number, 21, X179_col
    transmit
  End if
End if

'DEFINES THE EXCEL_ROW VARIABLE FOR WORKING WITH THE SPREADSHEET
excel_row = 2

'OPENS A NEW EXCEL SPREADSHEET
Set objExcel = CreateObject("Excel.Application") 
objExcel.Visible = True  
Set objWorkbook = objExcel.Workbooks.Add() 

'FORMATS THE EXCEL SPREADSHEET WITH THE HEADERS, AND SETS THE COLUMN WIDTH
ObjExcel.Cells(1, 1).Value = "M#"
objExcel.Cells(1, 1).Font.Bold = TRUE
objExcel.Cells(1, 1).ColumnWidth = 9
ObjExcel.Cells(1, 2).Value = "Name"
objExcel.Cells(1, 2).Font.Bold = TRUE
objExcel.Cells(1, 2).ColumnWidth = 27
ObjExcel.Cells(1, 3).Value = "SWKR name"
objExcel.Cells(1, 3).Font.Bold = TRUE
objExcel.Cells(1, 3).ColumnWidth = 35
ObjExcel.Cells(1, 4).Value = "Copy of Notice?"
objExcel.Cells(1, 4).Font.Bold = TRUE
objExcel.Cells(1, 4).ColumnWidth = 20

'THIS DO...LOOP DUMPS THE CASE NUMBER AND NAME OF EACH CLIENT INTO A SPREADSHEET
Do

  EMReadScreen last_page_check, 21, 24, 02

  'This Do...loop checks for the password prompt.
  Do
    EMReadScreen password_prompt, 38, 2, 23
    IF password_prompt = "ACF2/CICS PASSWORD VERIFICATION PROMPT" then MsgBox "You are locked out of your case. Type your password then try again."
  Loop until password_prompt <> "ACF2/CICS PASSWORD VERIFICATION PROMPT"

  row = 7 'defining the row to look at
  Do
    If REPT_panel = "REPT/ACTV" then
      EMReadScreen case_number, 8, row, 12 'grabbing case number
      EMReadScreen client_name, 18, row, 21 'grabbing client name
    Else
      EMReadScreen case_number, 8, row, 6 'grabbing case number
      EMReadScreen client_name, 15, row, 16 'grabbing client name
    End if
    ObjExcel.Cells(excel_row, 1).Value = trim(case_number)
    ObjExcel.Cells(excel_row, 2).Value = trim(client_name)
    excel_row = excel_row + 1
    row = row + 1
  Loop until row = 19 or trim(case_number) = ""

  PF8 'going to the next screen


Loop until last_page_check = "THIS IS THE LAST PAGE"

'NOW THE SCRIPT IS CHECKING STAT/AREP FOR EACH CASE.----------------------------------------------------------------------------------------------------

excel_row = 2 'Resetting the case row to investigate.

do until ObjExcel.Cells(excel_row, 1).Value = "" 'shuts down when there's no more case numbers
  SWKR_name = "" 'Resetting this variable in case a SWKR cannot be found.
  case_number = ObjExcel.Cells(excel_row, 1).Value 
  If case_number = "" then exit do

  'This Do...loop gets back to SELF
  do
    EMSendKey "<PF3>"
    EMWaitReady 0, 0
    EMReadScreen SELF_check, 27, 2, 28
  loop until SELF_check = "Select Function Menu (SELF)"

  'NAVIGATES TO STAT/SWKR
  EMWriteScreen "stat", 16, 43
  EMWriteScreen "        ", 18, 43
  EMWriteScreen case_number, 18, 43
  EMSetCursor 21, 70
  EMSendKey "swkr" + "<enter>"
  EMWaitReady 0, 0

  'CHECKS FOR ERROR PRONE CASES
  EMReadScreen error_prone_check, 4, 2, 52
  If error_prone_check = "ERRR" then
    EMWriteScreen "SWKR", 20, 71
    transmit
  End if

    'NAVIGATES TO SWKR, READS THE NAME, AND ADDS TO SPREADSHEET
  call navigate_to_screen("stat", "swkr")
  EMReadScreen SWKR_name, 34, 6, 32
  swkr_name = replace(swkr_name, "_", "")
  ObjExcel.Cells(excel_row, 3).Value = swkr_name

'Navigates to SWKR, CHECKS FOR NOTICE Y/N, ADDS TO SPREADSHEET
  call navigate_to_screen("stat", "swkr")
  EMReadScreen NOTC_Y_N, 1, 15, 63
  If NOTC_Y_N = "_" then NOTC_Y_N = ""
  ObjExcel.Cells(excel_row, 4).Value = NOTC_Y_N

  
      excel_row = excel_row + 1 'setting up the script to check the next row.
loop


script_end_procedure("")
