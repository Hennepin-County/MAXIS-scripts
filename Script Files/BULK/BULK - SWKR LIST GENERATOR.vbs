'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - SWKR LIST GENERATOR.vbs"
start_time = timer

'CONNECTS TO MAXIS
EMConnect ""

'DIALOG TO DETERMINE WHERE TO GO IN MAXIS TO GET THE INFO
BeginDialog SWKR_list_generator_dialog, 0, 0, 156, 115, "SWKR list generator dialog"
  DropListBox 65, 5, 85, 15, "REPT/ACTV"+chr(9)+"REPT/REVS"+chr(9)+"REPT/REVW", REPT_panel
  EditBox 55, 25, 20, 15, footer_month
  EditBox 130, 25, 20, 15, footer_year
  EditBox 75, 45, 75, 15, worker_number
  ButtonGroup ButtonPressed
    OkButton 20, 95, 50, 15
    CancelButton 85, 95, 50, 15
  Text 5, 10, 55, 10, "Create list from:"
  Text 5, 30, 45, 10, "Footer month:"
  Text 85, 30, 40, 10, "Footer year:"
  Text 5, 50, 65, 10, "Worker number(s):"
  Text 5, 65, 145, 25, "Enter last three digits of each, (ex: x100###). If entering multiple workers, separate each with a comma."
EndDialog


'DISPLAYS DIALOG
Dialog SWKR_list_generator_dialog
If buttonpressed = cancel then stopscript

'CHECKS FOR PASSWORD PROMPT/MAXIS STATUS
transmit
MAXIS_check_function

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

'DEFINES THE EXCEL_ROW VARIABLE FOR WORKING WITH THE SPREADSHEET
excel_row = 2

'OPENS A NEW EXCEL SPREADSHEET
Set objExcel = CreateObject("Excel.Application") 
objExcel.Visible = True  
Set objWorkbook = objExcel.Workbooks.Add() 

'FORMATS THE EXCEL SPREADSHEET WITH THE HEADERS, AND SETS THE COLUMN WIDTH
ObjExcel.Cells(1, 1).Value = "WORKER"
objExcel.Cells(1, 1).Font.Bold = TRUE
ObjExcel.Cells(1, 2).Value = "M#"
objExcel.Cells(1, 2).Font.Bold = TRUE
objExcel.Cells(1, 2).ColumnWidth = 9
ObjExcel.Cells(1, 3).Value = "Name"
objExcel.Cells(1, 3).Font.Bold = TRUE
objExcel.Cells(1, 3).ColumnWidth = 27
ObjExcel.Cells(1, 4).Value = "SWKR name"
objExcel.Cells(1, 4).Font.Bold = TRUE
objExcel.Cells(1, 4).ColumnWidth = 35
ObjExcel.Cells(1, 5).Value = "Copy of Notice?"
objExcel.Cells(1, 5).Font.Bold = TRUE
objExcel.Cells(1, 5).ColumnWidth = 20

'Splitting array for use by the for...next statement
worker_number_array = split(worker_number, ",")

For each worker in worker_number_array

	If trim(worker) = "" then exit for

	worker_ID = worker_county_code & trim(worker)
	
	If REPT_panel = "REPT/ACTV" then 'THE REPT PANEL HAS THE worker NUMBER IN DIFFERENT COLUMNS. THIS WILL DETERMINE THE CORRECT COLUMN FOR THE worker NUMBER TO GO
		worker_ID_col = 13
	Else
		worker_ID_col = 6
	End if  
	EMReadScreen default_worker_number, 3, 21, worker_ID_col 'CHECKING THE CURRENT worker NUMBER. IF IT DOESN'T NEED TO CHANGE IT WON'T. OTHERWISE, THE SCRIPT WILL INPUT THE CORRECT NUMBER.
	If ucase(worker_ID) <> default_worker_number then
		EMWriteScreen worker_ID, 21, worker_ID_col
		transmit
	End if


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
			ObjExcel.Cells(excel_row, 1).Value = worker_ID
			ObjExcel.Cells(excel_row, 2).Value = trim(case_number)
			ObjExcel.Cells(excel_row, 3).Value = trim(client_name)
			excel_row = excel_row + 1
			row = row + 1
		Loop until row = 19 or trim(case_number) = ""
		
		PF8 'going to the next screen
	
	
	Loop until last_page_check = "THIS IS THE LAST PAGE"

Next

'NOW THE SCRIPT IS CHECKING STAT/AREP FOR EACH CASE.----------------------------------------------------------------------------------------------------

excel_row = 2 'Resetting the case row to investigate.

do until ObjExcel.Cells(excel_row, 2).Value = "" 'shuts down when there's no more case numbers
	SWKR_name = "" 'Resetting this variable in case a SWKR cannot be found.
	case_number = ObjExcel.Cells(excel_row, 2).Value 
	If case_number = "" then exit do
	
	'This Do...loop gets back to SELF
	back_to_self
	
	'NAVIGATES TO STAT/SWKR
	call navigate_to_screen("STAT", "SWKR")
	
	'CHECKS FOR ERROR PRONE CASES
	ERRR_screen_check
	
	'NAVIGATES TO SWKR, READS THE NAME AND NOTICE Y/N, AND ADDS TO SPREADSHEET
	EMReadScreen SWKR_name, 34, 6, 32
	swkr_name = replace(swkr_name, "_", "")
	ObjExcel.Cells(excel_row, 4).Value = swkr_name
	EMReadScreen NOTC_Y_N, 1, 15, 63
	If NOTC_Y_N = "_" then NOTC_Y_N = ""
	ObjExcel.Cells(excel_row, 5).Value = NOTC_Y_N
	
	
	excel_row = excel_row + 1 'setting up the script to check the next row.
loop

MsgBox "Success! Your list has been created."

script_end_procedure("")
