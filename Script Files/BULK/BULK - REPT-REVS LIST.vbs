'PLEASE NOTE: this script was designed to run off of the BULK - pull data into Excel script.
'As such, it might not work if ran separately from that.

'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - REPT-REVS LIST.vbs"
start_time = timer

'LOADING ROUTINE FUNCTIONS FROM GITHUB REPOSITORY---------------------------------------------------------------------------
url = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a URL
req.open "GET", url, FALSE									'Attempts to open the URL
req.send													'Sends request
IF req.Status = 200 THEN									'200 means great success
	Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
	Execute req.responseText								'Executes the script code
ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
	MsgBox 	"Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_ 
			vbCr & _
			"Before contacting Veronica Cary, please check to make sure you can load the main page at www.GitHub.com." & vbCr &_
			vbCr & _
			"If you can reach GitHub.com, but this script still does not work, ask an alpha user to contact Veronica Cary and provide the following information:" & vbCr &_
			vbTab & "- The name of the script you are running." & vbCr &_
			vbTab & "- Whether or not the script is ""erroring out"" for any other users." & vbCr &_
			vbTab & "- The name and email for an employee from your IT department," & vbCr & _
			vbTab & vbTab & "responsible for network issues." & vbCr &_
			vbTab & "- The URL indicated below (a screenshot should suffice)." & vbCr &_
			vbCr & _
			"Veronica will work with your IT department to try and solve this issue, if needed." & vbCr &_ 
			vbCr &_
			"URL: " & url
			script_end_procedure("Script ended due to error connecting to GitHub.")
END IF

'DIALOGS-----------------------------------------------------------
BeginDialog pull_REPT_data_into_excel_dialog, 0, 0, 286, 120, "Pull REPT data into Excel dialog"
  EditBox 140, 20, 140, 15, worker_number
  CheckBox 70, 65, 150, 10, "Check here to run this query county-wide.", all_workers_check
  CheckBox 10, 35, 40, 10, "SNAP?", SNAP_check
  CheckBox 10, 50, 50, 10, "Cash/GRH?", cash_check
  CheckBox 10, 65, 40, 10, "HC?", HC_check
  ButtonGroup ButtonPressed
    OkButton 175, 100, 50, 15
    CancelButton 230, 100, 50, 15
  GroupBox 5, 20, 60, 60, "Progs to scan"
  Text 70, 25, 65, 10, "Worker(s) to check:"
  Text 70, 80, 210, 20, "NOTE: running queries county-wide can take a significant amount of time and resources. This should be done after hours."
  Text 80, 5, 125, 10, "***PULL REPT DATA INTO EXCEL***"
  Text 70, 40, 210, 20, "Enter last 3 digits of your workers' x1 numbers (ex: x100###), separated by a comma."
EndDialog

'THE SCRIPT---------------------------------------------------

'Asks if we want to navigate to current month + 2, which REVS is unique in that it can show this screen
current_month_plus_2 = MsgBox("Navigate to current month + 2 for REVS?", vbYesNo)
If current_month_plus_2 = vbCancel then stopscript
If current_month_plus_2 = vbYes then current_month_plus_2 = True
If current_month_plus_2 = vbNo then current_month_plus_2 = False

'Determining what current month + 2 is
future_footer_month = datepart("m", dateadd("m", 2, date))
If len(future_footer_month) = 1 then future_footer_month = "0" & future_footer_month
future_footer_year = right(datepart("yyyy", dateadd("m", 2, date)), 2)

'Connects to BlueZone
EMConnect ""

'Shows dialog
Dialog pull_rept_data_into_Excel_dialog
If buttonpressed = cancel then stopscript

'Starting the query start time (for the query runtime at the end)
query_start_time = timer

'Checking for MAXIS
PF3
EMReadScreen MAXIS_check, 5, 1, 39
If MAXIS_check <> "MAXIS" and MAXIS_check <> "AXIS " then script_end_procedure("You appear to be locked out of MAXIS.")

'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add() 
objExcel.DisplayAlerts = True

'Setting the first 3 col as worker, case number, and name
ObjExcel.Cells(1, 1).Value = "WORKER"
objExcel.Cells(1, 1).Font.Bold = TRUE
ObjExcel.Cells(1, 2).Value = "CASE NUMBER"
objExcel.Cells(1, 2).Font.Bold = TRUE
ObjExcel.Cells(1, 3).Value = "NAME"
objExcel.Cells(1, 3).Font.Bold = TRUE


'Figuring out what to put in each Excel col. To add future variables to this, add the checkbox variables below and copy/paste the same code!
'	Below, use the "[blank]_col" variable to recall which col you set for which option.
col_to_use = 4 'Starting with 4 because cols 1-3 are already used

If SNAP_check = checked then
	ObjExcel.Cells(1, col_to_use).Value = "SNAP?"
	objExcel.Cells(1, col_to_use).Font.Bold = TRUE
	snap_actv_col = col_to_use
	col_to_use = col_to_use + 1
	SNAP_letter_col = convert_digit_to_excel_column(snap_actv_col)
End if
If cash_check = checked then
	ObjExcel.Cells(1, col_to_use).Value = "CASH?"
	objExcel.Cells(1, col_to_use).Font.Bold = TRUE
	cash_actv_col = col_to_use
	col_to_use = col_to_use + 1
	cash_letter_col = convert_digit_to_excel_column(cash_actv_col)
End if
If HC_check = checked then
	ObjExcel.Cells(1, col_to_use).Value = "HC?"		'First does HC col, then does exempt IR col, then MAGI col
	objExcel.Cells(1, col_to_use).Font.Bold = TRUE
	HC_actv_col = col_to_use
	col_to_use = col_to_use + 1
	HC_letter_col = convert_digit_to_excel_column(HC_actv_col)
	ObjExcel.Cells(1, col_to_use).Value = "EXEMPT HC IR?"	'Exempt IR col
	objExcel.Cells(1, col_to_use).Font.Bold = TRUE
	exempt_IR_col = col_to_use
	col_to_use = col_to_use + 1
	exempt_IR_letter_col = convert_digit_to_excel_column(exempt_IR_col)
	ObjExcel.Cells(1, col_to_use).Value = "MAGI?"		'Here's that MAGI col
	objExcel.Cells(1, col_to_use).Font.Bold = TRUE
	MAGI_col = col_to_use
	col_to_use = col_to_use + 1
	MAGI_letter_col = convert_digit_to_excel_column(MAGI_col)
End if

'Does these two columns after the others, because they appear that way in the screen, but are always used.
'Only does these if current_month_plus_2 = False, as you cannot have a revw rec'd or an interview date before this point.
If current_month_plus_2 = False then
	ObjExcel.Cells(1, col_to_use).Value = "DATE REVW REC'D"
	objExcel.Cells(1, col_to_use).Font.Bold = TRUE
	revw_recd_date_col = col_to_use
	col_to_use = col_to_use + 1
	revw_recd_date_letter_col = convert_digit_to_excel_column(revw_recd_date_col)
	ObjExcel.Cells(1, col_to_use).Value = "INTERVIEW DATE"
	objExcel.Cells(1, col_to_use).Font.Bold = TRUE
	interview_date_col = col_to_use
	col_to_use = col_to_use + 1
	interview_date_letter_col = convert_digit_to_excel_column(interview_date_col)
End if

'If all workers are selected, the script will go to REPT/USER, and load all of the workers into an array. Otherwise it'll create a single-object "array" just for simplicity of code.
If all_workers_check = checked then
	call create_array_of_all_active_x_numbers_in_county(worker_array, two_digit_county_code)
Else
	x1s_from_dialog = split(worker_number, ",")	'Splits the worker array based on commas

	'Need to add the worker_county_code to each one
	For each x1_number in x1s_from_dialog
		If worker_array = "" then
			worker_array = worker_county_code & trim(replace(ucase(x1_number), worker_county_code, ""))		'replaces worker_county_code if found in the typed x1 number
		Else
			worker_array = worker_array & ", " & worker_county_code & trim(replace(ucase(x1_number), worker_county_code, "")) 'replaces worker_county_code if found in the typed x1 number
		End if
	Next

	'Split worker_array
	worker_array = split(worker_array, ", ")
End if

'Setting the variable for what's to come
excel_row = 2

For each worker in worker_array
	back_to_self	'Does this to prevent "ghosting" where the old info shows up on the new screen for some reason
	Call navigate_to_screen("rept", "revs")
	EMWriteScreen worker, 21, 6
	transmit

	'If current_month_plus_2 is selected, it pops that month into the footer month area.
	If current_month_plus_2 = True then
		EMWriteScreen future_footer_month, 20, 55
		EMWriteScreen future_footer_year, 20, 58
		transmit
	End if

	'Skips workers with no info
	EMReadScreen has_content_check, 1, 7, 8
	If has_content_check <> " " then

		'Grabbing each case number on screen
		Do
			'Set variable for next do...loop
			MAXIS_row = 7
			Do			
				EMReadScreen case_number, 8, MAXIS_row, 6			'Reading case number
				EMReadScreen client_name, 15, MAXIS_row, 16		'Reading client name
				EMReadScreen cash_status, 1, MAXIS_row, 35		'Reading cash status
				EMReadScreen SNAP_status, 1, MAXIS_row, 45		'Reading SNAP status
				EMReadScreen HC_status, 1, MAXIS_row, 49			'Reading HC status
				EMReadScreen exempt_IR_status, 1, MAXIS_row, 51		'Reading exempt IR status
				EMReadScreen MAGI_status, 8, MAXIS_row, 54		'Reading MAGI status
				EMReadScreen revw_recd_date, 8, MAXIS_row, 62		'Reading review received date
				EMReadScreen interview_date, 8, MAXIS_row, 72		'Reading interview date

				'Doing this because sometimes BlueZone registers a "ghost" of previous data when the script runs. This checks against an array and stops if we've seen this one before.
				If trim(case_number) <> "" and instr(all_case_numbers_array, case_number) <> 0 then exit do
				all_case_numbers_array = trim(all_case_numbers_array & " " & case_number)

				If case_number = "        " then exit do			'Exits do if we reach the end

				'For some goofy reason the dash key shows up instead of the space key. No clue why. This will turn them into null variables.
				If cash_status = "-" then cash_status = ""
				If SNAP_status = "-" then SNAP_status = ""
				If HC_status = "-" then HC_status = ""

				'The asterisk in the exempt IR column messes up the formula for Excel. Replacing with the word "exempt"
				If exempt_IR_status = "*" then exempt_IR_status = "exempt"

				'Using if...thens to decide if a case should be added (status isn't blank and respective box is checked)
				If trim(cash_status) <> "" and cash_check = checked then add_case_info_to_Excel = True
				If trim(SNAP_status) <> "" and SNAP_check = checked then add_case_info_to_Excel = True
				If trim(HC_status) <> "" and HC_check = checked then add_case_info_to_Excel = True

				'Cleaning up the blank revw_recd_date and interview_date variables
				revw_recd_date = trim(replace(revw_recd_date, "__ __ __", ""))
				interview_date = trim(replace(interview_date, "__ __ __", ""))

				'Adding the case to Excel
				If add_case_info_to_Excel = True then 
					ObjExcel.Cells(excel_row, 1).Value = worker
					ObjExcel.Cells(excel_row, 2).Value = case_number
					ObjExcel.Cells(excel_row, 3).Value = client_name
					If current_month_plus_2 = False then
						ObjExcel.Cells(excel_row, revw_recd_date_col).Value = replace(revw_recd_date, " ", "/")
						ObjExcel.Cells(excel_row, interview_date_col).Value = replace(interview_date, " ", "/")
					End if
					If SNAP_check = checked then ObjExcel.Cells(excel_row, snap_actv_col).Value = trim(SNAP_status)
					If cash_check = checked then ObjExcel.Cells(excel_row, cash_actv_col).Value = trim(cash_status)
					If HC_check = checked then 
						ObjExcel.Cells(excel_row, HC_actv_col).Value = trim(HC_status)
						ObjExcel.Cells(excel_row, exempt_IR_col).Value = trim(exempt_IR_status)
						ObjExcel.Cells(excel_row, MAGI_col).Value = trim(MAGI_status)
					End if
					excel_row = excel_row + 1
				End if
				MAXIS_row = MAXIS_row + 1
				add_case_info_to_Excel = ""	'Blanking out variable
				case_number = ""			'Blanking out variable
			Loop until MAXIS_row = 19
			PF8
			EMReadScreen last_page_check, 21, 24, 2	'checking to see if we're at the end
		Loop until last_page_check = "THIS IS THE LAST PAGE"
	End if
next

col_to_use = col_to_use + 2	'Doing two because the wrap-up is two columns
row_to_use = 3			'Declaring here before the following if...then statements

'Query date/time/runtime info
objExcel.Cells(1, col_to_use - 1).Font.Bold = TRUE
objExcel.Cells(2, col_to_use - 1).Font.Bold = TRUE
ObjExcel.Cells(1, col_to_use - 1).Value = "Query date and time:"	'Goes back one, as this is on the next row
ObjExcel.Cells(1, col_to_use).Value = now
ObjExcel.Cells(2, col_to_use - 1).Value = "Query runtime (in seconds):"	'Goes back one, as this is on the next row
ObjExcel.Cells(2, col_to_use).Value = timer - query_start_time

'SNAP info
If SNAP_check = checked then	
	ObjExcel.Cells(row_to_use, col_to_use - 1).Value = "SNAP cases with unapproved review:"	'Row header
	objExcel.Cells(row_to_use, col_to_use - 1).Font.Bold = TRUE						'Row header should be bold
	ObjExcel.Cells(row_to_use, col_to_use).Value = "=COUNTA(" & SNAP_letter_col & ":" & SNAP_letter_col & ") - 1"	'Excel formula
	ObjExcel.Cells(row_to_use + 1, col_to_use - 1).Value = "Percentage of SNAP cases coded " & chr(34) & "N" & chr(34) & ":"	'Row header
	objExcel.Cells(row_to_use + 1, col_to_use - 1).Font.Bold = TRUE								'Row header should be bold
	ObjExcel.Cells(row_to_use + 1, col_to_use).Value = "=COUNTIFS(" & SNAP_letter_col & ":" & SNAP_letter_col & ", " & chr(34) & "N" & chr(34) & ")/(COUNTA(" & SNAP_letter_col & ":" & SNAP_letter_col & ") - 1)" 'Excel formula
	ObjExcel.Cells(row_to_use + 1, col_to_use).NumberFormat = "0.00%"		'Formula should be percent
	ObjExcel.Cells(row_to_use + 2, col_to_use - 1).Value = "Percentage of SNAP cases coded " & chr(34) & "I" & chr(34) & ":"	'Row header
	objExcel.Cells(row_to_use + 2, col_to_use - 1).Font.Bold = TRUE								'Row header should be bold
	ObjExcel.Cells(row_to_use + 2, col_to_use).Value = "=COUNTIFS(" & SNAP_letter_col & ":" & SNAP_letter_col & ", " & chr(34) & "I" & chr(34) & ")/(COUNTA(" & SNAP_letter_col & ":" & SNAP_letter_col & ") - 1)" 'Excel formula
	ObjExcel.Cells(row_to_use + 2, col_to_use).NumberFormat = "0.00%"		'Formula should be percent
	ObjExcel.Cells(row_to_use + 3, col_to_use - 1).Value = "Percentage of SNAP cases coded " & chr(34) & "U" & chr(34) & ":"	'Row header
	objExcel.Cells(row_to_use + 3, col_to_use - 1).Font.Bold = TRUE								'Row header should be bold
	ObjExcel.Cells(row_to_use + 3, col_to_use).Value = "=COUNTIFS(" & SNAP_letter_col & ":" & SNAP_letter_col & ", " & chr(34) & "U" & chr(34) & ")/(COUNTA(" & SNAP_letter_col & ":" & SNAP_letter_col & ") - 1)" 'Excel formula
	ObjExcel.Cells(row_to_use + 3, col_to_use).NumberFormat = "0.00%"		'Formula should be percent
	row_to_use = row_to_use + 4	'It's four rows we jump, because the SNAP stat takes up four rows
End if

'HC info
If HC_check = checked then	
	ObjExcel.Cells(row_to_use, col_to_use - 1).Value = "HC cases with unapproved review:"	'Row header
	objExcel.Cells(row_to_use, col_to_use - 1).Font.Bold = TRUE						'Row header should be bold
	ObjExcel.Cells(row_to_use, col_to_use).Value = "=COUNTA(" & HC_letter_col & ":" & HC_letter_col & ") - 1"	'Excel formula
	ObjExcel.Cells(row_to_use + 1, col_to_use - 1).Value = "Percentage of HC cases coded " & chr(34) & "N" & chr(34) & ":"	'Row header
	objExcel.Cells(row_to_use + 1, col_to_use - 1).Font.Bold = TRUE								'Row header should be bold
	ObjExcel.Cells(row_to_use + 1, col_to_use).Value = "=COUNTIFS(" & HC_letter_col & ":" & HC_letter_col & ", " & chr(34) & "N" & chr(34) & ")/(COUNTA(" & HC_letter_col & ":" & HC_letter_col & ") - 1)" 'Excel formula
	ObjExcel.Cells(row_to_use + 1, col_to_use).NumberFormat = "0.00%"		'Formula should be percent
	ObjExcel.Cells(row_to_use + 2, col_to_use - 1).Value = "Percentage of HC cases coded " & chr(34) & "I" & chr(34) & ":"	'Row header
	objExcel.Cells(row_to_use + 2, col_to_use - 1).Font.Bold = TRUE								'Row header should be bold
	ObjExcel.Cells(row_to_use + 2, col_to_use).Value = "=COUNTIFS(" & HC_letter_col & ":" & HC_letter_col & ", " & chr(34) & "I" & chr(34) & ")/(COUNTA(" & HC_letter_col & ":" & HC_letter_col & ") - 1)" 'Excel formula
	ObjExcel.Cells(row_to_use + 2, col_to_use).NumberFormat = "0.00%"		'Formula should be percent
	ObjExcel.Cells(row_to_use + 3, col_to_use - 1).Value = "Percentage of HC cases coded " & chr(34) & "U" & chr(34) & ":"	'Row header
	objExcel.Cells(row_to_use + 3, col_to_use - 1).Font.Bold = TRUE								'Row header should be bold
	ObjExcel.Cells(row_to_use + 3, col_to_use).Value = "=COUNTIFS(" & HC_letter_col & ":" & HC_letter_col & ", " & chr(34) & "U" & chr(34) & ")/(COUNTA(" & HC_letter_col & ":" & HC_letter_col & ") - 1)" 'Excel formula
	ObjExcel.Cells(row_to_use + 3, col_to_use).NumberFormat = "0.00%"		'Formula should be percent
	ObjExcel.Cells(row_to_use + 4, col_to_use - 1).Value = "Percentage of HC cases with exempt IRs:"		'Row header
	objExcel.Cells(row_to_use + 4, col_to_use - 1).Font.Bold = TRUE								'Row header should be bold
	ObjExcel.Cells(row_to_use + 4, col_to_use).Value = "=COUNTIFS(" & exempt_IR_letter_col & ":" & exempt_IR_letter_col & ", " & chr(34) & "exempt" & chr(34) & ")/(COUNTA(" & HC_letter_col & ":" & HC_letter_col & ") - 1)" 'Excel formula
	ObjExcel.Cells(row_to_use + 4, col_to_use).NumberFormat = "0.00%"		'Formula should be percent
	row_to_use = row_to_use + 5	'It's four five we jump, because the HC stat takes up five rows
End if

'cash info
If cash_check = checked then	
	ObjExcel.Cells(row_to_use, col_to_use - 1).Value = "Cash cases with unapproved review:"	'Row header
	objExcel.Cells(row_to_use, col_to_use - 1).Font.Bold = TRUE						'Row header should be bold
	ObjExcel.Cells(row_to_use, col_to_use).Value = "=COUNTA(" & cash_letter_col & ":" & cash_letter_col & ") - 1"	'Excel formula
	ObjExcel.Cells(row_to_use + 1, col_to_use - 1).Value = "Percentage of cash cases coded " & chr(34) & "N" & chr(34) & ":"	'Row header
	objExcel.Cells(row_to_use + 1, col_to_use - 1).Font.Bold = TRUE								'Row header should be bold
	ObjExcel.Cells(row_to_use + 1, col_to_use).Value = "=COUNTIFS(" & cash_letter_col & ":" & cash_letter_col & ", " & chr(34) & "N" & chr(34) & ")/(COUNTA(" & cash_letter_col & ":" & cash_letter_col & ") - 1)" 'Excel formula
	ObjExcel.Cells(row_to_use + 1, col_to_use).NumberFormat = "0.00%"		'Formula should be percent
	ObjExcel.Cells(row_to_use + 2, col_to_use - 1).Value = "Percentage of cash cases coded " & chr(34) & "I" & chr(34) & ":"	'Row header
	objExcel.Cells(row_to_use + 2, col_to_use - 1).Font.Bold = TRUE								'Row header should be bold
	ObjExcel.Cells(row_to_use + 2, col_to_use).Value = "=COUNTIFS(" & cash_letter_col & ":" & cash_letter_col & ", " & chr(34) & "I" & chr(34) & ")/(COUNTA(" & cash_letter_col & ":" & cash_letter_col & ") - 1)" 'Excel formula
	ObjExcel.Cells(row_to_use + 2, col_to_use).NumberFormat = "0.00%"		'Formula should be percent
	ObjExcel.Cells(row_to_use + 3, col_to_use - 1).Value = "Percentage of cash cases coded " & chr(34) & "U" & chr(34) & ":"	'Row header
	objExcel.Cells(row_to_use + 3, col_to_use - 1).Font.Bold = TRUE								'Row header should be bold
	ObjExcel.Cells(row_to_use + 3, col_to_use).Value = "=COUNTIFS(" & cash_letter_col & ":" & cash_letter_col & ", " & chr(34) & "U" & chr(34) & ")/(COUNTA(" & cash_letter_col & ":" & cash_letter_col & ") - 1)" 'Excel formula
	ObjExcel.Cells(row_to_use + 3, col_to_use).NumberFormat = "0.00%"		'Formula should be percent
	row_to_use = row_to_use + 4	'It's four rows we jump, because the cash stat takes up four rows
End if

'Autofitting columns
For col_to_autofit = 1 to col_to_use
	ObjExcel.columns(col_to_autofit).AutoFit()
Next

'Logging usage stats
script_end_procedure("")
