Enter file contents here'PLEASE NOTE: this script was designed to run off of the BULK - pull data into Excel script.
'As such, it might not work if ran separately from that.

'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - REPT-PND1 list"
start_time = timer

'DIALOGS-------------------------------------------------------------------------------------------------------------
BeginDialog pull_REPT_data_into_excel_dialog, 0, 0, 286, 120, "Pull REPT PND1 data into Excel dialog"
  EditBox 150, 20, 130, 15, worker_number
  CheckBox 70, 65, 150, 10, "Check here to run this query county-wide.", all_workers_check
  ButtonGroup ButtonPressed
    OkButton 175, 100, 50, 15
    CancelButton 230, 100, 50, 15
  Text 5, 25, 65, 10, "Worker(s) to check:"
  Text 5, 80, 210, 20, "NOTE: running queries county-wide can take a significant amount of time and resources. This should be done after hours."
  Text 5, 5, 125, 10, "***PULL REPT DATA INTO EXCEL***"
  Text 5, 40, 210, 20, "Enter last 3 digits of your workers' x1 numbers (ex: x100###), separated by a comma."
EndDialog

'THE SCRIPT-----------------------------------------------------------------------------------------------------------

'Connects to BlueZone
EMConnect ""

'Shows dialog
Dialog pull_rept_data_into_Excel_dialog
If buttonpressed = cancel then stopscript

'Starting the query start time (for the query runtime at the end)
query_start_time = timer

'Checking for MAXIS
EMReadScreen MAXIS_check, 5, 1, 39
If MAXIS_check <> "MAXIS" and MAXIS_check <> "AXIS " then script_end_procedure("You appear to be locked out of MAXIS.")

'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add() 
objExcel.DisplayAlerts = True


'Setting the first 4 col as worker, case number, name, and APPL date
ObjExcel.Cells(1, 1).Value = "WORKER"
objExcel.Cells(1, 1).Font.Bold = TRUE
ObjExcel.Cells(1, 2).Value = "CASE NUMBER"
objExcel.Cells(1, 2).Font.Bold = TRUE
ObjExcel.Cells(1, 3).Value = "NAME"
objExcel.Cells(1, 3).Font.Bold = TRUE
ObjExcel.Cells(1, 4).Value = "APPL DATE"
objExcel.Cells(1, 4).Font.Bold = TRUE
ObjExcel.Cells(1, 5).Value = "NBR DAYS PENDING"
objExcel.Cells(1, 5).Font.Bold = TRUE
ObjExcel.Cells(1, 6).Value = "STOP AUTODENY"
objExcel.Cells(1, 6).Font.Bold = TRUE
ObjExcel.Cells(1, 7).Value = "AUTODENY DATE"
objExcel.Cells(1, 7).Font.Bold = TRUE

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
	Call navigate_to_screen("rept", "pnd1")
	EMWriteScreen worker, 21, 13
	transmit

	'Skips workers with no info
	EMReadScreen has_content_check, 1, 7, 5
	If has_content_check <> " " then

		'Grabbing each case number on screen
		Do
			'Set variable for next do...loop
			MAXIS_row = 7
			Do			
				EMReadScreen case_number, 4, MAXIS_row, 7			'Reading case number
				EMReadScreen client_name, 25, MAXIS_row, 13		'Reading client name
				EMReadScreen appl_date, 8, MAXIS_row, 41		      'Reading application date
				EMReadScreen nbr_days_pending, 2, MAXIS_row, 55		'Reading nbr days pending
				EMReadScreen stop_autodeny, 1, MAXIS_row, 65		'Reading stop autodeny
				EMReadScreen autodeny_date, 8, MAXIS_row, 72		'Reading autodeny date

				'Doing this because sometimes BlueZone registers a "ghost" of previous data when the script runs. This checks against an array and stops if we've seen this one before.
				If trim(case_number) <> "" and instr(all_case_numbers_array, case_number) <> 0 then exit do
				all_case_numbers_array = trim(all_case_numbers_array & " " & case_number)

				If case_number = "        " then exit do			'Exits do if we reach the end

				MAXIS_row = MAXIS_row + 1
				add_case_info_to_Excel = ""	'Blanking out variable
				autoclose_string = ""		'Blanking out variable
				case_number = ""			'Blanking out variable
			Loop until MAXIS_row = 19
			PF8
		Loop until last_page_check = "THIS IS THE LAST PAGE"
	End if
next

'logging usage state
script_end_procedure("")

col_to_use = col_to_use + 2	'Doing two because the wrap-up is two columns
row_to_use = 3			'For the individual program-breakdown of info

'Query date/time/runtime info
objExcel.Cells(1, col_to_use - 1).Font.Bold = TRUE
objExcel.Cells(2, col_to_use - 1).Font.Bold = TRUE
ObjExcel.Cells(1, col_to_use - 1).Value = "Query date and time:"	'Goes back one, as this is on the next row
ObjExcel.Cells(1, col_to_use).Value = now
ObjExcel.Cells(2, col_to_use - 1).Value = "Query runtime (in seconds):"	'Goes back one, as this is on the next row
ObjExcel.Cells(2, col_to_use).Value = timer - query_start_time

'Autofitting columns
For col_to_autofit = 1 to col_to_use
	ObjExcel.columns(col_to_autofit).AutoFit()
Next
