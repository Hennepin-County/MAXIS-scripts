'PLEASE NOTE: this script was designed to run off of the BULK - pull data into Excel script.
'As such, it might not work if ran separately from that.

'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - REPT-PND2 LIST.vbs"
start_time = timer

'LOADING ROUTINE FUNCTIONS---------------------------------------------------------------
url = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER FUNCTIONS LIBRARY.vbs"
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

'DIALOG-----------------------------------------------------------------------------------------
BeginDialog pull_REPT_data_into_excel_dialog, 0, 0, 286, 120, "Pull REPT data into Excel dialog"
  EditBox 150, 20, 130, 15, worker_number
  CheckBox 70, 65, 150, 10, "Check here to run this query county-wide.", all_workers_check
  CheckBox 10, 35, 40, 10, "SNAP?", SNAP_check
  CheckBox 10, 50, 40, 10, "Cash?", cash_check
  CheckBox 10, 65, 40, 10, "HC?", HC_check
  CheckBox 10, 80, 40, 10, "EA?", EA_check
  CheckBox 10, 95, 40, 10, "GRH?", GRH_check
  ButtonGroup ButtonPressed
    OkButton 175, 100, 50, 15
    CancelButton 230, 100, 50, 15
  GroupBox 5, 20, 60, 90, "Progs to scan"
  Text 70, 25, 65, 10, "Worker(s) to check:"
  Text 70, 80, 210, 20, "NOTE: running queries county-wide can take a significant amount of time and resources. This should be done after hours."
  Text 80, 5, 125, 10, "***PULL REPT DATA INTO EXCEL***"
  Text 70, 40, 210, 20, "Enter last 3 digits of your workers' x1 numbers (ex: x100###), separated by a comma."
EndDialog

'THE SCRIPT-------------------------------------------------------------------------

'Connects to BlueZone
EMConnect ""

'Dialog asks what stats are being pulled
Dialog pull_REPT_data_into_excel_dialog
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

'Changes name of Excel sheet to "Case information"
ObjExcel.ActiveSheet.Name = "Case information"

'Setting the first 4 col as worker, case number, name, and APPL date
ObjExcel.Cells(1, 1).Value = "WORKER"
objExcel.Cells(1, 1).Font.Bold = TRUE
ObjExcel.Cells(1, 2).Value = "CASE NUMBER"
objExcel.Cells(1, 2).Font.Bold = TRUE
ObjExcel.Cells(1, 3).Value = "NAME"
objExcel.Cells(1, 3).Font.Bold = TRUE
ObjExcel.Cells(1, 4).Value = "APPL DATE"
objExcel.Cells(1, 4).Font.Bold = TRUE
ObjExcel.Cells(1, 5).Value = "DAYS PENDING"	
objExcel.Cells(1, 5).Font.Bold = TRUE

'Figuring out what to put in each Excel col. To add future variables to this, add the checkbox variables below and copy/paste the same code!
'	Below, use the "[blank]_col" variable to recall which col you set for which option.
col_to_use = 6 'Starting with 6 because cols 1-5 are already used

If SNAP_check = checked then
	ObjExcel.Cells(1, col_to_use).Value = "SNAP?"
	objExcel.Cells(1, col_to_use).Font.Bold = TRUE
	snap_pends_col = col_to_use
	col_to_use = col_to_use + 1
	SNAP_letter_col = convert_digit_to_excel_column(snap_pends_col)
End if
If cash_check = checked then
	ObjExcel.Cells(1, col_to_use).Value = "CASH?"
	objExcel.Cells(1, col_to_use).Font.Bold = TRUE
	cash_pends_col = col_to_use
	col_to_use = col_to_use + 1
	cash_letter_col = convert_digit_to_excel_column(cash_pends_col)
End if
If HC_check = checked then
	ObjExcel.Cells(1, col_to_use).Value = "HC?"
	objExcel.Cells(1, col_to_use).Font.Bold = TRUE
	HC_pends_col = col_to_use
	col_to_use = col_to_use + 1
	HC_letter_col = convert_digit_to_excel_column(HC_pends_col)
End if
If EA_check = checked then
	ObjExcel.Cells(1, col_to_use).Value = "EA?"
	objExcel.Cells(1, col_to_use).Font.Bold = TRUE
	EA_pends_col = col_to_use
	col_to_use = col_to_use + 1
	EA_letter_col = convert_digit_to_excel_column(EA_pends_col)
End if
If GRH_check = checked then
	ObjExcel.Cells(1, col_to_use).Value = "GRH?"
	objExcel.Cells(1, col_to_use).Font.Bold = TRUE
	GRH_pends_col = col_to_use
	col_to_use = col_to_use + 1
	GRH_letter_col = convert_digit_to_excel_column(GRH_pends_col)
End if

'Setting the variable for what's to come
excel_row = 2

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

For each worker in worker_array
	back_to_self	'Does this to prevent "ghosting" where the old info shows up on the new screen for some reason
	Call navigate_to_screen("rept", "pnd2")
	EMWriteScreen worker, 21, 13
	transmit

	'Skips workers with no info
	EMReadScreen has_content_check, 6, 3, 74
	If has_content_check <> "0 Of 0" then

		'Grabbing each case number on screen
		Do
			MAXIS_row = 7
			Do
				EMReadScreen case_number, 8, MAXIS_row, 5			'Reading case number
				EMReadScreen client_name, 22, MAXIS_row, 16		'Reading client name
				EMReadScreen APPL_date, 8, MAXIS_row, 38			'Reading application date
				EMReadScreen days_pending, 4, MAXIS_row, 49		'Reading days pending
				EMReadScreen cash_status, 1, MAXIS_row, 54		'Reading cash status
				EMReadScreen SNAP_status, 1, MAXIS_row, 62		'Reading SNAP status
				EMReadScreen HC_status, 1, MAXIS_row, 65			'Reading HC status
				EMReadScreen EA_status, 1, MAXIS_row, 68			'Reading EA status
				EMReadScreen GRH_status, 1, MAXIS_row, 72			'Reading GRH status

				'Doing this because sometimes BlueZone registers a "ghost" of previous data when the script runs. This checks against an array and stops if we've seen this one before.
				If trim(case_number) <> "" and (instr(all_case_numbers_array, case_number) <> 0 and client_name <> " ADDITIONAL APP       ") then exit do
				all_case_numbers_array = trim(all_case_numbers_array & " " & case_number)

				If case_number = "        " then exit do			'Exits do if we reach the end

				'Cleaning up each program's status
				SNAP_status = trim(replace(SNAP_status, "_", ""))
				cash_status = trim(replace(cash_status, "_", ""))
				HC_status = trim(replace(HC_status, "_", ""))
				EA_status = trim(replace(EA_status, "_", ""))
				GRH_status = trim(replace(GRH_status, "_", ""))

				'Using if...thens to decide if a case should be added (status isn't blank and respective box is checked)
				If SNAP_status <> "" and SNAP_check = checked then add_case_info_to_Excel = True
				If cash_status <> "" and cash_check = checked then add_case_info_to_Excel = True
				If HC_status <> "" and HC_check = checked then add_case_info_to_Excel = True
				If EA_status <> "" and EA_check = checked then add_case_info_to_Excel = True
				If GRH_status <> "" and GRH_check = checked then add_case_info_to_Excel = True

				If add_case_info_to_Excel = True then 
					ObjExcel.Cells(excel_row, 1).Value = worker
					ObjExcel.Cells(excel_row, 2).Value = case_number
					ObjExcel.Cells(excel_row, 3).Value = client_name
					ObjExcel.Cells(excel_row, 4).Value = replace(APPL_date, " ", "/")
					ObjExcel.Cells(excel_row, 5).Value = abs(days_pending)
					If SNAP_check = checked then ObjExcel.Cells(excel_row, snap_pends_col).Value = SNAP_status
					If cash_check = checked then ObjExcel.Cells(excel_row, cash_pends_col).Value = cash_status
					If HC_check = checked then ObjExcel.Cells(excel_row, HC_pends_col).Value = HC_status
					If EA_check = checked then ObjExcel.Cells(excel_row, EA_pends_col).Value = EA_status
					If GRH_check = checked then ObjExcel.Cells(excel_row, GRH_pends_col).Value = GRH_status
					excel_row = excel_row + 1
				End if
				MAXIS_row = MAXIS_row + 1
				add_case_info_to_Excel = ""	'Blanking out variable
				case_number = ""			'Blanking out variable
			Loop until MAXIS_row = 19
			PF8
			EMReadScreen last_page_check, 21, 24, 2
		Loop until last_page_check = "THIS IS THE LAST PAGE"
	End if
next

'Resetting excel_row variable, now we need to start looking people up
excel_row = 2 

'Setting variables for the stats area
col_to_use = col_to_use + 2	'Doing two because the wrap-up is two columns
row_to_use = 3			'Setting variable for the if...thens below

'Query date/time/runtime info
objExcel.Cells(1, col_to_use - 1).Font.Bold = TRUE
objExcel.Cells(2, col_to_use - 1).Font.Bold = TRUE
ObjExcel.Cells(1, col_to_use - 1).Value = "Query date and time:"	'Goes back one, as this is on the next row
ObjExcel.Cells(1, col_to_use).Value = now
ObjExcel.Cells(2, col_to_use - 1).Value = "Query runtime (in seconds):"	'Goes back one, as this is on the next row
ObjExcel.Cells(2, col_to_use).Value = timer - query_start_time

'SNAP info
If SNAP_check = checked then	
	ObjExcel.Cells(row_to_use, col_to_use - 1).Value = "SNAP cases pending over 30 days:"	'Row header
	objExcel.Cells(row_to_use, col_to_use - 1).Font.Bold = TRUE						'Row header should be bold
	ObjExcel.Cells(row_to_use, col_to_use).Value = "=COUNTIFS(E:E, " & Chr(34) & ">30" & Chr(34) & ", " & SNAP_letter_col & ":" & SNAP_letter_col & ", " & is_not_blank_excel_string & ")"	'Excel formula
	ObjExcel.Cells(row_to_use + 1, col_to_use - 1).Value = "Percentage of SNAP cases pending over 30 days:"	'Row header
	objExcel.Cells(row_to_use + 1, col_to_use - 1).Font.Bold = TRUE								'Row header should be bold
	ObjExcel.Cells(row_to_use + 1, col_to_use).Value = "=(COUNTIFS(E:E, " & Chr(34) & ">30" & Chr(34) & ", " & SNAP_letter_col & ":" & SNAP_letter_col & ", " & is_not_blank_excel_string & "))/(COUNTIF(" & SNAP_letter_col & ":" & SNAP_letter_col & ", " & is_not_blank_excel_string & ") -1)" 'Excel formula
	ObjExcel.Cells(row_to_use + 1, col_to_use).NumberFormat = "0.00%"		'Formula should be percent
	row_to_use = row_to_use + 2	'It's two rows we jump, because the SNAP stat takes up two rows
End if

'cash info
If cash_check = checked then	
	ObjExcel.Cells(row_to_use, col_to_use - 1).Value = "Cash cases pending over 30 days:"	'Row header
	objExcel.Cells(row_to_use, col_to_use - 1).Font.Bold = TRUE						'Row header should be bold
	ObjExcel.Cells(row_to_use, col_to_use).Value = "=COUNTIFS(E:E, " & Chr(34) & ">30" & Chr(34) & ", " & cash_letter_col & ":" & cash_letter_col & ", " & is_not_blank_excel_string & ")"	'Excel formula
	ObjExcel.Cells(row_to_use + 1, col_to_use - 1).Value = "Percentage of cash cases pending over 30 days:"	'Row header
	objExcel.Cells(row_to_use + 1, col_to_use - 1).Font.Bold = TRUE								'Row header should be bold
	ObjExcel.Cells(row_to_use + 1, col_to_use).Value = "=(COUNTIFS(E:E, " & Chr(34) & ">30" & Chr(34) & ", " & cash_letter_col & ":" & cash_letter_col & ", " & is_not_blank_excel_string & "))/(COUNTIF(" & cash_letter_col & ":" & cash_letter_col & ", " & is_not_blank_excel_string & ") -1)" 'Excel formula
	ObjExcel.Cells(row_to_use + 1, col_to_use).NumberFormat = "0.00%"		'Formula should be percent
	row_to_use = row_to_use + 2	'It's two rows we jump, because the cash stat takes up two rows
End if

'HC info
If HC_check = checked then	
	ObjExcel.Cells(row_to_use, col_to_use - 1).Value = "HC cases pending over 45 days:"	'Row header
	objExcel.Cells(row_to_use, col_to_use - 1).Font.Bold = TRUE						'Row header should be bold
	ObjExcel.Cells(row_to_use, col_to_use).Value = "=COUNTIFS(E:E, " & Chr(34) & ">45" & Chr(34) & ", " & HC_letter_col & ":" & HC_letter_col & ", " & is_not_blank_excel_string & ")"	'Excel formula
	ObjExcel.Cells(row_to_use + 1, col_to_use - 1).Value = "Percentage of HC cases pending over 45 days:"	'Row header
	objExcel.Cells(row_to_use + 1, col_to_use - 1).Font.Bold = TRUE								'Row header should be bold
	ObjExcel.Cells(row_to_use + 1, col_to_use).Value = "=(COUNTIFS(E:E, " & Chr(34) & ">45" & Chr(34) & ", " & HC_letter_col & ":" & HC_letter_col & ", " & is_not_blank_excel_string & "))/(COUNTIF(" & HC_letter_col & ":" & HC_letter_col & ", " & is_not_blank_excel_string & ") -1)" 'Excel formula
	ObjExcel.Cells(row_to_use + 1, col_to_use).NumberFormat = "0.00%"		'Formula should be percent
	row_to_use = row_to_use + 2	'It's two rows we jump, because the HC stat takes up two rows
End if

'EA info
If EA_check = checked then	
	ObjExcel.Cells(row_to_use, col_to_use - 1).Value = "EA cases pending over 30 days:"	'Row header
	objExcel.Cells(row_to_use, col_to_use - 1).Font.Bold = TRUE						'Row header should be bold
	ObjExcel.Cells(row_to_use, col_to_use).Value = "=COUNTIFS(E:E, " & Chr(34) & ">30" & Chr(34) & ", " & EA_letter_col & ":" & EA_letter_col & ", " & is_not_blank_excel_string & ")"	'Excel formula
	ObjExcel.Cells(row_to_use + 1, col_to_use - 1).Value = "Percentage of EA cases pending over 30 days:"	'Row header
	objExcel.Cells(row_to_use + 1, col_to_use - 1).Font.Bold = TRUE								'Row header should be bold
	ObjExcel.Cells(row_to_use + 1, col_to_use).Value = "=(COUNTIFS(E:E, " & Chr(34) & ">30" & Chr(34) & ", " & EA_letter_col & ":" & EA_letter_col & ", " & is_not_blank_excel_string & "))/(COUNTIF(" & EA_letter_col & ":" & EA_letter_col & ", " & is_not_blank_excel_string & ") -1)" 'Excel formula
	ObjExcel.Cells(row_to_use + 1, col_to_use).NumberFormat = "0.00%"		'Formula should be percent
	row_to_use = row_to_use + 2	'It's two rows we jump, because the EA stat takes up two rows
End if

'GRH info
If GRH_check = checked then	
	ObjExcel.Cells(row_to_use, col_to_use - 1).Value = "GRH cases pending over 30 days:"	'Row header
	objExcel.Cells(row_to_use, col_to_use - 1).Font.Bold = TRUE						'Row header should be bold
	ObjExcel.Cells(row_to_use, col_to_use).Value = "=COUNTIFS(E:E, " & Chr(34) & ">30" & Chr(34) & ", " & GRH_letter_col & ":" & GRH_letter_col & ", " & is_not_blank_excel_string & ")"	'Excel formula
	ObjExcel.Cells(row_to_use + 1, col_to_use - 1).Value = "Percentage of GRH cases pending over 30 days:"	'Row header
	objExcel.Cells(row_to_use + 1, col_to_use - 1).Font.Bold = TRUE								'Row header should be bold
	ObjExcel.Cells(row_to_use + 1, col_to_use).Value = "=(COUNTIFS(E:E, " & Chr(34) & ">30" & Chr(34) & ", " & GRH_letter_col & ":" & GRH_letter_col & ", " & is_not_blank_excel_string & "))/(COUNTIF(" & GRH_letter_col & ":" & GRH_letter_col & ", " & is_not_blank_excel_string & ") -1)" 'Excel formula
	ObjExcel.Cells(row_to_use + 1, col_to_use).NumberFormat = "0.00%"		'Formula should be percent
	row_to_use = row_to_use + 2	'It's two rows we jump, because the GRH stat takes up two rows
End if

'Autofitting columns
For col_to_autofit = 1 to col_to_use
	ObjExcel.columns(col_to_autofit).AutoFit()
Next

'Provides additional statistics for SNAP cases
If SNAP_check = checked then

	'Going to another sheet, to enter worker-specific statistics
	ObjExcel.Worksheets.Add().Name = "SNAP stats by worker"
	
	'Headers
	ObjExcel.Cells(1, 2).Value = "SNAP STATS BY WORKER"
	ObjExcel.Cells(1, 2).Font.Bold = TRUE
	ObjExcel.Cells(2, 1).Value = "WORKER"
	objExcel.Cells(2, 1).Font.Bold = TRUE
	ObjExcel.Cells(2, 2).Value = "PENDING <= 30 DAYS"
	objExcel.Cells(2, 2).Font.Bold = TRUE
	ObjExcel.Cells(2, 3).Value = "TOTAL PENDING"
	objExcel.Cells(2, 3).Font.Bold = TRUE
	ObjExcel.Cells(2, 4).Value = "% PENDING <= 30 DAYS"
	objExcel.Cells(2, 4).Font.Bold = TRUE
	ObjExcel.Cells(2, 5).Value = "% OF SAMPLED WORKLOAD"
	objExcel.Cells(2, 5).Font.Bold = TRUE
	
	
	'Writes each worker from the worker_array in the Excel spreadsheet
	For x = 0 to ubound(worker_array)
		ObjExcel.Cells(x + 3, 1) = worker_array(x)
		ObjExcel.Cells(x + 3, 2) = "=COUNTIFS('Case information'!" & SNAP_letter_col & ":" & SNAP_letter_col & ", " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'Case information'!A:A, A" & x + 3 & ", 'Case information'!E:E, " & Chr(34) & "<=30" & Chr(34) & ")"
		ObjExcel.Cells(x + 3, 3) = "=COUNTIFS('Case information'!" & SNAP_letter_col & ":" & SNAP_letter_col & ", " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'Case information'!A:A, A" & x + 3 & ")"
		ObjExcel.Cells(x + 3, 4) = "=B" & x + 3 & "/C" & x + 3
		ObjExcel.Cells(x + 3, 4).NumberFormat = "0.00%"		'Formula should be percent
		ObjExcel.Cells(x + 3, 5) = "=C" & x + 3 & "/SUM(C:C)"
		ObjExcel.Cells(x + 3, 5).NumberFormat = "0.00%"		'Formula should be percent
	Next
	
	'Merging header cell.
	ObjExcel.Range(ObjExcel.Cells(1, 1), ObjExcel.Cells(1, 5)).Merge
	
	'Centering the cell
	objExcel.Cells(1, 2).HorizontalAlignment = -4108
	
	'Autofitting columns
	For col_to_autofit = 1 to 20
		ObjExcel.columns(col_to_autofit).AutoFit()
	Next
End if

'Provides additional statistics for cash cases
If cash_check = checked then

	'Going to another sheet, to enter worker-specific statistics
	ObjExcel.Worksheets.Add().Name = "cash stats by worker"
	
	'Headers
	ObjExcel.Cells(1, 2).Value = "CASH STATS BY WORKER"
	ObjExcel.Cells(1, 2).Font.Bold = TRUE
	ObjExcel.Cells(2, 1).Value = "WORKER"
	objExcel.Cells(2, 1).Font.Bold = TRUE
	ObjExcel.Cells(2, 2).Value = "PENDING <= 30 DAYS"
	objExcel.Cells(2, 2).Font.Bold = TRUE
	ObjExcel.Cells(2, 3).Value = "TOTAL PENDING"
	objExcel.Cells(2, 3).Font.Bold = TRUE
	ObjExcel.Cells(2, 4).Value = "% PENDING <= 30 DAYS"
	objExcel.Cells(2, 4).Font.Bold = TRUE
	ObjExcel.Cells(2, 5).Value = "% OF SAMPLED WORKLOAD"
	objExcel.Cells(2, 5).Font.Bold = TRUE
	
	
	'Writes each worker from the worker_array in the Excel spreadsheet
	For x = 0 to ubound(worker_array)
		ObjExcel.Cells(x + 3, 1) = worker_array(x)
		ObjExcel.Cells(x + 3, 2) = "=COUNTIFS('Case information'!" & cash_letter_col & ":" & cash_letter_col & ", " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'Case information'!A:A, A" & x + 3 & ", 'Case information'!E:E, " & Chr(34) & "<=30" & Chr(34) & ")"
		ObjExcel.Cells(x + 3, 3) = "=COUNTIFS('Case information'!" & cash_letter_col & ":" & cash_letter_col & ", " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'Case information'!A:A, A" & x + 3 & ")"
		ObjExcel.Cells(x + 3, 4) = "=B" & x + 3 & "/C" & x + 3
		ObjExcel.Cells(x + 3, 4).NumberFormat = "0.00%"		'Formula should be percent
		ObjExcel.Cells(x + 3, 5) = "=C" & x + 3 & "/SUM(C:C)"
		ObjExcel.Cells(x + 3, 5).NumberFormat = "0.00%"		'Formula should be percent
	Next
	
	'Merging header cell.
	ObjExcel.Range(ObjExcel.Cells(1, 1), ObjExcel.Cells(1, 5)).Merge
	
	'Centering the cell
	objExcel.Cells(1, 2).HorizontalAlignment = -4108
	
	'Autofitting columns
	For col_to_autofit = 1 to 20
		ObjExcel.columns(col_to_autofit).AutoFit()
	Next
End if

'Provides additional statistics for HC cases
If HC_check = checked then

	'Going to another sheet, to enter worker-specific statistics
	ObjExcel.Worksheets.Add().Name = "HC stats by worker"
	
	'Headers
	ObjExcel.Cells(1, 2).Value = "HC STATS BY WORKER"
	ObjExcel.Cells(1, 2).Font.Bold = TRUE
	ObjExcel.Cells(2, 1).Value = "WORKER"
	objExcel.Cells(2, 1).Font.Bold = TRUE
	ObjExcel.Cells(2, 2).Value = "PENDING <= 45 DAYS"
	objExcel.Cells(2, 2).Font.Bold = TRUE
	ObjExcel.Cells(2, 3).Value = "TOTAL PENDING"
	objExcel.Cells(2, 3).Font.Bold = TRUE
	ObjExcel.Cells(2, 4).Value = "% PENDING <= 45 DAYS"
	objExcel.Cells(2, 4).Font.Bold = TRUE
	ObjExcel.Cells(2, 5).Value = "% OF SAMPLED WORKLOAD"
	objExcel.Cells(2, 5).Font.Bold = TRUE
	
	
	'Writes each worker from the worker_array in the Excel spreadsheet
	For x = 0 to ubound(worker_array)
		ObjExcel.Cells(x + 3, 1) = worker_array(x)
		ObjExcel.Cells(x + 3, 2) = "=COUNTIFS('Case information'!" & HC_letter_col & ":" & HC_letter_col & ", " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'Case information'!A:A, A" & x + 3 & ", 'Case information'!E:E, " & Chr(34) & "<=45" & Chr(34) & ")"
		ObjExcel.Cells(x + 3, 3) = "=COUNTIFS('Case information'!" & HC_letter_col & ":" & HC_letter_col & ", " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'Case information'!A:A, A" & x + 3 & ")"
		ObjExcel.Cells(x + 3, 4) = "=B" & x + 3 & "/C" & x + 3
		ObjExcel.Cells(x + 3, 4).NumberFormat = "0.00%"		'Formula should be percent
		ObjExcel.Cells(x + 3, 5) = "=C" & x + 3 & "/SUM(C:C)"
		ObjExcel.Cells(x + 3, 5).NumberFormat = "0.00%"		'Formula should be percent
	Next
	
	'Merging header cell.
	ObjExcel.Range(ObjExcel.Cells(1, 1), ObjExcel.Cells(1, 5)).Merge
	
	'Centering the cell
	objExcel.Cells(1, 2).HorizontalAlignment = -4108
	
	'Autofitting columns
	For col_to_autofit = 1 to 20
		ObjExcel.columns(col_to_autofit).AutoFit()
	Next
End if

'Provides additional statistics for EA cases
If EA_check = checked then

	'Going to another sheet, to enter worker-specific statistics
	ObjExcel.Worksheets.Add().Name = "EA stats by worker"
	
	'Headers
	ObjExcel.Cells(1, 2).Value = "EA STATS BY WORKER"
	ObjExcel.Cells(1, 2).Font.Bold = TRUE
	ObjExcel.Cells(2, 1).Value = "WORKER"
	objExcel.Cells(2, 1).Font.Bold = TRUE
	ObjExcel.Cells(2, 2).Value = "PENDING <= 30 DAYS"
	objExcel.Cells(2, 2).Font.Bold = TRUE
	ObjExcel.Cells(2, 3).Value = "TOTAL PENDING"
	objExcel.Cells(2, 3).Font.Bold = TRUE
	ObjExcel.Cells(2, 4).Value = "% PENDING <= 30 DAYS"
	objExcel.Cells(2, 4).Font.Bold = TRUE
	ObjExcel.Cells(2, 5).Value = "% OF SAMPLED WORKLOAD"
	objExcel.Cells(2, 5).Font.Bold = TRUE
	
	
	'Writes each worker from the worker_array in the Excel spreadsheet
	For x = 0 to ubound(worker_array)
		ObjExcel.Cells(x + 3, 1) = worker_array(x)
		ObjExcel.Cells(x + 3, 2) = "=COUNTIFS('Case information'!" & EA_letter_col & ":" & EA_letter_col & ", " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'Case information'!A:A, A" & x + 3 & ", 'Case information'!E:E, " & Chr(34) & "<=30" & Chr(34) & ")"
		ObjExcel.Cells(x + 3, 3) = "=COUNTIFS('Case information'!" & EA_letter_col & ":" & EA_letter_col & ", " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'Case information'!A:A, A" & x + 3 & ")"
		ObjExcel.Cells(x + 3, 4) = "=B" & x + 3 & "/C" & x + 3
		ObjExcel.Cells(x + 3, 4).NumberFormat = "0.00%"		'Formula should be percent
		ObjExcel.Cells(x + 3, 5) = "=C" & x + 3 & "/SUM(C:C)"
		ObjExcel.Cells(x + 3, 5).NumberFormat = "0.00%"		'Formula should be percent
	Next
	
	'Merging header cell.
	ObjExcel.Range(ObjExcel.Cells(1, 1), ObjExcel.Cells(1, 5)).Merge
	
	'Centering the cell
	objExcel.Cells(1, 2).HorizontalAlignment = -4108
	
	'Autofitting columns
	For col_to_autofit = 1 to 20
		ObjExcel.columns(col_to_autofit).AutoFit()
	Next
End if

'Provides additional statistics for GRH cases
If GRH_check = checked then

	'Going to another sheet, to enter worker-specific statistics
	ObjExcel.Worksheets.Add().Name = "GRH stats by worker"
	
	'Headers
	ObjExcel.Cells(1, 2).Value = "GRH STATS BY WORKER"
	ObjExcel.Cells(1, 2).Font.Bold = TRUE
	ObjExcel.Cells(2, 1).Value = "WORKER"
	objExcel.Cells(2, 1).Font.Bold = TRUE
	ObjExcel.Cells(2, 2).Value = "PENDING <= 30 DAYS"
	objExcel.Cells(2, 2).Font.Bold = TRUE
	ObjExcel.Cells(2, 3).Value = "TOTAL PENDING"
	objExcel.Cells(2, 3).Font.Bold = TRUE
	ObjExcel.Cells(2, 4).Value = "% PENDING <= 30 DAYS"
	objExcel.Cells(2, 4).Font.Bold = TRUE
	ObjExcel.Cells(2, 5).Value = "% OF SAMPLED WORKLOAD"
	objExcel.Cells(2, 5).Font.Bold = TRUE
	
	
	'Writes each worker from the worker_array in the Excel spreadsheet
	For x = 0 to ubound(worker_array)
		ObjExcel.Cells(x + 3, 1) = worker_array(x)
		ObjExcel.Cells(x + 3, 2) = "=COUNTIFS('Case information'!" & GRH_letter_col & ":" & GRH_letter_col & ", " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'Case information'!A:A, A" & x + 3 & ", 'Case information'!E:E, " & Chr(34) & "<=30" & Chr(34) & ")"
		ObjExcel.Cells(x + 3, 3) = "=COUNTIFS('Case information'!" & GRH_letter_col & ":" & GRH_letter_col & ", " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'Case information'!A:A, A" & x + 3 & ")"
		ObjExcel.Cells(x + 3, 4) = "=B" & x + 3 & "/C" & x + 3
		ObjExcel.Cells(x + 3, 4).NumberFormat = "0.00%"		'Formula should be percent
		ObjExcel.Cells(x + 3, 5) = "=C" & x + 3 & "/SUM(C:C)"
		ObjExcel.Cells(x + 3, 5).NumberFormat = "0.00%"		'Formula should be percent
	Next
	
	'Merging header cell.
	ObjExcel.Range(ObjExcel.Cells(1, 1), ObjExcel.Cells(1, 5)).Merge
	
	'Centering the cell
	objExcel.Cells(1, 2).HorizontalAlignment = -4108
	
	'Autofitting columns
	For col_to_autofit = 1 to 20
		ObjExcel.columns(col_to_autofit).AutoFit()
	Next
End if

'Logging usage stats
script_end_procedure("")
