'PLEASE NOTE: this script was designed to run off of the BULK - pull data into Excel script.
'As such, it might not work if ran separately from that.

'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - REPT-MFCM LIST.vbs"
start_time = timer

'DIALOGS----------------------------------------------------------------------
BeginDialog pull_REPT_data_into_excel_dialog, 0, 0, 218, 120, "Pull REPT data into Excel dialog"
  EditBox 84, 20, 130, 15, worker_number
  CheckBox 4, 65, 150, 10, "Check here to run this query county-wide.", all_workers_check
  ButtonGroup ButtonPressed
    OkButton 109, 100, 50, 15
    CancelButton 164, 100, 50, 15
  Text 4, 25, 65, 10, "Worker(s) to check:"
  Text 4, 80, 210, 20, "NOTE: running queries county-wide can take a significant amount of time and resources. This should be done after hours."
  Text 14, 5, 125, 10, "***PULL REPT DATA INTO EXCEL***"
  Text 4, 40, 210, 20, "Enter last 3 digits of your workers' x1 numbers (ex: x100###), separated by a comma."
EndDialog

'THE SCRIPT-------------------------------------------------------------------------

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

'Setting the first 4 col as worker, case number, name, and APPL date
ObjExcel.Cells(1, 1).Value = "WORKER"
objExcel.Cells(1, 1).Font.Bold = TRUE
ObjExcel.Cells(1, 2).Value = "CASE NUMBER"
objExcel.Cells(1, 2).Font.Bold = TRUE
ObjExcel.Cells(1, 3).Value = "NAME"
objExcel.Cells(1, 3).Font.Bold = TRUE
ObjExcel.Cells(1, 4).Value = "SANCTION %"
objExcel.Cells(1, 4).Font.Bold = TRUE
ObjExcel.Cells(1, 5).Value = "VEND RSN"
objExcel.Cells(1, 5).Font.Bold = TRUE
ObjExcel.Cells(1, 6).Value = "EMPS STATUS"
objExcel.Cells(1, 6).Font.Bold = TRUE
ObjExcel.Cells(1, 7).Value = "HRS RETRO"
objExcel.Cells(1, 7).Font.Bold = TRUE
ObjExcel.Cells(1, 8).Value = "EMPL PRO"
objExcel.Cells(1, 8).Font.Bold = TRUE
ObjExcel.Cells(1, 9).Value = "TANF MOS"
objExcel.Cells(1, 9).Font.Bold = TRUE
ObjExcel.Cells(1, 10).Value = "60 MOS EXT RSN"
objExcel.Cells(1, 10).Font.Bold = TRUE

'Figuring out what to put in each Excel col. To add future variables to this, add the checkbox variables below and copy/paste the same code!
'	Below, use the "[blank]_col" variable to recall which col you set for which option.
col_to_use = 10 'Starting with 5 because cols 1-4 are already used

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
	Call navigate_to_screen("rept", "mfcm")
	EMWriteScreen worker, 21, 13
	transmit

	'Skips workers with no info
	EMReadScreen has_content_check, 29, 7, 6
  has_content_check = trim(has_content_check)
	If has_content_check <> "" then

		'Grabbing each case number on screen
		Do
			'Set variable for next do...loop
			MAXIS_row = 7
			
			'Checking for the last page of cases.
			EMReadScreen last_page_check, 21, 24, 2	'because on REPT/MFCF it displays right away, instead of when the second F8 is sent
			Do			
				EMReadScreen case_number, 8, MAXIS_row, 6		  'Reading case number
				EMReadScreen client_name, 20, MAXIS_row, 16		'Reading client name
				EMReadScreen sanc_perc, 2, MAXIS_row, 39	    'Reading Sanction Percentage
				EMReadScreen vend_rsn, 2, MAXIS_row, 45		    'Reading Vend Rsn
				EMReadScreen emps_status, 2, MAXIS_row, 52		'Reading Emps Status
				EMReadScreen hrs_retro, 3, MAXIS_row, 57			'Reading Hrs Retro
				EMReadScreen empl_pro, 3, MAXIS_row, 62			  'Reading Empl Pro
				EMReadScreen tanf_mos, 2, MAXIS_row, 69			  'Reading TANF Mos
				EMReadScreen sixty_ext_rsn, 2, MAXIS_row, 75	'Reading 60 Mos Ext Rsn

				'Doing this because sometimes BlueZone registers a "ghost" of previous data when the script runs. This checks against an array and stops if we've seen this one before.
				If trim(case_number) <> "" and instr(all_case_numbers_array, case_number) <> 0 then exit do
				all_case_numbers_array = trim(all_case_numbers_array & " " & case_number)

				If case_number = "        " and client_name = "                    " then exit do			'Exits do if we reach the end

        ObjExcel.Cells(excel_row, 1).Value = worker
        ObjExcel.Cells(excel_row, 2).Value = case_number
        ObjExcel.Cells(excel_row, 3).Value = client_name
        ObjExcel.Cells(excel_row, 4).Value = sanc_perc
        ObjExcel.Cells(excel_row, 5).Value = vend_rsn
        ObjExcel.Cells(excel_row, 6).Value = emps_status
        ObjExcel.Cells(excel_row, 7).Value = hrs_retro
        ObjExcel.Cells(excel_row, 8).Value = empl_pro
        ObjExcel.Cells(excel_row, 9).Value = tanf_mos
        ObjExcel.Cells(excel_row, 10).Value = sixty_ext_rsn
        
        excel_row = excel_row + 1
        
				MAXIS_row = MAXIS_row + 1
				case_number = ""			'Blanking out variable
			Loop until MAXIS_row = 19
			PF8
			
		Loop until last_page_check = "THIS IS THE LAST PAGE"
	End if
next

col_to_use = col_to_use + 2	'Doing two because the wrap-up is two columns

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

'Logging usage stats
script_end_procedure("")
