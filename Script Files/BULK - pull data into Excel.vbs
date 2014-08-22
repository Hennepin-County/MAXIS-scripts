'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - pull data into Excel"
start_time = timer

'LOADING ROUTINE FUNCTIONS----------------------------------------------------------------------------------------------------
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\MAXIS-BZ-Scripts-County-Beta\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script



'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog pull_PND2_data_into_excel_dialog, 0, 0, 281, 100, "Pull PND2 data into Excel dialog"
  EditBox 185, 20, 90, 15, worker_number
  CheckBox 120, 40, 150, 10, "Check here to run this query county-wide.", all_workers_check
  CheckBox 10, 35, 100, 10, "Days pending", days_pending_check
  CheckBox 10, 70, 105, 10, "SNAP cases only", SNAP_cases_only_check
  ButtonGroup ButtonPressed
    OkButton 140, 80, 50, 15
    CancelButton 195, 80, 50, 15
  GroupBox 5, 20, 110, 30, "Additional items to log"
  GroupBox 5, 55, 110, 30, "Filters"
  Text 120, 25, 60, 10, "Worker to check:"
  Text 120, 50, 155, 25, "NOTE: running queries county-wide can take a significant amount of time and resources. This should be done after hours."
  Text 40, 5, 220, 10, "***AT THIS TIME, THIS SCRIPT PULLS PND2 DATA ONLY***"
EndDialog



'VARIABLES TO DECLARE
all_case_numbers_array = " "					'Creating blank variable for the future array
screen_to_check = "PND2"					'Declaring this here, but in a future revision this will be a dropdown amongst other options

'THE SCRIPT----------------------------------------------------------------------------------------------------

'Connecting to BlueZone
EMConnect ""

If screen_to_check = "PND2" then

	'Dialog asks what stats are being pulled
	Dialog pull_PND2_data_into_excel_dialog
	If buttonpressed = 0 then stopscript

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
	ObjExcel.Cells(1, 4).Value = "APPL DATE"
	objExcel.Cells(1, 4).Font.Bold = TRUE

	'Figuring out what to put in each Excel col. To add future variables to this, add the checkbox variables below and copy/paste the same code!
	'	Below, use the "[blank]_col" variable to recall which col you set for which option.
	col_to_use = 5 'Starting with 5 because cols 1-4 are already used
	If days_pending_check = checked then	'NOTE: keep this first, or program a function to convert the numeric days_pending_col to an alpha for the totals later on
		ObjExcel.Cells(1, col_to_use).Value = "DAYS PENDING"	
		objExcel.Cells(1, col_to_use).Font.Bold = TRUE
		days_pending_col = col_to_use
		col_to_use = col_to_use + 1
	End if
	If preg_check = checked then
		ObjExcel.Cells(1, col_to_use).Value = "PREG EXISTS?"
		objExcel.Cells(1, col_to_use).Font.Bold = TRUE
		preg_col = col_to_use
		col_to_use = col_to_use + 1
	End if
	If all_HH_membs_19_plus_check = checked then
		ObjExcel.Cells(1, col_to_use).Value = "ALL MEMBS 19+?"
		objExcel.Cells(1, col_to_use).Font.Bold = TRUE
		all_HH_membs_19_plus_col = col_to_use
		col_to_use = col_to_use + 1
	End if
	If ages_of_HH_membs_check = checked then
		ObjExcel.Cells(1, col_to_use).Value = "AGES OF HH MEMBS"
		objExcel.Cells(1, col_to_use).Font.Bold = TRUE
		ages_of_HH_membs_col = col_to_use
		col_to_use = col_to_use + 1
	End if
	If number_of_HH_membs_check = checked then
		ObjExcel.Cells(1, col_to_use).Value = "NUMBER OF HH MEMBS?"
		objExcel.Cells(1, col_to_use).Font.Bold = TRUE
		number_of_HH_membs_col = col_to_use
		col_to_use = col_to_use + 1
	End if
	If ABAWD_code_check = checked then
		ObjExcel.Cells(1, col_to_use).Value = "ABAWD CODE"
		objExcel.Cells(1, col_to_use).Font.Bold = TRUE
		ABAWD_code_col = col_to_use
		col_to_use = col_to_use + 1
	End if

	'Setting the variable for what's to come
	excel_row = 2

	'If all workers are selected, the script will go to REPT/USER, and load all of the workers into an array. Otherwise it'll create a single-object "array" just for simplicity of code.
	If all_workers_check = checked then
		call create_array_of_all_active_x_numbers_in_county(worker_array, two_digit_county_code)
	Else
		worker_array = split(worker_county_code & worker_number)
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
					EMReadScreen SNAP_status, 1, MAXIS_row, 62
					EMReadScreen case_number, 8, MAXIS_row, 5
					EMReadScreen client_name, 22, MAXIS_row, 16
					EMReadScreen APPL_date, 8, MAXIS_row, 38
					EMReadScreen days_pending, 4, MAXIS_row, 49
					'Doing this because sometimes BlueZone registers a "ghost" of previous data when the script runs. This checks against an array and stops if we've seen this one before.
					If trim(case_number) <> "" and (instr(all_case_numbers_array, case_number) <> 0 and client_name <> " ADDITIONAL APP       ") then exit do
					all_case_numbers_array = trim(all_case_numbers_array & " " & case_number)
					If case_number = "        " then exit do
					SNAP_status = trim(replace(SNAP_status, "_", ""))
					If (SNAP_cases_only_check = checked and SNAP_status <> "") or SNAP_cases_only_check = 0 then 
						ObjExcel.Cells(excel_row, 1).Value = worker
						ObjExcel.Cells(excel_row, 2).Value = case_number
						ObjExcel.Cells(excel_row, 3).Value = client_name
						ObjExcel.Cells(excel_row, 4).Value = replace(APPL_date, " ", "/")
						If days_pending_check = checked then ObjExcel.Cells(excel_row, days_pending_col).Value = cint(days_pending)	'Only grabs this if requested
						excel_row = excel_row + 1

					End if
					MAXIS_row = MAXIS_row + 1
				Loop until MAXIS_row = 19
				PF8
				EMReadScreen last_page_check, 21, 24, 2
			Loop until last_page_check = "THIS IS THE LAST PAGE"
		End if
	next

	'Resetting excel_row variable, now we need to start looking people up
	excel_row = 2 

	Do 
		case_number = ObjExcel.Cells(excel_row, 2).Value
		If case_number = "" then exit do

		'Now pulling PREG info
		If preg_check = checked then
			call navigate_to_screen("STAT", "PREG")
			EMReadScreen ERRR_check, 4, 2, 52
			If ERRR_check = "ERRR" then transmit	'for error prone cases
			EMReadScreen PREG_panel_check, 1, 2, 78
			If PREG_panel_check <> "0" then 
				ObjExcel.Cells(excel_row, preg_col).Value = "Y"
			Else
				ObjExcel.Cells(excel_row, preg_col).Value = "N"
			End if
		End if

		'Now pulling 19+ y/o info
		If all_HH_membs_19_plus_check = checked then
			call navigate_to_screen("STAT", "MEMB")
			EMReadScreen ERRR_check, 4, 2, 52
			If ERRR_check = "ERRR" then transmit	'for error prone cases
			Do
				EMReadScreen MEMB_panel_current, 1, 2, 73
				EMReadScreen MEMB_panel_total, 1, 2, 78
				EMReadScreen MEMB_age, 3, 8, 76
				If MEMB_age = "   " then MEMB_age = "0"
				If cint(MEMB_age) < 19 then has_minor_in_case = True
				transmit
			Loop until MEMB_panel_current = MEMB_panel_total
			If has_minor_in_case <> True then 
				ObjExcel.Cells(excel_row, all_HH_membs_19_plus_col).Value = "Y"
			Else
				ObjExcel.Cells(excel_row, all_HH_membs_19_plus_col).Value = "N"
			End if
			has_minor_in_case = "" 'clearing variable
		End if

		'Now pulling ages of MEMBs info
		If ages_of_HH_membs_check = checked then
			call navigate_to_screen("STAT", "MEMB")
			EMReadScreen ERRR_check, 4, 2, 52
			If ERRR_check = "ERRR" then transmit	'for error prone cases
			Do
				EMReadScreen MEMB_panel_current, 1, 2, 73
				EMReadScreen MEMB_panel_total, 1, 2, 78
				EMReadScreen MEMB_age, 3, 8, 76
				EMReadScreen MEMB_number, 2, 4, 33
				If MEMB_age = "   " then MEMB_age = "0"
				MEMB_age_array = trim(MEMB_age_array & " MEMB " & MEMB_number & ": " & trim(MEMB_age) & " y/o.")
				transmit
			Loop until MEMB_panel_current = MEMB_panel_total
			ObjExcel.Cells(excel_row, ages_of_HH_membs_col).Value = MEMB_age_array
			MEMB_age_array = "" 'clearing variable
		End if

		'Now pulling number of membs info
		If number_of_HH_membs_check = checked then
			call navigate_to_screen("STAT", "MEMB")
			EMReadScreen ERRR_check, 4, 2, 52
			If ERRR_check = "ERRR" then transmit	'for error prone cases
			EMReadScreen MEMB_panel_total, 1, 2, 78
			ObjExcel.Cells(excel_row, number_of_HH_membs_col).Value = cint(MEMB_panel_total)
		End if

		'Now pulling ABAWD info
		If ABAWD_code_check = checked then
			call navigate_to_screen("STAT", "WREG")
			EMReadScreen ERRR_check, 4, 2, 52		'Error prone case checking
			If ERRR_check = "ERRR" then transmit	'transmitting if case is error prone
			EMReadScreen WREG_panel_total, 1, 2, 78
			If WREG_panel_total <> "0" then
				WREG_row = 5 'setting variable for do...loop
				WREG_membs_array = "" 'Clearing variable to use in the do...loop
				Do
					EMReadScreen WREG_ref_nbr, 2, WREG_row, 3
					If WREG_ref_nbr = "  " then exit do
					WREG_membs_array = WREG_membs_array & WREG_ref_nbr & ", "
					WREG_row = WREG_row + 1
				Loop until WREG_row = 19
				WREG_membs_array = split(WREG_membs_array, ", ")
				For each WREG_memb in WREG_membs_array
					EMWriteScreen WREG_memb, 20, 76
					transmit
					EMReadScreen ABAWD_status_code, 2, 13, 50
					If WREG_memb <> "" then ABAWD_status = ABAWD_status & WREG_memb & ": " & ABAWD_status_code & ", "
				Next
				ObjExcel.Cells(excel_row, ABAWD_code_col).Value = "'" & left(ABAWD_status, len(ABAWD_status) - 2)
				ABAWD_status = "" 'clearing variable
			End if
		End if

		excel_row = excel_row + 1
	Loop until case_number = ""

	'Wrap-up stats if case pending days was requested
	col_to_use = col_to_use + 2	'Doing two because the wrap-up is two columns

	objExcel.Cells(1, col_to_use - 1).Font.Bold = TRUE
	objExcel.Cells(2, col_to_use - 1).Font.Bold = TRUE
	ObjExcel.Cells(1, col_to_use - 1).Value = "Query date and time:"	'Goes back one, as this is on the next row
	ObjExcel.Cells(1, col_to_use).Value = now
	ObjExcel.Cells(2, col_to_use - 1).Value = "Query runtime (in seconds):"	'Goes back one, as this is on the next row
	ObjExcel.Cells(2, col_to_use).Value = timer - start_time
	If days_pending_check = checked then	'The following is just for days pending.
		ObjExcel.Cells(3, col_to_use - 1).Value = "Percentage of cases over 30 days:"
		ObjExcel.Cells(3, col_to_use).Value = "=(COUNTIFS(E:E," & Chr(34) & ">30" & Chr(34) & "))/COUNT(E:E)"
		ObjExcel.Cells(3, col_to_use).NumberFormat = "0.00%"
		ObjExcel.Cells(4, col_to_use - 1).Value = "Cases pending over 30 days:"	'Goes back one, as this is on the next row
		ObjExcel.Cells(4, col_to_use).Value = "=COUNTIFS(E:E, " & Chr(34) & ">30" & Chr(34) & ")"
		objExcel.Cells(3, col_to_use - 1).Font.Bold = TRUE
		objExcel.Cells(4, col_to_use - 1).Font.Bold = TRUE
	End if

	'Autofitting columns

	For col_to_autofit = 1 to col_to_use
		ObjExcel.columns(col_to_autofit).AutoFit()
	Next

	'Logging usage stats
	script_end_procedure("")

End if


