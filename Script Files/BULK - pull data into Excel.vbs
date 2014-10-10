'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - pull data into Excel"
start_time = timer

'LOADING ROUTINE FUNCTIONS----------------------------------------------------------------------------------------------------
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\MAXIS-BZ-Scripts-County-Beta\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'CUSTOM FUNCTION (WILL INCLUDE IN FUNCTIONS FILE BEFORE FULL DEPLOYMENT)

'This function converts a numeric digit to an Excel column, up to 104 digits (columns).
function convert_digit_to_excel_column(col_in_excel)
	'Create string with the alphabet
	alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"

	'Assigning a letter, based on that column. Uses "mid" function to determine it. If number > 26, it handles by adding a letter (per Excel).
	convert_digit_to_excel_column = Mid(alphabet, col_in_excel, 1)		
	If col_in_excel >= 27 and col_in_excel < 53 then convert_digit_to_excel_column = "A" & Mid(alphabet, col_in_excel - 26, 1)
	If col_in_excel >= 53 and col_in_excel < 79 then convert_digit_to_excel_column = "B" & Mid(alphabet, col_in_excel - 52, 1)
	If col_in_excel >= 79 and col_in_excel < 105 then convert_digit_to_excel_column = "C" & Mid(alphabet, col_in_excel - 78, 1)

	'Closes script if the number gets too high (very rare circumstance, just errorproofing)
	If col_in_excel >= 105 then script_end_procedure("This script is only able to assign excel columns to 104 rows. You've exceeded this number, and this script cannot continue.")
end function

'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog pull_PND2_data_into_excel_dialog, 0, 0, 286, 115, "Pull PND2 data into Excel dialog"
  EditBox 150, 20, 130, 15, worker_number
  CheckBox 70, 55, 150, 10, "Check here to run this query county-wide.", all_workers_check
  CheckBox 10, 35, 40, 10, "SNAP?", SNAP_check
  CheckBox 10, 50, 40, 10, "Cash?", cash_check
  CheckBox 10, 65, 40, 10, "HC?", HC_check
  CheckBox 10, 80, 40, 10, "EA?", EA_check
  CheckBox 10, 95, 40, 10, "GRH?", GRH_check
  ButtonGroup ButtonPressed
    OkButton 175, 95, 50, 15
    CancelButton 230, 95, 50, 15
  GroupBox 5, 20, 60, 90, "Progs to scan"
  Text 70, 25, 65, 10, "Worker(s) to check:"
  Text 70, 70, 215, 20, "NOTE: running queries county-wide can take a significant amount of time and resources. This should be done after hours."
  Text 65, 5, 125, 10, "***PULL PND2 DATA INTO EXCEL***"
  Text 70, 40, 215, 10, "Enter workers' x1 numbers (ex: x100###), separated by a comma."
EndDialog

'VARIABLES TO DECLARE
all_case_numbers_array = " "					'Creating blank variable for the future array
screen_to_check = "PND2"					'Declaring this here, but in a future revision this will be a dropdown amongst other options

call worker_county_code_determination(worker_county_code, two_digit_county_code)

'THE SCRIPT----------------------------------------------------------------------------------------------------

'Connecting to BlueZone
EMConnect ""

If screen_to_check = "PND2" then

	'Dialog asks what stats are being pulled
	Dialog pull_PND2_data_into_excel_dialog
	If buttonpressed = 0 then stopscript

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
		x1s_from_dialog = split(worker_number, ",")	'Splits the worker array based on commas

		'Need to add the worker_county_code to each one
		For each x1_number in x1s_from_dialog
			worker_array = worker_array & ", " & worker_county_code & trim(x1_number)
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

	'Setting variables for the stats area
	col_to_use = col_to_use + 2	'Doing two because the wrap-up is two columns
	row_to_use = 3			'Setting variable for the if...thens below
	is_not_blank_excel_string = Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34)	'Pre-declaring this, to clean up the code below

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
		ObjExcel.Cells(row_to_use, col_to_use - 1).Value = "HC cases pending over 30 days:"	'Row header
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

	'Logging usage stats
	script_end_procedure("")

End if


