'Required for statistical purposes===============================================================================
name_of_script = "BULK - REPT-PND2 LIST.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 13                      'manual run time in seconds
STATS_denomination = "C"       							'C is for each CASE
'END OF stats block==============================================================================================

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else											'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		End if
		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
		req.send													'Sends request
		IF req.Status = 200 THEN									'200 means great success
			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
			Execute req.responseText								'Executes the script code
		ELSE														'Error message
			critical_error_msgbox = MsgBox ("Something has gone wrong. The Functions Library code stored on GitHub was not able to be reached." & vbNewLine & vbNewLine &_
                                            "FuncLib URL: " & FuncLib_URL & vbNewLine & vbNewLine &_
                                            "The script has stopped. Please check your Internet connection. Consult a scripts administrator with any questions.", _
                                            vbOKonly + vbCritical, "BlueZone Scripts Critical Error")
            StopScript
		END IF
	ELSE
		FuncLib_URL = "C:\BZS-FuncLib\MASTER FUNCTIONS LIBRARY.vbs"
		Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
		Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
		text_from_the_other_script = fso_command.ReadAll
		fso_command.Close
		Execute text_from_the_other_script
	END IF
END IF
'END FUNCTIONS LIBRARY BLOCK================================================================================================

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
Call changelog_update("12/10/2016", "Added IV-E and Child Care cases statuses to script, and checkbox option to add information about the last case note (date, x number, header of case note) to the spreadsheet. Added navigation so that the 'Case information' worksheet is the first worksheet that is visable to the user. Also added closing message informing user that script has ended sucessfully.", "Ilse Ferris, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'DIALOG-----------------------------------------------------------------------------------------
BeginDialog pull_REPT_data_into_excel_dialog, 0, 0, 286, 135, "Pull REPT data into Excel dialog"
  EditBox 135, 20, 145, 15, worker_number
  CheckBox 70, 65, 150, 10, "Check here to run this query county-wide.", all_workers_check
  CheckBox 70, 80, 205, 10, "Check here to add last case note information to spreadsheet.", case_note_check
  CheckBox 10, 20, 40, 10, "SNAP?", SNAP_check
  CheckBox 10, 35, 40, 10, "Cash?", cash_check
  CheckBox 10, 50, 40, 10, "HC?", HC_check
  CheckBox 10, 65, 40, 10, "EA?", EA_check
  CheckBox 10, 80, 40, 10, "GRH?", GRH_check
  CheckBox 10, 95, 40, 10, "IV-E?", IVE_check
  CheckBox 10, 110, 50, 10, "Child care?", CC_check
  ButtonGroup ButtonPressed
    OkButton 175, 115, 50, 15
    CancelButton 230, 115, 50, 15
  Text 80, 5, 125, 10, "***PULL REPT DATA INTO EXCEL***"
  Text 70, 40, 210, 20, "Enter all 7 digits of your workers' x1 numbers (ex: x######), separated by a comma."
  GroupBox 5, 5, 60, 120, "Progs to scan"
  Text 70, 25, 65, 10, "Worker(s) to check:"
  Text 70, 95, 210, 20, "NOTE: running queries county-wide can take a significant amount of time and resources. This should be done after hours."
EndDialog

'THE SCRIPT-------------------------------------------------------------------------
'Gathering county code for multi-county...
get_county_code

'Connects to BlueZone
EMConnect ""

'Dialog asks what stats are being pulled
Do 	
	Dialog pull_REPT_data_into_excel_dialog
	If buttonpressed = cancel then stopscript
	Call check_for_password(are_we_passworded_out)
Loop until check_for_password(are_we_passworded_out) = False		'loops until user is password-ed out

'Starting the query start time (for the query runtime at the end)
query_start_time = timer

'Checking for MAXIS
Call check_for_MAXIS(True)

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
If IVE_check = checked then
	ObjExcel.Cells(1, col_to_use).Value = "IV-E?"
	objExcel.Cells(1, col_to_use).Font.Bold = TRUE
	ive_pends_col = col_to_use
	col_to_use = col_to_use + 1
	IVE_letter_col = convert_digit_to_excel_column(ive_pends_col)
End if
If CC_check = checked then
	ObjExcel.Cells(1, col_to_use).Value = "CC?"
	objExcel.Cells(1, col_to_use).Font.Bold = TRUE
	CC_pends_col = col_to_use
	col_to_use = col_to_use + 1
	CC_letter_col = convert_digit_to_excel_column(CC_pends_col)
End if

'Setting the variable for what's to come
excel_row = 2

'If all workers are selected, the script will go to REPT/USER, and load all of the workers into an array. Otherwise it'll create a single-object "array" just for simplicity of code.
If all_workers_check = checked then
	call create_array_of_all_active_x_numbers_in_county(worker_array, right(worker_county_code, 2))
Else
	x1s_from_dialog = split(worker_number, ",")	'Splits the worker array based on commas

	'Need to add the worker_county_code to each one
	For each x1_number in x1s_from_dialog
		If worker_array = "" then
			worker_array = trim(ucase(x1_number))		'replaces worker_county_code if found in the typed x1 number
		Else
			worker_array = worker_array & ", " & trim(ucase(x1_number)) 'replaces worker_county_code if found in the typed x1 number
		End if
	Next

	'Split worker_array
	worker_array = split(worker_array, ", ")
End if

For each worker in worker_array
	back_to_self	'Does this to prevent "ghosting" where the old info shows up on the new screen for some reason
	Call navigate_to_MAXIS_screen("rept", "pnd2")
	EMWriteScreen worker, 21, 13
	transmit

	'Skips workers with no info
	EMReadScreen has_content_check, 6, 3, 74
	If has_content_check <> "0 Of 0" then

		'Grabbing each case number on screen
		Do
			MAXIS_row = 7
			Do
				EMReadScreen MAXIS_case_number, 8, MAXIS_row, 5	'Reading case number
				EMReadScreen client_name, 22, MAXIS_row, 16		'Reading client name
				EMReadScreen APPL_date, 8, MAXIS_row, 38		'Reading application date
				EMReadScreen days_pending, 4, MAXIS_row, 49		'Reading days pending
				EMReadScreen cash_status, 1, MAXIS_row, 54		'Reading cash status
				EMReadScreen SNAP_status, 1, MAXIS_row, 62		'Reading SNAP status
				EMReadScreen HC_status, 1, MAXIS_row, 65		'Reading HC status
				EMReadScreen EA_status, 1, MAXIS_row, 68		'Reading EA status
				EMReadScreen GRH_status, 1, MAXIS_row, 72		'Reading GRH status
				EMReadScreen IVE_status, 1, MAXIS_row, 76		'Reading IV-E status
				EMReadScreen CC_status, 1, MAXIS_row, 80		'Reading CC status	

				'Doing this because sometimes BlueZone registers a "ghost" of previous data when the script runs. This checks against an array and stops if we've seen this one before.
				If trim(MAXIS_case_number) <> "" and (instr(all_case_numbers_array, MAXIS_case_number) <> 0 and client_name <> " ADDITIONAL APP       ") then exit do
				all_case_numbers_array = trim(all_case_numbers_array & " " & MAXIS_case_number)

				If MAXIS_case_number = "        " then exit do			'Exits do if we reach the end

				'Cleaning up each program's status
				SNAP_status = trim(replace(SNAP_status, "_", ""))
				cash_status = trim(replace(cash_status, "_", ""))
				HC_status = trim(replace(HC_status, "_", ""))
				EA_status = trim(replace(EA_status, "_", ""))
				GRH_status = trim(replace(GRH_status, "_", ""))
				IVE_status = trim(replace(IVE_status, "_", ""))
				CC_status = trim(replace(CC_status, "_", ""))

				'Using if...thens to decide if a case should be added (status isn't blank and respective box is checked)
				If SNAP_status <> "" and SNAP_check = checked then add_case_info_to_Excel = True
				If cash_status <> "" and cash_check = checked then add_case_info_to_Excel = True
				If HC_status <> "" and HC_check = checked then add_case_info_to_Excel = True
				If EA_status <> "" and EA_check = checked then add_case_info_to_Excel = True
				If GRH_status <> "" and GRH_check = checked then add_case_info_to_Excel = True
				If IVE_status <> "" and IVE_check = checked then add_case_info_to_Excel = True
				If CC_status <> "" and CC_check = checked then add_case_info_to_Excel = True
				
				If add_case_info_to_Excel = True then
					ObjExcel.Cells(excel_row, 1).Value = worker
					ObjExcel.Cells(excel_row, 2).Value = MAXIS_case_number
					ObjExcel.Cells(excel_row, 3).Value = client_name
					ObjExcel.Cells(excel_row, 4).Value = replace(APPL_date, " ", "/")
					ObjExcel.Cells(excel_row, 5).Value = abs(days_pending)
					If SNAP_check = checked then ObjExcel.Cells(excel_row, snap_pends_col).Value = SNAP_status
					If cash_check = checked then ObjExcel.Cells(excel_row, cash_pends_col).Value = cash_status
					If HC_check = checked then ObjExcel.Cells(excel_row, HC_pends_col).Value = HC_status
					If EA_check = checked then ObjExcel.Cells(excel_row, EA_pends_col).Value = EA_status
					If GRH_check = checked then ObjExcel.Cells(excel_row, GRH_pends_col).Value = GRH_status
					If IVE_check = checked then ObjExcel.Cells(excel_row, ive_pends_col).Value = IVE_status
					If CC_check = checked then ObjExcel.Cells(excel_row, CC_pends_col).Value = CC_status
					excel_row = excel_row + 1
				End if
				MAXIS_row = MAXIS_row + 1
				add_case_info_to_Excel = ""	'Blanking out variable
				MAXIS_case_number = ""			'Blanking out variable
			Loop until MAXIS_row = 19
			PF8
			EMReadScreen last_page_check, 21, 24, 2
		Loop until last_page_check = "THIS IS THE LAST PAGE"
	End if
	STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
next

'This section adds the most rencent case note information (date, x number and case note to the Excel list. The user will need to select this option in the checkbox on the dialog.)
If case_note_check = checked then 		'all this fun stuff happens
	col_to_use = col_to_use + 1			'scoots over 1 column 
	ObjExcel.Cells(1, col_to_use).Value = "Most recent case note information"	'title of col info
	
	excel_row = 2		'starting with row 2 (1st cell with case information)
	Do 
		MAXIS_case_number = ObjExcel.Cells(excel_row, 2).Value		'establishing what the case number is for each case
		If MAXIS_case_number = "" then exit do						'leaves do if no case number is on the next Excel row
		Call navigate_to_MAXIS_screen("CASE", "NOTE")				'headin' over to CASE/NOTE
		EMReadScreen case_note_info, 74 , 5, 6						'reads the most recent case note
		If trim(case_note_info) <> "" then ObjExcel.Cells(excel_row, col_to_use).Value = case_note_info	'If it's not blank, then it writes the information into Excel 
		excel_row = excel_row + 1									'moves Excel to next row 
	LOOP until MAXIS_case_number = ""								'Loops until all the case have been noted 
END IF 
	
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

is_not_blank_excel_string = chr(34) & "<>" & chr(34)

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

'IVE info
If IVE_check = checked then
	ObjExcel.Cells(row_to_use, col_to_use - 1).Value = "IVE cases pending over 30 days:"	'Row header
	objExcel.Cells(row_to_use, col_to_use - 1).Font.Bold = TRUE						'Row header should be bold
	ObjExcel.Cells(row_to_use, col_to_use).Value = "=COUNTIFS(E:E, " & Chr(34) & ">30" & Chr(34) & ", " & IVE_letter_col & ":" & IVE_letter_col & ", " & is_not_blank_excel_string & ")"	'Excel formula
	ObjExcel.Cells(row_to_use + 1, col_to_use - 1).Value = "Percentage of SNAP cases pending over 30 days:"	'Row header
	objExcel.Cells(row_to_use + 1, col_to_use - 1).Font.Bold = TRUE								'Row header should be bold
	ObjExcel.Cells(row_to_use + 1, col_to_use).Value = "=(COUNTIFS(E:E, " & Chr(34) & ">30" & Chr(34) & ", " & IVE_letter_col & ":" & IVE_letter_col & ", " & is_not_blank_excel_string & "))/(COUNTIF(" & IVE_letter_col & ":" & IVE_letter_col & ", " & is_not_blank_excel_string & ") -1)" 'Excel formula
	ObjExcel.Cells(row_to_use + 1, col_to_use).NumberFormat = "0.00%"		'Formula should be percent
	row_to_use = row_to_use + 2	'It's two rows we jump, because the IVE stat takes up two rows
End if

'CC
If CC_check = checked then
	ObjExcel.Cells(row_to_use, col_to_use - 1).Value = "SNAP cases pending over 30 days:"	'Row header
	objExcel.Cells(row_to_use, col_to_use - 1).Font.Bold = TRUE						'Row header should be bold
	ObjExcel.Cells(row_to_use, col_to_use).Value = "=COUNTIFS(E:E, " & Chr(34) & ">30" & Chr(34) & ", " & CC_letter_col & ":" & CC_letter_col & ", " & is_not_blank_excel_string & ")"	'Excel formula
	ObjExcel.Cells(row_to_use + 1, col_to_use - 1).Value = "Percentage of SNAP cases pending over 30 days:"	'Row header
	objExcel.Cells(row_to_use + 1, col_to_use - 1).Font.Bold = TRUE								'Row header should be bold
	ObjExcel.Cells(row_to_use + 1, col_to_use).Value = "=(COUNTIFS(E:E, " & Chr(34) & ">30" & Chr(34) & ", " & CC_letter_col & ":" & CC_letter_col & ", " & is_not_blank_excel_string & "))/(COUNTIF(" & CC_letter_col & ":" & CC_letter_col & ", " & is_not_blank_excel_string & ") -1)" 'Excel formula
	ObjExcel.Cells(row_to_use + 1, col_to_use).NumberFormat = "0.00%"		'Formula should be percent
	row_to_use = row_to_use + 2	'It's two rows we jump, because the CC stat takes up two rows
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

'Provides additional statistics for IV-E cases
If IVE_check = checked then

	'Going to another sheet, to enter worker-specific statistics
	ObjExcel.Worksheets.Add().Name = "IV-E stats by worker"

	'Headers
	ObjExcel.Cells(1, 2).Value = "IV-E STATS BY WORKER"
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
		ObjExcel.Cells(x + 3, 2) = "=COUNTIFS('Case information'!" & IVE_letter_col & ":" & IVE_letter_col & ", " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'Case information'!A:A, A" & x + 3 & ", 'Case information'!E:E, " & Chr(34) & "<=30" & Chr(34) & ")"
		ObjExcel.Cells(x + 3, 3) = "=COUNTIFS('Case information'!" & IVE_letter_col & ":" & IVE_letter_col & ", " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'Case information'!A:A, A" & x + 3 & ")"
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

'Provides additional statistics for CC cases
If CC_check = checked then

	'Going to another sheet, to enter worker-specific statistics
	ObjExcel.Worksheets.Add().Name = "CC stats by worker"

	'Headers
	ObjExcel.Cells(1, 2).Value = "CC STATS BY WORKER"
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
		ObjExcel.Cells(x + 3, 2) = "=COUNTIFS('Case information'!" & CC_letter_col & ":" & CC_letter_col & ", " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'Case information'!A:A, A" & x + 3 & ", 'Case information'!E:E, " & Chr(34) & "<=30" & Chr(34) & ")"
		ObjExcel.Cells(x + 3, 3) = "=COUNTIFS('Case information'!" & CC_letter_col & ":" & CC_letter_col & ", " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'Case information'!A:A, A" & x + 3 & ")"
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

'Navigates back to the case information worksheet so this is the 1st worksheet the user sees
objExcel.worksheets("Case Information").Activate			'Activates the selected worksheet'

'Logging usage stats
STATS_counter = STATS_counter - 1                      'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
script_end_procedure("Success! Your REPT/PND2 list has been created.")