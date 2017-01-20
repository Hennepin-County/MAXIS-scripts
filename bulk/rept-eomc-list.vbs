'Required for statistical purposes==========================================================================================
name_of_script = "BULK - REPT-EOMC LIST.vbs"
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
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'DIALOGS-------------------------------------------------------------------------------------------------------------
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
  Text 70, 40, 210, 20, "Enter all 7 digits of your workers' x1 numbers (ex: x######), separated by a comma."
EndDialog

'VARIABLES TO DECLARE------------------------------------------------------------------------------------------------------------------
all_case_numbers_array = " "					'Creating blank variable for the future array
get_county_code	'Determines worker county code
is_not_blank_excel_string = Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34)	'This is the string required to tell excel to ignore blank cells in a COUNTIFS function

'THE SCRIPT-----------------------------------------------------------------------------------------------------------
'Connects to BlueZone
EMConnect ""

'Shows dialog
Dialog pull_rept_data_into_Excel_dialog
If buttonpressed = cancel then stopscript

'Starting the query start time (for the query runtime at the end)
query_start_time = timer

'Checking for MAXIS
Call check_for_MAXIS(True)

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
ObjExcel.Cells(1, 4).Value = "AUTOCLOSE?"
objExcel.Cells(1, 4).Font.Bold = TRUE

'Figuring out what to put in each Excel col. To add future variables to this, add the checkbox variables below and copy/paste the same code!
'	Below, use the "[blank]_col" variable to recall which col you set for which option.
col_to_use = 5 'Starting with 5 because cols 1-4 are already used

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
	ObjExcel.Cells(1, col_to_use).Value = "HC?"
	objExcel.Cells(1, col_to_use).Font.Bold = TRUE
	HC_actv_col = col_to_use
	col_to_use = col_to_use + 1
	HC_letter_col = convert_digit_to_excel_column(HC_actv_col)
End if
If GRH_check = checked then
	ObjExcel.Cells(1, col_to_use).Value = "GRH?"
	objExcel.Cells(1, col_to_use).Font.Bold = TRUE
	GRH_actv_col = col_to_use
	col_to_use = col_to_use + 1
	GRH_letter_col = convert_digit_to_excel_column(GRH_actv_col)
End if

'If all workers are selected, the script will go to REPT/USER, and load all of the workers into an array. Otherwise it'll create a single-object "array" just for simplicity of code.
If all_workers_check = checked then
	call create_array_of_all_active_x_numbers_in_county(worker_array, two_digit_county_code)
Else
	x1s_from_dialog = split(worker_number, ",")	'Splits the worker array based on commas

	'building array
	For each x1_number in x1s_from_dialog
		If worker_array = "" then
			worker_array = trim(ucase(x1_number))	'replaces worker_county_code if found in the typed x1 number
		Else
			worker_array = worker_array & ", " & trim(ucase(x1_number)) 'replaces worker_county_code if found in the typed x1 number
		End if
	Next

	'Split worker_array
	worker_array = split(worker_array, ", ")
End if

'Setting the variable for what's to come
excel_row = 2

For each worker in worker_array
	back_to_self	'Does this to prevent "ghosting" where the old info shows up on the new screen for some reason
	Call navigate_to_MAXIS_screen("rept", "eomc")
	EMWriteScreen worker, 21, 16
	transmit

	'Skips workers with no info
	EMReadScreen has_content_check, 1, 7, 5
	If has_content_check <> " " then

		'Grabbing each case number on screen
		Do

			EMReadScreen last_page_check, 21, 24, 2	'because on REPT/EOMC it displays right away, instead of when the second F8 is sent

			'Set variable for next do...loop
			MAXIS_row = 7
			Do
				EMReadScreen MAXIS_case_number, 8, MAXIS_row, 7			'Reading case number
				EMReadScreen client_name, 25, MAXIS_row, 16		'Reading client name
				EMReadScreen cash_status, 4, MAXIS_row, 43		'Reading cash status
				EMReadScreen SNAP_status, 4, MAXIS_row, 53		'Reading SNAP status
				EMReadScreen HC_status, 4, MAXIS_row, 58			'Reading HC status
				EMReadScreen GRH_status, 4, MAXIS_row, 68			'Reading GRH status

				'Doing this because sometimes BlueZone registers a "ghost" of previous data when the script runs. This checks against an array and stops if we've seen this one before.
				If trim(MAXIS_case_number) <> "" and instr(all_case_numbers_array, MAXIS_case_number) <> 0 then exit do
				all_case_numbers_array = trim(all_case_numbers_array & " " & MAXIS_case_number)

				If MAXIS_case_number = "        " then exit do			'Exits do if we reach the end

				'Using if...thens to decide if a case should be added (status isn't blank or inactive and respective box is checked)
				If cash_status <> "    " and cash_check = checked then add_case_info_to_Excel = True
				If SNAP_status <> "    " and SNAP_check = checked then add_case_info_to_Excel = True
				If HC_status <> "    " and HC_check = checked then add_case_info_to_Excel = True
				If GRH_status <> "    " and GRH_check = checked then add_case_info_to_Excel = True

				'Determines if any programs are autoclosing, and creates an autoclose string containing that info
				If cash_check = checked and right(cash_status, 1) = "A" then autoclose_string = autoclose_string & left(cash_status, 2) & " "
				If SNAP_check = checked and right(SNAP_status, 1) = "A" then autoclose_string = autoclose_string & left(SNAP_status, 2) & " "
				If HC_check = checked and right(HC_status, 1) = "A" then autoclose_string = autoclose_string & left(HC_status, 2) & " "
				If GRH_check = checked and right(GRH_status, 1) = "A" then autoclose_string = autoclose_string & left(GRH_status, 2) & " "

				If add_case_info_to_Excel = True then
					ObjExcel.Cells(excel_row, 1).Value = worker
					ObjExcel.Cells(excel_row, 2).Value = MAXIS_case_number
					ObjExcel.Cells(excel_row, 3).Value = client_name
					ObjExcel.Cells(excel_row, 4).Value = trim(autoclose_string)
					If SNAP_check = checked then ObjExcel.Cells(excel_row, snap_actv_col).Value = trim(SNAP_status)
					If cash_check = checked then ObjExcel.Cells(excel_row, cash_actv_col).Value = trim(cash_status)
					If HC_check = checked then ObjExcel.Cells(excel_row, HC_actv_col).Value = trim(HC_status)
					If GRH_check = checked then ObjExcel.Cells(excel_row, GRH_actv_col).Value = trim(GRH_status)
					excel_row = excel_row + 1
				End if
				MAXIS_row = MAXIS_row + 1
				add_case_info_to_Excel = ""	'Blanking out variable
				autoclose_string = ""		'Blanking out variable
				MAXIS_case_number = ""			'Blanking out variable
			Loop until MAXIS_row = 19
			PF8
		Loop until last_page_check = "THIS IS THE LAST PAGE"
	End if
	STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
next

col_to_use = col_to_use + 2	'Doing two because the wrap-up is two columns
row_to_use = 3			'For the individual program-breakdown of info

'Query date/time/runtime info
objExcel.Cells(1, col_to_use - 1).Font.Bold = TRUE
objExcel.Cells(2, col_to_use - 1).Font.Bold = TRUE
ObjExcel.Cells(1, col_to_use - 1).Value = "Query date and time:"	'Goes back one, as this is on the next row
ObjExcel.Cells(1, col_to_use).Value = now
ObjExcel.Cells(2, col_to_use - 1).Value = "Query runtime (in seconds):"	'Goes back one, as this is on the next row
ObjExcel.Cells(2, col_to_use).Value = timer - query_start_time

'SNAP info
If SNAP_check = checked then
	ObjExcel.Cells(row_to_use, col_to_use - 1).Value = "SNAP cases that are closing:"	'Row header
	objExcel.Cells(row_to_use, col_to_use - 1).Font.Bold = TRUE						'Row header should be bold
	ObjExcel.Cells(row_to_use, col_to_use).Value = "=COUNTA(" & SNAP_letter_col & ":" & SNAP_letter_col & ") - 1"	'Excel formula
	ObjExcel.Cells(row_to_use + 1, col_to_use - 1).Value = "Percentage of SNAP cases autoclosing:"	'Row header
	objExcel.Cells(row_to_use + 1, col_to_use - 1).Font.Bold = TRUE								'Row header should be bold
	ObjExcel.Cells(row_to_use + 1, col_to_use).Value = "=(COUNTIFS(D:D, " & chr(34) & "*FS*" & chr(34) & ", " & SNAP_letter_col & ":" & SNAP_letter_col & ", " & is_not_blank_excel_string & "))/(COUNTA(" & SNAP_letter_col & ":" & SNAP_letter_col & ") - 1)" 'Excel formula
	ObjExcel.Cells(row_to_use + 1, col_to_use).NumberFormat = "0.00%"		'Formula should be percent
	row_to_use = row_to_use + 2	'It's two rows we jump, because the SNAP stat takes up two rows
End if

'HC info
If HC_check = checked then
	ObjExcel.Cells(row_to_use, col_to_use - 1).Value = "HC cases that are closing:"	'Row header
	objExcel.Cells(row_to_use, col_to_use - 1).Font.Bold = TRUE						'Row header should be bold
	ObjExcel.Cells(row_to_use, col_to_use).Value = "=COUNTA(" & HC_letter_col & ":" & HC_letter_col & ") - 1"	'Excel formula
	ObjExcel.Cells(row_to_use + 1, col_to_use - 1).Value = "Percentage of HC cases autoclosing:"	'Row header
	objExcel.Cells(row_to_use + 1, col_to_use - 1).Font.Bold = TRUE								'Row header should be bold
	ObjExcel.Cells(row_to_use + 1, col_to_use).Value = "=(COUNTIFS(D:D, " & chr(34) & "*HC*" & chr(34) & ", " & HC_letter_col & ":" & HC_letter_col & ", " & is_not_blank_excel_string & "))/(COUNTA(" & HC_letter_col & ":" & HC_letter_col & ") - 1)" 'Excel formula
	ObjExcel.Cells(row_to_use + 1, col_to_use).NumberFormat = "0.00%"		'Formula should be percent
	row_to_use = row_to_use + 2	'It's two rows we jump, because the HC stat takes up two rows
End if

'GRH info
If GRH_check = checked then
	ObjExcel.Cells(row_to_use, col_to_use - 1).Value = "GRH cases that are closing:"	'Row header
	objExcel.Cells(row_to_use, col_to_use - 1).Font.Bold = TRUE						'Row header should be bold
	ObjExcel.Cells(row_to_use, col_to_use).Value = "=COUNTA(" & GRH_letter_col & ":" & GRH_letter_col & ") - 1"	'Excel formula
	ObjExcel.Cells(row_to_use + 1, col_to_use - 1).Value = "Percentage of GRH cases autoclosing:"	'Row header
	objExcel.Cells(row_to_use + 1, col_to_use - 1).Font.Bold = TRUE								'Row header should be bold
	ObjExcel.Cells(row_to_use + 1, col_to_use).Value = "=(COUNTIFS(D:D, " & chr(34) & "*GR*" & chr(34) & ", " & GRH_letter_col & ":" & GRH_letter_col & ", " & is_not_blank_excel_string & "))/(COUNTA(" & GRH_letter_col & ":" & GRH_letter_col & ") - 1)" 'Excel formula
	ObjExcel.Cells(row_to_use + 1, col_to_use).NumberFormat = "0.00%"		'Formula should be percent
	row_to_use = row_to_use + 2	'It's two rows we jump, because the GRH stat takes up two rows
End if

'cash info
If cash_check = checked then
	ObjExcel.Cells(row_to_use, col_to_use - 1).Value = "Cash cases that are closing:"	'Row header
	objExcel.Cells(row_to_use, col_to_use - 1).Font.Bold = TRUE						'Row header should be bold
	ObjExcel.Cells(row_to_use, col_to_use).Value = "=COUNTA(" & cash_letter_col & ":" & cash_letter_col & ") - 1"	'Excel formula
	ObjExcel.Cells(row_to_use + 1, col_to_use - 1).Value = "Percentage of cash cases autoclosing:"	'Row header
	objExcel.Cells(row_to_use + 1, col_to_use - 1).Font.Bold = TRUE								'Row header should be bold
	ObjExcel.Cells(row_to_use + 1, col_to_use).Value = "=(COUNTIFS(D:D, " & chr(34) & "*" & chr(34) & ", " & cash_letter_col & ":" & cash_letter_col & ", " & is_not_blank_excel_string & ", " & cash_letter_col & ":" & cash_letter_col & ", " & chr(34) & "*/A*" & chr(34) & "))/(COUNTA(" & cash_letter_col & ":" & cash_letter_col & ") - 1)" 'Excel formula
	ObjExcel.Cells(row_to_use + 1, col_to_use).NumberFormat = "0.00%"		'Formula should be percent
	row_to_use = row_to_use + 2	'It's two rows we jump, because the cash stat takes up two rows
End if

'Autofitting columns
For col_to_autofit = 1 to col_to_use
	ObjExcel.columns(col_to_autofit).AutoFit()
Next

STATS_counter = STATS_counter - 1                      'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
script_end_procedure("")
