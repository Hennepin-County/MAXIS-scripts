'Required for statistical purposes===============================================================================
MMname_of_script = "BULK - REPT-MAMS LIST.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 13                      'manual run time in seconds
STATS_denomination = "C"       			   'C is for each CASE
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

'Checks for county info from global variables, or asks if it is not already defined.
get_county_code

'CONSTANTS FOR EXCEL SPREADSHEET--------------------------------------------------------------
const worker_col 	= 1
const case_col 		= 2
const name_col		= 3
const status_col 	= 4
const revw_date_col	= 5

'DIALOGS----------------------------------------------------------------------
BeginDialog pull_REPT_data_into_excel_dialog, 0, 0, 231, 130, "Pull REPT/MAMS data into Excel dialog"
  EditBox 80, 20, 145, 15, worker_number
  CheckBox 10, 65, 150, 10, "Check here to run this query county-wide.", all_workers_check
  ButtonGroup ButtonPressed
    OkButton 120, 110, 50, 15
    CancelButton 175, 110, 50, 15
  Text 10, 25, 65, 10, "Worker(s) to check:"
  Text 10, 80, 215, 20, "NOTE: running queries county-wide can take a significant amount of time/resources."
  Text 55, 5, 125, 10, "***PULL REPT MAMS INTO EXCEL***"
  Text 10, 40, 210, 20, "Enter the 7-digit worker number(s) (ex: x1#####), separated by a comma."
EndDialog

'THE SCRIPT---------------------------------------------------
'Connects to BlueZone
EMConnect ""

'Shows dialog
DO
	Dialog pull_rept_data_into_Excel_dialog
	If buttonpressed = cancel then script_end_procedure("")
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'Starting the query start time (for the query runtime at the end)
query_start_time = timer

'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

'Setting the first 3 col as worker, case number, and name
ObjExcel.Cells(1, worker_col).Value = "WORKER"
ObjExcel.Cells(1, case_col).Value = "CASE NUMBER"
ObjExcel.Cells(1, name_col).Value = "NAME"
ObjExcel.Cells(1, status_col).Value = "STATUS"
ObjExcel.Cells(1, revw_date_col).Value = "DATE REVW REC'D"

HC_letter_col = convert_digit_to_excel_column(status_col)

FOR i = worker_col to revw_date_col		'formatting the cells'
	objExcel.Cells(1, i).Font.Bold = True		'bold font'
	objExcel.Columns(i).AutoFit()						'sizing the colums'
NEXT

'Figuring out what to put in each Excel col. To add future variables to this, add the checkbox variables below and copy/paste the same code!
'	Below, use the "[blank]_col" variable to recall which col you set for which option.
col_to_use = 6 'Starting with 6 because cols 1-5 are already used

'If all workers are selected, the script will go to REPT/USER, and load all of the workers into an array. Otherwise it'll create a single-object "array" just for simplicity of code.
If all_workers_check = checked then
	call create_array_of_all_active_x_numbers_in_county(worker_array, two_digit_county_code)
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

'Setting the variable for what's to come
excel_row = 2

For each worker in worker_array
	back_to_self	'Does this to prevent "ghosting" where the old info shows up on the new screen for some reason
	Call navigate_to_MAXIS_screen("REPT", "MAMS")
	EMWriteScreen worker, 21, 6
	transmit

	'Skips workers with no info
	EMReadScreen has_content_check, 8, 7, 6
	If has_content_check <> "        " then

		'Grabbing each case number on screen
		Do
			'Set variable for next do...loop
			MAXIS_row = 7
			Do
				EMReadScreen MAXIS_case_number, 8, MAXIS_row, 6			'Reading case number
				EMReadScreen client_name, 15, MAXIS_row, 16		'Reading client name
				EMReadScreen HRF_status, 2, MAXIS_row, 45		'Reading HRF status
				EMReadScreen revw_recd_date, 8, MAXIS_row, 54	'Reading review received date

				'Doing this because sometimes BlueZone registers a "ghost" of previous data when the script runs. This checks against an array and stops if we've seen this one before.
				If trim(MAXIS_case_number) <> "" and instr(all_case_numbers_array, MAXIS_case_number) <> 0 then exit do
				all_case_numbers_array = trim(all_case_numbers_array & " " & MAXIS_case_number)

				If MAXIS_case_number = "        " then exit do			'Exits do if we reach the end

				'For some goofy reason the dash key shows up instead of the space key. No clue why. This will turn them into null variables.
				If HRF_status = "-" then HRF_status = ""

				'Cleaning up the blank revw_recd_date and interview_date variables
				revw_recd_date = trim(replace(revw_recd_date, "__ __ __", ""))

				'Adding the case to Excel
				ObjExcel.Cells(excel_row, worker_col).Value = worker
				ObjExcel.Cells(excel_row, case_col).Value = MAXIS_case_number
				ObjExcel.Cells(excel_row, name_col).Value = client_name
				ObjExcel.Cells(excel_row, status_col).Value = HRF_status
				ObjExcel.Cells(excel_row, revw_date_col).Value = replace(revw_recd_date, " ", "/")
				excel_row = excel_row + 1
				MAXIS_row = MAXIS_row + 1
				add_case_info_to_Excel = ""	'Blanking out variable
				MAXIS_case_number = ""			'Blanking out variable
			Loop until MAXIS_row = 19
			PF8
			EMReadScreen last_page_check, 21, 24, 2	'checking to see if we're at the end
		Loop until last_page_check = "THIS IS THE LAST PAGE"
	End if
	STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
next

'IF error_check = checked THEN col_to_use = col_to_use + 1 'need to add the extra column for this outside the DO...LOOP
col_to_use = col_to_use + 2	'Doing two because the wrap-up is two columns
row_to_use = 3			'Declaring here before the following if...then statements

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
STATS_counter = STATS_counter - 1                      'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
script_end_procedure("Success! Your list has been created.")
