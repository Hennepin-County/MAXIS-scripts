'Required for statistical purposes===============================================================================
name_of_script = "BULK - REPT-INTR LIST.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 35                               'manual run time, per line, in seconds
STATS_denomination = "I"       'I is for each ITEM
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

'DIALOG=============================================================================
BeginDialog bulk_paris_report_dialog, 0, 0, 361, 160, "Bulk DAIL report dialog"
  EditBox 10, 35, 345, 15, x_number_editbox
  DropListBox 230, 70, 120, 45, "AL - All"+chr(9)+"UR - Unresolved"+chr(9)+"AR - All Resolved Matches"+chr(9)+"PR - Person Removed"+chr(9)+"HM - Household Moved"+chr(9)+"RV - MN Residency Verified"+chr(9)+"FR - Failed Residency Verificaion"+chr(9)+"PC - Person Closed NOT PARIS"+chr(9)+"CC - Case Closed NOT PARIS", res_status_list
  EditBox 85, 90, 15, 15, start_month
  EditBox 100, 90, 15, 15, start_year
  EditBox 195, 90, 15, 15, end_month
  EditBox 210, 90, 15, 15, end_year
  CheckBox 10, 125, 150, 10, "Check here to run this query county-wide.", all_workers_check
  ButtonGroup ButtonPressed
    OkButton 250, 140, 50, 15
    CancelButton 305, 140, 50, 15
  Text 145, 5, 90, 10, "---BULK INTR REPORT---"
  Text 10, 20, 350, 10, "Please enter the x1 numbers of the caseloads you wish to check, separated by commas (if more than one):"
  Text 10, 55, 290, 10, "Note: please enter the entire 7-digit number x1 number. (Example: ''x100abc, x100abc'')"
  Text 10, 75, 220, 10, "Select the resolution status for the matches you would like to pull:"
  Text 10, 95, 75, 10, "Search Start (MM YY)"
  Text 125, 95, 70, 10, "SearchEnd (MM YY)"
  Text 20, 110, 190, 10, "NOTE: Leave dates blank to default to the past 12 months"
  Text 20, 140, 210, 20, "NOTE: running queries county-wide can take a significant amount of time and resources. This should be done after hours."
EndDialog

'=================================================================================
'Connects to MAXIS
EMConnect ""

'Looks up an existing user for autofilling the next dialog
CALL find_variable("User: ", x_number_editbox, 7)

'Shows the dialog.
DO
	Do
		err_msg = ""
		dialog bulk_paris_report_dialog
		cancel_confirmation
		If start_month <> "" AND isnumeric(start_month) = false then err_msg = err_msg & vbNewLine & "Please enter a number in the start month field."
		If start_year <> "" AND isnumeric(start_year) = false then err_msg = err_msg & vbNewLine & "Please enter a number in the start year field."
		If end_month <> "" AND isnumeric(end_month) = false then err_msg = err_msg & vbNewLine & "Please enter a number in the end month field."
		If end_year <> "" AND isnumeric(end_year) = false then err_msg = err_msg & vbNewLine & "Please enter a number in the year year field."
		If err_msg <> "" Then MsgBox "Please resolve:" & vbNewLine & err_msg
	Loop until err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'Starting the query start time (for the query runtime at the end)
query_start_time = timer
entire_county = FALSE

resolution_code = left(res_status_list, 2)

'If all workers are selected, the script will go to REPT/USER, and load all of the workers into an array. Otherwise it'll create a single-object "array" just for simplicity of code.
If all_workers_check = checked then
	call create_array_of_all_active_x_numbers_in_county(x_number_array, two_digit_county_code)
	entire_county = TRUE
Else
	'splits the results of the editbox into an array
	x_number_array = split(x_number_editbox, ",")
End if

'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

'Name for the current sheet'
ObjExcel.ActiveSheet.Name = "Case information"

'Excel headers and formatting the columns
objExcel.Cells(1, 1).Value     = "SUPERVISOR ID"
objExcel.Cells(1, 1).Font.Bold = True
objExcel.Cells(1, 2).Value     = "X1 NUMBER"
objExcel.Cells(1, 2).Font.Bold = True
objExcel.Cells(1, 3).Value     = "WORKER NAME"
objExcel.Cells(1, 3).Font.Bold = True
objExcel.Cells(1, 4).Value     = "CASE NBR"
objExcel.Cells(1, 4).Font.Bold = True
objExcel.Cells(1, 5).Value     = "CLIENT NAME"
objExcel.Cells(1, 5).Font.Bold = True
objExcel.Cells(1, 6).Value     = "MONTH"
objExcel.Cells(1, 6).Font.Bold = True
objExcel.Cells(1, 7).Value     = "RESOLUTION STATUS"
objExcel.Cells(1, 7).Font.Bold = True
objExcel.Cells(1, 8).Value     = "INTERSTATE MATCH NOTICE DATE"
objExcel.Cells(1, 8).Font.Bold = True


'This bit freezes the top row of the Excel sheet for better useability when there is a lot of information
ObjExcel.ActiveSheet.Range("A2").Select
objExcel.ActiveWindow.FreezePanes = True

'Sets variable for all of the Excel stuff
excel_row = 2

supervisor_number = FALSE

'This for...next contains each worker indicated above
For each x_number in x_number_array

	'Trims the x_number so that we don't have glitches
	x_number = trim(x_number)
	x_number = UCase(x_number)

	'Going to idetify if the x-number is for a supervisor
	back_to_SELF
	CALL navigate_to_MAXIS_screen("REPT", "USER")		'Go to REPT USER and refresh to supervisor view
	PF5
	PF5
	EMWriteScreen x_number, 21, 12						'Enter the X number'
	transmit
	If entire_county = FALSE Then 						'If running agency wide - this is not needed because every worker will be selected
		EMReadScreen worker_number, 7, 7, 5				'Otherwise if there is a worker number listed under sup in this view - they are a supervisor
		If worker_number <> "       " Then supervisor_number = TRUE
	End IF
	EMReadScreen worker_name, 30, 3, 47					'Gets the worker name from REPT USER
	worker_name = trim(worker_name)

	back_to_SELF
	CALL navigate_to_MAXIS_screen("REPT", "INTR")	'Navigates to the worker based report'
	If supervisor_number = FALSE then EMWriteScreen x_number, 5, 15		'For workers, the number is put on the worker line
	If supervisor_number = TRUE then 									'For supervisors, the worker line is blanked out and the enumber is put on the supervisor line
		EMWriteScreen x_number, 6, 15
		EMWriteScreen "       ", 5, 15
	End If
	EMWriteScreen resolution_code, 5, 67			'Entering the resolution code selected in dialog
	If start_month <> "" Then
		start_month = right("00" & start_month, 2)
		EMWriteScreen start_month, 6, 67			'Entering the date range if selected
	End If
	If start_year <> "" Then
		start_year = right("00" & start_year, 2)
		EMWriteScreen start_year, 6, 70
	End If
	If end_month <> "" Then
		end_month = right("00" & end_month, 2)
		EMWriteScreen end_month, 7, 67
	End If
	If end_year <> "" Then
		end_year = right("00" & end_year, 2)
		EMWriteScreen end_year, 7, 70
	End If
	transmit										'and GO

	EMReadScreen intr_exists, 8, 11, 5				'Looking if there are any matches listed under this worker
	intr_exists = trim(intr_exists)
	row = 11
	If intr_exists <> "" Then 	'If there are any matches the script will pull detail
		Do
			EMReadScreen maxis_case_number, 8, row, 5			'Reading the case number
			maxis_case_number = trim(maxis_case_number)			'removing the spaces
			If maxis_case_number = "" then exit Do 		'Once the script reaches the last line in the list, it will go to the next worker

			EMReadScreen match_x_number, 7, row, 14				'Reading the worker x-number listed on the match - necessary if the number in the array is a supervisor number
			EMReadScreen supervisor_id, 7, 6, 15				'Reading the x-number of the supervisor for that case
			EMReadScreen client_name, 20, row, 31				'Reading the client name and removing the blanks
			client_name = trim(client_name)
			EMReadScreen match_month, 5, row, 53				'Reading the month the match was issued
			match_month = replace(match_month, " ", "/01/")		'Formatting the date as a date for entry into Excel
			EMReadScreen res_status, 2, row, 64					'Reading the resolution status
			EMReadScreen notice_date, 8, row, 71				'Reading the notice date field
			if notice_date = "        " then notice_date = ""	'blanking out if there is no date
			notice_date = replace(notice_date, " ", "/")		'Formatting the date

			'Adding all the information to Excel
			objExcel.Cells(excel_row, 1).Value = supervisor_id
			objExcel.Cells(excel_row, 2).Value = match_x_number
			If supervisor_number = FALSE Then objExcel.Cells(excel_row, 3).Value = worker_name
			objExcel.Cells(excel_row, 4).Value = maxis_case_number
			objExcel.Cells(excel_row, 5).Value = client_name
			objExcel.Cells(excel_row, 6).Value = match_month
			objExcel.Cells(excel_row, 7).Value = res_status
			objExcel.Cells(excel_row, 8).Value = notice_date

			row = row + 1		'Go to the next excel row

			If row = 19 Then 		'If we have reached the end of the page, it will go to the next page
				PF8
				row = 11			'Resets the row
				EMReadScreen last_page_check, 21, 24, 2
			End If
			excel_row = excel_row + 1	'increments the excel row so we don't overwrite our data
			STATS_counter = STATS_counter + 1		'Counts 1 item for every Match found and entered into excel.			diff_notc_date = ""			'blanks this out so that the information is not carried over in the do-loop'
			maxis_case_number = ""

		Loop until last_page_check = "THIS IS THE LAST PAGE"

	Else
		objExcel.Cells(excel_row, 2).Value = x_number						'If there are no items on a worker's report, this adds that information to the excel sheet
		EMReadScreen supervisor_id, 7, 6, 15				'Finding Supervisor X Number and adding to Excel
		objExcel.Cells(excel_row, 1).Value = supervisor_id
		objExcel.Cells(excel_row, 3).Value = worker_name
		EMReadScreen INTR_check, 4, 2, 55
		objExcel.Cells(excel_row, 5).Value = "No PARIS Matches for this worker."		'Adds line to Excel sheet indicating no matches
		excel_row = excel_row + 1
	End If

	supervisor_number = FALSE
Next

last_excel_row = excel_row - 1

'Centers the text for the columns with days remaining and difference notice
objExcel.Columns(6).HorizontalAlignment = -4108
objExcel.Columns(7).HorizontalAlignment = -4108
objExcel.Columns(8).HorizontalAlignment = -4108

excel_is_not_blank = chr(34) & "<>" & chr(34)		'Setting up a variable for useable quote marks in Excel

'Query date/time/runtime info
objExcel.Cells(1, 11).Font.Bold = TRUE
objExcel.Cells(2, 11).Font.Bold = TRUE
objExcel.Cells(3, 11).Font.Bold = TRUE
objExcel.Cells(4, 11).Font.Bold = TRUE

ObjExcel.Cells(1, 11).Value = "Query date and time:"	'Goes back one, as this is on the next row
ObjExcel.Cells(1, 12).Value = now
ObjExcel.Cells(2, 11).Value = "Query runtime (in seconds):"	'Goes back one, as this is on the next row
ObjExcel.Cells(2, 12).Value = timer - query_start_time
ObjExcel.Cells(3, 11).Value = "Number of IEVS with No DAYS remaining:"
objExcel.Cells(3, 12).Value = "=COUNTIFS(G:G, " & Chr(34) & "<=0" & Chr(34) & ", H:H, " & excel_is_not_blank & ")"	'Excel formula
ObjExcel.Cells(4, 11).Value = "Number of total UNRESOLVED IEVS:"
objExcel.Cells(4, 12).Value = "=(COUNTIF(H:H, " & excel_is_not_blank & ")-1)"	'Excel formula

'Need a new array for the x-numbers of all the matches because if supervisor numbers are used all the numbers are not in the original array
Dim worker_number_array
ReDim worker_number_array (0)
array_counter = 0

'Going to REPT USER to get the worker's name
CALL navigate_to_MAXIS_screen ("REPT", "USER")
excel_row = 2
Do
	If objExcel.Cells(excel_row, 3).Value = "" Then 			'If the worker's name was not in the excel sheet - the script will look for it
		worker_number = objExcel.Cells(excel_row, 2).Value		'Gets the worker number from the spreadsheet
		in_array = FALSE 										'Setting the variable
		For each worker in worker_number_array					'Looping through the new array to compare it to the current worker number
			if worker_number = worker then
				in_array = TRUE
				Exit For
			End if
		Next
		If in_array = FALSE Then 	'If the x-number was not found already in the array the number will be added to the array
			ReDim Preserve worker_number_array(array_counter)
			worker_number_array(array_counter) = worker_number
			array_counter = array_counter + 1
		End If
		EMWriteScreen worker_number, 21, 12			'Entering the worker number and opening detail about the worker
		transmit
		EMWriteScreen "X", 7, 3
		transmit
		EMReadScreen worker_name, 40, 7, 27			'Reading the worker number
		PF3
		worker_name = trim(worker_name)
		objExcel.Cells(excel_row, 3).Value = worker_name	'Adding the number to Excel
		excel_row = excel_row + 1 					'Go to the next row
		Do
			If objExcel.Cells(excel_row, 2).Value = worker_number Then 	'If the next row has the same number - add the same name
				objExcel.Cells(excel_row, 3).Value = worker_name
				excel_row = excel_row + 1
			Else
				Exit Do
			End If
		loop until objExcel.Cells(excel_row, 3).Value <> worker_number
	Else 																'If the name is already in Excel the number will be reviewed for adding to the array
		worker_number = objExcel.Cells(excel_row, 2).Value		'Gets the worker number from the spreadsheet
		in_array = FALSE 										'Setting the variable
		For each worker in worker_number_array					'Looping through the new array to compare it to the current worker number
			if worker_number = worker then
				in_array = TRUE
				Exit For
			End if
		Next
		If in_array = FALSE Then 			'If the x-number was not found already in the array the number will be added to the array
			ReDim Preserve worker_number_array(array_counter)
			worker_number_array(array_counter) = worker_number
			array_counter = array_counter + 1
		End If
		excel_row = excel_row + 1
	End If
Loop until excel_row = last_excel_row + 1

'Formatting the column width.
FOR i = 1 to 12
	objExcel.Columns(i).AutoFit()
NEXT

'Going to another sheet, to enter worker-specific statistics
ObjExcel.Worksheets.Add().Name = "INTR stats by worker"

'Headers
ObjExcel.Cells(1, 2).Value = "INTR STATS BY WORKER"
ObjExcel.Cells(1, 2).Font.Bold = TRUE
ObjExcel.Cells(2, 1).Value = "WORKER"
objExcel.Cells(2, 1).Font.Bold = TRUE
ObjExcel.Cells(2, 2).Value = "NAME"
ObjExcel.Cells(2, 2).Font.Bold = TRUE
ObjExcel.Cells(2, 3).Value = "UNRESOLVED"
objExcel.Cells(2, 3).Font.Bold = TRUE

'This bit freezes the top 2 rows for scrolling ease of use
ObjExcel.ActiveSheet.Range("A3").Select
objExcel.ActiveWindow.FreezePanes = True

worker_row = 3
'Writes each worker from the worker_array in the Excel spreadsheet
For each x_number in worker_number_array
	'Trims the x_number so that we don't have glitches
	x_number = trim(x_number)
	x_number = UCase(x_number)
	IF right(x_number, 3) <> "CLS" then 	'This bit gets worker names from REPT ACTV
		Call navigate_to_MAXIS_screen ("REPT", "ACTV")
		EMWriteScreen x_number, 21, 13
		transmit
		EMReadScreen worker_name, 24, 3, 11
		worker_name = trim(worker_name)
	Else
		worker_name = "CLOSED RECORDS"		'Except CLS - which takes a long time to load and is Closed Records
	End IF
	'Adding all the information to Excel
	ObjExcel.Cells(worker_row, 1).Value = x_number
	ObjExcel.Cells(worker_row, 2).Value = worker_name
	'Writing a formula to excel - Count each row in which Column G on the first worksheet is not blank AND the x number in Column B on the first worksheet matches the X number on this row and the match is unresolved
	ObjExcel.Cells(worker_row, 3).Value = "=COUNTIFS('Case information'!G:G, " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'Case information'!B:B, A" & worker_row & ", 'Case information'!G:G," & Chr(34) & "UR" & Chr(34) & ")"

	worker_row = worker_row + 1
Next

'Merging header cell.
ObjExcel.Range(ObjExcel.Cells(1, 1), ObjExcel.Cells(1, 3)).Merge

'Centering the cell
objExcel.Cells(1, 2).HorizontalAlignment = -4108

'Autofitting columns
For col_to_autofit = 1 to 3
	ObjExcel.columns(col_to_autofit).AutoFit()
Next

STATS_counter = STATS_counter - 1		'removing the initial counter so that this number is correct.

script_end_procedure("Success! The spreadsheet has all requested information.")
