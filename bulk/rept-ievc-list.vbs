'Required for statistical purposes===============================================================================
name_of_script = "BULK - REPT-IEVC LIST.vbs"
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
BeginDialog bulk_ievs_report_dialog, 0, 0, 361, 105, "Bulk DAIL report dialog"
  EditBox 10, 35, 345, 15, x_number_editbox
  CheckBox 10, 70, 150, 10, "Check here to run this query county-wide.", all_workers_check
  ButtonGroup ButtonPressed
    OkButton 250, 85, 50, 15
    CancelButton 305, 85, 50, 15
  Text 145, 5, 90, 10, "---BULK IEVS REPORT---"
  Text 10, 20, 350, 10, "Please enter the x1 numbers of the caseloads you wish to check, separated by commas (if more than one):"
  Text 10, 55, 290, 10, "Note: please enter the entire 7-digit number x1 number. (Example: ''x100abc, x100abc'')"
  Text 20, 85, 210, 20, "NOTE: running queries county-wide can take a significant amount of time and resources. This should be done after hours."
EndDialog
'=================================================================================
'Connects to MAXIS
EMConnect ""

'Looks up an existing user for autofilling the next dialog
CALL find_variable("User: ", x_number_editbox, 7)

'Shows the dialog.
DO
	dialog bulk_ievs_report_dialog
	cancel_confirmation
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'Starting the query start time (for the query runtime at the end)
query_start_time = timer

'If all workers are selected, the script will go to REPT/USER, and load all of the workers into an array. Otherwise it'll create a single-object "array" just for simplicity of code.
If all_workers_check = checked then
	call create_array_of_all_active_x_numbers_in_county(x_number_array, two_digit_county_code)
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
objExcel.Cells(1, 6).Value     = "COVERED PERIOD"
objExcel.Cells(1, 6).Font.Bold = True
objExcel.Cells(1, 7).Value     = "DAYS REMAINING"
objExcel.Cells(1, 7).Font.Bold = True
objExcel.Cells(1, 8).Value     = "DIFF NOTICE SENT"
objExcel.Cells(1, 8).Font.Bold = True
objExcel.Cells(1, 9).Value     = "DATE DIFF NOTICE SENT"
objExcel.Cells(1, 9).Font.Bold = True

'This bit freezes the top row of the Excel sheet for better useability when there is a lot of information
ObjExcel.ActiveSheet.Range("A2").Select
objExcel.ActiveWindow.FreezePanes = True

'Sets variable for all of the Excel stuff
excel_row = 2

'This for...next contains each worker indicated above
For each x_number in x_number_array

	'Trims the x_number so that we don't have glitches
	x_number = trim(x_number)
	x_number = UCase(x_number)

	back_to_SELF
	CALL navigate_to_MAXIS_screen("REPT", "IEVC")	'Navigates to the worker based report'
	EMReadScreen non_disclosure_screen, 14, 2, 46	'Checks to make sure the NDA is current
	If non_disclosure_screen = "Non-disclosure" Then script_end_procedure ("It appears you need to confirm agreement to access IEVC. Please navigate there manually to confirm and then run the script again.")
	EMWriteScreen x_number, 4, 11					'goes to the specific worker's report
	transmit

	EMReadScreen unresolved_ievs_exists, 1, 8, 5	'Checks to see if there is something listed on the first line
	If unresolved_ievs_exists <> " " Then 			'If so, the script will gather data
		EMReadScreen supervisor_id, 7, 4, 32		'Pulls the X-Number of the supervisor
		EMReadScreen IEVC_check, 4, 2, 53			'Makes sure still on the IEVC report - sometimes this glitches and causes all kinds of errors
		If IEVC_check = "IEVC" Then
			EMSendKey "HOME"						'Sets the cursor at the first editable field - which is the worker X-Number
			PF1										'PF1 on the X-Number to pull up worker information
			EMReadScreen worker_name, 21, 19, 10	'Reads the worker name
			worker_name = trim(worker_name)			'Trims the worker name
			transmit 								'Closes the worker information pop-up
		End If
		EMWriteScreen x_number, 4, 11				'goes to the specific worker's report - again
		transmit									'This part happens again after looking at worker information due to some weird glitchy thing on this report
		IEVC_Row = 8
		DO
			'Reading and trimming the MAXIS case number and dumping it in Excel
			EMReadScreen maxis_case_number, 8, IEVC_Row, 31
			If maxis_case_number = "        " then exit Do 		'Once the script reaches the last line in the list, it will go to the next worker
			maxis_case_number = trim(maxis_case_number)
			objExcel.Cells(excel_row, 4).Value = maxis_case_number	'Adds case number to Excel

			objExcel.Cells(excel_row, 2).Value = x_number		'enters the worker number to the excel spreadsheet
			objExcel.Cells(excel_row, 1).Value = supervisor_id	'Adds Supervisor X-Numner to Excel
			objExcel.Cells(excel_row, 3).Value = worker_name	'Adds the worker name to Excel

			EMReadScreen client_name, 17, IEVC_Row, 14			'Reads the client name and adds to excel
			client_name = trim(client_name)
			objExcel.Cells(excel_row, 5).Value = client_name

			EMReadScreen covered_period, 11, IEVC_Row, 62		'Reads the dates of the match and adds to excel
			covered_period = trim(covered_period)
			objExcel.Cells(excel_row, 6).Value = covered_period

			EMReadScreen days_remaining, 6, IEVC_Row, 74		'Reads how the days left to resolve the match and adds to excel
			days_remaining = trim(days_remaining)
			objExcel.Cells(excel_row, 7).Value = days_remaining
			objExcel.Cells(excel_row, 7).NumberFormat = "0"
			If left(days_remaining, 1) = "(" Then 				'If this is a negative number - listed in () on the panel
				objExcel.Cells(excel_row, 10).Value = "OVERDUE!"		'Adds this to the spreadsheet
				objExcel.Cells(excel_row, 10).Font.Bold = True 		'Highlights the overdue word
				For col = 1 to 10
					objExcel.Cells(excel_row, col).Interior.ColorIndex = 3	'Fills the row with red
				Next
			End If

			EMWriteScreen "D", IEVC_Row, 3		'Opens the detail on the match
			transmit
			row = 1
			col = 1
			EMSearch "SEND IEVS DIFFERENCE NOTICE?", row, col 	'Finds where the difference notice code is - because it moves
			EMReadScreen diff_notc_sent, 1, row, 36				'Reads if diff notice was sent or not
			If diff_notc_sent = "Y" Then EMReadScreen diff_notc_date, 8, row, 72	'If notice was sent, reads the date it was sent
			objExcel.Cells(excel_row, 8).Value = diff_notc_sent	'Adding both of these to excel
			objExcel.Cells(excel_row, 9).Value = diff_notc_date

			PF3 		'Back to the list!

			IEVC_Row = IEVC_Row + 1 'increment to the next row on the panel

			If IEVC_Row = 18 Then 		'If we have reached the end of the page, it will go to the next page
				PF8
				IEVC_Row = 8			'Resets the row
				EMReadScreen last_page_check, 21, 24, 2
			End If
			excel_row = excel_row + 1	'increments the excel row so we don't overwrite our data
			STATS_counter = STATS_counter + 1		'Counts 1 item for every Match found and entered into excel.			diff_notc_date = ""			'blanks this out so that the information is not carried over in the do-loop'
			maxis_case_number = ""
		LOOP until last_page_check = "THIS IS THE LAST PAGE"
	Else
		objExcel.Cells(excel_row, 2).Value = x_number						'If there are no items on a worker's report, this adds that information to the excel sheet
		EMReadScreen supervisor_id, 7, 4, 32				'Finding Supervisor X Number and adding to Excel
		objExcel.Cells(excel_row, 1).Value = supervisor_id
		EMReadScreen IEVC_check, 4, 2, 53
		If IEVC_check = "IEVC" Then 						'Getting the worker name from the info pop-up and adding to Excel
			EMSendKey "HOME"
			PF1
			EMReadScreen worker_name, 21, 19, 10
			worker_name = trim(worker_name)
			objExcel.Cells(excel_row, 3).Value = worker_name
			transmit
		End If
		objExcel.Cells(excel_row, 5).Value = "No IEVS for this worker."		'Adds line to Excel sheet indicating no matches
		excel_row = excel_row + 1
	End If
Next

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


'Formatting the column width.
FOR i = 1 to 12
	objExcel.Columns(i).AutoFit()
NEXT

'Going to another sheet, to enter worker-specific statistics
ObjExcel.Worksheets.Add().Name = "IEVC stats by worker"

'Headers
ObjExcel.Cells(1, 2).Value = "IEVC STATS BY WORKER"
ObjExcel.Cells(1, 2).Font.Bold = TRUE
ObjExcel.Cells(2, 1).Value = "WORKER"
objExcel.Cells(2, 1).Font.Bold = TRUE
ObjExcel.Cells(2, 2).Value = "NAME"
ObjExcel.Cells(2, 2).Font.Bold = TRUE
ObjExcel.Cells(2, 3).Value = "OLDER THAN 45 DAYS"
objExcel.Cells(2, 3).Font.Bold = TRUE
ObjExcel.Cells(2, 4).Value = "UNRESOLVED"
objExcel.Cells(2, 4).Font.Bold = TRUE
ObjExcel.Cells(2, 5).Value = "% OF WORKERS IEVS OLDER THAN 45 DAYS"
objExcel.Cells(2, 5).Font.Bold = TRUE
ObjExcel.Cells(2, 6).Value = "% OF UNRESOLVED IEVS OWNED BY THIS WORKER"
objExcel.Cells(2, 6).Font.Bold = TRUE

'This bit freezes the top 2 rows for scrolling ease of use
ObjExcel.ActiveSheet.Range("A3").Select
objExcel.ActiveWindow.FreezePanes = True

worker_row = 3
'Writes each worker from the worker_array in the Excel spreadsheet
For each x_number in x_number_array
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
	'Writing a formula to excel - Count each row in which Column H on the first worksheet is not blank AND the x number in Column B on the first worksheet matches the X number on this row AND Column G is 0 or less - All OVERDUE matches for this worker
	ObjExcel.Cells(worker_row, 3).Value = "=COUNTIFS('Case information'!H:H, " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'Case information'!B:B, A" & worker_row & ", 'Case information'!G:G, " & Chr(34) & "<=0" & Chr(34) & ")"
	'Writing a formula to excel - Count each row in which Column H on the first worksheet is not blank AND the x number in Column B on the first worksheet matches the X number on this row - ALL matches for this worker
	ObjExcel.Cells(worker_row, 4).Value = "=COUNTIFS('Case information'!H:H, " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'Case information'!B:B, A" & worker_row & ")"
	IF ObjExcel.Cells(worker_row, 4).Value <> "0" Then	'Preventing a divide by 0 error
		ObjExcel.Cells(worker_row, 5).Value = "=C" & worker_row & "/D" & worker_row
	Else
		ObjExcel.Cells(worker_row, 5).Value = "0"
	End If
	ObjExcel.Cells(worker_row, 5).NumberFormat = "0.00%"		'Formula should be percent
	ObjExcel.Cells(worker_row, 6).Value = "=D" & worker_row & "/SUM(D:D)"
	ObjExcel.Cells(worker_row, 6).NumberFormat = "0.00%"		'Formula should be percent
	worker_row = worker_row + 1
Next

'Merging header cell.
ObjExcel.Range(ObjExcel.Cells(1, 1), ObjExcel.Cells(1, 6)).Merge

'Centering the cell
objExcel.Cells(1, 2).HorizontalAlignment = -4108

'Autofitting columns
For col_to_autofit = 1 to 20
	ObjExcel.columns(col_to_autofit).AutoFit()
Next

STATS_counter = STATS_counter - 1		'removing the initial counter so that this number is correct.

script_end_procedure("Success! The spreadsheet has all requested information.")
