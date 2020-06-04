'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ADMIN - DAIL TASK-BASED ASSIGNMENTS.vbs"
start_time = timer
STATS_counter = 1                       'sets the stats counter at one
STATS_manualtime = 20
STATS_denomination = "C"       			'C is for each CASE
'END OF stats block==============================================================================================

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/Hennepin-County/MAXIS-scripts/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else											'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/Hennepin-County/MAXIS-scripts/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
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
		FuncLib_URL = "C:\MAXIS-scripts\MASTER FUNCTIONS LIBRARY.vbs"
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
call changelog_update("12/17/2019", "Added function to evaluate DAIL messages.", "Ilse Ferris, Hennepin County")
call changelog_update("07/31/2019", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display

Function task_dail_selection
	'selecting the type of DAIl message
	EMWriteScreen "x", 4, 12		'transmits to the PICK screen
	transmit
	EMWriteScreen "_", 7, 39		'clears the all selection
    'Selecting the identified categories
    EMWriteScreen "X",  8, 39 'COLA
    EMWriteScreen "X", 10, 39 'CSES
    EMWriteScreen "X", 13, 39 'INFO
    EMWriteScreen "X", 18, 39 'PEPR
    EMWriteScreen "X", 19, 39 'TIKL
    Transmit
End Function

'END CHANGELOG BLOCK =======================================================================================================

'----------------------------------------------------------------------------------------------------THE SCRIPT
EMConnect ""

dail_to_decimate = "Task-Based"
this_month = CM_mo & " " & CM_yr
next_month = CM_plus_1_mo & " " & CM_plus_1_yr
CM_minus_2_mo =  right("0" & DatePart("m", DateAdd("m", -2, date)), 2)

'Finding the right folder to automatically save the file for the main DAIL Decimator
month_folder = "DAIL " & CM_mo & "-" & DatePart("yyyy", date) & ""
decimator_folder = replace(this_month, " ", "-") & " DAIL Decimator"
report_date = replace(date, "/", "-")

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 266, 110, "DAIL Task-Based Processing"
  ButtonGroup ButtonPressed
    OkButton 155, 90, 50, 15
    CancelButton 210, 90, 50, 15
  GroupBox 5, 5, 255, 80, "How the script works:"
  Text 20, 20, 210, 25, "This script will queue specific family baskets identified for DAIL task-based processing. The script will use the DAIL DECIMATOR functionality to evaluate DAIL and delete non-actionable DAILS."
  Text 15, 55, 235, 30, "The deleted and remaining DAIL messages are then automatically saved in the QI folders. Lastly, the remaining DAIL messages are saved in a designated file for task-based assignments."
EndDialog
Do
	Do
  		dialog Dialog1
        cancel_without_confirmation
    Loop until ButtonPressed = -1
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

worker_number = "X127FD4, X127FD5, X127EZ5, X127FD8, X127EZ8, X127FH6, X127FD6, X127EZ6, X127EZ7, X127FD7, X127EZ0, X127ES4, X127ET2, X127ES8, X127ET1, X127ES7, X127EM5, X127EM6, X127EZ2, X127EZ9, X127ES5, X127EX2, X127ES6, X127EZ4, X127EZ3, X127FF3, X127EU5, X127EX7, X127EU6, X127FJ5, X127EY2, X127F3W, X127FA1, X127EU8, X127F3Q, X127EX9, X127F3X, X127FA2, X127EU7, X127F3R, X127EX8, X127F3Z, X127EV1, X127FB9, X127FC1, X127EV2, X127EV4, X127EV3, X127FB8, X127ER8, X127ET4, X127F3B, X127ET6, X127ES1, X127FB6, X127F3A, X127F4C, X127F4F, X127ET7, X127FB3, X127ER9, X127ET5, X127ES2, X127ET9, X127EW3, X127EU1, X127EU2"

back_to_SELF 'navigates back to self in case the worker is working within the DAIL. All messages for a single number may not be captured otherwise.

'If all workers are selected, the script will go to REPT/USER, and load all of the workers into an array. Otherwise it'll create a single-object "array" just for simplicity of code.
'If all_workers_check = checked then
'	call create_array_of_all_active_x_numbers_in_county(worker_array, two_digit_county_code)
'Else
	x1s_from_dialog = split(worker_number, ", ")	'Splits the worker array based on commas

	'Need to add the worker_county_code to each one
	For each x1_number in x1s_from_dialog
		If worker_array = "" then
			worker_array = trim(ucase(x1_number))		'replaces worker_county_code if found in the typed x1 number
		Else
			worker_array = worker_array & "," & trim(ucase(x1_number)) 'replaces worker_county_code if found in the typed x1 number
		End if
	Next
	'Split worker_array
	worker_array = split(worker_array, ",")
'End if

'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

'Changes name of Excel sheet to "DAIL List"
ObjExcel.ActiveSheet.Name = "Deleted DAILS - " & dail_to_decimate

'Excel headers and formatting the columns
objExcel.Cells(1, 1).Value = "X NUMBER"
objExcel.Cells(1, 2).Value = "CASE #"
objExcel.Cells(1, 3).Value = "DAIL TYPE"
objExcel.Cells(1, 4).Value = "DAIL MO."
objExcel.Cells(1, 5).Value = "DAIL MESSAGE"

FOR i = 1 to 5		'formatting the cells'
	objExcel.Cells(1, i).Font.Bold = True		'bold font'
	ObjExcel.columns(i).NumberFormat = "@" 		'formatting as text
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

DIM DAIL_array()
ReDim DAIL_array(5, 0)
Dail_count = 0              'Incremental for the array

'constants for array
const worker_const	            = 0
const maxis_case_number_const   = 1
const dail_type_const 	        = 2
const dail_month_const		    = 3
const dail_msg_const		    = 4
const client_name_const         = 5

'Sets variable for all of the Excel stuff
excel_row = 2
deleted_dails = 0	'establishing the value of the count for deleted deleted_dails

MAXIS_case_number = ""
CALL navigate_to_MAXIS_screen("DAIL", "DAIL")

'This for...next contains each worker indicated above
For each worker in worker_array
	'msgbox worker
	DO
		EMReadScreen dail_check, 4, 2, 48
		If next_dail_check <> "DAIL" then
			MAXIS_case_number = ""
			CALL navigate_to_MAXIS_screen("DAIL", "DAIL")
		End if
	Loop until dail_check = "DAIL"

	EMWriteScreen worker, 21, 6
	transmit
	transmit 'transmit past 'not your dail message'

	Call task_dail_selection

	EMReadScreen number_of_dails, 1, 3, 67		'Reads where the count of DAILs is listed

	DO
		If number_of_dails = " " Then exit do		'if this space is blank the rest of the DAIL reading is skipped

		dail_row = 6			'Because the script brings each new case to the top of the page, dail_row starts at 6.
		DO
			dail_type = ""
			dail_msg = ""

		    'Determining if there is a new case number...
		    EMReadScreen new_case, 8, dail_row, 63
		    new_case = trim(new_case)
		    IF new_case <> "CASE NBR" THEN '...if there is NOT a new case number, the script will read the DAIL type, month, year, and message...
				Call write_value_and_transmit("T", dail_row, 3)
				dail_row = 6
			ELSEIF new_case = "CASE NBR" THEN
			    '...if the script does find that there is a new case number (indicated by "CASE NBR"), it will write a "T" in the next row and transmit, bringing that case number to the top of your DAIL
			    Call write_value_and_transmit("T", dail_row + 1, 3)
				dail_row = 6
			End if

			EMReadScreen maxis_case_number, 8, dail_row - 1, 73
            EMReadScreen dail_type, 4, dail_row, 6
			EMReadScreen dail_msg, 61, dail_row, 20
			dail_msg = trim(dail_msg)
            EMReadScreen dail_month, 8, dail_row, 11
            dail_month = trim(dail_month)

            'This bit of code grabs the client name. The do/loop expands the search area until the value for
    		'next_two equals "--" ... at which time the script determines that the cl name has ended
    		dail_col = 6
    		name_len = 1
    		DO
    			EMReadScreen client_name, name_len, 5, 5
    			EMReadScreen next_two, 2, 5, dail_col
    			IF next_two <> "--" THEN
    				name_len = name_len + 1
    				dail_col = dail_col + 1
    			END IF
    		LOOP UNTIL next_two = "--"

			stats_counter = stats_counter + 1
            Call non_actionable_dails   'Function to evaluate the DAIL messages

            IF add_to_excel = True then
				'--------------------------------------------------------------------...and put that in Excel.
				objExcel.Cells(excel_row, 1).Value = worker
				objExcel.Cells(excel_row, 2).Value = trim(MAXIS_case_number)
				objExcel.Cells(excel_row, 3).Value = trim(dail_type)
				objExcel.Cells(excel_row, 4).Value = trim(dail_month)
				objExcel.Cells(excel_row, 5).Value = trim(dail_msg)
				excel_row = excel_row + 1

				Call write_value_and_transmit("D", dail_row, 3)
				EMReadScreen other_worker_error, 13, 24, 2
				If other_worker_error = "** WARNING **" then transmit
				deleted_dails = deleted_dails + 1
			else
				add_to_excel = False
				dail_row = dail_row + 1
                ReDim Preserve DAIL_array(5, DAIL_count)	'This resizes the array based on the number of rows in the Excel File'
            	DAIL_array(worker_const,	           DAIL_count) = worker
            	DAIL_array(maxis_case_number_const,    DAIL_count) = MAXIS_case_number
            	DAIL_array(dail_type_const, 	       DAIL_count) = dail_type
                If len(dail_month) = 5 then dail_month = replace(dail_month, " ", "/1/")
            	DAIL_array(dail_month_const, 		   DAIL_count) = dail_month
            	DAIL_array(dail_msg_const, 		       DAIL_count) = dail_msg
                DAIL_array(client_name_const, 		   DAIL_count) = client_name
                Dail_count = DAIL_count + 1
			End if

			EMReadScreen message_error, 11, 24, 2		'Cases can also NAT out for whatever reason if the no messages instruction comes up.
			If message_error = "NO MESSAGES" then
				CALL navigate_to_MAXIS_screen("DAIL", "DAIL")
				Call write_value_and_transmit(worker, 21, 6)
				transmit   'transmit past 'not your dail message'
				Call task_dail_selection
				exit do
			End if

			'...going to the next page if necessary
			EMReadScreen next_dail_check, 4, dail_row, 4
			If trim(next_dail_check) = "" then
				PF8
				EMReadScreen last_page_check, 21, 24, 2
				If last_page_check = "THIS IS THE LAST PAGE" then
					all_done = true
					exit do
				Else
					dail_row = 6
				End if
			End if
		LOOP
		IF all_done = true THEN exit do
	LOOP
Next

STATS_counter = STATS_counter - 1
'Enters info about runtime for the benefit of folks using the script
objExcel.Cells(2, 7).Value = "Number of DAILs deleted:"
objExcel.Cells(3, 7).Value = "Average time to find/select/copy/paste one line (in seconds):"
objExcel.Cells(4, 7).Value = "Estimated manual processing time (lines x average):"
objExcel.Cells(5, 7).Value = "Script run time (in seconds):"
objExcel.Cells(6, 7).Value = "Estimated time savings by using script (in minutes):"
objExcel.Cells(7, 7).Value = "Number of messages reviewed/DAIL messages remaining:"
objExcel.Columns(7).Font.Bold = true
objExcel.Cells(2, 8).Value = deleted_dails
objExcel.Cells(3, 8).Value = STATS_manualtime
objExcel.Cells(4, 8).Value = STATS_counter * STATS_manualtime
objExcel.Cells(5, 8).Value = timer - start_time
objExcel.Cells(6, 8).Value = ((STATS_counter * STATS_manualtime) - (timer - start_time)) / 60
objExcel.Cells(7, 8).Value = STATS_counter

'Formatting the column width.
FOR i = 1 to 8
	objExcel.Columns(i).AutoFit()
NEXT

'Adding another sheet
ObjExcel.Worksheets.Add().Name = "Remaining DAIL messages"

excel_row = 2
'Excel headers and formatting the columns
objExcel.Cells(1, 1).Value = "X NUMBER"
objExcel.Cells(1, 2).Value = "CASE #"
objExcel.Cells(1, 3).Value = "DAIL TYPE"
objExcel.Cells(1, 4).Value = "DAIL MO."
objExcel.Cells(1, 5).Value = "DAIL MESSAGE"

FOR i = 1 to 5		'formatting the cells'
	objExcel.Cells(1, i).Font.Bold = True		'bold font'
	ObjExcel.columns(i).NumberFormat = "@" 		'formatting as text
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

'Export informaiton to Excel re: case status
For item = 0 to UBound(DAIL_array, 2)
	objExcel.Cells(excel_row, 1).Value = DAIL_array(worker_const, item)
	objExcel.Cells(excel_row, 2).Value = DAIL_array(maxis_case_number_const, item)
    objExcel.Cells(excel_row, 3).Value = DAIL_array(dail_type_const, item)
	objExcel.Cells(excel_row, 4).Value = DAIL_array(dail_month_const, item)
    objExcel.Cells(excel_row, 5).Value = DAIL_array(dail_msg_const, item)
	excel_row = excel_row + 1
Next

objExcel.Cells(1, 7).Value = "Remaning DAIL messages:"
objExcel.Columns(7).Font.Bold = true
objExcel.Cells(1, 8).Value = DAIL_count

'formatting the cells
FOR i = 1 to 8
	objExcel.Columns(i).AutoFit()				'sizing the columns
NEXT

'saving the Excel file
file_info = month_folder & "\" & decimator_folder & "\" & report_date & " " & dail_to_decimate & " " & deleted_dails

'Saves and closes the most recent Excel workbook with the Task based cases to process.
objExcel.ActiveWorkbook.SaveAs "T:\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\DAIL list\" & file_info & ".xlsx"
objExcel.ActiveWorkbook.Close
objExcel.Application.Quit
objExcel.Quit

'----------------------------------------------------------------------------------------------------Task-based assignment output
'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

'Changes name of Excel sheet to "DAIL List"
ObjExcel.ActiveSheet.Name = "FAD DAIL Assignments"

excel_row = 2
'Excel headers and formatting the columns
objExcel.Cells(1, 1).Value = "X NUMBER"
objExcel.Cells(1, 2).Value = "CASE #"
objExcel.Cells(1, 3).Value = "DAIL TYPE"
objExcel.Cells(1, 4).Value = "DAIL MO."
objExcel.Cells(1, 5).Value = "DAIL MESSAGE"
objExcel.Cells(1, 6).Value = "CLIENT NAME"

FOR i = 1 to 6		'formatting the cells'
	objExcel.Cells(1, i).Font.Bold = True		'bold font'
	ObjExcel.columns(i).NumberFormat = "@" 		'formatting as text
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

'Export informaiton to Excel re: case status
For item = 0 to UBound(DAIL_array, 2)
	objExcel.Cells(excel_row, 1).Value = DAIL_array(worker_const, item)
	objExcel.Cells(excel_row, 2).Value = DAIL_array(maxis_case_number_const, item)
    objExcel.Cells(excel_row, 3).Value = DAIL_array(dail_type_const, item)
	objExcel.Cells(excel_row, 4).Value = DAIL_array(dail_month_const, item)
    objExcel.Cells(excel_row, 5).Value = DAIL_array(dail_msg_const, item)
    objExcel.Cells(excel_row, 6).Value = DAIL_array(client_name_const, item)
	excel_row = excel_row + 1
Next

'formatting the cells
FOR i = 1 to 6
	objExcel.Columns(i).AutoFit()				'sizing the columns
NEXT

'Saves and closes the most recent Excel workbook with the Task based cases to process.
objExcel.ActiveWorkbook.SaveAs "T:\HSPH Restricted Access Workspace\EWS Work Structure\WS Data and Reports\Daily Form Reports\DAIL\" & report_date &".xlsx"
objExcel.ActiveWorkbook.Close
objExcel.Application.Quit
objExcel.Quit

script_end_procedure("Success! Please review the list created for accuracy.")
