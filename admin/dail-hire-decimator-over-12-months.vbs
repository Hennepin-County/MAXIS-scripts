'STATS GATHERING=============================================================================================================
name_of_script = "BULK - DAIL HIRE DECIMATOR OVER 12 MONTHS.vbs"       'REPLACE TYPE with either ACTIONS, BULK, DAIL, NAV, NOTES, NOTICES, or UTILITIES. The name of the script should be all caps. The ".vbs" should be all lower case.
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 1            'manual run time in seconds  -----REPLACE STATS_MANUALTIME = 1 with the actual manual time based on time study
STATS_denomination = "C"        'C is for each case; I is for Instance, M is for member; REPLACE with the denomination applicable to your script.
'END OF stats block==========================================================================================================

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

'CHANGELOG BLOCK===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
CALL changelog_update("04/17/24", "Initial version.", "Mark Riegel, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

Function create_array_of_all_active_x_numbers_in_county_with_restart(array_name, two_digit_county_code, restart_status, restart_worker_number)
'--- This function is used to grab all active X numbers in a county
'~~~~~ array_name: name of array that will contain all the x numbers
'~~~~~ county_code: inserted by reading the county code under REPT/USER
'===== Keywords: MAXIS, array, worker number, create
	'Getting to REPT/USER
	Call navigate_to_MAXIS_screen("REPT", "USER")
	PF5 'Hitting PF5 to force sorting, which allows directly selecting a county
	Call write_value_and_transmit(county_code, 21, 6)  	'Inserting county

	MAXIS_row = 7  'Declaring the MAXIS row
	array_name = ""    'Blanking out array_name in case this has been used already in the script

    Found_restart_worker = False    'defaulting to false. Will become true when the X number is found.
	Do
		Do
			'Reading MAXIS information for this row, adding to spreadsheet
			EMReadScreen worker_ID, 8, MAXIS_row, 5					'worker ID
			If worker_ID = "        " then exit do					'exiting before writing to array, in the event this is a blank (end of list)
            If restart_status = True then
                If trim(UCase(worker_ID)) = trim(UCase(restart_worker_number)) then
                    Found_restart_worker = True
                End if
                If Found_restart_worker = True then array_name = trim(array_name & " " & worker_ID)				'writing to variable
            Else
                array_name = trim(array_name & " " & worker_ID)				'writing to variable
            End if
			MAXIS_row = MAXIS_row + 1
		Loop until MAXIS_row = 19

		'Seeing if there are more pages. If so it'll grab from the next page and loop around, doing so until there's no more pages.
		EMReadScreen more_pages_check, 7, 19, 3
		If more_pages_check = "More: +" then
			PF8			'getting to next screen
			MAXIS_row = 7	're-declaring MAXIS row so as to start reading from the top of the list again
		End if
	Loop until more_pages_check = "More:  " or more_pages_check = "       "	'The or works because for one-page only counties, this will be blank

    array_name = split(array_name)
End function

'THE SCRIPT==================================================================================================================
EMConnect ""
Call Check_for_MAXIS(False)
dail_to_decimate = "ALL"
all_workers_check = 1

this_month = CM_mo & " " & CM_yr
this_month_date = CM_mo & "/01/" & CM_yr
this_month_date = DateAdd("m", 1, this_month_date)

'Finding the right folder to automatically save the file
month_folder = "DAIL " & CM_mo & "-" & DatePart("yyyy", date) & ""
decimator_folder = replace(this_month, " ", "-") & " DAIL Decimator"
report_date = replace(date, "/", "-")

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 266, 110, "DAIL Decimator - HIRE Messages over 12 Months Old"
  GroupBox 10, 5, 250, 40, "Using the DAIL Decimator script"
  Text 20, 20, 235, 20, "This script should be used to remove HIRE messages that are over 12 months old from current month."
  CheckBox 10, 50, 165, 10, "Check here to process for all workers (default).", all_workers_check
  Text 10, 65, 170, 10, "For restart only, enter the x number to restart from:"
  EditBox 180, 60, 50, 15, restart_worker_number
  ButtonGroup ButtonPressed
    OkButton 155, 90, 50, 15
    CancelButton 210, 90, 50, 15
EndDialog

Do
	Do
  		err_msg = ""
  		dialog Dialog1
  		cancel_without_confirmation
  		If trim(restart_worker_number) = "" and all_workers_check = 0 then err_msg = err_msg & vbNewLine & "* Select a worker number(s) or all cases."
  		If trim(restart_worker_number) <> "" and all_workers_check = 1 then err_msg = err_msg & vbNewLine & "* Select a worker number(s) or all cases, not both options."
  	  	IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
  	LOOP until err_msg = ""
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

back_to_SELF 'navigates back to self in case the worker is working within the DAIL. All messages for a single number may not be captured otherwise.

'determining if this is a restart or not in function below when gathering the x numbers.
If trim(restart_worker_number) = "" then
    restart_status = False
Else
	restart_status = True
End if

'If all workers are selected, the script will go to REPT/USER, and load all of the workers into an array. Otherwise it'll create a single-object "array" just for simplicity of code.
Call create_array_of_all_active_x_numbers_in_county_with_restart(worker_array, two_digit_county_code, restart_status, restart_worker_number)

'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

'Changes name of Excel sheet to "DAIL List"
ObjExcel.ActiveSheet.Name = "Deleted DAILS"

'Excel headers and formatting the columns
objExcel.Cells(1, 1).Value = "X NUMBER"
objExcel.Cells(1, 2).Value = "CASE #"
objExcel.Cells(1, 3).Value = "DAIL TYPE"
objExcel.Cells(1, 4).Value = "DAIL MO."
objExcel.Cells(1, 5).Value = "DAIL MESSAGE"
objExcel.Cells(1, 6).Value = "ACTION TAKEN"


FOR i = 1 to 6		'formatting the cells'
	objExcel.Cells(1, i).Font.Bold = True		'bold font'
	ObjExcel.columns(i).NumberFormat = "@" 		'formatting as text
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

DIM DAIL_array()
ReDim DAIL_array(5, 0)
Dail_count = 0              'Incremental for the array
false_count = 0

'constants for array
const worker_const	            = 0
const maxis_case_number_const   = 1
const dail_type_const 	        = 2
const dail_month_const		    = 3
const dail_msg_const		    = 4
const dail_action_const		    = 5

'Sets variable for all of the Excel stuff
excel_row = 2
deleted_dails = 0	'establishing the value of the count for deleted deleted_dails

'This for...next contains each worker indicated above
For each worker in worker_array
    MAXIS_case_number = ""
    back_to_SELF

    'Navigate to DAIL/PICK and select 'INFO' to find HIRE messages
    CALL navigate_to_MAXIS_screen("DAIL", "PICK")
    EMWriteScreen "_", 7, 39
    EMWriteScreen "X", 13, 39
    transmit
    
	Call write_value_and_transmit(worker, 21, 6)
	transmit  'transmits past 'not your dail message
    EMReadScreen number_of_dails, 1, 3, 67		'Reads where the count of DAILs is listed

	DO
		If number_of_dails = " " Then exit do		'if this space is blank the rest of the DAIL reading is skipped
		dail_row = 6			'Because the script brings each new case to the top of the page, dail_row starts at 6.
		DO
			dail_type = ""
			dail_msg = ""
            dail_month = ""
            MAXIS_case_number = ""

		    'Determining if there is a new case number...
		    EMReadScreen new_case, 8, dail_row, 63
		    new_case = trim(new_case)
		    IF new_case = "CASE NBR" THEN
			    '...if the script does find that there is a new case number (indicated by "CASE NBR"), it will write a "T" in the next row and transmit, bringing that case number to the top of your DAIL
			    Call write_value_and_transmit("T", dail_row + 1, 3)
			ELSEIF new_case <> "CASE NBR" THEN '...if there is NOT a new case number, the script will read the DAIL type, month, year, and message...
				Call write_value_and_transmit("T", dail_row, 3)
			End if

            dail_row = 6  'resetting the DAIL row

            'Reading the DAIL Information
			EMReadScreen MAXIS_case_number, 8, dail_row - 1, 73
            MAXIS_case_number = trim(MAXIS_case_number)

            EMReadScreen dail_type, 4, dail_row, 6

            EMReadScreen dail_msg, 61, dail_row, 20
			dail_msg = trim(dail_msg)

            EMReadScreen dail_month, 8, dail_row, 11
            dail_month = trim(dail_month)

            If dail_type = "HIRE" Then
                'process message
                dail_month_date = replace(dail_month, " ", "/01/20")
                dail_month_date = dateadd("m", 1, dail_month_date)
                dail_months_old = DateDiff("m", dail_month_date, this_month_date)

                If dail_months_old > 13 Then
                    'should be deleted
			    	objExcel.Cells(excel_row, 1).Value = worker
			    	objExcel.Cells(excel_row, 2).Value = MAXIS_case_number
			    	objExcel.Cells(excel_row, 3).Value = dail_type
			    	objExcel.Cells(excel_row, 4).Value = dail_month
			    	objExcel.Cells(excel_row, 5).Value = dail_msg
			    	
                    msgbox "Message would be deleted"
                    ' Call write_value_and_transmit("D", dail_row, 3)
			        EMReadScreen other_worker_error, 13, 24, 2
			        If other_worker_error = "** WARNING **" then transmit
			    	objExcel.Cells(excel_row, 6).Value = "Message deleted."

                    excel_row = excel_row + 1
                    deleted_dails = deleted_dails + 1
                    stats_counter = stats_counter + 1   'I increment thee

                Else
                    dail_row = dail_row + 1
                End If

            Else
                'Not a HIRE message so can be skipped
                dail_row = dail_row + 1
            End If

            'checking for the last DAIL message - If it's the last message, which can be blank OR _ then the script will exit the do. 
			EMReadScreen next_dail_check, 7, dail_row, 3
			If trim(next_dail_check) = "" or trim(next_dail_check) = "_" then
                PF8
                EMReadScreen next_dail_check, 7, dail_row, 3
			    If trim(next_dail_check) = "" or trim(next_dail_check) = "_" then
                    last_case = true
				    exit do
                End if 
			End if
		LOOP
		IF last_case = true THEN exit do
	LOOP
Next

STATS_counter = STATS_counter - 1
'Enters info about runtime for the benefit of folks using the script
objExcel.Cells(2, 8).Value = "Number of DAILs deleted:"
objExcel.Cells(3, 8).Value = "Average time to find/select/copy/paste one line (in seconds):"
objExcel.Cells(4, 8).Value = "Estimated manual processing time (lines x average):"
objExcel.Cells(5, 8).Value = "Script run time (in seconds):"
objExcel.Cells(6, 8).Value = "Estimated time savings by using script (in minutes):"
objExcel.Cells(7, 8).Value = "Number of messages reviewed/DAIL messages remaining:"
objExcel.Cells(8, 8).Value = "False count/duplicate DAIL Messages not counted:"
objExcel.Columns(8).Font.Bold = true
objExcel.Cells(2, 9).Value = deleted_dails
objExcel.Cells(3, 9).Value = STATS_manualtime
objExcel.Cells(4, 9).Value = STATS_counter * STATS_manualtime
objExcel.Cells(5, 9).Value = timer - start_time
objExcel.Cells(6, 9).Value = ((STATS_counter * STATS_manualtime) - (timer - start_time)) / 60
objExcel.Cells(7, 9).Value = STATS_counter
objExcel.Cells(8, 9).Value = false_count

'Formatting the column width.
FOR i = 1 to 9
	objExcel.Columns(i).AutoFit()
NEXT

'saving the Excel file
file_info = month_folder & "\" & decimator_folder & "\" & report_date & " " & "HIRE Messages over 12 months old" & " " & deleted_dails

'Saves and closes the most recent Excel workbook with the Task based cases to process.
objExcel.ActiveWorkbook.SaveAs "T:\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\DAIL list\" & file_info & ".xlsx"
objExcel.ActiveWorkbook.Close
objExcel.Application.Quit
objExcel.Quit

'End the script. Put any success messages in between the quotes, *always* starting with the word "Success!"
script_end_procedure("")

'----------------------------------------------------------------------------------------------------Closing Project Documentation - Version date 01/12/2023
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------
'--Tab orders reviewed & confirmed----------------------------------------------
'--Mandatory fields all present & Reviewed--------------------------------------
'--All variables in dialog match mandatory fields-------------------------------
'Review dialog names for content and content fit in dialog----------------------
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------
'--CASE:NOTE Header doesn't look funky------------------------------------------
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------
'--write_variable_in_CASE_NOTE function: confirm that proper punctuation is used -----------------------------------
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------
'--MAXIS_background_check reviewed (if applicable)------------------------------
'--PRIV Case handling reviewed -------------------------------------------------
'--Out-of-County handling reviewed----------------------------------------------
'--script_end_procedures (w/ or w/o error messaging)----------------------------
'--BULK - review output of statistics and run time/count (if applicable)--------
'--All strings for MAXIS entry are uppercase vs. lower case (Ex: "X")-----------
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------
'--Incrementors reviewed (if necessary)-----------------------------------------
'--Denomination reviewed -------------------------------------------------------
'--Script name reviewed---------------------------------------------------------
'--BULK - remove 1 incrementor at end of script reviewed------------------------

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------
'--comment Code-----------------------------------------------------------------
'--Update Changelog for release/update------------------------------------------
'--Remove testing message boxes-------------------------------------------------
'--Remove testing code/unnecessary code-----------------------------------------
'--Review/update SharePoint instructions----------------------------------------
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------
'--COMPLETE LIST OF SCRIPTS update policy references----------------------------
'--Complete misc. documentation (if applicable)---------------------------------
'--Update project team/issue contact (if applicable)----------------------------

