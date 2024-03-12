'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - DAIL DECIMATOR.vbs"
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
call changelog_update("06/05/2023", "Added option for removing SVES/TPQY responses sent for Ex Parte cases.", "Ilse Ferris, Hennepin County")
Call changelog_update("10/10/2022", "Added restart functionality when using all workers option.", "Ilse Ferris, Hennepin County")
call changelog_update("02/01/2021", "Updated boolean variable name for clarity.", "Ilse Ferris, Hennepin County")
call changelog_update("06/10/2020", "Added TIKL DAIL selection.", "Ilse Ferris, Hennepin County")
call changelog_update("12/17/2019", "Added function to evaluate DAIL messages.", "Ilse Ferris, Hennepin County")
call changelog_update("12/09/2019", "Added 01/20 COLA messages to removal list.", "Ilse Ferris, Hennepin County")
call changelog_update("12/02/2019", "Added 07/19 COLA messages to removal list.", "Ilse Ferris, Hennepin County")
call changelog_update("08/07/2019", "Updated to output 8-digit case numbers and 8-character dates.", "Ilse Ferris, Hennepin County")
call changelog_update("08/07/2019", "Added auto-save functionality to save to specified QI folders.", "Ilse Ferris, Hennepin County")
call changelog_update("02/12/2019", "Added COLA messages for 03/19 COLA - SSI and RSDI Updated.", "Ilse Ferris, Hennepin County")
call changelog_update("01/17/2019", "Added total of DAIL messages left after processing.", "Ilse Ferris, Hennepin County")
call changelog_update("12/17/2018", "Added PEPR messages older than CM, and BENDEX and SDX messages for this month only.", "Ilse Ferris, Hennepin County")
call changelog_update("12/15/2018", "Added TIKL's for exempt IR process over 2 months old.", "Ilse Ferris, Hennepin County")
call changelog_update("12/03/2018", "Added COLA messages for 01/19 COLA.", "Ilse Ferris, Hennepin County")
call changelog_update("11/02/2018", "Added additional ELIG messages older than CM.", "Ilse Ferris, Hennepin County")
call changelog_update("10/26/2018", "Added additional messages included TIKL's over 6 months old, STAT edits over 5 days old and EFUNDS messages.", "Ilse Ferris, Hennepin County")
call changelog_update("10/26/2018", "Added MEC2 messages.", "Ilse Ferris, Hennepin County")
call changelog_update("10/24/2018", "Reorganized messages by type and alphabetical. Cleaned up backup coding.", "Ilse Ferris, Hennepin County")
call changelog_update("10/22/2018", "Added support for ADDR INFO messages, STAT edits over 10 days old, temporary addition of COLA messages greater than 07/18 COLA and MSA SBUD/LBUD messages.", "Ilse Ferris, Hennepin County")
call changelog_update("01/02/2018", "Added supported PEPR and CSES messages.", "Ilse Ferris, Hennepin County")
call changelog_update("01/02/2018", "Added Casey Love as autorized user of the script, blanked out MAXIS case number for PRIV cases, and merged SVES and INFO messages together into one option.", "Ilse Ferris, Hennepin County")
call changelog_update("12/30/2017", "Complete updates for INFO, SVES, COLA and ELIG messages.", "Ilse Ferris, Hennepin County")
call changelog_update("12/11/2017", "Added Quality Improvement Team as authorized users of DAIL Decimator script.", "Ilse Ferris, Hennepin County")
call changelog_update("12/05/2017", "Added ELIG DAIL messages as DAILs to decimate!", "Ilse Ferris, Hennepin County")
call changelog_update("10/28/2017", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'local function for DAIL/PICK selection
Function dail_pick_selection()
    CALL navigate_to_MAXIS_screen("DAIL", "PICK")
    EMReadscreen pick_confirmation, 26, 4, 29

    If pick_confirmation = "View/Pick Selection (PICK)" then
        'selecting the type of DAIl message
        If dail_to_decimate = "ALL"   then EMWriteScreen "X", 7, 39
    	If dail_to_decimate = "COLA"  then EMWriteScreen "X", 8, 39
    	If dail_to_decimate = "CSES"  then EMWriteScreen "X", 10, 39
    	If dail_to_decimate = "ELIG"  then EMWriteScreen "X", 11, 39
    	If dail_to_decimate = "INFO"  then EMWriteScreen "X", 13, 39
    	If dail_to_decimate = "PEPR"  then EMWriteScreen "X", 18, 39
    	If dail_to_decimate = "TIKL"  then EMWriteScreen "X", 19, 39
    	transmit
    Else
        script_end_procedure("Unable to navigate to DAIL/PICK. The script will now end.")
    End if
End Function

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

'----------------------------------------------------------------------------------------------------THE SCRIPT
EMConnect ""
Call Check_for_MAXIS(False)
dail_to_decimate = "ALL"
all_workers_check = 1

this_month = CM_mo & " " & CM_yr
next_month = CM_plus_1_mo & " " & CM_plus_1_yr
last_month = CM_minus_1_mo & " " & CM_minus_1_yr
CM_minus_2_mo =  right("0" & DatePart("m", DateAdd("m", -2, date)), 2)

'Finding the right folder to automatically save the file
month_folder = "DAIL " & CM_mo & "-" & DatePart("yyyy", date) & ""
decimator_folder = replace(this_month, " ", "-") & " DAIL Decimator"
report_date = replace(date, "/", "-")

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 266, 150, "DAIL Decimator Dialog"
  GroupBox 10, 5, 250, 40, "Using the DAIL Decimator script"
  Text 20, 20, 235, 20, "This script should be used to remove DAIL messages that have been determined by Quality Improvement staff do not require action."
  Text 40, 55, 35, 10, "DAIL type:"
  DropListBox 80, 50, 60, 15, "Select one..."+chr(9)+"ALL"+chr(9)+"COLA"+chr(9)+"CSES"+chr(9)+"ELIG"+chr(9)+"INFO"+chr(9)+"PEPR"+chr(9)+"TIKL", dail_to_decimate
  Text 15, 75, 60, 10, "Worker number(s):"
  EditBox 80, 70, 180, 15, worker_number
  CheckBox 15, 90, 135, 10, "Check here to process for all workers.", all_workers_check
  CheckBox 15, 100, 145, 10, "Check to remove SVES/QURY messages.", TPQY_checkbox
  Text 25, 115, 170, 10, "If restarting, what x number are you restarting from?"
  EditBox 210, 110, 50, 15, restart_worker_number
  ButtonGroup ButtonPressed
    OkButton 155, 130, 50, 15
    CancelButton 210, 130, 50, 15
EndDialog

Do
	Do
  		err_msg = ""
  		dialog Dialog1
  		cancel_without_confirmation
  		If dail_to_decimate = "Select one..." then err_msg = err_msg & vbNewLine & "* Select the type of DAIL message to decimate!"
  		If trim(worker_number) = "" and all_workers_check = 0 then err_msg = err_msg & vbNewLine & "* Select a worker number(s) or all cases."
  		If trim(worker_number) <> "" and all_workers_check = 1 then err_msg = err_msg & vbNewLine & "* Select a worker number(s) or all cases, not both options."
		 If trim(restart_worker_number) <> "" then
            If all_workers_check = 0 then err_msg = err_msg & vbNewLine & "* The restart option only works with the all workers option. Please update your selections."
            If len(trim(restart_worker_number)) <> 7 then err_msg = err_msg & vbNewLine & "* Enter one 7-digit worker number to restart."
        End if
  	  	IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
  	LOOP until err_msg = ""
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

back_to_SELF 'navigates back to self in case the worker is working within the DAIL. All messages for a single number may not be captured otherwise.

'Ending message when there are no more DAIL's differs based on if you select ALL DAIL's or specific DAILs
If dail_to_decimate = "ALL" then
    dail_end_msg = "NO MESSAGES WORK"
Else
    'all specified selection(s) will get this ending user message.
    dail_end_msg = "NO MESSAGES TYPE"
End if

'determining if this is a restart or not in function below when gathering the x numbers.
If trim(restart_worker_number) = "" then
    restart_status = False
Else
	restart_status = True
End if

'If all workers are selected, the script will go to REPT/USER, and load all of the workers into an array. Otherwise it'll create a single-object "array" just for simplicity of code.
If all_workers_check = checked then
	Call create_array_of_all_active_x_numbers_in_county_with_restart(worker_array, two_digit_county_code, restart_status, restart_worker_number)
Else
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
End if

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
ReDim DAIL_array(4, 0)
Dail_count = 0              'Incremental for the array
all_dail_array = "*"    'setting up string to find duplicate DAIL messages. At times there is a glitch in the DAIL, and messages are reviewed a second time.
false_count = 0

'constants for array
const worker_const	            = 0
const maxis_case_number_const   = 1
const dail_type_const 	        = 2
const dail_month_const		    = 3
const dail_msg_const		    = 4

'Sets variable for all of the Excel stuff
excel_row = 2
deleted_dails = 0	'establishing the value of the count for deleted deleted_dails

'This for...next contains each worker indicated above
For each worker in worker_array
    DO
        EMReadScreen dail_check, 4, 2, 48
        If dail_check <> "DAIL" then Call dail_pick_selection
    Loop until dail_check = "DAIL"

	Call write_value_and_transmit(worker, 21, 6)
	transmit  'transmits past 'not your dail message
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
			ELSEIF new_case = "CASE NBR" THEN
			    '...if the script does find that there is a new case number (indicated by "CASE NBR"), it will write a "T" in the next row and transmit, bringing that case number to the top of your DAIL
			    Call write_value_and_transmit("T", dail_row + 1, 3)
			End if

            dail_row = 6  'resetting the DAIL row '

            'Reading the DAIL Information
			EMReadScreen MAXIS_case_number, 8, dail_row - 1, 73
            MAXIS_case_number = trim(MAXIS_case_number)

            EMReadScreen dail_type, 4, dail_row, 6

            EMReadScreen dail_msg, 61, dail_row, 20
			dail_msg = trim(dail_msg)

            EMReadScreen dail_month, 8, dail_row, 11
            dail_month = trim(dail_month)

            stats_counter = stats_counter + 1   'I increment thee
            Call non_actionable_dails(actionable_dail)   'Function to evaluate the DAIL messages

            If TPQY_checkbox = 1 then
                If instr(dail_msg, "TPQY RESPONSE RECEIVED FROM SSA") then actionable_dail = False  'cleaning up TPQY messages after BULK SVES/QURY for SSI/RSDI RAP project.
            End if

			'Accounting for duplicate DAIL messages
			dail_string = worker & " " & MAXIS_case_number & " " & dail_type & " " & dail_month & " " & dail_msg

            'special handling for duplicate PEPR messages in CM and CM + 1
            If dail_type = "PEPR" then 
                'If the message has already been determined to be non-actionable, we don't need to evaluate those.
                If actionable_dail = True then 
                    'PEPR determination for duplicate messages that are CM + 1
                    last_month_dail_string = dail_string = worker & " " & MAXIS_case_number & " " & dail_type & " " & last_month & " " & dail_msg
                    'if last month's message was found in the all_day_array then the CM + 1 messages is non-actionable
                    If instr(all_dail_array, "*" & last_month_dail_string & "*") then 
                        actionable_dail = False 
                    Else
                        'otherwise it's captured. This happens with a lot of HC program PEPR's. 
                        actionable_dail = True 
                    End if 
                End if 
            End if 

            'If the case number is found in the string of case numbers, it's not added again.
            If instr(all_dail_array, "*" & dail_string & "*") then
                If dail_type = "HIRE" then
                    capture_message = True
                Else    
                    capture_message = False
					false_count = false_count + 1
                End if
            else
                capture_message = True
            End if

			If capture_message = True then
				all_dail_array = trim(all_dail_array & dail_string & "*") 'Adding dail_string to all_daily_array
                IF actionable_dail = False then
			    	'--------------------------------------------------------------------actionable_dail = False will captured in Excel and deleted.
			    	objExcel.Cells(excel_row, 1).Value = worker
			    	objExcel.Cells(excel_row, 2).Value = MAXIS_case_number
			    	objExcel.Cells(excel_row, 3).Value = dail_type
			    	objExcel.Cells(excel_row, 4).Value = dail_month
			    	objExcel.Cells(excel_row, 5).Value = dail_msg
			    	excel_row = excel_row + 1
                    deleted_dails = deleted_dails + 1
			    else
			    	actionable_dail = True      'actionable_dail = True will NOT be deleted and will be captured and reported out as actionable.
                    ReDim Preserve DAIL_array(4, DAIL_count)	'This resizes the array based on the number of rows in the Excel File'
                	DAIL_array(worker_const,	           DAIL_count) = worker
                	DAIL_array(maxis_case_number_const,    DAIL_count) = MAXIS_case_number
                	DAIL_array(dail_type_const, 	       DAIL_count) = dail_type
                	DAIL_array(dail_month_const, 		   DAIL_count) = dail_month
                	DAIL_array(dail_msg_const, 		       DAIL_count) = dail_msg
                    Dail_count = DAIL_count + 1
			    End if
			End if

			'Navigation handling for if a case is actionable or not. If actionable the dail_row needs to increment
			If actionable_DAIL = False then
				Call write_value_and_transmit("D", dail_row, 3)
			    EMReadScreen other_worker_error, 13, 24, 2
			    If other_worker_error = "** WARNING **" then transmit
			Elseif actionable_DAIL = True then
				dail_row = dail_row + 1
			End if

            EMReadScreen message_error, 11, 24, 2		'Cases can also NAT out for whatever reason if the no messages instruction comes up.
            If message_error = "NO MESSAGES" then exit do

			'...going to the next page if necessary
			EMReadScreen next_dail_check, 4, dail_row, 4
			If trim(next_dail_check) = "" then
				PF8
				EMReadScreen last_page_check, 16, 24, 2
                'DAIL/PICK will look for 'no message worker X127XXX as the full message.
                If last_page_check = "THIS IS THE LAST" or last_page_check = dail_end_msg then
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
objExcel.Cells(8, 7).Value = "False count/duplicate DAIL Messages not counted:"
objExcel.Columns(7).Font.Bold = true
objExcel.Cells(2, 8).Value = deleted_dails
objExcel.Cells(3, 8).Value = STATS_manualtime
objExcel.Cells(4, 8).Value = STATS_counter * STATS_manualtime
objExcel.Cells(5, 8).Value = timer - start_time
objExcel.Cells(6, 8).Value = ((STATS_counter * STATS_manualtime) - (timer - start_time)) / 60
objExcel.Cells(7, 8).Value = STATS_counter
objExcel.Cells(8, 8).Value = false_count

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

'Export information to Excel re: case status
For item = 0 to UBound(DAIL_array, 2)
	objExcel.Cells(excel_row, 1).Value = DAIL_array(worker_const, item)
	objExcel.Cells(excel_row, 2).Value = DAIL_array(maxis_case_number_const, item)
    objExcel.Cells(excel_row, 3).Value = DAIL_array(dail_type_const, item)
	objExcel.Cells(excel_row, 4).Value = DAIL_array(dail_month_const, item)
    objExcel.Cells(excel_row, 5).Value = DAIL_array(dail_msg_const, item)
	excel_row = excel_row + 1
Next

objExcel.Cells(1, 7).Value = "Remaining DAIL messages:"
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

script_end_procedure("Success! Please review the list created for accuracy.")

'----------------------------------------------------------------------------------------------------Closing Project Documentation - Version date 01/12/2023
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------11/20/2023
'--Tab orders reviewed & confirmed----------------------------------------------11/20/2023
'--Mandatory fields all present & Reviewed--------------------------------------11/20/2023
'--All variables in dialog match mandatory fields-------------------------------11/20/2023
'Review dialog names for content and content fit in dialog----------------------11/20/2023
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------11/20/2023-------------------N/A
'--CASE:NOTE Header doesn't look funky------------------------------------------11/20/2023-------------------N/A
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------11/20/2023-------------------N/A
'--write_variable_in_CASE_NOTE function: confirm that proper punctuation is used-11/20/2023-------------------N/A
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------11/20/2023
'--MAXIS_background_check reviewed (if applicable)------------------------------11/20/2023-------------------N/A
'--PRIV Case handling reviewed -------------------------------------------------11/20/2023-------------------N/A
'--Out-of-County handling reviewed----------------------------------------------11/20/2023-------------------N/A
'--script_end_procedures (w/ or w/o error messaging)----------------------------11/20/2023
'--BULK - review output of statistics and run time/count (if applicable)--------11/20/2023
'--All strings for MAXIS entry are uppercase vs. lower case (Ex: "X")-----------11/20/2023
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------11/20/2023
'--Incrementors reviewed (if necessary)-----------------------------------------11/20/2023
'--Denomination reviewed -------------------------------------------------------11/20/2023
'--Script name reviewed---------------------------------------------------------11/20/2023
'--BULK - remove 1 incrementor at end of script reviewed------------------------11/20/2023

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------11/20/2023
'--comment Code-----------------------------------------------------------------11/20/2023
'--Update Changelog for release/update------------------------------------------11/20/2023-------------------N/A
'--Remove testing message boxes-------------------------------------------------11/20/2023
'--Remove testing code/unnecessary code-----------------------------------------11/20/2023
'--Review/update SharePoint instructions----------------------------------------11/20/2023-------------------N/A
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------11/20/2023-------------------N/A
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------11/20/2023
'--COMPLETE LIST OF SCRIPTS update policy references----------------------------11/20/2023-------------------N/A
'--Complete misc. documentation (if applicable)---------------------------------11/20/2023
'--Update project team/issue contact (if applicable)----------------------------11/20/2023