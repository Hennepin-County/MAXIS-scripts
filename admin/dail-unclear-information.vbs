'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ADMIN - DAIL UNCLEAR INFORMATION.vbs"
start_time = timer
STATS_counter = 1                       'sets the stats counter at one
STATS_manualtime = 30
STATS_denomination = "I"       			'I is for each item
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
call changelog_update("08/21/2023", "Initial version.", "Mark Riegel, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'----------------------------------------------------------------------------------------------------FUNCTIONS
'Add function for X Number restart functionality
Function create_array_of_all_active_x_numbers_in_county_with_restart(array_name, two_digit_county_code, restart_status, restart_x_number)
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
                If trim(UCase(worker_ID)) = trim(UCase(restart_x_number)) then
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
			MAXIS_row = 7	'redeclaring MAXIS row so as to start reading from the top of the list again
		End if
	Loop until more_pages_check = "More:  " or more_pages_check = "       "	'The or works because for one-page only counties, this will be blank

    array_name = split(array_name)
End function

'----------------------------------------------------------------------------------------------------THE SCRIPT
EMConnect ""
Call Check_for_MAXIS(False)

'Sets the county code for Hennepin County as X127
worker_county_code = "X127"

'Set footer month and year to current month and year
MAXIS_footer_month = CM_mo
MAXIS_footer_year = CM_yr
'To determine if DAIL message is in scope based on DAIL month, creating variable for date for current month, day, and year
footer_month_day_year = dateadd("d", 0, MAXIS_footer_month & "/1/20" & MAXIS_footer_year)

EMWriteScreen MAXIS_footer_month, 20, 43
EMWriteScreen MAXIS_footer_year, 20, 46

'Initial dialog - select whether to create a list or process a list
BeginDialog Dialog1, 0, 0, 306, 220, "DAIL Unclear Information"
  GroupBox 10, 5, 290, 65, "Using the DAIL Unclear Information Script"
  Text 20, 15, 275, 50, "A BULK script that gathers and processes selected (HIRE and/or CSES) DAIL messages for the agency that fall under the Food and Nutrition Service's unclear information rules. As the DAIL messages are reviewed, the script will process DAIL messages for 6-month reporters on SNAP-only and process the DAIL messages accordingly. The data will be exported in a Microsoft Excel file type (.xlsx) and saved in the LAN. "
  Text 15, 80, 175, 10, "Type of DAIL Messages to Process:"
  CheckBox 15, 90, 55, 10, "CSES", CSES_messages
  CheckBox 15, 100, 55, 10, "HIRE", HIRE_messages
  Text 15, 115, 185, 10, "Select the X Numbers to Process (check one box only):"
  CheckBox 15, 125, 90, 10, "Process ALL X Numbers", process_all_x_numbers
  CheckBox 15, 135, 225, 10, "RESTART Process ALL X Numbers (enter restart X Number below)", restart_process_all_x_numbers
  EditBox 25, 145, 85, 15, restart_x_number
  CheckBox 15, 165, 255, 10, "Process Specific X Numbers (enter X Numbers below separated by comma)", process_specific_x_numbers
  EditBox 25, 175, 270, 15, specific_x_numbers_to_process
  Text 15, 205, 60, 10, "Worker Signature:"
  EditBox 80, 200, 110, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 205, 200, 40, 15
    CancelButton 245, 200, 40, 15
EndDialog

DO
    Do
        err_msg = ""    'This is the error message handling
        Dialog Dialog1
        cancel_confirmation

        'Dialog field validation
        'Validation to ensure that at least CSES or HIRE messages checkbox is checked
        If CSES_messages = 0 AND HIRE_messages = 0 Then err_msg = err_msg & vbCr & "* Select either CSES or HIRE messages, or both. Both checkboxes cannot be left blank."
        'Validation to ensure that only one option is selected for X Numbers to process
        If process_specific_x_numbers + process_all_x_numbers + restart_process_all_x_numbers <> 1 Then err_msg = err_msg & vbCr & "* You can only select one option for processing X Numbers. Make sure only one box is checked."
        'Validation to ensure that Specific X Numbers and Restart Process All X Numbers fields are left blank if processing all X Numbers
        If process_all_x_numbers = 1 AND (trim(specific_x_numbers_to_process) <> "" OR trim(restart_x_number) <> "") Then err_msg = err_msg & vbCr & "* You selected the option to Process All X Numbers. The entry fields for Process Specific X Numbers and RESTART Process All X Numbers must be blank to proceed."
        'Validation to ensure that Process Specific X Numbers field is blank if Restart Process All X Numbers is selected
        If restart_process_all_x_numbers = 1 AND trim(specific_x_numbers_to_process) <> "" Then err_msg = err_msg & vbCr & "* You selected the option to Restart Process All X Numbers. The entry field for Process Specific X Numbers must be blank to proceed."
        If restart_process_all_x_numbers = 1 AND trim(restart_x_number) = "" Then err_msg = err_msg & vbCr & "* You selected the option to Restart Process All X Numbers. The entry field for Restart Process All X Numbers is empty. Enter the X Number that the script should restart on."
        'Validation to ensure that Restart Process All X Numbers field is blank if Process Specific X Numbers is selected
        If process_specific_x_numbers = 1 AND trim(restart_x_number) <> "" Then err_msg = err_msg & vbCr & "* You selected the option to Process Specific X Numbers. The entry field for RESTART Process All X Numbers must be empty. Clear the field to proceed."
        'Validation to ensure that if processing specific X numbers, the list of X numbers field is not blank
        If process_specific_x_numbers = 1 AND trim(specific_x_numbers_to_process) = "" Then err_msg = err_msg & vbCr & "* You selected the option to Process Specific X Numbers. You must enter a list of X Numbers separated by a comma to proceed. The entry field is currently empty."
        'Ensures worker signature is not blank
        IF trim(worker_signature) = "" THEN err_msg = err_msg & vbCr & "* Please enter your worker signature."
        IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
    Loop until err_msg = "" and ButtonPressed = OK
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in	

'Determining if this is a restart or not in function below when gathering the x numbers.
If restart_process_all_x_numbers = 0 then
    restart_status = False
Else 
	restart_status = True
End if 

'If all workers are selected, the script will go to REPT/USER, and load all of the workers into an array. Otherwise it'll create a single-object "array" just for simplicity of code.
If process_all_x_numbers = 1 OR restart_process_all_x_numbers = 1 then
	Call create_array_of_all_active_x_numbers_in_county_with_restart(worker_array, two_digit_county_code, restart_status, restart_x_number)
Else
	x_numbers_from_dialog = split(specific_x_numbers_to_process, ", ")	'Splits the worker array based on commas

	'Need to add the worker_county_code to each one
	For each x_number in x_numbers_from_dialog
		If worker_array = "" then
			worker_array = trim(ucase(x_number))		'replaces worker_county_code if found in the typed x1 number
		Else
			worker_array = worker_array & "," & trim(ucase(x_number)) 'replaces worker_county_code if found in the typed x1 number
		End if
	Next
	'Split worker_array
	worker_array = split(worker_array, ",")
End if

Call check_for_MAXIS(False)

'Opening the Excel file for list of DAIL messages
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

'Changes name of Excel sheet to DAIL Messages to capture details about in-scope DAIL messages
ObjExcel.ActiveSheet.Name = "DAIL Messages"

'Excel headers and formatting the columns
objExcel.Cells(1, 1).Value = "Case Number"
objExcel.Cells(1, 2).Value = "X Number"
objExcel.Cells(1, 3).Value = "DAIL Type"
objExcel.Cells(1, 4).Value = "DAIL Month"
'To do - determine if FULL DAIL message should be captured
objExcel.Cells(1, 5).Value = "DAIL Message"
objExcel.Cells(1, 6).Value = "Renewal Month Determination"
objExcel.Cells(1, 7).Value = "Processable based on DAIL"
objExcel.Cells(1, 8).Value = "Processing Notes for DAIL Message"

FOR i = 1 to 8		'formatting the cells'
    objExcel.Cells(1, i).Font.Bold = True		'bold font'
    ObjExcel.columns(i).NumberFormat = "@" 		'formatting as text
    objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

'Creating second Excel sheet for compiling case details
ObjExcel.Worksheets.Add().Name = "Case Details"

'Excel headers and formatting the columns
objExcel.Cells(1, 1).Value = "Case Number"
objExcel.Cells(1, 2).Value = "X Number"
objExcel.Cells(1, 3).Value = "SNAP Status"
objExcel.Cells(1, 4).Value = "Other Programs Present"
objExcel.Cells(1, 5).Value = "Reporting Status"
objExcel.Cells(1, 6).Value = "SR Report Date"
objExcel.Cells(1, 7).Value = "Recertification Date"
objExcel.Cells(1, 8).Value = "Processing Notes for Case"
objExcel.Cells(1, 9).Value = "Processable based on Case Details"

FOR i = 1 to 9		'formatting the cells'
    objExcel.Cells(1, i).Font.Bold = True		'bold font'
    ObjExcel.columns(i).NumberFormat = "@" 		'formatting as text
    objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

'Creates sheet to track stats for the script
ObjExcel.Worksheets.Add().Name = "Stats"

STATS_counter = STATS_counter - 1
'Enters info about runtime for the benefit of folks using the script
'To do - update to reflect actual stats needed/wanted
objExcel.Cells(1, 1).Value = "Number of DAIL Messages Added to List:"
objExcel.Cells(2, 1).Value = "Average time to find/select/copy/paste one line (in seconds):"
objExcel.Cells(3, 1).Value = "Estimated manual processing time (lines x average):"
objExcel.Cells(4, 1).Value = "Script run time (in seconds):"
objExcel.Cells(5, 1).Value = "Estimated time savings by using script (in minutes):"


FOR i = 1 to 5		'formatting the cells'
    objExcel.Cells(i, 1).Font.Bold = True		'bold font'
    ObjExcel.rows(i).NumberFormat = "@" 		'formatting as text
    objExcel.columns(1).AutoFit()				'sizing the columns'
NEXT

'Create an array to track in-scope DAIL messages
DIM DAIL_message_array()

ReDim DAIL_message_array(8, 0)
'Incrementor for the array
Dail_count = 0

'constants for array
const dail_maxis_case_number_const      = 0
const dail_worker_const	                = 1
const dail_type_const                   = 2
const dail_month_const		            = 3
const dail_msg_const		            = 4
const renewal_month_determination_const = 5
const processable_based_on_dail_const   = 6
'To do - processing notes, would these be captured in case details array?
const dail_processing_notes_const       = 7
' To Do - is the excel row constant needed?
const dail_excel_row_const              = 8

'Sets variable for the Excel row to export data to Excel sheet
dail_excel_row = 2

'Create an array to track case details
DIM case_details_array()

ReDim case_details_array(9, 0)
'Incrementor for the array
case_count = 0

'constants for array
const case_maxis_case_number_const      = 0
const case_worker_const	                = 1
const snap_status_const                 = 2
const other_programs_present_const      = 3
const reporting_status_const            = 4
const sr_report_date_const              = 5
const recertification_date_const        = 6
'To do - processing notes, would these be captured in case details array?
const case_processing_notes_const       = 7
const processable_based_on_case_const   = 8
' To Do - is the excel row constant needed?
const case_excel_row_const              = 9

'Sets variable for the Excel row to export data to Excel sheet
case_excel_row = 2

'Create an array with PMIs to match with CASE/PERS info
Dim PMI_and_ref_nbr_array()

'Reset the array 
ReDim PMI_and_ref_nbr_array(3, 0)

'Incrementor for the array
'To do - necessary?
member_count = 0

'Constants for the array
const ref_nbr_const           = 0
const PMI_const               = 1
const PMI_match_found_const   = 2
const hh_member_count_const   = 3

'To Do - add tracking of deleted dails once processing the list
'deleted_dails = 0	'establishing the value of the count for deleted deleted_dails

For each worker in worker_array
    ' MsgBox worker

    'Clearing out MAXIS case number so that it doesn't carry forward from previous case
    MAXIS_case_number = ""
    
    'To do - verify placement of these strings for lists of case numbers, dail messages, etc.

    'Resetting all of the string lists
    'Creating initial string for tracking list of valid case numbers pulled from REPT/ACTV. This is used to avoid triggering a privileged case and losing connection to DAIL
    valid_case_numbers_list = "*"

    'Create list of case numbers to be used for comparison purposes as the script navigates through the DAIL
    list_of_all_case_numbers = "*"

    'Create list of DAIL messages that should be deleted. If a DAIL message matches, then it will be deleted. This is needed because DAIL will reset to first DAIL message for case number anytime the script goes to CASE/CURR, CASE/PERS, STAT/UNEA, etc.
    list_of_DAIL_messages_to_delete = "*"

    'Create list of DAIL messages that should be skipped. If a DAIL message matches, then the script will skip past it to next DAIL row. This is needed because DAIL will reset to first DAIL message for case number anytime the script goes to CASE/CURR, CASE/PERS, STAT/UNEA, etc. 
    list_of_DAIL_messages_to_skip = "*"

    'Formatting the worker so there are no errors
    worker = trim(ucase(worker))

    'Does this to prevent "ghosting" where the old info shows up on the new screen for some reason					
	back_to_self	

	Call navigate_to_MAXIS_screen("REPT", "ACTV")
	EMWriteScreen worker, 21, 13
	TRANSMIT
	EMReadScreen user_worker, 7, 21, 71
	EMReadScreen p_worker, 7, 21, 13
	IF user_worker = p_worker THEN PF7		'If the user is checking their own REPT/ACTV, the script will back up to page 1 of the REPT/ACTV

	'msgbox "worker " & worker

	IF worker_number = "X127CCL" or worker = "127CCL" THEN
		DO
			EmReadScreen worker_confirmation, 20, 3, 11 'looking for CENTURY PLAZA CLOSED
			EMWaitReady 0, 0
		LOOP UNTIL worker_confirmation = "CENTURY PLAZA CLOSED"
	END IF

	'Skips workers with no info
	EMReadScreen has_content_check, 1, 7, 8
	If has_content_check <> " " then
		'Grabbing each case number on screen
		Do
		    'Set variable for next do...loop
			MAXIS_row = 7
			'Checking for the last page of cases.
			EMReadScreen last_page_check, 21, 24, 2	'because on REPT/ACTV it displays right away, instead of when the second F8 is sent
			EMReadscreen number_of_pages, 4, 3, 76 'getting page number because to ensure it doesnt fail'
			number_of_pages = trim(number_of_pages)
			Do
				EMReadScreen MAXIS_case_number, 8, MAXIS_row, 12	'Reading case number

				'Doing this because sometimes BlueZone registers a "ghost" of previous data when the script runs. This checks against an array and stops if we've seen this one before.
				MAXIS_case_number = trim(MAXIS_case_number)
				If MAXIS_case_number <> "" and instr(valid_case_numbers_list, "*" & MAXIS_case_number & "*") <> 0 then exit do
				valid_case_numbers_list = trim(valid_case_numbers_list & MAXIS_case_number & "*")

				If MAXIS_case_number = "" Then Exit Do			'Exits do if we reach the end

				MAXIS_row = MAXIS_row + 1
				MAXIS_case_number = ""			'Blanking out variable
			Loop until MAXIS_row = 19
			PF8
		Loop until last_page_check = "THIS IS THE LAST PAGE"
	END IF

    'Ensuring valid_case_numbers is blanked out
    ' msgbox valid_case_numbers_list

    'Navigates to DAIL to pull DAIL messages
    MAXIS_case_number = ""
    CALL navigate_to_MAXIS_screen("DAIL", "PICK")
    EMWriteScreen "_", 7, 39    'blank out ALL selection
    'Selects CSES DAIL Type based on dialog selection
    If CSES_messages = 1 Then EMWriteScreen "X", 10, 39    'Select CSES DAIL Type
    'Selects INFO (HIRE) DAIL Type based on dialog selection
    If HIRE_messages = 1 Then EMWriteScreen "X", 13, 39
    transmit

    'Enter the worker number on DAIL to pull up DAIL messages
    Call write_value_and_transmit(worker, 21, 6)
    'Transmits past not your dail message
    transmit 

    'Reads where the count of DAILs is listed. Used to verify DAIL is not empty.
    EMReadScreen number_of_dails, 1, 3, 67		

    DO
       'If this space is blank the rest of the DAIL reading is skipped
        If number_of_dails = " " Then exit do		
        'Because the script brings each new case to the top of the page, dail_row starts at 6.
        dail_row = 6	

        DO
            ' To do - verify if variables are resetting properly every do loop
            dail_type = ""
            dail_msg = ""
            dail_month = ""
            MAXIS_case_number = ""
            INFC_dail_msg = ""

            'To do - do we need to reset the full dail message or any other variables?


            'Determining if the script has moved to a new case number within the dail, in which case it needs to move down one more row to get to next dail message
            EMReadScreen new_case, 8, dail_row, 63
            new_case = trim(new_case)
            IF new_case <> "CASE NBR" THEN 
                'If there is NOT a new case number, the script will top the message
                Call write_value_and_transmit("T", dail_row, 3)
            ELSEIF new_case = "CASE NBR" THEN
                'If the script does find that there is a new case number (indicated by "CASE NBR"), it will write a "T" in the next row and transmit, bringing that case number to the top of your DAIL
                Call write_value_and_transmit("T", dail_row + 1, 3)
            End if

            'Resets the DAIL row since the message has now been topped
            dail_row = 6  

            'Determines the DAIL Type
            EMReadScreen dail_type, 4, dail_row, 6
            dail_type = trim(dail_type)

            'Determine the DAIL msg so that INFC messages can be excluded
            EMReadScreen dail_msg, 61, dail_row, 20
            dail_msg = trim(dail_msg)
            If InStr(dail_msg, "INFC") then 
                INFC_dail_msg = True
            Else
                INFC_dail_msg = False
            End If

            If (INFC_dail_msg = False AND dail_type = "CSES") OR dail_type = "HIRE" Then
                'Read the MAXIS Case Number, if it is a new case number then pull case details. If it is not a new case number, then do not pull new case details.
                
                ' Msgbox "(INFC_dail_msg = False AND dail_type = 'CSES') OR dail_type = 'HIRE' Then"

                EMReadScreen MAXIS_case_number, 8, dail_row - 1, 73
                MAXIS_case_number = trim(MAXIS_case_number)

                If InStr(valid_case_numbers_list, "*" & MAXIS_case_number & "*") Then
                    'If the case is in the valid_case_numbers_list, then it can be evaluated. If it is NOT in the valid_case_numbers_list then it is likely privileged or out of county so it will be skipped

                    If Instr(list_of_all_case_numbers, "*" & MAXIS_case_number & "*") = 0 Then
                        'If the MAXIS case number is NOT in the list of all case numbers, then it is a new case number and the script will gather case details
                        
                        'Redim the case details array and add to array
                        ReDim Preserve case_details_array(case_excel_row_const, case_count)
                        case_details_array(case_maxis_case_number_const, case_count) = MAXIS_case_number
                        case_details_array(case_worker_const, case_count) = worker
                
                        'Since case number is not in list of all case numbers, add it to the list
                        list_of_all_case_numbers = list_of_all_case_numbers & MAXIS_case_number & "*"

                        'Navigate to CASE/CURR to pull case details 
                        Call write_value_and_transmit("H", dail_row, 3)

                        'Handling if the case is out of county
                        EmReadscreen worker_county, 4, 21, 14
                        If worker_county <> worker_county_code then
                            case_details_array(case_processing_notes_const, case_count) = "Out-of-County Case"
                            case_details_array(processable_based_on_case_const, case_count) = False
                        Else
                            'Pull case details from CASE/CURR, maintains connection to DAIL
                            Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type, case_status, list_active_programs, list_pending_programs)

                            'Set SNAP status within array
                            case_details_array(snap_status_const, case_count) = trim(snap_status)

                            'Function (determine_program_and_case_status_from_CASE_CURR) sets dail_row equal to 4 so need to reset it.
                            dail_row = 6
                            

                            'If SNAP is not active, then the case is out-of-scope for Unclear Information.
                            If snap_status <> "ACTIVE" then 
                                case_details_array(processable_based_on_case_const, case_count) = False
                                case_details_array(reporting_status_const, case_count) = "N/A"
                                case_details_array(recertification_date_const, case_count) = "N/A"
                                case_details_array(sr_report_date_const, case_count) = "N/A"
                                case_details_array(case_processing_notes_const, case_count) = "N/A"
                            End If

                            'If other programs are active/pending then case does not fall under Unclear Information scope
                            If ga_case = True OR _
                                msa_case = True OR _
                                mfip_case = True OR _
                                dwp_case = True OR _
                                grh_case = True OR _
                                ma_case = True OR _
                                msp_case = True then
                                    case_details_array(other_programs_present_const, case_count) = True
                                    case_details_array(processable_based_on_case_const, case_count) = False
                            Else
                                'No other programs are active or pending so set value to false
                                case_details_array(other_programs_present_const, case_count) = False
                            End if

                            If snap_status = "ACTIVE" then
                                'To do - check if background check is needed, may break connection to DAIL
                                ' Call MAXIS_background_check

                                'Navigate to ELIG/FS from CASE/CURR to maintain tie to DAIL
                                EMWriteScreen "ELIG", 20, 22
                                Call write_value_and_transmit("FS  ", 20, 69)

                                EMReadScreen no_SNAP, 10, 24, 2
                                If no_SNAP = "NO VERSION" then						'NO SNAP version means no determination
                                    case_details_array(case_processing_notes_const, case_count) = case_details_array(case_processing_notes_const, case_count) & "No version of SNAP exists for " & MAXIS_footer_month & "/" & MAXIS_footer_year & ". "
                                    case_details_array(processable_based_on_case_const, case_count) = False
                                Else

                                    EMWriteScreen "99", 19, 78
                                    transmit
                                    'This brings up the FS versions of eligibility results to search for approved versions
                                    status_row = 7
                                    Do
                                        EMReadScreen app_status, 8, status_row, 50
                                        app_status = trim(app_status)
                                        If app_status = "" then
                                            PF3
                                            exit do 	'if end of the list is reached then exits the do loop
                                        End if
                                        If app_status = "UNAPPROV" Then status_row = status_row + 1
                                    Loop until app_status = "APPROVED" or app_status = ""

                                    If app_status = "" or app_status <> "APPROVED" then
                                        case_details_array(case_processing_notes_const, case_count) = case_details_array(case_processing_notes_const, case_count) & "No approved eligibility results exists in " & MAXIS_footer_month & "/" & MAXIS_footer_year & ". "
                                        case_details_array(processable_based_on_case_const, case_count) = False
                                        'To do - since pulling cases for REPT/ACTV, this may never trigger but using message box just in case. Delete after testing on all worker numbers
                                        MsgBox "Instance where SNAP is active but there is not app status or it is not approved"
                                    Elseif app_status = "APPROVED" then
                                        EMReadScreen vers_number, 1, status_row, 23
                                        Call write_value_and_transmit(vers_number, 18, 54)
                                        Call write_value_and_transmit("FSSM", 19, 70)
                                        EmReadscreen reporting_status, 12, 8, 31
                                        EmReadscreen recertification_date, 8, 11, 31
                                        'Converts date from string to date
                                        recertification_date = DateAdd("m", 0, recertification_date)

                                        If InStr(reporting_status, "SIX MONTH") Then 
                                            sr_report_date = DateAdd("m", -6, recertification_date)
                                        Else
                                            sr_report_date = "N/A"
                                        End If
                                        ' MsgBox "Updating the footer month and year"
                                        'Change the footer month and year back to CM/CY otherwise the DAIL will carry forward footer month and year from previous DAIL message as it moves to the next one and could cause error
                                        'To do - is this necessary?
                                        EMWriteScreen MAXIS_footer_month, 19, 54
                                        EMWriteScreen MAXIS_footer_year, 19, 57
                                        ' MsgBox "did footer month year update?"
                                    End if
                                    
                                    ' MsgBox "Updating the case_details_array"
                                    'Update the array with new case details
                                    case_details_array(reporting_status_const, case_count) = trim(reporting_status)
                                    case_details_array(recertification_date_const, case_count) = trim(recertification_date)
                                    case_details_array(sr_report_date_const, case_count) = trim(sr_report_date)
                                End if
                            Else
                                case_details_array(reporting_status_const, case_count) = "N/A"
                                ' MsgBox "Updating the footer month and year"
                                'Update the footer month and year to CM/CY on CASE/CURR before returning to DAIL
                                'To do - is this necessary?
                                EMWriteScreen MAXIS_footer_month, 20, 54
                                EMWriteScreen MAXIS_footer_year, 20, 57
                                ' MsgBox "did footer month year update?"
                            End If 

                            'To do - remove comments below after testing
                            'Determine if processable_based_on_case is True (case is within scope, MAY act on DAIL message) or if processable_based_on_case is false (case is outside of scope, WILL NOT act on DAIL message)
                            ' Msgbox "case_details_array(snap_status_const, case_count): " & case_details_array(snap_status_const, case_count)
                            ' Msgbox "case_details_array(other_programs_present_const, case_count): " & case_details_array(other_programs_present_const, case_count)
                            ' Msgbox "case_details_array(reporting_status_const, case_count): " & case_details_array(reporting_status_const, case_count)
                        End If    
                        
                        If case_details_array(snap_status_const, case_count) = "ACTIVE" AND case_details_array(other_programs_present_const, case_count) = False AND case_details_array(reporting_status_const, case_count) = "SIX MONTH" then
                            case_details_array(processable_based_on_case_const, case_count) = True
                        Else
                            case_details_array(processable_based_on_case_const, case_count) = False
                        End if

                        'Activate the case details sheet
                        objExcel.Worksheets("Case Details").Activate

                        'Update the Case Details sheet with case data
                        objExcel.Cells(case_excel_row, 1).Value = case_details_array(case_maxis_case_number_const, case_count)
                        objExcel.Cells(case_excel_row, 2).Value = case_details_array(case_worker_const, case_count)
                        objExcel.Cells(case_excel_row, 3).Value = case_details_array(snap_status_const, case_count)
                        objExcel.Cells(case_excel_row, 4).Value = case_details_array(other_programs_present_const, case_count)
                        objExcel.Cells(case_excel_row, 5).Value = case_details_array(reporting_status_const, case_count)
                        objExcel.Cells(case_excel_row, 6).Value = case_details_array(sr_report_date_const, case_count)
                        objExcel.Cells(case_excel_row, 7).Value = case_details_array(recertification_date_const, case_count)
                        objExcel.Cells(case_excel_row, 8).Value = case_details_array(case_processing_notes_const, case_count)
                        objExcel.Cells(case_excel_row, 9).Value = case_details_array(processable_based_on_case_const, case_count)
                        case_excel_row = case_excel_row + 1

                        'Return to DAIL by PF3
                        PF3
                      
                        'Increment the case_count for updating the array
                        case_count = case_count + 1
                        'Subtract one from dail_row so that the dail_row restarts evaluation of cases now with case details
                        ' Msgbox "subtract 1 from dail? Dail row is currently: " & dail_row
                        dail_row = dail_row - 1
                        ' Msgbox "After subtraction, dail_row = " & dail_row
                    
                    Else
                        'If the MAXIS case number IS in the list of all case numbers, then it is not a new case number and no case details need to be gathered. It can work off the already collected case details.

                        'Before determining whether the DAIL is processable, script determines if it has encountered this DAIL message previously. Based on determination, it then processes (deletes) the dail, skips it, or makes processable determination

                        'Resetting the full_dail_msg to ensure it is not carrying forward to subsequent loops
                        'To do - necessary?
                        full_dail_msg = ""

                        'Script opens the entire DAIL message to evaluate if it is a new message or not
                        Call write_value_and_transmit("X", dail_row, 3)

                        ' Script reads the full DAIL message so that it can process, or not process, as needed.
                        EMReadScreen full_dail_msg_line_1, 60, 9, 5

                        full_dail_msg_line_1 = trim(full_dail_msg_line_1)
                        ' Msgbox full_dail_msg_line_1

                        EMReadScreen full_dail_msg_line_2, 60, 10, 5
                        full_dail_msg_line_2 = trim(full_dail_msg_line_2)
                        ' Msgbox full_dail_msg_line_2

                        EMReadScreen full_dail_msg_line_3, 60, 11, 5
                        full_dail_msg_line_3 = trim(full_dail_msg_line_3)
                        ' If full_dail_msg_line_3 <> "" Then Msgbox full_dail_msg_line_3

                        EMReadScreen full_dail_msg_line_4, 60, 12, 5
                        full_dail_msg_line_4 = trim(full_dail_msg_line_4)
                        ' If full_dail_msg_line_4 <> "" Then Msgbox full_dail_msg_line_4

                        full_dail_msg = full_dail_msg_line_1 & " " & full_dail_msg_line_2 & " " & full_dail_msg_line_3 & " " & full_dail_msg_line_4

                        ' Msgbox full_dail_msg

                        'Transmit back to dail
                        transmit

                        'Confirming that dail message lists are updating properly
                        ' Msgbox "list_of_DAIL_messages_to_delete: " & list_of_DAIL_messages_to_delete
                        ' Msgbox "list_of_DAIL_messages_to_skip: " & list_of_DAIL_messages_to_skip

                        'The script has the full DAIL message and can compare against delete and skip lists to determine if it is a new message

                        'To do - consider more robust handling, should we validate that case number matches? That dail month matches? These could be added to the string - i.e. *123456 - CS DISB Type 36....*
                        If Instr(list_of_DAIL_messages_to_delete, "*" & full_dail_msg & "*") Then
                            'If the full dail message is within the list of dail messages to delete then the message should be deleted
                            'To do - Add handling here for deleting DAIL messages
                            ' MsgBox "This message is on the delete list. It would normally be deleted"
                            'To do - remove adding dail to row because it would happen automatically if deleted
                            ' MsgBox "Where is the dail row? Should it be increased?"
                            ' dail_row = dail_row + 1
                        ElseIf Instr(list_of_DAIL_messages_to_skip, "*" & full_dail_msg & "*") Then
                            'If the full message is on the list of dail messages to skip then the message should be skipped
                            'To do - Add handling for messages to skip
                            ' MsgBox "This message is on the skip list. It should be skipped."
                            ' MsgBox "Where is the dail row? It should be increased so that it is skipped?"
                            
                            'Go to next dail_row
                            ' dail_row = dail_row + 1

                        ElseIf Instr(list_of_DAIL_messages_to_delete, "*" & full_dail_msg & "*") = 0 AND Instr(list_of_DAIL_messages_to_skip, "*" & full_dail_msg & "*") = 0 Then
                            'If the full dail message is NOT in the list of dail messages to delete AND the full dail messages is NOT in the list of skip messages then it SHOULD be a new dail message and therefore it needs to be evaluated

                            'Gather details on DAIL message, should capture DAIL details in spreadsheet even if ultimately not actionable
                        
                            ' MsgBox "This is a new DAIL message. It should be processed accordingly."

                            'Reset the array
                            ReDim Preserve DAIL_message_array(DAIL_excel_row_const, dail_count)
                            DAIL_message_array(dail_maxis_case_number_const, DAIL_count) = MAXIS_case_number
                            DAIL_message_array(dail_worker_const, DAIL_count) = worker

                            ' Msgbox "DAIL_message_array(dail_maxis_case_number_const, DAIL_count): " & DAIL_message_array(dail_maxis_case_number_const, DAIL_count)
                            ' Msgbox "DAIL_message_array(dail_worker_const, DAIL_count): " & DAIL_message_array(dail_worker_const, DAIL_count)

                            'Use for next loop to match the individual DAIL message to the corresponding array item of matching Case Details
                            for each_case = 0 to UBound(case_details_array, 2)
                                'Iterate through each of the cases 
                                If DAIL_message_array(dail_maxis_case_number_const, dail_count) = case_details_array(case_maxis_case_number_const, each_case) Then
                                    'As the for to loop iterates through each case details array, if the dail maxis case number for the dail message array matches the maxis case number for the case details array then it can pull the case details from the array  
                                    
                                    'Clearing out process_dail_message
                                    process_dail_message = ""

                                    'Read dail message details
                                    EMReadScreen dail_type, 4, dail_row, 6
                                    dail_type = trim(dail_type)

                                    EMReadScreen dail_month, 8, dail_row, 11
                                    dail_month = trim(dail_month)

                                    EMReadScreen dail_msg, 61, dail_row, 20
                                    dail_msg = trim(dail_msg)

                                    ' MsgBox "dail_type: " & dail_type & " dail_month: " & dail_month & " dail_msg: " & dail_msg

                                    'Update the DAIL message array with details
                                    'To do - no need to update the maxis case number again. Remove?
                                    ' DAIL_message_array(dail_maxis_case_number_const, dail_count) = MAXIS_case_number
                                    DAIL_message_array(dail_type_const, dail_count) = dail_type
                                    DAIL_message_array(dail_month_const, dail_count) = dail_month
                                    DAIL_message_array(dail_msg_const, dail_count) = dail_msg

                                    'Activate the DAIL Messages sheet
                                    objExcel.Worksheets("DAIL Messages").Activate

                                    'Write dail details to the Excel sheet
                                    objExcel.Cells(dail_excel_row, 1).Value = DAIL_message_array(dail_maxis_case_number_const, dail_count)
                                    objExcel.Cells(dail_excel_row, 2).Value = DAIL_message_array(dail_worker_const, dail_count)
                                    objExcel.Cells(dail_excel_row, 3).Value = DAIL_message_array(dail_type_const, dail_count)
                                    objExcel.Cells(dail_excel_row, 4).Value = DAIL_message_array(dail_month_const, dail_count)
                                    objExcel.Cells(dail_excel_row, 5).Value = DAIL_message_array(dail_msg_const, dail_count)

                                    ' Msgbox "case_details_array(processable_based_on_case_const, each_case): " & case_details_array(processable_based_on_case_const, each_case)

                                    If case_details_array(processable_based_on_case_const, each_case) = False Then
                                        
                                        ' Msgbox "case_details_array(processable_based_on_case_const, each_case) = False"
                                        
                                        DAIL_message_array(renewal_month_determination_const, dail_count) = "N/A"
                                        DAIL_message_array(processable_based_on_dail_const, dail_count) = "Not Processable based on Case Details"

                                        'The dail message should not be processed due to case details
                                        process_dail_message = False

                                        'to do - do we need to add to skip list? It shouldn't ever process since it is false based on case details
                                        list_of_DAIL_messages_to_skip = list_of_DAIL_messages_to_skip & full_dail_msg & "*"

                                        'Activate the DAIL Messages sheet
                                        objExcel.Worksheets("DAIL Messages").Activate

                                        'Update the Excel sheet
                                        objExcel.Cells(dail_excel_row, 6).Value = DAIL_message_array(renewal_month_determination_const, dail_count)
                                        objExcel.Cells(dail_excel_row, 7).Value = DAIL_message_array(processable_based_on_dail_const, dail_count)
                                    
                                    ElseIf case_details_array(processable_based_on_case_const, each_case) = True Then     
                                        
                                        ' Msgbox "case_details_array(processable_based_on_case_const, each_case) = True Then " 

                                        ' Msgbox "DateAdd('m', 0, case_details_array(recertification_date_const, each_case)) " & DateAdd("m", 0, case_details_array(recertification_date_const, each_case)) 
                                        ' Msgbox "DateAdd('m', 1, footer_month_day_year) " & DateAdd("m", 1, footer_month_day_year) 
                                        ' Msgbox "DateAdd('m', 0, case_details_array(sr_report_date_const, each_case)) " & DateAdd("m", 0, case_details_array(sr_report_date_const, each_case)) 
                                        ' Msgbox "DateAdd('m', 1, footer_month_day_year) Then " & DateAdd("m", 1, footer_month_day_year)


                                        'If the recertification date or SR report date is next month, then we will check if the DAIL month matches based on the message type
                                        If DateAdd("m", 0, case_details_array(recertification_date_const, each_case)) = DateAdd("m", 1, footer_month_day_year) or DateAdd("m", 0, case_details_array(sr_report_date_const, each_case)) = DateAdd("m", 1, footer_month_day_year) Then
                                            ' Msgbox "The recertification date is equal to CM + 1 OR SR report date is equal to CM + 1"

                                            If dail_type = "CSES" Then
                                                
                                                ' Msgbox "dail type is CSES"
                                                ' Msgbox "dail_month: " & dail_month
                                                
                                                If DateAdd("m", 0, Replace(dail_month, " ", "/01/")) = DateAdd("m", 1, footer_month_day_year) Then
                                                    'To do - update language once finalized
                                                    ' Msgbox "DateAdd('m', 0, Replace(dail_month, ' ', '/01/')): " & DateAdd("m", 0, Replace(dail_month, " ", "/01/"))
                                                    ' Msgbox "DateAdd('m', 1, footer_month_day_year): " & DateAdd("m", 1, footer_month_day_year)

                                                    DAIL_message_array(processable_based_on_dail_const, dail_count) = "Not Processable due to DAIL Month & Recert/Renewal. DAIL Month is " & DateAdd("m", 0, Replace(dail_month, " ", "/01/")) & "."
                                                    objExcel.Cells(dail_excel_row, 7).Value = DAIL_message_array(processable_based_on_dail_const, dail_count)

                                                    'The dail message cannot be processed due to timing of recertification or SR report date
                                                    process_dail_message = False

                                                    'to do - do we need to add to skip list? It shouldn't ever process since it is false based on case details
                                                    list_of_DAIL_messages_to_skip = list_of_DAIL_messages_to_skip & full_dail_msg & "*"

                                                Else

                                                    'Process the CSES message here
                                                    process_dail_message = True

                                                End If
                                            ElseIf dail_type = "HIRE" Then
                                                ' MsgBox "dail type is HIRE"
                                                If DateAdd("m", 0, Replace(dail_month, " ", "/01/")) = DateAdd("m", 0, footer_month_day_year) Then
                                                    
                                                    ' Msgbox "DateAdd('m', 0, Replace(dail_month, ' ', '/01/')): " & DateAdd("m", 0, Replace(dail_month, " ", "/01/"))
                                                    ' Msgbox "DateAdd('m', 0, footer_month_day_year): " & DateAdd("m", 0, footer_month_day_year)
                                                    
                                                    'To do - update language once finalized
                                                    DAIL_message_array(processable_based_on_dail_const, dail_count) = "Not Processable due to DAIL Month & Recert/Renewal. DAIL Month is " & DateAdd("m", 0, Replace(dail_month, " ", "/01/")) & "."
                                                    objExcel.Cells(dail_excel_row, 7).Value = DAIL_message_array(processable_based_on_dail_const, dail_count)

                                                    'The dail message cannot be processed due to timing of recertification or SR report date
                                                    process_dail_message = False

                                                    'to do - do we need to add to skip list? It shouldn't ever process since it is false based on case details
                                                    list_of_DAIL_messages_to_skip = list_of_DAIL_messages_to_skip & full_dail_msg & "*"

                                                Else
                                                    'Process the HIRE message
                                                    process_dail_message = True
                                                End If
                                            Else
                                                MsgBox "something went wrong on line 809"
                                            End If

                                        Else
                                            'To do - ensure this is correct logic regarding handling based on dail months
                                            'If neither the recertification or SR report date is next month then we assume the dail message can be processed since processable based on case details is True. So set the process_dail_message to True to gather more information about the dail message
                                            process_dail_message = True
                                             
                                        End If

                                        'Make sure variables are correct
                                        ' Msgbox "process_dail_message: " & process_dail_message
                                        ' Msgbox "dail_type: " & dail_type

                                        'Process the CSES dail message
                                        If process_dail_message = True and dail_type = "CSES" Then

                                            If InStr(dail_msg, "DISB CS (TYPE 36) OF") Then

                                                ' Msgbox "InStr(dail_msg, 'DISB CS (TYPE 36) OF'): " & InStr(dail_msg, "DISB CS (TYPE 36) OF")
                                                'Enters “X” on DAIL message to open full message. 
                                                Call write_value_and_transmit("X", dail_row, 3)

                                                ' Script reads the full DAIL message so that it can process, or not process, as needed.
                                                'To do - may not need to double-check messages after fully tested
                                                EMReadScreen check_full_dail_msg_line_1, 60, 9, 5

                                                check_full_dail_msg_line_1 = trim(check_full_dail_msg_line_1)
                                                ' MsgBox check_full_dail_msg_line_1

                                                EMReadScreen check_full_dail_msg_line_2, 60, 10, 5
                                                check_full_dail_msg_line_2 = trim(check_full_dail_msg_line_2)
                                                ' MsgBox check_full_dail_msg_line_2

                                                EMReadScreen check_full_dail_msg_line_3, 60, 11, 5
                                                check_full_dail_msg_line_3 = trim(check_full_dail_msg_line_3)
                                                ' MsgBox check_full_dail_msg_line_3

                                                EMReadScreen check_full_dail_msg_line_4, 60, 12, 5
                                                check_full_dail_msg_line_4 = trim(check_full_dail_msg_line_4)
                                                ' MsgBox check_full_dail_msg_line_4

                                                check_full_dail_msg = check_full_dail_msg_line_1 & " " & check_full_dail_msg_line_2 & " " & check_full_dail_msg_line_3 & " " & check_full_dail_msg_line_4

                                                ' MsgBox check_full_dail_msg
                                                ' MsgBox full_dail_msg

                                                If check_full_dail_msg = full_dail_msg Then
                                                    ' MsgBox "They match"
                                                Else
                                                    MsgBox "Something went wrong. The DAIL messages do not match"
                                                    MsgBox "STOP THE SCRIPT HERE"
                                                End if

                                                ' Script reads information from full message, particularly the PMI number(s) listed. The script creates new variables for each PMI number.
                                                'To do - likely should validate that this is ALWAYS starting point for PMIs for Type 36
                                                EMReadScreen PMIs_line_one, 37, 10, 28
                                                ' MsgBox PMIs_line_one
                                                EMReadScreen PMIs_line_two, 60, 11, 5
                                                ' MsgBox PMIs_line_two
                                                EMReadScreen PMIs_line_three, 60, 12, 5
                                                ' MsgBox PMIs_line_three
                                                
                                                'Combine the PMIs into one string
                                                full_PMIs = replace(PMIs_line_one & PMIs_line_two & PMIs_line_three, " ", "")
                                                ' Msgbox full_PMIs
                                                'Splits the PMIs into an array
                                                PMIs_array = Split(full_PMIs, ",")

                                                'Reset the array 
                                                'To do - any issues with completely resetting array vs. adding to it each time?
                                                ReDim PMI_and_ref_nbr_array(3, 0)

                                                'Using list of PMIs in PMIs_array to update the PMI number in PMI_and_ref_nbr_array 
                                                for each_PMI = 0 to UBound(PMIs_array, 1)
                                                    ReDim Preserve PMI_and_ref_nbr_array(hh_member_count_const, each_PMI)
                                                    PMI_and_ref_nbr_array(PMI_const, each_PMI) = PMIs_array(each_PMI)
                                                    ' Msgbox "PMI_and_ref_nbr_array(PMI_const, each_PMI): " & PMI_and_ref_nbr_array(PMI_const, each_PMI)
                                                Next 

                                                'Transmit back to DAIL
                                                transmit

                                                ' Navigate to CASE/PERS to match PMIs and Ref Nbrs for checking UNEA panel
                                                ' Msgbox "Navigate to CASE/PERS"
                                                Call write_value_and_transmit("H", dail_row, 3)

                                                EMWriteScreen "PERS", 20, 69
                                                Transmit

                                                ' Iterate through CASE/PERS to match up PMI with Ref Nbr

                                                'the first member number starts at row 10
                                                pers_row = 10                  

                                                Do
                                                    'Reset reference number and PMI number so they don't carry through loop
                                                    ref_number_pers_panel = ""
                                                    pmi_number_pers_panel = ""

                                                    'reading the member number
                                                    EMReadScreen ref_number_pers_panel, 2, pers_row, 3
                                                    ref_number_pers_panel = trim(ref_number_pers_panel)

                                                    ' Msgbox "ref_number_pers_panel: " & ref_number_pers_panel

                                                    'Reading the PMI number
                                                    EMReadScreen pmi_number_pers_panel, 8, pers_row, 34  
                                                    pmi_number_pers_panel = trim(pmi_number_pers_panel)
                                                    ' Msgbox "pmi_number_pers_panel: " & PMI_number_pers_panel

                                                    for each_PMI = 0 to UBound(PMI_and_ref_nbr_array, 2)
                                                        ' Msgbox "pmi_number_pers_panel: " & PMI_number_pers_panel
                                                        ' Msgbox PMI_and_ref_nbr_array(PMI_const, each_PMI) 

                                                        If pmi_number_pers_panel = PMI_and_ref_nbr_array(PMI_const, each_PMI) Then
                                                            ' Msgbox "There is a match on the PMI"
                                                            PMI_and_ref_nbr_array(ref_nbr_const, each_PMI) = ref_number_pers_panel
                                                            ' Msgbox "PMI_and_ref_nbr_array(ref_nbr_const, each_PMI): " & PMI_and_ref_nbr_array(ref_nbr_const, each_PMI)
                                                            PMI_and_ref_nbr_array(PMI_match_found_const, each_PMI) = True
                                                            ' Msgbox "PMI_and_ref_nbr_array(PMI_match_found_const, each_PMI): " & PMI_and_ref_nbr_array(PMI_match_found_const, each_PMI)
                                                        End If
                                                    Next
                                                    
                                                    'go to the next member number - which is 3 rows down
                                                    pers_row = pers_row + 3

                                                    'if it reaches 19 - this is further down from the last member
                                                    If pers_row = 19 Then                       
                                                        'go to the next page and reset to line 10
                                                        PF8         
                                                        ' Msgbox "did last page show up"
                                                        EMReadScreen last_page_check, 21, 24, 2                          
                                                        ' Msgbox last_page_check
                                                        If last_page_check = "THIS IS THE LAST PAGE" Then Exit Do   
                                                        pers_row = 10
                                                    End If

                                                    EMReadScreen ref_number_pers_panel, 2, pers_row, 3
                                                    ' If ref_number_pers_panel = "  " Then Exit Do


                                                Loop until ref_number_pers_panel = "  "      
                                                
                                                'If there are PMIs listed on the DAIL message that do not align, then that DAIL message needs to be flagged for QI
                                                'To do - remove message boxes
                                                'To do - verify this approach works
                                                for each_individual = 0 to UBound(PMI_and_ref_nbr_array, 2)
                                                    If PMI_and_ref_nbr_array(PMI_match_found_const, each_individual) <> True Then
                                                        ' Msgbox "Some PMIs not matched"
                                                        DAIL_message_array(dail_processing_notes_const, DAIL_count) = trim(DAIL_message_array(dail_processing_notes_const, DAIL_count) & " PMI #: " & PMI_and_ref_nbr_array(PMI_const, each_individual) & " not found on case.")
                                                        ' objExcel.Cells(dail_excel_row, 8).Value = DAIL_message_array(dail_processing_notes_const, DAIL_count)
                                                    End If
                                                Next

                                                'Only check UNEA panels if ALL PMIs are matched for DAIL message and for case. There are no PMIs that did not match within the array.
                                                If InStr(DAIL_message_array(dail_processing_notes_const, DAIL_count), "not found on case") = 0 Then
                                                    'If all PMIs are found on the case, then the script will navigate directly to STAT/UNEA from CASE/PERS to verify that UNEA panels exist for CS Type 36 for each identified PMI/reference number

                                                    ' Msgbox "InStr(DAIL_message_array(dail_processing_notes_const, DAIL_count), 'not found on case') = 0: " & InStr(DAIL_message_array(dail_processing_notes_const, DAIL_count), "not found on case")
                                                    ' MsgBox "PMIs all found on case"

                                                    ' Msgbox "Moving to STAT"
                                                    EMWriteScreen "STAT", 19, 22
                                                    Call write_value_and_transmit("UNEA", 19, 69)

                                                    EmReadScreen no_unea_panels_exist, 34, 24, 2

                                                    ' MsgBox "no_unea_panels_exist: " & no_unea_panels_exist

                                                    If trim(no_unea_panels_exist) = "UNEA DOES NOT EXIST FOR ANY MEMBER" Then
                                                        'If no UNEA panels exist for the case, then the case needs to be flagged for QI
                                                        ' Msgbox "no_unea_panels_exist: " & no_unea_panels_exist
                                                        DAIL_message_array(dail_processing_notes_const, DAIL_count) = "No UNEA panels exist for any member on the case"

                                                        ' Add full dail msg to list of dail messages to skip
                                                        'To do - use check_full_dail_msg or just full_dail_msg
                                                        list_of_DAIL_messages_to_skip = list_of_DAIL_messages_to_skip & full_dail_msg & "*"

                                                        'Navigate back to DAIL
                                                        PF3

                                                        'To do - is it necessary to reset the footer month since it should update when going to CASE/CURR?
                                                        'Need to reset the footer month and footer year without interrupting script navigation in DAIL so open CASE/CURR
                                                        Msgbox "Resetting footer month and year by going to case curr. Needed?"
                                                        Call write_value_and_transmit("H", dail_row, 3)

                                                        MsgBox "update footer month and year"
                                                        'Update the footer month and year to CM/CY
                                                        EMWriteScreen MAXIS_footer_month, 20, 54
                                                        EMWriteScreen MAXIS_footer_year, 20, 57
                                                        MsgBox "Did footer month and year update?"
                                                        
                                                        'Navigate back to DAIL
                                                        PF3

                                                    ElseIf trim(no_unea_panels_exist) <> "UNEA DOES NOT EXIST FOR ANY MEMBER" Then
                                                        'There are at least some UNEA panels. Iterate through all of the PMI/reference numbers to ensure there are corresponding UNEA panels for the DISB Type
                                                        for each_individual = 0 to UBound(PMI_and_ref_nbr_array, 2)
                                                            'Navigate to first UNEA panel for member to determine if any exist
                                                            ' Msgbox "Write the PMI number to UNEA panel"
                                                            EMWriteScreen PMI_and_ref_nbr_array(ref_nbr_const, each_individual), 20, 76
                                                            Call write_value_and_transmit("01", 20, 79)

                                                            ' Msgbox "What panel did it end up on?"
                                                            'Check if no UNEA panel exists
                                                            EmReadScreen unea_panel_check, 25, 24, 2

                                                            ' Msgbox "unea_panel_check: " & unea_panel_check

                                                            If InStr(unea_panel_check, "DOES NOT EXIST") Then
                                                                'There are no UNEA panels for this HH member. Updates the processing notes for the DAIL message to reflect this
                                                                DAIL_message_array(dail_processing_notes_const, DAIL_count) = trim(DAIL_message_array(dail_processing_notes_const, DAIL_count) & " No UNEA panels exist for HH member: " & PMI_and_ref_nbr_array(ref_nbr_const, each_individual) & ".")
                                                            Else
                                                                'Read the UNEA type
                                                                EMReadScreen unea_type, 2, 5, 37
                                                                ' Msgbox "unea_type: " & unea_type
                                                                If unea_type = "36" Then
                                                                    'To do - add flagging that the panel does exist?
                                                                    'If it is a type 36 panel then it does not need to read the other panels
                                                                    ' Msgbox "unea_type: " & unea_type
                                                                    DAIL_message_array(dail_processing_notes_const, DAIL_count) = DAIL_message_array(dail_processing_notes_const, DAIL_count) & "A Type " & unea_type & " UNEA panel does exist for HH member: " & PMI_and_ref_nbr_array(ref_nbr_const, each_individual) & "."
                                                                Else
                                                                    'Check how many panels exist for the HH member
                                                                    EMReadScreen unea_panels_count, 1, 2, 78
                                                                    ' Msgbox "unea_panels_count: " & unea_panels_count
                                                                    MsgBox "Is it a number? " & IsNumeric(unea_panels_count)
                                                                    'If there are more than just a single UNEA panel, loop through them all to check for Type 36
                                                                    If unea_panels_count <> 1 Then
                                                                        'Set incrementor for do loop
                                                                        panel_count = 1

                                                                        Do
                                                                            panel_count = panel_count + 1
                                                                            ' Msgbox "panel_count: " & panel_count
                                                                            EMWriteScreen PMI_and_ref_nbr_array(ref_nbr_const, each_individual), 20, 76
                                                                            ' Msgbox "Where did it write the ref number?"
                                                                            Call write_value_and_transmit("0" & panel_count, 20, 79)

                                                                            'Read the UNEA type
                                                                            EMReadScreen unea_type, 2, 5, 37
                                                                            ' Msgbox "unea_type: " & unea_type
                                                                            If unea_type = "36" Then
                                                                                'To do - add flagging that the panel does exist?
                                                                                'If it is a type 36 panel then it does not need to read the other panels
                                                                                ' Msgbox "unea_type: " & unea_type
                                                                                DAIL_message_array(dail_processing_notes_const, DAIL_count) = DAIL_message_array(dail_processing_notes_const, DAIL_count) & "A Type " & unea_type & " UNEA panel does exist for HH member: " & PMI_and_ref_nbr_array(ref_nbr_const, each_individual) & "."
                                                                                Exit Do
                                                                            End if

                                                                            'If the loop has reached the final panel without encountering a Type 36 message, then it updates the processing notes accordingly
                                                                            If panel_count = unea_panels_count Then
                                                                                DAIL_message_array(dail_processing_notes_const, DAIL_count) = DAIL_message_array(dail_processing_notes_const, DAIL_count) & "A Type " & unea_type & " UNEA panel does not exist for HH member: " & PMI_and_ref_nbr_array(ref_nbr_const, each_individual) & "."
                                                                                Exit Do
                                                                            End If
                                                                        Loop
                                                                    End If
                                                                End If
                                                            End If
                                                        Next
                                                    End If

                                                    ' Msgbox "InStr(DAIL_message_array(dail_processing_notes_const, DAIL_count), 'does not exist'): " & InStr(DAIL_message_array(dail_processing_notes_const, DAIL_count), "does not exist")

                                                    If InStr(DAIL_message_array(dail_processing_notes_const, DAIL_count), "does not exist") Then
                                                        'There is at least one missing Type 36 UNEA panel for a HH member. The DAIL message should be added to the skip list as it cannot be deleted and requires QI review.
                                                        list_of_DAIL_messages_to_skip = list_of_DAIL_messages_to_skip & full_dail_msg & "*"
                                                        'To do - ensure this is at the correct spot
                                                        'Update the excel spreadsheet with processing notes
                                                        objExcel.Cells(dail_excel_row, 8).Value = "Message added to skip list." & DAIL_message_array(dail_processing_notes_const, DAIL_count)
                                                    ElseIf InStr(DAIL_message_array(dail_processing_notes_const, DAIL_count), "does not exist") = 0 Then
                                                        'All of the identified HH members have a corresponding Type 36 UNEA panel. The message can be deleted.
                                                        list_of_DAIL_messages_to_delete = list_of_DAIL_messages_to_delete & full_dail_msg & "*"
                                                        'To do - ensure this is at the correct spot
                                                        'Update the excel spreadsheet with processing notes
                                                        objExcel.Cells(dail_excel_row, 8).Value = "Message added to delete list." & DAIL_message_array(dail_processing_notes_const, DAIL_count)
                                                    End If


                                                Else
                                                    'There are PMIs in the DAIL message that are not on the case. Therefore, this message should be flagged for QI and added to the DAIL skip list when it is encountered again.
                                                    ' MsgBox "PMIs NOT ALL found on case"

                                                    list_of_DAIL_messages_to_skip = list_of_DAIL_messages_to_skip & full_dail_msg & "*"

                                                    'Update the excel spreadsheet with processing notes
                                                    'Ensure this is at correct spot
                                                    objExcel.Cells(dail_excel_row, 8).Value = "Message added to skip list." & DAIL_message_array(dail_processing_notes_const, DAIL_count)

                                                End If

                                                'To do - ensure this is at the correct spot
                                                'Update the excel spreadsheet with processing notes
                                                ' objExcel.Cells(dail_excel_row, 8).Value = DAIL_message_array(dail_processing_notes_const, DAIL_count)

                                                'Navigate back to the DAIL. This will reset to the top of the DAIL messages for the specific case number. Need to consider how to handle.
                                                ' MsgBox "navigate back to DAIL"
                                                PF3

                                            ElseIf InStr(dail_msg, "DISB SPOUSAL SUP (TYPE 37)") Then
                                                'Enters “X” on DAIL message to open full message. 
                                                EMWriteScreen "X", dail_row, 3
                                                ' MsgBox "Did it add X?"
                                                Transmit
                                                
                                                ' Script reads information from full message, specifically the caregiver reference number
                                                EMReadScreen caregiver_ref_nbr, 2, 10, 32
                                                ' MsgBox caregiver_ref_nbr
                                                PF3
                                                '1.	Enters “X” on DAIL message to open full message. Script reads information from full message, particularly the reference number provided. The script creates a new variable for the full DAIL message text and a variable for the reference number.
                                                ' 2.	Script PF3s out of DAIL message and navigates to STAT/UNEA from DAIL (Enters “S” on for DAIL row, then enters UNEA) 
                                                ' 3.	Script navigates to corresponding reference number’s UNEA panels.
                                                ' 4.	For identified reference number, script iterates through all UNEA panels to determine if there is a corresponding Type 37 UNEA panel 
                                                ' 5.	If there is a Type 37 UNEA panel for the reference number, script navigates back to DAIL (PF3)
                                                ' 1.	Script reads through the DAIL messages again until the full DAIL message matches accordingly
                                                ' 2.	Deletes the DAIL message (enters “D” on DAIL row)
                                                ' 3.	Updates spreadsheet with processing notes “UNEA Type 37 panel exists for Reference Number #. DAIL message deleted.”
                                                ' 6.	If there is NOT a UNEA panel for the reference number, the script navigates back to DAIL (PF3) but does NOT delete the panel
                                                ' 1.	Updates spreadsheet with processing notes “UNEA panel Type 37 missing for Reference Number #. DAIL message not deleted. Requires QI review.”
                                                ' 7.	Exits Do Loop back and moves to next row in the spreadsheet (excel_row = excel_row + 1)

                                                ' MsgBox "DISB SPOUSAL SUP (TYPE 37): " & dail_msg
                                            ElseIf InStr(dail_msg, "DISB CS ARREARS (TYPE 39) OF") Then
                                                '1.	Enters “X” on DAIL message to open full message. Script reads information from full message, particularly the PMI number(s) listed. The script creates new variables for each PMI number.

                                                'To do - once it is working, can use Call write value and transmit
                                                '1.	Enters “X” on DAIL message to open full message. 
                                                EMWriteScreen "X", dail_row, 3
                                                ' MsgBox "Did it add X?"
                                                Transmit
                                                
                                                ' Script reads information from full message, particularly the PMI number(s) listed. The script creates new variables for each PMI number.
                                                EMReadScreen PMIs_line_one, 30, 10, 35
                                                ' MsgBox PMIs_line_one
                                                EMReadScreen PMIs_line_two, 60, 11, 5
                                                ' MsgBox PMIs_line_two
                                                EMReadScreen PMIs_line_three, 60, 12, 5
                                                ' MsgBox PMIs_line_three
                                                
                                                'Combine the PMIs into one string
                                                full_PMIs = replace(PMIs_line_one & PMIs_line_two & PMIs_line_three, " ", "")
                                                ' MsgBox full_PMIs
                                                'Splits the PMIs into an array
                                                PMIs_array = Split(full_PMIs, ",")

                                                ' For each PMI in PMIs_array
                                                '     MsgBox PMI 
                                                ' Next

                                                'Backs out of full DAIL message to DAIL
                                                PF3

                                                ' 2.	Script PF3s out of DAIL message and navigates to STAT/UNEA from DAIL (Enters “S” on for DAIL row, then enters UNEA) 
                                                ' 3.	Script reads through each household member’s UNEA panels until each PMI is matched
                                                ' 4.	For each identified PMI number, determines if there is a corresponding Type 39 UNEA panel 
                                                ' 5.	If there is a Type 39 UNEA panel for every PMI, script navigates back to DAIL (PF3)
                                                ' 1.	Script reads through the messages again until the full DAIL message matches accordingly
                                                ' 2.	Deletes the DAIL message (enters “D” on DAIL row)
                                                ' 3.	Updates spreadsheet with processing notes “UNEA Type 39 panel exists for all PMI(s). DAIL message deleted.”
                                                ' 6.	If there is NOT a UNEA panel for every PMI, script navigates back to DAIL but does NOT delete the panel
                                                ' 1.	Updates spreadsheet with processing notes “UNEA panel Type 39 missing for PMI(S): #####. DAIL message not deleted. Requires QI review.”
                                                ' 7.	Exits Do Loop back and moves to next row in the spreadsheet (excel_row = excel_row + 1)

                                                ' MsgBox "DISB CS ARREARS (TYPE 39) OF: " & dail_msg
                                            ElseIf InStr(dail_msg, "DISB SPOUSAL SUP ARREARS (TYPE 40) OF") Then
                                                'Enters “X” on DAIL message to open full message. 
                                                EMWriteScreen "X", dail_row, 3
                                                ' MsgBox "Did it add X?"
                                                Transmit
                                                
                                                ' Script reads information from full message, specifically the caregiver reference number
                                                EMReadScreen caregiver_ref_nbr, 2, 10, 32
                                                ' MsgBox caregiver_ref_nbr
                                                PF3
                                                '1.	Enters “X” on DAIL message to open full message. Script reads information from full message, particularly the reference number provided. The script creates a new variable for the full DAIL message text and a variable for the reference number.
                                                ' 2.	Script PF3s out of DAIL message and navigates to STAT/UNEA from DAIL (Enters “S” on for DAIL row, then enters UNEA) 
                                                ' 3.	Script navigates to corresponding reference number’s UNEA panels.
                                                ' 4.	For identified reference number, script iterates through all UNEA panels to determine if there is a corresponding Type 40 UNEA panel 
                                                ' 5.	If there is a Type 40 UNEA panel for the reference number, script navigates back to DAIL (PF3)
                                                ' 1.	Script reads through the DAIL messages again until the full DAIL message matches accordingly
                                                ' 2.	Deletes the DAIL message (enters “D” on DAIL row)
                                                ' 3.	Updates spreadsheet with processing notes “UNEA Type 40 panel exists for Reference Number #. DAIL message deleted.”
                                                ' 6.	If there is NOT a UNEA panel for the reference number, the script navigates back to DAIL (PF3) but does NOT delete the panel
                                                ' 1.	Updates spreadsheet with processing notes “UNEA panel Type 40 missing for Reference Number #. DAIL message not deleted. Requires QI review.”
                                                ' 7.	Exits Do Loop back and moves to next row in the spreadsheet (excel_row = excel_row + 1)

                                                ' MsgBox "DISB SPOUSAL SUP ARREARS (TYPE 40) OF: " & dail_msg
                                            ElseIf InStr(dail_msg, "CS REPORTED: NEW EMPLOYER FOR CAREGIVER REF NBR:") Then

                                                'Enters “X” on DAIL message to open full message. 
                                                EMWriteScreen "X", dail_row, 3
                                                ' MsgBox "Did it add X?"
                                                Transmit
                                                
                                                ' Script reads information from full message, specifically the caregiver reference number
                                                EMReadScreen caregiver_ref_nbr, 2, 9, 54
                                                ' MsgBox caregiver_ref_nbr
                                                
                                                'Then it reads the employer name up to two extra lines just in case
                                                EMReadScreen employer_name_line_one, 8, 9, 57
                                                EMReadScreen employer_name_line_two, 60, 10, 5
                                                EMReadScreen employer_name_line_three, 60, 11, 5
                                                
                                                ' Combine the employer name lines together to form the full nameCombine the PMIs into one string
                                                full_employer_name = employer_name_line_one & employer_name_line_two & employer_name_line_three
                                                
                                                ' MsgBox full_employer_name

                                                PF3
                                                
                                                '1.	Enters “X” on DAIL message to open full message. Script reads information from full message. The script creates a new variable for the full DAIL message text, a variable for the reference number, and a variable for the full employer name.
                                                ' 2.	Script PF3s out of DAIL message and navigates to STAT/JOBS from DAIL (Enters “S” on for DAIL row, then enters JOBS) 
                                                ' 3.	Script navigates to corresponding reference number’s JOBS panels.
                                                ' 4.	For identified reference number, script iterates through all JOBS panels to determine if there is a matching employer name
                                                ' 1.	Consider handling for an approximate match vs exact match 
                                                ' 2.	Dialog box with list of employer names against the CSES message to choose manually?
                                                ' 5.	If there is a matching JOBS panel for the reference number, script navigates back to DAIL (PF3)
                                                ' 1.	Script reads through the DAIL messages again until the full DAIL message matches accordingly
                                                ' 2.	Deletes the DAIL message (enters “D” on DAIL row)
                                                ' 3.	Updates spreadsheet with processing notes “JOBS panel exists for Reference Number #. DAIL message deleted.”
                                                ' 4.	Script CASE/NOTEs information about deleting the DAIL message
                                                ' 6.	If there is NO matching JOBS panel for the reference number, the script creates a new JOBS panel
                                                ' 1.	Adds new JOBS panel for the reference number
                                                ' 2.	Use “Other” for JOBS panel and fill in rest (blank?)
                                                ' 3.	Navigate back to DAIL (PF3)
                                                ' 4.	Deletes the DAIL message (enters “D” on DAIL row)
                                                ' 5.	Updates spreadsheet with processing notes “Created new JOBS panel for employer [name] and CASE/NOTEd. DAIL message deleted.”
                                                ' 6.	Script CASE/NOTEs information about deleting the DAIL message
                                                ' 7.	Exits Do Loop back and moves to next row in the spreadsheet (excel_row = excel_row + 1)

                                                ' MsgBox "CS REPORTED: NEW EMPLOYER FOR CAREGIVER REF NBR: " & dail_msg
                                            ElseIf InStr(dail_msg, "REPORTED: CHILD REF NBR:") Then
                                                '1.	No action on these, simply note in spreadsheet that QI team to review
                                                ' MsgBox "REPORTED: CHILD REF NBR:" & dail_msg
                                            Else
                                                msgbox "Something went wrong - line 1248"
                                            End If


                                        ElseIf process_dail_message = True and dail_type = "HIRE" Then
                                            'Update here

                                            If InStr(dail_msg, "NDNH MEMB") Then
                                                'Add logic here
                                                MsgBox "NDNH MEMB: " & dail_msg
                                            ElseIf InStr(dail_msg, "NEW JOB DETAILS FOR SSN") Then
                                                'Add logic here
                                                MsgBox "NEW JOB DETAILS FOR SSN: " & dail_msg
                                            ElseIf InStr(dail_msg, "SDNH NEW JOB DETAILS") Then
                                                'Add logic here
                                                MsgBox "SDNH NEW JOB DETAILS: " & dail_msg
                                            ElseIf InStr(dail_msg, "JOB DETAILS FOR  ") Then
                                                'Add logic here
                                                MsgBox "JOB DETAILS FOR  " & dail_msg
                                            Else
                                                msgbox "Something went wrong - line 1268"
                                            End If
                                        Else
                                            MsgBox "Something went wrong = 1269"
                                            MsgBox "process_dail_message: " & process_dail_message
                                            MsgBox "dail_type: " & dail_type
                                            MsgBox "Stop here"
                                        End If

                                        
                                        


                                    End If

                                    'Increment the dail_excel_row so that data isn't overwritten
                                    dail_excel_row = dail_excel_row + 1
                                    
                                    'Increment dail_count for the dail array
                                    dail_count = dail_count + 1

                                    ' dail_excel_row = dail_excel_row + 1
                                End If 
                                'To do - validate placement of dail count incrementor
                                'To do - I think it is in wrong spot. Erroring out on line 680. The dail count is incrementing before it is redimmed so when it is called at higher dail count it errors.
                                ' dail_count = dail_count + 1
                            Next

                        Else
                            'Add handling for messages that are not meeting any criteria. May not be necessary but have this just in case
                        End If
                            
                    End If
                Else
                    'To do - add handling for cases that are not on valid case numbers list, just set processable to false and include processing note that it is likely out of county or privileged?
                
                End If
                        
            
            Else
                'To do - probably can remove this ELSE since it will just move to next DAIL message without doing anything
                'If dail_type is not CSES or HIRE then it is out of scope and there is no need to evaluate it
                ' MsgBox "NOT CSES OR HIRE"

            End If

            'Add in handling to determine DAIL details

            'Increment the stats counter
            stats_counter = stats_counter + 1
            
            ' MsgBox "dail increases by 1"
            'To do - validate placement of dail_row incrementor based on DAIL message processing outcome
            'To do - should dail_row + 1 be within each of the options (delete list, skip list, new list)
            dail_row = dail_row + 1

            ' MsgBox "dail_row = " & dail_row
            
            'TO DO - this is from DAIL decimator. Appears to handle for NAT errors. Is it needed?
            'EMReadScreen message_error, 11, 24, 2		'Cases can also NAT out for whatever reason if the no messages instruction comes up.
            'If message_error = "NO MESSAGES" then exit do

            '...going to the next page if necessary
            EMReadScreen next_dail_check, 4, dail_row, 4
            If trim(next_dail_check) = "" then
                PF8
                EMReadScreen last_page_check, 21, 24, 2
                'DAIL/PICK when searching for specific DAIL types has message check of NO MESSAGES TYPE vs. NO MESSAGES WORK (for ALL DAIL/PICK selection).
                If last_page_check = "THIS IS THE LAST PAGE" or last_page_check = "NO MESSAGES TYPE" then
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

'Update Stats Info
'Activate the stats sheet
objExcel.Worksheets("Stats").Activate
objExcel.Cells(1, 2).Value = STATS_counter
objExcel.Cells(2, 2).Value = STATS_manualtime
objExcel.Cells(3, 2).Value = STATS_counter * STATS_manualtime
objExcel.Cells(4, 2).Value = timer - start_time
objExcel.Cells(5, 2).Value = ((STATS_counter * STATS_manualtime) - (timer - start_time)) / 60