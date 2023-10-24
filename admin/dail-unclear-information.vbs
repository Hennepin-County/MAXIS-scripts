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
    
    'Creating initial string for tracking all case numbers
    valid_case_numbers_list = "*"

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
			'MsgBox "Are we waiting?"
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

    ' MsgBox valid_case_numbers_list

    'Navigates to DAIL to pull DAIL messages
    MAXIS_case_number = ""
    CALL navigate_to_MAXIS_screen("DAIL", "PICK")
    EMWriteScreen "_", 7, 39    'blank out ALL selection
    'Selects CSES DAIL Type based on dialog selection
    If CSES_messages = 1 Then EMWriteScreen "X", 10, 39    'Select CSES DAIL Type
    'Selects INFO (HIRE) DAIL Type based on dialog selection
    If HIRE_messages = 1 Then EMWriteScreen "X", 13, 39
    transmit

    'To do - verify placement of this
    list_of_all_case_numbers = "~"
    'To do - think about handling for a situation where first case in DAIL is privileged so there wouldn't be any prior examples

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
            ' MsgBox "Did it move one row?"
            'To do - verify if variables are resetting properly every do loop
            ' dail_type = ""
            ' dail_msg = ""

            'Determining if there is a new case number...
            EMReadScreen new_case, 8, dail_row, 63
            new_case = trim(new_case)
            IF new_case <> "CASE NBR" THEN 
                'If there is NOT a new case number, the script will read the DAIL type, month, year, and message...
                Call write_value_and_transmit("T", dail_row, 3)
            ELSEIF new_case = "CASE NBR" THEN
                '...if the script does find that there is a new case number (indicated by "CASE NBR"), it will write a "T" in the next row and transmit, bringing that case number to the top of your DAIL
                Call write_value_and_transmit("T", dail_row + 1, 3)
            End if

            dail_row = 6  'resetting the DAIL row '

            'Determines the DAIL Type
            EMReadScreen dail_type, 4, dail_row, 6
            dail_type = trim(dail_type)

            If dail_type = "CSES" OR dail_type = "HIRE" Then
                'Read the MAXIS Case Number, if it is a new case number then pull case details. If it is not a new case number, then do not pull new case details.
                EMReadScreen MAXIS_case_number, 8, dail_row - 1, 73
                MAXIS_case_number = trim(MAXIS_case_number)

                'If the case is in the valid_case_numbers_list, then it can be evaluated. If it is NOT in the valid_case_numbers_list then it is likely privileged or out of county so it will be skipped
                If InStr(valid_case_numbers_list, "*" & MAXIS_case_number & "*") Then
                    'If the MAXIS case number is NOT in the list of all case numbers, then it is a new case number and the script will gather case details
                    If Instr(list_of_all_case_numbers, "~" & MAXIS_case_number & "~") = 0 Then
                        'Redim the case details array and add to array
                        ReDim Preserve case_details_array(case_excel_row_const, case_count)
                        case_details_array(case_maxis_case_number_const, case_count) = MAXIS_case_number
                        case_details_array(case_worker_const, case_count) = worker

                        'Basically needs to be If Priv then this, elseif out-of-county then this, all other cases move forward
                        'Must add handling for privileged case
                        'Priv - determine priv case number and case number preceding priv case to then go back to, need to re-select DAIL/PICK correctly

                        
                
                        'Navigating to CASE/CURR and pulling case details
                        'Add new case number to list of all case numbers
                        list_of_all_case_numbers = list_of_all_case_numbers & MAXIS_case_number & "~"

                        'Navigate to CASE/CURR to pull case details 
                        Call write_value_and_transmit("H", dail_row, 3)
                        EMReadScreen priv_check, 24, 24, 2

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

                            'Function above (determine_program_and_case_status_from_CASE_CURR) changes dail_row to equal 4 so need to reset it.
                            dail_row = 6
                            

                            'If SNAP is not active, then the case is out-of-scope for Unclear Information
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
                                    ' MsgBox "other programs ARE present"
                                    case_details_array(other_programs_present_const, case_count) = True
                                    case_details_array(processable_based_on_case_const, case_count) = False
                            Else
                                case_details_array(other_programs_present_const, case_count) = False
                            End if

                            If snap_status = "ACTIVE" then
                                'To do - check if background check is needed, may break connection to DAIL
                                ' Call MAXIS_background_check
                                ' Call navigate_to_MAXIS_screen("ELIG", "FS  ")
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
                                        'To do - since pulling cases for REPT/ACTV, this may never trigger but using message box just in case
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
                                        'Change the footer month and year back to CM/CY
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
                                EMWriteScreen MAXIS_footer_month, 20, 54
                                EMWriteScreen MAXIS_footer_year, 20, 57
                                ' MsgBox "did footer month year update?"
                            End If 

                            'Determine if processable_based_on_case is True (case is within scope, MAY act on DAIL message) or if processable_based_on_case is false (case is outside of scope, WILL NOT act on DAIL message)
                            ' MsgBox "case_details_array(snap_status_const, case_count): " & case_details_array(snap_status_const, case_count)
                            ' MsgBox "case_details_array(other_programs_present_const, case_count): " & case_details_array(other_programs_present_const, case_count)
                            ' MsgBox "case_details_array(reporting_status_const, case_count): " & case_details_array(reporting_status_const, case_count)
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

                        ' MsgBox "place the footer month update here"

                        ' MsgBox "determine program and case status. Did it work?"
                        'Return to DAIL by PF3
                        PF3
                      
                        ' MsgBox "Where did the PF3 move to?"
                        'Increment the case_count for updating the array
                        case_count = case_count + 1
                        'Subtract one from dail_row so that the dail_row restarts evaluation of cases now with case details
                        ' MsgBox "subtract 1 from dail?"
                        dail_row = dail_row - 1
                        ' MsgBox "After subtraction, dail_row = " & dail_row
                    
                    Else
                        'If the MAXIS case number IS in the list of all case numbers, then it is not a new case number and no case details need to be gathered. It can work off the already collected case details.
                        'Gather details on DAIL message, should capture DAIL details in spreadsheet even if ultimately not actionable
                        
                        ' MsgBox "MAXIS_case_number: " & MAXIS_case_number

                        ReDim Preserve DAIL_message_array(DAIL_excel_row_const, dail_count)
                        DAIL_message_array(dail_maxis_case_number_const, DAIL_count) = MAXIS_case_number
                        DAIL_message_array(dail_worker_const, DAIL_count) = worker

                        'Use for next loop to match the individual DAIL message to the corresponding array item of matching Case Details
                        for each_case = 0 to UBound(case_details_array, 2)
                            If DAIL_message_array(dail_maxis_case_number_const, dail_count) = case_details_array(case_maxis_case_number_const, each_case) Then
                                'Clearing out process_dail_message
                                process_dail_message = ""

                                EMReadScreen dail_type, 4, dail_row, 6
                                dail_type = trim(dail_type)

                                EMReadScreen dail_month, 8, dail_row, 11
                                dail_month = trim(dail_month)

                                EMReadScreen dail_msg, 61, dail_row, 20
                                dail_msg = trim(dail_msg)
                                If InStr(dail_msg, "INFC") then 
                                    INFC_dail_message = True
                                Else
                                    INFC_dail_message = False
                                End If

                                ' MsgBox "This is the dail_msg: " & dail_msg
                                ' MsgBox infc_dail_message  

                                If INFC_dail_message = False Then
                                    DAIL_message_array(dail_maxis_case_number_const, dail_count) = MAXIS_case_number
                                    DAIL_message_array(dail_type_const, dail_count) = dail_type
                                    DAIL_message_array(dail_month_const, dail_count) = dail_month
                                    DAIL_message_array(dail_msg_const, dail_count) = dail_msg

                                    'Activate the DAIL Messages sheet
                                    objExcel.Worksheets("DAIL Messages").Activate

                                    objExcel.Cells(dail_excel_row, 1).Value = DAIL_message_array(dail_maxis_case_number_const, dail_count)
                                    objExcel.Cells(dail_excel_row, 2).Value = DAIL_message_array(dail_worker_const, dail_count)
                                    objExcel.Cells(dail_excel_row, 3).Value = DAIL_message_array(dail_type_const, dail_count)
                                    objExcel.Cells(dail_excel_row, 4).Value = DAIL_message_array(dail_month_const, dail_count)
                                    objExcel.Cells(dail_excel_row, 5).Value = DAIL_message_array(dail_msg_const, dail_count)

                                    If case_details_array(processable_based_on_case_const, each_case) = False Then
                                        ' MsgBox "not processable based on case details"
                                        DAIL_message_array(renewal_month_determination_const, dail_count) = "N/A"
                                        DAIL_message_array(processable_based_on_dail_const, dail_count) = "Not Processable based on Case Details"

                                        'Activate the DAIL Messages sheet
                                        objExcel.Worksheets("DAIL Messages").Activate

                                        objExcel.Cells(dail_excel_row, 6).Value = DAIL_message_array(renewal_month_determination_const, dail_count)
                                        objExcel.Cells(dail_excel_row, 7).Value = DAIL_message_array(processable_based_on_dail_const, dail_count)
                                    
                                    ElseIf case_details_array(processable_based_on_case_const, each_case) = True Then     
                                        
                                        ' MsgBox "Processable = true"
                                        
                                        'If the recertification date or SR report date is next month, then we will check if the DAIL month matches based on the message type
                                        If DateAdd("m", 0, case_details_array(recertification_date_const, each_case)) = DateAdd("m", 1, footer_month_day_year) or DateAdd("m", 0, case_details_array(sr_report_date_const, each_case)) = DateAdd("m", 1, footer_month_day_year) Then
                                            If dail_type = "CSES" Then
                                                ' MsgBox "dail type is CSES"
                                                If DateAdd("m", 0, Replace(dail_month, " ", "/01/")) = DateAdd("m", 1, footer_month_day_year) Then
                                                    'To do - update language once finalized
                                                    DAIL_message_array(processable_based_on_dail_const, dail_count) = "Not Processable due to DAIL Month & Recert/Renewal. DAIL Month is " & DateAdd("m", 0, Replace(dail_month, " ", "/01/")) & "."
                                                    objExcel.Cells(dail_excel_row, 7).Value = DAIL_message_array(processable_based_on_dail_const, dail_count)
                                                    process_dail_message = False
                                                Else
                                                    'Process the CSES message here
                                                    process_dail_message = True
                                                End If
                                            ElseIf dail_type = "HIRE" Then
                                                ' MsgBox "dail type is HIRE"
                                                If DateAdd("m", 0, Replace(dail_month, " ", "/01/")) = DateAdd("m", 0, footer_month_day_year) Then
                                                    'To do - update language once finalized
                                                    DAIL_message_array(processable_based_on_dail_const, dail_count) = "Not Processable due to DAIL Month & Recert/Renewal. DAIL Month is " & DateAdd("m", 0, Replace(dail_month, " ", "/01/")) & "."
                                                    objExcel.Cells(dail_excel_row, 7).Value = DAIL_message_array(processable_based_on_dail_const, dail_count)
                                                    process_dail_message = False
                                                Else
                                                    'Process the HIRE message
                                                    process_dail_message = True
                                                End If
                                            End If

                                        End If

                                        'Process CSES message
                                        ' DISB CS (TYPE 36) OF [$X.XX] FOR [#] CHILD(REN) ISSUED 
                                        ' REPLACED [DD/MM/YY] DISB CS (TYPE 36) OF [$X.XX] FOR 
                                        ' DISB SPOUSAL SUP (TYPE 37) OF [$X.XX] ISSUED ON 
                                        ' DISB CS ARREARS (TYPE 39) OF [$X.XX] FOR [#] CHILD(REN) 
                                        ' REPLACED [MM/DD/YY] DISB CS ARREARS (TYPE 39) OF
                                        ' DISB SPOUSAL SUP ARREARS (TYPE 40) OF [$X.XX] ISSUED 
                                        ' CS REPORTED: NEW EMPLOYER FOR CAREGIVER REF NBR: [##] [EMPLOYER NAME]
                                        ' REPORTED: CHILD REF NBR: [##] NO LONGER RESIDES WITH CAREGIVER (unable to find example)

                                        If process_dail_message = True and dail_type = "CSES" Then

                                            If InStr(dail_msg, "DISB CS (TYPE 36) OF") Then

                                                '1.	Enters “X” on DAIL message to open full message. 
                                                'To do - once it is working, can use Call write value and transmit
                                                EMWriteScreen "X", dail_row, 3
                                                ' MsgBox "Did it add X?"
                                                Transmit

                                                ' Script reads the full DAIL message so that it can process, or not process, as needed.
                                                EMReadScreen full_dail_msg_line_1, 60, 9, 5

                                                full_dail_msg_line_1 = trim(full_dail_msg_line_1)
                                                MsgBox full_dail_msg_line_1

                                                EMReadScreen full_dail_msg_line_2, 60, 10, 5
                                                full_dail_msg_line_2 = trim(full_dail_msg_line_2)
                                                MsgBox full_dail_msg_line_2

                                                EMReadScreen full_dail_msg_line_3, 60, 11, 5
                                                full_dail_msg_line_3 = trim(full_dail_msg_line_3)
                                                MsgBox full_dail_msg_line_3

                                                EMReadScreen full_dail_msg_line_4, 60, 12, 5
                                                full_dail_msg_line_4 = trim(full_dail_msg_line_4)
                                                MsgBox full_dail_msg_line_4

                                                full_dail_msg = full_dail_msg_line_1 & " " & full_dail_msg_line_2 & " " & full_dail_msg_line_3 & " " & full_dail_msg_line_4

                                                MsgBox full_dail_msg

                                                ' Script reads information from full message, particularly the PMI number(s) listed. The script creates new variables for each PMI number.
                                                EMReadScreen PMIs_line_one, 37, 10, 28
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

                                                'Reset the array 
                                                ReDim PMI_and_ref_nbr_array(3, 0)

                                                'Using list of PMIs in PMIs_array to update the PMI number in PMI_and_ref_nbr_array 
                                                for each_PMI = 0 to UBound(PMIs_array, 1)
                                                    ReDim Preserve PMI_and_ref_nbr_array(hh_member_count_const, each_PMI)
                                                    PMI_and_ref_nbr_array(PMI_const, each_PMI) = PMIs_array(each_PMI)
                                                Next 

                                                'Transmit back to DAIL
                                                transmit

                                                ' Navigate to CASE/PERS to match PMIs and Ref Nbrs for checking UNEA panel
                                                'To do - this will break tie to specific DAIL message, think about how to navigate back
                                                
                                                EMWriteScreen "H", dail_row, 3
                                                Transmit

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

                                                    MsgBox "ref_number_pers_panel: " & ref_number_pers_panel

                                                    'Reading the PMI number
                                                    EMReadScreen pmi_number_pers_panel, 8, pers_row, 34  
                                                    pmi_number_pers_panel = trim(pmi_number_pers_panel)
                                                    MsgBox "pmi_number_pers_panel: " & PMI_number_pers_panel

                                                    for each_PMI = 0 to UBound(PMI_and_ref_nbr_array, 2)
                                                        MsgBox "pmi_number_pers_panel: " & PMI_number_pers_panel
                                                        MsgBox PMI_and_ref_nbr_array(PMI_const, each_PMI) 

                                                        If pmi_number_pers_panel = PMI_and_ref_nbr_array(PMI_const, each_PMI) Then
                                                            PMI_and_ref_nbr_array(ref_nbr_const, each_PMI) = ref_number_pers_panel
                                                            MsgBox PMI_and_ref_nbr_array(ref_nbr_const, each_PMI)
                                                            PMI_and_ref_nbr_array(PMI_match_found_const, each_PMI) = True
                                                            MsgBox PMI_and_ref_nbr_array(PMI_match_found_const, each_PMI)
                                                        End If
                                                    Next
                                                    
                                                    'go to the next member number - which is 3 rows down
                                                    pers_row = pers_row + 3

                                                    'if it reaches 19 - this is further down from the last member
                                                    If pers_row = 19 Then                       
                                                        'go to the next page and reset to line 10
                                                        PF8                                     
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
                                                        MsgBox "Some PMIs not matched"
                                                        DAIL_message_array(dail_processing_notes_const, DAIL_count) = DAIL_message_array(dail_processing_notes_const, DAIL_count) & "PMI #: " & PMI_and_ref_nbr_array(PMI_const, each_individual) & " not found on case."
                                                        objExcel.Cells(dail_excel_row, 8).Value = DAIL_message_array(dail_processing_notes_const, DAIL_count)
                                                    End If
                                                Next

                                                'Only check UNEA panels if ALL PMIs are matched for DAIL message and for case
                                                If InStr(DAIL_message_array(dail_processing_notes_const, DAIL_count), "not found on case.") = 0 Then
                                                    'If all PMIs are found on the case, then the script will navigate directly to STAT/UNEA from CASE/PERS to verify that UNEA panels exist for CS Type 36
                                                    MsgBox "PMIs all found on case"

                                                    EMWriteScreen "STAT", 19, 22
                                                    Call write_value_and_transmit("UNEA", 19, 69)


                                                Else
                                                    MsgBox "PMIs NOT ALL found on case"

                                                    'To do - add functionality to go to the next message and skip the message because not processable

                                                End If

                                                'this is the end of the list

                                                ' for each_PMI = 0 to UBound(case_pers_ref_nbr_and_pmi_array, 2)
                                                '     for each PMI in PMIs_array
                                                '         If PMI = case_pers_ref_nbr_and_pmi_array(case_curr_PMI_const, each_PMI) Then
                                                '             msgbox "PMIs_array: " & PMI
                                                '             MsgBox "PMI ref nbr array: " & PMI_with_ref_nbrs_array(PMI_number_const, each_case)
                                                '             MsgBox "There is a match"
                                                '             case_pers_ref_nbr_and_pmi_array(case_curr_PMI_match_found_const, each_PMI) = True
                                                '             MsgBox case_pers_ref_nbr_and_pmi_array(case_curr_PMI_match_found_const, each_PMI)
                                                '         End If
                                                '     Next
                                                ' Next 

                                                'Navigate back to the DAIL. This will reset to the top of the DAIL messages for the specific case number. Need to consider how to handle.
                                                MsgBox "navigate back to DAIL"
                                                PF3


                                                'Delete after testing

                                                ' 'Navigate to STAT/UNEA to determine if UNEA panels exist for the case
                                                ' 'To do - once it is working, can use Call write value and transmit
                                                ' EMWriteScreen "S", dail_row, 3
                                                ' ' MsgBox "Did it add X?"
                                                ' Transmit
                                                
                                                ' EMWriteScreen "UNEA", 20, 71
                                                ' Transmit

                                                ' ' Check if no UNEA panels exist for the case, in which this makes it easy to determine whether to process the DAIL message
                                                ' EMReadScreen no_unea_panels_check, 34, 24, 2
                                                ' If no_unea_panels_check = "UNEA DOES NOT EXIST FOR ANY MEMBER" Then
                                                '     'Add functionality here to handle this situation




                                                ' 2.	Script PF3s out of DAIL message and navigates to STAT/UNEA from DAIL (Enters “S” on for DAIL row, then enters UNEA) 
                                                ' 3.	Script reads through each household member’s UNEA panels until each PMI is matched
                                                ' 4.	For each identified PMI number, determines if there is a corresponding Type 36 UNEA panel 
                                                ' 5.	If there is a Type 36 UNEA panel for every PMI, script navigates back to DAIL (PF3)
                                                ' 1.	Script reads through the messages again until the full DAIL message matches accordingly
                                                ' 2.	Deletes the DAIL message (enters “D” on DAIL row)
                                                ' 3.	Updates spreadsheet with processing notes “UNEA Type 36 panel exists for every PMI. DAIL message deleted.”
                                                ' 6.	If there is NOT a UNEA panel for every PMI, script navigates back to DAIL but does NOT delete the panel
                                                ' 1.	Updates spreadsheet with processing notes “UNEA panel TYPE 36 missing for PMI(S): #####. DAIL message not deleted. Requires QI review.”
                                                ' 7.	Exits Do Loop back and moves to next row in the spreadsheet (excel_row = excel_row + 1)

                                                ' MsgBox "DISB CS (TYPE 36) OF: " & dail_msg
                                            ElseIf InStr(dail_msg, "DISB SPOUSAL SUP (TYPE 37)") Then
                                                'Enters “X” on DAIL message to open full message. 
                                                EMWriteScreen "X", dail_row, 3
                                                ' MsgBox "Did it add X?"
                                                Transmit
                                                
                                                ' Script reads information from full message, specifically the caregiver reference number
                                                EMReadScreen caregiver_ref_nbr, 2, 10, 32
                                                MsgBox caregiver_ref_nbr
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
                                                MsgBox PMIs_line_one
                                                EMReadScreen PMIs_line_two, 60, 11, 5
                                                MsgBox PMIs_line_two
                                                EMReadScreen PMIs_line_three, 60, 12, 5
                                                MsgBox PMIs_line_three
                                                
                                                'Combine the PMIs into one string
                                                full_PMIs = replace(PMIs_line_one & PMIs_line_two & PMIs_line_three, " ", "")
                                                ' MsgBox full_PMIs
                                                'Splits the PMIs into an array
                                                PMIs_array = Split(full_PMIs, ",")

                                                For each PMI in PMIs_array
                                                    MsgBox PMI 
                                                Next

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
                                                MsgBox caregiver_ref_nbr
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
                                                MsgBox caregiver_ref_nbr
                                                
                                                'Then it reads the employer name up to two extra lines just in case
                                                EMReadScreen employer_name_line_one, 8, 9, 57
                                                EMReadScreen employer_name_line_two, 60, 10, 5
                                                EMReadScreen employer_name_line_three, 60, 11, 5
                                                
                                                ' Combine the employer name lines together to form the full nameCombine the PMIs into one string
                                                full_employer_name = employer_name_line_one & employer_name_line_two & employer_name_line_three
                                                
                                                MsgBox full_employer_name

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
                                                msgbox "Something went wrong"
                                            End If


                                        End if

                                        'Process HIRE message
                                        ' NDNH MEMB [##] NEW [STATE ABBREV] JOB DETAILS 
                                        ' NEW JOB DETAILS FOR SSN [###-##-####] ([LAST NAME, FIRST INITIAL])
                                        ' SDNH NEW JOB DETAILS FOR MEMB [##] 
                                        ' [SSN #] NEW [STATE ABBREV] JOB DETAILS FOR  [LAST NAME,FIRST INITIAL]


                                        If process_dail_message = True and dail_type = "HIRE" Then
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
                                            End If
                                        End If

                                        
                                        


                                    End If

                                    dail_excel_row = dail_excel_row + 1
                                    

                                End If
                                ' dail_excel_row = dail_excel_row + 1
                            End If 
                        next


                        dail_count = dail_count + 1

                        ' a.	Is case processable based on case details? 
                            ' i.	Yes (Processable based on Case Details = True)
                                ' 1.	Is DAIL message processable based on DAIL details?
                                    ' a.	Yes (Processable based on DAIL Details = True)
                                        ' i.	Process DAIL according to process
                                    ' b.	No (Processable based on DAIL Details = False)
                                        ' i.	Move to next DAIL row and restart loop
                            ' ii.	No (Processable based on Case Details = False)
                                ' 1.	Still capture DAIL details (month, message, etc.) but then indicate that it is not processable based on case details
                            
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