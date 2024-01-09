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
objExcel.Cells(1, 6).Value = "Full DAIL Message"
' objExcel.Cells(1, 6).Value = "Renewal Month Determination"
' objExcel.Cells(1, 7).Value = "Processable based on DAIL"
objExcel.Cells(1, 7).Value = "Processing Notes for DAIL Message"

FOR i = 1 to 7		'formatting the cells'
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
objExcel.Cells(1, 4).Value = "SNAP Only"
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

'Setting counters at 0
STATS_counter = STATS_counter - 1
not_processable_msg_count = 0
dail_msg_deleted_count = 0
QI_flagged_msg_count = 0

'Enters info about runtime for the benefit of folks using the script
'To do - update to reflect actual stats needed/wanted
objExcel.Cells(1, 1).Value = "Cases Evaluated:"
objExcel.Cells(2, 1).Value = "Evaluated DAIL Messages:"
objExcel.Cells(3, 1).Value = "Unprocessable DAIL Messages:"
objExcel.Cells(4, 1).Value = "Deleted DAIL Messages:"
objExcel.Cells(5, 1).Value = "QI Flagged Messages:"
objExcel.Cells(6, 1).Value = "Script run time (in seconds):"
objExcel.Cells(7, 1).Value = "Estimated time savings by using script (in minutes):"


FOR i = 1 to 7		'formatting the cells'
    objExcel.Cells(i, 1).Font.Bold = True		'bold font'
    ObjExcel.rows(i).NumberFormat = "@" 		'formatting as text
    objExcel.columns(1).AutoFit()				'sizing the columns'
NEXT

'Create an array to track in-scope DAIL messages
DIM DAIL_message_array()

ReDim DAIL_message_array(7, 0)
'Incrementor for the array
Dail_count = 0

'constants for array
const dail_maxis_case_number_const      = 0
const dail_worker_const	                = 1
const dail_type_const                   = 2
const dail_month_const		            = 3
const dail_msg_const		            = 4
const full_dail_msg_const		        = 5
'Unneccessary - info is recorded in processing notes field
' const renewal_month_determination_const = 5
'Removed constant because redundant with processing notes
' const processable_based_on_dail_const   = 6
'To do - processing notes, would these be captured in case details array?
const dail_processing_notes_const       = 6
' To Do - is the excel row constant needed?
const dail_excel_row_const              = 7

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
const snap_only_const      = 3
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

    ' Ensuring valid_case_numbers is blanked out
    ' msgbox valid_case_numbers_list

    'To do - delete this after testing, used to test specific case numbers
    ' valid_case_numbers_list = "**"


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
            actionable_dail = ""
            renewal_6_month_check = ""

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

            'Reads the DAIL msg to determine if it is an out-of-scope message
            EMReadScreen dail_msg, 61, dail_row, 20
            dail_msg = trim(dail_msg)
            'List of out of scope messages pulled from non-actionable dails function
            If instr(dail_msg, "AMT CHILD SUPP MOD/ORD") OR _
                instr(dail_msg, "AP OF CHILD REF NBR:") OR _
                instr(dail_msg, "ADDRESS DIFFERS W/ CS RECORDS:") OR _
                instr(dail_msg, "AMOUNT AS UNEARNED INCOME IN LBUD IN THE MONTH") OR _
                instr(dail_msg, "AMOUNT AS UNEARNED INCOME IN SBUD IN THE MONTH") OR _
                instr(dail_msg, "CHILD SUPP PAYMT FREQUENCY IS MONTHLY FOR CHILD REF NBR") OR _
                instr(dail_msg, "CHILD SUPP PAYMT FREQUENCY IS MONTHLY FOR CHILD REF NBR") OR _
                instr(dail_msg, "CHILD SUPP PAYMTS PD THRU THE COURT/AGENCY FOR CHILD") OR _
                instr(dail_msg, "COMPLETE INFC PANEL") OR _
                instr(dail_msg, "IS LIVING W/CAREGIVER") OR _
                instr(dail_msg, "IS COOPERATING WITH CHILD SUPPORT") OR _
                instr(dail_msg, "IS NOT COOPERATING WITH CHILD SUPPORT") OR _
                instr(dail_msg, "NAME DIFFERS W/ CS RECORDS:") OR _
                instr(dail_msg, "PATERNITY ON CHILD REF NBR") OR _
                instr(dail_msg, "REPORTED NAME CHG TO:") OR _
                instr(dail_msg, "BENEFITS RETURNED, IF IOC HAS NEW ADDRESS") OR _
                instr(dail_msg, "CASE IS CATEGORICALLY ELIGIBLE") OR _
                instr(dail_msg, "CASE NOT AUTO-APPROVED - HRF/SR/RECERT DUE") OR _
                instr(dail_msg, "CHANGE IN BUDGET CYCLE") OR _
                instr(dail_msg, "COMPLETE ELIG IN FIAT") OR _
                instr(dail_msg, "COUNTED IN LBUD AS UNEARNED INCOME") OR _
                instr(dail_msg, "COUNTED IN SBUD AS UNEARNED INCOME") OR _
                instr(dail_msg, "HAS EARNED INCOME IN 6 MONTH BUDGET BUT NO DCEX PANEL") OR _
                instr(dail_msg, "NEW DENIAL ELIG RESULTS EXIST") OR _
                instr(dail_msg, "NEW ELIG RESULTS EXIST") OR _
                instr(dail_msg, "POTENTIALLY CATEGORICALLY ELIGIBLE") OR _
                instr(dail_msg, "WARNING MESSAGES EXIST") OR _
                instr(dail_msg, "ADDR CHG*CHK SHEL") OR _
                instr(dail_msg, "APPLCT ID CHNGD") OR _
                instr(dail_msg, "CASE AUTOMATICALLY DENIED") OR _
                instr(dail_msg, "CASE FILE INFORMATION WAS SENT ON") OR _
                instr(dail_msg, "CASE NOTE ENTERED BY") OR _
                instr(dail_msg, "CASE NOTE TRANSFER FROM") OR _
                instr(dail_msg, "CASE VOLUNTARY WITHDRAWN") OR _
                instr(dail_msg, "CASE XFER") OR _
                instr(dail_msg, "CHANGE REPORT FORM SENT ON") OR _
                instr(dail_msg, "DIRECT DEPOSIT STATUS") OR _
                instr(dail_msg, "EFUNDS HAS NOTIFIED DHS THAT THIS CLIENT'S EBT CARD") OR _
                instr(dail_msg, "MEMB:NEEDS INTERPRETER HAS BEEN CHANGED") OR _
                instr(dail_msg, "MEMB:SPOKEN LANGUAGE HAS BEEN CHANGED") OR _
                instr(dail_msg, "MEMB:RACE CODE HAS BEEN CHANGED FROM UNABLE") OR _
                instr(dail_msg, "MEMB:SSN HAS BEEN CHANGED FROM") OR _
                instr(dail_msg, "MEMB:SSN VER HAS BEEN CHANGED FROM") OR _
                instr(dail_msg, "MEMB:WRITTEN LANGUAGE HAS BEEN CHANGED FROM") OR _
                instr(dail_msg, "MEMI: HAS BEEN DELETED BY THE PMI MERGE PROCESS") OR _
                instr(dail_msg, "NOT ACCESSED FOR 300 DAYS,SPEC NOT") OR _
                instr(dail_msg, "PMI MERGED") OR _
                instr(dail_msg, "THIS APPLICATION WILL BE AUTOMATICALLY DENIED") OR _
                instr(dail_msg, "THIS CASE IS ERROR PRONE") OR _
                instr(dail_msg, "EMPL SERV REF DATE IS > 60 DAYS; CHECK ES PROVIDER RESPONSE") OR _
                instr(dail_msg, "LAST GRADE COMPLETED") OR _
                instr(dail_msg, "~*~*~CLIENT WAS SENT AN APPT LETTER") OR _
                instr(dail_msg, "IF CLIENT HAS NOT COMPLETED RECERT, APPL CAF FOR") OR _
                instr(dail_msg, "UPDATE PND2 FOR CLIENT DELAY IF APPROPRIATE") OR _
                instr(dail_msg, "PERSON HAS A RENEWAL OR HRF DUE. STAT UPDATES") OR _
                instr(dail_msg, "PERSON HAS HC RENEWAL OR HRF DUE") OR _
                instr(dail_msg, "GA: REVIEW DUE FOR JANUARY - NOT AUTO") OR _
                instr(dail_msg, "GA: STATUS IS PENDING - NOT AUTO-APPROVED") OR _
                instr(dail_msg, "GA: STATUS IS REIN OR SUSPEND - NOT AUTO-APPROVED") OR _
                instr(dail_msg, "GRH: REVIEW DUE - NOT AUTO") or _
                instr(dail_msg, "GRH: APPROVED VERSION EXISTS FOR JANUARY - NOT AUTO-APPROVED") OR _
                instr(dail_msg, "HEALTH CARE IS IN REINSTATE OR PENDING STATUS") OR _
                instr(dail_msg, "MSA RECERT DUE - NOT AUTO") or _
                instr(dail_msg, "MSA IN PENDING STATUS - NOT AUTO") or _
                instr(dail_msg, "APPROVED MSA VERSION EXISTS - NOT AUTO-APPROVED") OR _
                instr(dail_msg, "SNAP: RECERT/SR DUE FOR JANUARY - NOT AUTO") or _
                instr(dail_msg, "GRH: STATUS IS REIN, PENDING OR SUSPEND - NOT AUTO") OR _
                instr(dail_msg, "SDNH NEW JOB DETAILS FOR MEMB 00") OR _
                instr(dail_msg, "SNAP: PENDING OR STAT EDITS EXIST") OR _
                instr(dail_msg, "SNAP: REIN STATUS - NOT AUTO-APPROVED") OR _
                instr(dail_msg, "SNAP: APPROVED VERSION ALREADY EXISTS - NOT AUTO-APPROVED") OR _
                instr(dail_msg, "SNAP: AUTO-APPROVED - PREVIOUS UNAPPROVED VERSION EXISTS") OR _
                instr(dail_msg, "SSN DIFFERS W/ CS RECORDS") OR _
                instr(dail_msg, "MFIP MASS CHANGE AUTO-APPROVED AN UNUSUAL INCREASE") OR _
                instr(dail_msg, "MFIP MASS CHANGE AUTO-APPROVED CASE WITH SANCTION") OR _
                instr(dail_msg, "DWP MASS CHANGE AUTO-APPROVED AN UNUSUAL INCREASE") OR _
                instr(dail_msg, "IV-D NAME DISCREPANCY") OR _
                instr(dail_msg, "CHECK HAS BEEN APPROVED") OR _
                instr(dail_msg, "SDX INFORMATION HAS BEEN STORED - CHECK INFC") OR _
                instr(dail_msg, "BENDEX INFORMATION HAS BEEN STORED - CHECK INFC") OR _
                instr(dail_msg, "- TRANS #") OR _
                instr(dail_msg, "RSDI UPDATED - (REF") OR _
                instr(dail_msg, "SSI UPDATED - (REF") OR _
                instr(dail_msg, "SNAP ABAWD ELIGIBILITY HAS EXPIRED, APPROVE NEW ELIG RESULTS") then 
                    actionable_dail = False
            Else
                actionable_dail = True
            End If

            If (actionable_dail = True AND dail_type = "CSES") OR dail_type = "HIRE" Then
                'Read the MAXIS Case Number, if it is a new case number then pull case details. If it is not a new case number, then do not pull new case details.
                
                ' Msgbox "(actionable_dail = False AND dail_type = 'CSES') OR dail_type = 'HIRE' Then"

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

                            
                            If case_active = TRUE AND list_active_programs = "SNAP" AND list_pending_programs = "" Then
                                'Active case, SNAP only, no other active or pending programs
                                ' MsgBox "SNAP ONLY - case_status: " & case_status & " list_active_programs: " & list_active_programs & " list_pending_programs: " & list_pending_programs
                                case_details_array(snap_only_const, case_count) = "SNAP Only"

                                ' To do - check if background check is needed, may break connection to DAIL
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
                                        reporting_status = trim(reporting_status)

                                        If reporting_status = "SIX MONTH" Then
                                            ' MsgBox "six month reporter"
                                            'Navigate to STAT/REVW to confirm recertification and SR report date
                                            EMWriteScreen "STAT", 19, 22
                                            Call write_value_and_transmit("REVW", 19, 70)
                                            'Pause here as it sometimes errors
                                            EMWaitReady 0, 0
                                            'Open the FS screen
                                            Call write_value_and_transmit("X", 5, 58)
                                            EMWaitReady 0, 0

                                            EmReadscreen sr_report_date, 8, 9, 26
                                            EmReadscreen recertification_date, 8, 9, 64

                                            'Add handling for missing SR report date or recertification
                                            'Adds slashes to dates then converts to datedate from string to date
                                            If sr_report_date = "__ 01 __" Then
                                                sr_report_date = "SR Report Date is Missing"
                                                ' MsgBox "SR Report Date Not Entered"
                                            Else
                                                sr_report_date = replace(sr_report_date, " ", "/")
                                                sr_report_date = DateAdd("m", 0, sr_report_date)
                                            End If

                                            If recertification_date = "__ 01 __" Then
                                                recertification_date = "Recertification Date is Missing"
                                                ' MsgBox "Recertification Date Not Entered"
                                            Else
                                                recertification_date = replace(recertification_date, " ", "/")
                                                recertification_date = DateAdd("m", 0, recertification_date)
                                            End If
                    
                                            If sr_report_date <> "SR Report Date is Missing" and recertification_date <> "Recertification Date is Missing" Then 
                                                ' MsgBox "Both SR and recert dates are present"
                                                renewal_6_month_difference = DateDiff("M", sr_report_date, recertification_date)

                                                If renewal_6_month_difference = "6" or renewal_6_month_difference = "-6" then 
                                                    renewal_6_month_check = True
                                                Else 
                                                    renewal_6_month_check = False
                                                    case_details_array(case_processing_notes_const, case_count) = "SR Report Date and Recertification are not 6 months apart"
                                                End if
                                            
                                            Else
                                                ' MsgBox "One or both dates are missing"
                                                renewal_6_month_check = False
                                                case_details_array(case_processing_notes_const, case_count) = "SR Report Date and/or Recertification Date is missing"
                                            End If
                                            
                                            'Close the FS screen
                                            transmit
                                        Else
                                            ' MsgBox "not 6 month reporter"
                                            sr_report_date = "N/A"
                                            recertification_date = "N/A"

                                        End If

                                        

                                        ' MsgBox "Updating the footer month and year"
                                        'Change the footer month and year back to CM/CY otherwise the DAIL will carry forward footer month and year from previous DAIL message as it moves to the next one and could cause error
                                        'To do - is this necessary?
                                        ' EMWriteScreen MAXIS_footer_month, 19, 54
                                        ' EMWriteScreen MAXIS_footer_year, 19, 57
                                        ' ' MsgBox "did footer month year update?"
                                    End if
                                    
                                    ' MsgBox "Updating the case_details_array"
                                    'Update the array with new case details
                                    case_details_array(reporting_status_const, case_count) = trim(reporting_status)
                                    case_details_array(recertification_date_const, case_count) = trim(recertification_date)
                                    case_details_array(sr_report_date_const, case_count) = trim(sr_report_date)
                                End If


                            Else
                                case_details_array(processable_based_on_case_const, case_count) = False
                                case_details_array(reporting_status_const, case_count) = "N/A"
                                case_details_array(recertification_date_const, case_count) = "N/A"
                                case_details_array(sr_report_date_const, case_count) = "N/A"
                                case_details_array(case_processing_notes_const, case_count) = "N/A"
                                case_details_array(snap_only_const, case_count) = "Not SNAP Only"

                                ' ' MsgBox "Updating the footer month and year"
                                ' 'Update the footer month and year to CM/CY on CASE/CURR before returning to DAIL
                                ' 'To do - is this necessary?
                                ' EMWriteScreen MAXIS_footer_month, 20, 54
                                ' EMWriteScreen MAXIS_footer_year, 20, 57
                                ' ' MsgBox "did footer month year update?"

                            End If

                        End If    
                        
                        If case_details_array(snap_status_const, case_count) = "ACTIVE" AND case_details_array(snap_only_const, case_count) = "SNAP Only" AND case_details_array(reporting_status_const, case_count) = "SIX MONTH" and renewal_6_month_check = True then
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
                        objExcel.Cells(case_excel_row, 4).Value = case_details_array(snap_only_const, case_count)
                        objExcel.Cells(case_excel_row, 5).Value = case_details_array(reporting_status_const, case_count)
                        objExcel.Cells(case_excel_row, 6).Value = case_details_array(sr_report_date_const, case_count)
                        objExcel.Cells(case_excel_row, 7).Value = case_details_array(recertification_date_const, case_count)
                        objExcel.Cells(case_excel_row, 8).Value = case_details_array(case_processing_notes_const, case_count)
                        objExcel.Cells(case_excel_row, 9).Value = case_details_array(processable_based_on_case_const, case_count)
                        case_excel_row = case_excel_row + 1

                        'Return to DAIL by PF3
                        PF3

                        'Reset the footer month/year to CM through CASE/CURR
                        Call write_value_and_transmit("H", dail_row, 3)
                        EMWriteScreen MAXIS_footer_month, 20, 54
                        EMWriteScreen MAXIS_footer_year, 20, 57
                        PF3
                        ' ' MsgBox "did footer month year update?"
                      
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

                        'Check if the full message is displayed
                        EMReadScreen full_message_check, 36, 24, 2
                        If InStr(full_message_check, "THE ENTIRE MESSAGE TEXT") Then
                            ' MsgBox "entire message is displayed"
                            EMReadScreen dail_msg, 61, dail_row, 20
                            dail_msg = trim(dail_msg)
                            full_dail_msg = dail_msg
                            'Remove x from dail message
                            EMWriteScreen " ", dail_row, 3
                        Else
                            
                            ' Script reads the full DAIL message so that it can process, or not process, as needed.
                            EMReadScreen full_dail_msg_line_1, 60, 9, 5

                            EMReadScreen full_dail_msg_line_2, 60, 10, 5

                            EMReadScreen full_dail_msg_line_3, 60, 11, 5
                            full_dail_msg_line_3 = trim(full_dail_msg_line_3)
                            ' If full_dail_msg_line_3 <> "" Then Msgbox full_dail_msg_line_3

                            EMReadScreen full_dail_msg_line_4, 60, 12, 5
                            full_dail_msg_line_4 = trim(full_dail_msg_line_4)
                            ' If full_dail_msg_line_4 <> "" Then Msgbox full_dail_msg_line_4

                            If trim(full_dail_msg_line_2) = "" Then 
                                ' MsgBox "empty!"
                                full_dail_msg_line_1 = trim(full_dail_msg_line_1)
                            End If

                            full_dail_msg = trim(full_dail_msg_line_1 & full_dail_msg_line_2 & full_dail_msg_line_3 & full_dail_msg_line_4)

                            ' Msgbox full_dail_msg

                            'Transmit back to dail
                            transmit

                        End If

                        

                        'Confirming that dail message lists are updating properly
                        ' Msgbox "list_of_DAIL_messages_to_delete: " & list_of_DAIL_messages_to_delete
                        ' Msgbox "list_of_DAIL_messages_to_skip: " & list_of_DAIL_messages_to_skip

                        'The script has the full DAIL message and can compare against delete and skip lists to determine if it is a new message

                        'To do - consider more robust handling, should we validate that case number matches? That dail month matches? These could be added to the string - i.e. *123456 - CS DISB Type 36....*
                        If Instr(list_of_DAIL_messages_to_delete, "*" & full_dail_msg & "*") Then
                            'If the full dail message is within the list of dail messages to delete then the message should be deleted

                            'Resetting variables so they do not carry forward
                            last_dail_check = ""
                            other_worker_error = ""
                            total_dail_msg_count_before = ""
                            total_dail_msg_count_after = ""
                            all_done = ""
                            final_dail_error = ""
                            
                            'Check if script is about to delete the last dail message to avoid DAIL bouncing backwards or issue with deleting only message in the DAIL
                            EMReadScreen last_dail_check, 12, 3, 67
                            last_dail_check = trim(last_dail_check)

                            'If the current dail message is equal to the final dail message then it will delete the message and then exit the do loop so the script does not restart
                            last_dail_check = split(last_dail_check, " ")

                            If last_dail_check(0) = last_dail_check(2) then 
                                'The script is about to delete the LAST message in the DAIL so script will exit do loop after deletion, also works if it is about to delete the ONLY message in the DAIL
                                all_done = true
                            End If

                            'Delete the message
                            Call write_value_and_transmit("D", dail_row, 3)

                            'Handling for deleting message under someone else's x number
                            EMReadScreen other_worker_error, 25, 24, 2
                            other_worker_error = trim(other_worker_error)

                            If other_worker_error = "ALL MESSAGES WERE DELETED" Then
                                'Script deleted the final message in the DAIL
                                dail_row = dail_row - 1
                                dail_msg_deleted_count = dail_msg_deleted_count + 1
                                objExcel.Cells(dail_excel_row - 1, 7).Value = "Message successfully deleted. " & DAIL_message_array(dail_processing_notes_const, DAIL_count - 1)
                                'Exit do loop as all messages are deleted
                                all_done = true

                            ElseIf other_worker_error = "" Then
                                'Script appears to have deleted the message but there was no warning, checking DAIL counts to confirm deletion

                                'Handling to check if message actually deleted
                                total_dail_msg_count_before = last_dail_check(2) * 1
                                EMReadScreen total_dail_msg_count_after, 12, 3, 67

                                total_dail_msg_count_after = split(trim(total_dail_msg_count_after), " ")
                                total_dail_msg_count_after(2) = total_dail_msg_count_after(2) * 1

                                If last_dail_check(2) - 1 = total_dail_msg_count_after(2) Then
                                    'The total DAILs decreased by 1, message deleted successfully
                                    dail_row = dail_row - 1
                                    dail_msg_deleted_count = dail_msg_deleted_count + 1
                                    objExcel.Cells(dail_excel_row - 1, 7).Value = "Message successfully deleted. " & DAIL_message_array(dail_processing_notes_const, DAIL_count - 1)
                                Else
                                    'The total DAILs did not decrease by 1, something went wrong
                                    objExcel.Cells(dail_excel_row - 1, 7).Value = "Message deletion failed. " & DAIL_message_array(dail_processing_notes_const, DAIL_count - 1)
                                    script_end_procedure_with_error_report("Script end error - something went wrong with deleting the message at line 881.")
                                End If

                            ElseIf other_worker_error = "** WARNING ** YOU WILL BE" then 
                                
                                'Since the script is deleting another worker's DAIL message, need to transmit again to delete the message
                                transmit

                                'Reads the total number of DAILS after deleting to determine if it decreased by 1
                                EMReadScreen total_dail_msg_count_after, 12, 3, 67

                                'Checks if final DAIL message deleted
                                EMReadScreen final_dail_error, 25, 24, 2

                                If final_dail_error = "ALL MESSAGES WERE DELETED" Then
                                    'All DAIL messages deleted so indicates deletion a success
                                    dail_row = dail_row - 1
                                    dail_msg_deleted_count = dail_msg_deleted_count + 1
                                    objExcel.Cells(dail_excel_row - 1, 7).Value = "Message successfully deleted. " & DAIL_message_array(dail_processing_notes_const, DAIL_count - 1)
                                    'No more DAIL messages so exit do loop
                                    all_done = True
                                ElseIf trim(final_dail_error) = "" Then
                                    'Handling to check if message actually deleted
                                    total_dail_msg_count_before = last_dail_check(2) * 1

                                    total_dail_msg_count_after = split(trim(total_dail_msg_count_after), " ")
                                    total_dail_msg_count_after(2) = total_dail_msg_count_after(2) * 1

                                    If last_dail_check(2) - 1 = total_dail_msg_count_after(2) Then
                                        'The total DAILs decreased by 1, message deleted successfully
                                        dail_row = dail_row - 1
                                        dail_msg_deleted_count = dail_msg_deleted_count + 1
                                        objExcel.Cells(dail_excel_row - 1, 7).Value = "Message successfully deleted. " & DAIL_message_array(dail_processing_notes_const, DAIL_count - 1)
                                    Else
                                        objExcel.Cells(dail_excel_row - 1, 7).Value = "Message deletion failed. " & DAIL_message_array(dail_processing_notes_const, DAIL_count - 1)
                                        script_end_procedure_with_error_report("Script end error - something went wrong with deleting the message at line 915.")
                                    End If

                                Else
                                    objExcel.Cells(dail_excel_row - 1, 7).Value = "Message deletion failed. " & DAIL_message_array(dail_processing_notes_const, DAIL_count - 1)
                                    script_end_procedure_with_error_report("Script end error - something went wrong with deleting the message at line 920.")
                                End if
                                
                            Else
                                objExcel.Cells(dail_excel_row - 1, 7).Value = "Message deletion failed. " & DAIL_message_array(dail_processing_notes_const, DAIL_count - 1)
                                script_end_procedure_with_error_report("Script end error - something went wrong with deleting the message at line 925.")
                            End If

                            ' MsgBox "The message has been deleted. Did anything go wrong? If so, stop here!"
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
                                    DAIL_message_array(full_dail_msg_const, dail_count) = full_dail_msg

                                    'Activate the DAIL Messages sheet
                                    objExcel.Worksheets("DAIL Messages").Activate

                                    'Write dail details to the Excel sheet
                                    objExcel.Cells(dail_excel_row, 1).Value = DAIL_message_array(dail_maxis_case_number_const, dail_count)
                                    objExcel.Cells(dail_excel_row, 2).Value = DAIL_message_array(dail_worker_const, dail_count)
                                    objExcel.Cells(dail_excel_row, 3).Value = DAIL_message_array(dail_type_const, dail_count)
                                    objExcel.Cells(dail_excel_row, 4).Value = DAIL_message_array(dail_month_const, dail_count)
                                    objExcel.Cells(dail_excel_row, 5).Value = DAIL_message_array(dail_msg_const, dail_count)
                                    objExcel.Cells(dail_excel_row, 6).Value = DAIL_message_array(full_dail_msg_const, dail_count)

                                    ' Msgbox "case_details_array(processable_based_on_case_const, each_case): " & case_details_array(processable_based_on_case_const, each_case)

                                    If case_details_array(processable_based_on_case_const, each_case) = False Then
                                        
                                        ' Msgbox "case_details_array(processable_based_on_case_const, each_case) = False"

                                        If case_details_array(case_processing_notes_const, each_case) = "SR Report Date and Recertification are not 6 months apart" Then
                                            DAIL_message_array(dail_processing_notes_const, dail_count) = "QI review needed. SR Report Date and Recertification are not 6 months apart."
                                            QI_flagged_msg_count = QI_flagged_msg_count + 1
                                        ElseIf case_details_array(case_processing_notes_const, each_case) = "SR Report Date and/or Recertification Date is missing" Then
                                            DAIL_message_array(dail_processing_notes_const, dail_count) = "QI review needed. SR Report Date and/or Recertification Date is missing."
                                            QI_flagged_msg_count = QI_flagged_msg_count + 1
                                        Else
                                            DAIL_message_array(dail_processing_notes_const, dail_count) = "Not Processable based on Case Details"
                                            not_processable_msg_count = not_processable_msg_count + 1
                                        End If
                                        
                                        'The dail message should not be processed due to case details
                                        process_dail_message = False

                                        'to do - do we need to add to skip list? It shouldn't ever process since it is false based on case details
                                        list_of_DAIL_messages_to_skip = list_of_DAIL_messages_to_skip & full_dail_msg & "*"

                                        'Activate the DAIL Messages sheet
                                        objExcel.Worksheets("DAIL Messages").Activate

                                        'Update the Excel sheet
                                        'To do - can delete, no longer needed
                                        ' objExcel.Cells(dail_excel_row, 6).Value = DAIL_message_array(renewal_month_determination_const, dail_count)
                                        objExcel.Cells(dail_excel_row, 7).Value = DAIL_message_array(dail_processing_notes_const, dail_count)
                                    
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

                                                    DAIL_message_array(dail_processing_notes_const, dail_count) = "Not Processable due to DAIL Month & Recert/Renewal. DAIL Month is " & DateAdd("m", 0, Replace(dail_month, " ", "/01/")) & "."
                                                    objExcel.Cells(dail_excel_row, 7).Value = DAIL_message_array(dail_processing_notes_const, dail_count)
                                                    not_processable_msg_count = not_processable_msg_count + 1

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
                                                    DAIL_message_array(dail_processing_notes_const, dail_count) = "Not Processable due to DAIL Month & Recert/Renewal. DAIL Month is " & DateAdd("m", 0, Replace(dail_month, " ", "/01/")) & "."
                                                    objExcel.Cells(dail_excel_row, 7).Value = DAIL_message_array(dail_processing_notes_const, dail_count)
                                                    not_processable_msg_count = not_processable_msg_count + 1

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

                                                ' MsgBox check_full_dail_msg_line_1

                                                EMReadScreen check_full_dail_msg_line_2, 60, 10, 5
                                                ' MsgBox check_full_dail_msg_line_2

                                                EMReadScreen check_full_dail_msg_line_3, 60, 11, 5
                                                ' MsgBox check_full_dail_msg_line_3

                                                EMReadScreen check_full_dail_msg_line_4, 60, 12, 5
                                                ' MsgBox check_full_dail_msg_line_4

                                                If trim(check_full_dail_msg_line_2) = "" Then 
                                                    check_full_dail_msg_line_1 = trim(check_full_dail_msg_line_1)
                                                End If

                                                check_full_dail_msg = trim(check_full_dail_msg_line_1 & check_full_dail_msg_line_2 & check_full_dail_msg_line_3 & check_full_dail_msg_line_4)

                                                ' MsgBox check_full_dail_msg
                                                ' MsgBox full_dail_msg

                                                'To do - delete after testing
                                                If check_full_dail_msg = full_dail_msg Then
                                                    ' MsgBox "They match"
                                                Else
                                                    MsgBox "Something went wrong. The DAIL messages do not match"
                                                    ' MsgBox "STOP THE SCRIPT HERE"
                                                End if

                                                ' Script reads information from full message, particularly the PMI number(s) listed. The script creates new variables for each PMI number.
                                                'To do - likely should validate that this is ALWAYS starting point for PMIs for Type 36
                                                'Identify where 'PMI(S):' text is so that script can account for Type 36 and replaced Type 36 is
                                                'Set row and col
                                                row = 1
                                                col = 1
                                                EMSearch "PMI(S):", row, col
                                                EMReadScreen PMIs_line_one, 65 - (col + 8), row, col + 8
                                                ' MsgBox "PMIs_line_one: " & PMIs_line_one 
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
                                                        ' objExcel.Cells(dail_excel_row, 7).Value = DAIL_message_array(dail_processing_notes_const, DAIL_count)
                                                    ElseIf PMI_and_ref_nbr_array(PMI_match_found_const, each_individual) = True Then
                                                        ' Msgbox "PMI matched"
                                                        DAIL_message_array(dail_processing_notes_const, DAIL_count) = trim(DAIL_message_array(dail_processing_notes_const, DAIL_count) & " PMI #: " & PMI_and_ref_nbr_array(PMI_const, each_individual) & " found on case (M" & PMI_and_ref_nbr_array(ref_nbr_const, each_individual) & ").")
                                                        ' objExcel.Cells(dail_excel_row, 7).Value = DAIL_message_array(dail_processing_notes_const, DAIL_count)
                                                    Else
                                                        MsgBox "Something went wrong at line 1014"
                                                    End If
                                                Next

                                                'Only check UNEA panels if ALL PMIs are matched for DAIL message and for case. There are no PMIs that did not match within the array.
                                                If InStr(DAIL_message_array(dail_processing_notes_const, DAIL_count), "not found on case") = 0 Then
                                                    'If all PMIs are found on the case, then the script will navigate directly to STAT/UNEA from CASE/PERS to verify that UNEA panels exist for CS Type 36 for each identified PMI/reference number

                                                    'Update the processing notes to indicate that all PMIs found on the case rather than listing out on by one
                                                    DAIL_message_array(dail_processing_notes_const, DAIL_count) = "All PMIs found on case. " & DAIL_message_array(dail_processing_notes_const, DAIL_count)

                                                    ' Msgbox "InStr(DAIL_message_array(dail_processing_notes_const, DAIL_count), 'not found on case') = 0: " & InStr(DAIL_message_array(dail_processing_notes_const, DAIL_count), "not found on case")
                                                    ' MsgBox "PMIs all found on case"

                                                    ' Msgbox "Moving to STAT"
                                                    EMWriteScreen "STAT", 19, 22
                                                    Call write_value_and_transmit("UNEA", 19, 69)

                                                    EmReadScreen no_unea_panels_exist, 34, 24, 2

                                                    ' MsgBox "no_unea_panels_exist: " & "|" & no_unea_panels_exist & "|"

                                                    If trim(no_unea_panels_exist) = "UNEA DOES NOT EXIST FOR ANY MEMBER" Then
                                                        'If no UNEA panels exist for the case, then the case needs to be flagged for QI
                                                        ' Msgbox "no_unea_panels_exist: " & no_unea_panels_exist
                                                        DAIL_message_array(dail_processing_notes_const, DAIL_count) = " No UNEA panels exist for any member on the case." & DAIL_message_array(dail_processing_notes_const, DAIL_count)
                                                        

                                                        'To do - confirm, could be problematic that originally PF3 here instead of at end, it would back out of DAIL
                                                        ' ' Add full dail msg to list of dail messages to skip
                                                        ' 'To do - use check_full_dail_msg or just full_dail_msg
                                                        ' list_of_DAIL_messages_to_skip = list_of_DAIL_messages_to_skip & full_dail_msg & "*"

                                                        ' 'Navigate back to DAIL
                                                        ' PF3

                                                        ' 'To do - is it necessary to reset the footer month since it should update when going to CASE/CURR?
                                                        ' 'Need to reset the footer month and footer year without interrupting script navigation in DAIL so open CASE/CURR
                                                        ' Msgbox "Resetting footer month and year by going to case curr. Needed?"
                                                        ' Call write_value_and_transmit("H", dail_row, 3)

                                                        ' MsgBox "update footer month and year"
                                                        ' 'Update the footer month and year to CM/CY
                                                        ' EMWriteScreen MAXIS_footer_month, 20, 54
                                                        ' EMWriteScreen MAXIS_footer_year, 20, 57
                                                        ' MsgBox "Did footer month and year update?"
                                                        
                                                        ' 'Navigate back to DAIL
                                                        ' PF3

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
                                                                DAIL_message_array(dail_processing_notes_const, DAIL_count) = trim(DAIL_message_array(dail_processing_notes_const, DAIL_count) & " No UNEA panels exist for M" & PMI_and_ref_nbr_array(ref_nbr_const, each_individual) & ".")
                                                            Else
                                                                'Read the UNEA type
                                                                EMReadScreen unea_type, 2, 5, 37
                                                                ' Msgbox "unea_type: " & unea_type
                                                                If unea_type = "36" Then
                                                                    'To do - add flagging that the panel exists?
                                                                    'If it is a type 36 panel then it does not need to read the other panels
                                                                    ' Msgbox "unea_type: " & unea_type
                                                                    DAIL_message_array(dail_processing_notes_const, DAIL_count) = DAIL_message_array(dail_processing_notes_const, DAIL_count) & " A UNEA panel (Type 36) exists for M" & PMI_and_ref_nbr_array(ref_nbr_const, each_individual) & "."
                                                                Else
                                                                    'Check how many panels exist for the HH member
                                                                    EMReadScreen unea_panels_count, 1, 2, 78
                                                                    ' Msgbox "unea_panels_count: " & unea_panels_count
                                                                    ' MsgBox "Is it a number? " & IsNumeric(unea_panels_count)

                                                                    If unea_panels_count = "1" Then
                                                                        'If there is only one UNEA panel and it is not a Type 36 then it will update processing notes
                                                                        DAIL_message_array(dail_processing_notes_const, DAIL_count) = DAIL_message_array(dail_processing_notes_const, DAIL_count) & " No UNEA panel (Type 36) exists for M" & PMI_and_ref_nbr_array(ref_nbr_const, each_individual) & "."
                                                                        
                                                                        
                                                                    ElseIf unea_panels_count <> "1" Then
                                                                        'If there are more than just a single UNEA panel, loop through them all to check for Type 36
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
                                                                                'To do - add flagging that the panel exists?
                                                                                'If it is a type 36 panel then it does not need to read the other panels
                                                                                ' Msgbox "unea_type: " & unea_type
                                                                                DAIL_message_array(dail_processing_notes_const, DAIL_count) = DAIL_message_array(dail_processing_notes_const, DAIL_count) & " A UNEA panel (Type 36) exists for M" & PMI_and_ref_nbr_array(ref_nbr_const, each_individual) & "."
                                                                                ' msgbox "1093 it is about to exit do"
                                                                                Exit Do
                                                                            End if

                                                                            'Ensuring that both panel_count and unea_panels_count are both numbers
                                                                            panel_count = panel_count * 1
                                                                            unea_panels_count = unea_panels_count * 1

                                                                            'If the loop has reached the final panel without encountering a Type 36 message, then it updates the processing notes accordingly
                                                                            If panel_count = unea_panels_count Then
                                                                                DAIL_message_array(dail_processing_notes_const, DAIL_count) = DAIL_message_array(dail_processing_notes_const, DAIL_count) & " No UNEA panel (Type 36) exists for M" & PMI_and_ref_nbr_array(ref_nbr_const, each_individual) & "."
                                                                                Exit Do
                                                                            End If
                                                                        Loop
                                                                    End If
                                                                End If
                                                            End If
                                                        Next
                                                    End If

                                                    ' Msgbox "InStr(DAIL_message_array(dail_processing_notes_const, DAIL_count), 'does not exist'): " & InStr(DAIL_message_array(dail_processing_notes_const, DAIL_count), "does not exist")

                                                    If InStr(DAIL_message_array(dail_processing_notes_const, DAIL_count), "No UNEA panel") Then
                                                        'There is at least one missing Type 36 UNEA panel for a HH member. The DAIL message should be added to the skip list as it cannot be deleted and requires QI review.
                                                        list_of_DAIL_messages_to_skip = list_of_DAIL_messages_to_skip & full_dail_msg & "*"
                                                        'To do - ensure this is at the correct spot
                                                        'Update the excel spreadsheet with processing notes
                                                        objExcel.Cells(dail_excel_row, 7).Value = "QI review needed. " & DAIL_message_array(dail_processing_notes_const, DAIL_count)
                                                        QI_flagged_msg_count = QI_flagged_msg_count + 1
                                                    ElseIf InStr(DAIL_message_array(dail_processing_notes_const, DAIL_count), "No UNEA panel") = 0 Then
                                                        'All of the identified HH members have a corresponding Type 36 UNEA panel. The message can be deleted.
                                                        list_of_DAIL_messages_to_delete = list_of_DAIL_messages_to_delete & full_dail_msg & "*"
                                                        'To do - ensure this is at the correct spot
                                                        'Update the excel spreadsheet with processing notes
                                                        objExcel.Cells(dail_excel_row, 7).Value = "Message added to delete list. " & DAIL_message_array(dail_processing_notes_const, DAIL_count)
                                                        dail_row = dail_row - 1
                                                    End If


                                                Else
                                                    'There are PMIs in the DAIL message that are not on the case. Therefore, this message should be flagged for QI and added to the DAIL skip list when it is encountered again.
                                                    ' MsgBox "PMIs NOT ALL found on case"

                                                    list_of_DAIL_messages_to_skip = list_of_DAIL_messages_to_skip & full_dail_msg & "*"

                                                    'Update the excel spreadsheet with processing notes
                                                    'Ensure this is at correct spot
                                                    objExcel.Cells(dail_excel_row, 7).Value = "QI review needed. " & DAIL_message_array(dail_processing_notes_const, DAIL_count)
                                                    QI_flagged_msg_count = QI_flagged_msg_count + 1
                                                End If

                                                'To do - ensure this is at the correct spot
                                                'Update the excel spreadsheet with processing notes
                                                ' objExcel.Cells(dail_excel_row, 7).Value = DAIL_message_array(dail_processing_notes_const, DAIL_count)

                                                'Navigate back to the DAIL. This will reset to the top of the DAIL messages for the specific case number. Need to consider how to handle.
                                                ' MsgBox "navigate back to DAIL"
                                                PF3

                                            ElseIf InStr(dail_msg, "DISB SPOUSAL SUP (TYPE 37)") Then
                                                'Reset the caregiver reference number
                                                caregiver_ref_nbr = ""

                                                'Enters “X” on DAIL message to open full message. 
                                                Call write_value_and_transmit("X", dail_row, 3)

                                                ' Script reads the full DAIL message so that it can process, or not process, as needed.
                                                'To do - may not need to double-check messages after fully tested
                                                EMReadScreen check_full_dail_msg_line_1, 60, 9, 5

                                                ' MsgBox check_full_dail_msg_line_1

                                                EMReadScreen check_full_dail_msg_line_2, 60, 10, 5
                                                ' MsgBox check_full_dail_msg_line_2

                                                EMReadScreen check_full_dail_msg_line_3, 60, 11, 5
                                                ' MsgBox check_full_dail_msg_line_3

                                                EMReadScreen check_full_dail_msg_line_4, 60, 12, 5
                                                ' MsgBox check_full_dail_msg_line_4

                                                If trim(check_full_dail_msg_line_2) = "" Then 
                                                    check_full_dail_msg_line_1 = trim(check_full_dail_msg_line_1)
                                                End If

                                                check_full_dail_msg = trim(check_full_dail_msg_line_1 & check_full_dail_msg_line_2 & check_full_dail_msg_line_3 & check_full_dail_msg_line_4)

                                                ' MsgBox check_full_dail_msg
                                                ' MsgBox full_dail_msg

                                                'To do - delete after testing
                                                If check_full_dail_msg = full_dail_msg Then
                                                    ' MsgBox "They match"
                                                Else
                                                    MsgBox "Something went wrong. The DAIL messages do not match"
                                                    ' MsgBox "STOP THE SCRIPT HERE"
                                                End if

                                                'Identify where 'Ref Nbr:' text is so that script can account for slight changes in location in MAXIS
                                                'Set row and col
                                                row = 1
                                                col = 1
                                                EMSearch "REF NBR:", row, col
                                                EMReadScreen caregiver_ref_nbr, 2, row, col + 9

                                                ' MsgBox "caregiver ref nbr: " & caregiver_ref_nbr

                                                'Transmit back to DAIL message
                                                transmit

                                                'Navigate to STAT/UNEA to check for corresponding Type 37 UNEA panel
                                                Call write_value_and_transmit("S", dail_row, 3)
                                                Call write_value_and_transmit("UNEA", 20, 71)

                                                'Open the first panel of the caregiver reference number
                                                EMWriteScreen caregiver_ref_nbr, 20, 76
                                                Call write_value_and_transmit("01", 20, 79)

                                                'Check if no UNEA panel exists
                                                EmReadScreen unea_panel_check, 25, 24, 2

                                                'Check if UNEA panels exist for the caregiver reference number
                                                If InStr(unea_panel_check, "DOES NOT EXIST") Then
                                                    'There are no UNEA panels for this HH member. Updates the processing notes for the DAIL message to reflect this
                                                    DAIL_message_array(dail_processing_notes_const, DAIL_count) = trim(DAIL_message_array(dail_processing_notes_const, DAIL_count) & " No UNEA panels exist for caregiver M" & caregiver_ref_nbr & ".")
                                                Else
                                                    'Read the UNEA type
                                                    EMReadScreen unea_type, 2, 5, 37
                                                    ' Msgbox "unea_type: " & unea_type
                                                    If unea_type = "37" Then
                                                        'To do - add flagging that the panel exists?
                                                        'If it is a type 37 panel then it does not need to read the other panels
                                                        ' Msgbox "unea_type: " & unea_type
                                                        DAIL_message_array(dail_processing_notes_const, DAIL_count) = DAIL_message_array(dail_processing_notes_const, DAIL_count) & " A UNEA panel (Type 37) exists for caregiver M" & caregiver_ref_nbr & "."
                                                    Else
                                                        'Check how many panels exist for the HH member
                                                        EMReadScreen unea_panels_count, 1, 2, 78
                                                        ' MsgBox "unea_panels_count: " & unea_panels_count
                                                        ' MsgBox IsNumeric(unea_panels_count)
                                                        
                                                        If unea_panels_count = "1" Then
                                                            'If there is only one UNEA panel and it is not a Type 37 then it will update processing notes
                                                            DAIL_message_array(dail_processing_notes_const, DAIL_count) = DAIL_message_array(dail_processing_notes_const, DAIL_count) & " No UNEA panel (Type 37) exists for caregiver M" & PMI_and_ref_nbr_array(ref_nbr_const, each_individual) & "."

                                                        
                                                        ElseIf unea_panels_count <> "1" Then
                                                            'If there are more than just a single UNEA panel, loop through them all to check for Type 37
                                                            'Set incrementor for do loop
                                                            panel_count = 1

                                                            Do
                                                                panel_count = panel_count + 1
                                                                ' Msgbox "panel_count: " & panel_count
                                                                EMWriteScreen caregiver_ref_nbr, 20, 76
                                                                ' Msgbox "Where did it write the ref number?"
                                                                Call write_value_and_transmit("0" & panel_count, 20, 79)

                                                                'Read the UNEA type
                                                                EMReadScreen unea_type, 2, 5, 37
                                                                ' Msgbox "unea_type: " & unea_type
                                                                If unea_type = "37" Then
                                                                    'To do - add flagging that the panel exists?
                                                                    'If it is a type 36 panel then it does not need to read the other panels
                                                                    ' Msgbox "unea_type: " & unea_type
                                                                    DAIL_message_array(dail_processing_notes_const, DAIL_count) = DAIL_message_array(dail_processing_notes_const, DAIL_count) & " A UNEA panel (Type 37) exists for caregiver M" & caregiver_ref_nbr & "."
                                                                    Exit Do
                                                                End if

                                                                'Ensuring that both panel_count and unea_panels_count are both numbers
                                                                panel_count = panel_count * 1
                                                                unea_panels_count = unea_panels_count * 1

                                                                'If the loop has reached the final panel without encountering a Type 37 message, then it updates the processing notes accordingly
                                                                If panel_count = unea_panels_count Then
                                                                    DAIL_message_array(dail_processing_notes_const, DAIL_count) = DAIL_message_array(dail_processing_notes_const, DAIL_count) & " No UNEA panel (Type 37) exists for caregiver M" & caregiver_ref_nbr & "."
                                                                    Exit Do
                                                                End If
                                                            Loop
                                                        End If
                                                    End If
                                                End If

                                                    ' Msgbox "InStr(DAIL_message_array(dail_processing_notes_const, DAIL_count), 'does not exist'): " & InStr(DAIL_message_array(dail_processing_notes_const, DAIL_count), "does not exist")

                                                    If InStr(DAIL_message_array(dail_processing_notes_const, DAIL_count), "No UNEA panel") Then
                                                        'There is at least one missing Type 36 UNEA panel for a HH member. The DAIL message should be added to the skip list as it cannot be deleted and requires QI review.
                                                        list_of_DAIL_messages_to_skip = list_of_DAIL_messages_to_skip & full_dail_msg & "*"
                                                        'To do - ensure this is at the correct spot
                                                        'Update the excel spreadsheet with processing notes
                                                        objExcel.Cells(dail_excel_row, 7).Value = "QI review needed. " & DAIL_message_array(dail_processing_notes_const, DAIL_count)
                                                        QI_flagged_msg_count = QI_flagged_msg_count + 1
                                                    ElseIf InStr(DAIL_message_array(dail_processing_notes_const, DAIL_count), "No UNEA panel") = 0 Then
                                                        'All of the identified HH members have a corresponding Type 36 UNEA panel. The message can be deleted.
                                                        list_of_DAIL_messages_to_delete = list_of_DAIL_messages_to_delete & full_dail_msg & "*"
                                                        'To do - ensure this is at the correct spot
                                                        'Update the excel spreadsheet with processing notes
                                                        objExcel.Cells(dail_excel_row, 7).Value = "Message added to delete list. " & DAIL_message_array(dail_processing_notes_const, DAIL_count)

                                                        dail_row = dail_row - 1
                                                    End If

                                                'PF3 back to DAIL
                                                PF3

                                                ' MsgBox "DISB SPOUSAL SUP (TYPE 37): " & dail_msg
                                            ElseIf InStr(dail_msg, "DISB CS ARREARS (TYPE 39) OF") Then
                                                ' Msgbox "InStr(dail_msg, 'DISB CS (TYPE 39) OF'): " & InStr(dail_msg, "DISB CS (TYPE 39) OF")
                                                'Enters “X” on DAIL message to open full message. 
                                                Call write_value_and_transmit("X", dail_row, 3)

                                                ' Script reads the full DAIL message so that it can process, or not process, as needed.
                                                'To do - may not need to double-check messages after fully tested
                                                EMReadScreen check_full_dail_msg_line_1, 60, 9, 5

                                                ' MsgBox check_full_dail_msg_line_1

                                                EMReadScreen check_full_dail_msg_line_2, 60, 10, 5
                                                ' MsgBox check_full_dail_msg_line_2

                                                EMReadScreen check_full_dail_msg_line_3, 60, 11, 5
                                                ' MsgBox check_full_dail_msg_line_3

                                                EMReadScreen check_full_dail_msg_line_4, 60, 12, 5
                                                ' MsgBox check_full_dail_msg_line_4

                                                If trim(check_full_dail_msg_line_2) = "" Then 
                                                    ' MsgBox "empty!"
                                                    check_full_dail_msg_line_1 = trim(check_full_dail_msg_line_1)
                                                End If

                                                check_full_dail_msg = trim(check_full_dail_msg_line_1 & check_full_dail_msg_line_2 & check_full_dail_msg_line_3 & check_full_dail_msg_line_4)

                                                ' MsgBox check_full_dail_msg
                                                ' MsgBox full_dail_msg

                                                'To do - delete after testing
                                                If check_full_dail_msg = full_dail_msg Then
                                                    ' MsgBox "They match"
                                                Else
                                                    MsgBox "Something went wrong. The DAIL messages do not match"
                                                    ' MsgBox "STOP THE SCRIPT HERE"
                                                End if

                                                ' Script reads information from full message, particularly the PMI number(s) listed. The script creates new variables for each PMI number.
                                                'To do - likely should validate that this is ALWAYS starting point for PMIs for Type 39
                                                'Identify where 'PMI(S):' text is so that script can account for Type 39 and replaced Type 39 is
                                                'Set row and col
                                                row = 1
                                                col = 1
                                                EMSearch "PMI(S):", row, col
                                                EMReadScreen PMIs_line_one, 65 - (col + 8), row, col + 8
                                                ' MsgBox "PMIs_line_one: " & PMIs_line_one 
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
                                                        ' objExcel.Cells(dail_excel_row, 7).Value = DAIL_message_array(dail_processing_notes_const, DAIL_count)
                                                    ElseIf PMI_and_ref_nbr_array(PMI_match_found_const, each_individual) = True Then
                                                        ' Msgbox "PMI matched"
                                                        DAIL_message_array(dail_processing_notes_const, DAIL_count) = trim(DAIL_message_array(dail_processing_notes_const, DAIL_count) & " PMI #: " & PMI_and_ref_nbr_array(PMI_const, each_individual) & " found on case (M" & PMI_and_ref_nbr_array(ref_nbr_const, each_individual) & ").")
                                                        ' objExcel.Cells(dail_excel_row, 7).Value = DAIL_message_array(dail_processing_notes_const, DAIL_count)
                                                    Else
                                                        MsgBox "Something went wrong at line 1014"
                                                    End If
                                                Next

                                                'Only check UNEA panels if ALL PMIs are matched for DAIL message and for case. There are no PMIs that did not match within the array.
                                                If InStr(DAIL_message_array(dail_processing_notes_const, DAIL_count), "not found on case") = 0 Then
                                                    'If all PMIs are found on the case, then the script will navigate directly to STAT/UNEA from CASE/PERS to verify that UNEA panels exist for CS Type 39 for each identified PMI/reference number

                                                    'Update the processing notes to indicate that all PMIs found on the case rather than listing out on by one
                                                    DAIL_message_array(dail_processing_notes_const, DAIL_count) = "All PMIs found on case. " & DAIL_message_array(dail_processing_notes_const, DAIL_count)

                                                    ' Msgbox "InStr(DAIL_message_array(dail_processing_notes_const, DAIL_count), 'not found on case') = 0: " & InStr(DAIL_message_array(dail_processing_notes_const, DAIL_count), "not found on case")
                                                    ' MsgBox "PMIs all found on case"

                                                    ' Msgbox "Moving to STAT"
                                                    EMWriteScreen "STAT", 19, 22
                                                    Call write_value_and_transmit("UNEA", 19, 69)

                                                    EmReadScreen no_unea_panels_exist, 34, 24, 2

                                                    ' MsgBox "no_unea_panels_exist: " & "|" & no_unea_panels_exist & "|"

                                                    If trim(no_unea_panels_exist) = "UNEA DOES NOT EXIST FOR ANY MEMBER" Then
                                                        'If no UNEA panels exist for the case, then the case needs to be flagged for QI
                                                        ' Msgbox "no_unea_panels_exist: " & no_unea_panels_exist
                                                        DAIL_message_array(dail_processing_notes_const, DAIL_count) = " No UNEA panels exist for any member on the case." & DAIL_message_array(dail_processing_notes_const, DAIL_count)

                                                        'To do - confirm, seems like this could be problematic if PF3 here and later
                                                        ' ' Add full dail msg to list of dail messages to skip
                                                        ' 'To do - use check_full_dail_msg or just full_dail_msg
                                                        ' list_of_DAIL_messages_to_skip = list_of_DAIL_messages_to_skip & full_dail_msg & "*"

                                                        ' 'Navigate back to DAIL
                                                        ' PF3

                                                        ' 'To do - is it necessary to reset the footer month since it should update when going to CASE/CURR?
                                                        ' 'Need to reset the footer month and footer year without interrupting script navigation in DAIL so open CASE/CURR
                                                        ' Msgbox "Resetting footer month and year by going to case curr. Needed?"
                                                        ' Call write_value_and_transmit("H", dail_row, 3)

                                                        ' MsgBox "update footer month and year"
                                                        ' 'Update the footer month and year to CM/CY
                                                        ' EMWriteScreen MAXIS_footer_month, 20, 54
                                                        ' EMWriteScreen MAXIS_footer_year, 20, 57
                                                        ' MsgBox "Did footer month and year update?"
                                                        
                                                        ' 'Navigate back to DAIL
                                                        ' PF3

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
                                                                DAIL_message_array(dail_processing_notes_const, DAIL_count) = trim(DAIL_message_array(dail_processing_notes_const, DAIL_count) & " No UNEA panels exist for M" & PMI_and_ref_nbr_array(ref_nbr_const, each_individual) & ".")
                                                            Else
                                                                'Read the UNEA type
                                                                EMReadScreen unea_type, 2, 5, 37
                                                                ' Msgbox "unea_type: " & unea_type
                                                                If unea_type = "39" Then
                                                                    'To do - add flagging that the panel exists?
                                                                    'If it is a type 39 panel then it does not need to read the other panels
                                                                    ' Msgbox "unea_type: " & unea_type
                                                                    DAIL_message_array(dail_processing_notes_const, DAIL_count) = DAIL_message_array(dail_processing_notes_const, DAIL_count) & " A UNEA panel (Type 39) exists for M" & PMI_and_ref_nbr_array(ref_nbr_const, each_individual) & "."
                                                                Else
                                                                    'Check how many panels exist for the HH member
                                                                    EMReadScreen unea_panels_count, 1, 2, 78
                                                                    ' Msgbox "unea_panels_count: " & unea_panels_count
                                                                    ' MsgBox "Is it a number? " & IsNumeric(unea_panels_count)

                                                                    If unea_panels_count = "1" Then
                                                                        'If there is only one UNEA panel and it is not a Type 39 then it will update processing notes
                                                                        DAIL_message_array(dail_processing_notes_const, DAIL_count) = DAIL_message_array(dail_processing_notes_const, DAIL_count) & " No UNEA panel (Type 39) exists for M" & PMI_and_ref_nbr_array(ref_nbr_const, each_individual) & "."
                                                                        
                                                                    'If there are more than just a single UNEA panel, loop through them all to check for Type 39
                                                                    ElseIf unea_panels_count <> "1" Then
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
                                                                            If unea_type = "39" Then
                                                                                'To do - add flagging that the panel exists?
                                                                                'If it is a type 39 panel then it does not need to read the other panels
                                                                                ' Msgbox "unea_type: " & unea_type
                                                                                DAIL_message_array(dail_processing_notes_const, DAIL_count) = DAIL_message_array(dail_processing_notes_const, DAIL_count) & " A UNEA panel (Type 39) exists for M" & PMI_and_ref_nbr_array(ref_nbr_const, each_individual) & "."
                                                                                Exit Do
                                                                            End if

                                                                            'Ensuring that both panel_count and unea_panels_count are both numbers
                                                                            panel_count = panel_count * 1
                                                                            unea_panels_count = unea_panels_count * 1

                                                                            'If the loop has reached the final panel without encountering a Type 39 message, then it updates the processing notes accordingly
                                                                            If panel_count = unea_panels_count Then
                                                                                DAIL_message_array(dail_processing_notes_const, DAIL_count) = DAIL_message_array(dail_processing_notes_const, DAIL_count) & " No UNEA panel (Type 39) exists for M" & PMI_and_ref_nbr_array(ref_nbr_const, each_individual) & "."
                                                                                Exit Do
                                                                            End If
                                                                        Loop
                                                                    End If
                                                                End If
                                                            End If
                                                        Next
                                                    End If

                                                    ' Msgbox "InStr(DAIL_message_array(dail_processing_notes_const, DAIL_count), 'does not exist'): " & InStr(DAIL_message_array(dail_processing_notes_const, DAIL_count), "does not exist")

                                                    If InStr(DAIL_message_array(dail_processing_notes_const, DAIL_count), "No UNEA panel") Then
                                                        'There is at least one missing Type 39 UNEA panel for a HH member. The DAIL message should be added to the skip list as it cannot be deleted and requires QI review.
                                                        list_of_DAIL_messages_to_skip = list_of_DAIL_messages_to_skip & full_dail_msg & "*"
                                                        'To do - ensure this is at the correct spot
                                                        'Update the excel spreadsheet with processing notes
                                                        objExcel.Cells(dail_excel_row, 7).Value = "QI review needed. " & DAIL_message_array(dail_processing_notes_const, DAIL_count)
                                                        QI_flagged_msg_count = QI_flagged_msg_count + 1
                                                    ElseIf InStr(DAIL_message_array(dail_processing_notes_const, DAIL_count), "No UNEA panel") = 0 Then
                                                        'All of the identified HH members have a corresponding Type 39 UNEA panel. The message can be deleted.
                                                        list_of_DAIL_messages_to_delete = list_of_DAIL_messages_to_delete & full_dail_msg & "*"
                                                        'To do - ensure this is at the correct spot
                                                        'Update the excel spreadsheet with processing notes
                                                        objExcel.Cells(dail_excel_row, 7).Value = "Message added to delete list. " & DAIL_message_array(dail_processing_notes_const, DAIL_count)

                                                        dail_row = dail_row - 1
                                                    End If


                                                Else
                                                    'There are PMIs in the DAIL message that are not on the case. Therefore, this message should be flagged for QI and added to the DAIL skip list when it is encountered again.
                                                    ' MsgBox "PMIs NOT ALL found on case"

                                                    list_of_DAIL_messages_to_skip = list_of_DAIL_messages_to_skip & full_dail_msg & "*"

                                                    'Update the excel spreadsheet with processing notes
                                                    'Ensure this is at correct spot
                                                    objExcel.Cells(dail_excel_row, 7).Value = "QI review needed. " & DAIL_message_array(dail_processing_notes_const, DAIL_count)
                                                    QI_flagged_msg_count = QI_flagged_msg_count + 1

                                                End If

                                                'To do - ensure this is at the correct spot
                                                'Update the excel spreadsheet with processing notes
                                                ' objExcel.Cells(dail_excel_row, 7).Value = DAIL_message_array(dail_processing_notes_const, DAIL_count)

                                                'Navigate back to the DAIL. This will reset to the top of the DAIL messages for the specific case number. Need to consider how to handle.
                                                ' MsgBox "navigate back to DAIL"
                                                PF3

                                            ElseIf InStr(dail_msg, "DISB SPOUSAL SUP ARREARS (TYPE 40) OF") Then
                                                'Reset the caregiver reference number
                                                caregiver_ref_nbr = ""

                                                'Enters “X” on DAIL message to open full message. 
                                                Call write_value_and_transmit("X", dail_row, 3)

                                                ' Script reads the full DAIL message so that it can process, or not process, as needed.
                                                'To do - may not need to double-check messages after fully tested
                                                EMReadScreen check_full_dail_msg_line_1, 60, 9, 5

                                                ' MsgBox check_full_dail_msg_line_1

                                                EMReadScreen check_full_dail_msg_line_2, 60, 10, 5
                                                ' MsgBox check_full_dail_msg_line_2

                                                EMReadScreen check_full_dail_msg_line_3, 60, 11, 5
                                                ' MsgBox check_full_dail_msg_line_3

                                                EMReadScreen check_full_dail_msg_line_4, 60, 12, 5
                                                ' MsgBox check_full_dail_msg_line_4

                                                If trim(check_full_dail_msg_line_2) = "" Then 
                                                    ' MsgBox "empty!"
                                                    check_full_dail_msg_line_1 = trim(check_full_dail_msg_line_1)
                                                End If

                                                check_full_dail_msg = trim(check_full_dail_msg_line_1 & check_full_dail_msg_line_2 & check_full_dail_msg_line_3 & check_full_dail_msg_line_4)

                                                ' MsgBox check_full_dail_msg
                                                ' MsgBox full_dail_msg

                                                'To do - delete after testing
                                                If check_full_dail_msg = full_dail_msg Then
                                                    ' MsgBox "They match"
                                                Else
                                                    MsgBox "Something went wrong. The DAIL messages do not match"
                                                    ' MsgBox "STOP THE SCRIPT HERE"
                                                End if

                                                'Identify where 'Ref Nbr:' text is so that script can account for slight changes in location in MAXIS
                                                'Set row and col
                                                row = 1
                                                col = 1
                                                EMSearch "REF NBR:", row, col
                                                EMReadScreen caregiver_ref_nbr, 2, row, col + 9

                                                ' MsgBox "caregiver ref nbr: " & caregiver_ref_nbr

                                                'Transmit back to DAIL message
                                                transmit

                                                'Navigate to STAT/UNEA to check for corresponding Type 37 UNEA panel
                                                Call write_value_and_transmit("S", dail_row, 3)
                                                Call write_value_and_transmit("UNEA", 20, 71)

                                                'Open the first panel of the caregiver reference number
                                                EMWriteScreen caregiver_ref_nbr, 20, 76
                                                Call write_value_and_transmit("01", 20, 79)

                                                'Check if no UNEA panel exists
                                                EmReadScreen unea_panel_check, 25, 24, 2

                                                'Check if UNEA panels exist for the caregiver reference number
                                                If InStr(unea_panel_check, "DOES NOT EXIST") Then
                                                    'There are no UNEA panels for this HH member. Updates the processing notes for the DAIL message to reflect this
                                                    DAIL_message_array(dail_processing_notes_const, DAIL_count) = trim(DAIL_message_array(dail_processing_notes_const, DAIL_count) & " No UNEA panels exist for caregiver M" & caregiver_ref_nbr & ".")
                                                Else
                                                    'Read the UNEA type
                                                    EMReadScreen unea_type, 2, 5, 37
                                                    ' Msgbox "unea_type: " & unea_type
                                                    If unea_type = "40" Then
                                                        'To do - add flagging that the panel exists?
                                                        'If it is a type 40 panel then it does not need to read the other panels
                                                        ' Msgbox "unea_type: " & unea_type
                                                        DAIL_message_array(dail_processing_notes_const, DAIL_count) = DAIL_message_array(dail_processing_notes_const, DAIL_count) & "A UNEA panel (Type 40) exists for caregiver M" & caregiver_ref_nbr & "."
                                                    Else
                                                        'Check how many panels exist for the HH member
                                                        EMReadScreen unea_panels_count, 1, 2, 78
                                                        ' MsgBox "unea_panels_count: " & unea_panels_count
                                                        ' MsgBox "IsNumeric(unea_panels_count): " & IsNumeric(unea_panels_count)
                                                        If unea_panels_count = "1" Then
                                                            'If there is only one UNEA panel and it is not a Type 40 then it will update processing notes
                                                            DAIL_message_array(dail_processing_notes_const, DAIL_count) = DAIL_message_array(dail_processing_notes_const, DAIL_count) & "No UNEA panel (Type 40) exists for caregiver M" & PMI_and_ref_nbr_array(ref_nbr_const, each_individual) & "."
                                                        
                                                        ElseIf unea_panels_count <> "1" Then
                                                            'If there are more than just a single UNEA panel, loop through them all to check for Type 40
                                                            'Set incrementor for do loop
                                                            panel_count = 1

                                                            Do
                                                                panel_count = panel_count + 1
                                                                ' Msgbox "panel_count: " & panel_count
                                                                EMWriteScreen caregiver_ref_nbr, 20, 76
                                                                ' Msgbox "Where did it write the ref number?"
                                                                Call write_value_and_transmit("0" & panel_count, 20, 79)

                                                                'Read the UNEA type
                                                                EMReadScreen unea_type, 2, 5, 37
                                                                ' Msgbox "unea_type: " & unea_type
                                                                If unea_type = "40" Then
                                                                    'To do - add flagging that the panel exists?
                                                                    'If it is a type 40 panel then it does not need to read the other panels
                                                                    ' Msgbox "unea_type: " & unea_type
                                                                    DAIL_message_array(dail_processing_notes_const, DAIL_count) = DAIL_message_array(dail_processing_notes_const, DAIL_count) & "A UNEA panel (Type 40) exists for caregiver M" & caregiver_ref_nbr & "."
                                                                    Exit Do
                                                                End if

                                                                'Ensuring that both panel_count and unea_panels_count are both numbers
                                                                panel_count = panel_count * 1
                                                                unea_panels_count = unea_panels_count * 1
                                                                
                                                                'If the loop has reached the final panel without encountering a Type 40 message, then it updates the processing notes accordingly
                                                                If panel_count = unea_panels_count Then
                                                                    DAIL_message_array(dail_processing_notes_const, DAIL_count) = DAIL_message_array(dail_processing_notes_const, DAIL_count) & "No UNEA panel (Type 40) exists for caregiver M" & caregiver_ref_nbr & "."
                                                                    ' msgbox "Line 1426. It worked. panel_count = unea_panels_count BUT HAD TO CONVERT TO NUMBERS FOR SOME REASON"
                                                                    Exit Do
                                                                End If
                                                            Loop
                                                        End If
                                                    End If
                                                End If

                                                    ' Msgbox "InStr(DAIL_message_array(dail_processing_notes_const, DAIL_count), 'does not exist'): " & InStr(DAIL_message_array(dail_processing_notes_const, DAIL_count), "does not exist")

                                                    If InStr(DAIL_message_array(dail_processing_notes_const, DAIL_count), "No UNEA panel") Then
                                                        'There is at least one missing Type 40 UNEA panel for a HH member. The DAIL message should be added to the skip list as it cannot be deleted and requires QI review.
                                                        list_of_DAIL_messages_to_skip = list_of_DAIL_messages_to_skip & full_dail_msg & "*"
                                                        'To do - ensure this is at the correct spot
                                                        'Update the excel spreadsheet with processing notes
                                                        objExcel.Cells(dail_excel_row, 7).Value = "QI review needed. " & DAIL_message_array(dail_processing_notes_const, DAIL_count)
                                                        QI_flagged_msg_count = QI_flagged_msg_count + 1
                                                    ElseIf InStr(DAIL_message_array(dail_processing_notes_const, DAIL_count), "No UNEA panel") = 0 Then
                                                        'All of the identified HH members have a corresponding Type 40 UNEA panel. The message can be deleted.
                                                        list_of_DAIL_messages_to_delete = list_of_DAIL_messages_to_delete & full_dail_msg & "*"
                                                        'To do - ensure this is at the correct spot
                                                        'Update the excel spreadsheet with processing notes
                                                        objExcel.Cells(dail_excel_row, 7).Value = "Message added to delete list. " & DAIL_message_array(dail_processing_notes_const, DAIL_count)

                                                        dail_row = dail_row - 1
                                                    End If

                                                'PF3 back to DAIL
                                                PF3

                                            ElseIf InStr(dail_msg, "CS REPORTED: NEW EMPLOYER FOR CAREGIVER REF NBR:") Then
                                              
                                                ' Comment/uncomment for testing purposes
                                                DAIL_message_array(dail_processing_notes_const, DAIL_count) = "New Employer reported. Ignore for now."

                                                list_of_DAIL_messages_to_skip = list_of_DAIL_messages_to_skip & full_dail_msg & "*"
                                                'To do - ensure this is at the correct spot
                                                'Update the excel spreadsheet with processing notes
                                                objExcel.Cells(dail_excel_row, 7).Value = "QI review needed. " & DAIL_message_array(dail_processing_notes_const, DAIL_count)
                                                QI_flagged_msg_count = QI_flagged_msg_count + 1

                                                ' 'Reset the caregiver reference number
                                                ' caregiver_ref_nbr = ""

                                                ' 'Reset the no_exact_JOBS_panel_matches 
                                                ' no_exact_JOBS_panel_matches = ""

                                                ' 'Reset the list of employers
                                                ' list_of_employers_on_jobs_panels = "*"

                                                ' 'Reset the JOBS footer month and footer year
                                                ' JOBS_footer_month = ""
                                                ' JOBS_footer_year = ""

                                                ' 'Enters “X” on DAIL message to open full message. 
                                                ' Call write_value_and_transmit("X", dail_row, 3)

                                                ' 'Check if the full message is displayed
                                                ' EMReadScreen full_message_check, 36, 24, 2
                                                ' If InStr(full_message_check, "THE ENTIRE MESSAGE TEXT") Then
                                                '     EMReadScreen dail_msg, 61, dail_row, 20
                                                '     dail_msg = trim(dail_msg)
                                                '     check_full_dail_msg = dail_msg

                                                '     'Since the entire message is displayed, script reads the reference number and employer name from the dail_msg string
                                                '     caregiver_ref_nbr = Mid(check_full_dail_msg, instr(check_full_dail_msg, "REF NBR: ") + 9, 2)
                                                '     employer_full_name = Mid(check_full_dail_msg, instr(check_full_dail_msg, "REF NBR: ") + 12, 8)
                                                '     MsgBox "caregiver_ref_nbr: " & caregiver_ref_nbr & "     employer_full_name: " & employer_full_name

                                                '     'Remove x from dail message
                                                '     EMWriteScreen " ", dail_row, 3
                                                ' Else
                                                '     ' Script reads the full DAIL message so that it can process, or not process, as needed.
                                                    
                                                '     EMReadScreen check_full_dail_msg_line_1, 60, 9, 5
                                                '     ' MsgBox check_full_dail_msg_line_1

                                                '     EMReadScreen check_full_dail_msg_line_2, 60, 10, 5
                                                '     ' MsgBox check_full_dail_msg_line_2

                                                '     EMReadScreen check_full_dail_msg_line_3, 60, 11, 5
                                                '     ' MsgBox check_full_dail_msg_line_3

                                                '     EMReadScreen check_full_dail_msg_line_4, 60, 12, 5
                                                '     ' MsgBox check_full_dail_msg_line_4

                                                '     If trim(check_full_dail_msg_line_2) = "" Then check_full_dail_msg_line_1 = trim(check_full_dail_msg_line_1)

                                                '     check_full_dail_msg = trim(check_full_dail_msg_line_1 & check_full_dail_msg_line_2 & check_full_dail_msg_line_3 & check_full_dail_msg_line_4)

                                                '     'Identify where 'Ref Nbr:' text is so that script can account for slight changes in location in MAXIS
                                                '     'Set row and col
                                                '     row = 1
                                                '     col = 1
                                                '     EMSearch "REF NBR:", row, col
                                                '     EMReadScreen caregiver_ref_nbr, 2, row, col + 9

                                                '     'Identify where 'Ref Nbr:' text is so that script can account for slight changes in location in MAXIS
                                                '     'Set row and col
                                                '     row = 1
                                                '     col = 1
                                                '     EMSearch "REF NBR:", row, col
                                                '     EMReadScreen employer_name_line_1, 65 - (col + 12), row, col + 12

                                                '     If trim(check_full_dail_msg_line_2) = "" Then 
                                                '         employer_name_line_1 = trim(employer_name_line_1)
                                                '     End If
                                                
                                                '     employer_full_name = trim(employer_name_line_1 & check_full_dail_msg_line_2 & check_full_dail_msg_line_3 & check_full_dail_msg_line_4)
                                                '     MsgBox employer_full_name

                                                '     MsgBox "caregiver_ref_nbr: " & caregiver_ref_nbr & "     employer_full_name: " & employer_full_name
                                                    
                                                '     'Transmit back to DAIL message
                                                '     transmit

                                                ' End If

                                                ' 'To do - delete after testing
                                                ' If check_full_dail_msg = full_dail_msg Then
                                                '     ' MsgBox "They match"
                                                ' Else
                                                '     MsgBox "Something went wrong. The DAIL messages do not match. Stop here"
                                                ' End if

                                                ' 'Navigate to STAT/JOBS to check if corresponding JOBS panel exists
                                                ' Call write_value_and_transmit("S", dail_row, 3)
                                                ' Call write_value_and_transmit("JOBS", 20, 71)

                                                ' 'Open the first JOBS panel of the caregiver reference number
                                                ' EMWriteScreen caregiver_ref_nbr, 20, 76
                                                ' Call write_value_and_transmit("01", 20, 79)
                                                
                                                ' 'Check if no JOBS panel exists
                                                ' EmReadScreen jobs_panel_check, 25, 24, 2
                                                
                                                ' msgbox "Script navigated to first JOBS panel. It will determine if no jobs exist, 1 job exists, or multiple jobs exist."

                                                ' 'Check if JOBS panels exist for the caregiver reference number
                                                ' If InStr(jobs_panel_check, "DOES NOT EXIST") Then
                                                '     'There are no JOBS panels for this HH member. The script will add a new JOBS panel for the member
                                                '     MsgBox "No JOBS panel exist. Script will create new panel and fill it out. STOP HERE in production."

                                                '     Call write_value_and_transmit("NN", 20, 79)				'Creates new panel

                                                '     'Reads footer month for updating the panel
                                                '     EMReadScreen JOBS_footer_month, 2, 20, 55	
                                                '     EMReadScreen JOBS_footer_year, 2, 20, 58	

                                                '     'Writes information to JOBS panel
                                                '     EMWriteScreen "O", 5, 34
                                                '     EMWriteScreen "4", 6, 34
                                                '     EMWriteScreen employer_full_name, 7, 42
                                                '     EmWriteScreen JOBS_footer_month, 12, 54
                                                '     EMWriteScreen "01", 12, 57
                                                '     EmWriteScreen JOBS_footer_year, 12, 60

                                                '     'Puts $0 in as the received income amt
                                                '     EMWriteScreen "0", 12, 67				
                                                '     'Puts 0 hours in as the worked hours
                                                '     EMWriteScreen "0", 18, 72		
                                                    
                                                '     'Opens FS PIC
                                                '     Call write_value_and_transmit("X", 19, 38)
                                                '     Call create_MAXIS_friendly_date(date, 0, 5, 34) 'Puts date hired if message is from same month as hire ex 01/16 new hire for 1/17/16 start date.

                                                '     'Entering PIC information - PIC will update no matter is SNAP is active or not. Following steps for coding from POLI TEMP TE02.05.108 Denying/Closing SNAP for No Income Verif
                                                '     EMWriteScreen "1", 5, 64
                                                '     EMWriteScreen "0", 8, 64
                                                '     EMWriteScreen "0", 9, 66

                                                '     transmit
                                                '     EmReadScreen PIC_warning, 7, 20, 6
                                                '     IF PIC_warning = "WARNING" then transmit 'to clear message
                                                '     transmit 'back to JOBS panel
                                                '     MsgBox "It is about save the JOBS panel. Stop here if in testing or production"
                                                '     MsgBox "It is about save the JOBS panel. Stop here if in testing or production"
                                                '     ' transmit 'to save JOBS panel
                                               
                                                '     'Check if information is expiring and needs to be added to CM + 1
                                                '     EMReadScreen expired_check, 6, 24, 17 


                                                '     If expired_check = "EXPIRE" THEN 
                                                '         'New JOBS panel is expiring so it needs to be added to CM + 1 as well
                                                '         msgbox "New JOBS panel is expiring so it needs to be added to CM + 1 as well"

                                                '         'PF3 to go to STAT/WRAP
                                                '         PF3

                                                '         'Enter Y to add JOBS panel to CM + 1
                                                '         Call write_value_and_transmit("Y", 16, 54)
                                                '         'Navigate to STAT/JOBS for CM + 1
                                                '         Call write_value_and_transmit("JOBS", 20, 71)

                                                '         'Add new panel to caregiver ref nbr
                                                '         EMWriteScreen caregiver_ref_nbr, 20, 76
                                                '         Call write_value_and_transmit("NN", 20, 79)

                                                '         'Reads footer month for updating the panel
                                                '         EMReadScreen JOBS_footer_month, 2, 20, 55	
                                                '         EMReadScreen JOBS_footer_year, 2, 20, 58	

                                                '         'Writes information to JOBS panel
                                                '         EMWriteScreen "O", 5, 34
                                                '         EMWriteScreen "4", 6, 34
                                                '         EMWriteScreen employer_full_name, 7, 42
                                                '         EmWriteScreen JOBS_footer_month, 12, 54
                                                '         EMWriteScreen "01", 12, 57
                                                '         EmWriteScreen JOBS_footer_year, 12, 60

                                                '         'Puts $0 in as the received income amt
                                                '         EMWriteScreen "0", 12, 67				
                                                '         'Puts 0 hours in as the worked hours
                                                '         EMWriteScreen "0", 18, 72				

                                                '         'Opens FS PIC
                                                '         Call write_value_and_transmit("X", 19, 38)
                                                '         Call create_MAXIS_friendly_date(date, 0, 5, 34)

                                                '         'Entering PIC information 
                                                '         EMWriteScreen "1", 5, 64
                                                '         EMWriteScreen "0", 8, 64
                                                '         EMWriteScreen "0", 9, 66
                                                '         transmit
                                                '         EmReadScreen PIC_warning, 7, 20, 6
                                                '         IF PIC_warning = "WARNING" then transmit 'to clear message
                                                '         transmit 'back to JOBS panel
                                                '         'To Do - Uncomment once finalized
                                                '         MsgBox "The script is about to save the JOBS panel for CM + 1. Stop here if in testing or production"
                                                '         MsgBox "The script is about to save the JOBS panel for CM + 1. Stop here if in testing or production"
                                                '         ' transmit 'to save JOBS panel

                                                '         MsgBox "Script will not CASE/NOTE information"
                                                '         'Write information to CASE/NOTE

                                                '         'PF4 to navigate to CASE/NOTE
                                                '         PF4
                                                '         'Open new CASE/NOTE
                                                '         PF9

                                                '         CALL write_variable_in_case_note("-CS: NEW EMPLOYER REPORTED FOR (M" & caregiver_ref_nbr & ") for " & trim(employer_full_name) & "-")
                                                '         CALL write_variable_in_case_note("DATE HIRED: " & JOBS_footer_month & " " & JOBS_footer_year)
                                                '         CALL write_variable_in_case_note("EMPLOYER: " & employer_full_name)
                                                '         CALL write_variable_in_case_note("---")
                                                '         CALL write_variable_in_case_note("STAT/JOBS UPDATED WITH NEW HIRE INFORMATION FROM CSES DAIL MESSAGE.")
                                                '         CALL write_variable_in_case_note("---")
                                                '         CALL write_variable_in_case_note("REVIEW INCOME WITH RESIDENT AT RENEWAL/RECERTIFICATION AS CASE IS A SNAP 6-MONTH REPORTING CASE.")
                                                '         CALL write_variable_in_case_note("---")
                                                '         CALL write_variable_in_case_note(worker_signature)


                                                '         MsgBox "The script is about to save the CASE/NOTE for CM + 1. Stop here if in testing or production"
                                                '         MsgBox "The script is about to save the CASE/NOTE for CM + 1. Stop here if in testing or production"

                                                '         'PF3 to save the CASE/NOTE
                                                '         ' PF3

                                                '         'PF3 to STAT/WRAP
                                                '         PF3

                                                '         'PF3 back to JOBS
                                                '         PF3

                                                '     Else
                                                '         'If the JOBS panel is not expiring then write the information to CASE/NOTE

                                                '         MsgBox "Information is not expiring in CM + 1. Script will navigate to CASE/NOTE"
                                                        
                                                '         'PF4 to navigate to CASE/NOTE
                                                '         PF4
                                                '         'Open new CASE/NOTE
                                                '         PF9

                                                '         CALL write_variable_in_case_note("-CS: NEW EMPLOYER REPORTED FOR (M" & caregiver_ref_nbr & ") for " & trim(employer_full_name) & "-")
                                                '         CALL write_variable_in_case_note("DATE HIRED: " & JOBS_footer_month & " " & JOBS_footer_year)
                                                '         CALL write_variable_in_case_note("EMPLOYER: " & employer_full_name)
                                                '         CALL write_variable_in_case_note("---")
                                                '         CALL write_variable_in_case_note("STAT/JOBS UPDATED WITH NEW HIRE INFORMATION FROM CSES DAIL MESSAGE.")
                                                '         CALL write_variable_in_case_note("---")
                                                '         CALL write_variable_in_case_note("REVIEW INCOME WITH RESIDENT AT RENEWAL/RECERTIFICATION AS CASE IS A SNAP 6-MONTH REPORTING CASE.")
                                                '         CALL write_variable_in_case_note("---")
                                                '         CALL write_variable_in_case_note(worker_signature)


                                                '         MsgBox "The script is about to save the CASE/NOTE for CM + 1. Stop here if in testing or production"
                                                '         MsgBox "The script is about to save the CASE/NOTE for CM + 1. Stop here if in testing or production"

                                                '         'PF3 to save the CASE/NOTE
                                                '         ' PF3

                                                '         'PF3 back to JOBS
                                                '         PF3


                                                        
                                                '     End If

                                                '     ' 'PF3 back to DAIL
                                                '     ' PF3

                                                '     'Updates the processing notes for the DAIL message to reflect this
                                                '     msgbox "No jobs panels exist"
                                                    
                                                '     DAIL_message_array(dail_processing_notes_const, DAIL_count) = trim(DAIL_message_array(dail_processing_notes_const, DAIL_count) & " No JOBS panels exist for caregiver reference number: " & caregiver_ref_nbr & ". JOBS Panel and CASE/NOTE added for employer noted in CSES message. Message should be deleted.")

                                                
                                                ' Else
                                                '     'Read the employer name
                                                '     EMReadScreen employer_name_jobs_panel, 30, 7, 42
                                                '     employer_name_jobs_panel = trim(replace(employer_name_jobs_panel, "_", " "))


                                                '     If employer_name_jobs_panel = employer_full_name Then
                                                '         MsgBox "The employer names match exactly. Message can be deleted."
                                                        
                                                '         DAIL_message_array(dail_processing_notes_const, DAIL_count) = DAIL_message_array(dail_processing_notes_const, DAIL_count) & "A JOBS panel exists for employer: " & employer_full_name & " for M" & caregiver_ref_nbr & ". No CASE/NOTE created. Message should be deleted."

                                                '     Else
                                                '         'Check how many panels exist for the HH member
                                                '         EMReadScreen jobs_panels_count, 1, 2, 78
                                                '         'Convert jobs_panels_count to a number
                                                '         jobs_panels_count = jobs_panels_count * 1
                                                '         'If there is more than just 1 JOBS panel, loop through them all to check for matching employers
                                                '         If jobs_panels_count = 1 Then
                                                '             MsgBox "There is only one JOBS panel and they do not match. It will open the dialog to compare"

                                                '             'It adds the employer name to the list of employers so that it can be displayed on the dialog for verification
                                                '             list_of_employers_on_jobs_panels = list_of_employers_on_jobs_panels & employer_name_jobs_panel & "*"

                                                '             'Set variable below to true to trigger dialog
                                                '             no_exact_JOBS_panel_matches = True

                                                        
                                                        
                                                '         ElseIf jobs_panels_count <> 1 Then
                                                '             MsgBox "There are multiple JOBS panels and they do not match. It will open the dialog to compare"
                                                            
                                                '             'It adds the first employer name to the list of employers so that it can be displayed on the dialog for verification
                                                '             list_of_employers_on_jobs_panels = list_of_employers_on_jobs_panels & employer_name_jobs_panel & "*"

                                                '             'Set incrementor for do loop
                                                '             panel_count = 1

                                                '             Do
                                                '                 panel_count = panel_count + 1
                                                '                 EMWriteScreen caregiver_ref_nbr, 20, 76
                                                '                 Call write_value_and_transmit("0" & panel_count, 20, 79)

                                                '                 'Read the employer name
                                                '                 EMReadScreen employer_name_jobs_panel, 30, 7, 42
                                                '                 employer_name_jobs_panel = trim(replace(employer_name_jobs_panel, "_", " "))

                                                '                 If employer_name_jobs_panel = employer_full_name Then
                                                '                     MsgBox "That's convenient. The employer names match exactly"
                                                                
                                                '                     DAIL_message_array(dail_processing_notes_const, DAIL_count) = DAIL_message_array(dail_processing_notes_const, DAIL_count) & "A JOBS panel exists for employer: " & employer_full_name & " for M" & caregiver_ref_nbr & ". No CASE/NOTE created. Message should be deleted."

                                                '                     'Exit the do loop since an exact match was found
                                                '                     Exit Do
                                                '                 Else
                                                '                     'If the employer names do not match, then it adds to the employer name to the list of employers for review in dialog
                                                '                     list_of_employers_on_jobs_panels = list_of_employers_on_jobs_panels & employer_name_jobs_panel & "*"

                                                '                 End If

                                                '                 'Ensuring that both panel_count and unea_panels_count are both numbers
                                                '                 panel_count = panel_count * 1
                                                '                 jobs_panels_count = jobs_panels_count * 1
                                                                
                                                '                 If panel_count = jobs_panels_count Then
                                                '                     msgbox "2931 Since there were no exact employer matches, setting no_exact_JOBS_panel_matches = True"
                                                '                     'Since there were no exact employer matches, setting no_exact_JOBS_panel_matches = True
                                                '                     no_exact_JOBS_panel_matches = True
                                                '                     Exit Do
                                                '                 End If
                                                '             Loop
                                                '         End If

                                                '         'Convert string of the employer names into an array for use in the dialog
                                                '         'To do - add handling for when it has already been determined that there is a match on the employer names
                                                '         If no_exact_JOBS_panel_matches = True Then
                                                '             'Remove ending *
                                                '             list_of_employers_on_jobs_panels = Left(list_of_employers_on_jobs_panels, len(list_of_employers_on_jobs_panels) - 1)
                                                '             'Remove starting *
                                                '             list_of_employers_on_jobs_panels = Right(list_of_employers_on_jobs_panels, len(list_of_employers_on_jobs_panels) - 1)
                                                '             'Convert list of employer names from a string to a single dimensional array
                                                '             list_of_employers_on_jobs_panels = split(list_of_employers_on_jobs_panels, "*")

                                                '             'Alternative dialog and approach
                                                '             BeginDialog Dialog1, 0, 0, 321, 255, "Employers on JOBS Panel"
                                                '                 Text 5, 5, 100, 10, "Caregiver Reference Number:"
                                                '                 Text 105, 5, 20, 10, caregiver_ref_nbr
                                                '                 Text 55, 20, 50, 10, "Case Number:"
                                                '                 Text 105, 20, 80, 10, MAXIS_case_number
                                                '                 GroupBox 5, 40, 310, 115, "Employer on JOBS Panels"
                                                '                 Text 25, 55, 75, 10, "CSES - New Employer:"
                                                '                 Text 100, 55, 210, 10, employer_full_name
                                                '                 Text 10, 75, 90, 10, "Employer - JOBS Panel 01: "
                                                '                 Text 100, 75, 210, 10, list_of_employers_on_jobs_panels(0)
                                                '                 If UBound(list_of_employers_on_jobs_panels) >= 1 Then
                                                '                     Text 10, 90, 90, 10, "Employer - JOBS Panel 02:"
                                                '                     Text 100, 90, 210, 10, list_of_employers_on_jobs_panels(1)
                                                '                 End if
                                                '                 If UBound(list_of_employers_on_jobs_panels) >= 2 Then
                                                '                     Text 10, 105, 90, 10, "Employer - JOBS Panel 03:"
                                                '                     Text 100, 105, 210, 10, list_of_employers_on_jobs_panels(2)
                                                '                 End If
                                                '                 If UBound(list_of_employers_on_jobs_panels) >= 3 Then
                                                '                     Text 10, 120, 90, 10, "Employer - JOBS Panel 04:"
                                                '                     Text 100, 120, 210, 10, list_of_employers_on_jobs_panels(3)
                                                '                 End if
                                                '                 If UBound(list_of_employers_on_jobs_panels) >= 4 Then
                                                '                     Text 10, 135, 90, 10, "Employer - JOBS Panel 05:"
                                                '                     Text 100, 135, 210, 10, list_of_employers_on_jobs_panels(4)
                                                '                 End If
                                                '                 GroupBox 5, 160, 310, 65, "Employer Match Verification"
                                                '                 Text 10, 175, 110, 10, "Indicate if any Employers Match:"
                                                '                 DropListBox 120, 175, 190, 15, "[Select Option]"+chr(9)+"No Exact Match - Create New JOBS Panel"+chr(9)+"Potential Match(es) - Flag for Review"+chr(9)+"Exact Match Found - Delete Message", employer_match_determination
                                                '                 Text 10, 195, 235, 10, "Select the matching panel or indicate N/A:"
                                                '                 'To do - use cleaner code once workaround for no item in array is found
                                                '                 If UBound(list_of_employers_on_jobs_panels) = 4 Then
                                                '                     DropListBox 10, 205, 225, 15, "Not Applicable - No Match"+chr(9)+list_of_employers_on_jobs_panels(0)+chr(9)+list_of_employers_on_jobs_panels(1)+chr(9)+list_of_employers_on_jobs_panels(2)+chr(9)+list_of_employers_on_jobs_panels(3)+chr(9)+list_of_employers_on_jobs_panels(4), matching_employer_panel
                                                '                 ElseIf UBound(list_of_employers_on_jobs_panels) = 3 Then
                                                '                     DropListBox 10, 205, 225, 15, "Not Applicable - No Match"+chr(9)+list_of_employers_on_jobs_panels(0)+chr(9)+list_of_employers_on_jobs_panels(1)+chr(9)+list_of_employers_on_jobs_panels(2)+chr(9)+list_of_employers_on_jobs_panels(3), matching_employer_panel
                                                '                 ElseIf UBound(list_of_employers_on_jobs_panels) = 2 Then
                                                '                     DropListBox 10, 205, 225, 15, "Not Applicable - No Match"+chr(9)+list_of_employers_on_jobs_panels(0)+chr(9)+list_of_employers_on_jobs_panels(1)+chr(9)+list_of_employers_on_jobs_panels(2), matching_employer_panel
                                                '                 ElseIf UBound(list_of_employers_on_jobs_panels) = 1 Then
                                                '                     DropListBox 10, 205, 225, 15, "Not Applicable - No Match"+chr(9)+list_of_employers_on_jobs_panels(0)+chr(9)+list_of_employers_on_jobs_panels(1), matching_employer_panel
                                                '                 ElseIf UBound(list_of_employers_on_jobs_panels) = 0 Then
                                                '                     DropListBox 10, 205, 225, 15, "Not Applicable - No Match"+chr(9)+list_of_employers_on_jobs_panels(0), matching_employer_panel
                                                '                 End If

                                                '                 ButtonGroup ButtonPressed
                                                '                     OkButton 200, 235, 50, 15
                                                '                     CancelButton 255, 235, 50, 15
                                                '             EndDialog

                                                '             'Show dialog
                                                '             DO
                                                '                 DO
                                                '                     err_msg = ""
                                                '                     Dialog Dialog1
                                                '                     cancel_confirmation
                                                '                     'Validation to ensure that match determination made
                                                '                     If employer_match_determination = "[Select Option]" Then err_msg = err_msg & vbNewLine & "* You must indicate if any employers match."
                                                                    
                                                '                     'If one match is indicated, then the matching employer must be selected
                                                '                     If employer_match_determination = "Exact Match Found - Delete Message" AND matching_employer_panel = "Not Applicable - No Match" Then err_msg = err_msg & vbNewLine & "* You must select the matching employer."

                                                '                     'If there isn't an exact match, then the N/A option must be selected
                                                '                     If (employer_match_determination = "Potential Match(es) - Flag for Review" OR employer_match_determination = "No Exact Match - Create New JOBS Panel") AND matching_employer_panel <> "Not Applicable - No Match" Then err_msg = err_msg & vbNewLine & "* You indicated there was no exact match so you must select the 'Not Applicable - No Match' option."

                                                '                     IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
                                                '                 LOOP UNTIL err_msg = ""									'loops until all errors are resolved
                                                '                 CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
                                                '             LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in


                                                '             'Handling for match determination
                                                '             If employer_match_determination = "Potential Match(es) - Flag for Review" Then
                                                '                 'The message cannot be processed since no exact match exists
                                                '                 'Add the message to the skip list since it cannot be processed

                                                '                 MsgBox "Potential Match(es) - Flag for Review"

                                                '                 DAIL_message_array(dail_processing_notes_const, DAIL_count) = DAIL_message_array(dail_processing_notes_const, DAIL_count) & "There does not appear to be an exactly matching JOBS panel for employer: " & employer_full_name & " for M" & caregiver_ref_nbr & ". Review needed." & " Message should not be deleted."


                                                '             ElseIf employer_match_determination = "No Exact Match - Create New JOBS Panel" Then
                                                                
                                                '                 MsgBox "No Exact Match - Create New JOBS Panel"
                                                '                 '5 panels, note in array and don't add panel, add to skip list
                                                '                 If UBound(list_of_employers_on_jobs_panels) = 4 Then
                                                '                     MsgBox "There are 5 panels. Cannot add another. Message will not be deleted."

                                                '                     DAIL_message_array(dail_processing_notes_const, DAIL_count) = DAIL_message_array(dail_processing_notes_const, DAIL_count) & "A JOBS panel does not exist for employer: " & employer_full_name & " for M" & caregiver_ref_nbr & ", but unable to add a panel because 5 JOBS panels exist already. Review needed." & " Message should not be deleted."
                                                '                 ElseIf UBound(list_of_employers_on_jobs_panels) < 4 Then
                                                '                     'Less than 5 panels, add panel

                                                '                     MsgBox "There are less than 5 panels. New JOBS panel will be added"
                                                '                     Call write_value_and_transmit("NN", 20, 79)				'Creates new panel

                                                '                     'Reads footer month for updating the panel
                                                '                     EMReadScreen JOBS_footer_month, 2, 20, 55	
                                                '                     EMReadScreen JOBS_footer_year, 2, 20, 58	

                                                '                     'Writes information to JOBS panel
                                                '                     EMWriteScreen "O", 5, 34
                                                '                     EMWriteScreen "4", 6, 34
                                                '                     EMWriteScreen employer_full_name, 7, 42
                                                '                     EmWriteScreen JOBS_footer_month, 12, 54
                                                '                     EMWriteScreen "01", 12, 57
                                                '                     EmWriteScreen JOBS_footer_year, 12, 60

                                                '                     'Puts $0 in as the received income amt
                                                '                     EMWriteScreen "0", 12, 67				
                                                '                     'Puts 0 hours in as the worked hours
                                                '                     EMWriteScreen "0", 18, 72				

                                                '                     'Opens FS PIC
                                                '                     Call write_value_and_transmit("X", 19, 38)
                                                '                     Call create_MAXIS_friendly_date(date, 0, 5, 34) 

                                                '                     'Entering PIC information
                                                '                     EMWriteScreen "1", 5, 64
                                                '                     EMWriteScreen "0", 8, 64
                                                '                     EMWriteScreen "0", 9, 66
                                                '                     transmit
                                                '                     EmReadScreen PIC_warning, 7, 20, 6
                                                '                     IF PIC_warning = "WARNING" then transmit 'to clear message
                                                '                     transmit 'back to JOBS panel
                                                                    
                                                '                     'To Do - Uncomment once finalized
                                                '                     MsgBox "It is about save the JOBS panel. Stop here if in testing or production"
                                                '                     MsgBox "It is about save the JOBS panel. Stop here if in testing or production"
                                                '                     ' transmit 'to save JOBS panel
                                                                    
                                                '                     'Checks to see if the jobs panel will carry over by looking for the "This information will expire" at the bottom of the page
                                                '                     EMReadScreen expired_check, 6, 24, 17 

                                                '                     If expired_check = "EXPIRE" THEN 

                                                '                         MsgBox "Info will expire at end of month. Navigating to CM + 1"

                                                '                         'PF3 to go to STAT/WRAP
                                                '                         PF3
                                                '                         'Enter Y to add JOBS panel to CM + 1
                                                '                         Call write_value_and_transmit("Y", 16, 54)
                                                '                         'Navigate to STAT/JOBS for CM + 1
                                                '                         Call write_value_and_transmit("JOBS", 20, 71)

                                                '                         'Check if there are 5 jobs already for CM + 1
                                                '                         EMReadScreen five_JOBS_panels_check, 1, 2, 78

                                                '                         If five_JOBS_panels_check = "5" Then
                                                '                             MsgBox "There are 5 panels in CM + 1. Cannot add another. Message will not be deleted"

                                                '                             DAIL_message_array(dail_processing_notes_const, DAIL_count) = DAIL_message_array(dail_processing_notes_const, DAIL_count) & "A JOBS panel does not exist for employer: " & employer_full_name & " for M" & caregiver_ref_nbr & ". JOBS panel added for CM, but unable to add a JOBS panel to CM + 1 because 5 JOBS panels exist already. Review needed." & " Message should not be deleted."

                                                '                         Else
                                                '                             'There are less than 5 JOBS panels so add new panel to caregiver ref nbr
                                                '                             EMWriteScreen caregiver_ref_nbr, 20, 76
                                                '                             Call write_value_and_transmit("NN", 20, 79)

                                                '                             'Reads footer month for updating the panel
                                                '                             EMReadScreen JOBS_footer_month, 2, 20, 55	
                                                '                             EMReadScreen JOBS_footer_year, 2, 20, 58	

                                                '                             'Writes information to JOBS panel
                                                '                             EMWriteScreen "O", 5, 34
                                                '                             EMWriteScreen "4", 6, 34
                                                '                             EMWriteScreen employer_full_name, 7, 42
                                                '                             EmWriteScreen JOBS_footer_month, 12, 54
                                                '                             EMWriteScreen "01", 12, 57
                                                '                             EmWriteScreen JOBS_footer_year, 12, 60

                                                '                             'Puts $0 in as the received income amt
                                                '                             EMWriteScreen "0", 12, 67				
                                                '                             'Puts 0 hours in as the worked hours
                                                '                             EMWriteScreen "0", 18, 72				

                                                '                             'Opens FS PIC
                                                '                             Call write_value_and_transmit("X", 19, 38)
                                                '                             Call create_MAXIS_friendly_date(date, 0, 5, 34)

                                                '                             'Entering PIC information
                                                '                             EMWriteScreen "1", 5, 64
                                                '                             EMWriteScreen "0", 8, 64
                                                '                             EMWriteScreen "0", 9, 66
                                                '                             transmit
                                                '                             EmReadScreen PIC_warning, 7, 20, 6
                                                '                             IF PIC_warning = "WARNING" then transmit 'to clear message
                                                '                             transmit 'back to JOBS panel
                                                                            
                                                '                             MsgBox "It is about save the JOBS panel. Stop here if in testing or production"
                                                '                             MsgBox "It is about save the JOBS panel. Stop here if in testing or production"

                                                '                             transmit 'to save JOBS panel

                                                '                             'Write information to CASE/NOTE

                                                '                             'PF4 to navigate to CASE/NOTE
                                                '                             PF4
                                                '                             'Open new CASE/NOTE
                                                '                             PF9

                                                '                             CALL write_variable_in_case_note("-CS: NEW EMPLOYER REPORTED FOR (M" & caregiver_ref_nbr & ") for " & trim(employer_full_name) & "-")
                                                '                             CALL write_variable_in_case_note("DATE HIRED: " & JOBS_footer_month & " " & JOBS_footer_year)
                                                '                             CALL write_variable_in_case_note("EMPLOYER: " & employer_full_name)
                                                '                             CALL write_variable_in_case_note("---")
                                                '                             CALL write_variable_in_case_note("STAT/JOBS UPDATED WITH NEW HIRE INFORMATION FROM CSES DAIL MESSAGE.")
                                                '                             CALL write_variable_in_case_note("---")
                                                '                             CALL write_variable_in_case_note("REVIEW INCOME WITH RESIDENT AT RENEWAL/RECERTIFICATION AS CASE IS A SNAP 6-MONTH REPORTING CASE.")
                                                '                             CALL write_variable_in_case_note("---")
                                                '                             CALL write_variable_in_case_note(worker_signature)

                                                '                             MsgBox "It is about save the JOBS panel. Stop here if in testing or production"
                                                '                             MsgBox "It is about save the JOBS panel. Stop here if in testing or production"

                                                                            
                                                '                             'PF3 to save the CASE/NOTE
                                                '                             PF3
                                                '                         End If
                                                                        
                                                '                         'PF3 to STAT/WRAP
                                                '                         PF3

                                                '                         'PF3 back to JOBS
                                                '                         PF3

                                                '                     Else
                                                '                         'If not expiring at end of month, then add a CASE/NOTE

                                                '                         'Write information to CASE/NOTE
                                                '                         'PF4 to navigate to CASE/NOTE
                                                '                         PF4
                                                                        
                                                '                         'Open new CASE/NOTE
                                                '                         PF9

                                                '                         CALL write_variable_in_case_note("-CS: NEW EMPLOYER REPORTED FOR (M" & caregiver_ref_nbr & ") for " & trim(employer_full_name) & "-")
                                                '                         CALL write_variable_in_case_note("DATE HIRED: " & JOBS_footer_month & " " & JOBS_footer_year)
                                                '                         CALL write_variable_in_case_note("EMPLOYER: " & employer_full_name)
                                                '                         CALL write_variable_in_case_note("---")
                                                '                         CALL write_variable_in_case_note("STAT/JOBS UPDATED WITH NEW HIRE INFORMATION FROM CSES DAIL MESSAGE.")
                                                '                         CALL write_variable_in_case_note("---")
                                                '                         CALL write_variable_in_case_note("REVIEW INCOME WITH RESIDENT AT RENEWAL/RECERTIFICATION AS CASE IS A SNAP 6-MONTH REPORTING CASE.")
                                                '                         CALL write_variable_in_case_note("---")
                                                '                         CALL write_variable_in_case_note(worker_signature)

                                                '                         MsgBox "It is about save the JOBS panel. Stop here if in testing or production"
                                                '                         MsgBox "It is about save the JOBS panel. Stop here if in testing or production"

                                                '                         'PF3 to save the CASE/NOTE
                                                '                         PF3

                                                '                         'PF3 back to JOBS
                                                '                         PF3

                                                '                     End If

                                                '                     DAIL_message_array(dail_processing_notes_const, DAIL_count) = DAIL_message_array(dail_processing_notes_const, DAIL_count) & "A JOBS panel did not exist for employer: " & employer_full_name & " for M" & caregiver_ref_nbr & ". A JOBS panel for this employer was successfully added, along with a CASE/NOTE." & " Message should be deleted."
                                                                   
                                                '                     'Unecessary to back out of DAIL here
                                                '                     ' PF3
                                                '                 End If

                                                '             ElseIf employer_match_determination = "Exact Match Found - Delete Message" Then
                                                '                 'There is an exact match for the employer, delete the message

                                                '                 MsgBox "Exact Match Found - Delete Message"

                                                '                 DAIL_message_array(dail_processing_notes_const, DAIL_count) = DAIL_message_array(dail_processing_notes_const, DAIL_count) & "A JOBS panel did match for employer: " & employer_name_jobs_panel & " for M" & caregiver_ref_nbr & "." & " Message should be deleted."

                                                '                 'Add message to delete list
                                                '                 'PF3 back to DAIL?

                                                '             End If
                                                                


                                                '         End if

                                                '     End If
                                                ' End If

                                                '     ' Msgbox "InStr(DAIL_message_array(dail_processing_notes_const, DAIL_count), 'does not exist'): " & InStr(DAIL_message_array(dail_processing_notes_const, DAIL_count), "does not exist")

                                                '     If InStr(DAIL_message_array(dail_processing_notes_const, DAIL_count), "Message should not be deleted") Then
                                                '         'The DAIL message should be added to the skip list as it cannot be deleted and requires QI review.
                                                '         list_of_DAIL_messages_to_skip = list_of_DAIL_messages_to_skip & full_dail_msg & "*"
                                                '         'Update the excel spreadsheet with processing notes
                                                '         objExcel.Cells(dail_excel_row, 7).Value = "QI review needed. " & DAIL_message_array(dail_processing_notes_const, DAIL_count)
                                                '         QI_flagged_msg_count = QI_flagged_msg_count + 1
                                                '     ElseIf InStr(DAIL_message_array(dail_processing_notes_const, DAIL_count), "Message should be deleted") Then
                                                '         'There is a corresponding JOBS panel or a JOBS panel was created. The message can be deleted.
                                                '         list_of_DAIL_messages_to_delete = list_of_DAIL_messages_to_delete & full_dail_msg & "*"
                                                '         'Update the excel spreadsheet with processing notes
                                                '         objExcel.Cells(dail_excel_row, 7).Value = "Message added to delete list. " & DAIL_message_array(dail_processing_notes_const, DAIL_count)
                                                '     End If

                                                ' 'PF3 back to DAIL
                                                ' PF3

                                                ' MsgBox "The message has been processed and script will navigate back to DAIL now."

                                            ElseIf InStr(dail_msg, "REPORTED: CHILD REF NBR:") Then
                                                'No action on these, simply note in spreadsheet that QI team to review

                                                ' MsgBox "REPORTED: CHILD REF NBR:" & dail_msg
                                                
                                                DAIL_message_array(dail_processing_notes_const, DAIL_count) = "QI Review. CHILD NO LONGER RESIDES WITH CAREGIVER."

                                                list_of_DAIL_messages_to_skip = list_of_DAIL_messages_to_skip & full_dail_msg & "*"
                                                'To do - ensure this is at the correct spot
                                                'Update the excel spreadsheet with processing notes
                                                objExcel.Cells(dail_excel_row, 7).Value = "QI review needed. " & DAIL_message_array(dail_processing_notes_const, DAIL_count)
                                                QI_flagged_msg_count = QI_flagged_msg_count + 1
                                            Else
                                                'To do - add handling for CSES type messages that are not in scope
                                                msgbox "Something went wrong - line 2189ish." & "dail_msg: " & dail_msg 

                                            End If


                                        ElseIf process_dail_message = True and dail_type = "HIRE" Then
                                            'Update here

                                            If InStr(dail_msg, "NDNH MEMB") Then
                                                'Add logic here
                                                ' MsgBox "NDNH MEMB: " & dail_msg
                                            ElseIf InStr(dail_msg, "NEW JOB DETAILS FOR SSN") Then
                                                ' MsgBox "NEW JOB DETAILS FOR SSN: " & dail_msg

                                                'No action on these, simply note in spreadsheet that QI team to review

                                                ' MsgBox "NEW JOB DETAILS FOR SSN:" & dail_msg
                                                
                                                DAIL_message_array(dail_processing_notes_const, DAIL_count) = "QI Review. Outdated HIRE message."

                                                list_of_DAIL_messages_to_skip = list_of_DAIL_messages_to_skip & full_dail_msg & "*"
                                                'To do - ensure this is at the correct spot
                                                'Update the excel spreadsheet with processing notes
                                                objExcel.Cells(dail_excel_row, 7).Value = "QI review needed. " & DAIL_message_array(dail_processing_notes_const, DAIL_count)
                                                QI_flagged_msg_count = QI_flagged_msg_count + 1
                                            ElseIf InStr(dail_msg, "SDNH NEW JOB DETAILS") Then
                                                'Add logic here
                                                ' MsgBox "SDNH NEW JOB DETAILS: " & dail_msg
                                            ElseIf InStr(dail_msg, "JOB DETAILS FOR  ") Then
                                                'No action on these, simply note in spreadsheet that QI team to review

                                                ' MsgBox "NEW JOB DETAILS FOR SSN:" & dail_msg
                                                
                                                DAIL_message_array(dail_processing_notes_const, DAIL_count) = "QI Review. Outdated HIRE message."

                                                list_of_DAIL_messages_to_skip = list_of_DAIL_messages_to_skip & full_dail_msg & "*"
                                                'To do - ensure this is at the correct spot
                                                'Update the excel spreadsheet with processing notes
                                                objExcel.Cells(dail_excel_row, 7).Value = "QI review needed. " & DAIL_message_array(dail_processing_notes_const, DAIL_count)
                                                QI_flagged_msg_count = QI_flagged_msg_count + 1
                                            Else
                                                ' msgbox "Something went wrong - line 1964"
                                            End If
                                        Else
                                            ' MsgBox "Something went wrong = 1269"
                                            ' MsgBox "process_dail_message: " & process_dail_message
                                            ' MsgBox "dail_type: " & dail_type
                                            ' MsgBox "Stop here"
                                        End If

                                        
                                        


                                    End If

                                    'Increment the dail_excel_row so that data isn't overwritten
                                    dail_excel_row = dail_excel_row + 1
                                    
                                    'Increment dail_count for the dail array
                                    dail_count = dail_count + 1

                                    'In instances where the case details are not the final item in the array, need to exit the for loop
                                    Exit For

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

            ' 'Increment the stats counter
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
objExcel.Cells(1, 2).Value = case_excel_row - 2
objExcel.Cells(2, 2).Value = dail_excel_row - 2
objExcel.Cells(3, 2).Value = not_processable_msg_count
objExcel.Cells(4, 2).Value = dail_msg_deleted_count
objExcel.Cells(5, 2).Value = QI_flagged_msg_count
objExcel.Cells(6, 2).Value = timer - start_time
objExcel.Cells(7, 2).Value = ((STATS_counter * STATS_manualtime) - (timer - start_time)) / 60

'Finding the right folder to automatically save the file
this_month = CM_mo & " " & CM_yr
month_folder = "DAIL " & CM_mo & "-" & DatePart("yyyy", date) & ""
unclear_info_folder = replace(this_month, " ", "-") & " DAIL Unclear Info"
report_date = replace(date, "/", "-")

'saving the Excel file
If CSES_messages = 1 and HIRE_messages = 1 Then 
    file_info = month_folder & "\" & unclear_info_folder & "\" & report_date & " Unclear Info" & " " & "CSES and HIRE" & " " & dail_msg_deleted_count
ElseIf CSES_messages = 0 and HIRE_messages = 1 Then 
    file_info = month_folder & "\" & unclear_info_folder & "\" & report_date & " Unclear Info" & " " & "HIRE" & " " & dail_msg_deleted_count
ElseIf CSES_messages = 1 and HIRE_messages = 0 Then 
    file_info = month_folder & "\" & unclear_info_folder & "\" & report_date & " Unclear Info" & " " & "CSES" & " " & dail_msg_deleted_count
End if

'Saves and closes the most recent Excel workbook with the Task based cases to process.
objExcel.ActiveWorkbook.SaveAs "T:\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\DAIL list\" & file_info & ".xlsx"
objExcel.ActiveWorkbook.Close
objExcel.Application.Quit
objExcel.Quit

script_end_procedure_with_error_report("Success! Please review the list created for accuracy.")
