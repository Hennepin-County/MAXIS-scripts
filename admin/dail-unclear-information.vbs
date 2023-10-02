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
objExcel.Cells(1, 2).Value = STATS_counter
objExcel.Cells(2, 2).Value = STATS_manualtime
objExcel.Cells(3, 2).Value = STATS_counter * STATS_manualtime
objExcel.Cells(4, 2).Value = timer - start_time
objExcel.Cells(5, 2).Value = ((STATS_counter * STATS_manualtime) - (timer - start_time)) / 60

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
const dail_maxis_case_number_const           = 0
const dail_worker_const	                    = 1
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

'To Do - add tracking of deleted dails once processing the list
'deleted_dails = 0	'establishing the value of the count for deleted deleted_dails

'Navigates to DAIL to pull DAIL messages
MAXIS_case_number = ""
CALL navigate_to_MAXIS_screen("DAIL", "PICK")
EMWriteScreen "_", 7, 39    'blank out ALL selection
'Selects CSES DAIL Type based on dialog selection
If CSES_messages = 1 Then EMWriteScreen "X", 10, 39    'Select CSES DAIL Type
'Selects INFO (HIRE) DAIL Type based on dialog selection
If HIRE_messages = 1 Then EMWriteScreen "X", 13, 39
transmit

For each worker in worker_array
    'To do - verify placement of this
    list_of_all_case_numbers = "~"
    'To do - think about handling for a situation where first case in DAIL is privileged so there wouldn't be any prior examples
    'Establish list of all PRIV cases so script knows to skip forward
    list_of_all_priv_case_numbers = "~"
    'EStablish list of all cases prior to PRIV so script knows where to start and where to skip forward
    list_of_all_prior_to_priv_case_numbers = "~"

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
            'To do - not sure if necessary
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

                    If InStr(priv_check, "YOU ARE NOT PRIVILEGED") Then
                        'Determine the case number and add to priv_check_list
                        'Determine case before priv case and add to prior to priv list
                        'Navigate back to DAIL/PICK for requested CSES or HIRE
                        'Enter the case number for prior to priv
                    Else
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
                                    Elseif app_status = "APPROVED" then
                                        EMReadScreen vers_number, 1, status_row, 23
                                        Call write_value_and_transmit(vers_number, 18, 54)
                                        Call write_value_and_transmit("FSSM", 19, 70)
                                    End if
                                    EmReadscreen reporting_status, 12, 8, 31
                                    EmReadscreen recertification_date, 8, 11, 31
                                    'Converts date from string to date
                                    recertification_date = DateAdd("m", 0, recertification_date)

                                    If InStr(reporting_status, "SIX MONTH") Then 
                                        sr_report_date = DateAdd("m", -6, recertification_date)
                                    Else
                                        sr_report_date = "N/A"
                                    End If
                                    ' MsgBox "Updating the case_details_array"
                                    'Update the array with new case details
                                    case_details_array(reporting_status_const, case_count) = trim(reporting_status)
                                    case_details_array(recertification_date_const, case_count) = trim(recertification_date)
                                    case_details_array(sr_report_date_const, case_count) = trim(sr_report_date)
                                End if
                            Else
                                case_details_array(reporting_status_const, case_count) = "N/A"
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
                    End if
                'If the MAXIS case number IS in the list of all case numbers, then it is not a new case number and no case details need to be gathered. It can work off the already collected case details.
                Else
                    ' MsgBox "NOT a new case number"
                        
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



' For each worker in worker_array
'     Call write_value_and_transmit(worker, 21, 6)
'     transmit 'transmits past not your dail message'

'     EMReadScreen number_of_dails, 1, 3, 67		'Reads where the count of DAILs is listed

'     DO
'         If number_of_dails = " " Then exit do		'if this space is blank the rest of the DAIL reading is skipped
'         dail_row = 6			'Because the script brings each new case to the top of the page, dail_row starts at 6.
'         DO
'             dail_type = ""
'             dail_msg = ""

'             'Determining if there is a new case number...
'             EMReadScreen new_case, 8, dail_row, 63
'             new_case = trim(new_case)
'             IF new_case <> "CASE NBR" THEN '...if there is NOT a new case number, the script will read the DAIL type, month, year, and message...
'                 Call write_value_and_transmit("T", dail_row, 3)
'             ELSEIF new_case = "CASE NBR" THEN
'                 '...if the script does find that there is a new case number (indicated by "CASE NBR"), it will write a "T" in the next row and transmit, bringing that case number to the top of your DAIL
'                 Call write_value_and_transmit("T", dail_row + 1, 3)
'             End if

'             dail_row = 6  'resetting the DAIL row '

'             EMReadScreen MAXIS_case_number, 8, dail_row - 1, 73
'             MAXIS_case_number = trim(MAXIS_case_number)
            
'             EMReadScreen dail_type, 4, dail_row, 6
            
'             EMReadScreen dail_month, 8, dail_row, 11
            
'             EMReadScreen dail_msg, 61, dail_row, 20
'             INFC_dail_msg = InStr(dail_msg, "INFC")

'             'Increment the stats counter
'             stats_counter = stats_counter + 1
            
'             If instr(dail_type,"HIRE") or (instr(dail_type, "CSES") and INFC_dail_msg = 0) Then  
'                 'To do - any issues with using actual count instead of excel_row_const
'                 ReDim Preserve DAIL_array(14, dail_count)	'This resizes the array based on the number of rows in the Excel File'
'                 ' TO DO - Only adding data from DAIL message to array, that is why not all constants are included
'                 DAIL_array(worker_const,	           DAIL_count) = trim(worker)
'                 DAIL_array(maxis_case_number_const,    DAIL_count) = right("00000000" & MAXIS_case_number, 8) 'outputs in 8 digits format
'                 DAIL_array(dail_type_const, 	       DAIL_count) = trim(dail_type)
'                 DAIL_array(dail_month_const, 		   DAIL_count) = trim(dail_month)
'                 DAIL_array(dail_msg_const, 		       DAIL_count) = trim(dail_msg)
'                 DAIL_array(excel_row_hire_const,       DAIL_count) = excel_row_hire
'                 DAIL_array(excel_row_cses_const, 	   DAIL_count) = excel_row_cses
'                 DAIL_count = DAIL_count + 1

'                 'add the data from DAIL to Excel
'                 If instr(dail_type,"HIRE") Then
'                     objExcel.Worksheets("HIRE").Activate
'                     objExcel.Cells(excel_row_hire, 1).Value = trim(worker)
'                     objExcel.Cells(excel_row_hire, 2).Value = trim(MAXIS_case_number)
'                     objExcel.Cells(excel_row_hire, 3).Value = trim(dail_type)
'                     objExcel.Cells(excel_row_hire, 4).Value = trim(dail_month)
'                     objExcel.Cells(excel_row_hire, 5).Value = trim(dail_msg)
'                     excel_row_hire = excel_row_hire + 1
'                     'Adding MAXIS case number to case number string
'                     'TO DO - verify functionality/need
'                     all_case_numbers_array = trim(all_case_numbers_array & MAXIS_case_number & "*") 
'                 End If

'                 If instr(dail_type,"CSES") Then
'                     objExcel.Worksheets("CSES").Activate
'                     objExcel.Cells(excel_row_cses, 1).Value = trim(worker)
'                     objExcel.Cells(excel_row_cses, 2).Value = trim(MAXIS_case_number)
'                     objExcel.Cells(excel_row_cses, 3).Value = trim(dail_type)
'                     objExcel.Cells(excel_row_cses, 4).Value = trim(dail_month)
'                     objExcel.Cells(excel_row_cses, 5).Value = trim(dail_msg)
'                     excel_row_cses = excel_row_cses + 1
'                     'Adding MAXIS case number to case number string
'                     'TO DO - verify if this is correct, necessary
'                     all_case_numbers_array = trim(all_case_numbers_array & MAXIS_case_number & "*") 
'                 End If
'             End if

'             dail_row = dail_row + 1
            
'             'TO DO - this is from DAIL decimator. Appears to handle for NAT errors. Is it needed?
'             'EMReadScreen message_error, 11, 24, 2		'Cases can also NAT out for whatever reason if the no messages instruction comes up.
'             'If message_error = "NO MESSAGES" then exit do

'             '...going to the next page if necessary
'             EMReadScreen next_dail_check, 4, dail_row, 4
'             If trim(next_dail_check) = "" then
'                 PF8
'                 EMReadScreen last_page_check, 21, 24, 2
'                 'DAIL/PICK when searching for specific DAIL types has message check of NO MESSAGES TYPE vs. NO MESSAGES WORK (for ALL DAIL/PICK selection).
'                 If last_page_check = "THIS IS THE LAST PAGE" or last_page_check = "NO MESSAGES TYPE" then
'                     all_done = true
'                     exit do
'                 Else
'                     dail_row = 6
'                 End if
'             End if
'         LOOP
'         IF all_done = true THEN exit do
'     LOOP
' Next

' Call back_to_SELF
' Call MAXIS_footer_month_confirmation

' For item = 0 to Ubound(DAIL_array, 2)
'     'Resets the dail_type so that it can switch between CSES and HIRE messages
'     'To do - double-check this is actually resetting information
'     dail_type = DAIL_array(dail_type_const, item)
'     MAXIS_case_number = DAIL_array(MAXIS_case_number_const, item)
'     dail_month = DAIL_array(dail_month_const, item)
'     worker = DAIL_array(worker_const, item)
'     dail_msg = DAIL_array(dail_msg_const, item)

'     Call navigate_to_MAXIS_screen_review_PRIV("CASE", "CURR", is_this_priv)
'     If is_this_priv = True then
'         DAIL_array(processing_notes_const, item) = DAIL_array(processing_notes_const, item) & "Privileged Case"
'     Else
'         EmReadscreen worker_county, 4, 21, 14
'         If worker_county <> worker_county_code then
'             DAIL_array(processing_notes_const, item) = DAIL_array(processing_notes_const, item) & "Out-of-County Case"
'         Else
'             Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type, case_status, list_active_programs, list_pending_programs)
'             'SNAP Information
'             'To Do - would there be instances when we would consider a case with Snap status other than active?
'             If snap_status <> "ACTIVE" then 
'                 DAIL_array(action_req_const, item) = True
'                 DAIL_array(reporting_status_const, item) = "N/A"
'                 DAIL_array(recertification_date_const, item) = "N/A"
'                 DAIL_array(sr_report_date_const, item) = "N/A"
'                 DAIL_array(renewal_month_determination_const, item) = "N/A"

'             End If

'             'If other programs are active/pending then no notice is necessary
'             If  ga_case = True OR _
'                 msa_case = True OR _
'                 mfip_case = True OR _
'                 dwp_case = True OR _
'                 grh_case = True OR _
'                 ma_case = True OR _
'                 msp_case = True then
'                     DAIL_array(other_programs_present_const, item) = True
'                     DAIL_array(action_req_const, item) = True
'             Else
'                 DAIL_array(other_programs_present_const, item) = False
'             End if

'             DAIL_array(snap_status_const, item) = snap_status


'             If snap_status = "ACTIVE" then
'                 Call MAXIS_background_check
'                 Call navigate_to_MAXIS_screen("ELIG", "FS  ")
'                 EMReadScreen no_SNAP, 10, 24, 2
'                 If no_SNAP = "NO VERSION" then						'NO SNAP version means no determination
'                     DAIL_array(processing_notes_const, item) = DAIL_array(processing_notes_const, item) & "No version of SNAP exists for " & MAXIS_footer_month & "/" & MAXIS_footer_year & ". "
'                     DAIL_array(action_req_const, item) = True
'                 Else

'                     EMWriteScreen "99", 19, 78
'                     transmit
'                     'This brings up the FS versions of eligibility results to search for approved versions
'                     status_row = 7
'                     Do
'                         EMReadScreen app_status, 8, status_row, 50
'                         app_status = trim(app_status)
'                         If app_status = "" then
'                             PF3
'                             exit do 	'if end of the list is reached then exits the do loop
'                         End if
'                         If app_status = "UNAPPROV" Then status_row = status_row + 1
'                     Loop until app_status = "APPROVED" or app_status = ""

'                     If app_status = "" or app_status <> "APPROVED" then
'                         DAIL_array(processing_notes_const, item) = DAIL_array(processing_notes_const, item) & "No approved eligibility results exists in " & MAXIS_footer_month & "/" & MAXIS_footer_year & ". "
'                         DAIL_array(action_req_const, item) = True
'                     Elseif app_status = "APPROVED" then
'                         EMReadScreen vers_number, 1, status_row, 23
'                         Call write_value_and_transmit(vers_number, 18, 54)
'                         Call write_value_and_transmit("FSSM", 19, 70)
'                     End if
'                     EmReadscreen reporting_status, 12, 8, 31
'                     EmReadscreen recertification_date, 8, 11, 31
'                     'Converts date from string to date
'                     recertification_date = DateAdd("m", 0, recertification_date)
'                     If InStr(reporting_status, "SIX MONTH") Then 
'                         ' MsgBox reporting_status
'                         ' MsgBox recertification_date    
'                         sr_report_date = DateAdd("m", -6, recertification_date)
'                     Else
'                         sr_report_date = "N/A"
'                     End If
'                     'To Do - verify that this is working properly
'                     'TO do - check on how to handle if SR or recertification is in CM
'                     'Add validation to determine if renewal/SR certification dates align with corresponding DAIL month
'                     'CSES - determine if dail_month = recertification OR dail_month = SR report date. If this is true, even in past, then should be flagged
'                     'Convert dail_month to date in MM/DD/YYYY format for comparison purposes
'                     dail_month = Left(dail_month, 2) & "/01/" & Right(dail_month, 2)
'                     dail_month = DateAdd("m", 0, dail_month)

'                     If dail_type = "CSES" Then
'                         If DateDiff("m", dail_month, recertification_date) = 0 Then
'                             renewal_month_determination = "Recertification month equals DAIL month."
'                         Else 
'                             renewal_month_determination = "Recertification month does not equal DAIL month."
'                         End If
'                     ElseIf dail_type = "HIRE" Then
'                         If DateDiff("m", dail_month, recertification_date) = 1 Then
'                             renewal_month_determination = "Recertification month equals DAIL month + 1."
'                         Else 
'                             renewal_month_determination = "Recertification month does not equal DAIL month + 1."
'                         End If
'                     End If

'                     If sr_report_date <> "N/A" Then
'                         If dail_type = "CSES" Then
'                             If DateDiff("m", dail_month, sr_report_date) = 0 Then
'                                 renewal_month_determination = "SR Report Date month equals DAIL month." & " " & renewal_month_determination
'                             Else 
'                                 renewal_month_determination = "SR Report Date month does not equal DAIL month." & " " & renewal_month_determination
'                             End If
'                         ElseIf dail_type = "HIRE" Then
'                             If DateDiff("m", dail_month, sr_report_date) = 1 Then
'                                 renewal_month_determination = "SR Report Date month equals DAIL month + 1." & " " & renewal_month_determination
'                             Else 
'                                 renewal_month_determination = "SR Report Date month does not equal DAIL month + 1." & " " & renewal_month_determination
'                             End If
'                         End If
'                     End If
                    
'                     'Determine if action is required due to the DAIL message aligning with SR report date or recertification date regardless of CSES or HIRE message
'                     If instr(renewal_month_determination, "equals") Then
'                         renewal_month_action = True
'                     Else
'                         renewal_month_action = False
'                     End If    
            
'                     DAIL_array(reporting_status_const, item) = trim(reporting_status)
'                     DAIL_array(recertification_date_const, item) = trim(recertification_date)
'                     DAIL_array(sr_report_date_const, item) = trim(sr_report_date)
'                     DAIL_array(renewal_month_determination_const, item) = trim(renewal_month_determination)
'                 End if
'             Else
'                 DAIL_array(reporting_status_const, item) = "N/A"
'             End if

'             'Determine if action_req is true (don't act on DAIL message) or if action_req is false (act on DAIL message)
'             If DAIL_array(snap_status_const, item) = "ACTIVE" AND DAIL_array(other_programs_present_const, item) = False AND DAIL_array(reporting_status_const, item) = "SIX MONTH" AND renewal_month_action = False then
'                 DAIL_array(action_req_const, item) = False
'             Else
'                 DAIL_array(action_req_const, item) = True
'             End if
'             reporting_status = ""   'blanking out variable
'         End if
'     End if

'     'Updates the corresponding Excel sheet (HIRE or CSES) with data about each case
'     If instr(dail_type,"HIRE") Then
'         objExcel.Worksheets("HIRE").Activate
'         objExcel.Cells(DAIL_array(excel_row_hire_const, item), 6).Value = DAIL_array(snap_status_const, item)
'         objExcel.Cells(DAIL_array(excel_row_hire_const, item), 7).Value = DAIL_array(other_programs_present_const, item)
'         objExcel.Cells(DAIL_array(excel_row_hire_const, item), 8).Value = DAIL_array(reporting_status_const, item)
'         objExcel.Cells(DAIL_array(excel_row_hire_const, item), 9).Value = DAIL_array(sr_report_date_const, item)
'         objExcel.Cells(DAIL_array(excel_row_hire_const, item), 10).Value = DAIL_array(recertification_date_const, item)
'         objExcel.Cells(DAIL_array(excel_row_hire_const, item), 11).Value = DAIL_array(renewal_month_determination_const, item)
'         objExcel.Cells(DAIL_array(excel_row_hire_const, item), 12).Value = DAIL_array(action_req_const, item)
'         objExcel.Cells(DAIL_array(excel_row_hire_const, item), 13).Value = DAIL_array(processing_notes_const, item)
'     End If

'     If instr(dail_type,"CSES") Then
'         objExcel.Worksheets("CSES").Activate
'         objExcel.Cells(DAIL_array(excel_row_cses_const, item), 6).Value = DAIL_array(snap_status_const, item)
'         objExcel.Cells(DAIL_array(excel_row_cses_const, item), 7).Value = DAIL_array(other_programs_present_const, item)
'         objExcel.Cells(DAIL_array(excel_row_cses_const, item), 8).Value = DAIL_array(reporting_status_const, item)
'         objExcel.Cells(DAIL_array(excel_row_cses_const, item), 9).Value = DAIL_array(sr_report_date_const, item)
'         objExcel.Cells(DAIL_array(excel_row_cses_const, item), 10).Value = DAIL_array(recertification_date_const, item)
'         objExcel.Cells(DAIL_array(excel_row_cses_const, item), 11).Value = DAIL_array(renewal_month_determination_const, item)
'         objExcel.Cells(DAIL_array(excel_row_cses_const, item), 12).Value = DAIL_array(action_req_const, item)
'         objExcel.Cells(DAIL_array(excel_row_cses_const, item), 13).Value = DAIL_array(processing_notes_const, item)
'     End If
' Next

' report_month = CM_mo & "-20" & CM_yr
' 'To DO - confirm file path and title is correct
' objExcel.ActiveWorkbook.SaveAs "T:\Eligibility Support\Restricted\QI - Quality Improvement\BZ scripts project\Projects\Unclear Information\" & report_month & " Unclear Information - DAIL Messages.xlsx" 
' objExcel.ActiveWorkbook.Close
' objExcel.Application.Quit
' objExcel.Quit

' script_end_procedure("Success! Please review the list created for accuracy.")

' 'Script actions if creating a new Excel list option is selected
' If script_action = "Process existing Excel list" Then

'     'Validation to ensure that processing correct Excel spreadsheet, otherwise script ends
'     'To do - should validation for Excel name be in the dialog instead?
'     If InStr(file_selection_path, "Unclear Information - DAIL Messages") Then 
'         Call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)
'     Else
'         script_end_procedure("The script must process an Unclear Information Excel list. The selected Excel file is not an Unclear Information Excel list. The script will now end.")
'     End If

'     'To do - should this be within the do loop?
'     objExcel.Worksheets("CSES").Activate

'     'Navigate to DAIL/PICK, blank out selection, and select CSES DAIL type
'     'To do - consider if navigation to DAIL/PICK should be within DO LOOP
'     CALL navigate_to_MAXIS_screen("DAIL", "PICK")
'     EMWriteScreen "_", 7, 39
'     Call write_value_and_transmit("X", 10, 39)

'     'Set initial Excel row value to iterate through
'     excel_row = 2

'     'Utilize Do Loop to iterate through each row of the sheet until an empty row is found
'     Do
'     'Reach through each row and set variables

'     'Reading case number from Excel Sheet
'     MAXIS_case_number = objExcel.cells(excel_row, 2).Value
'     MAXIS_case_number = trim(MAXIS_case_number)
'     'End do loop if the MAXIS_case_number is blank, script has reached end of sheet
'     IF MAXIS_case_number = "" THEN 
'         MsgBox "End of CSES Sheet has been reached."
'         EXIT DO
'     Else
'         'Set variables from Excel sheet
'         worker                          = objExcel.cells(excel_row, 1).Value 
'         dail_type                       = objExcel.cells(excel_row, 3).Value    
'         dail_month                      = objExcel.cells(excel_row, 4).Value 
'         dail_msg                        = objExcel.cells(excel_row, 5).Value        
'         snap_status                     = objExcel.cells(excel_row, 6).Value    
'         other_programs_present          = objExcel.cells(excel_row, 7).Value    
'         reporting_status                = objExcel.cells(excel_row, 8).Value  
'         sr_report_date                  = objExcel.cells(excel_row, 9).Value 
'         recertification_date            = objExcel.cells(excel_row, 10).Value    
'         renewal_month_determination     = objExcel.cells(excel_row, 11).Value
'         action_req                      = objExcel.cells(excel_row, 12).Value    
'         processing_notes                = objExcel.cells(excel_row, 13).Value

'         'If action is required, the row is skipped and processing notes is updated accordingly.
'         If action_req = TRUE Then
'             objExcel.Cells(excel_row, 13).Value = "Action required. DAIL message not processed."
'             excel_row = excel_row + 1

'         'If action is not required, then script processes DAIL message accordingly
'         ElseIf action_req = FALSE Then
'             'Clear out MAXIS case number if present
'             EMWriteScreen "________", 20, 38
'             'Write MAXIS_case_number to find corresponding DAIL message and transmits
'             EMWriteScreen MAXIS_case_number, 20, 38
'             transmit

'             'Reads where the count of DAILs is listed
'             EMReadScreen number_of_dails, 1, 3, 67		
            
'             'If the are no DAIL messages for the specific case number, it will exit the do loop and move to the next row
'             If number_of_dails = " " Then
'                 objExcel.Cells(excel_row, 13).Value = "Unable to find DAIL message. It may have been deleted."
'             Else
'                 'Do loop to find corresponding DAIL message
'                 'Set the starting dail_row, read the message, compare the messages. If they match then process, if they don't, add one to dail row and go to next
'                 dail_row = 6

'                 Do
'                     'Checking if DAIL has moved to a new case number, if that's the case then it should exit the do loop and move to the next row in the spreadsheet. Otherwise it can move the message to the top and read the message.
'                     EMReadScreen new_case, 8, dail_row, 63
'                     new_case = trim(new_case)
'                     IF new_case = "CASE NBR" THEN Exit Do
                
'                     'Make sure that top message is at the top
'                     Call write_value_and_transmit("T", dail_row, 3)

'                     'To do - verify this logic works
'                     'Reset dail_row to 6 since the message has been moved to the top
'                     dail_row = 6

'                     'Read the DAIL message month to check if it is a match
'                     EMReadScreen dail_month_check, 8, dail_row, 11
                    
'                     'Read the DAIL message month to check if it is a match
'                     EMReadScreen dail_msg_check, 61, dail_row, 20

'                     'If the dail month and dail message match, then open the message to read full message info
'                     If dail_month_check = dail_month and dail_msg_check = dail_msg Then
'                         'Process changes depending on the CSES message type
'                         If InStr(dail_msg, "DISB CS (TYPE 36)")

'                         'Open the full dail message
'                         Call write_value_and_transmit("X", dail_row, 3)

'                         'Reads the full message and identifies the PMI numbers
                        
'                         'Exit do loop once the match has been found and processed
'                         Exit Do

'                     Else
'                         'If the dail_msg and/or dail_month is different, then the script should go to next dail_row and restart do loop
'                         dail_row = dail_row + 1       
'                         'Go to next message to check, if it is a new case number then exit do
'                     End If

'                 Loop



'             End If
            

'             'Do loop to find 

'             'After processing is complete, add 1 to excel_row to go to next row
'             excel_row = excel_row + 1
'         'Add in handling in case there is no action required determination. May be unnecessary.
'         'To do - remove if unnecessary
'         Else
'             objExcel.Cells(excel_row, 13).Value = "No action required determination. DAIL message not processed"
'             excel_row = excel_row + 1
'         End If

'         'Increment excel_row to go to next row
'         excel_row = excel_row + 1


'     End If

'     Loop
' End If
