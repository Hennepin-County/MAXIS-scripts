'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "bulk-applications.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 335                      'manual run time in seconds
STATS_denomination = "C"       			   'C is for each CASE
'END OF stats block==============================================================================================
'TODO Add VGO things

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

' 'Reading Locally held FuncLib in leiu of issues with connecting to GitHub
' Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
' Set fso_command = run_another_script_fso.OpenTextFile("C:\MAXIS-scripts\MASTER FUNCTIONS LIBRARY.vbs")
' text_from_the_other_script = fso_command.ReadAll
' fso_command.Close
' Execute text_from_the_other_script

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
CALL changelog_update("02/05/2018", "Initial version.", "MiKayla Handley, Hennepin County")


'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'----------------------------------------------------------------------------------------------------Custom function
function start_a_new_spec_memo_and_continue(success_var)
'--- This function navigates user to SPEC/MEMO and starts a new SPEC/MEMO, selecting client, AREP, and SWKR if appropriate
'===== Keywords: MAXIS, notice, navigate, edit
  success_var = True
	call navigate_to_MAXIS_screen("SPEC", "MEMO")				'Navigating to SPEC/MEMO

	PF5															'Creates a new MEMO. If it's unable the script will stop.
	EMReadScreen memo_display_check, 12, 2, 33
	If memo_display_check = "Memo Display" then success_var = False

	'Checking for an AREP. If there's an AREP it'll navigate to STAT/AREP, check to see if the forms go to the AREP. If they do, it'll write X's in those fields below.
	row = 4                             'Defining row and col for the search feature.
	col = 1
	EMSearch "ALTREP", row, col         'Row and col are variables which change from their above declarations if "ALTREP" string is found.
	IF row > 4 THEN                     'If it isn't 4, that means it was found.
	    arep_row = row                                          'Logs the row it found the ALTREP string as arep_row
	    call navigate_to_MAXIS_screen("STAT", "AREP")           'Navigates to STAT/AREP to check and see if forms go to the AREP
	    EMReadscreen forms_to_arep, 1, 10, 45                   'Reads for the "Forms to AREP?" Y/N response on the panel.
	    call navigate_to_MAXIS_screen("SPEC", "MEMO")           'Navigates back to SPEC/MEMO
	    PF5                                                     'PF5s again to initiate the new memo process
	END IF
	'Checking for SWKR
	row = 4                             'Defining row and col for the search feature.
	col = 1
	EMSearch "SOCWKR", row, col         'Row and col are variables which change from their above declarations if "SOCWKR" string is found.
	IF row > 4 THEN                     'If it isn't 4, that means it was found.
	    swkr_row = row                                          'Logs the row it found the SOCWKR string as swkr_row
	    call navigate_to_MAXIS_screen("STAT", "SWKR")         'Navigates to STAT/SWKR to check and see if forms go to the SWKR
	    EMReadscreen forms_to_swkr, 1, 15, 63                'Reads for the "Forms to SWKR?" Y/N response on the panel.
	    call navigate_to_MAXIS_screen("SPEC", "MEMO")         'Navigates back to SPEC/MEMO
	    PF5                                           'PF5s again to initiate the new memo process
	END IF
	EMWriteScreen "x", 5, 12                                        'Initiates new memo to client
	IF forms_to_arep = "Y" THEN EMWriteScreen "x", arep_row, 12     'If forms_to_arep was "Y" (see above) it puts an X on the row ALTREP was found.
	IF forms_to_swkr = "Y" THEN EMWriteScreen "x", swkr_row, 12     'If forms_to_arep was "Y" (see above) it puts an X on the row ALTREP was found.
	transmit                                                        'Transmits to start the memo writing process
end function

function ONLY_create_MAXIS_friendly_date(date_variable)
'--- This function creates a MM DD YY date.
'~~~~~ date_variable: the name of the variable to output
	var_month = datepart("m", date_variable)
	If len(var_month) = 1 then var_month = "0" & var_month
	var_day = datepart("d", date_variable)
	If len(var_day) = 1 then var_day = "0" & var_day
	var_year = datepart("yyyy", date_variable)
	var_year = right(var_year, 2)
	date_variable = var_month &"/" & var_day & "/" & var_year
end function


Function check_pnd2_for_denial(coded_denial, SNAP_pnd2_code, cash_pnd2_code, emer_pnd2_code)
  Call navigate_to_MAXIS_screen("REPT", "PND2")
  row = 7
  col = 5
  EMSearch MAXIS_case_number, row, col      'finding correct case to check PND2 codes

  IF SNAP_check = checked Then
  	EMReadScreen SNAP_pnd2_code, 1, row, 62
  	IF SNAP_pnd2_code = "R" THEN coded_denial = coded_denial & " SNAP withdrawn on PND2."
  	IF SNAP_pnd2_code = "I" THEN coded_denial = coded_denial & " SNAP application incomplete, denied on PND2."
  	IF SNAP_pnd2_code = "_" THEN
  		'If SNAP is selected by the user but the SNAP column is empty on PND2, the script is going to look on the next row for ADDITIONAL APP...
  		EMReadScreen additional_maxis_application, 20, row + 1, 16
  		additional_maxis_application = trim(additional_maxis_application)
  		IF InStr(additional_maxis_application, "ADDITIONAL") <> 0 THEN
  			EMReadScreen SNAP_pnd2_code, 1, row + 1, 62
  			IF SNAP_pnd2_code = "R" THEN coded_denial = coded_denial & " SNAP withdrawn on PND2."
  			IF SNAP_pnd2_code = "I" THEN coded_denial = coded_denial & " SNAP application incomplete, denied on PND2."
  		END IF
  	END IF
  END IF
  IF cash_check = checked Then
  	EMReadScreen cash_pnd2_code, 1, row, 54
  	IF cash_pnd2_code = "R" THEN coded_denial = coded_denial & " CASH withdrawn on PND2."
  	IF cash_pnd2_code = "I" THEN coded_denial = coded_denial & " CASH application incomplete, denied on PND2."
  	IF cash_pnd2_code = "_" THEN
  		'If CASH is selected by the user but the CASH column is empty on PND2, the script is going to look on the next row for ADDITIONAL APP...
  		EMReadScreen additional_maxis_application, 20, row + 1, 16
  		additional_maxis_application = trim(additional_maxis_application)
  		IF InStr(additional_maxis_application, "ADDITIONAL") <> 0 THEN
  			EMReadScreen cash_pnd2_code, 1, row + 1, 54
  			IF cash_pnd2_code = "R" THEN coded_denial = coded_denial & " CASH withdrawn on PND2."
  			IF cash_pnd2_code = "I" THEN coded_denial = coded_denial & " CASH application incomplete, denied on PND2."
  		END IF
  	END IF
  END IF
End function


function convert_to_mainframe_date(date_var, yr_len)
    'This will change a variable to mm/dd/yy or mm/dd/yyyy format for comparison to dates in MX
    'yr_len should be a number - either 2 or 4
    'MsgBox date_var
    month_to_use = DatePart("m", date_var)
    month_to_use = right("00" & month_to_use, 2)

    day_to_use = DatePart("d", date_var)
    day_to_use = right("00" & day_to_use, 2)

    year_to_use = DatePart("yyyy", date_var)
    year_to_use = right(year_to_use, yr_len)

    date_var = month_to_use & "/" & day_to_use & "/" & year_to_use
end function

function confirm_memo_waiting(confirmation_var)
    memo_row = 7

    today_date = date
    Call convert_to_mainframe_date(today_date, 2)

    Do
        EMReadScreen create_date, 8, memo_row, 19                 'Reading the date of each memo and the status
        EMReadScreen print_status, 7, memo_row, 67
        'MsgBox print_status
        If create_date = today_date AND print_status = "Waiting" Then   'MEMOs created today and still waiting is likely our MEMO.
            confirmation_var = "Y"             'If we've found this then no reason to keep looking.
            successful_notices = successful_notices + 1
            'MsgBox ALL_PENDING_CASES_ARRAY(notc_confirm, case_entry)                 'For statistical purposes
            Exit Do
        End If
        memo_row = memo_row + 1           'Looking at next row'
    Loop Until create_date = "        "
end function

function go_to_top_of_notes()
    Do
        PF7
        EMReadScreen top_of_notes_check, 10, 24, 14
    Loop until top_of_notes_check = "FIRST PAGE"
end function

'THE SCRIPT-------------------------------------------------------------------------------------------------------------------------
EMConnect ""		'Connects to BlueZone
'Grabbing the worker's X number.
CALL find_variable("User: ", worker_number, 7)

'Starting the query start time (for the query runtime at the end)
query_start_time = timer

'if user is not Hennepin County - the script will end. Process is not approved for other counties
'------------------------------------------------------------------------------------------------------establishing date variables
MAXIS_footer_month = CM_plus_1_mo
MAXIS_footer_year = CM_plus_1_yr

CM_minus_1_mo = right("0" & DatePart("m", DateAdd("m", -1, date)), 2)
CM_minus_1_yr = right(DatePart("yyyy", DateAdd("m", -1, date)), 2)

current_date = date
Call ONLY_create_MAXIS_friendly_date(current_date)			'reformatting the dates to be MM/DD/YY format to measure against the panel dates

'Opens the current day's list
'dialog and dialog DO...Loop
Do
	Do
		'The dialog is defined in the loop as it can change as buttons are pressed
        BeginDialog file_select_dialog, 0, 0, 316, 175, "Select the source file"
          EditBox 5, 125, 260, 15, file_selection_path
          ButtonGroup ButtonPressed
            PushButton 270, 125, 40, 15, "Browse...", select_a_file_button
            OkButton 205, 155, 50, 15
            CancelButton 260, 155, 50, 15
          Text 5, 5, 305, 25, "This script will send Appointment Notices and NOMIs or update for denials when no interview has been completed. Once an interview has taken place, this script no longer takes action on the case."
          Text 5, 35, 255, 10, "Cases with an interview completed should have the interview listed on PROG."
          Text 5, 50, 310, 10, "An Appointment Notice will be sent on any case without a case note of appointment notice sent."
          Text 5, 65, 300, 10, "A NOMI will be sent once the appointment date indicated on Appointment Notice has passed."
          Text 5, 80, 305, 20, "A denial will be indicated when a case reaches day 30 (unless the NOMI did not go out until day 30 or after)."
          Text 10, 105, 295, 15, "Click the BROWSE button and select the BOBI report for today. Once selected, click 'OK'. There will be no additional input needed until the script run is complete."
          Text 5, 150, 160, 20, "Reminder, do not use Excel during the time the script is running. The script needs to use Excel."
        EndDialog

		err_msg = ""
		Dialog file_select_dialog
		If ButtonPressed = cancel then stopscript
		If ButtonPressed = select_a_file_button then
			If file_selection_path <> "" then 'This is handling for if the BROWSE button is pushed more than once'
				objExcel.Quit 'Closing the Excel file that was opened on the first push'
				objExcel = "" 	'Blanks out the previous file path'
			End If
			call file_selection_system_dialog(file_selection_path, ".xlsx") 'allows the user to select the file'
		End If
		If file_selection_path = "" then err_msg = err_msg & vbNewLine & "Use the Browse Button to select the file that has your client data"
		If IsNumeric(MAXIS_footer_month) = False or len(MAXIS_footer_month) > 2 or len(MAXIS_footer_month) < 2 then err_msg = err_msg & vbNewLine & "* Enter a valid footer month."
  		If IsNumeric(MAXIS_footer_year) = False or len(MAXIS_footer_year) > 2 or len(MAXIS_footer_year) < 2 then err_msg = err_msg & vbNewLine & "* Enter a valid footer year."
		If err_msg <> "" Then MsgBox err_msg
	Loop until err_msg = ""
	If objExcel = "" Then call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'
	If err_msg <> "" Then MsgBox err_msg
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

If trim(worker_signature) = "" Then
    worker_signature = InputBox("How would you like to sign you case notes:", "Worker Signature")
End If
'Activates worksheet based on user selection
objExcel.worksheets("Report 1").Activate

'Each of the case numbers will be stored at this position'
const case_number           = 0             'case_nbr_col
const excel_row             = 1
const client_name			= 2             'case_name_col
const program_group_ID		= 3             '
const worker_ID		   		= 4             ''
const program_status		= 5             ''
const priv_case             = 6
const out_of_co             = 7
const written_lang          = 8
const SNAP_status           = 9
const CASH_status           = 10
const application_date      = 11
const interview_date    	= 12
const appt_notc_sent        = 13 'dates'
const appt_notc_confirm     = 14
const nomi_sent             = 15 'dates'
const nomi_confirm          = 16
const deny_day30			= 17
const deny_memo_confirm     = 18
const need_appt_notc        = 19
const need_nomi             = 20
const appointment_date		= 21
const next_action_needed    = 22
const on_working_list       = 23
const questionable_intv     = 24
const take_action_today     = 25
const need_face_to_face     = 26
const error_notes 			= 27

'Constants for columns in the working excel sheet
const worker_id_col         = 1         'worker_ID
const case_nbr_col          = 2         'case_number
const case_name_col         = 3         'client_name
const snap_stat_col         = 4         'SNAP_status
const cash_stat_col         = 5         'CASH_status
const app_date_col          = 6         'application_date
const intvw_date_col        = 7         'interview_date
const quest_intvw_date_col  = 8         ''
const ftof_still_need_col   = 9         ''
const appt_notc_date_col    = 10        ''
const appt_date_col         = 11        ''
const appt_notc_confirm_col = 12        ''
const nomi_date_col         = 13        ''
const nomi_confirm_col      = 14        ''
const need_deny_col         = 15        ''
const deny_notc_confirm_col = 16        ''
const next_action_col       = 17        ''
const day_30_col            = 18
const correct_need_col      = 19        ''
const action_worker_col     = 20        ''
const action_sup_col        = 21        ''
const email_sent_col        = 22        ''

Dim TODAYS_CASES_ARRAY()
ReDim TODAYS_CASES_ARRAY(error_notes, 0)

todays_cases_list = "*"
case_entry = 0
row = 5
'Goes through the list, and creates an array of all cases - removing duplicates and removing cases with an interview date already listed
Do
    'If trim(objExcel.Cells(row, 8).value) = "" Then
        anything_number = trim(objExcel.Cells(row, 3).value)
        'MsgBox anything_number
        If instr(todays_cases_list, "*" & anything_number & "*") = 0 then
            'MsgBox anything_number
            todays_cases_list = todays_cases_list & anything_number & "*"
            ReDim Preserve TODAYS_CASES_ARRAY(error_notes, case_entry)
            TODAYS_CASES_ARRAY(worker_ID, case_entry) = trim(objExcel.Cells(row, 2).value)
            TODAYS_CASES_ARRAY(case_number, case_entry) = trim(objExcel.Cells(row, 3).value)
            TODAYS_CASES_ARRAY(excel_row, case_entry) = row
            TODAYS_CASES_ARRAY(client_name, case_entry) = trim(objExcel.cells(row, 4).value) 'storing all of the excel information
            TODAYS_CASES_ARRAY(application_date, case_entry) = trim(objExcel.cells(row, 7).value)
            TODAYS_CASES_ARRAY(interview_date, case_entry) = trim(objExcel.cells(row, 8).value)
            TODAYS_CASES_ARRAY(on_working_list, case_entry) = FALSE

            current_number = anything_number
            case_entry = case_entry + 1
        ElseIf anything_number = current_number Then
            If trim(objExcel.cells(row, 8).value) = "" Then TODAYS_CASES_ARRAY(interview_date, case_entry-1) = ""
        End If
        stats_counter = stats_counter + 1
    'End If
    row = row + 1
    next_case_number = trim(objExcel.Cells(row, 3).Value)
loop until next_case_number = ""

objExcel.quit
'MsgBox case_entry
'Opens the working excel spreadsheet.
working_excel_file_path = "T:\Eligibility Support\Restricted\QI - Quality Improvement\BZ scripts project\Projects\On Demand Waiver\Files for testing new application rewrite\Working Excel.xlsx"
'Opens Excel file here, as it needs to populate the dialog with the details from the spreadsheet.
call excel_open(working_excel_file_path, True, True, ObjWorkExcel, objWorkbook)

Dim ALL_PENDING_CASES_ARRAY()
ReDim ALL_PENDING_CASES_ARRAY(error_notes, 0)

Dim CASES_NO_LONGER_WORKING()
ReDim CASES_NO_LONGER_WORKING(error_notes, 0)

case_entry = 0      'incrementer to add a case to ALL_PENDING_CASES_ARRAY
case_removed = 0    'incrementer to add a case to CASES_NO_LONGER_WORKING
row = 2

'This do loops through all of the cases that are already on the working sheet to see if we can find them in today's array
Do
    case_number_to_assess = trim(objWorkExcel.Cells(row, 2).Value)
    found_case_on_todays_list = FALSE
    If trim(case_number_to_assess) = "" Then Exit DO

    For each_case = 0 to UBound(TODAYS_CASES_ARRAY, 2)
        'MsgBox "Excel case number: " & case_number_to_assess & vbNewLine & "Array case number: " & TODAYS_CASES_ARRAY(case_number, each_case)
        If case_number_to_assess = TODAYS_CASES_ARRAY(case_number, each_case) Then
            TODAYS_CASES_ARRAY(on_working_list, each_case) = TRUE
            found_case_on_todays_list = TRUE
            'MsgBox "Excel case number: " & case_number_to_assess & vbNewLine & "Array case number: " & TODAYS_CASES_ARRAY(case_number, each_case)
            If TODAYS_CASES_ARRAY(interview_date, each_case) <> "" Then
                'Remove from working sheet and add to list of cases removed
                'MsgBox "Interview Date: " & TODAYS_CASES_ARRAY(interview_date, each_case)
                ReDim Preserve CASES_NO_LONGER_WORKING(error_notes, case_removed)
                CASES_NO_LONGER_WORKING(worker_ID, case_removed) = TODAYS_CASES_ARRAY(worker_ID, each_case)
                CASES_NO_LONGER_WORKING(case_number, case_removed) = TODAYS_CASES_ARRAY(case_number, each_case)
                CASES_NO_LONGER_WORKING(excel_row, case_removed) = row
                CASES_NO_LONGER_WORKING(client_name, case_removed) = TODAYS_CASES_ARRAY(client_name, each_case)
                CASES_NO_LONGER_WORKING(application_date, case_removed) = ObjWorkExcel.Cells(row, app_date_col)
                'CASES_NO_LONGER_WORKING(interview_date, case_removed) = ObjWorkExcel.Cells(row, intvw_date_col)
                CASES_NO_LONGER_WORKING(interview_date, case_removed) = TODAYS_CASES_ARRAY(interview_date, each_case)
                CASES_NO_LONGER_WORKING(CASH_status, case_removed) = ObjWorkExcel.Cells(row, cash_stat_col)
                CASES_NO_LONGER_WORKING(SNAP_status, case_removed) = ObjWorkExcel.Cells(row, snap_stat_col)

                CASES_NO_LONGER_WORKING(appt_notc_sent, case_removed) = ObjWorkExcel.Cells(row, appt_notc_date_col)
                CASES_NO_LONGER_WORKING(appt_notc_confirm, case_removed) = ObjWorkExcel.Cells(row, appt_notc_confirm_col)
                CASES_NO_LONGER_WORKING(appointment_date, case_removed) = ObjWorkExcel.Cells(row, appt_date_col)
                CASES_NO_LONGER_WORKING(nomi_sent, case_removed) = ObjWorkExcel.Cells(row, nomi_date_col)
                CASES_NO_LONGER_WORKING(nomi_confirm, case_removed) = ObjWorkExcel.Cells(row, nomi_confirm_col)
                CASES_NO_LONGER_WORKING(next_action_needed, case_removed) = ObjWorkExcel.Cells(row, next_action_col)
                CASES_NO_LONGER_WORKING(questionable_intv, case_removed) = ObjWorkExcel.Cells(row, quest_intvw_date_col)
                CASES_NO_LONGER_WORKING(need_face_to_face, case_removed) = ObjWorkExcel.Cells(row, ftof_still_need_col)

                CASES_NO_LONGER_WORKING(error_notes, case_removed) = "Interview Completed on " & TODAYS_CASES_ARRAY(interview_date, each_case)

                case_removed = case_removed + 1
                SET objRange = ObjWorkExcel.Cells(row, 1).EntireRow
                objRange.Delete
            ElseIf ObjWorkExcel.Cells(row, next_action_col) = "REMOVE FROM LIST" Then
                'MsgBox "REMOVE FROM LIST"
                ReDim Preserve CASES_NO_LONGER_WORKING(error_notes, case_removed)
                CASES_NO_LONGER_WORKING(worker_ID, case_removed) = ObjWorkExcel.Cells(row, worker_id_col)
                CASES_NO_LONGER_WORKING(case_number, case_removed) = ObjWorkExcel.Cells(row, case_nbr_col)
                CASES_NO_LONGER_WORKING(excel_row, case_removed) = row
                CASES_NO_LONGER_WORKING(client_name, case_removed) = ObjWorkExcel.Cells(row, case_name_col)
                CASES_NO_LONGER_WORKING(application_date, case_removed) = ObjWorkExcel.Cells(row, app_date_col)
                'CASES_NO_LONGER_WORKING(interview_date, case_removed) = ObjWorkExcel.Cells(row, intvw_date_col)
                CASES_NO_LONGER_WORKING(interview_date, case_removed) = ObjWorkExcel.Cells(row, intvw_date_col)
                CASES_NO_LONGER_WORKING(CASH_status, case_removed) = ObjWorkExcel.Cells(row, cash_stat_col)
                CASES_NO_LONGER_WORKING(SNAP_status, case_removed) = ObjWorkExcel.Cells(row, snap_stat_col)

                CASES_NO_LONGER_WORKING(appt_notc_sent, case_removed) = ObjWorkExcel.Cells(row, appt_notc_date_col)
                CASES_NO_LONGER_WORKING(appt_notc_confirm, case_removed) = ObjWorkExcel.Cells(row, appt_notc_confirm_col)
                CASES_NO_LONGER_WORKING(appointment_date, case_removed) = ObjWorkExcel.Cells(row, appt_date_col)
                CASES_NO_LONGER_WORKING(nomi_sent, case_removed) = ObjWorkExcel.Cells(row, nomi_date_col)
                CASES_NO_LONGER_WORKING(nomi_confirm, case_removed) = ObjWorkExcel.Cells(row, nomi_confirm_col)
                CASES_NO_LONGER_WORKING(next_action_needed, case_removed) = ObjWorkExcel.Cells(row, next_action_col)
                CASES_NO_LONGER_WORKING(questionable_intv, case_removed) = ObjWorkExcel.Cells(row, quest_intvw_date_col)
                CASES_NO_LONGER_WORKING(need_face_to_face, case_removed) = ObjWorkExcel.Cells(row, ftof_still_need_col)

                CASES_NO_LONGER_WORKING(error_notes, case_removed) = "No programs pending."

                'TODO figure out why case is not on the list any more add add to error notes
                'CASES_NO_LONGER_WORKING(error_notes, case_removed) = "Interview Completed on " & TODAYS_CASES_ARRAY(interview_date, case_entry)
                'MsgBox row
                case_removed = case_removed + 1
                SET objRange = ObjWorkExcel.Cells(row, 1).EntireRow
                objRange.Delete
            Else

                ReDim Preserve ALL_PENDING_CASES_ARRAY(error_notes, case_entry)
                ALL_PENDING_CASES_ARRAY(worker_ID, case_entry) = TODAYS_CASES_ARRAY(worker_ID, each_case)
                ALL_PENDING_CASES_ARRAY(case_number, case_entry) = TODAYS_CASES_ARRAY(case_number, each_case)
                ALL_PENDING_CASES_ARRAY(excel_row, case_entry) = row
                ALL_PENDING_CASES_ARRAY(client_name, case_entry) = TODAYS_CASES_ARRAY(client_name, each_case)
                ALL_PENDING_CASES_ARRAY(application_date, case_entry) = ObjWorkExcel.Cells(row, app_date_col)
                ALL_PENDING_CASES_ARRAY(interview_date, case_entry) = ObjWorkExcel.Cells(row, intvw_date_col)
                ALL_PENDING_CASES_ARRAY(CASH_status, case_entry) = ObjWorkExcel.Cells(row, cash_stat_col)
                ALL_PENDING_CASES_ARRAY(SNAP_status, case_entry) = ObjWorkExcel.Cells(row, snap_stat_col)

                ALL_PENDING_CASES_ARRAY(appt_notc_sent, case_entry) = ObjWorkExcel.Cells(row, appt_notc_date_col)
                ALL_PENDING_CASES_ARRAY(appt_notc_confirm, case_entry) = ObjWorkExcel.Cells(row, appt_notc_confirm_col)
                ALL_PENDING_CASES_ARRAY(appointment_date, case_entry) = ObjWorkExcel.Cells(row, appt_date_col)
                ALL_PENDING_CASES_ARRAY(nomi_sent, case_entry) = ObjWorkExcel.Cells(row, nomi_date_col)
                ALL_PENDING_CASES_ARRAY(nomi_confirm, case_entry) = ObjWorkExcel.Cells(row, nomi_confirm_col)
                ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = ObjWorkExcel.Cells(row, next_action_col)
                ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry) = ObjWorkExcel.Cells(row, quest_intvw_date_col)
                ALL_PENDING_CASES_ARRAY(need_face_to_face, case_entry) = ObjWorkExcel.Cells(row, ftof_still_need_col)

                'ALL_PENDING_CASES_ARRAY(, case_entry) = ObjWorkExcel.Cells(row, )

                ALL_PENDING_CASES_ARRAY(need_appt_notc, case_entry) = TRUE
                ALL_PENDING_CASES_ARRAY(need_nomi, case_entry) = TRUE
                ALL_PENDING_CASES_ARRAY(take_action_today, case_entry) = FALSE


                case_entry = case_entry + 1
                row = row + 1
            End If
            Exit For
        End If
    Next



    If found_case_on_todays_list = FALSE Then
        'TODO figure out why case is not on the list any more and figure out what to do with it.
        'Remove from working sheet and add to list of cases removed
        'MsgBox "NOT ON TODAY'S LIST" & vbNewLine & ObjWorkExcel.Cells(row, case_nbr_col)
        ReDim Preserve CASES_NO_LONGER_WORKING(error_notes, case_removed)
        CASES_NO_LONGER_WORKING(worker_ID, case_removed) = ObjWorkExcel.Cells(row, worker_id_col)
        CASES_NO_LONGER_WORKING(case_number, case_removed) = ObjWorkExcel.Cells(row, case_nbr_col)
        CASES_NO_LONGER_WORKING(excel_row, case_removed) = row
        CASES_NO_LONGER_WORKING(client_name, case_removed) = ObjWorkExcel.Cells(row, case_name_col)
        CASES_NO_LONGER_WORKING(application_date, case_removed) = ObjWorkExcel.Cells(row, app_date_col)
        'CASES_NO_LONGER_WORKING(interview_date, case_removed) = ObjWorkExcel.Cells(row, intvw_date_col)
        CASES_NO_LONGER_WORKING(interview_date, case_removed) = ObjWorkExcel.Cells(row, intvw_date_col)
        CASES_NO_LONGER_WORKING(CASH_status, case_removed) = ObjWorkExcel.Cells(row, cash_stat_col)
        CASES_NO_LONGER_WORKING(SNAP_status, case_removed) = ObjWorkExcel.Cells(row, snap_stat_col)

        CASES_NO_LONGER_WORKING(appt_notc_sent, case_removed) = ObjWorkExcel.Cells(row, appt_notc_date_col)
        CASES_NO_LONGER_WORKING(appt_notc_confirm, case_removed) = ObjWorkExcel.Cells(row, appt_notc_confirm_col)
        CASES_NO_LONGER_WORKING(appointment_date, case_removed) = ObjWorkExcel.Cells(row, appt_date_col)
        CASES_NO_LONGER_WORKING(nomi_sent, case_removed) = ObjWorkExcel.Cells(row, nomi_date_col)
        CASES_NO_LONGER_WORKING(nomi_confirm, case_removed) = ObjWorkExcel.Cells(row, nomi_confirm_col)
        CASES_NO_LONGER_WORKING(next_action_needed, case_removed) = ObjWorkExcel.Cells(row, next_action_col)
        CASES_NO_LONGER_WORKING(questionable_intv, case_removed) = ObjWorkExcel.Cells(row, quest_intvw_date_col)

        CASES_NO_LONGER_WORKING(error_notes, case_removed) = ""
        'TODO figure out why case is not on the list any more add add to error notes
        'CASES_NO_LONGER_WORKING(error_notes, case_removed) = "Interview Completed on " & TODAYS_CASES_ARRAY(interview_date, case_entry)
        'MsgBox row
        case_removed = case_removed + 1
        SET objRange = ObjWorkExcel.Cells(row, 1).EntireRow
        objRange.Delete
    End If

    'row = row + 1
    next_case_number = trim(objWorkExcel.Cells(row, 1).Value)
Loop Until next_case_number = ""

add_a_case = case_entry
For case_entry = 0 to UBOUND(TODAYS_CASES_ARRAY, 2)
    'MsgBox TODAYS_CASES_ARRAY(on_working_list, case_entry)
    'MsgBox TODAYS_CASES_ARRAY(interview_date, case_entry)
    If TODAYS_CASES_ARRAY(on_working_list, case_entry) = FALSE AND TODAYS_CASES_ARRAY(interview_date, case_entry) = "" Then
        'MsgBox row
        ObjWorkExcel.Cells(row, worker_id_col) = TODAYS_CASES_ARRAY(worker_ID, case_entry)
        ObjWorkExcel.Cells(row, case_nbr_col) = TODAYS_CASES_ARRAY(case_number, case_entry)
        ObjWorkExcel.Cells(row, case_name_col) = TODAYS_CASES_ARRAY(client_name, case_entry)
        ObjWorkExcel.Cells(row, app_date_col) = TODAYS_CASES_ARRAY(application_date, case_entry)
        ObjWorkExcel.Cells(row, intvw_date_col) = TODAYS_CASES_ARRAY(interview_date, case_entry)

        'ObjWorkExcel.Cells(row, ) = TODAYS_CASES_ARRAY(, case_entry)


        ReDim Preserve ALL_PENDING_CASES_ARRAY(error_notes, add_a_case)
        ALL_PENDING_CASES_ARRAY(worker_ID, add_a_case) = TODAYS_CASES_ARRAY(worker_ID, case_entry)
        ALL_PENDING_CASES_ARRAY(case_number, add_a_case) = TODAYS_CASES_ARRAY(case_number, case_entry)
        ALL_PENDING_CASES_ARRAY(excel_row, add_a_case) = row
        ALL_PENDING_CASES_ARRAY(client_name, add_a_case) = TODAYS_CASES_ARRAY(client_name, case_entry)
        ALL_PENDING_CASES_ARRAY(application_date, add_a_case) = ObjWorkExcel.Cells(row, app_date_col)
        ALL_PENDING_CASES_ARRAY(interview_date, add_a_case) = ObjWorkExcel.Cells(row, intvw_date_col)
        ALL_PENDING_CASES_ARRAY(CASH_status, add_a_case) = ObjWorkExcel.Cells(row, cash_stat_col)
        ALL_PENDING_CASES_ARRAY(SNAP_status, add_a_case) = ObjWorkExcel.Cells(row, snap_stat_col)

        ALL_PENDING_CASES_ARRAY(appt_notc_sent, add_a_case) = ObjWorkExcel.Cells(row, appt_notc_date_col)
        ALL_PENDING_CASES_ARRAY(appt_notc_confirm, add_a_case) = ObjWorkExcel.Cells(row, appt_notc_confirm_col)
        ALL_PENDING_CASES_ARRAY(appointment_date, add_a_case) = ObjWorkExcel.Cells(row, appt_date_col)
        ALL_PENDING_CASES_ARRAY(nomi_sent, add_a_case) = ObjWorkExcel.Cells(row, nomi_date_col)
        ALL_PENDING_CASES_ARRAY(nomi_confirm, add_a_case) = ObjWorkExcel.Cells(row, nomi_confirm_col)
        ALL_PENDING_CASES_ARRAY(next_action_needed, add_a_case) = ObjWorkExcel.Cells(row, next_action_col)
        ALL_PENDING_CASES_ARRAY(questionable_intv, add_a_case) = ObjWorkExcel.Cells(row, quest_intvw_date_col)
        ALL_PENDING_CASES_ARRAY(need_face_to_face, add_a_case) = ObjWorkExcel.Cells(row, ftof_still_need_col)

        'ALL_PENDING_CASES_ARRAY(, add_a_case) = ObjWorkExcel.Cells(row, )

        ALL_PENDING_CASES_ARRAY(need_appt_notc, add_a_case) = TRUE
        ALL_PENDING_CASES_ARRAY(need_nomi, add_a_case) = TRUE
        ALL_PENDING_CASES_ARRAY(take_action_today, add_a_case) = FALSE

        add_a_case = add_a_case + 1
        row = row + 1
    End If
Next
MsgBox "Look at the list"
For case_entry = 0 to UBOUND(ALL_PENDING_CASES_ARRAY, 2)
    'Establishing values for each case in the array of cases
    MAXIS_case_number	= ALL_PENDING_CASES_ARRAY(case_number, case_entry)
    'MsgBox ALL_PENDING_CASES_ARRAY(case_number, case_entry)
    CALL navigate_to_MAXIS_screen("CASE", "CURR")
    'Checking for PRIV cases.
    EMReadScreen priv_check, 6, 24, 14 'If it can't get into the case needs to skip
    EMReadScreen county_check, 2, 21, 16    'Looking to see if case has Hennepin COunty worker
    EMReadScreen case_removed_in_MAXIS, 19, 24, 2
    If priv_check = "PRIVIL" THEN
        priv_case_list = priv_case_list & "|" & MAXIS_case_number
        ALL_PENDING_CASES_ARRAY(priv_case, case_entry) = TRUE
    ElseIf county_check <> "27" THEN
        ALL_PENDING_CASES_ARRAY(out_of_co, case_entry) = "OUT OF COUNTY - " & county_check
    ElseIf case_removed_in_MAXIS = "INVALID CASE NUMBER" Then
        ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "CASE HAS BEEN DELETED"
    Else
        ALL_PENDING_CASES_ARRAY(priv_case, case_entry) = FALSE

        IF ALL_PENDING_CASES_ARRAY(worker_ID, case_entry) = "X127EF8" or ALL_PENDING_CASES_ARRAY(worker_ID, case_entry) = "X127EJ1" THEN ALL_PENDING_CASES_ARRAY(error_notes, case_entry) = ALL_PENDING_CASES_ARRAY(error_notes, case_entry) & ", IMD CASE"
        'MEMB for written language

        'TODO - move the code for determineing 'take_action_today' to up here so it ONLY looks if action is needed today.
        'If no appointment notice sent date
        'if appointment date is tomorrow - look for an an interview or no pending progs
        'If NOMI is due
        'If at day 29 - loof for an interview or no pending progs
        'if at day 30 - or over
        'BETTER IDENTIFY IF INTERVIEW IS DONE'

        If ALL_PENDING_CASES_ARRAY(client_name, case_entry) = "XXXXX" Then
            Call navigate_to_MAXIS_screen("STAT", "MEMB")
            EMReadScreen last_name, 25, 6, 30
            EMReadScreen first_name, 12, 6, 63
            EMReadScreen middle_initial, 1, 6, 79

            last_name = replace(last_name, "_", "")
            first_name = replace(first_name, "_", "")
            middle_initial = replace(middle_initial, "_", "")

            ALL_PENDING_CASES_ARRAY(client_name, case_entry) = last_name & ", " & first_name & " " & middle_initial
        End If

        'PROG to determine programs active
        Call navigate_to_MAXIS_screen("STAT", "PROG")
        fs_intv = ""
        cash_intv_one = ""
        cash_intv_two = ""

        EMReadScreen cash_prog_one, 2, 6, 67               'reading for active MFIP program - which has different requirements
        EMReadScreen cash_stat_one, 4, 6, 74

        EMReadScreen cash_prog_two, 2, 7, 67
        EMReadScreen cash_stat_two, 4, 7, 74

        EMReadScreen fs_pend, 4, 10, 74

        cash_pend = FALSE
        cash_interview_done = FALSE
        snap_interview_done = FALSE

        If cash_stat_one = "PEND" Then
            cash_pend = TRUE
            EMReadScreen cash_intv_one, 8, 6, 55
            If cash_intv_one <> "__ __ __" Then
                cash_intv_one = replace(cash_intv_one, " ", "/")
                cash_interview_done = TRUE
            Else
                cash_intv_one = ""
            End If
        ElseIf cash_stat_one = "ACTV" Then
            ALL_PENDING_CASES_ARRAY(CASH_status, case_entry) = "Active"
        End If

        If cash_stat_two = "PEND" Then
            cash_pend = TRUE
            EMReadScreen cash_intv_two, 8, 7, 55
            If cash_intv_two <> "__ __ __" Then
                cash_intv_two = replace(cash_intv_two, " ", "/")
                cash_interview_done = TRUE
            Else
                cash_intv_two = ""
            End If
        ElseIf cash_stat_one = "ACTV" Then
            ALL_PENDING_CASES_ARRAY(CASH_status, case_entry) = "Active"
        Else
            ALL_PENDING_CASES_ARRAY(CASH_status, case_entry) = ""
        End If

        If cash_pend = TRUE then ALL_PENDING_CASES_ARRAY(CASH_status, case_entry) = "Pending"

        If fs_pend = "PEND" Then
            ALL_PENDING_CASES_ARRAY(SNAP_status, case_entry) = "Pending"
            EMReadScreen fs_intv, 8, 10, 55
            If fs_intv <> "__ __ __" Then
                fs_intv = replace(fs_intv, " ", "/")
                snap_interview_done = TRUE
            Else
                fs_intv = ""
            End If
        ElseIf fs_pend = "ACTV" Then
            ALL_PENDING_CASES_ARRAY(SNAP_status, case_entry) = "Active"
        Else
            ALL_PENDING_CASES_ARRAY(SNAP_status, case_entry) = ""
        End If

        If ALL_PENDING_CASES_ARRAY(SNAP_status, case_entry) <> "Pending" AND ALL_PENDING_CASES_ARRAY(CASH_status, case_entry) <> "Pending" Then
            ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "REMOVE FROM LIST"
            ALL_PENDING_CASES_ARRAY(error_notes, case_entry) = "Neither SNAP nor CASH is pending."
        Else
            If ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "REVIEW DENIAL" Then
                ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "*** DENY ***"
                ALL_PENDING_CASES_ARRAY(error_notes, case_entry) = "Denial Failed"
            End If
            If cash_pend = TRUE Then
                If cash_interview_done = TRUE Then
                    If cash_intv_one <> "" Then ALL_PENDING_CASES_ARRAY(interview_date, case_entry) = cash_intv_one
                    If cash_intv_two <> "" Then ALL_PENDING_CASES_ARRAY(interview_date, case_entry) = cash_intv_two
                    ALL_PENDING_CASES_ARRAY(need_face_to_face, case_entry) = ""
                    ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "NONE - Interview Completed"
                Else
                    If fs_pend = "PEND" Then
                        If fs_intv = "" THen
                            ALL_PENDING_CASES_ARRAY(interview_date, case_entry) = ""
                        Else
                            ALL_PENDING_CASES_ARRAY(interview_date, case_entry) = fs_intv
                            If ALL_PENDING_CASES_ARRAY(need_face_to_face, case_entry) = "" Then ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "CHECK FOR F2F NEEDED"
                            If ALL_PENDING_CASES_ARRAY(need_face_to_face, case_entry) = "N" Then ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "NONE - Interview Completed"
                            If ALL_PENDING_CASES_ARRAY(need_face_to_face, case_entry) = "Y" Then
                                ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "DENY AT DAY 30"
                                If ALL_PENDING_CASES_ARRAY(nomi_sent, case_entry) = "" Then ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "SEND NOMI"
                                IF ALL_PENDING_CASES_ARRAY(sppt_notc_sent, case_entry) = "" Then ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "SEND APPOINTMENT NOTICE"
                            End If
                            ALL_PENDING_CASES_ARRAY(error_notes, case_entry) = ", Cash interview incomplete."
                    'WHAT TO DO WITH F2F Cases'
                        End If
                    End If
                End If
            ElseIf fs_pend = "PEND" Then
                If fs_intv <> "" Then
                    ALL_PENDING_CASES_ARRAY(interview_date, case_entry) = fs_intv
                    ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "NONE - Interview Completed"
                    ALL_PENDING_CASES_ARRAY(need_face_to_face, case_entry) = ""
                End If
            End If
        End If

        If ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "" Then ALL_PENDING_CASES_ARRAY(take_action_today, case_entry) = TRUE
        If ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "SEND APPOINTMENT NOTICE" Then ALL_PENDING_CASES_ARRAY(take_action_today, case_entry) = TRUE
        If ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "SEND NOMI" AND DateDiff("d", date, ALL_PENDING_CASES_ARRAY(appointment_date, case_entry)) <= 1 Then ALL_PENDING_CASES_ARRAY(take_action_today, case_entry) = TRUE
        If ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "DENY AT DAY 30" AND DateDiff("d", ALL_PENDING_CASES_ARRAY(application_date, case_entry), date) >= 29 Then ALL_PENDING_CASES_ARRAY(take_action_today, case_entry) = TRUE

        If ALL_PENDING_CASES_ARRAY(take_action_today, case_entry) = TRUE Then

            ' COMMENTED OUT BECAUSE THIS DOESN"T CHANGE ANYTHING RIGHT NOW
            ' Call navigate_to_MAXIS_screen("STAT", "MEMB")
            ' EMReadScreen language_code, 2, 13, 42
            ' ALL_PENDING_CASES_ARRAY(written_lang, case_entry) = language_code

            Call navigate_to_MAXIS_screen("CASE", "NOTE")

            day_before_app = DateAdd("d", -1, ALL_PENDING_CASES_ARRAY(application_date, case_entry)) 'will set the date one day prior to app date'
            If ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "CHECK FOR F2F NEEDED" Then
                note_row = 5
                note_date = ""
                note_title = ""
                appt_date = ""
                Do
                    EMReadScreen note_date, 8, note_row, 6
                    EMReadScreen note_title, 55, note_row, 25
                    note_title = trim(note_title)

                    If left(note_title, 50) = "~ Application interview for cash is still needed ~" Then ALL_PENDING_CASES_ARRAY(need_face_to_face, case_entry) = "Y"
                    If left(note_title, 52) = "~ MFIP face to face application interview required ~" Then ALL_PENDING_CASES_ARRAY(need_face_to_face, case_entry) = "Y"
                    If note_title = "~ MFIP face to face application interview not required" Then ALL_PENDING_CASES_ARRAY(need_face_to_face, case_entry) = "N"
                    If left(note_title, 52) = "~ CASH face to face application interview required ~" Then ALL_PENDING_CASES_ARRAY(need_face_to_face, case_entry) = "Y"
                    'THIS IS THE CASE/NOTE FOR CLIENTS REQUIRING NO INTERVIEW'
                    ' ~ Application interview ~
                    ' * Client is in an IMD FACI and is not required to complete interview
                    ' ---
                    ' EWS Quality Improvement Team
                    IF note_date = "        " then Exit Do
                    note_row = note_row + 1
                    IF note_row = 19 THEN
                        PF8
                        note_row = 5
                    END IF
                    EMReadScreen next_note_date, 8, note_row, 6
                    IF next_note_date = "        " then Exit Do
                Loop until datevalue(next_note_date) < day_before_app 'looking ahead at the next case note kicking out the dates before app'
                If ALL_PENDING_CASES_ARRAY(need_face_to_face, case_entry) <> "" Then ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = ""
            End If
            go_to_top_of_notes

            If ALL_PENDING_CASES_ARRAY(appt_notc_sent, case_entry) = "" Then
                note_row = 5
                note_date = ""
                note_title = ""
                appt_date = ""
                Do
                    EMReadScreen note_date, 8, note_row, 6
                    EMReadScreen note_title, 55, note_row, 25
                    note_title = trim(note_title)

                    IF left(note_title, 35) = "~ Appointment letter sent in MEMO ~" then
                        ALL_PENDING_CASES_ARRAY(appt_notc_sent, case_entry) = note_date
    				ElseIF left(note_title, 42) = "~ Appointment letter sent in MEMO for SNAP" then
                        ALL_PENDING_CASES_ARRAY(appt_notc_sent, case_entry) = note_date
    				ElseIF left(note_title, 37) = "~ Appointment letter sent in MEMO for" then
    					EMReadScreen appt_date, 10, note_row, 63
    					appt_date = replace(appt_date, "~", "")
    				 	appt_date = trim(appt_date)
    					ALL_PENDING_CASES_ARRAY(appointment_date, case_entry) = appt_date
                        ALL_PENDING_CASES_ARRAY(appt_notc_sent, case_entry) = note_date
                        'MsgBox ALL_PENDING_CASES_ARRAY(appointment_date, case_entry)
    				END IF

                    IF note_date = "        " then Exit Do
                    note_row = note_row + 1
                    IF note_row = 19 THEN
                        PF8
                        note_row = 5
                    END IF
                    EMReadScreen next_note_date, 8, note_row, 6
                    IF next_note_date = "        " then Exit Do
                Loop until datevalue(next_note_date) < day_before_app 'looking ahead at the next case note kicking out the dates before app'
            End If
            go_to_top_of_notes

            If ALL_PENDING_CASES_ARRAY(nomi_sent, case_entry) = "" Then
                note_row = 5
                note_date = ""
                note_title = ""
                appt_date = ""
                Do
                    EMReadScreen note_date, 8, note_row, 6
                    EMReadScreen note_title, 55, note_row, 25
                    note_title = trim(note_title)

                    IF note_title = "~ Client missed application interview, NOMI sent via sc" then ALL_PENDING_CASES_ARRAY(nomi_sent, case_entry) = note_date
                    IF left(note_title, 32) = "**Client missed SNAP interview**" then ALL_PENDING_CASES_ARRAY(nomi_sent, case_entry) = note_date
    				IF left(note_title, 32) = "**Client missed CASH interview**" then ALL_PENDING_CASES_ARRAY(nomi_sent, case_entry) = note_date
    				IF left(note_title, 37) = "**Client missed SNAP/CASH interview**" then ALL_PENDING_CASES_ARRAY(nomi_sent, case_entry) = note_date
    				IF note_title = "~ Client has not completed application interview, NOMI" then ALL_PENDING_CASES_ARRAY(nomi_sent, case_entry) = note_date
                    IF note_title = "~ Client has not completed CASH APP interview, NOMI sen" then ALL_PENDING_CASES_ARRAY(nomi_sent, case_entry) = note_date

                    IF note_date = "        " then Exit Do
                    note_row = note_row + 1
                    IF note_row = 19 THEN
                        PF8
                        note_row = 5
                    END IF
                    EMReadScreen next_note_date, 8, note_row, 6
                    IF next_note_date = "        " then Exit Do
                Loop until datevalue(next_note_date) < day_before_app 'looking ahead at the next case note kicking out the dates before app'
            End If

            If ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "SEND APPOINTMENT NOTICE" AND ALL_PENDING_CASES_ARRAY(appt_notc_sent, case_entry) <> "" Then ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "SEND NOMI"
            If ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "SEND NOMI" AND ALL_PENDING_CASES_ARRAY(nomi_sent, case_entry) <> "" Then ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "DENY AT DAY 30"

            If ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "" Then
                ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "DENY AT DAY 30"
                If ALL_PENDING_CASES_ARRAY(nomi_sent, case_entry) = "" Then ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "SEND NOMI"
                If ALL_PENDING_CASES_ARRAY(appt_notc_sent, case_entry) = "" THen ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "SEND APPOINTMENT NOTICE"
            End If

            If ALL_PENDING_CASES_ARRAY(appt_notc_sent, case_entry) <> "" AND ALL_PENDING_CASES_ARRAY(appointment_date, case_entry) = "" Then
                Call navigate_to_MAXIS_screen ("SPEC", "MEMO")
                memo_mo = DatePart("m", ALL_PENDING_CASES_ARRAY(appt_notc_sent, case_entry))
                memo_mo = right("00"&memo_mo, 2)
                memo_yr = DatePart("yyyy", ALL_PENDING_CASES_ARRAY(appt_notc_sent, case_entry))
                memo_yr = right(memo_yr, 2)

                EmWriteScreen memo_mo, 3, 48
                EmWriteScreen memo_yr, 3, 53
                transmit

                'creating a variable in the MM/DD/YY format to compare with date read from MAXIS
                look_date = ALL_PENDING_CASES_ARRAY(appt_notc_sent, case_entry)
                'MsgBox MAXIS_case_number & " - Date 1"
                CAll convert_to_mainframe_date(look_date, 2)

                Do
                    EMReadScreen create_date, 8, memo_row, 19                 'Reading the date of each memo and the status
                    EMReadScreen print_status, 7, memo_row, 67
                    'MsgBox print_status
                    If create_date = look_date AND print_status = "Printed" Then   'MEMOs created today and still waiting is likely our MEMO.
                        EmWriteScreen "X", memo_row, 16
                        transmit
                        PF8

                        EMReadScreen start_of_msg, 35, 15,12
                        If start_of_msg = "You recently applied for assistance" Then
                            EMReadScreen date_in_memo, 10, 19, 47
                            date_in_memo = trim(date_in_memo)
                            ALL_PENDING_CASES_ARRAY(appointment_date, case_entry) = replace(date_in_memo, ".", "")
                            'MsgBox ALL_PENDING_CASES_ARRAY(notc_confirm, case_entry)                 'For statistical purposes
                            Pf3
                            Exit Do
                        End If
                        PF3
                    End If
                    memo_row = memo_row + 1           'Looking at next row'
                Loop Until create_date = "        "
            End If
            'NOW CASES SHOULD HAVE CORRECT NEXT ACTION NEEDED
        End If

        ALL_PENDING_CASES_ARRAY(take_action_today, case_entry) = FALSE
        If ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "SEND NOMI" AND ALL_PENDING_CASES_ARRAY(appointment_date, case_entry) = "" Then MsgBox "This case needs a NOMI but script cannot find an appointment date."

        If ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "" Then ALL_PENDING_CASES_ARRAY(take_action_today, case_entry) = TRUE
        If ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "SEND APPOINTMENT NOTICE" Then ALL_PENDING_CASES_ARRAY(take_action_today, case_entry) = TRUE
        If ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "SEND NOMI" AND DateDiff("d", date, ALL_PENDING_CASES_ARRAY(appointment_date, case_entry)) <= 1 Then ALL_PENDING_CASES_ARRAY(take_action_today, case_entry) = TRUE
        If ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "DENY AT DAY 30" AND DateDiff("d", ALL_PENDING_CASES_ARRAY(application_date, case_entry), date) >= 29 Then ALL_PENDING_CASES_ARRAY(take_action_today, case_entry) = TRUE

        If ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "" Then MsgBox "Case Number: " & ALL_PENDING_CASES_ARRAY(case_number, case_entry) & vbNewLine & "Does not have an action to take!!!"

        If ALL_PENDING_CASES_ARRAY(take_action_today, case_entry) = TRUE and ALL_PENDING_CASES_ARRAY(interview_date, case_entry) = "" Then
            Call navigate_to_MAXIS_screen("CASE", "NOTE")
            note_row = 5
            start_dates = ""
            day_before_app = DateAdd("d", -1, ALL_PENDING_CASES_ARRAY(application_date, case_entry)) 'will set the date one day prior to app date'
            If InStr(ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry), "~") <> 0 Then start_dates = ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry)
            If ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry) <> "" Then
                Call convert_to_mainframe_date(ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry), 2)
                start_dates = ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry)
            End If
            Do
                EMReadScreen note_date, 8, note_row, 6
                EMReadScreen note_title, 55, note_row, 25
                note_title = trim(note_title)
                check_this_date = TRUE

                array_of_dates = ""
                If InStr(ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry), "~") <> 0 Then
                    array_of_dates = split(ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry), "~")
                    If array_of_dates(0) <> "" Then
                        For each dates in array_of_dates
                            'MsgBox MAXIS_case_number & " - Date 2"
                            Call convert_to_mainframe_date(dates, 2)
                            'MsgBox "Already known questionable date: " & dates & vbNewLine & "Note Date: " & note_date
                            if DateValue(dates) = DateValue(note_date) Then check_this_date = FALSE
                        Next
                    End If
                Else
                    'MsgBox "Already known questionable date: " & dates & vbNewLine & "Note Date: " & note_date
                    If ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry) <> "" Then
                        Call convert_to_mainframe_date(ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry), 2)
                        'MsgBox "Already known questionable date: " & ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry) & vbNewLine & "Note Date: " & note_date
                        if DateValue(ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry)) = DateValue(note_date) Then check_this_date = FALSE
                    End If
                End If

                If check_this_date = TRUE Then 'if a questionable interview date is left on the spreadsheet - that means it has been reviewed and is NOT an interview.
                    IF left(note_title, 15) = "***Add program:" then
                        ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry) = ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry) & "~" & note_date
                        ALL_PENDING_CASES_ARRAY(error_notes, case_entry) = ALL_PENDING_CASES_ARRAY(error_notes, case_entry) & ", Case Note suggests interview completed but interview not listed on PROG."
                    END IF
                    IF left(note_title, 33) = "***Intake Interview Completed ***" then
                        ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry) = ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry) & "~" & note_date
                        ALL_PENDING_CASES_ARRAY(error_notes, case_entry) = ALL_PENDING_CASES_ARRAY(error_notes, case_entry) & ", Case Note suggests interview completed but interview not listed on PROG."
                    END IF
                    IF left(note_title, 40) = "***Reapplication Interview Completed ***" then
                        ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry) = ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry) & "~" & note_date
                        ALL_PENDING_CASES_ARRAY(error_notes, case_entry) = ALL_PENDING_CASES_ARRAY(error_notes, case_entry) & ", Case Note suggests interview completed but interview not listed on PROG."
                    END IF
                    IF left(note_title, 42) = "~ Interview Completed for SNAP ~" then
                        ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry) = ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry) & "~" & note_date
                        ALL_PENDING_CASES_ARRAY(error_notes, case_entry) = ALL_PENDING_CASES_ARRAY(error_notes, case_entry) & ", Case Note suggests interview completed but interview not listed on PROG."
                    END IF
                    IF left(note_title, 42) = "*client interviewed* onboarding processing" then
                        ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry) = ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry) & "~" & note_date
                        ALL_PENDING_CASES_ARRAY(error_notes, case_entry) = ALL_PENDING_CASES_ARRAY(error_notes, case_entry) & ", Case Note suggests interview completed but interview not listed on PROG."
                    END IF
                    IF left(note_title, 34) = "***Intake: pending mentor review**" then
                        ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry) = ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry) & "~" & note_date
                        ALL_PENDING_CASES_ARRAY(error_notes, case_entry) = ALL_PENDING_CASES_ARRAY(error_notes, case_entry) & ", Case Note suggests interview completed but interview not listed on PROG."
                    END IF
                    IF left(note_title, 23) = "~ Interview Completed ~" then
                        ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry) = ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry) & "~" & note_date
                        ALL_PENDING_CASES_ARRAY(error_notes, case_entry) = ALL_PENDING_CASES_ARRAY(error_notes, case_entry) & ", Case Note suggests interview completed but interview not listed on PROG."
                    END IF
                    IF left(note_title, 10) = "***Intake:" then
                        ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry) = ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry) & "~" & note_date
                        ALL_PENDING_CASES_ARRAY(error_notes, case_entry) = ALL_PENDING_CASES_ARRAY(error_notes, case_entry) & ", Case Note suggests interview completed but interview not listed on PROG."
                    END IF
                    IF left(note_title, 24) = "~ Application interview ~" then
                        ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry) = ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry) & "~" & note_date
                        ALL_PENDING_CASES_ARRAY(error_notes, case_entry) = ALL_PENDING_CASES_ARRAY(error_notes, case_entry) & ", This case may not require an interview."
                    END IF
                    IF left(note_title, 33) = "***Intake Interview Completed ***" then
                        ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry) = ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry) & "~" & note_date
                        ALL_PENDING_CASES_ARRAY(error_notes, case_entry) = ALL_PENDING_CASES_ARRAY(error_notes, case_entry) & ", Case Note suggests interview completed but interview not listed on PROG."
                    END IF
                    IF left(UCase(note_title), 51) = "Phone call from client re: Phone interview Complete" then
                        ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry) = ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry) & "~" & note_date
                        ALL_PENDING_CASES_ARRAY(error_notes, case_entry) = ALL_PENDING_CASES_ARRAY(error_notes, case_entry) & ", Possible case note indicating an interview. If an interview happened, PROG was not updated and an incorrect script was used."
                    END IF
                    IF left(UCase(note_title), 41) = "Phone call from client re: SNAP interview" then
                        ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry) = ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry) & "~" & note_date
                        ALL_PENDING_CASES_ARRAY(error_notes, case_entry) = ALL_PENDING_CASES_ARRAY(error_notes, case_entry) & ", Possible case note indicating an interview. If an interview happened, PROG was not updated and an incorrect script was used."
                    END IF
                End If
                IF note_date = "        " then Exit Do
                note_row = note_row + 1
                IF note_row = 19 THEN
                    PF8
                    note_row = 5
                END IF
                EMReadScreen next_note_date, 8, note_row, 6
                IF next_note_date = "        " then Exit Do
            Loop until datevalue(next_note_date) < day_before_app 'looking ahead at the next case note kicking out the dates before app'

            If left(ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry), 1) = "~" Then ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry) = right(ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry), len(ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry))-1)
            if ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry) <> start_dates Then ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "REVIEW QUESTIONABLE INTERVIEW DATE(S)"

        End If

        ALL_PENDING_CASES_ARRAY(take_action_today, case_entry) = FALSE

        If ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "SEND APPOINTMENT NOTICE" Then ALL_PENDING_CASES_ARRAY(take_action_today, case_entry) = TRUE
        If ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "SEND NOMI" AND DateDiff("d", date, ALL_PENDING_CASES_ARRAY(appointment_date, case_entry)) <= 0 Then ALL_PENDING_CASES_ARRAY(take_action_today, case_entry) = TRUE
        If ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "DENY AT DAY 30" AND DateDiff("d", ALL_PENDING_CASES_ARRAY(application_date, case_entry), date) >= 30 Then ALL_PENDING_CASES_ARRAY(take_action_today, case_entry) = TRUE

        If ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "" Then MsgBox "Case Number: " & ALL_PENDING_CASES_ARRAY(case_number, case_entry) & vbNewLine & "Does not have an action to take!!!"
    End If
Next
back_to_SELF

Dim ACTION_TODAY_CASES_ARRAY()
ReDim ACTION_TODAY_CASES_ARRAY(error_notes, 0)

todays_cases = 0

For case_entry = 0 to UBOUND(ALL_PENDING_CASES_ARRAY, 2)
    MAXIS_case_number	= ALL_PENDING_CASES_ARRAY(case_number, case_entry)

    forms_to_swkr = ""
    forms_to_arep = ""
    memo_started = TRUE

    'MsgBox MAXIS_case_number & vbNewLine & "Take action: " & ALL_PENDING_CASES_ARRAY(take_action_today, case_entry) & vbNewLine & "Next action: " & ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry)
    If ALL_PENDING_CASES_ARRAY(take_action_today, case_entry) = TRUE Then
        if ALL_PENDING_CASES_ARRAY(CASH_status, case_entry) = "Pending" then           'setting the language for the notices - MFIP or SNAP
            if ALL_PENDING_CASES_ARRAY(SNAP_status, case_entry) = "Pending" then
                programs = "CASH/SNAP"
            else
                programs = "CASH"
            end if
        else
            programs = "SNAP"
        end if

        If ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "SEND APPOINTMENT NOTICE" Then
            'Call Navigate_to_MAXIS_screen("CASE", "NOTE")
            'MsgBox "We're Sending an Appointment Notice."

            'THIS IS FOR TESTING'
            need_intv_date = dateadd("d", 7, ALL_PENDING_CASES_ARRAY(application_date, case_entry))    'NOTE - had to change this - it did not call the full array - dates were wrong.
            If need_intv_date <= date then need_intv_date = dateadd("d", 7, date)

            ALL_PENDING_CASES_ARRAY(appointment_date, case_entry) = need_intv_date

            ALL_PENDING_CASES_ARRAY(appt_notc_sent, case_entry) = date
            ALL_PENDING_CASES_ARRAY(appt_notc_confirm, case_entry) = "Y"
            ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "SEND NOMI"

            ReDim Preserve ACTION_TODAY_CASES_ARRAY(error_notes, todays_cases)
            ACTION_TODAY_CASES_ARRAY(case_number, todays_cases)         = ALL_PENDING_CASES_ARRAY(case_number, case_entry)
            ACTION_TODAY_CASES_ARRAY(client_name, todays_cases)         = ALL_PENDING_CASES_ARRAY(client_name, case_entry)
            ACTION_TODAY_CASES_ARRAY(worker_ID, todays_cases)           = ALL_PENDING_CASES_ARRAY(worker_ID, case_entry)
            ACTION_TODAY_CASES_ARRAY(SNAP_status, todays_cases)         = ALL_PENDING_CASES_ARRAY(SNAP_status, case_entry)
            ACTION_TODAY_CASES_ARRAY(CASH_status, todays_cases)         = ALL_PENDING_CASES_ARRAY(CASH_status, case_entry)
            ACTION_TODAY_CASES_ARRAY(application_date, todays_cases)    = ALL_PENDING_CASES_ARRAY(application_date, case_entry)
            ACTION_TODAY_CASES_ARRAY(interview_date, todays_cases)      = ALL_PENDING_CASES_ARRAY(interview_date, case_entry)
            ACTION_TODAY_CASES_ARRAY(questionable_intv, todays_cases)   = ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry)
            ACTION_TODAY_CASES_ARRAY(need_face_to_face, todays_cases)   = ALL_PENDING_CASES_ARRAY(need_face_to_face, case_entry)
            ACTION_TODAY_CASES_ARRAY(appt_notc_sent, todays_cases)      = ALL_PENDING_CASES_ARRAY(appt_notc_sent, case_entry)
            ACTION_TODAY_CASES_ARRAY(appt_notc_confirm, todays_cases)   = ALL_PENDING_CASES_ARRAY(appt_notc_confirm, case_entry)
            ACTION_TODAY_CASES_ARRAY(appointment_date, todays_cases)    = ALL_PENDING_CASES_ARRAY(appointment_date, case_entry)
            ACTION_TODAY_CASES_ARRAY(nomi_sent, todays_cases)           = ALL_PENDING_CASES_ARRAY(nomi_sent, case_entry)
            ACTION_TODAY_CASES_ARRAY(nomi_confirm, todays_cases)        = ALL_PENDING_CASES_ARRAY(nomi_confirm, case_entry)
            ACTION_TODAY_CASES_ARRAY(deny_day30, todays_cases)          = ALL_PENDING_CASES_ARRAY(deny_day30, case_entry)
            ACTION_TODAY_CASES_ARRAY(deny_memo_confirm, todays_cases)   = ALL_PENDING_CASES_ARRAY(deny_memo_confirm, case_entry)
            ACTION_TODAY_CASES_ARRAY(next_action_needed, todays_cases)  = ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry)
            ACTION_TODAY_CASES_ARRAY(error_notes, todays_cases)         = ALL_PENDING_CASES_ARRAY(error_notes, case_entry) & " - " & "Appointment Notice Sent today"
            todays_cases = todays_cases + 1

            ' 'THIS IS FOR REAL LIFE'
            ' need_intv_date = dateadd("d", 7, ALL_PENDING_CASES_ARRAY(application_date, case_entry))    'NOTE - had to change this - it did not call the full array - dates were wrong.
            ' If need_intv_date <= date then need_intv_date = dateadd("d", 7, date)
            '
            ' ALL_PENDING_CASES_ARRAY(appointment_date, case_entry) = need_intv_date
            ' need_intv_date = need_intv_date & ""		'turns interview date into string for variable
            '
            ' start_a_new_spec_memo_and_continue(memo_started)		'Writes the appt letter into the MEMO.
			' IF memo_started = True THEN
            '     EMsendkey("************************************************************")
            '     Call write_variable_in_SPEC_MEMO("You recently applied for assistance in Hennepin County on " & ALL_PENDING_CASES_ARRAY(application_date, case_entry) & ".")
            '     Call write_variable_in_SPEC_MEMO("An interview is required to process your application.")
            '     Call write_variable_in_SPEC_MEMO(" ")
            '     Call write_variable_in_SPEC_MEMO("The interview must be completed by " & need_intv_date & ".")
            '     Call write_variable_in_SPEC_MEMO("To complete a phone interview, call the EZ Info Line at")
            '     Call write_variable_in_SPEC_MEMO("612-596-1300 between 9:00am and 4:00pm Monday through Friday.")
            '     Call write_variable_in_SPEC_MEMO(" ")
            '     Call write_variable_in_SPEC_MEMO("If we do not hear from you by " & last_contact_day & " your application will be denied.") 'add 30 days
            '     Call write_variable_in_SPEC_MEMO("If you are applying for a cash program for pregnant women or minor children, you may need a face-to-face interview.")
            '     Call write_variable_in_SPEC_MEMO("Domestic violence brochures are available at https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3477-ENG.")
            '     Call write_variable_in_SPEC_MEMO("You can also request a paper copy.")
            '     Call write_variable_in_SPEC_MEMO("Auth: 7CFR 273.2(e)(3). ")
            '     Call write_variable_in_SPEC_MEMO("************************************************************")
            '     ALL_PENDING_CASES_ARRAY(appt_notc_sent, case_entry) = date
            '     PF4
			' ELSE
			' 	ALL_PENDING_CASES_ARRAY(appt_notc_confirm, case_entry) = "N" 'Setting this as N if the MEMO failed
			' END IF
            '
            ' If ALL_PENDING_CASES_ARRAY(appt_notc_confirm, case_entry) <> "N" Then
            '     Call confirm_memo_waiting(ALL_PENDING_CASES_ARRAY(appt_notc_confirm, case_entry))
            ' End If
            '
            ' If ALL_PENDING_CASES_ARRAY(appt_notc_confirm, case_entry) = "N" Then
            '     ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "Send Manual Appt Notice"
            ' ElseIf ALL_PENDING_CASES_ARRAY(appt_notc_confirm, case_entry) = "Y"
            '     ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "SEND NOMI"
            '     Call start_a_blank_case_note
            '
            '     Call write_variable_in_CASE_NOTE("~ Appointment letter sent in MEMO for " & need_intv_date & "~")
            '     Call write_variable_in_CASE_NOTE("* A notice has been sent via SPEC/MEMO informing the client of needed interview.")
            '     Call write_variable_in_CASE_NOTE("* Households failing to complete the interview within 30 days of the date they file an application will receive a denial notice")
            '     Call write_variable_in_CASE_NOTE("* A link to the domestic violence brochure sent to client in SPEC/MEMO as a part of interview notice.")
            '     Call write_variable_in_CASE_NOTE("---")
            '     Call write_variable_in_CASE_NOTE(worker_signature & " via bulk on demand waiver script")
            '     'MsgBox "What casenote was sent?"
            '     PF3
            '
            ' Else
            '     ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "???"
            ' End If
            ' Call back_to_SELF

            ' ReDim Preserve ACTION_TODAY_CASES_ARRAY(error_notes, todays_cases)
            ' ACTION_TODAY_CASES_ARRAY(case_number, todays_cases)         = ALL_PENDING_CASES_ARRAY(case_number, case_entry)
            ' ACTION_TODAY_CASES_ARRAY(client_name, todays_cases)         = ALL_PENDING_CASES_ARRAY(client_name, case_entry)
            ' ACTION_TODAY_CASES_ARRAY(worker_ID, todays_cases)           = ALL_PENDING_CASES_ARRAY(worker_ID, case_entry)
            ' ACTION_TODAY_CASES_ARRAY(SNAP_status, todays_cases)         = ALL_PENDING_CASES_ARRAY(SNAP_status, case_entry)
            ' ACTION_TODAY_CASES_ARRAY(CASH_status, todays_cases)         = ALL_PENDING_CASES_ARRAY(CASH_status, case_entry)
            ' ACTION_TODAY_CASES_ARRAY(application_date, todays_cases)    = ALL_PENDING_CASES_ARRAY(application_date, case_entry)
            ' ACTION_TODAY_CASES_ARRAY(interview_date, todays_cases)      = ALL_PENDING_CASES_ARRAY(interview_date, case_entry)
            ' ACTION_TODAY_CASES_ARRAY(questionable_intv, todays_cases)   = ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry)
            ' ACTION_TODAY_CASES_ARRAY(need_face_to_face, todays_cases)   = ALL_PENDING_CASES_ARRAY(need_face_to_face, case_entry)
            ' ACTION_TODAY_CASES_ARRAY(appt_notc_sent, todays_cases)      = ALL_PENDING_CASES_ARRAY(appt_notc_sent, case_entry)
            ' ACTION_TODAY_CASES_ARRAY(appt_notc_confirm, todays_cases)   = ALL_PENDING_CASES_ARRAY(appt_notc_confirm, case_entry)
            ' ACTION_TODAY_CASES_ARRAY(appointment_date, todays_cases)    = ALL_PENDING_CASES_ARRAY(appointment_date, case_entry)
            ' ACTION_TODAY_CASES_ARRAY(nomi_sent, todays_cases)           = ALL_PENDING_CASES_ARRAY(nomi_sent, case_entry)
            ' ACTION_TODAY_CASES_ARRAY(nomi_confirm, todays_cases)        = ALL_PENDING_CASES_ARRAY(nomi_confirm, case_entry)
            ' ACTION_TODAY_CASES_ARRAY(deny_day30, todays_cases)          = ALL_PENDING_CASES_ARRAY(deny_day30, case_entry)
            ' ACTION_TODAY_CASES_ARRAY(deny_memo_confirm, todays_cases)   = ALL_PENDING_CASES_ARRAY(deny_memo_confirm, case_entry)
            ' ACTION_TODAY_CASES_ARRAY(next_action_needed, todays_cases)  = ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry)
            ' ACTION_TODAY_CASES_ARRAY(error_notes, todays_cases)         = ALL_PENDING_CASES_ARRAY(error_notes, case_entry) & " - " & "Appointment Notice Sent today"
            ' todays_cases = todays_cases + 1


        ElseIf ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "SEND NOMI" Then
            'THIS IS FOR TESTING'
            ALL_PENDING_CASES_ARRAY(nomi_sent, case_entry) = date
            ALL_PENDING_CASES_ARRAY(nomi_confirm, case_entry) = "Y"
            ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "DENY AT DAY 30"
            ReDim Preserve ACTION_TODAY_CASES_ARRAY(error_notes, todays_cases)
            ACTION_TODAY_CASES_ARRAY(case_number, todays_cases)         = ALL_PENDING_CASES_ARRAY(case_number, case_entry)
            ACTION_TODAY_CASES_ARRAY(client_name, todays_cases)         = ALL_PENDING_CASES_ARRAY(client_name, case_entry)
            ACTION_TODAY_CASES_ARRAY(worker_ID, todays_cases)           = ALL_PENDING_CASES_ARRAY(worker_ID, case_entry)
            ACTION_TODAY_CASES_ARRAY(SNAP_status, todays_cases)         = ALL_PENDING_CASES_ARRAY(SNAP_status, case_entry)
            ACTION_TODAY_CASES_ARRAY(CASH_status, todays_cases)         = ALL_PENDING_CASES_ARRAY(CASH_status, case_entry)
            ACTION_TODAY_CASES_ARRAY(application_date, todays_cases)    = ALL_PENDING_CASES_ARRAY(application_date, case_entry)
            ACTION_TODAY_CASES_ARRAY(interview_date, todays_cases)      = ALL_PENDING_CASES_ARRAY(interview_date, case_entry)
            ACTION_TODAY_CASES_ARRAY(questionable_intv, todays_cases)   = ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry)
            ACTION_TODAY_CASES_ARRAY(need_face_to_face, todays_cases)   = ALL_PENDING_CASES_ARRAY(need_face_to_face, case_entry)
            ACTION_TODAY_CASES_ARRAY(appt_notc_sent, todays_cases)      = ALL_PENDING_CASES_ARRAY(appt_notc_sent, case_entry)
            ACTION_TODAY_CASES_ARRAY(appt_notc_confirm, todays_cases)   = ALL_PENDING_CASES_ARRAY(appt_notc_confirm, case_entry)
            ACTION_TODAY_CASES_ARRAY(appointment_date, todays_cases)    = ALL_PENDING_CASES_ARRAY(appointment_date, case_entry)
            ACTION_TODAY_CASES_ARRAY(nomi_sent, todays_cases)           = ALL_PENDING_CASES_ARRAY(nomi_sent, case_entry)
            ACTION_TODAY_CASES_ARRAY(nomi_confirm, todays_cases)        = ALL_PENDING_CASES_ARRAY(nomi_confirm, case_entry)
            ACTION_TODAY_CASES_ARRAY(deny_day30, todays_cases)          = ALL_PENDING_CASES_ARRAY(deny_day30, case_entry)
            ACTION_TODAY_CASES_ARRAY(deny_memo_confirm, todays_cases)   = ALL_PENDING_CASES_ARRAY(deny_memo_confirm, case_entry)
            ACTION_TODAY_CASES_ARRAY(next_action_needed, todays_cases)  = ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry)
            ACTION_TODAY_CASES_ARRAY(error_notes, todays_cases)         = ALL_PENDING_CASES_ARRAY(error_notes, case_entry) & " - " & "NOMI Sent today"
            todays_cases = todays_cases + 1
            '
            ' 'THIS IS FOR REAL LIFE'
            ' nomi_last_contact_day = dateadd("d", 30, ALL_PENDING_CASES_ARRAY(application_date, case_entry))
            ' 'ensuring that we have given the client an additional10days fromt he day nomi sent'
            ' IF DateDiff("d", ALL_PENDING_CASES_ARRAY(nomi_sent, case_entry), nomi_last_contact_day) < 1 then nomi_last_contact_day = dateadd("d", 10, ALL_PENDING_CASES_ARRAY(nomi_sent, case_entry))
            '
            ' start_a_new_spec_memo_and_continue(memo_started)		'Writes the NOMI into the MEMO.
            ' IF memo_started = TRUE THEN
            '     EMsendkey("************************************************************")
            '     Call write_variable_in_SPEC_MEMO("You recently applied for assistance on " & ALL_PENDING_CASES_ARRAY(application_date, case_entry) & ".")
            '     Call write_variable_in_SPEC_MEMO("Your interview should have been completed by " & ALL_PENDING_CASES_ARRAY(appointment_date, case_entry) & ".")
            '     Call write_variable_in_SPEC_MEMO("An interview is required to process your application.")
            '     Call write_variable_in_SPEC_MEMO("To complete a phone interview, call the EZ Info Line at ")
            '     Call write_variable_in_SPEC_MEMO("612-596-1300 between 9:00am and 4:00pm Monday through Friday.")
            '     Call write_variable_in_SPEC_MEMO(" ")
            '     Call write_variable_in_SPEC_MEMO("If you do not complete the interview by " & nomi_last_contact_day & " your application will be denied.") 'add 30 days
            '     Call write_variable_in_SPEC_MEMO(" ")
            '     Call write_variable_in_SPEC_MEMO("If you are applying for a cash program for pregnant women or minor children, you may need a face-to- face interview.")
            '     Call write_variable_in_SPEC_MEMO("Domestic violence brochures are available at https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3477-ENG.")
            '     Call write_variable_in_SPEC_MEMO("You can also request a paper copy.")
            '     Call write_variable_in_SPEC_MEMO("Auth: 7CFR 273.2(e)(3). ")
            '     Call write_variable_in_SPEC_MEMO("************************************************************")
            '     ALL_PENDING_CASES_ARRAY(nomi_sent, case_entry) = date
            '     PF4
            ' Else
            '     ALL_PENDING_CASES_ARRAY(nomi_confirm, case_entry) = "N"
            ' End If
            '
            ' If ALL_PENDING_CASES_ARRAY(nomi_confirm, case_entry) <> "N" Then
            '     Call confirm_memo_waiting(ALL_PENDING_CASES_ARRAY(nomi_confirm, case_entry))
            ' End If
            '
            ' If ALL_PENDING_CASES_ARRAY(nomi_confirm, case_entry) = "N" Then
            '     ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "Send Manual NOMI"
            ' ElseIf ALL_PENDING_CASES_ARRAY(nomi_confirm, case_entry) = "Y"
            '     ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "DENY AT DAY 30"
            '     Call start_a_blank_case_note
            '     Call write_variable_in_CASE_NOTE("~ Client has not completed application interview, NOMI sent via script ~ ")
            '     Call write_variable_in_CASE_NOTE("* A notice was previously sent to client with detail about completing an interview. ")
            '     Call write_variable_in_CASE_NOTE("* Households failing to complete the interview within 30 days of the date they file an application will receive a denial notice")
            '     Call write_variable_in_CASE_NOTE("* A link to the domestic violence brochure sent to client in SPEC/MEMO as a part of interview notice.")
            '     Call write_variable_in_CASE_NOTE("---")
            '     Call write_variable_in_CASE_NOTE(worker_signature & " via bulk on demand waiver script")
            '     'MsgBox "What casenote was sent?"
            '     PF3
            ' Else
            '     ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "???"
            ' End If
            ' Call back_to_SELF

            ' ReDim Preserve ACTION_TODAY_CASES_ARRAY(error_notes, todays_cases)
            ' ACTION_TODAY_CASES_ARRAY(case_number, todays_cases)         = ALL_PENDING_CASES_ARRAY(case_number, case_entry)
            ' ACTION_TODAY_CASES_ARRAY(client_name, todays_cases)         = ALL_PENDING_CASES_ARRAY(client_name, case_entry)
            ' ACTION_TODAY_CASES_ARRAY(worker_ID, todays_cases)           = ALL_PENDING_CASES_ARRAY(worker_ID, case_entry)
            ' ACTION_TODAY_CASES_ARRAY(SNAP_status, todays_cases)         = ALL_PENDING_CASES_ARRAY(SNAP_status, case_entry)
            ' ACTION_TODAY_CASES_ARRAY(CASH_status, todays_cases)         = ALL_PENDING_CASES_ARRAY(CASH_status, case_entry)
            ' ACTION_TODAY_CASES_ARRAY(application_date, todays_cases)    = ALL_PENDING_CASES_ARRAY(application_date, case_entry)
            ' ACTION_TODAY_CASES_ARRAY(interview_date, todays_cases)      = ALL_PENDING_CASES_ARRAY(interview_date, case_entry)
            ' ACTION_TODAY_CASES_ARRAY(questionable_intv, todays_cases)   = ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry)
            ' ACTION_TODAY_CASES_ARRAY(need_face_to_face, todays_cases)   = ALL_PENDING_CASES_ARRAY(need_face_to_face, case_entry)
            ' ACTION_TODAY_CASES_ARRAY(appt_notc_sent, todays_cases)      = ALL_PENDING_CASES_ARRAY(appt_notc_sent, case_entry)
            ' ACTION_TODAY_CASES_ARRAY(appt_notc_confirm, todays_cases)   = ALL_PENDING_CASES_ARRAY(appt_notc_confirm, case_entry)
            ' ACTION_TODAY_CASES_ARRAY(appointment_date, todays_cases)    = ALL_PENDING_CASES_ARRAY(appointment_date, case_entry)
            ' ACTION_TODAY_CASES_ARRAY(nomi_sent, todays_cases)           = ALL_PENDING_CASES_ARRAY(nomi_sent, case_entry)
            ' ACTION_TODAY_CASES_ARRAY(nomi_confirm, todays_cases)        = ALL_PENDING_CASES_ARRAY(nomi_confirm, case_entry)
            ' ACTION_TODAY_CASES_ARRAY(deny_day30, todays_cases)          = ALL_PENDING_CASES_ARRAY(deny_day30, case_entry)
            ' ACTION_TODAY_CASES_ARRAY(deny_memo_confirm, todays_cases)   = ALL_PENDING_CASES_ARRAY(deny_memo_confirm, case_entry)
            ' ACTION_TODAY_CASES_ARRAY(next_action_needed, todays_cases)  = ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry)
            ' ACTION_TODAY_CASES_ARRAY(error_notes, todays_cases)         = ALL_PENDING_CASES_ARRAY(error_notes, case_entry) & " - " & "NOMI Sent today"
            ' todays_cases = todays_cases + 1

        ElseIf ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "DENY AT DAY 30" Then
            IF datediff("d", ALL_PENDING_CASES_ARRAY(application_date, case_entry), date) >= 30 and ALL_PENDING_CASES_ARRAY(interview_date, case_entry) = "" THEN
                'MsgBox "Both false notice"
                'MsgBox ALL_PENDING_CASES_ARRAY(nomi_sent, case_entry)
                IF ALL_PENDING_CASES_ARRAY(nomi_sent, case_entry) <> "" then
                    day_30 = dateadd("d", 30, ALL_PENDING_CASES_ARRAY(application_date, case_entry))
                    IF datediff("d", ALL_PENDING_CASES_ARRAY(nomi_sent, case_entry), date) >= 10 or datediff("d", ALL_PENDING_CASES_ARRAY(nomi_sent, case_entry), day_30) > 0 THEN
                    'MsgBox datediff("d", ALL_PENDING_CASES_ARRAY(nomi_sent, case_entry), date)
                        'Add handling to read PND2 for MSA (36005) and delay denial for 60 days.'
                        Call navigate_to_MAXIS_screen("REPT", "PND2")
                        Row = 1
                        Col = 1
                        EMSearch MAXIS_case_number, row, col
                        EMReadScreen nbr_days_pending, 3, row, 50
                        nbr_days_pending = trim(nbr_days_pending)
                        nbr_days_pending = nbr_days_pending * 1
                        IF nbr_days_pending >= 30 THEN ALL_PENDING_CASES_ARRAY(deny_day30, case_entry) = TRUE

                        If ALL_PENDING_CASES_ARRAY(SNAP_status, case_entry) <> "Pending" and ALL_PENDING_CASES_ARRAY(CASH_status, case_entry) = "Pending" Then
                            EMReadScreen cash_prog, 2, row, 56
                            If cash_prog = "MS" Then
                                ALL_PENDING_CASES_ARRAY(error_notes, case_entry) = ALL_PENDING_CASES_ARRAY(error_notes, case_entry) & ", MSA pending only."
                                If datediff("d", ALL_PENDING_CASES_ARRAY(application_date, case_entry), date) >= 60 and ALL_PENDING_CASES_ARRAY(interview_date, case_entry) = "" THEN
                                    ALL_PENDING_CASES_ARRAY(deny_day30, case_entry) = TRUE
                                Else
                                    ALL_PENDING_CASES_ARRAY(deny_day30, case_entry) = FALSE
                                End If
                            End If
                        End If
                        back_to_SELF

                        'THIS IS FOR TESTING'
                        If ALL_PENDING_CASES_ARRAY(deny_day30, case_entry) = TRUE Then

                            ALL_PENDING_CASES_ARRAY(deny_memo_confirm, case_entry) = "Y"
                            ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "REVIEW DENIAL"

                            ReDim Preserve ACTION_TODAY_CASES_ARRAY(error_notes, todays_cases)
                            ACTION_TODAY_CASES_ARRAY(case_number, todays_cases)         = ALL_PENDING_CASES_ARRAY(case_number, case_entry)
                            ACTION_TODAY_CASES_ARRAY(client_name, todays_cases)         = ALL_PENDING_CASES_ARRAY(client_name, case_entry)
                            ACTION_TODAY_CASES_ARRAY(worker_ID, todays_cases)           = ALL_PENDING_CASES_ARRAY(worker_ID, case_entry)
                            ACTION_TODAY_CASES_ARRAY(SNAP_status, todays_cases)         = ALL_PENDING_CASES_ARRAY(SNAP_status, case_entry)
                            ACTION_TODAY_CASES_ARRAY(CASH_status, todays_cases)         = ALL_PENDING_CASES_ARRAY(CASH_status, case_entry)
                            ACTION_TODAY_CASES_ARRAY(application_date, todays_cases)    = ALL_PENDING_CASES_ARRAY(application_date, case_entry)
                            ACTION_TODAY_CASES_ARRAY(interview_date, todays_cases)      = ALL_PENDING_CASES_ARRAY(interview_date, case_entry)
                            ACTION_TODAY_CASES_ARRAY(questionable_intv, todays_cases)   = ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry)
                            ACTION_TODAY_CASES_ARRAY(need_face_to_face, todays_cases)   = ALL_PENDING_CASES_ARRAY(need_face_to_face, case_entry)
                            ACTION_TODAY_CASES_ARRAY(appt_notc_sent, todays_cases)      = ALL_PENDING_CASES_ARRAY(appt_notc_sent, case_entry)
                            ACTION_TODAY_CASES_ARRAY(appt_notc_confirm, todays_cases)   = ALL_PENDING_CASES_ARRAY(appt_notc_confirm, case_entry)
                            ACTION_TODAY_CASES_ARRAY(appointment_date, todays_cases)    = ALL_PENDING_CASES_ARRAY(appointment_date, case_entry)
                            ACTION_TODAY_CASES_ARRAY(nomi_sent, todays_cases)           = ALL_PENDING_CASES_ARRAY(nomi_sent, case_entry)
                            ACTION_TODAY_CASES_ARRAY(nomi_confirm, todays_cases)        = ALL_PENDING_CASES_ARRAY(nomi_confirm, case_entry)
                            ACTION_TODAY_CASES_ARRAY(deny_day30, todays_cases)          = ALL_PENDING_CASES_ARRAY(deny_day30, case_entry)
                            ACTION_TODAY_CASES_ARRAY(deny_memo_confirm, todays_cases)   = ALL_PENDING_CASES_ARRAY(deny_memo_confirm, case_entry)
                            ACTION_TODAY_CASES_ARRAY(next_action_needed, todays_cases)  = ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry)
                            ACTION_TODAY_CASES_ARRAY(error_notes, todays_cases)         = ALL_PENDING_CASES_ARRAY(error_notes, case_entry) & " - " & "DENY today"
                            todays_cases = todays_cases + 1
                        End If

                        'THIS IS FOR REAL'
                        ' If ALL_PENDING_CASES_ARRAY(deny_day30, case_entry) = TRUE Then
                        '     nomi_last_contact_day = dateadd("d", 30, ALL_PENDING_CASES_ARRAY(application_date, case_entry))
                        '     'ensuring that we have given the client an additional10days fromt he day nomi sent'
                        '     IF DateDiff("d", ALL_PENDING_CASES_ARRAY(nomi_sent, case_entry), nomi_last_contact_day) < 1 then nomi_last_contact_day = dateadd("d", 10, ALL_PENDING_CASES_ARRAY(nomi_sent, case_entry))
                        '
                        '     start_a_new_spec_memo_and_continue(memo_started)		'Writes the denial into the MEMO.
                		' 	IF memo_started = True THEN
                		' 		EMsendkey("************************************************************")
                		' 		Call write_variable_in_SPEC_MEMO("We received your application on " & ALL_PENDING_CASES_ARRAY(application_date, case_entry) & ".")
                		' 		Call write_variable_in_SPEC_MEMO("Your interview was not completed by " & nomi_last_contact_day & ".")
                		' 		call write_variable_in_spec_memo("Due to failing to complete the interview within 30 days of your application date your case has been denied.")
                		' 		Call write_variable_in_SPEC_MEMO("************************************************************")
                		' 		PF4
                		' 	ELSE
                		' 		ALL_PENDING_CASES_ARRAY(deny_memo_confirm, case_entry) = "N"         'Setting this as N if the MEMO failed
                		' 		'MsgBox "What memo was sent?"
                		' 	END IF
                        '
                        '     If ALL_PENDING_CASES_ARRAY(deny_memo_confirm, case_entry) <> "N" Then
                        '         Call confirm_memo_waiting(ALL_PENDING_CASES_ARRAY(deny_memo_confirm, case_entry))
                        '     End If
                        '
                        '     If ALL_PENDING_CASES_ARRAY(deny_memo_confirm, case_entry) = "N" Then
                        '         ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "Send DENY MEMO Manually"
                        '     ElseIf ALL_PENDING_CASES_ARRAY(deny_memo_confirm, case_entry) = "Y"
                        '         ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "REVIEW DENIAL"
                        '         Call start_a_blank_case_note
                        '
                        '         Call write_variable_in_case_note("~ Denied " & programs & " via script ~")
                        '         Call write_bullet_and_variable_in_case_note("Application date", ALL_PENDING_CASES_ARRAY(application_date, case_entry))
                        '         Call write_variable_in_case_note("* Reason for denial: interview was not completed timely.")
                        '         Call write_variable_in_case_note("* Confirmed client was provided sufficient 10 day notice.")
                        '         Call write_bullet_and_variable_in_case_note("NOMI sent to client on ", ALL_PENDING_CASES_ARRAY(nomi_sent, case_entry))
                        '         Call write_variable_in_case_note("---")
                        '         Call write_variable_in_CASE_NOTE(worker_signature & " via bulk on demand waiver script")
                        '
                        '         'MsgBox "What casenote was sent?"
                        '         PF3
                        '         END IF
                        '     Else
                        '         ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "???"
                        '     End If
                        '     'msgbox nbr_days_pending
                        '     Call back_to_SELF

                            ' ReDim Preserve ACTION_TODAY_CASES_ARRAY(error_notes, todays_cases)
                            ' ACTION_TODAY_CASES_ARRAY(case_number, todays_cases)         = ALL_PENDING_CASES_ARRAY(case_number, case_entry)
                            ' ACTION_TODAY_CASES_ARRAY(client_name, todays_cases)         = ALL_PENDING_CASES_ARRAY(client_name, case_entry)
                            ' ACTION_TODAY_CASES_ARRAY(worker_ID, todays_cases)           = ALL_PENDING_CASES_ARRAY(worker_ID, case_entry)
                            ' ACTION_TODAY_CASES_ARRAY(SNAP_status, todays_cases)         = ALL_PENDING_CASES_ARRAY(SNAP_status, case_entry)
                            ' ACTION_TODAY_CASES_ARRAY(CASH_status, todays_cases)         = ALL_PENDING_CASES_ARRAY(CASH_status, case_entry)
                            ' ACTION_TODAY_CASES_ARRAY(application_date, todays_cases)    = ALL_PENDING_CASES_ARRAY(application_date, case_entry)
                            ' ACTION_TODAY_CASES_ARRAY(interview_date, todays_cases)      = ALL_PENDING_CASES_ARRAY(interview_date, case_entry)
                            ' ACTION_TODAY_CASES_ARRAY(questionable_intv, todays_cases)   = ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry)
                            ' ACTION_TODAY_CASES_ARRAY(need_face_to_face, todays_cases)   = ALL_PENDING_CASES_ARRAY(need_face_to_face, case_entry)
                            ' ACTION_TODAY_CASES_ARRAY(appt_notc_sent, todays_cases)      = ALL_PENDING_CASES_ARRAY(appt_notc_sent, case_entry)
                            ' ACTION_TODAY_CASES_ARRAY(appt_notc_confirm, todays_cases)   = ALL_PENDING_CASES_ARRAY(appt_notc_confirm, case_entry)
                            ' ACTION_TODAY_CASES_ARRAY(appointment_date, todays_cases)    = ALL_PENDING_CASES_ARRAY(appointment_date, case_entry)
                            ' ACTION_TODAY_CASES_ARRAY(nomi_sent, todays_cases)           = ALL_PENDING_CASES_ARRAY(nomi_sent, case_entry)
                            ' ACTION_TODAY_CASES_ARRAY(nomi_confirm, todays_cases)        = ALL_PENDING_CASES_ARRAY(nomi_confirm, case_entry)
                            ' ACTION_TODAY_CASES_ARRAY(deny_day30, todays_cases)          = ALL_PENDING_CASES_ARRAY(deny_day30, case_entry)
                            ' ACTION_TODAY_CASES_ARRAY(deny_memo_confirm, todays_cases)   = ALL_PENDING_CASES_ARRAY(deny_memo_confirm, case_entry)
                            ' ACTION_TODAY_CASES_ARRAY(next_action_needed, todays_cases)  = ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry)
                            ' ACTION_TODAY_CASES_ARRAY(error_notes, todays_cases)         = ALL_PENDING_CASES_ARRAY(error_notes, case_entry) & " - " & "NOMI Sent today"
                            ' todays_cases = todays_cases + 1
                        ' End If
                    END IF
                END IF
            END IF

        End If

        ALL_PENDING_CASES_ARRAY(deny_day30, case_entry) = ALL_PENDING_CASES_ARRAY(deny_day30, case_entry) & ""
    End If
    row = ALL_PENDING_CASES_ARRAY(excel_row, case_entry)


    ObjWorkExcel.Cells(row, worker_id_col) = ALL_PENDING_CASES_ARRAY(worker_ID, case_entry)
    ObjWorkExcel.Cells(row, case_nbr_col) = ALL_PENDING_CASES_ARRAY(case_number, case_entry)
    ObjWorkExcel.Cells(row, case_name_col) = ALL_PENDING_CASES_ARRAY(client_name, case_entry)
    ObjWorkExcel.Cells(row, snap_stat_col) = ALL_PENDING_CASES_ARRAY(SNAP_status, case_entry)
    ObjWorkExcel.Cells(row, cash_stat_col) = ALL_PENDING_CASES_ARRAY(CASH_status, case_entry)
    ObjWorkExcel.Cells(row, app_date_col) = ALL_PENDING_CASES_ARRAY(application_date, case_entry)

    ObjWorkExcel.Cells(row, intvw_date_col) = ALL_PENDING_CASES_ARRAY(interview_date, case_entry)
    ObjWorkExcel.Cells(row, quest_intvw_date_col) = ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry)
    ObjWorkExcel.Cells(row, ftof_still_need_col) = ALL_PENDING_CASES_ARRAY(need_face_to_face, case_entry)
    ObjWorkExcel.Cells(row, appt_notc_date_col) = ALL_PENDING_CASES_ARRAY(appt_notc_sent, case_entry)
    ObjWorkExcel.Cells(row, appt_date_col) = ALL_PENDING_CASES_ARRAY(appointment_date, case_entry)
    ObjWorkExcel.Cells(row, appt_notc_confirm_col) = ALL_PENDING_CASES_ARRAY(appt_notc_confirm, case_entry)
    ObjWorkExcel.Cells(row, nomi_date_col) = ALL_PENDING_CASES_ARRAY(nomi_sent, case_entry)
    ObjWorkExcel.Cells(row, nomi_confirm_col) = ALL_PENDING_CASES_ARRAY(nomi_confirm, case_entry)
    ObjWorkExcel.Cells(row, need_deny_col) = ALL_PENDING_CASES_ARRAY(deny_day30, case_entry) & ""
    If ALL_PENDING_CASES_ARRAY(deny_day30, case_entry) = "TRUE" Then
        ObjWorkExcel.Rows(row).Font.ColorIndex = 3  'Red'
        ObjWorkExcel.Rows(row).Font.Bold = TRUE
    End If
    ObjWorkExcel.Cells(row, deny_notc_confirm_col) = ALL_PENDING_CASES_ARRAY(deny_memo_confirm, case_entry)
    ObjWorkExcel.Cells(row, next_action_col) = ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry)

    ObjWorkExcel.Cells(row, correct_need_col) = ALL_PENDING_CASES_ARRAY(error_notes, case_entry)
    'ObjWorkExcel.Cells(row, ) = ALL_PENDING_CASES_ARRAY(, case_entry)
Next

'--------------------------CHANGING LINE HERE ---------------------------------------------------'
'         'Finding if an appointment notice has been sent
'         If ALL_PENDING_CASES_ARRAY(appt_notc_sent, case_entry) = "" Then
'             Call navigate_to_MAXIS_screen("CASE", "NOTE")
'             note_row = 5
'             day_before_app = DateAdd("d", -1, ALL_PENDING_CASES_ARRAY(application_date, case_entry)) 'will set the date one day prior to app date'
'             Do
'                 EMReadScreen note_date, 8, note_row, 6
'                 EMReadScreen note_title, 55, note_row, 25
'                 note_title = trim(note_title)
'                 IF left(note_title, 35) = "~ Appointment letter sent in MEMO ~" then
'                     ALL_PENDING_CASES_ARRAY(appt_notc_sent, case_entry) = note_date
' 				ElseIF left(note_title, 42) = "~ Appointment letter sent in MEMO for SNAP" then
'                     ALL_PENDING_CASES_ARRAY(appt_notc_sent, case_entry) = note_date
' 				ElseIF left(note_title, 37) = "~ Appointment letter sent in MEMO for" then
' 					EMReadScreen appt_date, 10, note_row, 63
' 					appt_date = replace(appt_date, "~", "")
' 				 	appt_date = trim(appt_date)
' 					ALL_PENDING_CASES_ARRAY(appointment_date, case_entry) = appt_date
'                     ALL_PENDING_CASES_ARRAY(appt_notc_sent, case_entry) = note_date
'                     'MsgBox ALL_PENDING_CASES_ARRAY(appointment_date, case_entry)
' 				END IF
'                 IF note_date = "        " then Exit Do
'                 note_row = note_row + 1
'                 IF note_row = 19 THEN
'                     PF8
'                     note_row = 5
'                 END IF
'                 EMReadScreen next_note_date, 8, note_row, 6
'                 IF next_note_date = "        " then Exit Do
'             Loop until datevalue(next_note_date) < day_before_app 'looking ahead at the next case note kicking out the dates before app'
'         End If
'
'         If ALL_PENDING_CASES_ARRAY(appt_notc_sent, case_entry) = "" AND ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry)<> "REMOVE FROM LIST" AND ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) <> "NONE - Interview Completed"Then
'             ALL_PENDING_CASES_ARRAY(take_action_today, case_entry) = TRUE
'             ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "SEND APPOINTMENT NOTICE"
'         ElseIf ALL_PENDING_CASES_ARRAY(appointment_date, case_entry) = "" Then
'             Call navigate_to_MAXIS_screen ("SPEC", "MEMO")
'             memo_mo = DatePart("m", ALL_PENDING_CASES_ARRAY(appt_notc_sent, case_entry))
'             memo_mo = right("00"&memo_mo, 2)
'             memo_yr = DatePart("yyyy", ALL_PENDING_CASES_ARRAY(appt_notc_sent, case_entry))
'             memo_yr = right(memo_yr, 2)
'
'             EmWriteScreen memo_mo, 3, 48
'             EmWriteScreen memo_yr, 3, 53
'             transmit
'
'             'creating a variable in the MM/DD/YY format to compare with date read from MAXIS
'             look_mo = DatePart("m", ALL_PENDING_CASES_ARRAY(appt_notc_sent, case_entry))
'             look_mo = right("00" & look_mo, 2)
'
'             look_day = DatePart("d", ALL_PENDING_CASES_ARRAY(appt_notc_sent, case_entry))
'             look_day = right("00" & look_day, 2)
'
'             look_yr = DatePart("yyyy", ALL_PENDING_CASES_ARRAY(appt_notc_sent, case_entry))
'             look_yr = right(look_yr, 2)
'
'             look_date = look_mo & "/" & look_day & "/" & look_yr
'             Do
'                 EMReadScreen create_date, 8, memo_row, 19                 'Reading the date of each memo and the status
'                 EMReadScreen print_status, 7, memo_row, 67
'                 'MsgBox print_status
'                 If create_date = look_date AND print_status = "Printed" Then   'MEMOs created today and still waiting is likely our MEMO.
'                     EmWriteScreen "X", memo_row, 16
'                     transmit
'                     PF8
'
'                     EMReadScreen start_of_msg, 35, 15,12
'                     If start_of_msg = "You recently applied for assistance" Then
'                         EMReadScreen date_in_memo, 10, 19, 47
'                         date_in_memo = trim(date_in_memo)
'                         ALL_PENDING_CASES_ARRAY(appointment_date, case_entry) = replace(date_in_memo, ".", "")
'                         'MsgBox ALL_PENDING_CASES_ARRAY(notc_confirm, case_entry)                 'For statistical purposes
'                         Pf3
'                         Exit Do
'                     End If
'                     PF3
'                 End If
'                 memo_row = memo_row + 1           'Looking at next row'
'             Loop Until create_date = "        "
'         End If
'
'         ALL_PENDING_CASES_ARRAY(interview_date, case_entry) = trim(ALL_PENDING_CASES_ARRAY(interview_date, case_entry))
'         If ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "" Then
'             If ALL_PENDING_CASES_ARRAY(interview_date, case_entry) <> "" Then
'                 ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "NONE - Interview Completed"
'             ElseIf ALL_PENDING_CASES_ARRAY(appt_notc_sent, case_entry) = "" Then
'                 ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "SEND APPOINTMENT NOTICE"
'             ElseIf ALL_PENDING_CASES_ARRAY(appointment_date, case_entry) = "" Then
'                 ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "DETERMINE APPOINTMENT DATE"
'             ElseIf ALL_PENDING_CASES_ARRAY(nomi_sent, case_entry) = "" Then
'                 ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "SEND NOMI"
'             Else
'                 ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "DENY AT DAY 30"
'             End If
'         End If
'
'         search_case_notes_for_interview = FALSE
'         upcoming_nomi = FALSE
'         If ALL_PENDING_CASES_ARRAY(appointment_date, case_entry) <> "" Then
'             If DateDiff("d", ALL_PENDING_CASES_ARRAY(appointment_date, case_entry), date) >= -1 then upcoming_nomi = TRUE
'         End If
'
'         If ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "SEND APPOINTMENT NOTICE" Then search_case_notes_for_interview = TRUE
'         IF upcoming_nomi = TRUE AND ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "SEND NOMI" Then search_case_notes_for_interview = TRUE
'         If ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "DENY AT DAY 30" AND DateDiff("d", ALL_PENDING_CASES_ARRAY(application_date, case_entry), date) >= 29 Then search_case_notes_for_interview = TRUE
'
'         If search_case_notes_for_interview = TRUE Then
'             Call navigate_to_MAXIS_screen("CASE", "NOTE")
'             note_row = 5
'             day_before_app = DateAdd("d", -1, ALL_PENDING_CASES_ARRAY(application_date, case_entry)) 'will set the date one day prior to app date'
'             start_dates = ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry)
'             Do
'                 EMReadScreen note_date, 8, note_row, 6
'                 EMReadScreen note_title, 55, note_row, 25
'                 note_title = trim(note_title)
'                 check_this_date = TRUE
'
'                 If InStr(ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry), "~") <> 0 Then
'                     array_of_dates = ""
'                     array_of_dates = split(ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry), "~")
'                     For each dates in array_of_dates
'                         Call convert_to_mainframe_date(dates, 2)
'                         MsgBox "Already known questionable date: " & dates
'                         if dates = note_date Then check_this_date = FALSE
'                     Next
'                 Else
'                     if ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry) = note_date Then check_this_date = FALSE
'                 End If
'
'                 If check_this_date = TRUE Then 'if a questionable interview date is left on the spreadsheet - that means it has been reviewed and is NOT an interview.
'                     IF left(note_title, 15) = "***Add program:" then
'                         ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry) = ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry) & "~" & note_date
'                         ALL_PENDING_CASES_ARRAY(error_notes, case_entry) = ALL_PENDING_CASES_ARRAY(error_notes, case_entry) & ", Case Note suggests interview completed but interview not listed on PROG."
'                     END IF
'                     IF left(note_title, 33) = "***Intake Interview Completed ***" then
'                         ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry) = ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry) & "~" & note_date
'                         ALL_PENDING_CASES_ARRAY(error_notes, case_entry) = ALL_PENDING_CASES_ARRAY(error_notes, case_entry) & ", Case Note suggests interview completed but interview not listed on PROG."
'                     END IF
'                     IF left(note_title, 40) = "***Reapplication Interview Completed ***" then
'                         ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry) = ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry) & "~" & note_date
'                         ALL_PENDING_CASES_ARRAY(error_notes, case_entry) = ALL_PENDING_CASES_ARRAY(error_notes, case_entry) & ", Case Note suggests interview completed but interview not listed on PROG."
'                     END IF
'                     IF left(note_title, 42) = "~ Interview Completed for SNAP ~" then
'                         ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry) = ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry) & "~" & note_date
'                         ALL_PENDING_CASES_ARRAY(error_notes, case_entry) = ALL_PENDING_CASES_ARRAY(error_notes, case_entry) & ", Case Note suggests interview completed but interview not listed on PROG."
'                     END IF
'                     IF left(note_title, 42) = "*client interviewed* onboarding processing" then
'                         ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry) = ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry) & "~" & note_date
'                         ALL_PENDING_CASES_ARRAY(error_notes, case_entry) = ALL_PENDING_CASES_ARRAY(error_notes, case_entry) & ", Case Note suggests interview completed but interview not listed on PROG."
'                     END IF
'                     IF left(note_title, 34) = "***Intake: pending mentor review**" then
'                         ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry) = ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry) & "~" & note_date
'                         ALL_PENDING_CASES_ARRAY(error_notes, case_entry) = ALL_PENDING_CASES_ARRAY(error_notes, case_entry) & ", Case Note suggests interview completed but interview not listed on PROG."
'                     END IF
'                     IF left(note_title, 23) = "~ Interview Completed ~" then
'                         ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry) = ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry) & "~" & note_date
'                         ALL_PENDING_CASES_ARRAY(error_notes, case_entry) = ALL_PENDING_CASES_ARRAY(error_notes, case_entry) & ", Case Note suggests interview completed but interview not listed on PROG."
'                     END IF
'                     IF left(note_title, 10) = "***Intake:" then
'                         ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry) = ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry) & "~" & note_date
'                         ALL_PENDING_CASES_ARRAY(error_notes, case_entry) = ALL_PENDING_CASES_ARRAY(error_notes, case_entry) & ", Case Note suggests interview completed but interview not listed on PROG."
'                     END IF
'                     IF left(note_title, 24) = "~ Application interview ~" then
'                         ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry) = ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry) & "~" & note_date
'                         ALL_PENDING_CASES_ARRAY(error_notes, case_entry) = ALL_PENDING_CASES_ARRAY(error_notes, case_entry) & ", This case may not require an interview."
'                     END IF
'                     IF left(note_title, 33) = "***Intake Interview Completed ***" then
'                         ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry) = ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry) & "~" & note_date
'                         ALL_PENDING_CASES_ARRAY(error_notes, case_entry) = ALL_PENDING_CASES_ARRAY(error_notes, case_entry) & ", Case Note suggests interview completed but interview not listed on PROG."
'                     END IF
'                     IF left(UCase(note_title), 51) = "Phone call from client re: Phone interview Complete" then
'                         ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry) = ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry) & "~" & note_date
'                         ALL_PENDING_CASES_ARRAY(error_notes, case_entry) = ALL_PENDING_CASES_ARRAY(error_notes, case_entry) & ", Possible case note indicating an interview. If an interview happened, PROG was not updated and an incorrect script was used."
'                     END IF
'                     IF left(UCase(note_title), 41) = "Phone call from client re: SNAP interview" then
'                         ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry) = ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry) & "~" & note_date
'                         ALL_PENDING_CASES_ARRAY(error_notes, case_entry) = ALL_PENDING_CASES_ARRAY(error_notes, case_entry) & ", Possible case note indicating an interview. If an interview happened, PROG was not updated and an incorrect script was used."
'                     END IF
'                 End If
'                 IF note_date = "        " then Exit Do
'                 note_row = note_row + 1
'                 IF note_row = 19 THEN
'                     PF8
'                     note_row = 5
'                 END IF
'                 EMReadScreen next_note_date, 8, note_row, 6
'                 IF next_note_date = "        " then Exit Do
'             Loop until datevalue(next_note_date) < day_before_app 'looking ahead at the next case note kicking out the dates before app'
'
'             If left(ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry), 1) = "~" Then ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry) = right(ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry), len(ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry))-1)
'             if ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry) <> start_dates Then ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "REVIEW QUESTIONABLE INTERVIEW DATE(S)"
'         End If
'
'
'         If ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "SEND NOMI" Then
'             nomi_due = FALSE
'             If ALL_PENDING_CASES_ARRAY(appointment_date, case_entry) <> "" Then
'                 If DateDiff("d", ALL_PENDING_CASES_ARRAY(appointment_date, case_entry), date) >= -1 then nomi_due = TRUE
'             End If
'             If nomi_due = TRUE and ALL_PENDING_CASES_ARRAY(nomi_sent, case_entry) = "" Then
'                 Call navigate_to_MAXIS_screen("CASE", "NOTE")
'                 note_row = 5
'                 day_before_app = DateAdd("d", -1, ALL_PENDING_CASES_ARRAY(application_date, case_entry)) 'will set the date one day prior to app date'
'                 Do
'                     EMReadScreen note_date, 8, note_row, 6
'                     EMReadScreen note_title, 55, note_row, 25
'                     note_title = trim(note_title)
'
'                     IF note_title = "~ Client missed application interview, NOMI sent via sc" then ALL_PENDING_CASES_ARRAY(nomi_sent, case_entry) = note_date
'                     IF left(note_title, 32) = "**Client missed SNAP interview**" then ALL_PENDING_CASES_ARRAY(nomi_sent, case_entry) = note_date
'     				IF left(note_title, 32) = "**Client missed CASH interview**" then ALL_PENDING_CASES_ARRAY(nomi_sent, case_entry) = note_date
'     				IF left(note_title, 37) = "**Client missed SNAP/CASH interview**" then ALL_PENDING_CASES_ARRAY(nomi_sent, case_entry) = note_date
'     				IF note_title = "~ Client has not completed application interview, NOMI" then ALL_PENDING_CASES_ARRAY(nomi_sent, case_entry) = note_date
'
'                     IF note_date = "        " then Exit Do
'                     note_row = note_row + 1
'                     IF note_row = 19 THEN
'                         PF8
'                         note_row = 5
'                     END IF
'                     EMReadScreen next_note_date, 8, note_row, 6
'                     IF next_note_date = "        " then Exit Do
'                 Loop until datevalue(next_note_date) < day_before_app 'looking ahead at the next case note kicking out the dates before app'
'             End If
'             If ALL_PENDING_CASES_ARRAY(nomi_sent, case_entry) <> "" Then ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "DENY AT DAY 30"
'         End If
'
'     End If
' Next
'
' For case_entry = 0 to UBOUND(ALL_PENDING_CASES_ARRAY, 2)
'     MAXIS_case_number	= ALL_PENDING_CASES_ARRAY(case_number, case_entry)
'
'     If ALL_PENDING_CASES_ARRAY(priv_case, case_entry) = FALSE Then
'
'         If ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "SEND APPOINTMENT NOTICE" Then
'             Call Navigate_to_MAXIS_screen("CASE", "NOTE")
'             'MsgBox "We're Sending an Appointment Notice."
'
'             ALL_PENDING_CASES_ARRAY(appt_notc_sent, case_entry) = date
'             ALL_PENDING_CASES_ARRAY(appt_notc_confirm, case_entry) = "Y"
'             ALL_PENDING_CASES_ARRAY(appointment_date, case_entry) = DateAdd("d", 7, date)
'
'             ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "ApptNotc Sent - SEND NOMI"
'             Call back_to_SELF
'         End If
'         If ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "SEND NOMI" Then
'             nomi_due = FALSE
'             If ALL_PENDING_CASES_ARRAY(appointment_date, case_entry) <> "" Then
'                 If DateDiff("d", ALL_PENDING_CASES_ARRAY(appointment_date, case_entry), date) >= 0 then nomi_due = TRUE
'             End If
'
'             If nomi_due Then
'                 Call Navigate_to_MAXIS_screen("CASE", "NOTE")
'                 'MsgBox "We're Sending a NOMI."
'
'                 ALL_PENDING_CASES_ARRAY(nomi_sent, case_entry) = date
'                 ALL_PENDING_CASES_ARRAY(nomi_confirm, case_entry) = "Y"
'
'                 ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "NOMI sent - DENY AT DAY 30"
'                 Call back_to_SELF
'             ENd If
'         End If
'         If ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "DENY AT DAY 30" Then
'             IF datediff("d", ALL_PENDING_CASES_ARRAY(application_date, case_entry), date) >= 30 and ALL_PENDING_CASES_ARRAY(interview_date, case_entry) = "" THEN
'     			'MsgBox "Both false notice"
'     			'MsgBox ALL_PENDING_CASES_ARRAY(nomi_sent, case_entry)
'                 IF ALL_PENDING_CASES_ARRAY(nomi_sent, case_entry) <> "" then
'                     day_30 = dateadd("d", 30, ALL_PENDING_CASES_ARRAY(application_date, case_entry))
'     				IF datediff("d", ALL_PENDING_CASES_ARRAY(nomi_sent, case_entry), date) >= 10 or datediff("d", ALL_PENDING_CASES_ARRAY(nomi_sent, case_entry), day_30) > 0 THEN
'     				'MsgBox datediff("d", ALL_PENDING_CASES_ARRAY(nomi_sent, case_entry), date)
'     					Call navigate_to_MAXIS_screen("REPT", "PND2")
'     					Row = 1
'     					Col = 1
'     					EMSearch MAXIS_case_number, row, col
'     					EMReadScreen nbr_days_pending, 3, row, 50
'     		  		    nbr_days_pending = trim(nbr_days_pending)
'     					nbr_days_pending = nbr_days_pending * 1
'     					IF nbr_days_pending >= 30 THEN ALL_PENDING_CASES_ARRAY(deny_day30, case_entry) = TRUE
'     					'msgbox nbr_days_pending
'     				END IF
'                 END IF
'     		END IF
'
'             If ALL_PENDING_CASES_ARRAY(deny_day30, case_entry) = TRUE Then
'                 Call Navigate_to_MAXIS_screen("CASE", "NOTE")
'                 'MsgBox "DENIAL time."
'
'                 ALL_PENDING_CASES_ARRAY(deny_memo_confirm, case_entry) = "Y"
'                 ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "REVIEW DENIAL"
'
'                 Call back_to_SELF
'             End If
'         End If
'
'
'         'If
'         row = ALL_PENDING_CASES_ARRAY(excel_row, case_entry)
'
'
'         ObjWorkExcel.Cells(row, worker_id_col) = ALL_PENDING_CASES_ARRAY(worker_ID, case_entry)
'         ObjWorkExcel.Cells(row, case_nbr_col) = ALL_PENDING_CASES_ARRAY(case_number, case_entry)
'         ObjWorkExcel.Cells(row, case_name_col) = ALL_PENDING_CASES_ARRAY(client_name, case_entry)
'         ObjWorkExcel.Cells(row, snap_stat_col) = ALL_PENDING_CASES_ARRAY(SNAP_status, case_entry)
'         ObjWorkExcel.Cells(row, cash_stat_col) = ALL_PENDING_CASES_ARRAY(CASH_status, case_entry)
'         ObjWorkExcel.Cells(row, app_date_col) = ALL_PENDING_CASES_ARRAY(application_date, case_entry)
'
'         ObjWorkExcel.Cells(row, intvw_date_col) = ALL_PENDING_CASES_ARRAY(interview_date, case_entry)
'         ObjWorkExcel.Cells(row, quest_intvw_date_col) = ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry)
'         ObjWorkExcel.Cells(row, appt_notc_date_col) = ALL_PENDING_CASES_ARRAY(appt_notc_sent, case_entry)
'         ObjWorkExcel.Cells(row, appt_date_col) = ALL_PENDING_CASES_ARRAY(appointment_date, case_entry)
'         ObjWorkExcel.Cells(row, appt_notc_confirm_col) = ALL_PENDING_CASES_ARRAY(appt_notc_confirm, case_entry)
'         ObjWorkExcel.Cells(row, nomi_date_col) = ALL_PENDING_CASES_ARRAY(nomi_sent, case_entry)
'         ObjWorkExcel.Cells(row, nomi_confirm_col) = ALL_PENDING_CASES_ARRAY(nomi_confirm, case_entry)
'         ObjWorkExcel.Cells(row, need_deny_col) = ALL_PENDING_CASES_ARRAY(deny_day30, case_entry)
'         ObjWorkExcel.Cells(row, deny_notc_confirm_col) = ALL_PENDING_CASES_ARRAY(deny_memo_confirm, case_entry)
'         ObjWorkExcel.Cells(row, next_action_col) = ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry)
'
'         ObjWorkExcel.Cells(row, correct_need_col) = ALL_PENDING_CASES_ARRAY(error_notes, case_entry)
'         'ObjWorkExcel.Cells(row, ) = ALL_PENDING_CASES_ARRAY(, case_entry)
'
'
'
'
'     End If
' Next
'Goes through the working excel spreadsheet and compares the other list
'identifies the cases that are not already on the working list
'identifies the cases that are on the working list and NOT on the current day pending list to move them off of the working list
'uses the working list to identify cases that need action taken on them or need to be checked for something.

'Undoing the autofit because this list should remain set up the way the worker wants it.
' For col_to_autofit = 1 to 20
'     ObjWorkExcel.Columns(col_to_autofit).AutoFit()
' Next

call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)

ObjExcel.Worksheets.Add().Name = "Cases Removed From Working LIST"

ObjExcel.Cells(1, worker_id_col)        = "Worker ID"
ObjExcel.Cells(1, case_nbr_col)         = "Case Number"
ObjExcel.Cells(1, case_name_col)        = "Case Name"
ObjExcel.Cells(1, snap_stat_col)        = "SNAP"
ObjExcel.Cells(1, cash_stat_col)        = "CASH"
ObjExcel.Cells(1, app_date_col)         = "Application Date"
ObjExcel.Cells(1, intvw_date_col)       = "Interview Date"
ObjExcel.Cells(1, quest_intvw_date_col) = "Questionable Interview Date"
ObjExcel.Cells(1, ftof_still_need_col)  = "Face To Face Still Needed"
ObjExcel.Cells(1, appt_notc_date_col)   = "Appt Notice Sent"
ObjExcel.Cells(1, appt_date_col)        = "Appointment Date"
ObjExcel.Cells(1, appt_notc_confirm_col)= "Confirm"
ObjExcel.Cells(1, nomi_date_col)        = "NOMI Sent"
ObjExcel.Cells(1, nomi_confirm_col)     = "Confirm"
ObjExcel.Cells(1, need_deny_col)        = "Denial"
ObjExcel.Cells(1, deny_notc_confirm_col)= "Confirm"
ObjExcel.Cells(1, next_action_col)      = "Next Action"
ObjExcel.Cells(1, correct_need_col)     = "Detail"
' ObjExcel.Cells(1, action_worker_col)    =
' ObjExcel.Cells(1, action_sup_col)       =
' ObjExcel.Cells(1, email_sent_col)       =

ObjExcel.Rows(1).Font.Bold = TRUE

removed_row = 2
For case_removed = 0 to UBOUND(CASES_NO_LONGER_WORKING, 2)
    If CASES_NO_LONGER_WORKING(error_notes, case_removed) = "" OR CASES_NO_LONGER_WORKING(client_name, case_removed) = "XXXXX" Then
        'PROG to determine programs active
        MAXIS_case_number = CASES_NO_LONGER_WORKING(case_number, case_removed)
        CALL navigate_to_MAXIS_screen("CASE", "CURR")
        'Checking for PRIV cases.
        EMReadScreen priv_check, 6, 24, 14 'If it can't get into the case needs to skip
        EMReadScreen county_check, 2, 21, 16    'Looking to see if case has Hennepin COunty worker
        If priv_check = "PRIVIL" THEN
            CASES_NO_LONGER_WORKING(error_notes, case_removed) = "PRIV"
        ElseIf county_check <> "27" THEN
            CASES_NO_LONGER_WORKING(error_notes, case_removed) = "Transferred out of county - " & county_check
        ElseIf CASES_NO_LONGER_WORKING(client_name, case_removed) = "XXXXX" Then
            Call navigate_to_MAXIS_screen("STAT", "MEMB")
            EMReadScreen last_name, 25, 6, 30
            EMReadScreen first_name, 12, 6, 63
            EMReadScreen middle_initial, 1, 6, 79

            last_name = replace(last_name, "_", "")
            first_name = replace(first_name, "_", "")
            middle_initial = replace(middle_initial, "_", "")

            CASES_NO_LONGER_WORKING(client_name, case_removed) = last_name & ", " & first_name & " " & middle_initial
        End If

        If CASES_NO_LONGER_WORKING(error_notes, case_removed) = "" Then

            Call navigate_to_MAXIS_screen("STAT", "PROG")
            fs_intv = ""
            cash_intv_one = ""
            cash_intv_two = ""

            EMReadScreen cash_prog_one, 2, 6, 67               'reading for active MFIP program - which has different requirements
            EMReadScreen cash_stat_one, 4, 6, 74

            EMReadScreen cash_prog_two, 2, 7, 67
            EMReadScreen cash_stat_two, 4, 7, 74

            EMReadScreen fs_pend, 4, 10, 74

            cash_pend = FALSE
            cash_interview_done = FALSE
            snap_interview_done = FALSE

            If cash_stat_one = "PEND" Then
                cash_pend = TRUE
                EMReadScreen cash_intv_one, 8, 6, 55
                If cash_intv_one <> "__ __ __" Then
                    cash_intv_one = replace(cash_intv_one, " ", "/")
                    cash_interview_done = TRUE
                Else
                    cash_intv_one = ""
                End If
            ElseIf cash_stat_one = "ACTV" Then
                CASES_NO_LONGER_WORKING(CASH_status, case_removed) = "Active"
            End If

            If cash_stat_two = "PEND" Then
                cash_pend = TRUE
                EMReadScreen cash_intv_two, 8, 7, 55
                If cash_intv_two <> "__ __ __" Then
                    cash_intv_two = replace(cash_intv_two, " ", "/")
                    cash_interview_done = TRUE
                Else
                    cash_intv_two = ""
                End If
            ElseIf cash_stat_one = "ACTV" Then
                CASES_NO_LONGER_WORKING(CASH_status, case_removed) = "Active"
            Else
                CASES_NO_LONGER_WORKING(CASH_status, case_removed) = ""
            End If

            If cash_pend = TRUE then CASES_NO_LONGER_WORKING(CASH_status, case_removed) = "Pending"

            If fs_pend = "PEND" Then
                CASES_NO_LONGER_WORKING(SNAP_status, case_removed) = "Pending"
                EMReadScreen fs_intv, 8, 10, 55
                If fs_intv <> "__ __ __" Then
                    fs_intv = replace(fs_intv, " ", "/")
                    snap_interview_done = TRUE
                Else
                    fs_intv = ""
                End If
            ElseIf fs_pend = "ACTV" Then
                CASES_NO_LONGER_WORKING(SNAP_status, case_removed) = "Active"
            Else
                CASES_NO_LONGER_WORKING(SNAP_status, case_removed) = ""
            End If

            If CASES_NO_LONGER_WORKING(SNAP_status, case_removed) <> "Pending" AND CASES_NO_LONGER_WORKING(CASH_status, case_removed) <> "Pending" Then
                CASES_NO_LONGER_WORKING(next_action_needed, case_removed) = "REMOVE FROM LIST"
                CASES_NO_LONGER_WORKING(error_notes, case_removed) = "Neither SNAP nor CASH is pending."
            Else
                If cash_pend = TRUE Then
                    If cash_interview_done = TRUE Then
                        If cash_intv_one <> "" Then CASES_NO_LONGER_WORKING(interview_date, case_removed) = cash_intv_one
                        If cash_intv_two <> "" Then CASES_NO_LONGER_WORKING(interview_date, case_removed) = cash_intv_two
                        CASES_NO_LONGER_WORKING(need_face_to_face, case_removed) = ""
                        CASES_NO_LONGER_WORKING(next_action_needed, case_removed) = "NONE - Interview Completed"
                    Else
                        If fs_pend = "PEND" Then
                            If fs_intv = "" THen
                                CASES_NO_LONGER_WORKING(interview_date, case_removed) = ""
                            Else
                                CASES_NO_LONGER_WORKING(interview_date, case_removed) = fs_intv
                                If CASES_NO_LONGER_WORKING(need_face_to_face, case_removed) = "" Then CASES_NO_LONGER_WORKING(next_action_needed, case_removed) = "CHECK FOR F2F NEEDED"
                                If CASES_NO_LONGER_WORKING(need_face_to_face, case_removed) = "N" Then CASES_NO_LONGER_WORKING(next_action_needed, case_removed) = "NONE - Interview Completed"
                                If CASES_NO_LONGER_WORKING(need_face_to_face, case_removed) = "Y" Then
                                    CASES_NO_LONGER_WORKING(next_action_needed, case_removed) = "DENY AT DAY 30"
                                    If CASES_NO_LONGER_WORKING(nomi_sent, case_removed) = "" Then CASES_NO_LONGER_WORKING(next_action_needed, case_removed) = "SEND NOMI"
                                    IF CASES_NO_LONGER_WORKING(sppt_notc_sent, case_removed) = "" Then CASES_NO_LONGER_WORKING(next_action_needed, case_removed) = "SEND APPOINTMENT NOTICE"
                                End If
                                CASES_NO_LONGER_WORKING(error_notes, case_removed) = ", Cash interview incomplete."
                        'WHAT TO DO WITH F2F Cases'
                            End If
                        End If
                    End If
                ElseIf fs_pend = "PEND" Then
                    If fs_intv <> "" Then
                        CASES_NO_LONGER_WORKING(interview_date, case_removed) = fs_intv
                        CASES_NO_LONGER_WORKING(next_action_needed, case_removed) = "NONE - Interview Completed"
                        CASES_NO_LONGER_WORKING(need_face_to_face, case_removed) = ""
                    End If
                End If
            End If
        End If
    End If

    ObjExcel.Worksheets("Cases Removed From Working LIST").Activate
    'MsgBox "Row is " & removed_row & vbNewLine & "Worker ID " & CASES_NO_LONGER_WORKING(worker_ID, case_removed)
    ObjExcel.Cells(removed_row, worker_id_col).Value            = CASES_NO_LONGER_WORKING(worker_ID, case_removed)
    ObjExcel.Cells(removed_row, case_nbr_col).Value             = CASES_NO_LONGER_WORKING(case_number, case_removed)
    'CASES_NO_LONGER_WORKING(excel_removed_row, case_removed) = removed_row
    ObjExcel.Cells(removed_row, case_name_col).Value            = CASES_NO_LONGER_WORKING(client_name, case_removed)
    ObjExcel.Cells(removed_row, app_date_col).Value             = CASES_NO_LONGER_WORKING(application_date, case_removed)
    'CASES_NO_LONGER_WORKING(interview_date, case_removed) = ObjExcel.Cells(removed_row, intvw_date_col)
    ObjExcel.Cells(removed_row, intvw_date_col).Value           = CASES_NO_LONGER_WORKING(interview_date, case_removed)
    ObjExcel.Cells(removed_row, cash_stat_col).Value            = CASES_NO_LONGER_WORKING(CASH_status, case_removed)
    ObjExcel.Cells(removed_row, snap_stat_col).Value            = CASES_NO_LONGER_WORKING(SNAP_status, case_removed)

    ObjExcel.Cells(removed_row, appt_notc_date_col).Value       = CASES_NO_LONGER_WORKING(appt_notc_sent, case_removed)
    ObjExcel.Cells(removed_row, appt_notc_confirm_col).Value    = CASES_NO_LONGER_WORKING(appt_notc_confirm, case_removed)
    ObjExcel.Cells(removed_row, appt_date_col).Value            = CASES_NO_LONGER_WORKING(appointment_date, case_removed)
    ObjExcel.Cells(removed_row, nomi_date_col).Value            = CASES_NO_LONGER_WORKING(nomi_sent, case_removed)
    ObjExcel.Cells(removed_row, nomi_confirm_col).Value         = CASES_NO_LONGER_WORKING(nomi_confirm, case_removed)
    ObjExcel.Cells(removed_row, next_action_col).Value          = CASES_NO_LONGER_WORKING(next_action_needed, case_removed)
    ObjExcel.Cells(removed_row, quest_intvw_date_col).Value     = CASES_NO_LONGER_WORKING(questionable_intv, case_removed)
    ObjExcel.Cells(removed_row, ftof_still_need_col).Value     = CASES_NO_LONGER_WORKING(need_face_to_face, case_removed)

    ObjExcel.Cells(removed_row, correct_need_col).Value         = CASES_NO_LONGER_WORKING(error_notes, case_removed)

    'MsgBox removed_row & " = " & removed_row & " + 1"
    removed_row = removed_row + 1
Next

For col_to_autofit =1 to  correct_need_col
    ObjExcel.Columns(col_to_autofit).AutoFit()
Next


ObjExcel.Worksheets.Add().Name = "Actions Today"

ObjExcel.Cells(1, worker_id_col)        = "Worker ID"
ObjExcel.Cells(1, case_nbr_col)         = "Case Number"
ObjExcel.Cells(1, case_name_col)        = "Case Name"
ObjExcel.Cells(1, snap_stat_col)        = "SNAP"
ObjExcel.Cells(1, cash_stat_col)        = "CASH"
ObjExcel.Cells(1, app_date_col)         = "Application Date"
ObjExcel.Cells(1, intvw_date_col)       = "Interview Date"
ObjExcel.Cells(1, quest_intvw_date_col) = "Questionable Interview Date"
ObjExcel.Cells(1, ftof_still_need_col)  = "Face To Face Still Needed"
ObjExcel.Cells(1, appt_notc_date_col)   = "Appt Notice Sent"
ObjExcel.Cells(1, appt_date_col)        = "Appointment Date"
ObjExcel.Cells(1, appt_notc_confirm_col)= "Confirm"
ObjExcel.Cells(1, nomi_date_col)        = "NOMI Sent"
ObjExcel.Cells(1, nomi_confirm_col)     = "Confirm"
ObjExcel.Cells(1, need_deny_col)        = "Denial"
ObjExcel.Cells(1, deny_notc_confirm_col)= "Confirm"
ObjExcel.Cells(1, next_action_col)      = "Next Action"
ObjExcel.Cells(1, correct_need_col)     = "Detail"
' ObjExcel.Cells(1, action_worker_col)    =
' ObjExcel.Cells(1, action_sup_col)       =
' ObjExcel.Cells(1, email_sent_col)       =

ObjExcel.Rows(1).Font.Bold = TRUE

action_row = 2
For action_case = 0 to UBOUND(ACTION_TODAY_CASES_ARRAY, 2)

    ObjExcel.Cells(action_row, worker_id_col)        = ACTION_TODAY_CASES_ARRAY(worker_ID, action_case)
    ObjExcel.Cells(action_row, case_nbr_col)         = ACTION_TODAY_CASES_ARRAY(case_number, action_case)
    ObjExcel.Cells(action_row, case_name_col)        = ACTION_TODAY_CASES_ARRAY(client_name, action_case)
    ObjExcel.Cells(action_row, snap_stat_col)        = ACTION_TODAY_CASES_ARRAY(SNAP_status, action_case)
    ObjExcel.Cells(action_row, cash_stat_col)        = ACTION_TODAY_CASES_ARRAY(CASH_status, action_case)
    ObjExcel.Cells(action_row, app_date_col)         = ACTION_TODAY_CASES_ARRAY(application_date, action_case)
    ObjExcel.Cells(action_row, intvw_date_col)       = ACTION_TODAY_CASES_ARRAY(interview_date, action_case)
    ObjExcel.Cells(action_row, quest_intvw_date_col) = ACTION_TODAY_CASES_ARRAY(questionable_intv, action_case)
    ObjExcel.Cells(action_row, ftof_still_need_col)  = ACTION_TODAY_CASES_ARRAY(need_face_to_face, action_case)
    ObjExcel.Cells(action_row, appt_notc_date_col)   = ACTION_TODAY_CASES_ARRAY(appt_notc_sent, action_case)
    ObjExcel.Cells(action_row, appt_notc_confirm_col)= ACTION_TODAY_CASES_ARRAY(appt_notc_confirm, action_case)
    ObjExcel.Cells(action_row, appt_date_col)        = ACTION_TODAY_CASES_ARRAY(appointment_date, action_case)
    ObjExcel.Cells(action_row, nomi_date_col)        = ACTION_TODAY_CASES_ARRAY(nomi_sent, action_case)
    ObjExcel.Cells(action_row, nomi_confirm_col)     = ACTION_TODAY_CASES_ARRAY(nomi_confirm, action_case)
    ObjExcel.Cells(action_row, need_deny_col)        = ACTION_TODAY_CASES_ARRAY(deny_day30, action_case)
    ObjExcel.Cells(action_row, deny_notc_confirm_col)= ACTION_TODAY_CASES_ARRAY(deny_memo_confirm, action_case)
    ObjExcel.Cells(action_row, next_action_col)      = ACTION_TODAY_CASES_ARRAY(next_action_needed, action_case)
    ObjExcel.Cells(action_row, correct_need_col)     = ACTION_TODAY_CASES_ARRAY(error_notes, action_case)

    action_row = action_row + 1
Next

For col_to_autofit =1 to  correct_need_col
    ObjExcel.Columns(col_to_autofit).AutoFit()
Next

script_end_procedure("It worked!")
