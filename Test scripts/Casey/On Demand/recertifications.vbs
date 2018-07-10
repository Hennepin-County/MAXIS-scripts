'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - RECERTIFICATIONS.vbs"
start_time = timer
STATS_counter = 0			 'sets the stats counter at one
STATS_manualtime = 304			 'manual run time in seconds
STATS_denomination = "C"		 'C is for each case
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
CALL changelog_update("12/29/2017", "Coordinates for sending MEMO's has changed in SPEC/MEMO. Updated script to support change.", "Ilse Ferris, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================


'----------FUNCTIONS----------
'-----This function needs to be added to the FUNCTIONS FILE-----
'>>>>> This function converts the letter for a number so the script can work with it <<<<<
FUNCTION convert_excel_letter_to_excel_number(excel_col)
	IF isnumeric(excel_col) = FALSE THEN
		alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
		excel_col = ucase(excel_col)
		IF len(excel_col) = 1 THEN
			excel_col = InStr(alphabet, excel_col)
		ELSEIF len(excel_col) = 2 THEN
			excel_col = (26 * InStr(alphabet, left(excel_col, 1))) + (InStr(alphabet, right(excel_col, 1)))
		END IF
	ELSE
		excel_col = CInt(excel_col)
	END IF
END FUNCTION

'defining this function here because it needs to not end the script if a MEMO fails.
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

Function HCRE_panel_bypass()
	'handling for cases that do not have a completed HCRE panel
	PF3		'exits PROG to prommpt HCRE if HCRE insn't complete
	Do
		EMReadscreen HCRE_panel_check, 4, 2, 50
		If HCRE_panel_check = "HCRE" then
			PF10	'exists edit mode in cases where HCRE isn't complete for a member
			PF3
		END IF
	Loop until HCRE_panel_check <> "HCRE"
End Function

'DIALOGS ===================================================================================================================

'Initial Dialog which requests a file path for the excel file
BeginDialog recert_list_dlg, 0, 0, 361, 105, "On Demand Recertifications"
  EditBox 130, 60, 175, 15, recertification_cases_excel_file_path
  ButtonGroup ButtonPressed
    PushButton 310, 60, 45, 15, "Browse...", select_a_file_button
  EditBox 75, 85, 140, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 250, 85, 50, 15
    CancelButton 305, 85, 50, 15
  Text 10, 10, 170, 10, "Welcome to the On Demand Recertification Notifier."
  Text 10, 25, 340, 30, "This script will send an Appointment Notice or NOMI for recertification for a list of cases in a county that currently has an On Demand Waiver in effect for interviews. If your county does not have this waiver, this script should not be used."
  Text 10, 65, 120, 10, "Select an Excel file for recert cases:"
  Text 10, 90, 60, 10, "Worker Signature"
EndDialog


'Confirmation Diaglog will require worker to afirm the appointment notices/NOMIs should actually be sent

'END DIALOGS ===============================================================================================================

'SCRIPT ====================================================================================================================
'Connects to BlueZone
EMConnect ""

'Grabbing the worker's X number.
CALL find_variable("User: ", worker_number, 7)
get_county_code

'Starting the query start time (for the query runtime at the end)
query_start_time = timer

'if user is not Hennepin County - the script will end. Process is not approved for other counties
if worker_county_code <> "x127" Then script_end_procedure("This script is built specifically for a process defined in the 'On Demand Waiver' for interviews. Currenly only Hennepin County has this waiver. This script should not be run for any cases outside of Hennepin County.")

'Script will determine if this is likely being run for the appointment notice or the NOMI
'Appointment notice will be run around the 18th of 2 months prior to the recert month
'NOMI will be run on the 15th of the month before the recert month

'Setting the initial path for the excel file to be found at - so we don't have to clickity click a bunch to get to the right file.
recertification_cases_excel_file_path = "T:\Eligibility Support\Restricted\QI - Quality Improvement\BZ scripts project\Projects\Auto Dialer\Auto-dialer data\03-18 renewals enhanced data.xlsx"

'Initial Dialog to determine the excel file to use, column with case numbers, and which process should be run
'Show initial dialog
Do
	Dialog recert_list_dlg
	If ButtonPressed = cancel then stopscript
	If ButtonPressed = select_a_file_button then call file_selection_system_dialog(recertification_cases_excel_file_path, ".xlsx")
Loop until ButtonPressed = OK and recertification_cases_excel_file_path <> "" and worker_signature <> ""

if worker_signature = "UUDDLRLRBA" then
	developer_mode = true
	MsgBox "You have enabled Developer Mode." & vbCr & vbCr & "The script will not enter information into MAXIS, but it will navigate, showing you where the script would otherwise have been."
END IF

'Opens Excel file here, as it needs to populate the dialog with the details from the spreadsheet.
call excel_open(recertification_cases_excel_file_path, True, True, ObjExcel, objWorkbook)

'Set objWorkSheet = objWorkbook.Worksheet
For Each objWorkSheet In objWorkbook.Worksheets
	If instr(objWorkSheet.Name, "Sheet") = 0 and objWorkSheet.Name <> "controls" then scenario_list = scenario_list & chr(9) & objWorkSheet.Name
Next

'Dialog to select worksheet
'DIALOG is defined here so that the dropdown can be populated with the above code
BeginDialog recert_worksheet_dlg, 0, 0, 151, 75, "On Demand Recertifications"
  DropListBox 5, 35, 140, 15, "Select One..." & scenario_list, scenario_dropdown
  ButtonGroup ButtonPressed
    OkButton 40, 55, 50, 15
    CancelButton 95, 55, 50, 15
  Text 5, 10, 130, 20, "Select the correct worksheet to run for recertification interview notifications:"
EndDialog

'Shows the dialog to select the correct worksheet
Do
    Dialog recert_worksheet_dlg
    If ButtonPressed = cancel then stopscript
Loop until scenario_dropdown <> "Select One..."

'Activates worksheet based on user selection
objExcel.worksheets(scenario_dropdown).Activate

'Creating an array of letters to loop through
col_ind = "A~B~C~D~E~F~G~H~I~J~K~L~M~N~O~P~Q~R~S~T~U~V~W~X~Y~Z"
col_array = split(col_ind, "~")
'setting the start of the list of column options
column_list = "Select One..."
cell_val = 1        'starting the value for reading the top cell of each column to use header information

'looping through the array
For each letter in col_array
    col_header = UCase(objExcel.Cells(1, cell_val).Value)
    col_header = trim(col_header)

    If col_header <> ""  then                                              'if the column is not blank - add to dropdown
        column_list = column_list & chr(9) & letter & " - " & col_header
        if col_header = "CASE NUMBER" then case_number_column = letter & " - " & col_header    'if the first cell says 'Case Number' then it is likely the correct column
    Else
        last_col = letter       'setting this for adding additional columns with information
        Exit For
    End If
    cell_val = cell_val + 1
Next

excel_row_to_start = "2"

'Next dialog determines the column the case numbers are in and the type of notification to be sent.
'Defining the dialog here so that the list of columns can be dynamically generated
BeginDialog recert_list_details_dlg, 0, 0, 266, 135, "On Demand Recertifications"
  DropListBox 160, 70, 100, 45, column_list, case_number_column
  DropListBox 160, 90, 100, 45, "Select One..."+chr(9)+"Appointment Notice"+chr(9)+"NOMI"+chr(9)+"Data Only", notice_type
  EditBox 75, 115, 40, 15, excel_row_to_start
  ButtonGroup ButtonPressed
    OkButton 155, 115, 50, 15
    CancelButton 210, 115, 50, 15
  Text 10, 10, 245, 20, "Check the Excel File that has been opened. Be sure it is the correct file to run at this time."
  Text 10, 35, 245, 30, "Choose the column that has all the case numbers listed and select which type of notice should be sent. The script will run very differently based on these answers."
  Text 10, 70, 145, 10, "Indicate the column with the case numbers:"
  Text 10, 95, 140, 10, "Which type of notice do you want to send?"
  Text 10, 120, 60, 10, "Excel row to start:"
EndDialog

'Displaying the dialog to select the correct column and type of notice.
Do
    Dialog recert_list_details_dlg
    If ButtonPressed = cancel then stopscript
Loop until case_number_column <> "Select One..." AND notice_type <> "Select One..."

'Determining important dates
'IF Appointment Notice - need the last day of recert, dates of interview options (beginning and end),
if notice_type = "Appointment Notice" then
    'creating a last day of recert variable - for appt notice this is the last day of the current month plus one - which is determined here
    last_day_of_recert = CM_plus_2_mo & "/01/" & CM_plus_2_yr
    last_day_of_recert = dateadd("D", -1, last_day_of_recert)

    'creating first interview date
    snap_interview_date_begin = CM_mo & "/25/" & CM_yr                     'SNAP interviews can be completed as early as the 25 of the month before the processing month
    mfip_interview_date_begin = CM_plus_1_mo & "/01/" & CM_plus_1_yr       'MFIP interviews cannot be completed until the report month is over - so must happen in the processing month

    'creating last interview date
    interview_end_date = CM_plus_1_mo & "/15/" & CM_plus_1_yr

    'MsgBox "Last day of recert - " & last_day_of_recert & vbNewLine & "SNAP interview begin - " & snap_interview_date_begin & vbNewLine & "MFIP interview begin - " & mfip_interview_date_begin & vbNewLine & "Last interview date - " & interview_end_date
end if


'IF NOMI - need last day of recert, interview deadline, response deadline
if notice_type = "NOMI" then
    'creating a last day of recert variable - for NOMI this is the last day of the current month - which is determined here
    last_day_of_recert = CM_plus_1_mo & "/01/" & CM_plus_1_yr
    last_day_of_recert = dateadd("D", -1, last_day_of_recert)

    'creating the interview deadline date - this was the last day provided in the appointment notice
    interview_deadline_date = CM_mo & "/15/" & CM_yr

    'creating the response deadline date - this is the day client must respond by in order to prevent case closure
    response_deadline_date = CM_plus_1_mo & "/01/" & CM_plus_1_yr
    response_deadline_date = dateadd("D", -11, last_day_of_recert)

    'MsgBox "Last day of recert - " & last_day_of_recert & vbNewLine & "Interview deadline date - " & interview_deadline_date & vbNewLine & "Response deadline date - " & response_deadline_date
end if

MAXIS_footer_month = CM_plus_1_mo
MAXIS_footer_year = CM_plus_1_yr

if notice_type = "Data Only" then

    'Creating an array of letters to loop through
    col_hdr = "A~B~C~D~E~F~G~H~I~J~K~L~M~N~O~P~Q~R~S~T~U~V~W~X~Y~Z"
    header_array = split(col_hdr, "~")
    'setting the start of the list of column options
    header_list = "Select One..."
    cell_val = 1        'starting the value for reading the top cell of each column to use header information

    'looping through the array
    For each letter in header_array
        col_header = UCase(objExcel.Cells(1, cell_val).Value)
        col_header = trim(col_header)

        If col_header <> ""  then                                              'if the column is not blank - add to dropdown
            header_list = header_list & chr(9) & letter & " - " & col_header
            if col_header = " NUMBER" then worker_col = letter & " - " & col_header    'if the first cell says 'Case Number' then it is likely the correct column
        Else
            Exit For
        End If
        cell_val = cell_val + 1
    Next

    interview_deadline_checkbox = checked
    interview_deadline = date & ""
    verify_appointment_notices_checkbox  = checked
    interview_frequency_checkbox = checked
    recvd_appl_frequency_checkbox = checked
    worker_interview_checkbox = checked
    worker_recert_status_checkbox = checked

    BeginDialog stats_dlg, 0, 0, 311, 110, "On Demand Recertifications"
      CheckBox 15, 30, 95, 10, "Interviews completed by ", interview_deadline_checkbox
      EditBox 110, 25, 50, 15, interview_deadline
      CheckBox 15, 45, 120, 10, "Cases with appointment notices", verify_appointment_notices_checkbox
      CheckBox 15, 60, 100, 10, "Interview Dates Frequency", interview_frequency_checkbox
      CheckBox 15, 75, 120, 10, "Application Received Frequency", recvd_appl_frequency_checkbox
      CheckBox 185, 30, 65, 10, "Interview Status", worker_interview_checkbox
      CheckBox 185, 45, 55, 10, "Recert Status", worker_recert_status_checkbox
      DropListBox 220, 65, 85, 45, header_list, worker_col
      EditBox 145, 90, 20, 15, MAXIS_footer_month
      EditBox 170, 90, 20, 15, MAXIS_footer_year
      ButtonGroup ButtonPressed
        OkButton 200, 90, 50, 15
        CancelButton 255, 90, 50, 15
      Text 10, 10, 170, 10, "Statistics to collect"
      GroupBox 175, 20, 90, 40, "Statistics by X-Number"
      Text 10, 95, 130, 10, "Footer Month of BENEFIT Month"
      Text 175, 70, 40, 10, "Worker Col"
    EndDialog


    Dialog stats_dlg
    If ButtonPressed = cancel then stopscript

    interview_deadline = DateAdd("d", 0, interview_deadline)
end if

check_for_MAXIS(false)

If developer_mode <> True and notice_type <> "Data Only" Then

    call back_to_self
    EMReadScreen mx_region, 10, 22, 48

    If mx_region = "INQUIRY DB" Then
        continue_in_inquiry = MsgBox("It appears you are attempting to have the script send notices for these cases." & vbNewLine & vbNewLine & "However, you appear to be in MAXIS Inquiry." &vbNewLine & "*************************" & vbNewLine & "Do you want to continue?", vbQuestion + vbYesNo, "Confirm Inquiry")
        If continue_in_inquiry = vbNo Then script_end_procedure("Live script run was attempted in Inquiry and aborted.")
    End If

    BeginDialog confirm_dialog, 0, 0, 196, 150, "Confirm Selections"
      Text 10, 10, 175, 20, "You are running a BULK script that will send notices. Review the Excel Spreadsheet that opened."
      Text 10, 35, 175, 10, "Worksheet selected: " & scenario_dropdown
      Text 10, 55, 175, 10, "Case Number Column: " & case_number_column
      Text 10, 75, 175, 10, "Notice to be sent: " & notice_type
      Text 10, 90, 180, 35, "This is a long running script and you will be unable to use any Excel document or the current session of MAXIS while the script runs. Review the selected options to be sure the script will takethe correct action."
      ButtonGroup ButtonPressed
        PushButton 80, 130, 50, 15, "Confirm", cnfrm_btn
        CancelButton 140, 130, 50, 15
    EndDialog

    Do
        Dialog confirm_dialog
        If ButtonPressed = cancel then stopscript
    Loop until buttonpressed = cnfrm_btn

End If

'Creating an array of all the cases to have a notice sent
dim ALL_CASES_ARRAY()
redim ALL_CASES_ARRAY(notc_confirm, 0)

'constants because words are easier to read than numbers when calling array elements
const case_number           = 0
const excel_row             = 1
const wrkr_x_numb           = 2
const priv_case             = 3
const out_of_co             = 4
const written_lang          = 5
const REVW_code             = 6
const SNAP_case             = 7
const MFIP_case             = 8
const recvd_appl            = 9
const date_of_app           = 10
const completed_interview   = 11
const date_of_interview     = 12
const appt_notc_sent        = 13
const nomi_sent             = 14
const notc_confirm          = 15

'setting variables for the loop
row = excel_row_to_start * 1
case_entry = 0
list_all_workers = ""

'converting the column letter to a number because cell values are called by number
col = left(case_number_column, 1)
call convert_excel_letter_to_excel_number(col)

if notice_type = "Data Only" Then
    wrkr_col = left(worker_col, 1)
    call convert_excel_letter_to_excel_number(worker_col)
end if

'reading each line of the Excel file and adding case number information to the array
do
    ReDim Preserve ALL_CASES_ARRAY(notc_confirm, case_entry)
    ALL_CASES_ARRAY(case_number, case_entry) = trim(objExcel.Cells(row, col).Value)
    ALL_CASES_ARRAY(excel_row, case_entry) = row

    if notice_type = "Data Only" then
        ALL_CASES_ARRAY(wrkr_x_numb, case_entry) = trim(objExcel.Cells(row, wrkr_col).Value)
        if InStr(list_all_workers, trim(objExcel.Cells(row, wrkr_col).Value)) = 0 then list_all_workers = list_all_workers & trim(objExcel.Cells(row, wrkr_col).Value) & "~"
    end if

    ALL_CASES_ARRAY(MFIP_case, case_entry) = FALSE
    ALL_CASES_ARRAY(SNAP_case, case_entry) = FALSE

    case_entry = case_entry + 1
    row = row + 1
    next_case_number = trim(objExcel.Cells(row, col).Value)
loop until next_case_number = ""
last_excel_row = row -1

total_cases = case_entry
if notice_type = "Data Only" then
    list_all_workers = left(list_all_workers, len(list_all_workers)-1)
    worker_array = split(list_all_workers, "~")
end if

running_total = 0
running_interviews_complete = 0
running_interviews_not_done = 0
running_no_app = 0
running_app_recvd = 0

'Looping through the array to gather all required information
for case_entry = 0 to UBound(ALL_CASES_ARRAY, 2)
    MAXIS_case_number = ALL_CASES_ARRAY(case_number, case_entry)        'setting this for using navigate functions

    CALL navigate_to_MAXIS_screen("CASE", "CURR")
    'Checking for PRIV cases.
	EMReadScreen priv_check, 6, 24, 14 'If it can't get into the case needs to skip
    EMReadScreen county_check, 2, 21, 16    'Looking to see if case has Hennepin COunty worker
	IF priv_check = "PRIVIL" THEN
        priv_case_list = priv_case_list & "|" & MAXIS_case_number
        ALL_CASES_ARRAY(priv_case, case_entry) = TRUE
    ELSE
        IF county_check <> "27" THEN ALL_CASES_ARRAY(out_of_co, case_entry) = "OUT OF COUNTY - " & county_check
        ALL_CASES_ARRAY(priv_case, case_entry) = FALSE
        running_total = running_total + 1
        'MEMB for written language
        Call navigate_to_MAXIS_screen("STAT", "MEMB")
        EMReadScreen language_code, 2, 13, 42
        ALL_CASES_ARRAY(written_lang, case_entry) = language_code

        'PROG to determine programs active
        Call navigate_to_MAXIS_screen("STAT", "PROG")
        EMReadScreen cash_prog_one, 2, 6, 67               'reading for active MFIP program - which has different requirements
        EMReadScreen cash_stat_one, 4, 6, 74
        EMReadScreen cash_prog_two, 2, 7, 67
        EMReadScreen cash_stat_two, 4, 7, 74

        'MFIP is defaulted to FALSE and will only be changed if PROG reads MFIP as active
        If cash_prog_one = "MF" AND cash_stat_one = "ACTV" then ALL_CASES_ARRAY(MFIP_case, case_entry) = TRUE
        If cash_prog_two = "MF" AND cash_stat_two = "ACTV" then ALL_CASES_ARRAY(MFIP_case, case_entry) = TRUE

        EMReadScreen snap_status, 4, 10, 74                'reading the status of SNAP

        'SNAP is defaulted to TRUE and will only be changed to FALSE if the status us not active or pending
        If snap_status = "ACTV" then ALL_CASES_ARRAY(SNAP_case, case_entry) = TRUE
        If snap_status = "PEND" then ALL_CASES_ARRAY(SNAP_case, case_entry) = TRUE

        Call HCRE_panel_bypass

        If notice_type = "NOMI" or notice_type = "Data Only" then
            'REVW to check the REVW code
            Call navigate_to_MAXIS_screen("STAT", "REVW")

            app_recvd = TRUE
            interview_complete = TRUE

            EMReadScreen recvd_date, 8, 13, 37
            recvd_date = replace(recvd_date, " ", "/")
            if recvd_date = "__/__/__" then
                app_recvd = FALSE
                running_no_app = running_no_app + 1
                recvd_date = ""
            else
                running_app_recvd = running_app_recvd + 1
            end if

            EMReadScreen interview_date, 8, 15, 37
            interview_date = replace(interview_date, " ", "/")
            if interview_date = "__/__/__" then
                interview_complete = FALSE
                interview_date = ""
                running_interviews_not_done = running_interviews_not_done + 1
            else
                running_interviews_complete = running_interviews_complete + 1
            end if

            ALL_CASES_ARRAY(recvd_appl, case_entry) = app_recvd
            ALL_CASES_ARRAY(date_of_app, case_entry) = recvd_date
            ALL_CASES_ARRAY(completed_interview, case_entry) = interview_complete
            ALL_CASES_ARRAY(date_of_interview, case_entry) = interview_date

            EMReadScreen review_status, 1, 7, 60
            if review_status = "_" then EMReadScreen review_status, 1, 7, 40

            ALL_CASES_ARRAY(REVW_code, case_entry) = review_status

            'Going to check to see if an appointment notice was sent
            Call navigate_to_MAXIS_screen("CASE", "NOTE")

            ALL_CASES_ARRAY(appt_notc_sent, case_entry) = FALSE

            the_date_to_look_at = MAXIS_footer_month & "/1/" & MAXIS_footer_year
            too_old_date = DateAdd("M", -2, the_date_to_look_at)
            'MsgBox too_old_date
            note_row = 5
            ' MsgBox "End of case note look - " & end_look_mo & "/" & end_look_yr
            Do
                EMReadScreen note_date, 8, note_row, 6

                EMReadScreen note_title, 55, note_row, 25
                note_title = trim(note_title)
                'MsgBox note_title
                if note_title = "*** Notice of SNAP Recertification Interview Sent ***" then ALL_CASES_ARRAY(appt_notc_sent, case_entry) = TRUE
                if note_title = "*** Notice of MFIP Recertification Interview Sent ***" then ALL_CASES_ARRAY(appt_notc_sent, case_entry) = TRUE
                if note_title = "*** Notice of MFIP/SNAP Recertification Interview Sent" then ALL_CASES_ARRAY(appt_notc_sent, case_entry) = TRUE

                if note_mo = "  " then Exit Do
                if ALL_CASES_ARRAY(appt_notc_sent, case_entry) = TRUE then Exit Do

                note_row = note_row + 1
                if note_row = 19 then
                    'MsgBox "Next Page" & vbNewLine & "Note Date:" & note_date
                    note_row = 5
                    PF8
                    EMReadScreen check_for_last_page, 9, 24, 14
                    If check_for_last_page = "LAST PAGE" Then Exit Do
                End If
                EMReadScreen next_note_date, 8, note_row, 6
            Loop until DateDiff("d", too_old_date, next_note_date) <= 0
            'MsgBox ALL_CASES_ARRAY(appt_notc_sent, case_entry)
        end if

        if notice_type = "Data Only" then
            Call navigate_to_MAXIS_screen("CASE", "NOTE")

            ALL_CASES_ARRAY(appt_notc_sent, case_entry) = FALSE
            ALL_CASES_ARRAY(nomi_sent, case_entry) = FALSE

            the_date_to_look_at = MAXIS_footer_month & "/1/" & MAXIS_footer_year
            too_old_date = DateAdd("d", -1, (DateAdd("M", -2, the_date_to_look_at)))
            '
            note_row = 5
            ' MsgBox "End of case note look - " & end_look_mo & "/" & end_look_yr
            Do
                EMReadScreen note_date, 8, note_row, 6

                EMReadScreen note_title, 55, note_row, 25
                note_title = trim(note_title)
                'MsgBox note_title
                if note_title = "*** Notice of SNAP Recertification Interview Sent ***" then ALL_CASES_ARRAY(appt_notc_sent, case_entry) = TRUE
                if note_title = "*** Notice of MFIP Recertification Interview Sent ***" then ALL_CASES_ARRAY(appt_notc_sent, case_entry) = TRUE
                if note_title = "*** Notice of MFIP/SNAP Recertification Interview Sent" then ALL_CASES_ARRAY(appt_notc_sent, case_entry) = TRUE
                if note_title = "*** NOMI Sent for SNAP Reertification***" then ALL_CASES_ARRAY(nomi_sent, case_entry) = TRUE
                if note_title = "*** NOMI Sent for SNAP Recertification***" then ALL_CASES_ARRAY(nomi_sent, case_entry) = TRUE

                if note_mo = "  " then Exit Do
                if ALL_CASES_ARRAY(appt_notc_sent, case_entry) = TRUE then Exit Do

                note_row = note_row + 1
                if note_row = 19 then
                    'MsgBox "Next Page" & vbNewLine & "Note Date:" & note_date
                    note_row = 5
                    PF8
                    EMReadScreen check_for_last_page, 9, 24, 14
                    If check_for_last_page = "LAST PAGE" Then Exit Do
                End If
                EMReadScreen next_note_date, 8, note_row, 6
            Loop until DateDiff("d", too_old_date, next_note_date) <= 0
        end if
    END IF
    call back_to_SELF
next
'All information has been gathered


'Insert columns in excel for additional information to be added
column_end = last_col & "1"

Set objRange = objExcel.Range(column_end).EntireColumn
objRange.Insert(xlShiftToRight)                             'inserting one column to the end of the data in the spreadsheet

notc_col = last_col                                         'setting the a variable with the notice column for later updating of excel
notc_letter_col = notc_col
call convert_excel_letter_to_excel_number(notc_col)

if notice_type = "Appointment Notice" Then
    objExcel.Cells(1, notc_col).Value = "Appt Notice Success"   'Adding header to Excel

    stats_header_col = notc_col + 2         'Setting variables with coumn locations for statistics
    stats_col = notc_col + 3
End If

If notice_type = "NOMI" or notice_type = "Data Only" Then
    objRange.Insert(xlShiftToRight)     'add column for review status
    objRange.Insert(xlShiftToRight)     'add column with interview date
    objRange.Insert(xlShiftToRight)     'add column with app date
    If notice_type = "Data Only" Then objRange.Insert(xlShiftToRight)     'add another column for the other notice confirmation

    If notice_type = "NOMI" Then
        revw_code_col = notc_col + 1
        nomi_letter_col = convert_digit_to_excel_column(notc_col)
    ElseIf notice_type = "Data Only" Then
        appt_lrt_col = notc_col
        nomi_col = notc_col + 1
        revw_code_col = nomi_col + 1        'setting variables for writing to excel'

        appt_ltr_letter_col = convert_digit_to_excel_column(appt_lrt_col)
        nomi_letter_col = convert_digit_to_excel_column(nomi_col)
    End If
    intvw_date_col = revw_code_col + 1
    app_date_col = intvw_date_col + 1

    revw_code_letter_col = convert_digit_to_excel_column(revw_code_col)
    intvw_date_letter_col = convert_digit_to_excel_column(intvw_date_col)
    app_date_letter_col = convert_digit_to_excel_column(app_date_col)

    If notice_type = "NOMI" Then objExcel.Cells(1, notc_col).Value = "NOMI Success"          'adding headers to Excel
    If notice_type = "Data Only" Then
        objExcel.Cells(1, appt_lrt_col).Value = "Appt LTR Confirm"
        objExcel.Cells(1, nomi_col).Value = "NOMI Confirm"
    End If
    objExcel.Cells(1, revw_code_col).Value = "REVW Status"
    objExcel.Cells(1, intvw_date_col).Value = "Interview Date"
    objExcel.Cells(1, app_date_col).Value = "Date App Rec'vd"

    stats_header_col = app_date_col + 2    'Setting variables with coumn locations for statistics
    stats_col = app_date_col + 3
End If

'creating a variable in the MM/DD/YY format to compare with date read from MAXIS
today_mo = DatePart("m", date)
today_mo = right("00" & today_mo, 2)

today_day = DatePart("d", date)
today_day = right("00" & today_day, 2)

today_yr = DatePart("yyyy", date)
today_yr = right(today_yr, 2)

today_date = today_mo & "/" & today_day & "/" & today_yr

interviews_completed_timely = 0     'setting this here to the spreadsheet will read 0 if none found.


'Looping through the array again to send the notice and confirm the notice will be sent (that it saved correctly)
'For NOMI option, this will only send a notice IF the interview is incomplete.
for case_entry = 0 to UBound(ALL_CASES_ARRAY, 2)
    MAXIS_case_number = ALL_CASES_ARRAY(case_number, case_entry)        'setting this for using navigate functions
    row = ALL_CASES_ARRAY(excel_row, case_entry)

    forms_to_swkr = ""
    forms_to_arep = ""

    memo_started = True

    if ALL_CASES_ARRAY(priv_case, case_entry) = FALSE and ALL_CASES_ARRAY(out_of_co, case_entry) = "" then                  'PRIV cases will not have a MEMO attempted
        If notice_type = "Appointment Notice" Then                          'For cases requiring an Appointment Notice'
            STATS_counter = STATS_counter + 1
            if ALL_CASES_ARRAY(MFIP_case, case_entry) = TRUE then           'setting the language for the notices - MFIP or SNAP
                if ALL_CASES_ARRAY(SNAP_case, case_entry) = TRUE then
                    programs = "MFIP/SNAP"
                else
                    programs = "MFIP"
                end if
            else
                programs = "SNAP"
            end if

            if developer_mode = true then           'Running the script in developer mode will display a message box with the wording of the MEMO and CASENOTE
                Call navigate_to_MAXIS_screen ("SPEC", "MEMO")
                Select Case ALL_CASES_ARRAY(written_lang, case_entry)

                Case "02"   'Hmong
                    MsgBox "HMONG"
                Case "21"   'Karen
                    MsgBox "KAREN"
                Case "06"   'Russian
                    MsgBox "RUSSIAN"
                Case "07"   'Somali
                    MsgBox "SOMALI"
                Case "01"   'Spanish
                    MsgBox "SPANISH"
                Case Else
                    MsgBox "ENGLISH"
                End Select

                Memo_to_display = "The Department of Human Services sent you a packet of paperwork. This paperwork is to renew your " & programs & " case. Your " &_
                                                programs & " case will close on " & last_day_of_recert &_
                                                " if we do not receive your paperwork. Please sign, date and return your renewal paperwork by " &_
                                                CM_plus_1_mo & "/08/" & CM_plus_1_yr & "."
                Memo_to_display = Memo_to_display & vbNewLine & vbNewLine & "You must also complete an interview for your " & programs &_
                    " case to continue. To complete a phone interview, call the EZ Info Line at 612-596-1300 between 9:00am and 4:00pm Monday through Friday. Please complete your interview by " & interview_end_date & "."
                Memo_to_display = Memo_to_display & vbNewLine & vbNewLine & "We must have your renewal paperwork to do your interview. Please send proofs with your renewal paperwork."
                Memo_to_display = Memo_to_display & vbNewLine & vbNewLine & " * Examples of income proofs: paystubs, income reports, business ledgers, income tax forms, etc."
                Memo_to_display = Memo_to_display & vbNewLine & vbNewLine & " * Examples of housing cost proofs(if changed): rent/house payment receipt, mortgage, lease, subsidy, etc."
                Memo_to_display = Memo_to_display & vbNewLine & vbNewLine & " * Examples of medical cost proofs(if changed): prescription and medical bills, etc."
                Memo_to_display = Memo_to_display & vbNewLine & vbNewLine & "Domestic violence brochures are available at this website: https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3477-ENG. You can also request a paper copy."

                MsgBox Memo_to_display

                Call navigate_to_MAXIS_screen ("CASE", "NOTE")

                Case_note_to_display = "*** Notice of " & programs & " Recertification Interview Sent ***"
                Case_note_to_display = Case_note_to_display & vbNewLine & vbNewLine & "* A notice has been sent to client with detail about how to call in for an interview."
                Case_note_to_display = Case_note_to_display & vbNewLine & vbNewLine & "* Client must submit paperwork and call 612-596-1300 to complete interview."
                If forms_to_arep = "Y" then Case_note_to_display = Case_note_to_display & vbNewLine & vbNewLine & "* Copy of notice sent to AREP."
                If forms_to_swkr = "Y" then Case_note_to_display = Case_note_to_display & vbNewLine & vbNewLine & "* Copy of notice sent to Social Worker."
                Case_note_to_display = Case_note_to_display & vbNewLine & vbNewLine & "---"
                Case_note_to_display = Case_note_to_display & vbNewLine & vbNewLine & "Link to Domestic Violence Brochure sent to client in SPEC/MEMO as a part of interview notice."
                Case_note_to_display = Case_note_to_display & vbNewLine & vbNewLine & "---"
                Case_note_to_display = Case_note_to_display & vbNewLine & vbNewLine & "Worker Name"     'hard coded because when developer mode is run the worker signature is UUDDLRLRBA

                MsgBox Case_note_to_display

                successful_notices = successful_notices + 1                 'For statistical purposes
                ALL_CASES_ARRAY(notc_confirm, case_entry) = "Y"         'setting this here because in dev mode - the notice is always successful

            else


                'Looking for the written language code.
                'Once we have the memo translated into other languages, the MEMO portion will be put here and will be specific to the language needs.
                Select Case ALL_CASES_ARRAY(written_lang, case_entry)

                Case "02"   'Hmong
                    'MsgBox "HMONG"
                Case "21"   'Karen
                    'MsgBox "KAREN"
                Case "06"   'Russian
                    'MsgBox "RUSSIAN"
                Case "07"   'Somali
                    'MsgBox "SOMALI"
                Case "01"   'Spanish
                    'MsgBox "SPANISH"
                Case Else
                    'MsgBox "ENGLISH"
                End Select

                'Writing the SPEC MEMO - dates will be input from the determination made earlier.
                Call start_a_new_spec_memo_and_continue(memo_started)

                IF memo_started = True THEN         'The function will return this as FALSE if PF5 does not move past MEMO DISPLAY

                    EMSendKey("************************************************************")           'for some reason this is more stable then using write_variable
                    CALL write_variable_in_SPEC_MEMO("The Department of Human Services sent you a packet of paperwork. This paperwork is to renew your " & programs & " case. Your " &_
                                                    programs & " case will close on " & last_day_of_recert &_
                                                    " if we do not receive your paperwork. Please sign, date and return your renewal paperwork by " &_
                                                    CM_plus_1_mo & "/08/" & CM_plus_1_yr & ".")
                    CALL write_variable_in_SPEC_MEMO("")
                    CALL write_variable_in_SPEC_MEMO("You must also complete an interview for your " & programs &_
                        " case to continue. To complete a phone interview, call the EZ Info Line at 612-596-1300 between 9:00am and 4:00pm Monday through Friday. Please complete your interview by " & interview_end_date & ".")
                    CALL write_variable_in_SPEC_MEMO("")
                    CALL write_variable_in_SPEC_MEMO("We must have your renewal paperwork to do your interview. Please send proofs with your renewal paperwork.")
            		CALL write_variable_in_SPEC_MEMO("")
            		CALL write_variable_in_SPEC_MEMO(" * Examples of income proofs: paystubs, income reports, business ledgers, income tax forms, etc.")
            		CALL write_variable_in_SPEC_MEMO("")
            		CALL write_variable_in_SPEC_MEMO(" * Examples of housing cost proofs(if changed): rent/house payment receipt, mortgage, lease, subsidy, etc.")
            		CALL write_variable_in_SPEC_MEMO("")
            		CALL write_variable_in_SPEC_MEMO(" * Examples of medical cost proofs(if changed): prescription and medical bills, etc.")
                    CALL write_variable_in_SPEC_MEMO("")
            		CALL write_variable_in_SPEC_MEMO("Domestic violence brochures are available at this website: https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3477-ENG. You can also request a paper copy.")


                    PF4         'Submit the MEMO

                    memo_row = 7                                            'Setting the row for the loop to read MEMOs
                    ALL_CASES_ARRAY(notc_confirm, case_entry) = "N"         'Defaulting this to 'N'
                    Do
                        EMReadScreen create_date, 8, memo_row, 19                 'Reading the date of each memo and the status
                        EMReadScreen print_status, 7, memo_row, 67
                        If create_date = today_date AND print_status = "Waiting" Then   'MEMOs created today and still waiting is likely our MEMO.
                            ALL_CASES_ARRAY(notc_confirm, case_entry) = "Y"             'If we've found this then no reason to keep looking.
                            successful_notices = successful_notices + 1                 'For statistical purposes
                            Exit Do
                        End If

                        memo_row = memo_row + 1           'Looking at next row'
                    Loop Until create_date = "        "

                ELSE
                    ALL_CASES_ARRAY(notc_confirm, case_entry) = "N"         'Setting this as N if the MEMO failed
                    call back_to_SELF
                END IF

                if ALL_CASES_ARRAY(notc_confirm, case_entry) = "Y" then         'IF the notice was confirmed a CASENOTE will be entered
                    start_a_blank_case_note
                    EMSendKey("*** Notice of " & programs & " Recertification Interview Sent ***")
                    CALL write_variable_in_case_note("* A notice has been sent to client with detail about how to call in for an interview.")
                    CALL write_variable_in_case_note("* Client must submit paperwork and call 612-596-1300 to complete interview.")
                    If forms_to_arep = "Y" then call write_variable_in_case_note("* Copy of notice sent to AREP.")
                    If forms_to_swkr = "Y" then call write_variable_in_case_note("* Copy of notice sent to Social Worker.")
                    call write_variable_in_case_note("---")
                    CALL write_variable_in_case_note("Link to Domestic Violence Brochure sent to client in SPEC/MEMO as a part of interview notice.")
                    call write_variable_in_case_note("---")
                    call write_variable_in_case_note(worker_signature)
                end if
            end if
            'Entering the code from notice confirmation to excel
            row = ALL_CASES_ARRAY(excel_row, case_entry)
            objExcel.Cells(row, notc_col).Value = ALL_CASES_ARRAY(notc_confirm, case_entry)

        End If

        If notice_type = "NOMI" then                                        'different notice and note for the NOMI option
            If ALL_CASES_ARRAY(completed_interview, case_entry) = FALSE Then    'only the cases that have not had an interview need a NOMI.
                STATS_counter = STATS_counter + 1
                incomplete_reviews = incomplete_reviews + 1                     'this is for statistical gathering

                If ALL_CASES_ARRAY(appt_notc_sent, case_entry) = TRUE Then
                    if developer_mode = true then                                   'msgbox to display a memo instead of writing an actual memo'
                        Call navigate_to_MAXIS_screen ("SPEC", "MEMO")
                        Select Case ALL_CASES_ARRAY(written_lang, case_entry)       'Sending notice by language if possible

                        Case "02"   'Hmong
                            MsgBox "HMONG"
                        Case "21"   'Karen
                            MsgBox "KAREN"
                        Case "06"   'Russian
                            MsgBox "RUSSIAN"
                        Case "07"   'Somali
                            MsgBox "SOMALI"
                        Case "01"   'Spanish
                            MsgBox "SPANISH"
                        Case Else
                            MsgBox "ENGLISH"
                        End Select

                        'creating the memo message and displaying it.
                        if ALL_CASES_ARRAY(recvd_appl, case_entry) = TRUE then Memo_to_display = "We received your Recertification on " & ALL_CASES_ARRAY(date_of_app, case_entry) & "."
                        if ALL_CASES_ARRAY(recvd_appl, case_entry) = FALSE then Memo_to_display = "Your Recertification has not yet been received."

                        Memo_to_display = Memo_to_display & vbNewLine & vbNewLine & "You must have an interview by " & last_day_of_recert & " or your benefits will end. "
                        Memo_to_display = Memo_to_display & vbNewLine & vbNewLine & "To complete a phone interview, call the EZ Info Line at 612-596-1300 between 9:00am and 4:00pm Monday through Friday."
                        Memo_to_display = Memo_to_display & vbNewLine & vbNewLine & "You may also come in to the office to complete an interview between 8:00 am and 4:30pm Monday through Friday."
                        Memo_to_display = Memo_to_display & vbNewLine & vbNewLine & "If we do not hear from you by " & last_day_of_recert & ", your benefits will end on " & last_day_of_recert & "."

                        MsgBox Memo_to_display

                        Call navigate_to_MAXIS_screen ("CASE", "NOTE")
                        'creating the case note message and displaying it
                        Case_note_to_display = "*** NOMI Sent for SNAP Recertification***"
                        if ALL_CASES_ARRAY(recvd_appl, case_entry) = TRUE then Case_note_to_display = Case_note_to_display & vbNewLine & vbNewLine & "* Recertification app received on " & ALL_CASES_ARRAY(date_of_app, case_entry)
                        if ALL_CASES_ARRAY(recvd_appl, case_entry) = FALSE then Case_note_to_display = Case_note_to_display & vbNewLine & vbNewLine & "* Recertification app has NOT been received. Client must submit paperwork."
                        Case_note_to_display = Case_note_to_display & vbNewLine & vbNewLine & "* A notice was previously sent to client with detail about how to call in for an interview."
                        Case_note_to_display = Case_note_to_display & vbNewLine & vbNewLine & "* Client must call 612-596-1300 to complete interview."
                        If forms_to_arep = "Y" then Case_note_to_display = Case_note_to_display & vbNewLine & vbNewLine & "* Copy of notice sent to AREP."
                        If forms_to_swkr = "Y" then Case_note_to_display = Case_note_to_display & vbNewLine & vbNewLine & "* Copy of notice sent to Social Worker."
                        Case_note_to_display = Case_note_to_display & vbNewLine & vbNewLine & "---"
                        Case_note_to_display = Case_note_to_display & vbNewLine & vbNewLine & "Worker Name"

                        MsgBox Case_note_to_display

                        successful_notices = successful_notices + 1                 'For statistical purposes
                        ALL_CASES_ARRAY(notc_confirm, case_entry) = "Y"         'setting this here because in dev mode - the notice is always successful
                    else

                        Select Case ALL_CASES_ARRAY(written_lang, case_entry)       'selecting  the language and writing the memo by language

                        Case "02"   'Hmong
                            'MsgBox "HMONG"
                        Case "21"   'Karen
                            'MsgBox "KAREN"
                        Case "06"   'Russian
                            'MsgBox "RUSSIAN"
                        Case "07"   'Somali
                            'MsgBox "SOMALI"
                        Case "01"   'Spanish
                            'MsgBox "SPANISH"
                        Case Else
                            'MsgBox "ENGLISH"
                        End Select

                        'writing a SPEC MEMO with the NOMI wording.
                        Call start_a_new_spec_memo_and_continue(memo_started)
                        'MsgBox memo_started
                        IF memo_started = True THEN         'The function will return this as FALSE if PF5 does not move past MEMO DISPLAY

                            if ALL_CASES_ARRAY(recvd_appl, case_entry) = TRUE then CALL write_variable_in_SPEC_MEMO("We received your Recertification on " & ALL_CASES_ARRAY(date_of_app, case_entry) & ".")
                            if ALL_CASES_ARRAY(recvd_appl, case_entry) = FALSE then CALL write_variable_in_SPEC_MEMO("Your Recertification has not yet been received.")
                            CALL write_variable_in_SPEC_MEMO("")
                            CALL write_variable_in_SPEC_MEMO("You must have an interview by " & last_day_of_recert & " or your benefits will end. ")
                            CALL write_variable_in_SPEC_MEMO("")
                            CALL write_variable_in_SPEC_MEMO("To complete a phone interview, call the EZ Info Line at 612-596-1300 between 9:00am and 4:00pm Monday through Friday.")
                            CALL write_variable_in_SPEC_MEMO("")
                            CALL write_variable_in_SPEC_MEMO("You may also come in to the office to complete an interview between 8:00 am and 4:30pm Monday through Friday.")
                            CALL write_variable_in_SPEC_MEMO("")
                            CALL write_variable_in_SPEC_MEMO("If we do not hear from you by " & last_day_of_recert & ", your benefits will end on " & last_day_of_recert & ".")

                            PF4         'Submit the MEMO

                            memo_row = 7                                            'Setting the row for the loop to read MEMOs
                            ALL_CASES_ARRAY(notc_confirm, case_entry) = "N"         'Defaulting this to 'N'
                            Do
                                EMReadScreen create_date, 8, memo_row, 19                 'Reading the date of each memo and the status
                                EMReadScreen print_status, 7, memo_row, 67
                                If create_date = today_date AND print_status = "Waiting" Then   'MEMOs created today and still waiting is likely our MEMO.
                                    ALL_CASES_ARRAY(notc_confirm, case_entry) = "Y"             'If we've found this then no reason to keep looking.
                                    successful_notices = successful_notices + 1                 'For statistical purposes
                                    Exit Do
                                End If

                                memo_row = memo_row + 1           'Looking at next row'
                            Loop Until create_date = "        "

                        ELSE
                            ALL_CASES_ARRAY(notc_confirm, case_entry) = "N"         'Setting this as N if the MEMO failed
                            call back_to_SELF
                        END IF

                        'writing a case note only if the notice was successfully sent.
                        if ALL_CASES_ARRAY(notc_confirm, case_entry) = "Y" then

                            start_a_blank_case_note

                            EMSendKey("*** NOMI Sent for SNAP Recertification***")
                            if ALL_CASES_ARRAY(recvd_appl, case_entry) = TRUE then CALL write_variable_in_CASE_NOTE("* Recertification app received on " & ALL_CASES_ARRAY(date_of_app, case_entry))
                            if ALL_CASES_ARRAY(recvd_appl, case_entry) = FALSE then CALL write_variable_in_CASE_NOTE("* Recertification app has NOT been received. Client must submit paperwork.")
                            CALL write_variable_in_CASE_NOTE("* A notice was previously sent to client with detail about how to call in for an interview.")
                            CALL write_variable_in_CASE_NOTE("* Client must call 612-596-1300 to complete interview.")
                            If forms_to_arep = "Y" then CALL write_variable_in_CASE_NOTE("* Copy of notice sent to AREP.")
                            If forms_to_swkr = "Y" then CALL write_variable_in_CASE_NOTE("* Copy of notice sent to Social Worker.")
                            call write_variable_in_case_note("---")
                            call write_variable_in_case_note(worker_signature)
                        end if
                    end if
                Else
                    ALL_CASES_ARRAY(notc_confirm, case_entry) = "APPT NOTICE FAILED"
                End If
                'submit MEMO and check for status.
                'If MEMO sent set ALL_CASES_ARRAY(notc_confirm, case_entry) = "Y" (else "N")

            End If
            'MsgBox ALL_CASES_ARRAY(notc_confirm, case_entry)
            'Writing important details to excel
            row = ALL_CASES_ARRAY(excel_row, case_entry)
            objExcel.Cells(row, notc_col).Value = ALL_CASES_ARRAY(notc_confirm, case_entry)              'indicator if notice was successful or not
            objExcel.Cells(row, intvw_date_col).Value = ALL_CASES_ARRAY(date_of_interview, case_entry)   'date of interview
            objExcel.Cells(row, app_date_col).Value = ALL_CASES_ARRAY(date_of_app, case_entry)           'date of application
            objExcel.Cells(row, revw_code_col).Value = ALL_CASES_ARRAY(REVW_code, case_entry)            'Code from REVW for status of RECERT
        End If
        if notice_type = "Data Only" then

            ' if ALL_CASES_ARRAY(date_of_interview, case_entry) <> "" then
            '     'to_date = FormatDateTime(ALL_CASES_ARRAY(date_of_interview, case_entry))
            '     'If DateDiff("d", interview_deadline, to_date) >= 0 then interviews_completed_timely = interviews_completed_timely + 1
            '     'MsgBox DateDiff("d", interview_deadline, ALL_CASES_ARRAY(date_of_interview, case_entry))
            '     If DateDiff("d", ALL_CASES_ARRAY(date_of_interview, case_entry), interview_deadline) >= 0 then interviews_completed_timely = interviews_completed_timely + 1
            ' end if

            if ALL_CASES_ARRAY(appt_notc_sent, case_entry) = TRUE then objExcel.Cells(row, appt_lrt_col).Value = "Y"
            if ALL_CASES_ARRAY(nomi_sent, case_entry) = TRUE then objExcel.Cells(row, nomi_col).Value = "Y"

            objExcel.Cells(row, intvw_date_col).Value = ALL_CASES_ARRAY(date_of_interview, case_entry)   'date of interview
            objExcel.Cells(row, app_date_col).Value = ALL_CASES_ARRAY(date_of_app, case_entry)           'date of application
            objExcel.Cells(row, revw_code_col).Value = ALL_CASES_ARRAY(REVW_code, case_entry)            'Code from REVW for status of RECERT
        end if
    Else
        If ALL_CASES_ARRAY(priv_case, case_entry) = True Then
            objExcel.Cells(row, notc_col).Value = "PRIV"
        Else
            objExcel.Cells(row, notc_col).Value = ALL_CASES_ARRAY(out_of_co, case_entry)

            if notice_type = "Data Only" or notice_type = "NOMI" then
                objExcel.Cells(row, intvw_date_col).Value = ALL_CASES_ARRAY(date_of_interview, case_entry)   'date of interview
                objExcel.Cells(row, app_date_col).Value = ALL_CASES_ARRAY(date_of_app, case_entry)           'date of application
                objExcel.Cells(row, revw_code_col).Value = ALL_CASES_ARRAY(REVW_code, case_entry)            'Code from REVW for status of RECERT
            end if
        End If
    End If
next

'Statistics will be created on number of cases and percentage of notices sent/interviews completed/NOMIs sent
'Variables for stats - successful_notices, incomplete_reviews, total_cases'
sheet_row = 1
is_not_blank_excel_string = chr(34) & "<>" & chr(34)
Do
    stats_header_cell = objExcel.Cells(sheet_row, stats_header_col).Value   'Readding the cell where stats headers go
    stats_value_cell = objExcel.Cells(sheet_row, stats_col).Value           'Reading the cell where stats values go
    stats_letter_col = convert_digit_to_excel_column(stats_col)

    stats_header_cell = trim(stats_header_cell)
    stats_value_cell = trim(stats_value_cell)

    if stats_header_cell = "" and stats_value_cell = "" then        'If both are blank we have reached the end of the list of stats info
        entry_row = sheet_row                                       'Setting to start the next set of stats here
    else
        sheet_row = sheet_row + 1
    end if
Loop until stats_header_cell = "" and stats_value_cell = ""

objExcel.Cells(entry_row, stats_header_col).Value       = "--------------------------------------"      'Entering a seperator since multiple stats information are listed on the same sheet
entry_row = entry_row + 1

'For when this is run on appointment notices
if notice_type = "Appointment Notice" then
    objExcel.Cells(entry_row, stats_header_col).Value       = "Appointment Notices run on:"     'Date and time the script was completed
    objExcel.Cells(entry_row, stats_header_col).Font.Bold 	= TRUE
    objExcel.Cells(entry_row, stats_col).Value              = now
    entry_row = entry_row + 1

    objExcel.Cells(entry_row, stats_header_col).Value       = "Runtime (in seconds)"            'Enters the amount of time it took the script to run
    objExcel.Cells(entry_row, stats_header_col).Font.Bold 	= TRUE
    objExcel.Cells(entry_row, stats_col).Value              = timer - query_start_time
    entry_row = entry_row + 1

    objExcel.Cells(entry_row, stats_header_col).Value       = "Total Cases assesed"             'All cases from the spreadsheet
    objExcel.Cells(entry_row, stats_header_col).Font.Bold 	= TRUE
    objExcel.Cells(entry_row, stats_col).Value              = total_cases
    total_row = entry_row
    entry_row = entry_row + 1

    if successful_notices = "" then successful_notices = 0
    objExcel.Cells(entry_row, stats_header_col).Value       = "Appointment Notices Sent"        'number of notices that were successful
    objExcel.Cells(entry_row, stats_header_col).Font.Bold 	= TRUE
    objExcel.Cells(entry_row, stats_col).Value              = "=COUNTIF(" & notc_letter_col & ":" & notc_letter_col & ", " & Chr(34) & "Y" & Chr(34) & ")"                'This was incremented on the For Next loop where the memos were written
    appt_row = entry_row
    entry_row = entry_row + 1

    objExcel.Cells(entry_row, stats_header_col).Value       = "Percentage successful"           'calculation of the percent of successful notices
    objExcel.Cells(entry_row, stats_header_col).Font.Bold 	= TRUE
    objExcel.Cells(entry_row, stats_col).Value              = "=" & stats_letter_col & appt_row & "/" & stats_letter_col & total_row
    objExcel.Cells(entry_row, stats_col).NumberFormat       = "0.00%"		'Formula should be percent
    entry_row = entry_row + 1

    objExcel.Cells(entry_row, stats_header_col).Value       = "Privleged Cases:"                'Enter the header of the priv cases
    objExcel.Cells(entry_row, stats_header_col).Font.Bold 	= TRUE
end if

'For if the script was run for sending NOMIs
if notice_type = "NOMI" then
    objExcel.Cells(entry_row, stats_header_col).Value       = "NOMIs run on:"           'Date and time script run was completed
    objExcel.Cells(entry_row, stats_header_col).Font.Bold 	= TRUE
    objExcel.Cells(entry_row, stats_col).Value              = now
    entry_row = entry_row + 1

    objExcel.Cells(entry_row, stats_header_col).Value       = "Runtime (in seconds)"    'amount of time it took for the script to run'
    objExcel.Cells(entry_row, stats_header_col).Font.Bold 	= TRUE
    objExcel.Cells(entry_row, stats_col).Value              = timer - query_start_time
    entry_row = entry_row + 1

    objExcel.Cells(entry_row, stats_header_col).Value       = "Cases with no Interview"     'number of cases that potentially need a NOMI'
    objExcel.Cells(entry_row, stats_header_col).Font.Bold 	= TRUE
    objExcel.Cells(entry_row, stats_col).Value              = "=COUNTBLANK(" & intvw_date_letter_col & "2:" & intvw_date_letter_col & last_excel_row & ")"
    no_intv_row = entry_row
    entry_row = entry_row + 1

    if successful_notices = "" then successful_notices = 0
    objExcel.Cells(entry_row, stats_header_col).Value       = "NOMIs Sent"              'Number of successful NOMIs sent'
    objExcel.Cells(entry_row, stats_header_col).Font.Bold 	= TRUE
    objExcel.Cells(entry_row, stats_col).Value              = "=COUNTIF(" & nomi_letter_col & ":" & nomi_letter_col & ", " & Chr(34) & "Y" & Chr(34) & ")"        'This is incremented in the For Next loop above'
    nomi_row = entry_row
    entry_row = entry_row + 1

    objExcel.Cells(entry_row, stats_header_col).Value       = "Percentage successful"   'Calculates the percentage of NOMIs siucessful (from attempted)'
    objExcel.Cells(entry_row, stats_header_col).Font.Bold 	= TRUE
    objExcel.Cells(entry_row, stats_col).Value              = "=" & stats_letter_col & nomi_row & "/" & stats_letter_col & no_intv_row
    objExcel.Cells(entry_row, stats_col).NumberFormat       = "0.00%"		'Formula should be percent
    entry_row = entry_row + 1

    objExcel.Cells(entry_row, stats_header_col).Value       = "Interviews Completed"   'Calculates the percentage of NOMIs siucessful (from attempted)'
    objExcel.Cells(entry_row, stats_header_col).Font.Bold 	= TRUE
    objExcel.Cells(entry_row, stats_col).Value              = "=COUNTIF(" & intvw_date_letter_col & "2:" & intvw_date_letter_col & total_cases + 1 & ", " & is_not_blank_excel_string & ")"
    entry_row = entry_row + 1

    objExcel.Cells(entry_row, stats_header_col).Value       = "Privleged Cases:"        'PRIV cases header'
    objExcel.Cells(entry_row, stats_header_col).Font.Bold 	= TRUE
end if

'IF the Data Only option was selected
if notice_type = "Data Only" then
    row_to_use = 2      'Setting row and column because we are going to create a new worksheet
    col_to_use = 1

    'Determining the number of days in the calendar month.
    benefit_month = MAXIS_footer_month & "/01/" & MAXIS_footer_year			'Converts whatever the next_month variable is to a MM/01/YYYY format
    num_of_days = DatePart("D", (DateAdd("D", -1, benefit_month)))			'Determines the number of days in the processing month

    processing_month = DateAdd("M", -1, benefit_month)                      'Finding the processing month
    last_day = DatePart("D", (DateAdd("D", -1, processing_month)))          'Finding the last day of the report month
    first_day = last_day - 15                                               'Getting the day 15 days before - some interviews are completed prior to the processing month

    report_month = DateAdd("M", -1, processing_month)       'Setting the report month information
    rept_mo = DatePart("M", report_month)
    rept_yr = DatePart("YYYY", report_month)

    'Going to another sheet, to enter worker-specific statistics and naming it
    excel_friendly_date = replace(date, "/", "-")
    sheet_name = "Statistics from  " & excel_friendly_date
	ObjExcel.Worksheets.Add().Name = sheet_name

    'If option to gather date frequency information was selected
    if interview_frequency_checkbox = checked or recvd_appl_frequency_checkbox = checked then
        objExcel.Cells(1, col_to_use).Value = "DATES"
        objExcel.Cells(1, col_to_use).Font.Bold = TRUE
        month_dates_col = col_to_use
        month_dates_letter_col = convert_digit_to_excel_column(month_dates_col)
        col_to_use = col_to_use + 1

        disp_mo = DatePart("M", processing_month)
        disp_yr = DatePart("YYYY", processing_month)

        if DatePart("M", benefit_month) = DatePart("M", date) then          'If data is being collected in the benefit month - we will add some more days'
            day_today = DatePart("D", date)
            bene_mo = DatePart("M", benefit_month)
            bene_yr = DatePart("YYYY", benefit_month)
        end if

        if interview_frequency_checkbox = checked then              'Setting the header if Interview frequency is requested
            objExcel.Cells(1, col_to_use).Value = "Interviews"
            objExcel.Cells(1, col_to_use).Font.Bold = TRUE
            intv_freq_col = col_to_use                              'Setting the column to a name for future use
            col_to_use = col_to_use + 1
        end if

        if recvd_appl_frequency_checkbox = checked then             'Setting the header if App Date frequence is requested
            objExcel.Cells(1, col_to_use).Value = "Apps Recvd"
            objExcel.Cells(1, col_to_use).Font.Bold = TRUE
            app_freq_col = col_to_use                               'Setting the column to a name for future use
            col_to_use = col_to_use + 1
        end if

        for each_day = first_day to last_day                                'adding 15 days of the report month to the spreadsheet
            day_entry = rept_mo & "/" & each_day & "/" & rept_yr            'writing each day in the same column
            objExcel.Cells(row_to_use, month_dates_col).Value = day_entry

            if interview_frequency_checkbox = checked then                  'counts the number of interviews on the list worksheet that match each entry
                objExcel.Cells(row_to_use, intv_freq_col).Value = "=COUNTIF('" & scenario_dropdown & "'!" & intvw_date_letter_col & ":" & intvw_date_letter_col & ", " & month_dates_letter_col & row_to_use & ")"
            end if

            if recvd_appl_frequency_checkbox = checked then                 'counts the number of applications on the list worksheet that match each entry
                objExcel.Cells(row_to_use, app_freq_col).Value = "=COUNTIF('" & scenario_dropdown & "'!" & app_date_letter_col & ":" & app_date_letter_col & ", " & month_dates_letter_col & row_to_use & ")"
            end if

            row_to_use = row_to_use + 1
        next

        for each_day = 1 to num_of_days                                         'adding day 1 to today of the benefit month.
            day_entry = disp_mo & "/" & each_day & "/" & disp_yr                'writing each day in the same column
            objExcel.Cells(row_to_use, month_dates_col).Value = day_entry

            if interview_frequency_checkbox = checked then                      'counts the number of interviews on the list worksheet that match each entry
                objExcel.Cells(row_to_use, intv_freq_col).Value = "=COUNTIF('" & scenario_dropdown & "'!" & intvw_date_letter_col & ":" & intvw_date_letter_col & ", " & month_dates_letter_col & row_to_use & ")"
            end if

            if recvd_appl_frequency_checkbox = checked then                     'counts the number of applications on the list worksheet that match each entry
                objExcel.Cells(row_to_use, app_freq_col).Value = "=COUNTIF('" & scenario_dropdown & "'!" & app_date_letter_col & ":" & app_date_letter_col & ", " & month_dates_letter_col & row_to_use & ")"
            end if

            row_to_use = row_to_use + 1
        next

        if DatePart("M", benefit_month) = DatePart("M", date) Then          'If data is being collected in the benefit month - we will add some more days'
            for each_day = 1 to day_today
                day_entry = bene_mo & "/" & each_day & "/" & bene_yr                'writing each day in the same column
                objExcel.Cells(row_to_use, month_dates_col).Value = day_entry

                if interview_frequency_checkbox = checked then                      'counts the number of interviews on the list worksheet that match each entry
                    objExcel.Cells(row_to_use, intv_freq_col).Value = "=COUNTIF('" & scenario_dropdown & "'!" & intvw_date_letter_col & ":" & intvw_date_letter_col & ", " & month_dates_letter_col & row_to_use & ")"
                end if

                if recvd_appl_frequency_checkbox = checked then                     'counts the number of applications on the list worksheet that match each entry
                    objExcel.Cells(row_to_use, app_freq_col).Value = "=COUNTIF('" & scenario_dropdown & "'!" & app_date_letter_col & ":" & app_date_letter_col & ", " & month_dates_letter_col & row_to_use & ")"
                end if

                row_to_use = row_to_use + 1
            next
        end if
        col_to_use = col_to_use + 1     'Moving over to the right
    end if

    'If statistics selected included either of the worker specific requests - this code will run
    if worker_interview_checkbox = checked or worker_recert_status_checkbox = checked then
        row_to_use = 2                  'Resetting to the top

        objExcel.Cells(1, col_to_use).Value = "WORKERS"         'Adding headers
        objExcel.Cells(1, col_to_use).Font.Bold = TRUE
        wrkr_col = col_to_use
        col_to_use = col_to_use + 1

        objExcel.Cells(1, col_to_use).Value = "TOTAL CASES"
        objExcel.Cells(1, col_to_use).Font.Bold = TRUE
        case_tot_col = col_to_use
        col_to_use = col_to_use + 1

        if worker_interview_checkbox = checked then             'Headers for intervie counts by worker
            objExcel.Cells(1, col_to_use).Value = "INCOMPLETE INTERVIEWS"
            objExcel.Cells(1, col_to_use).Font.Bold = TRUE
            incomplete_col = col_to_use
            col_to_use = col_to_use + 1

            objExcel.Cells(1, col_to_use).Value = "COMPLETE INTERVIEWS"
            objExcel.Cells(1, col_to_use).Font.Bold = TRUE
            complete_col = col_to_use
            col_to_use = col_to_use + 1

            objExcel.Cells(1, col_to_use).Value = "PERCENT COMPLETED"
            objExcel.Cells(1, col_to_use).Font.Bold = TRUE
            compl_perc_col = col_to_use
            col_to_use = col_to_use + 1

            objExcel.Cells(1, col_to_use).Value = "APPS RECEIVED"
            objExcel.Cells(1, col_to_use).Font.Bold = TRUE
            received_col = col_to_use
            col_to_use = col_to_use + 1

            objExcel.Cells(1, col_to_use).Value = "NO APP"
            objExcel.Cells(1, col_to_use).Font.Bold = TRUE
            not_recvd_col = col_to_use
            col_to_use = col_to_use + 1

            objExcel.Cells(1, col_to_use).Value = "PERCENT RECEIVED"
            objExcel.Cells(1, col_to_use).Font.Bold = TRUE
            recvd_perc_col = col_to_use
            col_to_use = col_to_use + 1
        end if

        if worker_recert_status_checkbox = checked then         'Headers for status on REVW by worker
            objExcel.Cells(1, col_to_use).Value = "REVW - I"    'For number and percent
            objExcel.Cells(1, col_to_use).Font.Bold = TRUE
            revw_i_col = col_to_use
            col_to_use = col_to_use + 1
            objExcel.Cells(1, col_to_use).Value = "%"
            objExcel.Cells(1, col_to_use).Font.Bold = TRUE
            col_to_use = col_to_use + 1

            objExcel.Cells(1, col_to_use).Value = "REVW - U"
            objExcel.Cells(1, col_to_use).Font.Bold = TRUE
            revw_u_col = col_to_use
            col_to_use = col_to_use + 1
            objExcel.Cells(1, col_to_use).Value = "%"
            objExcel.Cells(1, col_to_use).Font.Bold = TRUE
            col_to_use = col_to_use + 1

            objExcel.Cells(1, col_to_use).Value = "REVW - N"
            objExcel.Cells(1, col_to_use).Font.Bold = TRUE
            revw_n_col = col_to_use
            col_to_use = col_to_use + 1
            objExcel.Cells(1, col_to_use).Value = "%"
            objExcel.Cells(1, col_to_use).Font.Bold = TRUE
            col_to_use = col_to_use + 1

            objExcel.Cells(1, col_to_use).Value = "REVW - A"
            objExcel.Cells(1, col_to_use).Font.Bold = TRUE
            revw_a_col = col_to_use
            col_to_use = col_to_use + 1
            objExcel.Cells(1, col_to_use).Value = "%"
            objExcel.Cells(1, col_to_use).Font.Bold = TRUE
            col_to_use = col_to_use + 1

            objExcel.Cells(1, col_to_use).Value = "REVW - O"
            objExcel.Cells(1, col_to_use).Font.Bold = TRUE
            revw_o_col = col_to_use
            col_to_use = col_to_use + 1
            objExcel.Cells(1, col_to_use).Value = "%"
            objExcel.Cells(1, col_to_use).Font.Bold = TRUE
            col_to_use = col_to_use + 1

            objExcel.Cells(1, col_to_use).Value = "REVW - T"
            objExcel.Cells(1, col_to_use).Font.Bold = TRUE
            revw_t_col = col_to_use
            col_to_use = col_to_use + 1
            objExcel.Cells(1, col_to_use).Value = "%"
            objExcel.Cells(1, col_to_use).Font.Bold = TRUE
            col_to_use = col_to_use + 1

            objExcel.Cells(1, col_to_use).Value = "REVW - D"
            objExcel.Cells(1, col_to_use).Font.Bold = TRUE
            revw_d_col = col_to_use
            col_to_use = col_to_use + 1
            objExcel.Cells(1, col_to_use).Value = "%"
            objExcel.Cells(1, col_to_use).Font.Bold = TRUE
            col_to_use = col_to_use + 1

        end if

        running_i_status = 0        'Setting running counts for totals
        running_u_status = 0
        running_n_status = 0
        running_a_status = 0
        running_o_status = 0
        running_t_status = 0
        running_d_status = 0

        for each x1_number in worker_array      'looping through the list of workers
            objExcel.Cells(row_to_use, wrkr_col).Value = x1_number

            case_total = 0          'Setting counts for each loop
            complt_intvw = 0
            incomplt_intvw = 0
            no_app_rcvd = 0
            app_recvd = 0
            wrkr_i_status = 0
            wrkr_u_status = 0
            wrkr_n_status = 0
            wrkr_a_status = 0
            wrkr_o_status = 0
            wrkr_t_status = 0
            wrkr_d_status = 0

            'This loop will look at each case and add to the above set counts for the selected worker
            for case_entry = 0 to UBound(ALL_CASES_ARRAY, 2)
                'if the case is in the selected basket - and is not PRIV and is in Hennepin - the case will be assesed to add to counts
                if ALL_CASES_ARRAY(wrkr_x_numb, case_entry) = x1_number AND ALL_CASES_ARRAY(priv_case, case_entry) = FALSE AND ALL_CASES_ARRAY(out_of_co, case_entry) = "" then
                    case_total = case_total + 1     'Cases looked at
                    if worker_interview_checkbox = checked then
                        If ALL_CASES_ARRAY(date_of_interview, case_entry) = "" then
                            incomplt_intvw = incomplt_intvw + 1     'incomplete interviews
                        ELSE
                            complt_intvw = complt_intvw + 1         'complete interviews
                        end if

                        If ALL_CASES_ARRAY(date_of_app, case_entry) = "" then
                            no_app_rcvd = no_app_rcvd + 1           'no apps
                        ELSE
                            app_recvd = app_recvd + 1               'apps received
                        end if
                    end if

                    if worker_recert_status_checkbox = checked then
                        SELECT CASE ALL_CASES_ARRAY(REVW_code, case_entry)

                        CASE "I"                                    'count for each REVW status
                            wrkr_i_status = wrkr_i_status + 1
                            running_i_status = running_i_status + 1
                        CASE "U"
                            wrkr_u_status = wrkr_u_status + 1
                            running_u_status = running_u_status + 1
                        CASE "N"
                            wrkr_n_status = wrkr_n_status + 1
                            running_n_status = running_n_status + 1
                        CASE "A"
                            wrkr_a_status = wrkr_a_status + 1
                            running_a_status = running_a_status + 1
                        CASE "O"
                            wrkr_o_status = wrkr_o_status + 1
                            running_o_status = running_o_status + 1
                        CASE "T"
                            wrkr_t_status = wrkr_t_status + 1
                            running_t_status = running_t_status + 1
                        CASE "D"
                            wrkr_d_status = wrkr_d_status + 1
                            running_d_status = running_d_status + 1
                        END SELECT
                    end if
                end if
            next

            'each count will be added to the spreadsheet
            objExcel.Cells(row_to_use, case_tot_col).Value = case_total
            if worker_interview_checkbox = checked then
                objExcel.Cells(row_to_use, incomplete_col).Value = incomplt_intvw
                objExcel.Cells(row_to_use, complete_col).Value = complt_intvw
                if case_total <> 0 then
                    objExcel.Cells(row_to_use, compl_perc_col).Value = complt_intvw / case_total
                else
                    objExcel.Cells(row_to_use, compl_perc_col).Value = 0
                end if
                objExcel.Cells(row_to_use, compl_perc_col).NumberFormat = "0.00%"
                objExcel.Cells(row_to_use, received_col).Value = app_recvd
                objExcel.Cells(row_to_use, not_recvd_col).Value = no_app_rcvd
                if case_total <> 0 then
                    objExcel.Cells(row_to_use, recvd_perc_col).Value = app_recvd / case_total
                else
                    objExcel.Cells(row_to_use, recvd_perc_col).Value = 0
                end if
                objExcel.Cells(row_to_use, recvd_perc_col).NumberFormat = "0.00%"
            end if

            if worker_recert_status_checkbox = checked then
                objExcel.Cells(row_to_use, revw_i_col).Value = wrkr_i_status
                objExcel.Cells(row_to_use, revw_u_col).Value = wrkr_u_status
                objExcel.Cells(row_to_use, revw_n_col).Value = wrkr_n_status
                objExcel.Cells(row_to_use, revw_a_col).Value = wrkr_a_status
                objExcel.Cells(row_to_use, revw_o_col).Value = wrkr_o_status
                objExcel.Cells(row_to_use, revw_t_col).Value = wrkr_t_status
                objExcel.Cells(row_to_use, revw_d_col).Value = wrkr_d_status

                if case_total <> 0 then
                    objExcel.Cells(row_to_use, revw_i_col + 1).Value = wrkr_i_status / case_total
                    objExcel.Cells(row_to_use, revw_u_col + 1).Value = wrkr_u_status / case_total
                    objExcel.Cells(row_to_use, revw_n_col + 1).Value = wrkr_n_status / case_total
                    objExcel.Cells(row_to_use, revw_a_col + 1).Value = wrkr_a_status / case_total
                    objExcel.Cells(row_to_use, revw_o_col + 1).Value = wrkr_o_status / case_total
                    objExcel.Cells(row_to_use, revw_t_col + 1).Value = wrkr_t_status / case_total
                    objExcel.Cells(row_to_use, revw_d_col + 1).Value = wrkr_d_status / case_total
                else
                    objExcel.Cells(row_to_use, revw_i_col + 1).Value = 0
                    objExcel.Cells(row_to_use, revw_u_col + 1).Value = 0
                    objExcel.Cells(row_to_use, revw_n_col + 1).Value = 0
                    objExcel.Cells(row_to_use, revw_a_col + 1).Value = 0
                    objExcel.Cells(row_to_use, revw_o_col + 1).Value = 0
                    objExcel.Cells(row_to_use, revw_t_col + 1).Value = 0
                    objExcel.Cells(row_to_use, revw_d_col + 1).Value = 0
                end if
                objExcel.Cells(row_to_use, revw_i_col + 1).NumberFormat = "0.00%"
                objExcel.Cells(row_to_use, revw_u_col + 1).NumberFormat = "0.00%"
                objExcel.Cells(row_to_use, revw_n_col + 1).NumberFormat = "0.00%"
                objExcel.Cells(row_to_use, revw_a_col + 1).NumberFormat = "0.00%"
                objExcel.Cells(row_to_use, revw_o_col + 1).NumberFormat = "0.00%"
                objExcel.Cells(row_to_use, revw_t_col + 1).NumberFormat = "0.00%"
                objExcel.Cells(row_to_use, revw_d_col + 1).NumberFormat = "0.00%"
            end if


            row_to_use = row_to_use + 1
        next

        'Now the total counts will be added
        objExcel.Cells(row_to_use, wrkr_col).Value = "TOTAL"
        objExcel.Cells(row_to_use, case_tot_col).Value = running_total

        if worker_recert_status_checkbox = checked then
            objExcel.Cells(row_to_use, revw_i_col).Value = running_i_status
            objExcel.Cells(row_to_use, revw_i_col + 1).Value = running_i_status / running_total
            objExcel.Cells(row_to_use, revw_i_col + 1).NumberFormat = "0.00%"

            objExcel.Cells(row_to_use, revw_u_col).Value = running_u_status
            objExcel.Cells(row_to_use, revw_u_col + 1).Value = running_u_status / running_total
            objExcel.Cells(row_to_use, revw_u_col + 1).NumberFormat = "0.00%"

            objExcel.Cells(row_to_use, revw_n_col).Value = running_n_status
            objExcel.Cells(row_to_use, revw_n_col + 1).Value = running_n_status / running_total
            objExcel.Cells(row_to_use, revw_n_col + 1).NumberFormat = "0.00%"

            objExcel.Cells(row_to_use, revw_a_col).Value = running_a_status
            objExcel.Cells(row_to_use, revw_a_col + 1).Value = running_a_status / running_total
            objExcel.Cells(row_to_use, revw_a_col + 1).NumberFormat = "0.00%"

            objExcel.Cells(row_to_use, revw_o_col).Value = running_o_status
            objExcel.Cells(row_to_use, revw_o_col + 1).Value = running_o_status / running_total
            objExcel.Cells(row_to_use, revw_o_col + 1).NumberFormat = "0.00%"

            objExcel.Cells(row_to_use, revw_t_col).Value = running_t_status
            objExcel.Cells(row_to_use, revw_t_col + 1).Value = running_t_status / running_total
            objExcel.Cells(row_to_use, revw_t_col + 1).NumberFormat = "0.00%"

            objExcel.Cells(row_to_use, revw_d_col).Value = running_d_status
            objExcel.Cells(row_to_use, revw_d_col + 1).Value = running_d_status / running_total
            objExcel.Cells(row_to_use, revw_d_col + 1).NumberFormat = "0.00%"
        end if
        col_to_use = col_to_use + 1
    end if

    'adding a stats block of information
    stats_header_col = col_to_use
    stats_col = col_to_use + 1
    stats_letter_col = convert_digit_to_excel_column(stats_col)
    entry_row = 1

    objExcel.Cells(entry_row, stats_header_col).Value       = "Query run on: "
    objExcel.Cells(entry_row, stats_header_col).Font.Bold 	= TRUE
    objExcel.Cells(entry_row, stats_col).Value              = now
    entry_row = entry_row + 1

    objExcel.Cells(entry_row, stats_header_col).Value       = "Runtime(in seconds):"
    objExcel.Cells(entry_row, stats_header_col).Font.Bold 	= TRUE
    objExcel.Cells(entry_row, stats_col).Value              = timer - query_start_time
    entry_row = entry_row + 1

    objExcel.Cells(entry_row, stats_header_col).Value       = "Cases Assessed:"
    objExcel.Cells(entry_row, stats_header_col).Font.Bold 	= TRUE
    objExcel.Cells(entry_row, stats_col).Value              = total_cases
    total_entry_row = entry_row
    entry_row = entry_row + 1

    objExcel.Cells(entry_row, stats_header_col).Value       = "All interviews completed:"
    objExcel.Cells(entry_row, stats_header_col).Font.Bold 	= TRUE
    objExcel.Cells(entry_row, stats_col).Value              = "=COUNTIF('" & scenario_dropdown & "'!" & intvw_date_letter_col & excel_row_to_start & ":" & intvw_date_letter_col & last_excel_row & ", " & is_not_blank_excel_string & ")"
    objExcel.Cells(entry_row, stats_col+1).Value            = "=" & stats_letter_col & entry_row & "/" & stats_letter_col & total_entry_row
    objExcel.Cells(entry_row, stats_col+1).NumberFormat     = "0.00%"
    entry_row = entry_row + 1

    if interview_deadline_checkbox = checked then
        objExcel.Cells(entry_row, stats_header_col).Value       = "Interviews completed before " & interview_deadline & ":"
        objExcel.Cells(entry_row, stats_header_col).Font.Bold 	= TRUE
        objExcel.Cells(entry_row, stats_col).Value              = "=COUNTIF('" & scenario_dropdown & "'!" & intvw_date_letter_col & excel_row_to_start & ":" & intvw_date_letter_col & last_excel_row & ", " & Chr(34) & "<" & interview_deadline & Chr(34) & ")"
        objExcel.Cells(entry_row, stats_col+1).Value            = "=" & stats_letter_col & entry_row & "/" & stats_letter_col & total_entry_row
        objExcel.Cells(entry_row, stats_col+1).NumberFormat     = "0.00%"
        entry_row = entry_row + 1
    end if

    objExcel.Cells(entry_row, stats_header_col).Value       = "Total applications received:"
    objExcel.Cells(entry_row, stats_header_col).Font.Bold 	= TRUE
    objExcel.Cells(entry_row, stats_col).Value              = "=COUNTIF('" & scenario_dropdown & "'!" & app_date_letter_col & excel_row_to_start & ":" & app_date_letter_col & last_excel_row & ", " & is_not_blank_excel_string & ")"
    objExcel.Cells(entry_row, stats_col+1).Value            = "=" & stats_letter_col & entry_row & "/" & stats_letter_col & total_entry_row
    objExcel.Cells(entry_row, stats_col+1).NumberFormat     = "0.00%"
    entry_row = entry_row + 1

    if verify_appointment_notices_checkbox = checked then
        appt_notc_count = 0
        nomi_count = 0

        for case_entry = 0 to UBound(ALL_CASES_ARRAY, 2)
            if ALL_CASES_ARRAY(appt_notc_sent, case_entry) = TRUE then appt_notc_count = appt_notc_count + 1
            if ALL_CASES_ARRAY(nomi_sent, case_entry) = TRUE then nomi_count = nomi_count + 1
        next
        objExcel.Cells(entry_row, stats_header_col).Value       = "Total APPT Notices Confirmed:"
        objExcel.Cells(entry_row, stats_header_col).Font.Bold 	= TRUE
        objExcel.Cells(entry_row, stats_col).Value              = "=COUNTIF('" & scenario_dropdown & "'!" & appt_ltr_letter_col & excel_row_to_start & ":" & appt_ltr_letter_col & last_excel_row & ", " & Chr(34) & "Y" & Chr(34) & ")"
        entry_row = entry_row + 1

        objExcel.Cells(entry_row, stats_header_col).Value       = "Total NOMIs confirmed:"
        objExcel.Cells(entry_row, stats_header_col).Font.Bold 	= TRUE
        objExcel.Cells(entry_row, stats_col).Value              = "=COUNTIF('" & scenario_dropdown & "'!" & nomi_letter_col & excel_row_to_start & ":" & nomi_letter_col & last_excel_row & ", " & Chr(34) & "Y" & Chr(34) & ")"
        entry_row = entry_row + 1
    end if

    benefit_month = the_date_to_look_at
    processing_month = DateAdd("m", -1, benefit_month)
    reporting_month = DateAdd("m", -2, benefit_month)

    objExcel.Cells(entry_row, stats_header_col).Value       = "Applications Received in " & MonthName(Month(reporting_month), true)
    objExcel.Cells(entry_row, stats_header_col).Font.Bold 	= TRUE
    objExcel.Cells(entry_row, stats_col).Value              = "=COUNTIF('" & scenario_dropdown & "'!" & app_date_letter_col & excel_row_to_start & ":" & app_date_letter_col & last_excel_row & ", " & Chr(34) & "<" & processing_month & Chr(34) & ")"
    objExcel.Cells(entry_row, stats_col+1).Value            = "=" & stats_letter_col & entry_row & "/" & stats_letter_col & total_entry_row
    objExcel.Cells(entry_row, stats_col+1).NumberFormat     = "0.00%"
    entry_row = entry_row + 1

    objExcel.Cells(entry_row, stats_header_col).Value       = "Interviews Completed in " & MonthName(Month(reporting_month), true)
    objExcel.Cells(entry_row, stats_header_col).Font.Bold 	= TRUE
    objExcel.Cells(entry_row, stats_col).Value              = "=COUNTIF('" & scenario_dropdown & "'!" & intvw_date_letter_col & excel_row_to_start & ":" & intvw_date_letter_col & last_excel_row & ", " & Chr(34) & "<" & processing_month & Chr(34) & ")"
    objExcel.Cells(entry_row, stats_col+1).Value            = "=" & stats_letter_col & entry_row & "/" & stats_letter_col & total_entry_row
    objExcel.Cells(entry_row, stats_col+1).NumberFormat     = "0.00%"
    entry_row = entry_row + 1

    objExcel.Cells(entry_row, stats_header_col).Value       = "Interviews Completed in " & MonthName(Month(benefit_month), true)
    objExcel.Cells(entry_row, stats_header_col).Font.Bold 	= TRUE
    objExcel.Cells(entry_row, stats_col).Value              = "=COUNTIF('" & scenario_dropdown & "'!" & intvw_date_letter_col & excel_row_to_start & ":" & intvw_date_letter_col & last_excel_row & ", " & Chr(34) & ">=" & benefit_month & Chr(34) & ")"
    objExcel.Cells(entry_row, stats_col+1).Value            = "=" & stats_letter_col & entry_row & "/" & stats_letter_col & total_entry_row
    objExcel.Cells(entry_row, stats_col+1).NumberFormat     = "0.00%"
    entry_row = entry_row + 1

    objExcel.Cells(entry_row, stats_header_col).Value       = MonthName(Month(benefit_month), true) & " Interviews that have been Approved"
    objExcel.Cells(entry_row, stats_header_col).Font.Bold 	= TRUE
    objExcel.Cells(entry_row, stats_col).Value              = "=COUNTIFS('" & scenario_dropdown & "'!" & intvw_date_letter_col & excel_row_to_start & ":" & intvw_date_letter_col & last_excel_row & ", " & Chr(34) & ">=" & benefit_month & Chr(34) & ",'" & scenario_dropdown & "'!" & revw_code_letter_col & excel_row_to_start & ":" & revw_code_letter_col & last_excel_row & ", " & Chr(34) & "A" & Chr(34) & ")"
    entry_row = entry_row + 1

    objExcel.Cells(entry_row, stats_header_col).Value       = "Privleged Cases:"
    objExcel.Cells(entry_row, stats_header_col).Font.Bold 	= TRUE
end if

'Creating the list of privileged cases and adding to the spreadsheet
If priv_case_list <> "" Then
	priv_case_list = right(priv_case_list, (len(priv_case_list)-1))
	prived_case_array = split(priv_case_list, "|")

	FOR EACH MAXIS_case_number in prived_case_array
		objExcel.cells(entry_row, stats_col).value = MAXIS_case_number
		entry_row = entry_row + 1
	NEXT
Else
    objExcel.cells(entry_row, stats_col).value = "None"
End If

'set column size
For col_to_autofit = 1 to stats_col
    ObjExcel.columns(col_to_autofit).AutoFit()
Next

script_end_procedure("Notices have been sent. Detail of script run is on the spreadsheet that was opened.")
