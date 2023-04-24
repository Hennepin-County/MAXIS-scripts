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
CALL changelog_update("04/10/2019", "Updated the default for running the script in Data Only.", "Casey Love, Hennepin County")
CALL changelog_update("07/20/2018", "Updated verbiage for Appointment Notices and NOMIs", "Casey Love, Hennepin County")
call changelog_update("06/01/2018", "Initial version.", "Casey Love, Hennepin County")

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

function convert_date_to_day_first(date_to_convert, date_to_output)
    If IsDate(date_to_convert) = TRUE Then
        intv_date_mo = DatePart("m", date_to_convert)
        intv_date_day = DatePart("d", date_to_convert)
        intv_date_yr = DatePart("yyyy", date_to_convert)
        date_to_output = intv_date_day & "/" & intv_date_mo & "/" & intv_date_yr
    End If
end function

'SCRIPT ====================================================================================================================
'Connects to BlueZone
EMConnect ""

' 'Grabbing the worker's X number.
' CALL find_variable("User: ", worker_number, 7)
' get_county_code

' 'if user is not Hennepin County - the script will end. Process is not approved for other counties
' if worker_county_code <> "x127" Then script_end_procedure("This script is built specifically for a process defined in the 'On Demand Waiver' for interviews. Currenly only Hennepin County has this waiver. This script should not be run for any cases outside of Hennepin County.")


'Starting the query start time (for the query runtime at the end)
query_start_time = timer

'Script will determine if this is likely being run for the appointment notice or the NOMI
'Appointment notice will be run around the 18th of 2 months prior to the recert month
'NOMI will be run on the 15th of the month before the recert month

'Setting the initial path for the excel file to be found at - so we don't have to clickity click a bunch to get to the right file.
recertification_cases_excel_file_path = ""

'Initial Dialog which requests a file path for the excel file
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 361, 105, "On Demand Recertifications"
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

'Initial Dialog to determine the excel file to use, column with case numbers, and which process should be run
'Show initial dialog
Do
	Dialog Dialog1
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
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 151, 75, "On Demand Recertifications"
  DropListBox 5, 35, 140, 15, "Select One..." & scenario_list, scenario_dropdown
  ButtonGroup ButtonPressed
    OkButton 40, 55, 50, 15
    CancelButton 95, 55, 50, 15
  Text 5, 10, 130, 20, "Select the correct worksheet to run for recertification interview notifications:"
EndDialog

'Shows the dialog to select the correct worksheet
Do
    Dialog Dialog1
    If ButtonPressed = cancel then stopscript
Loop until scenario_dropdown <> "Select One..."

'Activates worksheet based on user selection
objExcel.worksheets(scenario_dropdown).Activate

'Creating an array of letters to loop through
col_ind = "A~B~C~D~E~F~G~H~I~J~K~L~M~N~O~P~Q~R~S~T~U~V~W~X~Y~Z~AA~AB~AC~AD~AE~AF~AG~AH~AI~AJ~AK~AL~AM~AN~AO~AP~AQ~AR~AS~AT~AU~AV~AW~AX~AY~AZ~BA~BB~BC~BD~BE~BF~BG~BH~BI~BJ~BK~BL~BM~BN~BO~BP~BQ~BR~BS~BT~BU~BV~BW~BX~BY~BZ"
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
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 266, 135, "On Demand Recertifications"
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
    Dialog Dialog1
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

    MAXIS_footer_month = CM_mo
    MAXIS_footer_year = CM_yr

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
            if ucase(col_header) = "X NUMBER" then worker_col = letter & " - " & col_header    'if the first cell says 'Case Number' then it is likely the correct column
        Else
            Exit For
        End If
        cell_val = cell_val + 1
    Next

    interview_deadline_checkbox = checked
    beg_of_recert_pd = MAXIS_footer_month & "/1/" & MAXIS_footer_year
    interview_deadline = DateAdd("d", -1, beg_of_recert_pd)
    interview_deadline = interview_deadline & ""
    verify_appointment_notices_checkbox  = checked
    interview_frequency_checkbox = checked
    recvd_appl_frequency_checkbox = checked
    worker_interview_checkbox = checked
    worker_recert_status_checkbox = unchecked

    Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 311, 110, "On Demand Recertifications"
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


    Dialog Dialog1
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

    Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 196, 150, "Confirm Selections"
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
        Dialog Dialog1
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
	letter_worker_col = wrkr_col
    call convert_excel_letter_to_excel_number(worker_col)
end if

'reading each line of the Excel file and adding case number information to the array
do
    ReDim Preserve ALL_CASES_ARRAY(notc_confirm, case_entry)
    ALL_CASES_ARRAY(case_number, case_entry) = trim(objExcel.Cells(row, col).Value)
    ALL_CASES_ARRAY(excel_row, case_entry) = row

    if notice_type = "Data Only" then
        ALL_CASES_ARRAY(wrkr_x_numb, case_entry) = trim(objExcel.Cells(row, wrkr_col).Value)
        ' if InStr(list_all_workers, trim(objExcel.Cells(row, wrkr_col).Value)) = 0 then list_all_workers = list_all_workers & trim(objExcel.Cells(row, wrkr_col).Value) & "~"
    end if

    ALL_CASES_ARRAY(MFIP_case, case_entry) = FALSE
    ALL_CASES_ARRAY(SNAP_case, case_entry) = FALSE

    case_entry = case_entry + 1
    row = row + 1
    next_case_number = trim(objExcel.Cells(row, col).Value)
loop until next_case_number = ""
last_excel_row = row -1

' total_cases = case_entry
' if notice_type = "Data Only" then
'     list_all_workers = left(list_all_workers, len(list_all_workers)-1)
'     worker_array = split(list_all_workers, "~")
' end if

running_total = 0
running_interviews_complete = 0
running_interviews_not_done = 0
running_no_app = 0
running_app_recvd = 0

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
appt_lrt_col = 15
nomi_col = 16
revw_code_col = 17
intvw_date_col = 18
app_date_col = 19

appt_ltr_letter_col = convert_digit_to_excel_column(appt_lrt_col)
nomi_letter_col = convert_digit_to_excel_column(nomi_col)

revw_code_letter_col = convert_digit_to_excel_column(revw_code_col)
intvw_date_letter_col = convert_digit_to_excel_column(intvw_date_col)
app_date_letter_col = convert_digit_to_excel_column(app_date_col)
' If notice_type = "NOMI" or notice_type = "Data Only" Then
'     objRange.Insert(xlShiftToRight)     'add column for review status
'     objRange.Insert(xlShiftToRight)     'add column with interview date
'     objRange.Insert(xlShiftToRight)     'add column with app date
'     If notice_type = "Data Only" Then objRange.Insert(xlShiftToRight)     'add another column for the other notice confirmation
'
'     If notice_type = "NOMI" Then
'         revw_code_col = notc_col + 1
'         nomi_letter_col = convert_digit_to_excel_column(notc_col)
'     ElseIf notice_type = "Data Only" Then
'         appt_lrt_col = notc_col
'         nomi_col = notc_col + 1
'         revw_code_col = nomi_col + 1        'setting variables for writing to excel'
'
'         appt_ltr_letter_col = convert_digit_to_excel_column(appt_lrt_col)
'         nomi_letter_col = convert_digit_to_excel_column(nomi_col)
'     End If
'     intvw_date_col = revw_code_col + 1
'     app_date_col = intvw_date_col + 1
'
'     revw_code_letter_col = convert_digit_to_excel_column(revw_code_col)
'     intvw_date_letter_col = convert_digit_to_excel_column(intvw_date_col)
'     app_date_letter_col = convert_digit_to_excel_column(app_date_col)
'
'     If notice_type = "NOMI" Then objExcel.Cells(1, notc_col).Value = "NOMI Success"          'adding headers to Excel
'     If notice_type = "Data Only" Then
'         objExcel.Cells(1, appt_lrt_col).Value = "Appt LTR Confirm"
'         objExcel.Cells(1, nomi_col).Value = "NOMI Confirm"
'     End If
'     objExcel.Cells(1, revw_code_col).Value = "REVW Status"
'     objExcel.Cells(1, intvw_date_col).Value = "Interview Date"
'     objExcel.Cells(1, app_date_col).Value = "Date App Rec'vd"
'
'     stats_header_col = app_date_col + 2    'Setting variables with coumn locations for statistics
'     stats_col = app_date_col + 3
' End If

'creating a variable in the MM/DD/YY format to compare with date read from MAXIS

'reading each line of the Excel file and adding case number information to the array

if notice_type = "Data Only" then
	row = 2
	do
        if InStr(list_all_workers, trim(objExcel.Cells(row, wrkr_col).Value)) = 0 then list_all_workers = list_all_workers & trim(objExcel.Cells(row, wrkr_col).Value) & "~"
    	row = row + 1
    	next_case_number = trim(objExcel.Cells(row, col).Value)
	loop until next_case_number = ""

    list_all_workers = left(list_all_workers, len(list_all_workers)-1)
    worker_array = split(list_all_workers, "~")
end if


today_mo = DatePart("m", date)
today_mo = right("00" & today_mo, 2)

today_day = DatePart("d", date)
today_day = right("00" & today_day, 2)

today_yr = DatePart("yyyy", date)
today_yr = right(today_yr, 2)

today_date = today_mo & "/" & today_day & "/" & today_yr

interviews_completed_timely = 0     'setting this here to the spreadsheet will read 0 if none found.

if notice_type = "Data Only" then
    row_to_use = 2      'Setting row and column because we are going to create a new worksheet
    col_to_use = 1

    'Determining the number of days in the calendar month.
    benefit_month = MAXIS_footer_month & "/01/" & MAXIS_footer_year			'Converts whatever the next_month variable is to a MM/01/YYYY format
    month_after_bene = DateAdd("M", 1, benefit_month)
    num_of_days = DatePart("D", (DateAdd("D", -1, benefit_month)))			'Determines the number of days in the processing month

    processing_month = DateAdd("M", -1, benefit_month)                      'Finding the processing month
    last_day = DatePart("D", (DateAdd("D", -1, processing_month)))          'Finding the last day of the report month
    first_day = last_day - 15                                               'Getting the day 15 days before - some interviews are completed prior to the processing month

    report_month = DateAdd("M", -1, processing_month)       'Setting the report month information
    rept_mo = DatePart("M", report_month)
    rept_yr = DatePart("YYYY", report_month)

    'Going to another sheet, to enter worker-specific statistics and naming it
    excel_friendly_date = replace(date, "/", "-")
	sheet_start = left(scenario_dropdown, 3)
	sheet_start = trim(sheet_start)
    sheet_name = sheet_start & " Statistics from  " & excel_friendly_date
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
        else
            last_day_of_benefit_month = DatePart("D", (DateAdd("D", -1, month_after_bene)))
            end_of_benefit_month = MAXIS_footer_month & "/" & last_day_of_benefit_month & "/" & MAXIS_footer_year

            bene_mo = DatePart("M", benefit_month)
            bene_yr = DatePart("YYYY", benefit_month)

            If DateValue(end_of_benefit_month) < date Then
                for each_day = 1 to last_day_of_benefit_month
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
            End If

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

			letter_case_tot_col = convert_digit_to_excel_column(case_tot_col)
			letter_incomplete_col = convert_digit_to_excel_column(incomplete_col)
			letter_complete_col = convert_digit_to_excel_column(complete_col)
			letter_compl_perc_col = convert_digit_to_excel_column(compl_perc_col)
			letter_received_col = convert_digit_to_excel_column(received_col)
			letter_not_recvd_col = convert_digit_to_excel_column(not_recvd_col)
			letter_recvd_perc_col = convert_digit_to_excel_column(recvd_perc_col)
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

			objExcel.Cells(row_to_use, revw_i_col + 1).Value = 0
			objExcel.Cells(row_to_use, revw_u_col + 1).Value = 0
			objExcel.Cells(row_to_use, revw_n_col + 1).Value = 0
			objExcel.Cells(row_to_use, revw_a_col + 1).Value = 0
			objExcel.Cells(row_to_use, revw_o_col + 1).Value = 0
			objExcel.Cells(row_to_use, revw_t_col + 1).Value = 0
			objExcel.Cells(row_to_use, revw_d_col + 1).Value = 0

			letter_revw_i_col = convert_digit_to_excel_column(revw_i_col)
			letter_revw_u_col = convert_digit_to_excel_column(revw_u_col)
			letter_revw_n_col = convert_digit_to_excel_column(revw_n_col)
			letter_revw_a_col = convert_digit_to_excel_column(revw_a_col)
			letter_revw_o_col = convert_digit_to_excel_column(revw_o_col)
			letter_revw_t_col = convert_digit_to_excel_column(revw_t_col)
			letter_revw_d_col = convert_digit_to_excel_column(revw_d_col)

        end if

        ' running_i_status = 0        'Setting running counts for totals
        ' running_u_status = 0
        ' running_n_status = 0
        ' running_a_status = 0
        ' running_o_status = 0
        ' running_t_status = 0
        ' running_d_status = 0

		objExcel.ActiveSheet.Range("A2").Select
		objExcel.ActiveWindow.FreezePanes = True

		is_not_blank_excel_string = chr(34) & "<>" & chr(34)

		letter_wrkr_col = convert_digit_to_excel_column(wrkr_col)



        for each x1_number in worker_array      'looping through the list of workers
            objExcel.Cells(row_to_use, wrkr_col).Value = x1_number

			' MsgBOx "=COUNTIF('" & scenario_dropdown & "'!" & letter_worker_col & "2:" & letter_worker_col & last_excel_row & ", '" & sheet_name &"'!" & letter_wrkr_col & row_to_use & ")"
			objExcel.Cells(row_to_use, case_tot_col).Value = "=COUNTIF('" & scenario_dropdown & "'!" & letter_worker_col & "2:" & letter_worker_col & last_excel_row & ", '" & sheet_name &"'!" & letter_wrkr_col & row_to_use & ")"
			if worker_interview_checkbox = checked then

				objExcel.Cells(row_to_use, incomplete_col).Value = "=COUNTIFS('" & scenario_dropdown & "'!" & letter_worker_col & "2:" & letter_worker_col & last_excel_row & ", '" & sheet_name &"'!" & letter_wrkr_col & row_to_use & ", '" & scenario_dropdown & "'!" & intvw_date_letter_col & "2:" & intvw_date_letter_col & last_excel_row & ", " & chr(34) & chr(34) & ")"
				objExcel.Cells(row_to_use, complete_col).Value = "=COUNTIFS('" & scenario_dropdown & "'!" & letter_worker_col & "2:" & letter_worker_col & last_excel_row & ", '" & sheet_name &"'!" & letter_wrkr_col & row_to_use & ", '" & scenario_dropdown & "'!" & intvw_date_letter_col & "2:" & intvw_date_letter_col & last_excel_row & ", " & is_not_blank_excel_string & ")"


				objExcel.Cells(row_to_use, compl_perc_col).Value = "=" & letter_complete_col & row_to_use & "/" & letter_case_tot_col & row_to_use
				objExcel.Cells(row_to_use, compl_perc_col).NumberFormat = "0.00%"

				objExcel.Cells(row_to_use, received_col).Value = "=COUNTIFS('" & scenario_dropdown & "'!" & letter_worker_col & "2:" & letter_worker_col & last_excel_row & ", '" & sheet_name &"'!" & letter_wrkr_col & row_to_use & ", '" & scenario_dropdown & "'!" & app_date_letter_col & "2:" & app_date_letter_col & last_excel_row & ", " & is_not_blank_excel_string & ")"
				objExcel.Cells(row_to_use, not_recvd_col).Value = "=COUNTIFS('" & scenario_dropdown & "'!" & letter_worker_col & "2:" & letter_worker_col & last_excel_row & ", '" & sheet_name &"'!" & letter_wrkr_col & row_to_use & ", '" & scenario_dropdown & "'!" & app_date_letter_col & "2:" & app_date_letter_col & last_excel_row & ", " & chr(34) & chr(34) & ")"

				objExcel.Cells(row_to_use, recvd_perc_col).Value = "=" & letter_received_col & row_to_use & "/" & letter_case_tot_col & row_to_use
				objExcel.Cells(row_to_use, recvd_perc_col).NumberFormat = "0.00%"
			end if

			if worker_recert_status_checkbox = checked then
				objExcel.Cells(row_to_use, revw_i_col).Value = "=COUNTIFS('" & scenario_dropdown & "'!" & letter_worker_col & "2:" & letter_worker_col & last_excel_row & ", '" & sheet_name &"'!" & letter_wrkr_col & row_to_use & ", '" & scenario_dropdown & "'!" & revw_code_letter_col & "2:" & revw_code_letter_col & last_excel_row & ", " & chr(34) & "I" & chr(34) & ")"
				objExcel.Cells(row_to_use, revw_u_col).Value = "=COUNTIFS('" & scenario_dropdown & "'!" & letter_worker_col & "2:" & letter_worker_col & last_excel_row & ", '" & sheet_name &"'!" & letter_wrkr_col & row_to_use & ", '" & scenario_dropdown & "'!" & revw_code_letter_col & "2:" & revw_code_letter_col & last_excel_row & ", " & chr(34) & "U" & chr(34) & ")"
				objExcel.Cells(row_to_use, revw_n_col).Value = "=COUNTIFS('" & scenario_dropdown & "'!" & letter_worker_col & "2:" & letter_worker_col & last_excel_row & ", '" & sheet_name &"'!" & letter_wrkr_col & row_to_use & ", '" & scenario_dropdown & "'!" & revw_code_letter_col & "2:" & revw_code_letter_col & last_excel_row & ", " & chr(34) & "N" & chr(34) & ")"
				objExcel.Cells(row_to_use, revw_a_col).Value = "=COUNTIFS('" & scenario_dropdown & "'!" & letter_worker_col & "2:" & letter_worker_col & last_excel_row & ", '" & sheet_name &"'!" & letter_wrkr_col & row_to_use & ", '" & scenario_dropdown & "'!" & revw_code_letter_col & "2:" & revw_code_letter_col & last_excel_row & ", " & chr(34) & "A" & chr(34) & ")"
				objExcel.Cells(row_to_use, revw_o_col).Value = "=COUNTIFS('" & scenario_dropdown & "'!" & letter_worker_col & "2:" & letter_worker_col & last_excel_row & ", '" & sheet_name &"'!" & letter_wrkr_col & row_to_use & ", '" & scenario_dropdown & "'!" & revw_code_letter_col & "2:" & revw_code_letter_col & last_excel_row & ", " & chr(34) & "O" & chr(34) & ")"
				objExcel.Cells(row_to_use, revw_t_col).Value = "=COUNTIFS('" & scenario_dropdown & "'!" & letter_worker_col & "2:" & letter_worker_col & last_excel_row & ", '" & sheet_name &"'!" & letter_wrkr_col & row_to_use & ", '" & scenario_dropdown & "'!" & revw_code_letter_col & "2:" & revw_code_letter_col & last_excel_row & ", " & chr(34) & "T" & chr(34) & ")"
				objExcel.Cells(row_to_use, revw_d_col).Value = "=COUNTIFS('" & scenario_dropdown & "'!" & letter_worker_col & "2:" & letter_worker_col & last_excel_row & ", '" & sheet_name &"'!" & letter_wrkr_col & row_to_use & ", '" & scenario_dropdown & "'!" & revw_code_letter_col & "2:" & revw_code_letter_col & last_excel_row & ", " & chr(34) & "D" & chr(34) & ")"


				objExcel.Cells(row_to_use, revw_i_col + 1).Value = "=" & letter_revw_i_col & row_to_use & "/" & letter_case_tot_col & row_to_use
				objExcel.Cells(row_to_use, revw_u_col + 1).Value = "=" & letter_revw_u_col & row_to_use & "/" & letter_case_tot_col & row_to_use
				objExcel.Cells(row_to_use, revw_n_col + 1).Value = "=" & letter_revw_n_col & row_to_use & "/" & letter_case_tot_col & row_to_use
				objExcel.Cells(row_to_use, revw_a_col + 1).Value = "=" & letter_revw_a_col & row_to_use & "/" & letter_case_tot_col & row_to_use
				objExcel.Cells(row_to_use, revw_o_col + 1).Value = "=" & letter_revw_o_col & row_to_use & "/" & letter_case_tot_col & row_to_use
				objExcel.Cells(row_to_use, revw_t_col + 1).Value = "=" & letter_revw_t_col & row_to_use & "/" & letter_case_tot_col & row_to_use
				objExcel.Cells(row_to_use, revw_d_col + 1).Value = "=" & letter_revw_d_col & row_to_use & "/" & letter_case_tot_col & row_to_use
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

		objExcel.Cells(row_to_use, wrkr_col).Value = "TOTAL"

		objExcel.Cells(row_to_use, case_tot_col).Value = "=SUM(" & letter_case_tot_col & "2:" & letter_case_tot_col & row_to_use - 1 & ")"

		if worker_interview_checkbox = checked then
			objExcel.Cells(row_to_use, incomplete_col).Value = "=SUM(" & letter_incomplete_col & "2:" & letter_incomplete_col & row_to_use - 1 & ")"
			objExcel.Cells(row_to_use, complete_col).Value = "=SUM(" & letter_complete_col & "2:" & letter_complete_col &  row_to_use - 1 & ")"

			objExcel.Cells(row_to_use, compl_perc_col).Value = "=" & letter_complete_col & row_to_use & "/" & letter_case_tot_col & row_to_use
			objExcel.Cells(row_to_use, compl_perc_col).NumberFormat = "0.00%"

			objExcel.Cells(row_to_use, received_col).Value = "=SUM(" & letter_received_col & "2:" & letter_received_col & row_to_use - 1 & ")"
			objExcel.Cells(row_to_use, not_recvd_col).Value = "=SUM(" & letter_not_recvd_col & "2:" & letter_not_recvd_col & row_to_use - 1 & ")"

			objExcel.Cells(row_to_use, recvd_perc_col).Value = "=" & letter_received_col & row_to_use & "/" & letter_case_tot_col & row_to_use
			objExcel.Cells(row_to_use, recvd_perc_col).NumberFormat = "0.00%"
		end if

		if worker_recert_status_checkbox = checked then
			objExcel.Cells(row_to_use, revw_i_col).Value = "=SUM(" & letter_revw_i_col & "2:" & letter_revw_i_col & row_to_use - 1 & ")"
			objExcel.Cells(row_to_use, revw_u_col).Value = "=SUM(" & letter_revw_u_col & "2:" & letter_revw_u_col & row_to_use - 1 & ")"
			objExcel.Cells(row_to_use, revw_n_col).Value = "=SUM(" & letter_revw_n_col & "2:" & letter_revw_n_col & row_to_use - 1 & ")"
			objExcel.Cells(row_to_use, revw_a_col).Value = "=SUM(" & letter_revw_a_col & "2:" & letter_revw_a_col & row_to_use - 1 & ")"
			objExcel.Cells(row_to_use, revw_o_col).Value = "=SUM(" & letter_revw_o_col & "2:" & letter_revw_o_col & row_to_use - 1 & ")"
			objExcel.Cells(row_to_use, revw_t_col).Value = "=SUM(" & letter_revw_t_col & "2:" & letter_revw_t_col & row_to_use - 1 & ")"
			objExcel.Cells(row_to_use, revw_d_col).Value = "=SUM(" & letter_revw_d_col & "2:" & letter_revw_d_col & row_to_use - 1 & ")"

			objExcel.Cells(row_to_use, revw_i_col + 1).Value = "=" & letter_revw_i_col & row_to_use & "/" & letter_case_tot_col & row_to_use
			objExcel.Cells(row_to_use, revw_u_col + 1).Value = "=" & letter_revw_u_col & row_to_use & "/" & letter_case_tot_col & row_to_use
			objExcel.Cells(row_to_use, revw_n_col + 1).Value = "=" & letter_revw_n_col & row_to_use & "/" & letter_case_tot_col & row_to_use
			objExcel.Cells(row_to_use, revw_a_col + 1).Value = "=" & letter_revw_a_col & row_to_use & "/" & letter_case_tot_col & row_to_use
			objExcel.Cells(row_to_use, revw_o_col + 1).Value = "=" & letter_revw_o_col & row_to_use & "/" & letter_case_tot_col & row_to_use
			objExcel.Cells(row_to_use, revw_t_col + 1).Value = "=" & letter_revw_t_col & row_to_use & "/" & letter_case_tot_col & row_to_use
			objExcel.Cells(row_to_use, revw_d_col + 1).Value = "=" & letter_revw_d_col & row_to_use & "/" & letter_case_tot_col & row_to_use
			objExcel.Cells(row_to_use, revw_i_col + 1).NumberFormat = "0.00%"
			objExcel.Cells(row_to_use, revw_u_col + 1).NumberFormat = "0.00%"
			objExcel.Cells(row_to_use, revw_n_col + 1).NumberFormat = "0.00%"
			objExcel.Cells(row_to_use, revw_a_col + 1).NumberFormat = "0.00%"
			objExcel.Cells(row_to_use, revw_o_col + 1).NumberFormat = "0.00%"
			objExcel.Cells(row_to_use, revw_t_col + 1).NumberFormat = "0.00%"
			objExcel.Cells(row_to_use, revw_d_col + 1).NumberFormat = "0.00%"

        end if

		objExcel.Rows(row_to_use).Font.Bold = TRUE
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
        objExcel.Cells(entry_row, stats_col).Value              = "=COUNTIF('" & scenario_dropdown & "'!" & intvw_date_letter_col & excel_row_to_start & ":" & intvw_date_letter_col & last_excel_row & ", " & Chr(34) & "<=" & interview_deadline & Chr(34) & ")"
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

	objExcel.ActiveSheet.Range("A2").Select
	objExcel.ActiveWindow.FreezePanes = True
end if
