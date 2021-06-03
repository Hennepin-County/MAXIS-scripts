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
call changelog_update("03/02/2021", "Update EZ Info Phone hours from 9-4 pm to 8-4:30 pm.", "Ilse Ferris, Hennepin County")
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

'Creating an array of letters to loop through
col_hdr = "A~B~C~D~E~F~G~H~I~J~K~L~M~N~O~P~Q~R~S~T~U~V~W~X~Y~Z~AA~AB~AC~AD~AE~AF~AG~AH~AI~AJ~AK~AL~AM~AN~AO~AP~AQ~AR~AS~AT~AU~AV~AW~AX~AY~AZ~BA~BB~BC~BD~BE~BF~BG~BH~BI~BJ~BK~BL~BM~BN~BO~BP~BQ~BR~BS~BT~BU~BV~BW~BX~BY~BZ"
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
		if notice_type = "Data Only" then
			if ucase(col_header) = "X NUMBER" then worker_col = letter & " - " & col_header    'if the first cell says 'Case Number' then it is likely the correct column
		End If
	Else
		Exit For
	End If
	cell_val = cell_val + 1
Next

if notice_type = "Data Only" then

    MAXIS_footer_month = CM_mo
    MAXIS_footer_year = CM_yr



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
    call convert_excel_letter_to_excel_number(wrkr_col)
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

total_cases = case_entry
' if notice_type = "Data Only" then
'     list_all_workers = left(list_all_workers, len(list_all_workers)-1)
'     worker_array = split(list_all_workers, "~")
' end if

running_total = 0
running_interviews_complete = 0
running_interviews_not_done = 0
running_no_app = 0
running_app_recvd = 0

revw_status_col_known = "Needs New"
interview_date_col_known = "Needs New"
app_date_col_known = "Needs New"
nomi_recvd_col_known = "Needs New"
appt_ltr_recvd_col_known = "Needs New"
nomi_success_col_known = "Needs New"
appt_ltr_success_col_known = "Needs New"

If excel_row_to_start <> "2" Then

	header_list = replace(header_list, "Select One...", "Needs New")
	MsgBox header_list

	Dialog1 = ""
	If notice_type = "Appointment Notice" Then
		BeginDialog Dialog1, 0, 0, 236, 80, "Select a Known Column"
		  DropListBox 90, 40, 140, 45, header_list, appt_ltr_success_col_known
		  ButtonGroup ButtonPressed
		    OkButton 120, 60, 50, 15
		    CancelButton 180, 60, 50, 15
		  Text 10, 10, 185, 20, "Since you are not starting at the beginning, you can select if a column is already created to store the data."
		  Text 20, 45, 70, 10, "Appt Notice Success"
		EndDialog
	ElseIf notice_type = "NOMI" Then
		BeginDialog Dialog1, 0, 0, 236, 135, "Select a Known Column"
		  DropListBox 90, 35, 140, 45, header_list, nomi_success_col_known
		  DropListBox 90, 55, 140, 45, header_list, revw_status_col_known
		  DropListBox 90, 75, 140, 45, header_list, interview_date_col_known
		  DropListBox 90, 95, 140, 45, header_list, app_date_col_known
		  ButtonGroup ButtonPressed
		    OkButton 120, 115, 50, 15
		    CancelButton 180, 115, 50, 15
		  Text 10, 10, 185, 20, "Since you are not starting at the beginning, you can select if a column is already created to store the data."
		  Text 40, 40, 50, 10, "NOMI Success"
		  Text 40, 60, 50, 10, "REVW Status"
		  Text 35, 80, 50, 10, "Interview Date"
		  Text 35, 100, 55, 10, "Date App Recvd"
		EndDialog
	ElseIf notice_type = "Data Only" Then
		BeginDialog Dialog1, 0, 0, 236, 165, "Select a Known Column"
		  ComboBox 90, 40, 140, 45, header_list, appt_ltr_recvd_col_known
		  ComboBox 90, 60, 140, 45, header_list, nomi_recvd_col_known
		  ComboBox 90, 80, 140, 45, header_list, revw_status_col_known
		  ComboBox 90, 100, 140, 45, header_list, interview_date_col_known
		  ComboBox 90, 120, 140, 45, header_list, app_date_col_known
		  ButtonGroup ButtonPressed
		    OkButton 120, 140, 50, 15
		    CancelButton 180, 140, 50, 15
		  Text 10, 10, 185, 20, "Since you are not starting at the beginning, you can select if a column is already created to store the data."
		  Text 20, 45, 70, 10, "APPT Letter Confirm"
		  Text 40, 65, 50, 10, "NOMI Confirm"
		  Text 40, 85, 50, 10, "REVW Status"
		  Text 35, 105, 50, 10, "Interview Date"
		  Text 35, 125, 55, 10, "Date App Recvd"
		EndDialog
	End If

	Dialog Dialog1
	If ButtonPressed = cancel then stopscript
	if notice_type = "Data Only" Then
	    wrkr_col = left(worker_col, 1)
		letter_worker_col = wrkr_col
	    call convert_excel_letter_to_excel_number(wrkr_col)
	end if
End If

'Insert columns in excel for additional information to be added
column_end = last_col & "1"

Set objRange = objExcel.Range(column_end).EntireColumn

notc_col = last_col                                         'setting the a variable with the notice column for later updating of excel
notc_letter_col = notc_col
call convert_excel_letter_to_excel_number(notc_col)

if notice_type = "Appointment Notice" Then
    If appt_ltr_success_col_known = "Needs New" THen
		objExcel.Cells(1, notc_col).Value = "Appt Notice Success"   'Adding header to Excel
	Else
		appt_ltr_letter_col = left(appt_ltr_success_col_known, 2)
		appt_ltr_letter_col = trim(appt_ltr_letter_col)
		appt_lrt_col = appt_ltr_letter_col
		call convert_excel_letter_to_excel_number(appt_lrt_col)
		notc_col = appt_lrt_col
	End If
	stats_header_col = notc_col + 2         'Setting variables with coumn locations for statistics
	stats_col = notc_col + 3
End If

If notice_type = "NOMI" Then
	If nomi_success_col_known = "Needs New" Then
		objRange.Insert(xlShiftToRight)     'add column for review status
		nomi_letter_col = convert_digit_to_excel_column(notc_col)
		nomi_col = notc_col
		objExcel.Cells(1, notc_col).Value = "NOMI Success"
	Else
		nomi_letter_col = left(nomi_success_col_known, 2)
		nomi_letter_col = trim(nomi_letter_col)
		nomi_ltr_col = nomi_letter_col
		call convert_excel_letter_to_excel_number(nomi_ltr_col)
		notc_col = nomi_ltr_col
	End If
End If

If notice_type = "Data Only" Then
	If appt_ltr_recvd_col_known = "Needs New" Then
		objRange.Insert(xlShiftToRight)                             'inserting one column to the end of the data in the spreadsheet
		appt_lrt_col = notc_col
		appt_ltr_letter_col = convert_digit_to_excel_column(appt_lrt_col)
		objExcel.Cells(1, appt_lrt_col).Value = "Appt LTR Confirm"
	Else
		appt_ltr_letter_col = left(appt_ltr_recvd_col_known, 2)
		appt_ltr_letter_col = trim(appt_ltr_letter_col)
		appt_lrt_col = appt_ltr_letter_col
		call convert_excel_letter_to_excel_number(appt_lrt_col)
		notc_col = appt_lrt_col
	End If
	If nomi_recvd_col_known = "Needs New" Then
		objRange.Insert(xlShiftToRight)     'add another column for the other notice confirmation
		nomi_col = notc_col + 1
		nomi_letter_col = convert_digit_to_excel_column(nomi_col)
		objExcel.Cells(1, nomi_col).Value = "NOMI Confirm"
	Else
		nomi_letter_col = left(nomi_recvd_col_known, 2)
		nomi_letter_col = trim(nomi_letter_col)
		nomi_col = nomi_letter_col
		call convert_excel_letter_to_excel_number(nomi_col)
	End If
End If

If notice_type = "NOMI" or notice_type = "Data Only" Then
	If revw_status_col_known = "Needs New" Then
		objRange.Insert(xlShiftToRight)     'add column for review status
		revw_code_col = nomi_col + 1        'setting variables for writing to excel'
		revw_code_letter_col = convert_digit_to_excel_column(revw_code_col)
		objExcel.Cells(1, revw_code_col).Value = "REVW Status"
	Else
		revw_code_letter_col = left(revw_status_col_known, 2)
		revw_code_letter_col = trim(revw_code_letter_col)
		revw_code_col = revw_code_letter_col
		call convert_excel_letter_to_excel_number(revw_code_col)
	End If
	If interview_date_col_known = "Needs New" Then
		objRange.Insert(xlShiftToRight)     'add column with interview date
		intvw_date_col = revw_code_col + 1
		intvw_date_letter_col = convert_digit_to_excel_column(intvw_date_col)
		objExcel.Cells(1, intvw_date_col).Value = "Interview Date"
	Else
		intvw_date_letter_col = left(interview_date_col_known, 2)
		intvw_date_letter_col = trim(intvw_date_letter_col)
		intvw_date_col = intvw_date_letter_col
		call convert_excel_letter_to_excel_number(intvw_date_col)
	End If
	If app_date_col_known = "Needs New" Then
		objRange.Insert(xlShiftToRight)     'add column with app date
		app_date_col = intvw_date_col + 1
		app_date_letter_col = convert_digit_to_excel_column(app_date_col)
		objExcel.Cells(1, app_date_col).Value = "Date App Rec'vd"
		stats_header_col = app_date_col + 2    'Setting variables with coumn locations for statistics
		stats_col = app_date_col + 3
	Else
		app_date_letter_col = left(app_date_col_known, 2)
		app_date_letter_col = trim(app_date_letter_col)
		app_date_col = app_date_letter_col
		call convert_excel_letter_to_excel_number(app_date_col)

		final_col = last_col
		call convert_excel_letter_to_excel_number(final_col)
		stats_header_col = final_col + 2
		stats_col = final_col + 3
	End If

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
						' next
						' 'All information has been gathered
						' 'MsgBox "We are going back to EXCEL!"
						'
						' 'Looping through the array again to send the notice and confirm the notice will be sent (that it saved correctly)
						' 'For NOMI option, this will only send a notice IF the interview is incomplete.
						' for case_entry = 0 to UBound(ALL_CASES_ARRAY, 2)
						'     MAXIS_case_number = ALL_CASES_ARRAY(case_number, case_entry)        'setting this for using navigate functions

	'NOW we are going to actually take action and save the information to Excel.
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

                Case "07"   'Somali (2nd)
                    'MsgBox "SOMALI"
                    Memo_to_display = "Waaxda Adeegyada Aadanaha waxay kuu soo dirtay baakad warqado ah. Waraaqahani waxay cusbooneysiiyaan kiiskaaga " & programs & "."
                    Memo_to_display = Memo_to_display & vbNewLine & ""
                    Memo_to_display = Memo_to_display & vbNewLine & "Fadlan saxiix, taariikhdana ku qor oo soo celi waraaqaha cusboonaysiinta" & CM_plus_1_mo & "/08/" & CM_plus_1_yr & ". Waa inaad sidoo kale buuxusaa wareysiga " & programs & "-gaaga si kiisku u sii socdo."
                    Memo_to_display = Memo_to_display & vbNewLine & ""
                    Memo_to_display = Memo_to_display & vbNewLine & "*Fadlan dhammaystir wareysigaaga inta ka horreysa " & interview_end_date & "*"
                    Memo_to_display = Memo_to_display & vbNewLine & "Si aad u dhamaystirto wareysiga telefoonka, wac laynka taleefanka EZ 612-596-1300 inta u dhaxaysa 8:00 subaxnimo ilaa 4:30 galabnimo Isniinta ilaa Jimcaha."
                    Memo_to_display = Memo_to_display & vbNewLine & ""
                    Memo_to_display = Memo_to_display & vbNewLine & "* Kiiskaaga " & programs & " wuxuu xirmi doonaa " & last_day_of_recert & " haddii *"
                    Memo_to_display = Memo_to_display & vbNewLine & "* aynan helin waraaqahaaga iyo dhamaystirka wareysiga. *"
					'removal of in person verbiage during the COVID-19 PEACETIME STATE OF EMERGENCY
                    ' Memo_to_display = Memo_to_display & vbNewLine & "Haddii aad rabto inaad samaysato ballan wareysi, wac 612-596-1300. Waxa kale oo aad iman kartaa mid ka mid ah lixda xafiis ee hoos ku qoran si loo sameeyo wareysi gof ahaaneed inta u dhexeeya 8 ilaa 4:30, Isniinta ilaa jmcaha."
                    ' Memo_to_display = Memo_to_display & vbNewLine & "- 7051 Brooklyn Blvd Brooklyn Center 55429"
                    ' Memo_to_display = Memo_to_display & vbNewLine & "- 1011 1st St S Hopkins 55343"
                    ' Memo_to_display = Memo_to_display & vbNewLine & "- 9600 Aldrich Ave S Bloomington 55420 Th hrs: 8:30-6:30 "
                    ' Memo_to_display = Memo_to_display & vbNewLine & "- 1001 Plymouth Ave N Minneapolis 55411"
                    ' Memo_to_display = Memo_to_display & vbNewLine & "- 525 Portland Ave S Minneapolis 55415"
                    ' Memo_to_display = Memo_to_display & vbNewLine & "- 2215 East Lake Street Minneapolis 55407"
                    ' Memo_to_display = Memo_to_display & vbNewLine & "(Saacaduhu waa Isniinta - Jimcaha 8-4:30 haddii aan si kale loo sheegin.)"
					Memo_to_display = Memo_to_display & vbNewLine & ""
					Memo_to_display = Memo_to_display & vbNewLine & "You now have an option to use an email to return documents to Hennepin County. Write the case number and full name associated with the case in the body of the email. Only the following types are accepted PNG, JPG, TIFF, DOC, PDF, and HTML. You will not receive confirmation of receipt or failure. To obtain information about your case please contact your worker. EMAIL: hhsews@hennepin.us "
					Memo_to_display = Memo_to_display & vbNewLine & ""
                    Memo_to_display = Memo_to_display & vbNewLine & "Qoraallada rabshadaha qoysaska waxaad ka heli kartaa https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3477-ENG. Waxaad kaloo codsan kartaa qoraalkan oo warqad ah."

                Case "01"   'Spanish (3rd)
                    'MsgBox "SPANISH"
                    CALL convert_date_to_day_first(interview_end_date, day_first_intv_date)
                    CALL convert_date_to_day_first(last_day_of_recert, day_first_last_recert)

                    Memo_to_display = "El Departamento de Servicios Humanos le envió un paquete con papeles. Son los papeles para renovar su caso " & programs & "."
                    Memo_to_display = Memo_to_display & vbNewLine & ""
                    Memo_to_display = Memo_to_display & vbNewLine & "Por favor, fírmelos, coloque la fecha y envíe de regreso los papeles para el 08/" & CM_plus_1_mo & "/" & CM_plus_1_yr & ". También debe realizar una entrevista para que continúe su caso " & programs & "."
                    Memo_to_display = Memo_to_display & vbNewLine & "***Por favor, complete su entrevista para el " & day_first_intv_date & ".***"
                    Memo_to_display = Memo_to_display & vbNewLine & "Para completar una entrevista telefónica, llame a la línea de información EZ al 612-596-1300 entre las 9 am y las 4 pm de lunes a viernes."
                    Memo_to_display = Memo_to_display & vbNewLine & ""
                    Memo_to_display = Memo_to_display & vbNewLine & "**Su caso " & programs & " será cerrado el " & day_first_last_recert & " a menos que recibamos sus papeles y realice la entrevista**"
                    Memo_to_display = Memo_to_display & vbNewLine & ""
					'removal of in person verbiage during the COVID-19 PEACETIME STATE OF EMERGENCY
                    ' Memo_to_display = Memo_to_display & vbNewLine & "Si desea programar una entrevista, llame al 612-596-1300."
                    ' Memo_to_display = Memo_to_display & vbNewLine & "También puede acercarse a cualquiera de las seis oficinas mencionadas debajo para tener una entrevista personal entre las 8 y las 4:30 de lunes a viernes."
                    ' Memo_to_display = Memo_to_display & vbNewLine & "- 7051 Brooklyn Blvd Brooklyn Center 55429"
                    ' Memo_to_display = Memo_to_display & vbNewLine & "- 1011 1st St S Hopkins 55343"
                    ' Memo_to_display = Memo_to_display & vbNewLine & "- 9600 Aldrich Ave S Bloomington 55420 J h.: 8:30-6:30 "
                    ' Memo_to_display = Memo_to_display & vbNewLine & "- 1001 Plymouth Ave N Minneapolis 55411"
                    ' Memo_to_display = Memo_to_display & vbNewLine & "- 525 Portland Ave S Minneapolis 55415"
                    ' Memo_to_display = Memo_to_display & vbNewLine & "- 2215 East Lake Street Minneapolis 55407"
                    ' Memo_to_display = Memo_to_display & vbNewLine & "(Los horarios son de lunes a viernes de 8 a 4:30 a menos   que se remarque lo contrario)"
					Memo_to_display = Memo_to_display & vbNewLine & ""
					Memo_to_display = Memo_to_display & vbNewLine & "You now have an option to use an email to return documents to Hennepin County. Write the case number and full name associated with the case in the body of the email. Only the following types are accepted PNG, JPG, TIFF, DOC, PDF, and HTML. You will not receive confirmation of receipt or failure. To obtain information about your case please contact your worker. EMAIL: hhsews@hennepin.us "
                    Memo_to_display = Memo_to_display & vbNewLine & ""
                    Memo_to_display = Memo_to_display & vbNewLine & "Los folletos de violencia doméstica están disponibles en"
                    Memo_to_display = Memo_to_display & vbNewLine & "https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3477-ENG."
                    Memo_to_display = Memo_to_display & vbNewLine & "También puede solicitar una copia en papel."
                Case "02"   'Hmong (4th)
                    'MsgBox "HMONG"
                    Memo_to_display = "Lub Koos Haum Department of Human Services tau xa ib pob ntawv tuaj rau koj sent. Cov ntawv no yog tuaj tauj koj txoj kev pab " & programs & "."
                    Memo_to_display = Memo_to_display & vbNewLine & ""
                    Memo_to_display = Memo_to_display & vbNewLine & "Thov kos npe, tso sij hawm thiab muaj xa cov ntawv tauj rov qab tuaj ua ntej " & CM_plus_1_mo & "/08/" & CM_plus_1_yr & ". Koj yuav tsum mus xam phaj txog koj cov kev pab " & programs & " mas thiaj li tauj tau."
                    Memo_to_display = Memo_to_display & vbNewLine & "*** Thov mus xam phaj ua ntej " & interview_end_date & ". ***"
                    Memo_to_display = Memo_to_display & vbNewLine & "Yog xam phaj hauv xov tooj, hu rau EZ Info Line ntawm 612-596-1300 thaum 8:00am thib 4:30pm hnub Mon txog Fri."
                    Memo_to_display = Memo_to_display & vbNewLine & "** Koj cov kev pab " & programs & " yuav muab kaw thaum     **"
                    Memo_to_display = Memo_to_display & vbNewLine & "** " & last_day_of_recert & " tsis li mas peb yuav tsum tau txais koj cov**"
                    Memo_to_display = Memo_to_display & vbNewLine & "**      ntaub ntawvthiab koj txoj kev xam phaj.          **"
					'removal of in person verbiage during the COVID-19 PEACETIME STATE OF EMERGENCY
                    ' Memo_to_display = Memo_to_display & vbNewLine & "  Yog hais tias koj xav teem tuaj xam phaj, hu 612-596-1300 Koj kuj tuaj tau rau ib lub ntawm rau lub hoob kas nyob hauv qab no tuaj xam phaj tim ntej muag thaum 8 thiab 4:30, hnub Monday txog Friday."
                    ' Memo_to_display = Memo_to_display & vbNewLine & "- 7051 Brooklyn Blvd Brooklyn Center 55429"
                    ' Memo_to_display = Memo_to_display & vbNewLine & "- 1011 1st St S Hopkins 55343"
                    ' Memo_to_display = Memo_to_display & vbNewLine & "- 9600 Aldrich Ave S Bloomington 55420 Th hrs: 8:30-6:30 "
                    ' Memo_to_display = Memo_to_display & vbNewLine & "- 1001 Plymouth Ave N Minneapolis 55411"
                    ' Memo_to_display = Memo_to_display & vbNewLine & "- 525 Portland Ave S Minneapolis 55415"
                    ' Memo_to_display = Memo_to_display & vbNewLine & "- 2215 East Lake Street Minneapolis 55407"
                    ' Memo_to_display = Memo_to_display & vbNewLine & " (Cov sij hawm qhib yog M - F 8-4:30 tsis li mas yuav tsum qhia ua ntej)"
					Memo_to_display = Memo_to_display & vbNewLine & ""
					Memo_to_display = Memo_to_display & vbNewLine & "You now have an option to use an email to return documents to Hennepin County. Write the case number and full name associated with the case in the body of the email. Only the following types are accepted PNG, JPG, TIFF, DOC, PDF, and HTML. You will not receive confirmation of receipt or failure. To obtain information about your case please contact your worker. EMAIL: hhsews@hennepin.us "
                    Memo_to_display = Memo_to_display & vbNewLine & ""
                    Memo_to_display = Memo_to_display & vbNewLine & "Cov ntaub ntawv qhia txog kev raug tsim txom los ntawm cov txheeb ze kuj muaj nyob rau ntawm"
                    Memo_to_display = Memo_to_display & vbNewLine & "https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3477-ENG."
                    Memo_to_display = Memo_to_display & vbNewLine & "Koj kuj thov tau ib qauv thiab."
                ' Case "06"   'Russian (5th)
                '     'MsgBox "RUSSIAN"
                '     Memo_to_display = "Otdel soczial'ny'x sluzhb otpravil vam paket dokumentaczii."
                '     Memo_to_display = Memo_to_display & vbNewLine & "E'ti dokumenty' dlya obnovleniya vashego " & programs & " dela."
                '     Memo_to_display = Memo_to_display & vbNewLine & ""
                '     Memo_to_display = Memo_to_display & vbNewLine & "Podpishite, ukazhite datu i vernite dokumenty' o prodlenii do " & CM_plus_1_mo & "/08/" & CM_plus_1_yr & ". Vy' takzhe dolzhny' projti sobesedovanie dlya prodleniya svoego " & programs & " dela."
                '     Memo_to_display = Memo_to_display & vbNewLine & ""
                '     Memo_to_display = Memo_to_display & vbNewLine & "*** Pozhalujsta, projdite sobesedovanie do " & interview_end_date & ". ***"
                '     Memo_to_display = Memo_to_display & vbNewLine & "Chtoby' zavershit' sobesedovanie po telefonu, pozvonite v Informaczionnuyu liniyu EZ po telefonu 612-596-1300 s 8:00 do 16:30 s ponedel'nika po pyatniczu."
                '     Memo_to_display = Memo_to_display & vbNewLine & ""
                '     Memo_to_display = Memo_to_display & vbNewLine & "**Vash delo " & programs & " zakroetsya " & last_day_of_recert & ", za**"
                '     Memo_to_display = Memo_to_display & vbNewLine & "** isklyucheniem esli my' poluchim vashi dokumenty'  **"
                '     Memo_to_display = Memo_to_display & vbNewLine & "**          i vy' projdyote sobesedobanie.           **"
                '     Memo_to_display = Memo_to_display & vbNewLine & "   Esli vy' xotite naznachit' sobesedovanie, pozvonite po telefonu 612-596-1300. Vy' takzhe mozhete obratit'sya v lyuboj iz shesti ofisov. Dlya sobesedovanie s 8 i do 4:30, s ponedel'nika po pyatniczu."
                '     Memo_to_display = Memo_to_display & vbNewLine & "- 7051 Brooklyn Blvd Brooklyn Center 55429"
                '     Memo_to_display = Memo_to_display & vbNewLine & "- 1011 1st St S Hopkins 55343"
                '     Memo_to_display = Memo_to_display & vbNewLine & "- 9600 Aldrich Ave S Bloomington 55420 Th hrs: 8:30-6:30 "
                '     Memo_to_display = Memo_to_display & vbNewLine & "- 1001 Plymouth Ave N Minneapolis 55411"
                '     Memo_to_display = Memo_to_display & vbNewLine & "- 525 Portland Ave S Minneapolis 55415"
                '     Memo_to_display = Memo_to_display & vbNewLine & "- 2215 East Lake Street Minneapolis 55407"
                '     Memo_to_display = Memo_to_display & vbNewLine & "(Chasy' priyoma s ponedel'nika po pyatniczu s 8 do 4:30, esli ne ukazano inoe.)"
                '     Memo_to_display = Memo_to_display & vbNewLine & "Broshyupy' o nasilii v sem'e dostupny' po adresu https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3477-ENG Vy' takzhe mozhete zaprosit' bumazhnuyu kopiyu."

                ' Case "12"   'Oromo (6th)
                '     'MsgBox "OROMO"
                ' Case "03"   'Vietnamese (7th)
                '     'MsgBox "VIETNAMESE"
                Case Else  'English (1st)
                    'MsgBox "ENGLISH"
                    Memo_to_display = "The Department of Human Services sent you a packet of paperwork. This paperwork is to renew your " & programs & " case."
                    Memo_to_display = Memo_to_display & vbNewLine & ""
                    Memo_to_display = Memo_to_display & vbNewLine & "Please sign, date and return the renewal paperwork by " & CM_plus_1_mo & "/08/" & CM_plus_1_yr & ". You must also complete an interview for your " & programs & " case to continue."
                    Memo_to_display = Memo_to_display & vbNewLine & ""
                    Memo_to_display = Memo_to_display & vbNewLine & "  *** Please complete your interview by " & interview_end_date & ". ***"
                    Memo_to_display = Memo_to_display & vbNewLine & "To complete a phone interview, call the EZ Info Line at"
                    Memo_to_display = Memo_to_display & vbNewLine & "612-596-1300 between 8:00am and 4:30pm Monday thru Friday."
                    Memo_to_display = Memo_to_display & vbNewLine & ""
                    Memo_to_display = Memo_to_display & vbNewLine & "**  Your " & programs & " case will close on " & last_day_of_recert & " unless    **"
                    Memo_to_display = Memo_to_display & vbNewLine & "** we receive your paperwork and complete the interview. **"
                    Memo_to_display = Memo_to_display & vbNewLine & ""
					'removal of in person verbiage during the COVID-19 PEACETIME STATE OF EMERGENCY
                    ' Memo_to_display = Memo_to_display & vbNewLine & "If you wish to schedule an interview, call 612-596-1300. You may also come to any of the six offices below for an in-person interview between 8 and 4:30, Monday thru Friday."
                    ' Memo_to_display = Memo_to_display & vbNewLine & "- 7051 Brooklyn Blvd Brooklyn Center 55429"
                    ' Memo_to_display = Memo_to_display & vbNewLine & "- 1011 1st St S Hopkins 55343"
                    ' Memo_to_display = Memo_to_display & vbNewLine & "- 9600 Aldrich Ave S Bloomington 55420 Th hrs: 8:30-6:30 "
                    ' Memo_to_display = Memo_to_display & vbNewLine & "- 1001 Plymouth Ave N Minneapolis 55411"
                    ' Memo_to_display = Memo_to_display & vbNewLine & "- 525 Portland Ave S Minneapolis 55415"
                    ' Memo_to_display = Memo_to_display & vbNewLine & "- 2215 East Lake Street Minneapolis 55407"
                    ' Memo_to_display = Memo_to_display & vbNewLine & "(Hours are M - F 8-4:30 unless otherwise noted)"
					' Memo_to_display = Memo_to_display & vbNewLine & ""
					Memo_to_display = Memo_to_display & vbNewLine & "You now have an option to use an email to return documents to Hennepin County. Write the case number and full name associated with the case in the body of the email. Only the following types are accepted PNG, JPG, TIFF, DOC, PDF, and HTML. You will not receive confirmation of receipt or failure. To obtain information about your case please contact your worker. EMAIL: hhsews@hennepin.us "
                    Memo_to_display = Memo_to_display & vbNewLine & " "
                    Memo_to_display = Memo_to_display & vbNewLine & "Domestic violence brochures are available at this website: https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3477-ENG. You can also request a paper copy."

                End Select

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

                'Writing the SPEC MEMO - dates will be input from the determination made earlier.
                Call start_a_new_spec_memo_and_continue(memo_started)

                IF memo_started = True THEN         'The function will return this as FALSE if PF5 does not move past MEMO DISPLAY


                    'Looking for the written language code.
                    'Once we have the memo translated into other languages, the MEMO portion will be put here and will be specific to the language needs.
                    Select Case ALL_CASES_ARRAY(written_lang, case_entry)


                        Case "07"   'Somali (2nd)
                            'MsgBox "SOMALI"
                            CALL write_variable_in_SPEC_MEMO("Waaxda Adeegyada Aadanaha waxay kuu soo dirtay baakad warqado ah. Waraaqahani waxay cusbooneysiiyaan kiiskaaga " & programs & ".")
                            CALL write_variable_in_SPEC_MEMO("")
                            CALL write_variable_in_SPEC_MEMO("Fadlan saxiix, taariikhdana ku qor oo soo celi waraaqaha cusboonaysiinta" & CM_plus_1_mo & "/08/" & CM_plus_1_yr & ". Waa inaad sidoo kale buuxusaa wareysiga " & programs & "-gaaga si kiisku u sii socdo.")
                            'CALL write_variable_in_SPEC_MEMO("")
                            CALL write_variable_in_SPEC_MEMO("*Fadlan dhammaystir wareysigaaga inta ka horreysa " & interview_end_date & "*")
                            CALL write_variable_in_SPEC_MEMO("Si aad u dhamaystirto wareysiga telefoonka, wac laynka taleefanka EZ 612-596-1300 inta u dhaxaysa 8:00 subaxnimo ilaa 4:30 galabnimo Isniinta ilaa Jimcaha.")
                            'CALL write_variable_in_SPEC_MEMO("")
                            CALL write_variable_in_SPEC_MEMO("* Kiiskaaga " & programs & " wuxuu xirmi doonaa " & last_day_of_recert & " haddii *")
                            CALL write_variable_in_SPEC_MEMO("* aynan helin waraaqahaaga iyo dhamaystirka wareysiga. *")
							'removal of in person verbiage during the COVID-19 PEACETIME STATE OF EMERGENCY
                            ' CALL write_variable_in_SPEC_MEMO("Haddii aad rabto inaad samaysato ballan wareysi, wac 612-596-1300. Waxa kale oo aad iman kartaa mid ka mid ah lixda xafiis ee hoos ku qoran si loo sameeyo wareysi gof ahaaneed inta u dhexeeya 8 ilaa 4:30, Isniinta ilaa jmcaha.")
                            ' Call write_variable_in_SPEC_MEMO("- 7051 Brooklyn Blvd Brooklyn Center 55429")
                            ' Call write_variable_in_SPEC_MEMO("- 1011 1st St S Hopkins 55343")
                            ' Call write_variable_in_SPEC_MEMO("- 9600 Aldrich Ave S Bloomington 55420 Th hrs: 8:30-6:30 ")
                            ' Call write_variable_in_SPEC_MEMO("- 1001 Plymouth Ave N Minneapolis 55411")
                            ' Call write_variable_in_SPEC_MEMO("- 525 Portland Ave S Minneapolis 55415")
                            ' Call write_variable_in_SPEC_MEMO("- 2215 East Lake Street Minneapolis 55407")
                            ' CALL write_variable_in_SPEC_MEMO("(Saacaduhu waa Isniinta - Jimcaha 8-4:30 haddii aan si kale loo sheegin.)")
							CALL write_variable_in_SPEC_MEMO("You now have an option to use an email to return documents to Hennepin County. Write the case number and full name associated with the case in the body of the email. Only the following types are accepted PNG, JPG, TIFF, DOC, PDF, and HTML. You will not receive confirmation of receipt or failure. To obtain information about your case please contact your worker. EMAIL: hhsews@hennepin.us ")
							Call write_variable_in_SPEC_MEMO(" ")
                            CALL write_variable_in_SPEC_MEMO("Qoraallada rabshadaha qoysaska waxaad ka heli kartaa https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3477-ENG. Waxaad kaloo codsan kartaa qoraalkan oo warqad ah.")

                        Case "01"   'Spanish (3rd)
                            'MsgBox "SPANISH"
                            CALL convert_date_to_day_first(interview_end_date, day_first_intv_date)
                            CALL convert_date_to_day_first(last_day_of_recert, day_first_last_recert)

                            CALL write_variable_in_SPEC_MEMO("El Departamento de Servicios Humanos le envio un paquete con papeles. Son los papeles para renovar su caso " & programs & ".")
                            CALL write_variable_in_SPEC_MEMO("")
                            CALL write_variable_in_SPEC_MEMO("Por favor, firmelos, coloque la fecha y envie de regreso los papeles para el 08/" & CM_plus_1_mo & "/" & CM_plus_1_yr & ". Tambien debe realizar una entrevista para que continue su caso " & programs & ".")
                            CALL write_variable_in_SPEC_MEMO("***Por favor, complete su entrevista para el " & day_first_intv_date & ".***")
                            CALL write_variable_in_SPEC_MEMO("Para completar una entrevista telefonica, llame a la linea de informacion EZ al 612-596-1300 entre las 9 am y las 4 pm de lunes a viernes.")
                            CALL write_variable_in_SPEC_MEMO("**Su caso " & programs & " sera cerrado el " & day_first_last_recert & " a menos que recibamos sus papeles y realice la entrevista**")
                            CALL write_variable_in_SPEC_MEMO("")
							'removal of in person verbiage during the COVID-19 PEACETIME STATE OF EMERGENCY
                            ' CALL write_variable_in_SPEC_MEMO("Si desea programar una entrevista, llame al 612-596-1300.")
                            ' CALL write_variable_in_SPEC_MEMO("Tambien puede acercarse a cualquiera de las seis oficinas mencionadas debajo para tener una entrevista personal entre las 8 y las 4:30 de lunes a viernes.")
                            ' Call write_variable_in_SPEC_MEMO("- 7051 Brooklyn Blvd Brooklyn Center 55429")
                            ' Call write_variable_in_SPEC_MEMO("- 1011 1st St S Hopkins 55343")
                            ' Call write_variable_in_SPEC_MEMO("- 9600 Aldrich Ave S Bloomington 55420 J h.: 8:30-6:30 ")
                            ' Call write_variable_in_SPEC_MEMO("- 1001 Plymouth Ave N Minneapolis 55411")
                            ' Call write_variable_in_SPEC_MEMO("- 525 Portland Ave S Minneapolis 55415")
                            ' Call write_variable_in_SPEC_MEMO("- 2215 East Lake Street Minneapolis 55407")
                            ' CALL write_variable_in_SPEC_MEMO("(Los horarios son de lunes a viernes de 8 a 4:30 a menos   que se remarque lo contrario)")
                            ' CALL write_variable_in_SPEC_MEMO("")
							CALL write_variable_in_SPEC_MEMO("You now have an option to use an email to return documents to Hennepin County. Write the case number and full name associated with the case in the body of the email. Only the following types are accepted PNG, JPG, TIFF, DOC, PDF, and HTML. You will not receive confirmation of receipt or failure. To obtain information about your case please contact your worker. EMAIL: hhsews@hennepin.us ")
							Call write_variable_in_SPEC_MEMO(" ")
                            CALL write_variable_in_SPEC_MEMO("Los folletos de violencia domestica estan disponibles en")
                            CALL write_variable_in_SPEC_MEMO("https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3477-ENG.")
                            CALL write_variable_in_SPEC_MEMO("Tambien puede solicitar una copia en papel.")

                        Case "02"   'Hmong (4th)
                            'MsgBox "HMONG"
                            CALL write_variable_in_SPEC_MEMO("Lub Koos Haum Department of Human Services tau xa ib pob ntawv tuaj rau koj sent. Cov ntawv no yog tuaj tauj koj txoj kev pab " & programs & ".")
                            'CALL write_variable_in_SPEC_MEMO("")
                            CALL write_variable_in_SPEC_MEMO("Thov kos npe, tso sij hawm thiab muaj xa cov ntawv tauj rov qab tuaj ua ntej " & CM_plus_1_mo & "/08/" & CM_plus_1_yr & ". Koj yuav tsum mus xam phaj txog koj cov kev pab " & programs & " mas thiaj li tauj tau.")
                            CALL write_variable_in_SPEC_MEMO("     *** Thov mus xam phaj ua ntej " & interview_end_date & ". ***")
                            CALL write_variable_in_SPEC_MEMO("Yog xam phaj hauv xov tooj, hu rau EZ Info Line ntawm 612-596-1300 thaum 8:00am thib 4:30pm hnub Mon txog Fri.")
                            CALL write_variable_in_SPEC_MEMO("**    Koj cov kev pab " & programs & " yuav muab kaw thaum     **")
                            CALL write_variable_in_SPEC_MEMO("** " & last_day_of_recert & " tsis li mas peb yuav tsum tau txais koj cov **")
                            CALL write_variable_in_SPEC_MEMO("**      ntaub ntawvthiab koj txoj kev xam phaj.         **")
							'removal of in person verbiage during the COVID-19 PEACETIME STATE OF EMERGENCY
                            ' CALL write_variable_in_SPEC_MEMO("  Yog hais tias koj xav teem tuaj xam phaj, hu 612-596-1300 Koj kuj tuaj tau rau ib lub ntawm rau lub hoob kas nyob hauv qab no tuaj xam phaj tim ntej muag thaum 8 thiab 4:30, hnub Monday txog Friday.")
                            ' Call write_variable_in_SPEC_MEMO("- 7051 Brooklyn Blvd Brooklyn Center 55429")
                            ' Call write_variable_in_SPEC_MEMO("- 1011 1st St S Hopkins 55343")
                            ' Call write_variable_in_SPEC_MEMO("- 9600 Aldrich Ave S Bloomington 55420 Th hrs: 8:30-6:30 ")
                            ' Call write_variable_in_SPEC_MEMO("- 1001 Plymouth Ave N Minneapolis 55411")
                            ' Call write_variable_in_SPEC_MEMO("- 525 Portland Ave S Minneapolis 55415")
                            ' Call write_variable_in_SPEC_MEMO("- 2215 East Lake Street Minneapolis 55407")
                            ' CALL write_variable_in_SPEC_MEMO(" (Cov sij hawm qhib yog M - F 8-4:30 tsis li mas yuav tsum qhia ua ntej)")
                            CALL write_variable_in_SPEC_MEMO("")
							CALL write_variable_in_SPEC_MEMO("You now have an option to use an email to return documents to Hennepin County. Write the case number and full name associated with the case in the body of the email. Only the following types are accepted PNG, JPG, TIFF, DOC, PDF, and HTML. You will not receive confirmation of receipt or failure. To obtain information about your case please contact your worker. EMAIL: hhsews@hennepin.us ")
							Call write_variable_in_SPEC_MEMO(" ")
                            CALL write_variable_in_SPEC_MEMO("Cov ntaub ntawv qhia txog kev raug tsim txom los ntawm cov txheeb ze kuj muaj nyob rau ntawm")
                            CALL write_variable_in_SPEC_MEMO("https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3477-ENG.")
                            CALL write_variable_in_SPEC_MEMO("Koj kuj thov tau ib qauv thiab.")

                        ' Case "06"   'Russian (5th)
                        '     'MsgBox "RUSSIAN"
                        '     CALL write_variable_in_SPEC_MEMO("Otdel soczial'ny'x sluzhb otpravil vam paket dokumentaczii.")
                        '     CALL write_variable_in_SPEC_MEMO("E'ti dokumenty' dlya obnovleniya vashego " & programs & " dela.")
                        '     CALL write_variable_in_SPEC_MEMO("")
                        '     CALL write_variable_in_SPEC_MEMO("Podpishite, ukazhite datu i vernite dokumenty' o prodlenii do " & CM_plus_1_mo & "/08/" & CM_plus_1_yr & ". Vy' takzhe dolzhny' projti sobesedovanie dlya prodleniya svoego " & programs & " dela.")
                        '     CALL write_variable_in_SPEC_MEMO("")
                        '     CALL write_variable_in_SPEC_MEMO("*** Pozhalujsta, projdite sobesedovanie do " & interview_end_date & ". ***")
                        '     CALL write_variable_in_SPEC_MEMO("Chtoby' zavershit' sobesedovanie po telefonu, pozvonite v Informaczionnuyu liniyu EZ po telefonu 612-596-1300 s 8:00 do 16:30 s ponedel'nika po pyatniczu.")
                        '     CALL write_variable_in_SPEC_MEMO("")
                        '     CALL write_variable_in_SPEC_MEMO("**    Vash delo " & programs & " zakroetsya " & last_day_of_recert & ", za    **")
                        '     CALL write_variable_in_SPEC_MEMO("** isklyucheniem esli my' poluchim vashi dokumenty'  **")
                        '     CALL write_variable_in_SPEC_MEMO("**          i vy' projdyote sobesedobanie.           **")
                        '     CALL write_variable_in_SPEC_MEMO("   Esli vy' xotite naznachit' sobesedovanie, pozvonite po telefonu 612-596-1300. Vy' takzhe mozhete obratit'sya v lyuboj iz shesti ofisov. Dlya sobesedovanie s 8 i do 4:30, s ponedel'nika po pyatniczu.")
                        '     Call write_variable_in_SPEC_MEMO("- 7051 Brooklyn Blvd Brooklyn Center 55429")
                        '     Call write_variable_in_SPEC_MEMO("- 1011 1st St S Hopkins 55343")
                        '     Call write_variable_in_SPEC_MEMO("- 9600 Aldrich Ave S Bloomington 55420 Th hrs: 8:30-6:30 ")
                        '     Call write_variable_in_SPEC_MEMO("- 1001 Plymouth Ave N Minneapolis 55411")
                        '     Call write_variable_in_SPEC_MEMO("- 525 Portland Ave S Minneapolis 55415")
                        '     Call write_variable_in_SPEC_MEMO("- 2215 East Lake Street Minneapolis 55407")
                        '     CALL write_variable_in_SPEC_MEMO("(Chasy' priyoma s ponedel'nika po pyatniczu s 8 do 4:30, esli ne ukazano inoe.)")
                        '     CALL write_variable_in_SPEC_MEMO("Broshyupy' o nasilii v sem'e dostupny' po adresu https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3477-ENG Vy' takzhe mozhete zaprosit' bumazhnuyu kopiyu.")
                        ' Case "12"   'Oromo (6th)
                        '     'MsgBox "OROMO"
                        ' Case "03"   'Vietnamese (7th)
                        '     'MsgBox "VIETNAMESE"
                        Case Else  'English (1st)
                            'MsgBox "ENGLISH"
                            CALL write_variable_in_SPEC_MEMO("The Department of Human Services sent you a packet of paperwork. This paperwork is to renew your " & programs & " case.")
                            CALL write_variable_in_SPEC_MEMO("")
                            CALL write_variable_in_SPEC_MEMO("Please sign, date and return the renewal paperwork by " & CM_plus_1_mo & "/08/" & CM_plus_1_yr & ". You must also complete an interview for your " & programs & " case to continue.")
                            CALL write_variable_in_SPEC_MEMO("")
                            Call write_variable_in_SPEC_MEMO("  *** Please complete your interview by " & interview_end_date & ". ***")
                            Call write_variable_in_SPEC_MEMO("To complete a phone interview, call the EZ Info Line at")
                            Call write_variable_in_SPEC_MEMO("612-596-1300 between 8:00am and 4:30pm Monday thru Friday.")
                            CALL write_variable_in_SPEC_MEMO("")
                            CALL write_variable_in_SPEC_MEMO("**  Your " & programs & " case will close on " & last_day_of_recert & " unless    **")
                            CALL write_variable_in_SPEC_MEMO("** we receive your paperwork and complete the interview. **")
                            CALL write_variable_in_SPEC_MEMO("")
							'removal of in person verbiage during the COVID-19 PEACETIME STATE OF EMERGENCY
                            ' Call write_variable_in_SPEC_MEMO("If you wish to schedule an interview, call 612-596-1300. You may also come to any of the six offices below for an in-person interview between 8 and 4:30, Monday thru Friday.")
                            ' Call write_variable_in_SPEC_MEMO("- 7051 Brooklyn Blvd Brooklyn Center 55429")
                            ' Call write_variable_in_SPEC_MEMO("- 1011 1st St S Hopkins 55343")
                            ' Call write_variable_in_SPEC_MEMO("- 9600 Aldrich Ave S Bloomington 55420 Th hrs: 8:30-6:30 ")
                            ' Call write_variable_in_SPEC_MEMO("- 1001 Plymouth Ave N Minneapolis 55411")
                            ' Call write_variable_in_SPEC_MEMO("- 525 Portland Ave S Minneapolis 55415")
                            ' Call write_variable_in_SPEC_MEMO("- 2215 East Lake Street Minneapolis 55407")
							' Call write_variable_in_SPEC_MEMO("(Hours are M - F 8-4:30 unless otherwise noted)")
							' Call write_variable_in_SPEC_MEMO(" ")
							CALL write_variable_in_SPEC_MEMO("You now have an option to use an email to return documents to Hennepin County. Write the case number and full name associated with the case in the body of the email. Only the following types are accepted PNG, JPG, TIFF, DOC, PDF, and HTML. You will not receive confirmation of receipt or failure. To obtain information about your case please contact your worker. EMAIL: hhsews@hennepin.us ")
							Call write_variable_in_SPEC_MEMO(" ")
                            CALL write_variable_in_SPEC_MEMO("Domestic violence brochures are available at this website: https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3477-ENG. You can also request a paper copy.")

                    End Select

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

                    Call start_a_new_spec_memo_and_continue(memo_started)   'Starting a MEMO to send information about verifications

                    IF memo_started = True THEN         'The function will return this as FALSE if PF5 does not move past MEMO DISPLAY

                        Select Case ALL_CASES_ARRAY(written_lang, case_entry)       'Sending notice by language if possible

                        Case "07"   'Somali (2nd)
                            'MsgBox "SOMALI"
                            CALL write_variable_in_SPEC_MEMO("Nidaamka dib-u-cusboonaysiinta waxaa qayb ka ah inaan heno dhammaan xaqiijinta macaluumaadka. Si loo dedejiyo nidaamka dib-u-cusboonaysiinta, fadlan soo raacicaddaymnaha waraaqaha dib-u-cusboonaysiinta.")
                            CALL write_variable_in_SPEC_MEMO("")
                            CALL write_variable_in_SPEC_MEMO("* Tusaalooyinka caddaynta dakhliga: Qaybta dambe ee")
                            CALL write_variable_in_SPEC_MEMO("  jeegaga, qoraalka loo shaqeeyaha, warbixinta dakhliga,")
                            CALL write_variable_in_SPEC_MEMO("  xisaabaadka ganacsiga, foomamka canshuurta dakhliga, iwm.")
                            CALL write_variable_in_SPEC_MEMO("  * Haddii shaqo kaa dhammaatay, soo dir caddeynta")
                            CALL write_variable_in_SPEC_MEMO("    dhamaadka shaqada iyo mushaharka ugu dambeeya.")
                            CALL write_variable_in_SPEC_MEMO("")
                            CALL write_variable_in_SPEC_MEMO("* Tusaalooyinka caddaynta Kharashaadka guryaha (haddii wax")
                            CALL write_variable_in_SPEC_MEMO("  isbeddelay): kirada/guriga rasiidka lacag bixinta,")
                            CALL write_variable_in_SPEC_MEMO("  bixinta, amaah guri, ijaarka, kabitaanka, iwm.")
                            CALL write_variable_in_SPEC_MEMO("")
                            CALL write_variable_in_SPEC_MEMO("* Tusaalooyinka caddaymaha kharashka caafimaadka (haddii")
                            CALL write_variable_in_SPEC_MEMO("  wax isbeddelay): wargadda daawada dhaktarka iyo biilal")
                            CALL write_variable_in_SPEC_MEMO("  caafimaad, iwm.")
                            CALL write_variable_in_SPEC_MEMO("")
							CALL write_variable_in_SPEC_MEMO("You now have an option to use an email to return documents to Hennepin County. Write the case number and full name associated with the case in the body of the email. Only the following types are accepted PNG, JPG, TIFF, DOC, PDF, and HTML. You will not receive confirmation of receipt or failure. To obtain information about your case please contact your worker. EMAIL: hhsews@hennepin.us ")
                            CALL write_variable_in_SPEC_MEMO("Haddii aad qabto su'aalo ku saabsan nooca xaqiijinta loo baahan yahay, wac 612-596-1300 qof ayaa ku caawin doona.")

                        Case "01"   'Spanish (3rd)
                            'MsgBox "SPANISH"
                            CALL write_variable_in_SPEC_MEMO("Como parte del Proceso de Renovacion, debemos recibir una verificacion reciente de su informacion. Para acelerar el proceso de renovacion, por favor, envie pruebas de sus papeles de renovacion.")
                            CALL write_variable_in_SPEC_MEMO("")
                            CALL write_variable_in_SPEC_MEMO("* Ejemplos de pruebas de ingresos: resumenes de pagos,")
                            CALL write_variable_in_SPEC_MEMO("  declaracion del empleador, reportes de ingresos, libros")
                            CALL write_variable_in_SPEC_MEMO("  de contabilidad, formularios de impuestos, etc.")
                            CALL write_variable_in_SPEC_MEMO("  * Si un trabajo se ha terminado, envie pruebas de dicha")
                            CALL write_variable_in_SPEC_MEMO("    situacion y el ultimo pago.")
                            CALL write_variable_in_SPEC_MEMO("")
                            CALL write_variable_in_SPEC_MEMO("* Ejemplos de pruebas de costos de vivienda (si cambio):")
                            CALL write_variable_in_SPEC_MEMO("  recibo de la renta/casa, hipoteca, prestamo, subsidio,")
                            CALL write_variable_in_SPEC_MEMO("  etc.")
                            CALL write_variable_in_SPEC_MEMO("")
                            CALL write_variable_in_SPEC_MEMO("* Ejemplos de pruebas de gastos medicos (si cambio):")
                            CALL write_variable_in_SPEC_MEMO("  prescripciones y cuentas medicas, etc.")
                            CALL write_variable_in_SPEC_MEMO("")
							CALL write_variable_in_SPEC_MEMO("You now have an option to use an email to return documents to Hennepin County. Write the case number and full name associated with the case in the body of the email. Only the following types are accepted PNG, JPG, TIFF, DOC, PDF, and HTML. You will not receive confirmation of receipt or failure. To obtain information about your case please contact your worker. EMAIL: hhsews@hennepin.us ")
                            CALL write_variable_in_SPEC_MEMO("Si tiene preguntas sobre el tipo de verificacion necesaria, llame al 612-596-1300 y alguien lo/la asistira.")
                        Case "02"   'Hmong (4th)
                            'MsgBox "HMONG"
                            CALL write_variable_in_SPEC_MEMO("Raws li peb txoj kev Rov Tauj Dua mas peb yuav tsum tau txais cov xov tseem ceeb los ntawm koj. Yuav kom tauj tau sai, thov xa cov pov thawj nrog koj ntaub ntawv tauj dua tshiab.")
                            CALL write_variable_in_SPEC_MEMO("")
                            CALL write_variable_in_SPEC_MEMO("* Piv txwv pov thawj txog nyiaj txiag: cov tw tshev, ntawv")
                            CALL write_variable_in_SPEC_MEMO("  tom chaw ua hauj lwm, ntawv qhia txog nyiaj txiag, ntawv")
                            CALL write_variable_in_SPEC_MEMO("  ua lag luam, ntawv ua se, lwm yam.")
                            CALL write_variable_in_SPEC_MEMO("  *Yog hais tias koj txoj hauj lwm tu lawm, xa pav thawj")
                            CALL write_variable_in_SPEC_MEMO("   txog hnub kawg ua hauj lwm thiab daim tshev uas yog daim")
                            CALL write_variable_in_SPEC_MEMO("   kawg.")
                            CALL write_variable_in_SPEC_MEMO("")
                            CALL write_variable_in_SPEC_MEMO("* Piv txwv cov pov thawj them nqi tsev(yog hais tias")
                            CALL write_variable_in_SPEC_MEMO("  hloov): pov thawj xauj tsev/them tsev, ntawv them tuam")
                            CALL write_variable_in_SPEC_MEMO("  txhab qiv nyiaj yuav tsev, ntawv cog lus xauj tsev, ntawv")
                            CALL write_variable_in_SPEC_MEMO("  them tsev luam, lwm yam.")
                            CALL write_variable_in_SPEC_MEMO("")
                            CALL write_variable_in_SPEC_MEMO("* Piv txwv cov pov thawj txog nqi kho mob(yog hais tias")
                            CALL write_variable_in_SPEC_MEMO("  hloov lawm): Ntawv yuav tshuaj thiab nqi kho mob, lwm")
                            CALL write_variable_in_SPEC_MEMO("  yam.")
                            CALL write_variable_in_SPEC_MEMO("")
							CALL write_variable_in_SPEC_MEMO("You now have an option to use an email to return documents to Hennepin County. Write the case number and full name associated with the case in the body of the email. Only the following types are accepted PNG, JPG, TIFF, DOC, PDF, and HTML. You will not receive confirmation of receipt or failure. To obtain information about your case please contact your worker. EMAIL: hhsews@hennepin.us ")
                            CALL write_variable_in_SPEC_MEMO("Yog hais tias koj muaj lus nug txog cov yuav tsum muaj cov pov thaqwj twg, hu 612-596-1300 ces neeg mam los pab koj.")
                        '
                        ' Case "06"   'Russian (5th)
                        '     'MsgBox "RUSSIAN"
                        '     CALL write_variable_in_SPEC_MEMO("V czelyax obnovleniya proczessa my' dolzhny' poluchit' podtverzhdenie vashej unformaczii.  Chtoby' uskorit' proczess obnovlenie, pozhalujsta, otprav'te dokazatel'stva s vashej dokumentacziej na obnovlenie.")
                        '     CALL write_variable_in_SPEC_MEMO("")
                        '     CALL write_variable_in_SPEC_MEMO("* Primery' dokazatel'stv doxoda: koreshki chekov,")
                        '     CALL write_variable_in_SPEC_MEMO("  zayavlenie rabotodatelya, otchety' o doxodax,")
                        '     CALL write_variable_in_SPEC_MEMO("  buxgalterskie knigi, formy' podoxodnogo naloga i t.d.")
                        '     CALL write_variable_in_SPEC_MEMO("  * Esli vy' prekratili rabotat', otprav'te podtberzhdenie")
                        '     CALL write_variable_in_SPEC_MEMO("    o prekrashhenii raboty' i poslednyuyu oplatu.")
                        '     CALL write_variable_in_SPEC_MEMO("")
                        '     CALL write_variable_in_SPEC_MEMO("* Primery' dokazatel'stv stoimosti zhil'ya (esli oni")
                        '     CALL write_variable_in_SPEC_MEMO("  ezmeneny'): arenda/dom kvitancziya ob oplate, ipoteka,")
                        '     CALL write_variable_in_SPEC_MEMO("  arenda, subsidiya i t.d.")
                        '     CALL write_variable_in_SPEC_MEMO("")
                        '     CALL write_variable_in_SPEC_MEMO("* Primery' dokazatel'ctv mediczinskix rassxodov (esli oni")
                        '     CALL write_variable_in_SPEC_MEMO("  izmeneny'): oplata za lekarstva i medeczinskie scheta i")
                        '     CALL write_variable_in_SPEC_MEMO("  t. d.")
                        '     CALL write_variable_in_SPEC_MEMO("")
                        '     CALL write_variable_in_SPEC_MEMO("Esli u vas est' voprosy' o tipe dokazatel'stv pozvonite po telefonu 612-596-1300, u kto-to pomozhet vam.")
                        ' Case "12"   'Oromo (6th)
                        '     'MsgBox "OROMO"
                        ' Case "03"   'Vietnamese (7th)
                        '     'MsgBox "VIETNAMESE"
                        Case Else  'English (1st)
                            'MsgBox "ENGLISH"
                            CALL write_variable_in_SPEC_MEMO("As a part of the Renewal Process we must receive recent verification of your information. To speed the renewal process, please send proofs with your renewal paperwork.")
                            CALL write_variable_in_SPEC_MEMO("")
                            CALL write_variable_in_SPEC_MEMO(" * Examples of income proofs: paystubs, employer statement,")
                            CALL write_variable_in_SPEC_MEMO("   income reports, business ledgers, income tax forms, etc.")
                            CALL write_variable_in_SPEC_MEMO("   *If a job has ended, send proof of the end of employment")
                            CALL write_variable_in_SPEC_MEMO("   and last pay.")
                            CALL write_variable_in_SPEC_MEMO("")
                            CALL write_variable_in_SPEC_MEMO(" * Examples of housing cost proofs(if changed): rent/house")
                            CALL write_variable_in_SPEC_MEMO("   payment receipt, mortgage, lease, subsidy, etc.")
                            CALL write_variable_in_SPEC_MEMO("")
                            CALL write_variable_in_SPEC_MEMO(" * Examples of medical cost proofs(if changed):")
                            CALL write_variable_in_SPEC_MEMO("   prescription and medical bills, etc.")
                            CALL write_variable_in_SPEC_MEMO("")
							CALL write_variable_in_SPEC_MEMO("You now have an option to use an email to return documents to Hennepin County. Write the case number and full name associated with the case in the body of the email. Only the following types are accepted PNG, JPG, TIFF, DOC, PDF, and HTML. You will not receive confirmation of receipt or failure. To obtain information about your case please contact your worker. EMAIL: hhsews@hennepin.us ")
                            CALL write_variable_in_SPEC_MEMO("If you have questions about the type of verifications needed, call 612-596-1300 and someone will assist you.")

                        End Select

                        PF4 'Submit the MEMO'

                    End If

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

                    PF3
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

                        Case "07"   'Somali (2nd)
                            'MsgBox "SOMALI"
                            if ALL_CASES_ARRAY(recvd_appl, case_entry) = TRUE then Memo_to_display = "Waraaqahaagii dib-u-cusboonaysiinta waxaan helnay" & ALL_CASES_ARRAY(date_of_app, case_entry) & "."
                            if ALL_CASES_ARRAY(recvd_appl, case_entry) = FALSE then Memo_to_display = "Waraaqahaagii dib-u-cusboonaysiinta weli ma aynaan helin."
                            Memo_to_display = Memo_to_display & vbNewLine & ""
                            Memo_to_display = Memo_to_display & vbNewLine & "Waa inaad wareysi martaa inta ka horreysa " & last_day_of_recert & " haddii kale waxaa joogsan doona waxtarrada aad hesho."
                            Memo_to_display = Memo_to_display & vbNewLine & ""
                            Memo_to_display = Memo_to_display & vbNewLine & "Si aad u dhamaystirto wareysiga telefoonka, wac laynka taleefanka EZ 612-596-1300 inta u dhaxaysa 8:00 subaxnimo ilaa 4:30 galabnimo Isniinta ilaa Jimcaha."
                            Memo_to_display = Memo_to_display & vbNewLine & ""
							'removal of in person verbiage during the COVID-19 PEACETIME STATE OF EMERGENCY
                            ' Memo_to_display = Memo_to_display & vbNewLine & "Haddii aad rabto inaad samaysato ballan wareysi, wac 612-596-1300. Waxa kale oo aad iman kartaa mid ka mid ah lixda xafiis ee hoos ku qoran si loo sameeyo wareysi gof ahaaneed inta u dhexeeya 8 ilaa 4:30, Isniinta ilaa jmcaha."
                            ' Memo_to_display = Memo_to_display & vbNewLine & "- 7051 Brooklyn Blvd Brooklyn Center 55429"
                            ' Memo_to_display = Memo_to_display & vbNewLine & "- 1011 1st St S Hopkins 55343"
                            ' Memo_to_display = Memo_to_display & vbNewLine & "- 9600 Aldrich Ave S Bloomington 55420 Th hrs: 8:30-6:30 "
                            ' Memo_to_display = Memo_to_display & vbNewLine & "- 1001 Plymouth Ave N Minneapolis 55411"
                            ' Memo_to_display = Memo_to_display & vbNewLine & "- 525 Portland Ave S Minneapolis 55415"
                            ' Memo_to_display = Memo_to_display & vbNewLine & "- 2215 East Lake Street Minneapolis 55407"
                            ' Memo_to_display = Memo_to_display & vbNewLine & "(Saacaduhu waa Isniinta - Jimcaha 8-4:30 haddii aan si kale loo sheegin.)"
							Memo_to_display = Memo_to_display & vbNewLine & ""
							Memo_to_display = Memo_to_display & vbNewLine & "You now have an option to use an email to return documents to Hennepin County. Write the case number and full name associated with the case in the body of the email. Only the following types are accepted PNG, JPG, TIFF, DOC, PDF, and HTML. You will not receive confirmation of receipt or failure. To obtain information about your case please contact your worker. EMAIL: hhsews@hennepin.us "
                            Memo_to_display = Memo_to_display & vbNewLine & ""
                            Memo_to_display = Memo_to_display & vbNewLine & "* Haddii aynaan war kaa helin inta ka horreysa " & last_day_of_recert & " *"
                            Memo_to_display = Memo_to_display & vbNewLine & "*   Macaawinada aad hesho waxay instaageysaa " & last_day_of_recert & ". *"

                        Case "01"   'Spanish (3rd)
                            'MsgBox "SPANISH"
                            CALL convert_date_to_day_first(ALL_CASES_ARRAY(date_of_app, case_entry), day_first_app_date)
                            CALL convert_date_to_day_first(last_day_of_recert, day_first_last_recert)

                            if ALL_CASES_ARRAY(recvd_appl, case_entry) = TRUE then Memo_to_display = "Recibimos sus papeles de recertificación el " & day_first_app_date & "."
                            if ALL_CASES_ARRAY(recvd_appl, case_entry) = FALSE then Memo_to_display = "Aún no se han recibido sus Papeles de Recertificación."
                            Memo_to_display = Memo_to_display & vbNewLine & ""
                            Memo_to_display = Memo_to_display & vbNewLine & "Debe realizar una entrevista para el " & day_first_last_recert & " o sus beneficios se terminarán."
                            Memo_to_display = Memo_to_display & vbNewLine & ""
                            Memo_to_display = Memo_to_display & vbNewLine & "Para completar una entrevista telefónica, llame a la línea de información EZ al 612-596-1300 entre las 9 am y las 4 pm de lunes a viernes."
                            Memo_to_display = Memo_to_display & vbNewLine & ""
							'removal of in person verbiage during the COVID-19 PEACETIME STATE OF EMERGENCY
                            ' Memo_to_display = Memo_to_display & vbNewLine & "Si desea programar una entrevista, llame al 612-596-1300."
                            ' Memo_to_display = Memo_to_display & vbNewLine & "También puede acercarse a cualquiera de las seis oficinas mencionadas debajo para tener una entrevista personal entre las 8 y las 4:30 de lunes a viernes."
                            ' Memo_to_display = Memo_to_display & vbNewLine & "- 7051 Brooklyn Blvd Brooklyn Center 55429"
                            ' Memo_to_display = Memo_to_display & vbNewLine & "- 1011 1st St S Hopkins 55343"
                            ' Memo_to_display = Memo_to_display & vbNewLine & "- 9600 Aldrich Ave S Bloomington 55420 J h.: 8:30-6:30 "
                            ' Memo_to_display = Memo_to_display & vbNewLine & "- 1001 Plymouth Ave N Minneapolis 55411"
                            ' Memo_to_display = Memo_to_display & vbNewLine & "- 525 Portland Ave S Minneapolis 55415"
                            ' Memo_to_display = Memo_to_display & vbNewLine & "- 2215 East Lake Street Minneapolis 55407"
                            ' Memo_to_display = Memo_to_display & vbNewLine & "(Los horarios son de lunes a viernes de 8 a 4:30 a menos   que se remarque lo contrario)"
							Memo_to_display = Memo_to_display & vbNewLine & ""
							Memo_to_display = Memo_to_display & vbNewLine & "You now have an option to use an email to return documents to Hennepin County. Write the case number and full name associated with the case in the body of the email. Only the following types are accepted PNG, JPG, TIFF, DOC, PDF, and HTML. You will not receive confirmation of receipt or failure. To obtain information about your case please contact your worker. EMAIL: hhsews@hennepin.us "
                            Memo_to_display = Memo_to_display & vbNewLine & ""
                            Memo_to_display = Memo_to_display & vbNewLine & "**Si no tenemos novedades suyas para el " & day_first_last_recert & ", sus beneficios se terminarán el " & day_first_last_recert & "**"

                        Case "02"   'Hmong (4th)
                            'MsgBox "HMONG"
                            if ALL_CASES_ARRAY(recvd_appl, case_entry) = TRUE then Memo_to_display = "Peb twb txais tau koj cov Ntaub Ntawv Rov Qab Tauj Dua thaum " & ALL_CASES_ARRAY(date_of_app, case_entry) & "."
                            if ALL_CASES_ARRAY(recvd_appl, case_entry) = FALSE then Memo_to_display = "Peb tsis tau txais koj cov Ntaub Ntawv Rov Qab Tauj Duu."
                            Memo_to_display = Memo_to_display & vbNewLine & ""
                            Memo_to_display = Memo_to_display & vbNewLine & "Koj yuav tsum mus xam pphaj ua ntej " & last_day_of_recert & " los yog yuav txiav koj cov kev pab."
                            Memo_to_display = Memo_to_display & vbNewLine & ""
                            Memo_to_display = Memo_to_display & vbNewLine & "Yog xam phaj hauv xov tooj, hu rau EZ Info Line ntawm 612-596-1300 thaum 8:00am thib 4:30pm hnub Monday txog Friday."
							'removal of in person verbiage during the COVID-19 PEACETIME STATE OF EMERGENCY
                            ' Memo_to_display = Memo_to_display & vbNewLine & "  Yog hais tias koj xav teem tuaj xam phaj, hu 612-596-1300 Koj kuj tuaj tau rau ib lub ntawm rau lub hoob kas nyob hauv qab no tuaj xam phaj tim ntej muag thaum 8 thiab 4:30, hnub Monday txog Friday."
                            ' Memo_to_display = Memo_to_display & vbNewLine & "- 7051 Brooklyn Blvd Brooklyn Center 55429"
                            ' Memo_to_display = Memo_to_display & vbNewLine & "- 1011 1st St S Hopkins 55343"
                            ' Memo_to_display = Memo_to_display & vbNewLine & "- 9600 Aldrich Ave S Bloomington 55420 Th hrs: 8:30-6:30 "
                            ' Memo_to_display = Memo_to_display & vbNewLine & "- 1001 Plymouth Ave N Minneapolis 55411"
                            ' Memo_to_display = Memo_to_display & vbNewLine & "- 525 Portland Ave S Minneapolis 55415"
                            ' Memo_to_display = Memo_to_display & vbNewLine & "- 2215 East Lake Street Minneapolis 55407"
                            ' Memo_to_display = Memo_to_display & vbNewLine & " (Cov sij hawm qhib yog M - F 8-4:30 tsis li mas yuav tsum qhia ua ntej)"
							Memo_to_display = Memo_to_display & vbNewLine & ""
							Memo_to_display = Memo_to_display & vbNewLine & "You now have an option to use an email to return documents to Hennepin County. Write the case number and full name associated with the case in the body of the email. Only the following types are accepted PNG, JPG, TIFF, DOC, PDF, and HTML. You will not receive confirmation of receipt or failure. To obtain information about your case please contact your worker. EMAIL: hhsews@hennepin.us "
                            Memo_to_display = Memo_to_display & vbNewLine & ""
                            Memo_to_display = Memo_to_display & vbNewLine & "** Yog hais tias tsis hnov koj teb ua ntej " & last_day_of_recert & "  **"
                            Memo_to_display = Memo_to_display & vbNewLine & "**   koj cov kev pab yuav raug kaw thaum " & last_day_of_recert & ".   **"
                        ' Case "06"   'Russian (5th)
                        '     'MsgBox "RUSSIAN"
                        '     if ALL_CASES_ARRAY(recvd_appl, case_entry) = TRUE then Memo_to_display = "My' poluchili vashu dokumentacziyu o pereodicheskoj attestaczii " & ALL_CASES_ARRAY(date_of_app, case_entry) & "."
                        '     if ALL_CASES_ARRAY(recvd_appl, case_entry) = FALSE then Memo_to_display = "Vasha dokumentacziya o pereodicheskoj attestaczii eshhyo ne poluchena."
                        '     Memo_to_display = Memo_to_display & vbNewLine & ""
                        '     Memo_to_display = Memo_to_display & vbNewLine & "Vy' dolzhny' projti sobesedovanie do " & last_day_of_recert & " ili vasha programma zakroetsya."
                        '     Memo_to_display = Memo_to_display & vbNewLine & ""
                        '     Memo_to_display = Memo_to_display & vbNewLine & "Chtoby' projti sobesedovanie po telefonu, pozvonite v Informaczionnuyu liniyu EZ po telefonu 612-596-1300 s 8:00 do 16:30 s ponedel'nika po pyatniczu."
                        '     Memo_to_display = Memo_to_display & vbNewLine & ""
                        '     Memo_to_display = Memo_to_display & vbNewLine & "   Esli vy' xotite naznachit' sobesedovanie, pozvonite po telefonu 612-596-1300. Vy' takzhe mozhete obratit'sya v lyuboj iz shesti ofisov. Dlya sobesedovanie s 8 i do 4:30, s ponedel'nika po pyatniczu."
                        '     Memo_to_display = Memo_to_display & vbNewLine & "- 7051 Brooklyn Blvd Brooklyn Center 55429"
                        '     Memo_to_display = Memo_to_display & vbNewLine & "- 1011 1st St S Hopkins 55343"
                        '     Memo_to_display = Memo_to_display & vbNewLine & "- 9600 Aldrich Ave S Bloomington 55420 Th hrs: 8:30-6:30 "
                        '     Memo_to_display = Memo_to_display & vbNewLine & "- 1001 Plymouth Ave N Minneapolis 55411"
                        '     Memo_to_display = Memo_to_display & vbNewLine & "- 525 Portland Ave S Minneapolis 55415"
                        '     Memo_to_display = Memo_to_display & vbNewLine & "- 2215 East Lake Street Minneapolis 55407"
                        '     Memo_to_display = Memo_to_display & vbNewLine & "(Chasy' priyoma s ponedel'nika po pyatniczu s 8 do 4:30, esli ne ukazano inoe.)"
                        '     Memo_to_display = Memo_to_display & vbNewLine & ""
                        '     Memo_to_display = Memo_to_display & vbNewLine & "** Esli my' ne usly'shim ot vas do " & last_day_of_recert & " **"
                        '     Memo_to_display = Memo_to_display & vbNewLine & "**   vasha programma zakroetsya " & last_day_of_recert & "    **"
                        ' Case "12"   'Oromo (6th)
                        '     'MsgBox "OROMO"
                        ' Case "03"   'Vietnamese (7th)
                        '     'MsgBox "VIETNAMESE"
                        Case Else  'English (1st)
                            'MsgBox "ENGLISH"
                            'creating the memo message and displaying it.
                            if ALL_CASES_ARRAY(recvd_appl, case_entry) = TRUE then Memo_to_display = "We received your Recertification Paperwork on " & ALL_CASES_ARRAY(date_of_app, case_entry) & "."
                            if ALL_CASES_ARRAY(recvd_appl, case_entry) = FALSE then Memo_to_display = "Your Recertification Paperwork has not yet been received."
                            Memo_to_display = Memo_to_display & vbNewLine & vbNewLine & ""
                            Memo_to_display = Memo_to_display & vbNewLine & vbNewLine & "You must have an interview by " & last_day_of_recert & " or your benefits will end. "
                            Memo_to_display = Memo_to_display & vbNewLine & vbNewLine & ""
                            Memo_to_display = Memo_to_display & vbNewLine & vbNewLine & "To complete a phone interview, call the EZ Info Line at"
                            Memo_to_display = Memo_to_display & vbNewLine & vbNewLine & "612-596-1300 between 8:00am and 4:30pm Monday thru Friday."
                            Memo_to_display = Memo_to_display & vbNewLine & vbNewLine & ""
							'removal of in person verbiage during the COVID-19 PEACETIME STATE OF EMERGENCY
                            ' Memo_to_display = Memo_to_display & vbNewLine & vbNewLine & "If you wish to schedule an interview, call 612-596-1300. You may also come to any of the six offices below for an in-person interview between 8 and 4:30, Monday thru Friday."
                            ' Memo_to_display = Memo_to_display & vbNewLine & vbNewLine & "- 7051 Brooklyn Blvd Brooklyn Center 55429"
                            ' Memo_to_display = Memo_to_display & vbNewLine & vbNewLine & "- 1011 1st St S Hopkins 55343"
                            ' Memo_to_display = Memo_to_display & vbNewLine & vbNewLine & "- 9600 Aldrich Ave S Bloomington 55420 Th hrs: 8:30-6:30 "
                            ' Memo_to_display = Memo_to_display & vbNewLine & vbNewLine & "- 1001 Plymouth Ave N Minneapolis 55411"
                            ' Memo_to_display = Memo_to_display & vbNewLine & vbNewLine & "- 525 Portland Ave S Minneapolis 55415"
                            ' Memo_to_display = Memo_to_display & vbNewLine & vbNewLine & "- 2215 East Lake Street Minneapolis 55407"
                            ' Memo_to_display = Memo_to_display & vbNewLine & vbNewLine & "(Hours are M - F 8-4:30 unless otherwise noted)"
							Memo_to_display = Memo_to_display & vbNewLine & vbNewLine & ""
							Memo_to_display = Memo_to_display & vbNewLine & vbNewLine & "You now have an option to use an email to return documents to Hennepin County. Write the case number and full name associated with the case in the body of the email. Only the following types are accepted PNG, JPG, TIFF, DOC, PDF, and HTML. You will not receive confirmation of receipt or failure. To obtain information about your case please contact your worker. EMAIL: hhsews@hennepin.us "
                            Memo_to_display = Memo_to_display & vbNewLine & vbNewLine & ""
                            Memo_to_display = Memo_to_display & vbNewLine & vbNewLine & "  ** If we do not hear from you by " & last_day_of_recert & "  **"
                            Memo_to_display = Memo_to_display & vbNewLine & vbNewLine & "  **   your benefits will end on " & last_day_of_recert & ".   **"

                        End Select

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
                        'writing a SPEC MEMO with the NOMI wording.
                        Call start_a_new_spec_memo_and_continue(memo_started)
                        'MsgBox memo_started
                        IF memo_started = True THEN         'The function will return this as FALSE if PF5 does not move past MEMO DISPLAY

                            Select Case ALL_CASES_ARRAY(written_lang, case_entry)       'selecting  the language and writing the memo by language

                                Case "07"   'Somali (2nd)
                                    'MsgBox "SOMALI"
                                    if ALL_CASES_ARRAY(recvd_appl, case_entry) = TRUE then CALL write_variable_in_SPEC_MEMO("Waraaqahaagii dib-u-cusboonaysiinta waxaan helnay" & ALL_CASES_ARRAY(date_of_app, case_entry) & ".")
                                    if ALL_CASES_ARRAY(recvd_appl, case_entry) = FALSE then CALL write_variable_in_SPEC_MEMO("Waraaqahaagii dib-u-cusboonaysiinta weli ma aynaan helin.")
                                    CALL write_variable_in_SPEC_MEMO("")
                                    CALL write_variable_in_SPEC_MEMO("Waa inaad wareysi martaa inta ka horreysa " & last_day_of_recert & " haddii kale waxaa joogsan doona waxtarrada aad hesho.")
                                    CALL write_variable_in_SPEC_MEMO("")
                                    CALL write_variable_in_SPEC_MEMO("Si aad u dhamaystirto wareysiga telefoonka, wac laynka taleefanka EZ 612-596-1300 inta u dhaxaysa 8:00 subaxnimo ilaa 4:30 galabnimo Isniinta ilaa Jimcaha.")
                                    CALL write_variable_in_SPEC_MEMO("")
									'removal of in person verbiage during the COVID-19 PEACETIME STATE OF EMERGENCY
                                    ' CALL write_variable_in_SPEC_MEMO("Haddii aad rabto inaad samaysato ballan wareysi, wac 612-596-1300. Waxa kale oo aad iman kartaa mid ka mid ah lixda xafiis ee hoos ku qoran si loo sameeyo wareysi gof ahaaneed inta u dhexeeya 8 ilaa 4:30, Isniinta ilaa jmcaha.")
                                    ' Call write_variable_in_SPEC_MEMO("- 7051 Brooklyn Blvd Brooklyn Center 55429")
                                    ' Call write_variable_in_SPEC_MEMO("- 1011 1st St S Hopkins 55343")
                                    ' Call write_variable_in_SPEC_MEMO("- 9600 Aldrich Ave S Bloomington 55420 Th hrs: 8:30-6:30 ")
                                    ' Call write_variable_in_SPEC_MEMO("- 1001 Plymouth Ave N Minneapolis 55411")
                                    ' Call write_variable_in_SPEC_MEMO("- 525 Portland Ave S Minneapolis 55415")
                                    ' Call write_variable_in_SPEC_MEMO("- 2215 East Lake Street Minneapolis 55407")
                                    ' CALL write_variable_in_SPEC_MEMO("(Saacaduhu waa Isniinta - Jimcaha 8-4:30 haddii aan si kale loo sheegin.)")
									' CALL write_variable_in_SPEC_MEMO("")
									CALL write_variable_in_SPEC_MEMO("You now have an option to use an email to return documents to Hennepin County. Write the case number and full name associated with the case in the body of the email. Only the following types are accepted PNG, JPG, TIFF, DOC, PDF, and HTML. You will not receive confirmation of receipt or failure. To obtain information about your case please contact your worker. EMAIL: hhsews@hennepin.us ")
									Call write_variable_in_SPEC_MEMO(" ")
                                    CALL write_variable_in_SPEC_MEMO("* Haddii aynaan war kaa helin inta ka horreysa " & last_day_of_recert & " *")
                                    CALL write_variable_in_SPEC_MEMO("*   Macaawinada aad hesho waxay instaageysaa " & last_day_of_recert & ".  *")

                                Case "01"   'Spanish (3rd)
                                    'MsgBox "SPANISH"
                                    If ALL_CASES_ARRAY(date_of_app, case_entry) <> "" Then CALL convert_date_to_day_first(ALL_CASES_ARRAY(date_of_app, case_entry), day_first_app_date)
                                    CALL convert_date_to_day_first(last_day_of_recert, day_first_last_recert)

                                    if ALL_CASES_ARRAY(recvd_appl, case_entry) = TRUE then CALL write_variable_in_SPEC_MEMO("Recibimos sus papeles de recertificacion el " & day_first_app_date & ".")
                                    if ALL_CASES_ARRAY(recvd_appl, case_entry) = FALSE then CALL write_variable_in_SPEC_MEMO("Aun no se han recibido sus Papeles de Recertificacion.")
                                    CALL write_variable_in_SPEC_MEMO("")
                                    CALL write_variable_in_SPEC_MEMO("Debe realizar una entrevista para el " & day_first_last_recert & " o sus beneficios se terminaran.")
                                    CALL write_variable_in_SPEC_MEMO("")
                                    CALL write_variable_in_SPEC_MEMO("Para completar una entrevista telefonica, llame a la linea de informacion EZ al 612-596-1300 entre las 9 am y las 4 pm de lunes a viernes.")
                                    CALL write_variable_in_SPEC_MEMO("")
									'removal of in person verbiage during the COVID-19 PEACETIME STATE OF EMERGENCY
                                    ' CALL write_variable_in_SPEC_MEMO("Si desea programar una entrevista, llame al 612-596-1300.")
                                    ' CALL write_variable_in_SPEC_MEMO("Tambien puede acercarse a cualquiera de las seis oficinas mencionadas debajo para tener una entrevista personal entre las 8 y las 4:30 de lunes a viernes.")
                                    ' Call write_variable_in_SPEC_MEMO("- 7051 Brooklyn Blvd Brooklyn Center 55429")
                                    ' Call write_variable_in_SPEC_MEMO("- 1011 1st St S Hopkins 55343")
                                    ' Call write_variable_in_SPEC_MEMO("- 9600 Aldrich Ave S Bloomington 55420 J h.: 8:30-6:30 ")
                                    ' Call write_variable_in_SPEC_MEMO("- 1001 Plymouth Ave N Minneapolis 55411")
                                    ' Call write_variable_in_SPEC_MEMO("- 525 Portland Ave S Minneapolis 55415")
                                    ' Call write_variable_in_SPEC_MEMO("- 2215 East Lake Street Minneapolis 55407")
                                    ' CALL write_variable_in_SPEC_MEMO("(Los horarios son de lunes a viernes de 8 a 4:30 a menos   que se remarque lo contrario)")
                                    ' CALL write_variable_in_SPEC_MEMO("")
									CALL write_variable_in_SPEC_MEMO("You now have an option to use an email to return documents to Hennepin County. Write the case number and full name associated with the case in the body of the email. Only the following types are accepted PNG, JPG, TIFF, DOC, PDF, and HTML. You will not receive confirmation of receipt or failure. To obtain information about your case please contact your worker. EMAIL: hhsews@hennepin.us ")
									Call write_variable_in_SPEC_MEMO(" ")
                                    CALL write_variable_in_SPEC_MEMO("**Si no tenemos novedades suyas para el " & day_first_last_recert & ", sus beneficios se terminaran el " & day_first_last_recert & "**")

                                Case "02"   'Hmong (4th)
                                    'MsgBox "HMONG"
                                    if ALL_CASES_ARRAY(recvd_appl, case_entry) = TRUE then CALL write_variable_in_SPEC_MEMO("Peb twb txais tau koj cov Ntaub Ntawv Rov Qab Tauj Dua thaum " & ALL_CASES_ARRAY(date_of_app, case_entry) & ".")
                                    if ALL_CASES_ARRAY(recvd_appl, case_entry) = FALSE then CALL write_variable_in_SPEC_MEMO("Peb tsis tau txais koj cov Ntaub Ntawv Rov Qab Tauj Duu.")
                                    CALL write_variable_in_SPEC_MEMO("")
                                    CALL write_variable_in_SPEC_MEMO("Koj yuav tsum mus xam pphaj ua ntej " & last_day_of_recert & " los yog yuav txiav koj cov kev pab.")
                                    CALL write_variable_in_SPEC_MEMO("")
                                    CALL write_variable_in_SPEC_MEMO("Yog xam phaj hauv xov tooj, hu rau EZ Info Line ntawm 612-596-1300 thaum 8:00am thib 4:30pm hnub Monday txog Friday.")
									'removal of in person verbiage during the COVID-19 PEACETIME STATE OF EMERGENCY
                                    ' CALL write_variable_in_SPEC_MEMO("  Yog hais tias koj xav teem tuaj xam phaj, hu 612-596-1300 Koj kuj tuaj tau rau ib lub ntawm rau lub hoob kas nyob hauv qab no tuaj xam phaj tim ntej muag thaum 8 thiab 4:30, hnub Monday txog Friday.")
                                    ' Call write_variable_in_SPEC_MEMO("- 7051 Brooklyn Blvd Brooklyn Center 55429")
                                    ' Call write_variable_in_SPEC_MEMO("- 1011 1st St S Hopkins 55343")
                                    ' Call write_variable_in_SPEC_MEMO("- 9600 Aldrich Ave S Bloomington 55420 Th hrs: 8:30-6:30 ")
                                    ' Call write_variable_in_SPEC_MEMO("- 1001 Plymouth Ave N Minneapolis 55411")
                                    ' Call write_variable_in_SPEC_MEMO("- 525 Portland Ave S Minneapolis 55415")
                                    ' Call write_variable_in_SPEC_MEMO("- 2215 East Lake Street Minneapolis 55407")
                                    ' CALL write_variable_in_SPEC_MEMO(" (Cov sij hawm qhib yog M - F 8-4:30 tsis li mas yuav tsum qhia ua ntej)")
                                    CALL write_variable_in_SPEC_MEMO("")
									CALL write_variable_in_SPEC_MEMO("You now have an option to use an email to return documents to Hennepin County. Write the case number and full name associated with the case in the body of the email. Only the following types are accepted PNG, JPG, TIFF, DOC, PDF, and HTML. You will not receive confirmation of receipt or failure. To obtain information about your case please contact your worker. EMAIL: hhsews@hennepin.us ")
									Call write_variable_in_SPEC_MEMO(" ")
                                    CALL write_variable_in_SPEC_MEMO("** Yog hais tias tsis hnov koj teb ua ntej " & last_day_of_recert & "  **")
                                    CALL write_variable_in_SPEC_MEMO("**   koj cov kev pab yuav raug kaw thaum " & last_day_of_recert & ".   **")

                                ' Case "06"   'Russian (5th)
                                '     'MsgBox "RUSSIAN"
                                '     if ALL_CASES_ARRAY(recvd_appl, case_entry) = TRUE then CALL write_variable_in_SPEC_MEMO("My' poluchili vashu dokumentacziyu o pereodicheskoj attestaczii " & ALL_CASES_ARRAY(date_of_app, case_entry) & ".")
                                '     if ALL_CASES_ARRAY(recvd_appl, case_entry) = FALSE then CALL write_variable_in_SPEC_MEMO("Vasha dokumentacziya o pereodicheskoj attestaczii eshhyo ne poluchena.")
                                '     CALL write_variable_in_SPEC_MEMO("")
                                '     CALL write_variable_in_SPEC_MEMO("Vy' dolzhny' projti sobesedovanie do " & last_day_of_recert & " ili vasha programma zakroetsya.")
                                '     CALL write_variable_in_SPEC_MEMO("")
                                '     CALL write_variable_in_SPEC_MEMO("Chtoby' projti sobesedovanie po telefonu, pozvonite v Informaczionnuyu liniyu EZ po telefonu 612-596-1300 s 8:00 do 16:30 s ponedel'nika po pyatniczu.")
                                '     CALL write_variable_in_SPEC_MEMO("")
                                '     CALL write_variable_in_SPEC_MEMO("   Esli vy' xotite naznachit' sobesedovanie, pozvonite po telefonu 612-596-1300. Vy' takzhe mozhete obratit'sya v lyuboj iz shesti ofisov. Dlya sobesedovanie s 8 i do 4:30, s ponedel'nika po pyatniczu.")
                                '     Call write_variable_in_SPEC_MEMO("- 7051 Brooklyn Blvd Brooklyn Center 55429")
                                '     Call write_variable_in_SPEC_MEMO("- 1011 1st St S Hopkins 55343")
                                '     Call write_variable_in_SPEC_MEMO("- 9600 Aldrich Ave S Bloomington 55420 Th hrs: 8:30-6:30 ")
                                '     Call write_variable_in_SPEC_MEMO("- 1001 Plymouth Ave N Minneapolis 55411")
                                '     Call write_variable_in_SPEC_MEMO("- 525 Portland Ave S Minneapolis 55415")
                                '     Call write_variable_in_SPEC_MEMO("- 2215 East Lake Street Minneapolis 55407")
                                '     CALL write_variable_in_SPEC_MEMO("(Chasy' priyoma s ponedel'nika po pyatniczu s 8 do 4:30, esli ne ukazano inoe.)")
                                '     CALL write_variable_in_SPEC_MEMO("")
                                '     CALL write_variable_in_SPEC_MEMO("** Esli my' ne usly'shim ot vas do " & last_day_of_recert & " **")
                                '     CALL write_variable_in_SPEC_MEMO("**   vasha programma zakroetsya " & last_day_of_recert & "    **")
                                '
                                ' Case "12"   'Oromo (6th)
                                '     'MsgBox "OROMO"
                                ' Case "03"   'Vietnamese (7th)
                                '     'MsgBox "VIETNAMESE"
                                Case Else  'English (1st)
                                    'MsgBox "ENGLISH"
                                    if ALL_CASES_ARRAY(recvd_appl, case_entry) = TRUE then CALL write_variable_in_SPEC_MEMO("We received your Recertification Paperwork on " & ALL_CASES_ARRAY(date_of_app, case_entry) & ".")
                                    if ALL_CASES_ARRAY(recvd_appl, case_entry) = FALSE then CALL write_variable_in_SPEC_MEMO("Your Recertification Paperwork has not yet been received.")
                                    CALL write_variable_in_SPEC_MEMO("")
                                    CALL write_variable_in_SPEC_MEMO("You must have an interview by " & last_day_of_recert & " or your benefits will end. ")
                                    CALL write_variable_in_SPEC_MEMO("")
                                    Call write_variable_in_SPEC_MEMO("To complete a phone interview, call the EZ Info Line at")
                                    Call write_variable_in_SPEC_MEMO("612-596-1300 between 8:00am and 4:30pm Monday thru Friday.")
                                    CALL write_variable_in_SPEC_MEMO("")
									'removal of in person verbiage during the COVID-19 PEACETIME STATE OF EMERGENCY
						            ' Call write_variable_in_SPEC_MEMO("If you wish to schedule an interview, call 612-596-1300. You may also come to any of the six offices below for an in-person interview between 8 and 4:30, Monday thru Friday.")
						            ' Call write_variable_in_SPEC_MEMO("- 7051 Brooklyn Blvd Brooklyn Center 55429")
						            ' Call write_variable_in_SPEC_MEMO("- 1011 1st St S Hopkins 55343")
						            ' Call write_variable_in_SPEC_MEMO("- 9600 Aldrich Ave S Bloomington 55420 Th hrs: 8:30-6:30 ")
						            ' Call write_variable_in_SPEC_MEMO("- 1001 Plymouth Ave N Minneapolis 55411")
						            ' Call write_variable_in_SPEC_MEMO("- 525 Portland Ave S Minneapolis 55415")
						            ' Call write_variable_in_SPEC_MEMO("- 2215 East Lake Street Minneapolis 55407")
						            ' Call write_variable_in_SPEC_MEMO("(Hours are M - F 8-4:30 unless otherwise noted)")
									CALL write_variable_in_SPEC_MEMO("You now have an option to use an email to return documents to Hennepin County. Write the case number and full name associated with the case in the body of the email. Only the following types are accepted PNG, JPG, TIFF, DOC, PDF, and HTML. You will not receive confirmation of receipt or failure. To obtain information about your case please contact your worker. EMAIL: hhsews@hennepin.us ")
                                    CALL write_variable_in_SPEC_MEMO("")
                                    CALL write_variable_in_SPEC_MEMO("  ** If we do not hear from you by " & last_day_of_recert & "  **")
                                    CALL write_variable_in_SPEC_MEMO("  **   your benefits will end on " & last_day_of_recert & ".   **")

                            End Select

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
                            PF3
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


    objExcel.Cells(entry_row, stats_col+1).Value            = "=" & stats_letter_col & entry_row & "/" & stats_letter_col & total_entry_row
    objExcel.Cells(entry_row, stats_col+1).NumberFormat     = "0.00%"


    objExcel.Cells(entry_row, stats_header_col).Value       = "Cases with no Interview"     'number of cases that potentially need a NOMI'
    objExcel.Cells(entry_row, stats_header_col).Font.Bold 	= TRUE
    objExcel.Cells(entry_row, stats_col).Value              = "=COUNTBLANK(" & intvw_date_letter_col & "2:" & intvw_date_letter_col & last_excel_row & ")"
    objExcel.Cells(entry_row, stats_col+1).Value            = "=" & stats_letter_col & entry_row & "/" & stats_letter_col & "4"
    objExcel.Cells(entry_row, stats_col+1).NumberFormat     = "0.00%"
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
    objExcel.Cells(entry_row, stats_col).Value              = "=COUNTIF(" & intvw_date_letter_col & "2:" & intvw_date_letter_col & last_excel_row & ", " & is_not_blank_excel_string & ")"
    objExcel.Cells(entry_row, stats_col+1).Value            = "=" & stats_letter_col & entry_row & "/" & stats_letter_col & "4"
    objExcel.Cells(entry_row, stats_col+1).NumberFormat     = "0.00%"
    entry_row = entry_row + 1

    objExcel.Cells(entry_row, stats_header_col).Value       = "Applications Received"   'Calculates the percentage of NOMIs siucessful (from attempted)'
    objExcel.Cells(entry_row, stats_header_col).Font.Bold 	= TRUE
    objExcel.Cells(entry_row, stats_col).Value              = "=COUNTIF(" & app_date_letter_col & "2:" & app_date_letter_col & last_excel_row & ", " & is_not_blank_excel_string & ")"
    objExcel.Cells(entry_row, stats_col+1).Value            = "=" & stats_letter_col & entry_row & "/" & stats_letter_col & "4"
    objExcel.Cells(entry_row, stats_col+1).NumberFormat     = "0.00%"
    entry_row = entry_row + 1

    objExcel.Cells(entry_row, stats_header_col).Value       = "Privleged Cases:"        'PRIV cases header'
    objExcel.Cells(entry_row, stats_header_col).Font.Bold 	= TRUE
end if

'IF the Data Only option was selected
if notice_type = "Data Only" then
	row = 2
	cn_col = left(case_number_column, 1)
	call convert_excel_letter_to_excel_number(cn_col)
	do
		if InStr(list_all_workers, trim(objExcel.Cells(row, wrkr_col).Value)) = 0 then list_all_workers = list_all_workers & trim(objExcel.Cells(row, wrkr_col).Value) & "~"
		row = row + 1
		next_case_number = trim(objExcel.Cells(row, cn_col).Value)
	loop until next_case_number = ""

	list_all_workers = left(list_all_workers, len(list_all_workers)-1)
	worker_array = split(list_all_workers, "~")

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

If notice_type = "NOMI" Then

	ObjExcel.Worksheets.Add().Name = "NOMI Count"
	ObjExcel.Cells(1, 1).Value = "NOMIs Sent for this Month"
	ObjExcel.Cells(1, 2).Value = "=COUNTIF('" & scenario_dropdown & "'!" & nomi_letter_col & ":" & nomi_letter_col & ", "& chr(34) & "Y" & chr(34) & ")"
	ObjExcel.columns(1).AutoFit()
	ObjExcel.columns(2).AutoFit()

End If

script_end_procedure("Notices have been sent. Detail of script run is on the spreadsheet that was opened.")
