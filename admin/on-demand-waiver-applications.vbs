'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - APPLICATIONS.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 335                      'manual run time in seconds
STATS_denomination = "C"       			   'C is for each CASE
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
call changelog_update("09/03/2022", "Replaced Jennifer Frey's email contact with Tanya Payne, new HSS for QI.", "Ilse Ferris, Hennepin County")
call changelog_update("08/10/2022", "Added checkbox option in the main dialog to select if user wants Excel output warning message.", "Ilse Ferris, Hennepin County")
call changelog_update("08/01/2022", "Added Excel output warning message.", "Ilse Ferris, Hennepin County")
call changelog_update("05/01/2022", "Updated the Appointment Notice and NOMI to have information for residents about in person support.", "Casey Love, Hennepin County")
call changelog_update("01/25/2022", "Added new QI members and MiKayla to the list of On Demand Application assignment staff.", "Ilse Ferris, Hennepin County")
call changelog_update("12/17/2021", "Updated new MNBenefits website from MNBenefits.org to MNBenefits.mn.gov.", "Ilse Ferris, Hennepin County")
call changelog_update("10/04/2021", "GitHub #475 ADMIN-ON DEMAND APPLICATION: Remove F2F interview supports.", "MiKayla Handley, Hennepin County")
call changelog_update("08/01/2021", "Changed the notices sent in 2 ways:##~## ##~## - Updated verbiage on how to submit documents to Hennepin.##~## ##~## - Appointment Notices will now be sent with a date of 5 days from the date of application.##~##", "Casey Love, Hennepin County")
call changelog_update("03/02/2021", "Update EZ Info Phone hours from 9-4 pm to 8-4:30 pm.", "Ilse Ferris, Hennepin County")
call changelog_update("07/24/2020", "Updated the script to hold the comments section each day.", "MiKayla Handley, Hennepin County")
call changelog_update("05/28/2020", "Update to the notice wording, added virtual drop box information.", "MiKayla Handley, Hennepin County")
call changelog_update("05/13/2020", "Update to the notice wording. Information and direction for in-person interview option removed. County offices are not currently open due to the COVID-19 Peacetime Emergency.", "Casey Love, Hennepin County")
call changelog_update("10/07/2019", "Added HCRE panel bypass in case wonky HCRE panels exist.", "Ilse Ferris, Hennepin County")
call changelog_update("08/21/2019", "Bug on the script when a large PND2 list is accessed.", "Casey Love, Hennepin County")
CALL changelog_update("02/19/2019", "Script will now automatically save the Daily List.", "Casey Love, Hennepin County")
CALL changelog_update("01/30/2019", "Adding tracking of statistics, particularly around NOMIs and Correction Emails.", "Casey Love, Hennepin County")
CALL changelog_update("10/23/2018", "Bug Fixes: Next Action Needed update, Daily List Detail, Cases with Only a Face to Face interview required.", "Casey Love, Hennepin County")
CALL changelog_update("10/22/2018", "Removed denial memo.", "MiKayla Handley, Hennepin County")
CALL changelog_update("07/20/2018", "Updated verbiage of Appointment Notice and NOMI, changed appointment date to 10 days from application date.", "Casey Love, Hennepin County")
CALL changelog_update("07/11/2018", "Adding check to ensure script is not being run in Inquiry.", "Casey Love, Hennepin County")
CALL changelog_update("02/05/2018", "Initial version.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'FUNCTIONS =================================================================================================================
function confirm_memo_waiting(confirmation_var)
    'Function to read for a MEMO created and waiting today
    'This is used to confirm that MEMO creation was successful
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

Function File_Exists(file_name, does_file_exist)
    ' Set objFSO = CreateObject("Scripting.FileSystemObject")
    If (objFSO.FileExists(file_name)) Then
        does_file_exist = True
    Else
      does_file_exist = False
    End If
End Function

'DECLARATIONS =================================================================================================================
'Setting constants for SQL
Const adOpenStatic = 3
Const adLockOptimistic = 3

'Setting constants to make the arrays easier to read
const case_number           = 0
const excel_row             = 1
const client_name			= 2
const program_group_ID		= 3
const worker_ID		   		= 4
const program_status		= 5
const priv_case             = 6
const out_of_co             = 7
const written_lang          = 8
const SNAP_status           = 9
const CASH_status           = 10
const application_date      = 11
const interview_date    	= 12
const appt_notc_sent        = 13
const appt_notc_confirm     = 14
const nomi_sent             = 15
const nomi_confirm          = 16
const deny_day30			= 17
const case_over_30_days     = 18
const need_appt_notc        = 19
const need_nomi             = 20
const appointment_date		= 21
const next_action_needed    = 22
const on_working_list       = 23
const questionable_intv     = 24
const data_day_30			= 25
const data_days_pend		= 26
const take_action_today     = 27

const worker_name_one       = 26
const sup_name_one          = 27
const issue_item_one        = 28
const email_ym_one          = 29
const qi_worker_one         = 30

const worker_name_two       = 31
const sup_name_two          = 32
const issue_item_two        = 33
const email_ym_two          = 34
const qi_worker_two         = 35

const worker_name_three     = 36
const sup_name_three        = 37
const issue_item_three      = 38
const email_ym_three        = 39
const qi_worker_three       = 40

const add_to_daily_worklist = 41
const intvw_quest_resolve	= 42
const email_worker_from_wl	= 43
const email_issue_from_wl	= 44
const rept_pnd2_listed_days	= 45
const additional_app_date 	= 46
const yesterday_action_taken = 47

const case_in_other_co		= 48
const case_closed_in_30		= 49
const line_update_date		= 50
const script_action_taken	= 51
const script_notes_info		= 52
const last_wl_date			= 53
const out_of_co_resolve		= 54
const closed_in_30_resolve	= 55
const subsqt_appl_resolve	= 56
const deleted_today 		= 57

const error_notes 			= 58

'Constants for columns in the working excel sheet - to make the excel code easier to read.
const worker_id_col         = 1
const case_nbr_col          = 2
const case_name_col         = 3
const snap_stat_col         = 4
const cash_stat_col         = 5
const app_date_col          = 6
const second_app_date_col	= 7
const rept_pnd2_days_col	= 8
const intvw_date_col        = 9
const quest_intvw_date_col  = 10
' const resolve_quest_intvw_col = 11
const other_county_col 		= 11
const closed_in_30_col		= 12

const appt_notc_date_col    = 13
const appt_date_col         = 14
const appt_notc_confirm_col = 15
const nomi_date_col         = 16
const nomi_confirm_col      = 17
const need_deny_col         = 18
const next_action_col       = 19
const recent_wl_date_col	= 20
const day_30_col            = 21
const worker_notes_col      = 22
const script_notes_col		= 23
const script_revw_date_col	= 24

const worker_name_one_col   = 25
const sup_name_one_col      = 26
const issue_item_one_col    = 27
const email_ym_one_col      = 28
const qi_worker_one_col     = 29

const worker_name_two_col   = 30
const sup_name_two_col      = 31
const issue_item_two_col    = 32
const email_ym_two_col      = 33
const qi_worker_two_col     = 34

const worker_name_three_col = 35
const sup_name_three_col    = 36
const issue_item_three_col  = 37
const email_ym_three_col    = 38
const qi_worker_three_col   = 39

const list_update_date_col 	= 42

const wl_rept_pnd2_days_col			= 6		'worklist'
const wl_app_date_col 				= 7		'worklist'
const wl_second_app_date_col		= 8		'worklist'
const wl_resolve_2nd_app_date_col	= 9		'worklist'
const wl_intvw_date_col        		= 10	'worklist'
const wl_quest_intvw_date_col  		= 11	'worklist'
const wl_resolve_quest_intvw_col	= 12	'worklist'
const wl_other_county_col			= 13	'worklist'
const wl_resolve_othr_co_col		= 14	'worklist'
const wl_closed_in_30_col			= 15	'worklist'
const wl_resolve_closed_in_30_col	= 16	'worklist'
const wl_appt_notc_date_col   		= 17	'worklist'
const wl_appt_date_col         		= 18	'worklist'
const wl_nomi_date_col         		= 19	'worklist'
const wl_day_30_col 				= 20	'worklist'
const wl_cannot_deny_col 			= 21	'worklist'
' const wl_ecf_doc_accepted_col	= 18	'worklist'
const wl_action_taken_col 			= 22	'worklist'
const wl_work_notes_col				= 23	'worklist'
' const wl_email_worker_col		= 21	'worklist'
' const wl_email_issue_col		= 22	'worklist'

'ARRAY used to store ALL the cases listed on the BOBI today
Dim TODAYS_CASES_ARRAY()
ReDim TODAYS_CASES_ARRAY(error_notes, 0)
'ARRAY of all the cases that are on the working spreadsheet (this is essentially the spreadsheet dumped into a script array for use)
Dim ALL_PENDING_CASES_ARRAY()
ReDim ALL_PENDING_CASES_ARRAY(error_notes, 0)
'ARRAY of all the cases that are on the working spreadsheet (this is essentially the spreadsheet dumped into a script array for use)
Dim WORKING_LIST_CASES_ARRAY()
ReDim WORKING_LIST_CASES_ARRAY(error_notes, 0)
'ARRAY of all the cases that are removed from the working spreadsheet so that they can be reported out after the script run
Dim CASES_NO_LONGER_WORKING()
ReDim CASES_NO_LONGER_WORKING(error_notes, 0)
'creating a new ARRAY of all the cases that we take an action on so that we can add them to a sheet in the daily list
Dim ACTION_TODAY_CASES_ARRAY()
ReDim ACTION_TODAY_CASES_ARRAY(error_notes, 0)
todays_cases = 0        'incrementor for adding to this new array

Dim YESTERDAYS_PENDING_CASES_ARRAY()
ReDim YESTERDAYS_PENDING_CASES_ARRAY(error_notes, 0)

list_of_baskets_at_display_limit = ""					'defaulting some variables that will be added to through the review of the cases
cases_to_alert_BZST = ""

'THE SCRIPT ================================================================================================================
EMConnect ""

'------------------------------------------------------------------------------------------------------establishing date variables
MAXIS_footer_month = CM_plus_1_mo   'Setting footer month and year
MAXIS_footer_year = CM_plus_1_yr

'Opens the current day's list
current_date = date

'setting up information and variables for accessing yesterday's worklist
previous_date = dateadd("d", -1, date)
Call change_date_to_soonest_working_day(previous_date, "back")       'finds the most recent previous working day
archive_folder = DatePart("yyyy", previous_date) & "-" & right("0" & DatePart("m", previous_date), 2)

previous_date_month = DatePart("m", previous_date)
previous_date_day = DatePart("d", previous_date)
previous_date_year = DatePart("yyyy", previous_date)
previous_date_header = previous_date_month & "-" & previous_date_day & "-" & previous_date_year

next_working_day = dateadd("d", 1, date)
Call change_date_to_soonest_working_day(next_working_day, "FORWARD")

'The dialog is defined in the loop as it can change as buttons are pressed
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 316, 160, "Select the source file"
  DropListBox 185, 75, 125, 45, "Select One..."+chr(9)+"Amber Stone"+chr(9)+"Brooke Reilley"+chr(9)+"Deborah Lechner"+chr(9)+"Jacob Arco"+chr(9)+"Jessica Hall"+chr(9)+"Keith Semmelink"+chr(9)+"Kerry Walsh"+chr(9)+"Louise Kinzer"+chr(9)+"Mandora Young"+chr(9)+"MiKayla Handley"+chr(9)+"Ryan Kierth"+chr(9)+"Yeng Yang", qi_member_on_ONDEMAND
  CheckBox 10, 95, 230, 10, "Check here for warning message before Excel output/email creation.", warning_checkbox
  EditBox 90, 110, 220, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 205, 140, 50, 15
    CancelButton 260, 140, 50, 15
  Text 5, 5, 305, 25, "This script will send Appointment Notices and NOMIs after reviewing cases from the BOBI for today. Once completed, this script will create a WorkList for QI to complete any additional manual review or updates."
  Text 5, 35, 80, 10, "Scrript Requirements:"
  Text 10, 50, 45, 10, "- Production"
  Text 10, 60, 75, 10, "- Heavy use of Excel"
  Text 10, 80, 175, 10, "Select the QI Member assigned to On Demand today:"
  Text 10, 115, 80, 10, "Sign your CASE/NOTE:"
  Text 5, 135, 195, 20, "Reminder, do not use Excel during the time the script is running. The script needs to use Excel."
EndDialog

Do
    Do
        err_msg = ""
        dialog Dialog1
        cancel_without_confirmation

		If qi_member_on_ONDEMAND = "Select One..." Then err_msg = err_msg & vbcr & "* Indicate which member of QI is assigned to On Demand today."
		If trim(worker_signature) = "" Then err_msg = err_msg & vbcr & "* Sign your CASE/NOTE."
        If err_msg <> "" and left(err_msg, 4) <> "LOOP" Then MsgBox "Please resolve to continue:" & vbCr & err_msg
    Loop until err_msg = ""
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

call back_to_self
EMReadScreen mx_region, 10, 22, 48

If mx_region = "INQUIRY DB" Then
    continue_in_inquiry = MsgBox("It appears you are attempting to have the script send notices for these cases." & vbNewLine & vbNewLine & "However, you appear to be in MAXIS Inquiry." &vbNewLine & "*************************" & vbNewLine & "Do you want to continue?", vbQuestion + vbYesNo, "Confirm Inquiry")
    If continue_in_inquiry = vbNo Then script_end_procedure_with_error_report("Live script run was attempted in Inquiry and aborted.")
End If

'setting up file paths for accessing yesterday's worklist
archive_files = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\On Demand Waiver\QI On Demand Daily Assignment\Archive\" & archive_folder

previous_list_file_selection_path = t_drive & "/Eligibility Support/Restricted/QI - Quality Improvement/REPORTS/On Demand Waiver/QI On Demand Daily Assignment/QI " & previous_date_header & " Worklist.xlsx"
Call File_Exists(previous_list_file_selection_path, does_file_exist)
previous_worksheet_header = "Work List for " & previous_date_month & "-" & previous_date_day & "-" & previous_date_year



'Checking the working list to see when last updated
'declare the SQL statement that will query the database
objWorkSQL = "SELECT * FROM ES.ES_OnDemanCashAndSnapBZProcessed"

'Creating objects for Access
Set objWorkConnection = CreateObject("ADODB.Connection")
Set objWorkRecordSet = CreateObject("ADODB.Recordset")

'This is the file path for the statistics Access database.
objWorkConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
objWorkRecordSet.Open objWorkSQL, objWorkConnection

'pulling the date changed from the first record in the working list.
'This it to identify if this is a restart or not.
first_item_change = objWorkRecordSet("AuditChangeDate")
first_item_array = split(first_item_change, " ")
first_item_date = first_item_array(0)
first_item_date = DateAdd("d", 0, first_item_date)
first_item_date = #11/4/22#

'If the first item has not been changed today, this is NOT a restart and we need to compare today's list
If first_item_date <> date Then
	'declare the SQL statement that will query the database
	objSQL = "SELECT * FROM ES.ES_OnDemandCashAndSnap"

	'Creating objects for Access
	Set objConnection = CreateObject("ADODB.Connection")
	Set objRecordSet = CreateObject("ADODB.Recordset")

	'This is the file path for the statistics Access database.
	' stats_database_path = "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;"
	objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
	objRecordSet.Open objSQL, objConnection

	'Setting a starting value for a list of cases so that every case is bracketed by * on both sides.
	todays_cases_list = "*"
	case_entry = 0      'Setting an incrementor for the array to be filled

	Do While NOT objRecordSet.Eof
		anything_number = objRecordSet("CaseNumber")
		case_basket = objRecordSet("WorkerID")

		If left(case_basket, 4) = "X127" then
			If instr(todays_cases_list, "*" & anything_number & "*") = 0 then       'This indicates that the case number was not already found on the BOBI
				'MsgBox anything_number
				todays_cases_list = todays_cases_list & anything_number & "*"       'adding the case number on the current row to the list of all the case numbers found.
				ReDim Preserve TODAYS_CASES_ARRAY(error_notes, case_entry)          'resizing the array to add this case to the array

				'Saving each piece of case information from the BOBI to the array
				TODAYS_CASES_ARRAY(worker_ID, case_entry) = objRecordSet("WorkerID")
				TODAYS_CASES_ARRAY(case_number, case_entry) = objRecordSet("CaseNumber")
				Do
					If left(TODAYS_CASES_ARRAY(case_number, case_entry), 1) = "0" Then TODAYS_CASES_ARRAY(case_number, case_entry) = right(TODAYS_CASES_ARRAY(case_number, case_entry), len(TODAYS_CASES_ARRAY(case_number, case_entry)-1))
				Loop until left(TODAYS_CASES_ARRAY(case_number, case_entry), 1) <> "0"
				' TODAYS_CASES_ARRAY(excel_row, case_entry) = row
				TODAYS_CASES_ARRAY(client_name, case_entry) = objRecordSet("CaseName")
				TODAYS_CASES_ARRAY(application_date, case_entry) = objRecordSet("ApplDate")
				TODAYS_CASES_ARRAY(application_date, case_entry) = DateAdd("d", 0, TODAYS_CASES_ARRAY(application_date, case_entry))
				TODAYS_CASES_ARRAY(interview_date, case_entry) = ""
				TODAYS_CASES_ARRAY(data_day_30, case_entry) = objRecordSet("Day_30")
				TODAYS_CASES_ARRAY(data_days_pend, case_entry) = objRecordSet("DaysPending")
				TODAYS_CASES_ARRAY(on_working_list, case_entry) = FALSE         'defaulting this to FALSE

				current_number = anything_number    'saving the case number that is being looked at for the next loop because these are sorted by case number
				case_entry = case_entry + 1         'incrementing for the array to resize on the next loop

			End If
			stats_counter = stats_counter + 1       'incrementing for stats
		End If
		objRecordSet.MoveNext
	Loop

	'close the connection and recordset objects to free up resources
	objRecordSet.Close
	objConnection.Close
	Set objRecordSet=nothing
	Set objConnection=nothing

	case_entry = 0      'incrementor to add a case to ALL_PENDING_CASES_ARRAY
	case_removed = 0    'incrementor to add a case to CASES_NO_LONGER_WORKING
	row = 2             'Working Excel sheet starts with cases on row 2
	list_of_all_cases = ""

	'Reading through each item on the Workking SQP table'
	Do While NOT objWorkRecordSet.Eof
		case_number_to_assess = objWorkRecordSet("CaseNumber")  			'getting the case number in the Working Excel sheet
		case_name_to_assess = objWorkRecordSet("CaseName")
		found_case_on_todays_list = FALSE                               	'this Boolean is used to determine if the case number is on the BOBI run today
		If InStr(list_of_all_cases, "*" & case_number_to_assess & "*") = 0 Then 		'making sure we don't have repeat case numbers
			' list_of_all_cases = list_of_all_cases & case_number_to_assess & "*"

			For each_case = 0 to UBound(TODAYS_CASES_ARRAY, 2)              'This loops through each case that was on the BOBI today
		        'MsgBox "Excel case number: " & case_number_to_assess & vbNewLine & "Array case number: " & TODAYS_CASES_ARRAY(case_number, each_case)
				' If case_number_to_assess = TODAYS_CASES_ARRAY(case_number, each_case) Then  'If a matching case number is found this means the case was on the working excel AND is on the BOBI
		        If case_name_to_assess = TODAYS_CASES_ARRAY(client_name, each_case) Then  'If a matching case number is found this means the case was on the working excel AND is on the BOBI
		            TODAYS_CASES_ARRAY(on_working_list, each_case) = TRUE                   'Idetifying in the list of the cases on the BOBI that this case was also on the working list - and so won't need to be added later
		            found_case_on_todays_list = TRUE                                        'Identifying that this row on the working list was also found on the BOBI - so it won't necessarily have to be removed from the working list later

	                ReDim Preserve ALL_PENDING_CASES_ARRAY(error_notes, case_entry)     'resizing the WORKING CASES ARRAY

					' ObjWorkExcel.Cells(row, script_notes_col).Value = script_notes_var
	                ALL_PENDING_CASES_ARRAY(worker_ID, case_entry) 				= TODAYS_CASES_ARRAY(worker_ID, each_case)
	                ALL_PENDING_CASES_ARRAY(case_number, case_entry) 			= TODAYS_CASES_ARRAY(case_number, each_case)
	                ' ALL_PENDING_CASES_ARRAY(excel_row, case_entry) = row
	                ALL_PENDING_CASES_ARRAY(client_name, case_entry) 			= objWorkRecordSet("CaseName") 'ObjWorkExcel.Cells(row, case_name_col)       'This is gathered from the Working Excel instead of the BOBI list because we may have populated a priv case with an actual name
	                ALL_PENDING_CASES_ARRAY(application_date, case_entry) 		= TODAYS_CASES_ARRAY(application_date, each_case)
					ALL_PENDING_CASES_ARRAY(data_day_30, case_entry) 			= objWorkRecordSet("Day_30")
	                ALL_PENDING_CASES_ARRAY(interview_date, case_entry) 		= objWorkRecordSet("InterviewDate") 		'ObjWorkExcel.Cells(row, intvw_date_col)   'This is gathered from the Working Excel as we may have found an interview date that is NOT in PROG
	                ALL_PENDING_CASES_ARRAY(CASH_status, case_entry) 			= objWorkRecordSet("CashStatus") 			'ObjWorkExcel.Cells(row, cash_stat_col)
	                ALL_PENDING_CASES_ARRAY(SNAP_status, case_entry) 			= objWorkRecordSet("SnapStatus") 			'ObjWorkExcel.Cells(row, snap_stat_col)

	                ALL_PENDING_CASES_ARRAY(appt_notc_sent, case_entry) 		= objWorkRecordSet("ApptNoticeDate") 		'ObjWorkExcel.Cells(row, appt_notc_date_col)
	                ALL_PENDING_CASES_ARRAY(appt_notc_confirm, case_entry) 		= objWorkRecordSet("Confirmation") 			'ObjWorkExcel.Cells(row, appt_notc_confirm_col).Value
	                ALL_PENDING_CASES_ARRAY(appointment_date, case_entry) 		= objWorkRecordSet("ApptDate") 				'ObjWorkExcel.Cells(row, appt_date_col)
					ALL_PENDING_CASES_ARRAY(additional_app_date, case_entry) 	= objWorkRecordSet("SecondApplicationDate") 'ObjWorkExcel.Cells(row, second_app_date_col).Value
					ALL_PENDING_CASES_ARRAY(rept_pnd2_listed_days, case_entry) 	= objWorkRecordSet("REPT_PND2Days") 		'ObjWorkExcel.Cells(row, rept_pnd2_days_col).Value
					ALL_PENDING_CASES_ARRAY(data_days_pend, case_entry) 		= TODAYS_CASES_ARRAY(data_days_pend, each_case)
	                ALL_PENDING_CASES_ARRAY(nomi_sent, case_entry) 				= objWorkRecordSet("NOMIDate") 				'ObjWorkExcel.Cells(row, nomi_date_col)
	                ALL_PENDING_CASES_ARRAY(nomi_confirm, case_entry) 			= objWorkRecordSet("Confirmation2") 		'ObjWorkExcel.Cells(row, nomi_confirm_col)
	                ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) 	= objWorkRecordSet("NextActionNeeded") 			'ObjWorkExcel.Cells(row, next_action_col)
					ALL_PENDING_CASES_ARRAY(subsqt_appl_resolve, case_entry)	= objWorkRecordSet("SecondApplicationDateNotes")
					ALL_PENDING_CASES_ARRAY(case_closed_in_30, case_entry)		= objWorkRecordSet("ClosedInPast30Days")
					ALL_PENDING_CASES_ARRAY(closed_in_30_resolve, case_entry)	= objWorkRecordSet("ClosedInPast30DaysNotes")
					ALL_PENDING_CASES_ARRAY(case_in_other_co, case_entry)		= objWorkRecordSet("StartedOutOfCounty")
					ALL_PENDING_CASES_ARRAY(out_of_co_resolve, case_entry)		= objWorkRecordSet("StartedOutOfCountyNotes")
					ALL_PENDING_CASES_ARRAY(script_notes_info, case_entry)		= objWorkRecordSet("TrackingNotes")
					' ALL_PENDING_CASES_ARRAY(error_notes, case_entry)			= trim(array_of_script_notes(1))
					' ALL_PENDING_CASES_ARRAY(priv_case, case_entry)				= trim(array_of_script_notes(5))

	                ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry) 		= objWorkRecordSet("QuestionableInterview") 'ObjWorkExcel.Cells(row, quest_intvw_date_col)
					ALL_PENDING_CASES_ARRAY(intvw_quest_resolve, case_entry)	= objWorkRecordSet("Resolved")
					ALL_PENDING_CASES_ARRAY(last_wl_date, case_entry) 			= objWorkRecordSet("AddedtoWorkList")
					change_date_time = objWorkRecordSet("AuditChangeDate")
					change_date_time_array = split(change_date_time, " ")
					change_date = change_date_time_array(0)
					ALL_PENDING_CASES_ARRAY(line_update_date, case_entry) 		= DateAdd("d", 0, change_date)

					' "DenialNeeded"
					' ALL_PENDING_CASES_ARRAY(error_notes, case_entry) 			= objWorkRecordSet("AddedtoWorkList") 'ObjWorkExcel.Cells(row, worker_notes_col)
					' ALL_PENDING_CASES_ARRAY(line_update_date, case_entry) 		= objWorkRecordSet("AddedtoWorkList") 'ObjWorkExcel.Cells(row, script_revw_date_col)
					' ALL_PENDING_CASES_ARRAY(line_update_date, case_entry) = dateAdd("d", 0, ALL_PENDING_CASES_ARRAY(line_update_date, case_entry))

	                'Defaulting this values at this time as we will determine them to be different as the script proceeds.
	                ALL_PENDING_CASES_ARRAY(take_action_today, case_entry) = FALSE
					ALL_PENDING_CASES_ARRAY(add_to_daily_worklist, case_entry) = False
					array_of_script_notes = ""
					actions_detail_var = ""
					script_notes_var = ""

	                case_entry = case_entry + 1     'increasing the count for '
		            ' End If
		            ' Exit For                            'This is to leave the loop of looking through all of the cases in the BOBI list ARRAY because we found the match - and there should never be duplicates
		        End If
		    Next

			'If the script has looked through ALL the cases on the BOBI list for today and there was no match for the case number of the row of the Working Excel that we are on
		    'It means that the case is no longer pending for CASH nor for SNAP and we no longer need to look at it.
		    If found_case_on_todays_list = FALSE Then   'this was defaulted to FALSE and is only changed to TRUE when a case number match on today's BOBI list
		        'MsgBox "NOT ON TODAY'S LIST" & vbNewLine & ObjWorkExcel.Cells(row, case_nbr_col)
		        ReDim Preserve CASES_NO_LONGER_WORKING(error_notes, case_removed)       'increasing the size of the array
		        'Gathering all the detail from the working Excel and adding to the removed CASES ARRAY so that we can list it at the end.
		        ' MISSING - CASES_NO_LONGER_WORKING(worker_ID, case_removed) = ObjWorkExcel.Cells(row, worker_id_col)
		        CASES_NO_LONGER_WORKING(case_number, case_removed) 				= objWorkRecordSet("CaseNumber") 'ObjWorkExcel.Cells(row, case_nbr_col)
		        ' CASES_NO_LONGER_WORKING(excel_row, case_removed) = row
		        CASES_NO_LONGER_WORKING(client_name, case_removed) 				= objWorkRecordSet("CaseName") 'ObjWorkExcel.Cells(row, case_name_col)
		        CASES_NO_LONGER_WORKING(application_date, case_removed) 		= objWorkRecordSet("ApplDate") 'ObjWorkExcel.Cells(row, app_date_col)
		        CASES_NO_LONGER_WORKING(interview_date, case_removed) 			= objWorkRecordSet("InterviewDate") 'ObjWorkExcel.Cells(row, intvw_date_col)
		        CASES_NO_LONGER_WORKING(CASH_status, case_removed) 				= objWorkRecordSet("CashStatus") 'ObjWorkExcel.Cells(row, cash_stat_col)
		        CASES_NO_LONGER_WORKING(SNAP_status, case_removed) 				= objWorkRecordSet("SnapStatus") 'ObjWorkExcel.Cells(row, snap_stat_col)
		        CASES_NO_LONGER_WORKING(appt_notc_sent, case_removed) 			= objWorkRecordSet("ApptNoticeDate") 'ObjWorkExcel.Cells(row, appt_notc_date_col)
		        CASES_NO_LONGER_WORKING(appt_notc_confirm, case_removed) 		= objWorkRecordSet("Confirmation") 'ObjWorkExcel.Cells(row, appt_notc_confirm_col).Value
		        CASES_NO_LONGER_WORKING(appointment_date, case_removed) 		= objWorkRecordSet("ApptDate") 'ObjWorkExcel.Cells(row, appt_date_col)
				CASES_NO_LONGER_WORKING(additional_app_date, case_removed) 		= objWorkRecordSet("SecondApplicationDate") 'ObjWorkExcel.Cells(row, second_app_date_col).Value
				CASES_NO_LONGER_WORKING(rept_pnd2_listed_days, case_removed) 	= objWorkRecordSet("Day_30") 'ObjWorkExcel.Cells(row, rept_pnd2_days_col).Value
		        CASES_NO_LONGER_WORKING(nomi_sent, case_removed) 				= objWorkRecordSet("NOMIDate") 'ObjWorkExcel.Cells(row, nomi_date_col)
		        CASES_NO_LONGER_WORKING(nomi_confirm, case_removed) 			= objWorkRecordSet("Confirmation2") 'ObjWorkExcel.Cells(row, nomi_confirm_col)
		        CASES_NO_LONGER_WORKING(next_action_needed, case_removed) 		= objWorkRecordSet("NextActionNeeded") 'ObjWorkExcel.Cells(row, next_action_col)
		        CASES_NO_LONGER_WORKING(questionable_intv, case_removed) 		= objWorkRecordSet("QuestionableInterview") 'ObjWorkExcel.Cells(row, quest_intvw_date_col)

				'TODO - MISSING - CASES_NO_LONGER_WORKING(case_in_other_co, case_removed) = ObjWorkExcel.Cells(row, other_county_col)
				'TODO - MISSING - CASES_NO_LONGER_WORKING(case_closed_in_30, case_removed) = ObjWorkExcel.Cells(row, closed_in_30_col)

				' CASES_NO_LONGER_WORKING(intvw_quest_resolve, case_removed) = ObjWorkExcel.Cells(row, resolve_quest_intvw_col)

		        CASES_NO_LONGER_WORKING(error_notes, case_removed) = ""
		        'CASES_NO_LONGER_WORKING(error_notes, case_removed) = "Interview Completed on " & TODAYS_CASES_ARRAY(interview_date, case_entry)
		        'MsgBox row
		        case_removed = case_removed + 1     'adding to the incrementer for the removed cases ARRAY
		    End If
		End  If
		objWorkRecordSet.MoveNext
	Loop

	objWorkRecordSet.Close
	objWorkConnection.Close

	Set objWorkRecordSet=nothing
	Set objWorkConnection=nothing
	Set objWorkSQL=nothing

	'Actually deleting the row in the Working Excel - notice that ROW does not increase as the curent row is now new
	' For delete_case = 0 to UBound(CASES_NO_LONGER_WORKING, 2)
	' 	case_number_to_review = CASES_NO_LONGER_WORKING(case_number, delete_case)
	' 	objWorkRecordSet.Open "DELETE FROM ES.ES_OnDemanCashAndSnapBZProcessed WHERE CaseNumber = '" & case_number_to_review & "'", objWorkConnection
	' Next
	' objWorkRecordSet.Close
	' objWorkConnection.Close

	'BE SURE TO ALWAYS LEAVE THE row VARIABLE ALONE HERE AS WE USE IT IN THIS FOR NEXT TO ADD TO THE END OF THE WORKING EXCEL
	add_a_case = case_entry     'creating an incrementer that starts where the last one ended for the ALL PENDING CASES ARRAY
	' case_added = False
	For case_entry = 0 to UBOUND(TODAYS_CASES_ARRAY, 2)     'now we are going to look at each of the cases in the ARRAY for today's BOBI list
	    'MsgBox TODAYS_CASES_ARRAY(on_working_list, case_entry)
	    'MsgBox TODAYS_CASES_ARRAY(interview_date, case_entry)
	    If TODAYS_CASES_ARRAY(on_working_list, case_entry) = FALSE AND TODAYS_CASES_ARRAY(interview_date, case_entry) = "" Then
	        'These are all the cases on todays list that were NOT on the Working Excel AND have not already had an interview
	        'adding the information known from the BOBI to the Working Excel

	        ReDim Preserve ALL_PENDING_CASES_ARRAY(error_notes, add_a_case)         'resizing the array of the Working Excel

			' case_added = True
	        ALL_PENDING_CASES_ARRAY(worker_ID, add_a_case) 				= TODAYS_CASES_ARRAY(worker_ID, case_entry)
	        ALL_PENDING_CASES_ARRAY(case_number, add_a_case) 			= TODAYS_CASES_ARRAY(case_number, case_entry)
	        ' ALL_PENDING_CASES_ARRAY(excel_row, add_a_case) = row
	        ALL_PENDING_CASES_ARRAY(client_name, add_a_case) 			= TODAYS_CASES_ARRAY(client_name, case_entry)
	        ALL_PENDING_CASES_ARRAY(application_date, add_a_case) 		= TODAYS_CASES_ARRAY(application_date, case_entry)
	        ALL_PENDING_CASES_ARRAY(interview_date, add_a_case) 		= TODAYS_CASES_ARRAY(interview_date, case_entry)
			ALL_PENDING_CASES_ARRAY(data_days_pend, add_a_case) 		= TODAYS_CASES_ARRAY(data_days_pend, case_entry)
			ALL_PENDING_CASES_ARRAY(data_day_30, add_a_case) 			= TODAYS_CASES_ARRAY(data_day_30, case_entry)
			ALL_PENDING_CASES_ARRAY(take_action_today, add_a_case) 		= FALSE
			ALL_PENDING_CASES_ARRAY(add_to_daily_worklist, add_a_case) 	= False
			ALL_PENDING_CASES_ARRAY(client_name, add_a_case) = replace(ALL_PENDING_CASES_ARRAY(client_name, add_a_case), "'", "")

			' objWorkRecordSet.Open "INSERT INTO ES.ES_OnDemanCashAndSnapBZProcessed (CaseNumber, CaseName, ApplDate, InterviewDate, Day_30, DaysPending)" & _
			' 				  "VALUES ('" & ALL_PENDING_CASES_ARRAY(case_number, add_a_case) &  "', '" & _
			' 				  				ALL_PENDING_CASES_ARRAY(client_name, add_a_case) &  "', '" & _
			' 								ALL_PENDING_CASES_ARRAY(application_date, add_a_case) &  "', '" & _
			' 								ALL_PENDING_CASES_ARRAY(interview_date, add_a_case) &  "', '" & _
			' 								ALL_PENDING_CASES_ARRAY(data_day_30, add_a_case) &  "', '" & _
			' 								ALL_PENDING_CASES_ARRAY(data_days_pend, add_a_case) & "')", objWorkConnection, adOpenStatic, adLockOptimistic

	        add_a_case = add_a_case + 1     'incrementing the counter for this ARRAY
	        ' row = row + 1                   'going to the next row so that we don't overwrite the information we just added
	    End If
	Next
	' If case_added = True Then objWorkRecordSet.Close


	'TODO - remove this part when we movve off of the worklist process.
	yesterday_case_list = 0

	If does_file_exist = True Then
		'open the file
		call excel_open(previous_list_file_selection_path, True, False, ObjYestExcel, objYestWorkbook)

		objYestWorkbook.Worksheets("Statistics").visible = True
		objYestWorkbook.worksheets("Statistics").Activate
		' yesterday_worker = ObjYestExcel.Cells(2, 2).Value

		objYestWorkbook.worksheets(previous_worksheet_header).Activate

		objYestWorkbook.Worksheets("Statistics").visible = False

		'Pull info into a NEW array of prevvious day work.
		xl_row = 2
		Do
			this_case = trim(ObjYestExcel.Cells(xl_row, case_nbr_col).Value)
			If this_case <> "" Then
				ReDim Preserve YESTERDAYS_PENDING_CASES_ARRAY(error_notes, yesterday_case_list)

				YESTERDAYS_PENDING_CASES_ARRAY(worker_ID, yesterday_case_list) = trim(ObjYestExcel.Cells(xl_row, worker_id_col).Value)
				YESTERDAYS_PENDING_CASES_ARRAY(case_number, yesterday_case_list) = trim(ObjYestExcel.Cells(xl_row, case_nbr_col).Value)
				YESTERDAYS_PENDING_CASES_ARRAY(client_name, yesterday_case_list) = trim(ObjYestExcel.Cells(xl_row, case_name_col).Value)
				YESTERDAYS_PENDING_CASES_ARRAY(SNAP_status, yesterday_case_list) = trim(ObjYestExcel.Cells(xl_row, snap_stat_col).Value)
				YESTERDAYS_PENDING_CASES_ARRAY(CASH_status, yesterday_case_list) = trim(ObjYestExcel.Cells(xl_row, cash_stat_col).Value)
				' YESTERDAYS_PENDING_CASES_ARRAY(, yesterday_case_list) = trim(ObjYestExcel.Cells(xl_row, wl_rept_pnd2_days_col).Value)
				YESTERDAYS_PENDING_CASES_ARRAY(application_date, yesterday_case_list) = trim(ObjYestExcel.Cells(xl_row, wl_app_date_col).Value)
				YESTERDAYS_PENDING_CASES_ARRAY(additional_app_date, yesterday_case_list) = trim(ObjYestExcel.Cells(xl_row, wl_second_app_date_col).Value)
				YESTERDAYS_PENDING_CASES_ARRAY(subsqt_appl_resolve, yesterday_case_list) = trim(ObjYestExcel.Cells(xl_row, wl_resolve_2nd_app_date_col).Value)
				YESTERDAYS_PENDING_CASES_ARRAY(interview_date, yesterday_case_list) = trim(ObjYestExcel.Cells(xl_row, wl_intvw_date_col).Value)
				YESTERDAYS_PENDING_CASES_ARRAY(questionable_intv, yesterday_case_list) = trim(ObjYestExcel.Cells(xl_row, wl_quest_intvw_date_col).Value)
				YESTERDAYS_PENDING_CASES_ARRAY(intvw_quest_resolve, yesterday_case_list) = trim(ObjYestExcel.Cells(xl_row, wl_resolve_quest_intvw_col).Value)

				YESTERDAYS_PENDING_CASES_ARRAY(case_in_other_co, yesterday_case_list) = ObjYestExcel.Cells(row, wl_other_county_col)
				YESTERDAYS_PENDING_CASES_ARRAY(out_of_co_resolve, yesterday_case_list) = trim(ObjYestExcel.Cells(xl_row, wl_resolve_othr_co_col).Value)
				YESTERDAYS_PENDING_CASES_ARRAY(case_closed_in_30, yesterday_case_list) = ObjYestExcel.Cells(row, wl_closed_in_30_col)
				YESTERDAYS_PENDING_CASES_ARRAY(closed_in_30_resolve, yesterday_case_list) = trim(ObjYestExcel.Cells(xl_row, wl_resolve_closed_in_30_col).Value)

				' YESTERDAYS_PENDING_CASES_ARRAY(intvw_quest_resolve, yesterday_case_list) = trim(ObjYestExcel.Cells(xl_row, wl_resolve_quest_intvw_col).Value)

				YESTERDAYS_PENDING_CASES_ARRAY(appt_notc_sent, yesterday_case_list) = trim(ObjYestExcel.Cells(xl_row, wl_appt_notc_date_col).Value)
				YESTERDAYS_PENDING_CASES_ARRAY(appointment_date, yesterday_case_list) = trim(ObjYestExcel.Cells(xl_row, wl_appt_date_col).Value)
				YESTERDAYS_PENDING_CASES_ARRAY(nomi_sent, yesterday_case_list) = trim(ObjYestExcel.Cells(xl_row, wl_nomi_date_col).Value)
				' YESTERDAYS_PENDING_CASES_ARRAY(, yesterday_case_list) = trim(ObjYestExcel.Cells(xl_row, wl_day_30_col).Value)
				' YESTERDAYS_PENDING_CASES_ARRAY(, yesterday_case_list) = trim(ObjYestExcel.Cells(xl_row, wl_deny_col).Value)
				YESTERDAYS_PENDING_CASES_ARRAY(yesterday_action_taken, yesterday_case_list) = trim(ObjYestExcel.Cells(xl_row, wl_action_taken_col).Value)
				YESTERDAYS_PENDING_CASES_ARRAY(error_notes, yesterday_case_list) = trim(ObjYestExcel.Cells(xl_row, wl_work_notes_col).Value)

				yesterday_case_list = yesterday_case_list + 1
				xl_row = xl_row + 1
			End If
		Loop until this_case = ""

		'close the file
		ObjYestExcel.ActiveWorkbook.Close
		ObjYestExcel.Application.Quit
		ObjYestExcel.Quit

		For case_entry = 0 to UBOUND(ALL_PENDING_CASES_ARRAY, 2)

			'CHECK THE LIST and compare it against the previous day work to capture any important details
			For yest_entry = 0 to UBound(YESTERDAYS_PENDING_CASES_ARRAY, 2)
				If ALL_PENDING_CASES_ARRAY(case_number, case_entry) = YESTERDAYS_PENDING_CASES_ARRAY(case_number, yest_entry) Then
					ALL_PENDING_CASES_ARRAY(case_in_other_co, case_entry) = YESTERDAYS_PENDING_CASES_ARRAY(case_in_other_co, yest_entry)
					ALL_PENDING_CASES_ARRAY(case_closed_in_30, case_entry) = YESTERDAYS_PENDING_CASES_ARRAY(case_closed_in_30, yest_entry)

					ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry) = YESTERDAYS_PENDING_CASES_ARRAY(questionable_intv, yest_entry)
					ALL_PENDING_CASES_ARRAY(intvw_quest_resolve, case_entry) = YESTERDAYS_PENDING_CASES_ARRAY(intvw_quest_resolve, yest_entry)
					ALL_PENDING_CASES_ARRAY(additional_app_date, case_entry) = YESTERDAYS_PENDING_CASES_ARRAY(additional_app_date, yest_entry)
					ALL_PENDING_CASES_ARRAY(subsqt_appl_resolve, case_entry) = YESTERDAYS_PENDING_CASES_ARRAY(subsqt_appl_resolve, yest_entry)

					ALL_PENDING_CASES_ARRAY(closed_in_30_resolve, case_entry)	= YESTERDAYS_PENDING_CASES_ARRAY(closed_in_30_resolve, yest_entry)
					ALL_PENDING_CASES_ARRAY(out_of_co_resolve, case_entry)		= YESTERDAYS_PENDING_CASES_ARRAY(out_of_co_resolve, yest_entry)

					If ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) <> "" Then YESTERDAYS_PENDING_CASES_ARRAY(error_notes, yest_entry) = replace(YESTERDAYS_PENDING_CASES_ARRAY(error_notes, yest_entry), ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry),"")
					yesterdays_notes = YESTERDAYS_PENDING_CASES_ARRAY(error_notes, yest_entry)
					If ALL_PENDING_CASES_ARRAY(subsqt_appl_resolve, case_entry) <> "" Then yesterdays_notes = "Subsqnt APPL: " & ALL_PENDING_CASES_ARRAY(subsqt_appl_resolve, case_entry) & " - " & yesterdays_notes
					yesterdays_action_info = YESTERDAYS_PENDING_CASES_ARRAY(yesterday_action_taken, yest_entry)
					' If yesterday_worker = qi_member_on_ONDEMAND Then
					ALL_PENDING_CASES_ARRAY(error_notes, case_entry) = yesterdays_action_info & " - " & yesterdays_notes
					yesterdays_action_info = UCase(yesterdays_action_info)
					If InStr(yesterdays_action_info, "FOLLOW UP NEEDED") <> 0 Then
						ALL_PENDING_CASES_ARRAY(add_to_daily_worklist, case_entry) = True
						If InStr(ALL_PENDING_CASES_ARRAY(error_notes, case_entry), "Carry over from Yesterday Worklist") = 0 Then ALL_PENDING_CASES_ARRAY(error_notes, case_entry) = "Carry over from Yesterday Worklist. " & ALL_PENDING_CASES_ARRAY(error_notes, case_entry)
					End If
					IF ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "REVIEW CANNOT DENY - No Appt Notc" Then ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "DENY AT DAY 30"
					IF ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "REVIEW CANNOT DENY - No NOMI" Then ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "DENY AT DAY 30"
					IF ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "REVIEW CANNOT DENY - NOMI after Day 30" Then ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "DENY AT DAY 30"
					IF ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "PREP FOR DENIAL" Then ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "DENY AT DAY 30"

					IF ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "ALIGN INTERVIEW DATES" Then ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = ""

					IF ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "REVIEW QUESTIONABLE INTERVIEW DATE(S)" and YESTERDAYS_PENDING_CASES_ARRAY(intvw_quest_resolve, yest_entry) <> "" THEN ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = ""
					' IF ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "REVIEW OTHER COUNTY CASE"	and YESTERDAYS_PENDING_CASES_ARRAY(out_of_co_resolve, yest_entry) <> "" Then ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = ""
					IF ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "REVIEW OTHER COUNTY CASE" Then ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = ""
					IF ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "RESOLVE SUBSEQUENT APPLICATION DATE" and YESTERDAYS_PENDING_CASES_ARRAY(subsqt_appl_resolve, yest_entry) <> "" Then ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = ""
					If ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "REVIEW RECENT CLOSURE/DENIAL" and YESTERDAYS_PENDING_CASES_ARRAY(closed_in_30_resolve, yest_entry) <> "" Then ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = ""

					' working_row = ALL_PENDING_CASES_ARRAY(excel_row, case_entry)
					' ObjWorkExcel.Cells(working_row, other_county_col).Value = ALL_PENDING_CASES_ARRAY(case_in_other_co, case_entry)
					' ObjWorkExcel.Cells(working_row, closed_in_30_col).Value = ALL_PENDING_CASES_ARRAY(case_closed_in_30, case_entry)
					' ObjWorkExcel.Cells(working_row, worker_notes_col).Value = ALL_PENDING_CASES_ARRAY(error_notes, case_entry)
					' If ALL_PENDING_CASES_ARRAY(add_to_daily_worklist, case_entry) = True Then ObjWorkExcel.Cells(working_row, script_notes_col).Value = ObjWorkExcel.Cells(working_row, script_notes_col).Value & "-ADD TO TODAY'S WORKLIST"
				End If
			Next
		Next
	End If



	For case_entry = 0 to UBOUND(ALL_PENDING_CASES_ARRAY, 2)
		ALL_PENDING_CASES_ARRAY(deleted_today, case_entry) = False
		' If ALL_PENDING_CASES_ARRAY(line_update_date, case_entry) = date Then
		' 	If ALL_PENDING_CASES_ARRAY(script_action_taken, case_entry) = True Then
		' 		ReDim Preserve ACTION_TODAY_CASES_ARRAY(error_notes, todays_cases)
		' 		ACTION_TODAY_CASES_ARRAY(case_number, todays_cases)         = ALL_PENDING_CASES_ARRAY(case_number, case_entry)
		' 		ACTION_TODAY_CASES_ARRAY(client_name, todays_cases)         = ALL_PENDING_CASES_ARRAY(client_name, case_entry)
		' 		ACTION_TODAY_CASES_ARRAY(worker_ID, todays_cases)           = ALL_PENDING_CASES_ARRAY(worker_ID, case_entry)
		' 		ACTION_TODAY_CASES_ARRAY(SNAP_status, todays_cases)         = ALL_PENDING_CASES_ARRAY(SNAP_status, case_entry)
		' 		ACTION_TODAY_CASES_ARRAY(CASH_status, todays_cases)         = ALL_PENDING_CASES_ARRAY(CASH_status, case_entry)
		' 		ACTION_TODAY_CASES_ARRAY(application_date, todays_cases)    = ALL_PENDING_CASES_ARRAY(application_date, case_entry)
		' 		ACTION_TODAY_CASES_ARRAY(interview_date, todays_cases)      = ALL_PENDING_CASES_ARRAY(interview_date, case_entry)
		' 		ACTION_TODAY_CASES_ARRAY(questionable_intv, todays_cases)   = ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry)
		' 		ACTION_TODAY_CASES_ARRAY(appt_notc_sent, todays_cases)      = ALL_PENDING_CASES_ARRAY(appt_notc_sent, case_entry)
		' 		ACTION_TODAY_CASES_ARRAY(appt_notc_confirm, todays_cases)   = ALL_PENDING_CASES_ARRAY(appt_notc_confirm, case_entry)
		' 		ACTION_TODAY_CASES_ARRAY(appointment_date, todays_cases)    = ALL_PENDING_CASES_ARRAY(appointment_date, case_entry)
		' 		ACTION_TODAY_CASES_ARRAY(nomi_sent, todays_cases)           = ALL_PENDING_CASES_ARRAY(nomi_sent, case_entry)
		' 		ACTION_TODAY_CASES_ARRAY(nomi_confirm, todays_cases)        = ALL_PENDING_CASES_ARRAY(nomi_confirm, case_entry)
		' 		ACTION_TODAY_CASES_ARRAY(deny_day30, todays_cases)          = ALL_PENDING_CASES_ARRAY(deny_day30, case_entry)
		' 		ACTION_TODAY_CASES_ARRAY(deny_memo_confirm, todays_cases)   = ALL_PENDING_CASES_ARRAY(deny_memo_confirm, case_entry)
		' 		ACTION_TODAY_CASES_ARRAY(next_action_needed, todays_cases)  = ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry)
		' 		ACTION_TODAY_CASES_ARRAY(error_notes, todays_cases)         = ALL_PENDING_CASES_ARRAY(error_notes, case_entry)
		' 		ALL_PENDING_CASES_ARRAY(script_action_taken, case_entry) = True
		' 		todays_cases = todays_cases + 1
		' 	End If
		' End If
		date_zero =  #1/1/1900#
		If IsDate(ALL_PENDING_CASES_ARRAY(interview_date, case_entry)) = True Then
			If DateDiff("d", ALL_PENDING_CASES_ARRAY(interview_date, case_entry),date_zero) = 0 Then ALL_PENDING_CASES_ARRAY(interview_date, case_entry) = ""
		Else
			ALL_PENDING_CASES_ARRAY(interview_date, case_entry) = ""
		End if
		If IsDate(ALL_PENDING_CASES_ARRAY(additional_app_date, case_entry)) = True Then
			If DateDiff("d", ALL_PENDING_CASES_ARRAY(additional_app_date, case_entry),date_zero) = 0 Then ALL_PENDING_CASES_ARRAY(additional_app_date, case_entry) = ""
		Else
			ALL_PENDING_CASES_ARRAY(additional_app_date, case_entry) = ""
		End if
		If IsDate(ALL_PENDING_CASES_ARRAY(appt_notc_sent, case_entry)) = True Then
			If DateDiff("d", ALL_PENDING_CASES_ARRAY(appt_notc_sent, case_entry),date_zero) = 0 Then ALL_PENDING_CASES_ARRAY(appt_notc_sent, case_entry) = ""
		Else
			ALL_PENDING_CASES_ARRAY(appt_notc_sent, case_entry) = ""
		End if
		If IsDate(ALL_PENDING_CASES_ARRAY(appointment_date, case_entry)) = True Then
			If DateDiff("d", ALL_PENDING_CASES_ARRAY(appointment_date, case_entry),date_zero) = 0 Then ALL_PENDING_CASES_ARRAY(appointment_date, case_entry) = ""
		Else
			ALL_PENDING_CASES_ARRAY(appointment_date, case_entry) = ""
		End if
		If IsDate(ALL_PENDING_CASES_ARRAY(nomi_sent, case_entry)) = True Then
			If DateDiff("d", ALL_PENDING_CASES_ARRAY(nomi_sent, case_entry),date_zero) = 0 Then ALL_PENDING_CASES_ARRAY(nomi_sent, case_entry) = ""
		Else
			ALL_PENDING_CASES_ARRAY(nomi_sent, case_entry) = ""
		End if
		If IsDate(ALL_PENDING_CASES_ARRAY(last_wl_date, case_entry)) = True Then
			If DateDiff("d", ALL_PENDING_CASES_ARRAY(last_wl_date, case_entry),date_zero) = 0 Then ALL_PENDING_CASES_ARRAY(last_wl_date, case_entry) = ""
		Else
			ALL_PENDING_CASES_ARRAY(last_wl_date, case_entry) = ""
		End if
		' If ALL_PENDING_CASES_ARRAY(additional_app_date, case_entry) = date_zero Then ALL_PENDING_CASES_ARRAY(additional_app_date, case_entry) = ""
		' If ALL_PENDING_CASES_ARRAY(appt_notc_sent, case_entry) = date_zero Then ALL_PENDING_CASES_ARRAY(appt_notc_sent, case_entry) = ""
		' If ALL_PENDING_CASES_ARRAY(appointment_date, case_entry) = date_zero Then ALL_PENDING_CASES_ARRAY(appointment_date, case_entry) = ""
		' If ALL_PENDING_CASES_ARRAY(nomi_sent, case_entry) = date_zero Then ALL_PENDING_CASES_ARRAY(nomi_sent, case_entry) = ""
		' If ALL_PENDING_CASES_ARRAY(last_wl_date, case_entry) = date_zero Then ALL_PENDING_CASES_ARRAY(last_wl_date, case_entry) = ""

		If UCase(ALL_PENDING_CASES_ARRAY(add_to_daily_worklist, case_entry)) = "TRUE" Then ALL_PENDING_CASES_ARRAY(add_to_daily_worklist, case_entry) = True
		If UCase(ALL_PENDING_CASES_ARRAY(priv_case, case_entry)) = "TRUE" Then ALL_PENDING_CASES_ARRAY(priv_case, case_entry) = True
		If UCase(ALL_PENDING_CASES_ARRAY(case_in_other_co, case_entry)) = "TRUE" Then ALL_PENDING_CASES_ARRAY(case_in_other_co, case_entry) = True
		If UCase(ALL_PENDING_CASES_ARRAY(case_closed_in_30, case_entry)) = "TRUE" Then ALL_PENDING_CASES_ARRAY(case_closed_in_30, case_entry) = True

		If UCase(ALL_PENDING_CASES_ARRAY(add_to_daily_worklist, case_entry)) = "FALSE" Then ALL_PENDING_CASES_ARRAY(add_to_daily_worklist, case_entry) = False
		If UCase(ALL_PENDING_CASES_ARRAY(priv_case, case_entry)) = "FALSE" Then ALL_PENDING_CASES_ARRAY(priv_case, case_entry) = False
		If UCase(ALL_PENDING_CASES_ARRAY(case_in_other_co, case_entry)) = "FALSE" Then ALL_PENDING_CASES_ARRAY(case_in_other_co, case_entry) = False
		If UCase(ALL_PENDING_CASES_ARRAY(case_closed_in_30, case_entry)) = "FALSE" Then ALL_PENDING_CASES_ARRAY(case_closed_in_30, case_entry) = False

		If  ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "Send Manual NOMI" Then  ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = ""
		If  ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "Send Manual Appt Notice" Then  ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = ""

		' If ALL_PENDING_CASES_ARRAY(line_update_date, case_entry) <> date Then

			MAXIS_case_number = ALL_PENDING_CASES_ARRAY(case_number, case_entry)        'setting this so that nav functionality works
			ALL_PENDING_CASES_ARRAY(script_action_taken, case_entry) = False

			' If ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "REVIEW QUESTIONABLE INTERVIEW DATE(S)" Then ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = ""
			ALL_PENDING_CASES_ARRAY(error_notes, case_entry) = Replace(ALL_PENDING_CASES_ARRAY(error_notes, case_entry), "Display Limit", "")
			ALL_PENDING_CASES_ARRAY(error_notes, case_entry) = Replace(ALL_PENDING_CASES_ARRAY(error_notes, case_entry), "PRIVILEGED CASE.", "")
			ALL_PENDING_CASES_ARRAY(error_notes, case_entry) = Replace(ALL_PENDING_CASES_ARRAY(error_notes, case_entry), "Cash interview incomplete.", "")
			ALL_PENDING_CASES_ARRAY(error_notes, case_entry) = Replace(ALL_PENDING_CASES_ARRAY(error_notes, case_entry), "SCRIPT DENIAL ALREADY NOTED", "")
			ALL_PENDING_CASES_ARRAY(error_notes, case_entry) = trim(ALL_PENDING_CASES_ARRAY(error_notes, case_entry))

			CALL back_to_SELF
			Call navigate_to_MAXIS_screen_review_PRIV("STAT", "PROG", is_this_priv)
			EMReadScreen county_check, 2, 21, 23            'Looking to see if case has Hennepin County worker
		    EMReadScreen case_removed_in_MAXIS, 19, 24, 2   'There was one case that was removed from MX and it got a little weird.

			If is_this_priv = True Then
				ALL_PENDING_CASES_ARRAY(priv_case, case_entry) = TRUE
				ALL_PENDING_CASES_ARRAY(add_to_daily_worklist, case_entry) = True
				If Instr(ALL_PENDING_CASES_ARRAY(error_notes, case_entry), "PRIVILEGED CASE") = 0 Then ALL_PENDING_CASES_ARRAY(error_notes, case_entry) = "PRIVILEGED CASE. " & ALL_PENDING_CASES_ARRAY(error_notes, case_entry)
			Elseif county_check <> "27" Then
				ALL_PENDING_CASES_ARRAY(out_of_co, case_entry) = "OUT OF COUNTY - " & county_check
			ElseIf case_removed_in_MAXIS = "INVALID CASE NUMBER" Then
				ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "CASE HAS BEEN DELETED"
			Else
				ALL_PENDING_CASES_ARRAY(priv_case, case_entry) = FALSE

				'These caseloads have IMD cases and it is important to note them.
		        IF ALL_PENDING_CASES_ARRAY(worker_ID, case_entry) = "X127EF8" or ALL_PENDING_CASES_ARRAY(worker_ID, case_entry) = "X127EJ1" THEN
					IF InStr(ALL_PENDING_CASES_ARRAY(error_notes, case_entry), "IMD CASE") = 0 THEN ALL_PENDING_CASES_ARRAY(error_notes, case_entry) = ALL_PENDING_CASES_ARRAY(error_notes, case_entry) & ", IMD CASE"
				END IF
		        'Some PRIV cases do not have the client name in BOBI - this will find them

				'PROG to determine programs pending and interview dates
		        fs_intv = ""            'These need to be blanked out for each run as sometimes they are not found for each run and so there is carryover
		        cash_intv_one = ""
		        cash_intv_two = ""

		        'reading programs types and statuses
		        EMReadScreen cash_prog_one, 2, 6, 67
		        EMReadScreen cash_stat_one, 4, 6, 74
		        EMReadScreen cash_prog_two, 2, 7, 67
		        EMReadScreen cash_stat_two, 4, 7, 74
		        EMReadScreen fs_pend, 4, 10, 74

		        'defaulting these to false for each run through the loop
		        cash_pend = FALSE
		        cash_interview_done = FALSE
		        snap_interview_done = FALSE

				ALL_PENDING_CASES_ARRAY(CASH_status, case_entry) = ""       'resetting this so that if it has changed we get good information

				If cash_stat_one = "PEND" Then                              'If the first cash line indicates pending - look for interview information
					cash_pend = TRUE                                        'defining cash as a pending program
					EMReadScreen cash_intv_one, 8, 6, 55                    'read the interview date
					If cash_intv_one <> "__ __ __" Then                     'if it is not blank
						cash_intv_one = replace(cash_intv_one, " ", "/")    'convert it to an actual date
						cash_interview_done = TRUE                          'define that the interview for cash has been done
					Else
						cash_intv_one = ""
					End If
				ElseIf cash_stat_one = "ACTV" Then
					ALL_PENDING_CASES_ARRAY(CASH_status, case_entry) = "Active" 'setting the array to identify that cash is active
				End If

				'it is impportant that line 2 is looked at second because we could ahve an active cash program BUT line 2 indicates that another cash program is PENDING
				'having line 2 second will overwrite the line 1 happenings.
				If cash_stat_two = "PEND" Then                              'otherwirse, if the second cash line indicated pending, we will look at that line for information
					cash_pend = TRUE                                        'note that cash is pending
					EMReadScreen cash_intv_two, 8, 7, 55                    'reading the interview date
					If cash_intv_two <> "__ __ __" Then                     'will convert to a date
						cash_intv_two = replace(cash_intv_two, " ", "/")
						cash_interview_done = TRUE                          'dfines that n interview is done
					Else
						cash_intv_two = ""                                  'making that blank interview date a true blank
					End If
				ElseIf cash_stat_two = "ACTV" Then
					ALL_PENDING_CASES_ARRAY(CASH_status, case_entry) = "Active" 'setting the array to identify that cash is active'
				End If

				If cash_pend = TRUE then ALL_PENDING_CASES_ARRAY(CASH_status, case_entry) = "Pending"       'setting the cash status if a pending cash was found

				If fs_pend = "PEND" Then                                            'if the SNAP status is pending
		            ALL_PENDING_CASES_ARRAY(SNAP_status, case_entry) = "Pending"    'define the pending status in the ARRAY
		            EMReadScreen fs_intv, 8, 10, 55                                 'read the interview date and reformat
		            If fs_intv <> "__ __ __" Then
		                fs_intv = replace(fs_intv, " ", "/")
		                snap_interview_done = TRUE                                  'define the interview done
		            Else
		                fs_intv = ""
		            End If
		        ElseIf fs_pend = "ACTV" Then        'setting the correct infomration in the array otherwise
		            ALL_PENDING_CASES_ARRAY(SNAP_status, case_entry) = "Active"
		        Else
		            ALL_PENDING_CASES_ARRAY(SNAP_status, case_entry) = ""
		        End If

				'Here we have a chain of logic that will help to identify if what needs to happen from this point on
				'first, something needs to be PENDING for this process to apply - if neither are pending - we need to get rid of it
				If ALL_PENDING_CASES_ARRAY(SNAP_status, case_entry) <> "Pending" AND ALL_PENDING_CASES_ARRAY(CASH_status, case_entry) <> "Pending" Then
					ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "REMOVE FROM LIST"            'set this variable because we can't just delete it now - the rows have all been defined to the array and everything will get messed up
					ALL_PENDING_CASES_ARRAY(error_notes, case_entry) = ALL_PENDING_CASES_ARRAY(error_notes, case_entry) & ", Neither SNAP nor CASH is pending."  'explain the removal - the case will be deleted at tomorrow's run
				Else                                                                                        'if one of these is pending
					ALL_PENDING_CASES_ARRAY(interview_date, case_entry) = ""
					If cash_pend = TRUE and cash_interview_done = TRUE Then
						If cash_intv_one <> "" Then ALL_PENDING_CASES_ARRAY(interview_date, case_entry) = cash_intv_one     'setting the interview date from what was found in PROG
						If cash_intv_two <> "" Then ALL_PENDING_CASES_ARRAY(interview_date, case_entry) = cash_intv_two
					End If
					If cash_pend = TRUE and fs_pend = "PEND" Then
						If cash_interview_done = TRUE and fs_intv <> "" Then
							ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "NONE - Interview Completed"
						ElseIf cash_interview_done = TRUE and fs_intv = "" Then
							ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "ALIGN INTERVIEW DATES"
						ElseIf cash_interview_done = False and fs_intv <> "" Then
							ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "ALIGN INTERVIEW DATES"
						ElseIf cash_interview_done = False and fs_intv = "" Then
							ALL_PENDING_CASES_ARRAY(interview_date, case_entry) = ""
						End if

					ElseIf cash_pend = TRUE Then
						If cash_interview_done = TRUE Then ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "NONE - Interview Completed"
					ElseIf fs_pend = "PEND" Then
						If fs_intv <> "" Then
							ALL_PENDING_CASES_ARRAY(interview_date, case_entry) = fs_intv
							If cash_interview_done = TRUE Then ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "NONE - Interview Completed"
						End If
					End If
				End If


				'HCRE bypass coding
		    	PF3		'exits PROG to prommpt HCRE if HCRE insn't complete
		    	Do
		    		EMReadscreen HCRE_panel_check, 4, 2, 50
		    		If HCRE_panel_check = "HCRE" then
		    			PF10	'exists edit mode in cases where HCRE isn't complete for a member
		    			PF3
		    		END IF
		    	Loop until HCRE_panel_check <> "HCRE"

				If ALL_PENDING_CASES_ARRAY(client_name, case_entry) = "XXXXX" Then
					Call navigate_to_MAXIS_screen("STAT", "MEMB")       'go to MEMB - do not need to chose a different memb number because we are looking for the case name
					EMReadScreen last_name, 25, 6, 30       'read each name
					EMReadScreen first_name, 12, 6, 63
					EMReadScreen middle_initial, 1, 6, 79
					last_name = replace(last_name, "_", "") 'format so there are no underscores
					first_name = replace(first_name, "_", "")
					middle_initial = replace(middle_initial, "_", "")

					ALL_PENDING_CASES_ARRAY(client_name, case_entry) = last_name & ", " & first_name & " " & middle_initial     'this is how the BOBI lists names so we want them to match
				End If
				Call back_to_SELF

			End If


			'DELETE THESE FROM SQL'
			delete_from_sql = False
			If ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "CASE HAS BEEN DELETED" Then delete_from_sql = True
			If ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "NONE - Interview Completed" Then delete_from_sql = True
			If ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "REMOVE FROM LIST" Then delete_from_sql = True

			If delete_from_sql = True Then
				ALL_PENDING_CASES_ARRAY(line_update_date, case_entry) = date
				ALL_PENDING_CASES_ARRAY(deleted_today, case_entry) = TRUE
				ReDim Preserve CASES_NO_LONGER_WORKING(error_notes, case_removed)       'increasing the size of the array
				'Gathering all the detail from the working Excel and adding to the removed CASES ARRAY so that we can list it at the end.
				' MISSING - CASES_NO_LONGER_WORKING(worker_ID, case_removed) = ObjWorkExcel.Cells(row, worker_id_col)
				CASES_NO_LONGER_WORKING(case_number, case_removed) 				= ALL_PENDING_CASES_ARRAY(case_number, case_entry)
				' CASES_NO_LONGER_WORKING(excel_row, case_removed) = row
				CASES_NO_LONGER_WORKING(client_name, case_removed) 				= ALL_PENDING_CASES_ARRAY(client_name, case_entry)
				CASES_NO_LONGER_WORKING(application_date, case_removed) 		= ALL_PENDING_CASES_ARRAY(application_date, case_entry)
				CASES_NO_LONGER_WORKING(interview_date, case_removed) 			= ALL_PENDING_CASES_ARRAY(interview_date, case_entry)
				CASES_NO_LONGER_WORKING(CASH_status, case_removed) 				= ALL_PENDING_CASES_ARRAY(CASH_status, case_entry)
				CASES_NO_LONGER_WORKING(SNAP_status, case_removed) 				= ALL_PENDING_CASES_ARRAY(SNAP_status, case_entry)
				CASES_NO_LONGER_WORKING(appt_notc_sent, case_removed) 			= ALL_PENDING_CASES_ARRAY(appt_notc_sent, case_entry)
				CASES_NO_LONGER_WORKING(appt_notc_confirm, case_removed) 		= ALL_PENDING_CASES_ARRAY(appt_notc_confirm, case_entry)
				CASES_NO_LONGER_WORKING(appointment_date, case_removed) 		= ALL_PENDING_CASES_ARRAY(appointment_date, case_entry)
				CASES_NO_LONGER_WORKING(additional_app_date, case_removed) 		= ALL_PENDING_CASES_ARRAY(additional_app_date, case_entry)
				CASES_NO_LONGER_WORKING(rept_pnd2_listed_days, case_removed) 	= ALL_PENDING_CASES_ARRAY(rept_pnd2_listed_days, case_entry)
				CASES_NO_LONGER_WORKING(nomi_sent, case_removed) 				= ALL_PENDING_CASES_ARRAY(nomi_sent, case_entry)
				CASES_NO_LONGER_WORKING(nomi_confirm, case_removed) 			= ALL_PENDING_CASES_ARRAY(nomi_confirm, case_entry)
				CASES_NO_LONGER_WORKING(next_action_needed, case_removed) 		= ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry)
				CASES_NO_LONGER_WORKING(questionable_intv, case_removed) 		= ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry)

				CASES_NO_LONGER_WORKING(error_notes, case_removed) = ""
				'CASES_NO_LONGER_WORKING(error_notes, case_removed) = "Interview Completed on " & TODAYS_CASES_ARRAY(interview_date, case_entry)
				'MsgBox row
				case_removed = case_removed + 1     'adding to the incrementer for the removed cases ARRAY

				'Actually deleting the row in the Working Excel - notice that ROW does not increase as the curent row is now new
				' objWorkRecordSet.Open "DELETE FROM ES.ES_OnDemanCashAndSnapBZProcessed WHERE CaseNumber = '" & MAXIS_case_number & "'", objWorkConnection
				' objWorkRecordSet.Close
				' objWorkConnection.Close
			End If
		' End If
	Next

	count_case_to_add = 0
	For case_entry = 0 to UBOUND(ALL_PENDING_CASES_ARRAY, 2)
		If ALL_PENDING_CASES_ARRAY(deleted_today, case_entry) = False Then count_case_to_add = count_case_to_add + 1
	Next
	MsgBox "count_case_to_add - " & count_case_to_add

	On Error Resume Next
	'declare the SQL statement that will query the database
	objWorkSQL = "SELECT * FROM ES.ES_OnDemanCashAndSnapBZProcessed"

	'Creating objects for Access
	Set objWorkConnection = CreateObject("ADODB.Connection")
	Set objWorkRecordSet = CreateObject("ADODB.Recordset")

	'This is the file path for the statistics Access database.
	objWorkConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
	' objWorkRecordSet.Open objWorkSQL, objWorkConnection

	objWorkRecordSet.Open "DELETE FROM ES.ES_OnDemanCashAndSnapBZProcessed", objWorkConnection, adOpenStatic, adLockOptimistic

	objWorkRecordSet.Close
	objWorkConnection.Close

	Set objWorkRecordSet=nothing
	Set objWorkConnection=nothing
	Set objWorkSQL=nothing


	'declare the SQL statement that will query the database
	objWorkSQL = "SELECT * FROM ES.ES_OnDemanCashAndSnapBZProcessed"

	'Creating objects for Access
	Set objWorkConnection = CreateObject("ADODB.Connection")
	Set objWorkRecordSet = CreateObject("ADODB.Recordset")

	'This is the file path for the statistics Access database.
	objWorkConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""

	For case_entry = 0 to UBOUND(ALL_PENDING_CASES_ARRAY, 2)
		' If ALL_PENDING_CASES_ARRAY(line_update_date, case_entry) <> date and ALL_PENDING_CASES_ARRAY(priv_case, case_entry) = False Then
		If ALL_PENDING_CASES_ARRAY(deleted_today, case_entry) = False Then
			If ALL_PENDING_CASES_ARRAY(add_to_daily_worklist, case_entry) = True Then ALL_PENDING_CASES_ARRAY(last_wl_date, case_entry) = date

			objWorkRecordSet.Open "INSERT INTO ES.ES_OnDemanCashAndSnapBZProcessed (CaseNumber, CaseName, ApplDate, InterviewDate, Day_30, DaysPending, SnapStatus, CashStatus, SecondApplicationDate, REPT_PND2Days, QuestionableInterview, Resolved, ApptNoticeDate, ApptDate, Confirmation, NOMIDate, Confirmation2, DenialNeeded, NextActionNeeded, AddedtoWorkList, SecondApplicationDateNotes, ClosedInPast30Days, ClosedInPast30DaysNotes, StartedOutOfCounty, StartedOutOfCountyNotes, TrackingNotes)" & _
							  "VALUES ('" & ALL_PENDING_CASES_ARRAY(case_number, case_entry) &  "', '" & _
											ALL_PENDING_CASES_ARRAY(client_name, case_entry) &  "', '" & _
											ALL_PENDING_CASES_ARRAY(application_date, case_entry) &  "', '" & _
											ALL_PENDING_CASES_ARRAY(additional_app_date, case_entry) &  "', '" & _
											ALL_PENDING_CASES_ARRAY(data_day_30, case_entry) &  "', '" & _
											ALL_PENDING_CASES_ARRAY(data_days_pend, case_entry) &  "', '" & _
											ALL_PENDING_CASES_ARRAY(SNAP_status, case_entry) &  "', '" & _
											ALL_PENDING_CASES_ARRAY(CASH_status, case_entry) &  "', '" & _
											ALL_PENDING_CASES_ARRAY(additional_app_date, case_entry) &  "', '" & _
											ALL_PENDING_CASES_ARRAY(rept_pnd2_listed_days, case_entry) &  "', '" & _
											ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry) &  "', '" & _
											ALL_PENDING_CASES_ARRAY(intvw_quest_resolve, case_entry) &  "', '" & _
											ALL_PENDING_CASES_ARRAY(appt_notc_sent, case_entry) &  "', '" & _
											ALL_PENDING_CASES_ARRAY(appointment_date, case_entry) &  "', '" & _
											ALL_PENDING_CASES_ARRAY(appt_notc_confirm, case_entry) &  "', '" & _
											ALL_PENDING_CASES_ARRAY(nomi_sent, case_entry) &  "', '" & _
											ALL_PENDING_CASES_ARRAY(nomi_confirm, case_entry) &  "', '" & _
											ALL_PENDING_CASES_ARRAY(case_over_30_days, case_entry) &  "', '" & _
											ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) &  "', '" & _
											ALL_PENDING_CASES_ARRAY(last_wl_date, case_entry) &  "', '" & _
											ALL_PENDING_CASES_ARRAY(subsqt_appl_resolve, case_entry) &  "', '" & _
											ALL_PENDING_CASES_ARRAY(case_closed_in_30, case_entry) &  "', '" & _
											ALL_PENDING_CASES_ARRAY(closed_in_30_resolve, case_entry) &  "', '" & _
											ALL_PENDING_CASES_ARRAY(case_in_other_co, case_entry) &  "', '" & _
											ALL_PENDING_CASES_ARRAY(out_of_co_resolve, case_entry) &  "', '" & _
											ALL_PENDING_CASES_ARRAY(script_notes_info, case_entry) & "')", objWorkConnection, adOpenStatic, adLockOptimistic
		End If
	Next
	objWorkRecordSet.Close
	objWorkConnection.Close

	Set objWorkRecordSet=nothing
	Set objWorkConnection=nothing
	Set objWorkSQL=nothing




	'Now the script reopens the daily list that was identified in the beginning
	file_date = replace(current_date, "/", "-")   'Changing the format of the date to use as file path selection default
	daily_case_list_folder = right("0" & DatePart("m", file_date), 2) & "-" & DatePart("yyyy", file_date)
	file_selection_path = t_drive & "/Eligibility Support/Restricted/QI - Quality Improvement/REPORTS/On Demand Waiver/Daily case lists/" & daily_case_list_folder & "/" & file_date & ".xlsx" 'single assignment file

	' call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)
	'Opening the Excel file, (now that the dialog is done)
	'creating a new file to create the 'Daily List'
	Set objExcel = CreateObject("Excel.Application")
	objExcel.Visible = True
	Set objWorkbook = objExcel.Workbooks.Add()
	objExcel.DisplayAlerts = True

	'Changes name of Excel sheet to "Case information"
	ObjExcel.ActiveSheet.Name = "Report 1"

	ObjExcel.Cells(2, 4) = "Report 1"
	ObjExcel.Cells(4, 2) = "Case Worker ID"
	ObjExcel.Cells(4, 3) = "Case Number"
	ObjExcel.Cells(4, 4) = "Case Name"
	ObjExcel.Cells(4, 5) = "Program ID"
	ObjExcel.Cells(4, 6) = "Program Status"
	ObjExcel.Cells(4, 7) = "Program Application Date"
	ObjExcel.Cells(4, 8) = "Interview Date"
	ObjExcel.Range("B4:H4").Interior.ColorIndex = 5
	ObjExcel.Range("B4:H4").Font.ColorIndex = 2
	ObjExcel.Range("B4:H4").Font.Bold = True

	rept_one_row = 5
	For each_case = 0 to UBound(TODAYS_CASES_ARRAY, 2)
		ObjExcel.Cells(rept_one_row, 2) = TODAYS_CASES_ARRAY(worker_ID, each_case)
		ObjExcel.Cells(rept_one_row, 3) = TODAYS_CASES_ARRAY(case_number, each_case)
		ObjExcel.Cells(rept_one_row, 4) = TODAYS_CASES_ARRAY(client_name, each_case)
		ObjExcel.Cells(rept_one_row, 7) = TODAYS_CASES_ARRAY(application_date, each_case)
		rept_one_row = rept_one_row + 1
	Next

	For col_to_autofit =2 to  8
	    ObjExcel.Columns(col_to_autofit).AutoFit()
	Next

	'It creates a new worksheet and names it
	ObjExcel.Worksheets.Add().Name = "Cases Removed From Working LIST"

	'Then it creates column headers
	ObjExcel.Cells(1, worker_id_col)        = "Worker ID"
	ObjExcel.Cells(1, case_nbr_col)         = "Case Number"
	ObjExcel.Cells(1, case_name_col)        = "Case Name"
	ObjExcel.Cells(1, snap_stat_col)        = "SNAP"
	ObjExcel.Cells(1, cash_stat_col)        = "CASH"
	ObjExcel.Cells(1, app_date_col)         = "Application Date"
	ObjExcel.Cells(1, second_app_date_col) 	= "Second App Date"
	ObjExcel.Cells(1, rept_pnd2_days_col)	= "REPT/PND2 Days"
	ObjExcel.Cells(1, intvw_date_col)       = "Interview Date"
	ObjExcel.Cells(1, quest_intvw_date_col) = "Questionable Interview Date"

	ObjExcel.Cells(1, other_county_col)		= "Case was in Other Co."
	ObjExcel.Cells(1, closed_in_30_col)		= "Closed in Past 30 Days"
	' ObjExcel.Cells(1, resolve_quest_intvw_col) = "Resolved?"

	ObjExcel.Cells(1, appt_notc_date_col)   = "Appt Notice Sent"
	ObjExcel.Cells(1, appt_date_col)        = "Appointment Date"
	ObjExcel.Cells(1, appt_notc_confirm_col)= "Confirm"
	ObjExcel.Cells(1, nomi_date_col)        = "NOMI Sent"
	ObjExcel.Cells(1, nomi_confirm_col)     = "Confirm"
	ObjExcel.Cells(1, need_deny_col)        = "Denial"
	ObjExcel.Cells(1, next_action_col)      = "Next Action"
	ObjExcel.Cells(1, worker_notes_col)     = "Detail"

	ObjExcel.Rows(1).Font.Bold = TRUE   'Making the header row bold

	removed_row = 2     'setting a row counter
	For case_removed = 0 to UBOUND(CASES_NO_LONGER_WORKING, 2)      'looping through each of the cases in the ARRAY from the beginning of cases that were taken off of the Working Excel
	    ' If CASES_NO_LONGER_WORKING(error_notes, case_removed) = "" OR CASES_NO_LONGER_WORKING(client_name, case_removed) = "XXXXX" Then     'if we do not know WHY the case was removed or if the client's name is not filled in - we will go searching for a reason
	    '     'PROG to determine programs active
	    '     MAXIS_case_number = CASES_NO_LONGER_WORKING(case_number, case_removed)      'setting this for nav functions'
	    '     CALL navigate_to_MAXIS_screen("CASE", "CURR")
	    '     'Checking for PRIV cases.
	    '     EMReadScreen priv_check, 6, 24, 14 'If it can't get into the case needs to skip
	    '     EMReadScreen county_check, 2, 21, 16    'Looking to see if case has Hennepin COunty worker
	    '     If priv_check = "PRIVIL" THEN       'idetifying PRIV cases '
	    '         CASES_NO_LONGER_WORKING(error_notes, case_removed) = "PRIV"
	    '     ElseIf county_check <> "27" THEN        'Identifying cases out of county -they would no longer show up on our BOBI and so would be removed from the Working Excel
	    '         CASES_NO_LONGER_WORKING(error_notes, case_removed) = "Transferred out of county - " & county_check
	    '     ElseIf CASES_NO_LONGER_WORKING(client_name, case_removed) = "XXXXX" Then        'Some priv cases we have access to - we can look up the names where the BOBI doesn't have them
	    '         Call navigate_to_MAXIS_screen("STAT", "MEMB")       'Going to MEMB for 01
	    '         EMReadScreen last_name, 25, 6, 30                   'getting name information
	    '         EMReadScreen first_name, 12, 6, 63
	    '         EMReadScreen middle_initial, 1, 6, 79
		'
	    '         last_name = replace(last_name, "_", "")             'reformatting
	    '         first_name = replace(first_name, "_", "")
	    '         middle_initial = replace(middle_initial, "_", "")
		'
	    '         CASES_NO_LONGER_WORKING(client_name, case_removed) = last_name & ", " & first_name & " " & middle_initial   'saving to the ARRAY in the same structure as the BOBI does
	    '     End If
		'
		' 	If CASES_NO_LONGER_WORKING(error_notes, case_removed) = "" Then     'If we STILL don't know why the case was removed then we are going to look at PROG
	    '     'most cases are removed because an interview has been completed OR SNAP/Cash have been acted upon
		'
	    '         Call navigate_to_MAXIS_screen("STAT", "PROG")       'this is the same code as above
	    '         fs_intv = ""            'blanking out these variables
	    '         cash_intv_one = ""
	    '         cash_intv_two = ""
	    '         CASES_NO_LONGER_WORKING(CASH_status, case_removed) = "" 'setting this a blank as we will reread it
		'
	    '         EMReadScreen cash_prog_one, 2, 6, 67    'reading each of the programs and statuses
	    '         EMReadScreen cash_stat_one, 4, 6, 74
		'
	    '         EMReadScreen cash_prog_two, 2, 7, 67
	    '         EMReadScreen cash_stat_two, 4, 7, 74
		'
	    '         EMReadScreen fs_pend, 4, 10, 74
		'
	    '         cash_pend = FALSE           'resetting these for each loop - we will look for TRUEs next
	    '         cash_interview_done = FALSE
	    '         snap_interview_done = FALSE
		'
	    '         If cash_stat_one = "PEND" Then      'if this is pending we will look for an interview date
	    '             cash_pend = TRUE                'setting this to true
	    '             EMReadScreen cash_intv_one, 8, 6, 55
	    '             If cash_intv_one <> "__ __ __" Then     'formatting the date field read
	    '                 cash_intv_one = replace(cash_intv_one, " ", "/")
	    '                 cash_interview_done = TRUE
	    '             Else
	    '                 cash_intv_one = ""
	    '             End If
	    '         ElseIf cash_stat_one = "ACTV" Then      'if this is active - saving that to the ARRAY
	    '             CASES_NO_LONGER_WORKING(CASH_status, case_removed) = "Active"
	    '         End If
		'
	    '         If cash_stat_two = "PEND" Then      'if this is pending we will look for an interview date
	    '             cash_pend = TRUE                'setting this to true
	    '             EMReadScreen cash_intv_two, 8, 7, 55    'reading and formatting the date
	    '             If cash_intv_two <> "__ __ __" Then
	    '                 cash_intv_two = replace(cash_intv_two, " ", "/")
	    '                 cash_interview_done = TRUE
	    '             Else
	    '                 cash_intv_two = ""
	    '             End If
	    '         ElseIf cash_stat_one = "ACTV" Then      'if active, setting that to the ARRAY
	    '             CASES_NO_LONGER_WORKING(CASH_status, case_removed) = "Active"
	    '         End If
		'
	    '         'Setting ARRAY if either case programs is pending
	    '         If cash_pend = TRUE then CASES_NO_LONGER_WORKING(CASH_status, case_removed) = "Pending"
		'
	    '         If fs_pend = "PEND" Then    'if the SNAP is pending we are going to look for an interview
	    '             CASES_NO_LONGER_WORKING(SNAP_status, case_removed) = "Pending"  'setting the status in the ARRAY
	    '             EMReadScreen fs_intv, 8, 10, 55     'reading and formatting the interview date
	    '             If fs_intv <> "__ __ __" Then
	    '                 fs_intv = replace(fs_intv, " ", "/")
	    '                 snap_interview_done = TRUE
	    '             Else
	    '                 fs_intv = ""
	    '             End If
	    '         ElseIf fs_pend = "ACTV" Then        'setting to active if SNAP is active
	    '             CASES_NO_LONGER_WORKING(SNAP_status, case_removed) = "Active"
	    '         Else
	    '             CASES_NO_LONGER_WORKING(SNAP_status, case_removed) = ""
	    '         End If
		'
	    '         'if nothing is pending then the application process is over
	    '         If CASES_NO_LONGER_WORKING(SNAP_status, case_removed) <> "Pending" AND CASES_NO_LONGER_WORKING(CASH_status, case_removed) <> "Pending" Then
	    '             CASES_NO_LONGER_WORKING(next_action_needed, case_removed) = "REMOVE FROM LIST"  'indicate in the ARRAY that there is no pening programs
	    '             CASES_NO_LONGER_WORKING(error_notes, case_removed) = "Neither SNAP nor CASH is pending."
	    '         Else                                'if either program is pending, we are going to look at interview logic
	    '             If cash_pend = TRUE Then        'if cash is pending we will check for cash interviews first
	    '                 If cash_interview_done = TRUE Then  'if the cash interview is done then the interview is done. and we will add the right information to the ARRAY
	    '                     If cash_intv_one <> "" Then CASES_NO_LONGER_WORKING(interview_date, case_removed) = cash_intv_one
	    '                     If cash_intv_two <> "" Then CASES_NO_LONGER_WORKING(interview_date, case_removed) = cash_intv_two
	    '                     CASES_NO_LONGER_WORKING(next_action_needed, case_removed) = "NONE - Interview Completed"
	    '                 Else        'if the cash interview is NOT done we are going to look for a possibel SNAP interview
	    '                     If fs_pend = "PEND" Then        'this looks for a SNAP interview and then looks to see if we need a seperate Face to Face interview for the cash program
	    '                         If fs_intv = "" THen
	    '                             CASES_NO_LONGER_WORKING(interview_date, case_removed) = ""
	    '                         Else
	    '                             CASES_NO_LONGER_WORKING(interview_date, case_removed) = fs_intv
	    '                             CASES_NO_LONGER_WORKING(error_notes, case_removed) = ", Cash interview incomplete."
	    '                         End If
	    '                     End If
	    '                 End If
	    '             ElseIf fs_pend = "PEND" Then    'if cash is not pending but SNAP is, we will look for a SNAP interview
	    '                 If fs_intv <> "" Then
	    '                     CASES_NO_LONGER_WORKING(interview_date, case_removed) = fs_intv
	    '                     CASES_NO_LONGER_WORKING(next_action_needed, case_removed) = "NONE - Interview Completed"
	    '                 End If
	    '             End If
	    '         End If
		'
	    '         'HCRE bypass coding
	    '         PF3		'exits PROG to prompt HCRE if HCRE insn't complete
	    '         Do
	    '             EMReadscreen HCRE_panel_check, 4, 2, 50
	    '             If HCRE_panel_check = "HCRE" then
	    '                 PF10	'exists edit mode in cases where HCRE isn't complete for a member
	    '                 PF3
	    '             END IF
	    '         Loop until HCRE_panel_check <> "HCRE"
	    '     End If
	    ' End If

	    'making sure the script has the Excel Daily List up and saves the information about the case to the next blank row
	    ObjExcel.Worksheets("Cases Removed From Working LIST").Activate
	    'MsgBox "Row is " & removed_row & vbNewLine & "Worker ID " & CASES_NO_LONGER_WORKING(worker_ID, case_removed)
	    ObjExcel.Cells(removed_row, worker_id_col).Value            = CASES_NO_LONGER_WORKING(worker_ID, case_removed)
	    ObjExcel.Cells(removed_row, case_nbr_col).Value             = CASES_NO_LONGER_WORKING(case_number, case_removed)
	    'CASES_NO_LONGER_WORKING(excel_removed_row, case_removed) = removed_row
	    ObjExcel.Cells(removed_row, case_name_col).Value            = CASES_NO_LONGER_WORKING(client_name, case_removed)
	    ObjExcel.Cells(removed_row, app_date_col).Value             = CASES_NO_LONGER_WORKING(application_date, case_removed)
		ObjExcel.Cells(removed_row, second_app_date_col).Value		= CASES_NO_LONGER_WORKING(additional_app_date, case_removed)
		ObjExcel.Cells(removed_row, rept_pnd2_days_col).Value		= CASES_NO_LONGER_WORKING(rept_pnd2_listed_days, case_removed)

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

		ObjExcel.Cells(removed_row, other_county_col).Value     	= CASES_NO_LONGER_WORKING(case_in_other_co, case_removed)
		ObjExcel.Cells(removed_row, closed_in_30_col).Value     	= CASES_NO_LONGER_WORKING(case_closed_in_30, case_removed)

	    ' ObjExcel.Cells(removed_row, resolve_quest_intvw_col).Value	= CASES_NO_LONGER_WORKING(intvw_quest_resolve, case_removed)

	    ObjExcel.Cells(removed_row, worker_notes_col).Value         = CASES_NO_LONGER_WORKING(next_action_needed, case_removed)  & " - " & CASES_NO_LONGER_WORKING(error_notes, case_removed)

		removed_row = removed_row + 1   'moving to the next row for the next loop
	Next

	'formatting the spreadsheet
	For col_to_autofit =1 to  qi_worker_three_col
	    ObjExcel.Columns(col_to_autofit).AutoFit()
	Next

	'Saving the Daily List
	ObjExcel.ActiveWorkbook.SaveAs file_selection_path
	ObjExcel.Quit


	On Error GoTo 0

End If



'Checking the working list to see when last updated
'declare the SQL statement that will query the database
objWorkSQL = "SELECT * FROM ES.ES_OnDemanCashAndSnapBZProcessed"

'Creating objects for Access
Set objWorkConnection = CreateObject("ADODB.Connection")
Set objWorkRecordSet = CreateObject("ADODB.Recordset")

'This is the file path for the statistics Access database.
objWorkConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
objWorkRecordSet.Open objWorkSQL, objWorkConnection



case_entry = 0      'incrementor to add a case to WORKING_LIST_CASES_ARRAY
list_of_all_cases = ""

'Reading through each item on the Workking SQP table'
Do While NOT objWorkRecordSet.Eof
	case_number_to_assess = objWorkRecordSet("CaseNumber")  			'getting the case number in the Working Excel sheet
	' case_name_to_assess = objWorkRecordSet("CaseName")
	' found_case_on_todays_list = FALSE                               	'this Boolean is used to determine if the case number is on the BOBI run today
	If InStr(list_of_all_cases, "*" & case_number_to_assess & "*") = 0 Then 		'making sure we don't have repeat case numbers
		list_of_all_cases = list_of_all_cases & case_number_to_assess & "*"
  		ReDim Preserve WORKING_LIST_CASES_ARRAY(error_notes, case_entry)     'resizing the WORKING CASES ARRAY

        ' WORKING_LIST_CASES_ARRAY(worker_ID, case_entry) 				= TODAYS_CASES_ARRAY(worker_ID, each_case)
        WORKING_LIST_CASES_ARRAY(case_number, case_entry) 			= objWorkRecordSet("CaseNumber")
        ' WORKING_LIST_CASES_ARRAY(excel_row, case_entry) = row
        WORKING_LIST_CASES_ARRAY(client_name, case_entry) 			= objWorkRecordSet("CaseName") 'ObjWorkExcel.Cells(row, case_name_col)       'This is gathered from the Working Excel instead of the BOBI list because we may have populated a priv case with an actual name
        WORKING_LIST_CASES_ARRAY(application_date, case_entry) 		= objWorkRecordSet("ApplDate")	'TODAYS_CASES_ARRAY(application_date, each_case)
		WORKING_LIST_CASES_ARRAY(data_day_30, case_entry) 			= objWorkRecordSet("Day_30")
        WORKING_LIST_CASES_ARRAY(interview_date, case_entry) 		= objWorkRecordSet("InterviewDate") 		'ObjWorkExcel.Cells(row, intvw_date_col)   'This is gathered from the Working Excel as we may have found an interview date that is NOT in PROG
        WORKING_LIST_CASES_ARRAY(CASH_status, case_entry) 			= objWorkRecordSet("CashStatus") 			'ObjWorkExcel.Cells(row, cash_stat_col)
        WORKING_LIST_CASES_ARRAY(SNAP_status, case_entry) 			= objWorkRecordSet("SnapStatus") 			'ObjWorkExcel.Cells(row, snap_stat_col)

        WORKING_LIST_CASES_ARRAY(appt_notc_sent, case_entry) 		= objWorkRecordSet("ApptNoticeDate") 		'ObjWorkExcel.Cells(row, appt_notc_date_col)
        WORKING_LIST_CASES_ARRAY(appt_notc_confirm, case_entry) 		= objWorkRecordSet("Confirmation") 			'ObjWorkExcel.Cells(row, appt_notc_confirm_col).Value
        WORKING_LIST_CASES_ARRAY(appointment_date, case_entry) 		= objWorkRecordSet("ApptDate") 				'ObjWorkExcel.Cells(row, appt_date_col)
		WORKING_LIST_CASES_ARRAY(additional_app_date, case_entry) 	= objWorkRecordSet("SecondApplicationDate") 'ObjWorkExcel.Cells(row, second_app_date_col).Value
		WORKING_LIST_CASES_ARRAY(rept_pnd2_listed_days, case_entry) 	= objWorkRecordSet("REPT_PND2Days") 		'ObjWorkExcel.Cells(row, rept_pnd2_days_col).Value
		WORKING_LIST_CASES_ARRAY(data_days_pend, case_entry) 		= objWorkRecordSet("DaysPending") 		'TODAYS_CASES_ARRAY(data_days_pend, each_case)
        WORKING_LIST_CASES_ARRAY(nomi_sent, case_entry) 				= objWorkRecordSet("NOMIDate") 				'ObjWorkExcel.Cells(row, nomi_date_col)
        WORKING_LIST_CASES_ARRAY(nomi_confirm, case_entry) 			= objWorkRecordSet("Confirmation2") 		'ObjWorkExcel.Cells(row, nomi_confirm_col)
        WORKING_LIST_CASES_ARRAY(next_action_needed, case_entry) 	= objWorkRecordSet("NextActionNeeded") 			'ObjWorkExcel.Cells(row, next_action_col)
		WORKING_LIST_CASES_ARRAY(subsqt_appl_resolve, case_entry)	= objWorkRecordSet("SecondApplicationDateNotes")
		WORKING_LIST_CASES_ARRAY(case_closed_in_30, case_entry)		= objWorkRecordSet("ClosedInPast30Days")
		WORKING_LIST_CASES_ARRAY(closed_in_30_resolve, case_entry)	= objWorkRecordSet("ClosedInPast30DaysNotes")
		WORKING_LIST_CASES_ARRAY(case_in_other_co, case_entry)		= objWorkRecordSet("StartedOutOfCounty")
		WORKING_LIST_CASES_ARRAY(out_of_co_resolve, case_entry)		= objWorkRecordSet("StartedOutOfCountyNotes")
		WORKING_LIST_CASES_ARRAY(script_notes_info, case_entry)		= objWorkRecordSet("TrackingNotes")

        WORKING_LIST_CASES_ARRAY(questionable_intv, case_entry) 		= objWorkRecordSet("QuestionableInterview") 'ObjWorkExcel.Cells(row, quest_intvw_date_col)
		WORKING_LIST_CASES_ARRAY(intvw_quest_resolve, case_entry)	= objWorkRecordSet("Resolved")
		WORKING_LIST_CASES_ARRAY(last_wl_date, case_entry) 			= objWorkRecordSet("AddedtoWorkList")
		change_date_time = objWorkRecordSet("AuditChangeDate")
		change_date_time_array = split(change_date_time, " ")
		change_date = change_date_time_array(0)
		WORKING_LIST_CASES_ARRAY(line_update_date, case_entry) 		= DateAdd("d", 0, change_date)

		' "DenialNeeded"
		' WORKING_LIST_CASES_ARRAY(error_notes, case_entry) 			= objWorkRecordSet("AddedtoWorkList") 'ObjWorkExcel.Cells(row, worker_notes_col)
		' WORKING_LIST_CASES_ARRAY(line_update_date, case_entry) 		= objWorkRecordSet("AddedtoWorkList") 'ObjWorkExcel.Cells(row, script_revw_date_col)
		' WORKING_LIST_CASES_ARRAY(line_update_date, case_entry) = dateAdd("d", 0, WORKING_LIST_CASES_ARRAY(line_update_date, case_entry))

        'Defaulting this values at this time as we will determine them to be different as the script proceeds.
        WORKING_LIST_CASES_ARRAY(take_action_today, case_entry) = FALSE
		WORKING_LIST_CASES_ARRAY(add_to_daily_worklist, case_entry) = False

		For case_info = 0 to UBOUND(ALL_PENDING_CASES_ARRAY, 2)
			If ALL_PENDING_CASES_ARRAY(case_number, case_info) = WORKING_LIST_CASES_ARRAY(case_number, case_entry) Then
				WORKING_LIST_CASES_ARRAY(error_notes, case_entry) = ALL_PENDING_CASES_ARRAY(error_notes, case_info)
			End If
		Next


        case_entry = case_entry + 1     'increasing the count for '
	End If

	objWorkRecordSet.MoveNext
Loop

objWorkRecordSet.Close
objWorkConnection.Close

Set objWorkRecordSet=nothing
Set objWorkConnection=nothing
Set objWorkSQL=nothing


For case_entry = 0 to UBOUND(WORKING_LIST_CASES_ARRAY, 2)

	date_zero =  #1/1/1900#
	If IsDate(WORKING_LIST_CASES_ARRAY(interview_date, case_entry)) = True Then
		If DateDiff("d", WORKING_LIST_CASES_ARRAY(interview_date, case_entry),date_zero) = 0 Then WORKING_LIST_CASES_ARRAY(interview_date, case_entry) = ""
	Else
		WORKING_LIST_CASES_ARRAY(interview_date, case_entry) = ""
	End if
	If IsDate(WORKING_LIST_CASES_ARRAY(additional_app_date, case_entry)) = True Then
		If DateDiff("d", WORKING_LIST_CASES_ARRAY(additional_app_date, case_entry),date_zero) = 0 Then WORKING_LIST_CASES_ARRAY(additional_app_date, case_entry) = ""
	Else
		WORKING_LIST_CASES_ARRAY(additional_app_date, case_entry) = ""
	End if
	If IsDate(WORKING_LIST_CASES_ARRAY(appt_notc_sent, case_entry)) = True Then
		If DateDiff("d", WORKING_LIST_CASES_ARRAY(appt_notc_sent, case_entry),date_zero) = 0 Then WORKING_LIST_CASES_ARRAY(appt_notc_sent, case_entry) = ""
	Else
		WORKING_LIST_CASES_ARRAY(appt_notc_sent, case_entry) = ""
	End if
	If IsDate(WORKING_LIST_CASES_ARRAY(appointment_date, case_entry)) = True Then
		If DateDiff("d", WORKING_LIST_CASES_ARRAY(appointment_date, case_entry),date_zero) = 0 Then WORKING_LIST_CASES_ARRAY(appointment_date, case_entry) = ""
	Else
		WORKING_LIST_CASES_ARRAY(appointment_date, case_entry) = ""
	End if
	If IsDate(WORKING_LIST_CASES_ARRAY(nomi_sent, case_entry)) = True Then
		If DateDiff("d", WORKING_LIST_CASES_ARRAY(nomi_sent, case_entry),date_zero) = 0 Then WORKING_LIST_CASES_ARRAY(nomi_sent, case_entry) = ""
	Else
		WORKING_LIST_CASES_ARRAY(nomi_sent, case_entry) = ""
	End if
	If IsDate(WORKING_LIST_CASES_ARRAY(last_wl_date, case_entry)) = True Then
		If DateDiff("d", WORKING_LIST_CASES_ARRAY(last_wl_date, case_entry),date_zero) = 0 Then WORKING_LIST_CASES_ARRAY(last_wl_date, case_entry) = ""
	Else
		WORKING_LIST_CASES_ARRAY(last_wl_date, case_entry) = ""
	End if
	If WORKING_LIST_CASES_ARRAY(last_wl_date, case_entry) = date Then WORKING_LIST_CASES_ARRAY(add_to_daily_worklist, case_entry) = True
	' If WORKING_LIST_CASES_ARRAY(additional_app_date, case_entry) = date_zero Then WORKING_LIST_CASES_ARRAY(additional_app_date, case_entry) = ""
	' If WORKING_LIST_CASES_ARRAY(appt_notc_sent, case_entry) = date_zero Then WORKING_LIST_CASES_ARRAY(appt_notc_sent, case_entry) = ""
	' If WORKING_LIST_CASES_ARRAY(appointment_date, case_entry) = date_zero Then WORKING_LIST_CASES_ARRAY(appointment_date, case_entry) = ""
	' If WORKING_LIST_CASES_ARRAY(nomi_sent, case_entry) = date_zero Then WORKING_LIST_CASES_ARRAY(nomi_sent, case_entry) = ""
	' If WORKING_LIST_CASES_ARRAY(last_wl_date, case_entry) = date_zero Then WORKING_LIST_CASES_ARRAY(last_wl_date, case_entry) = ""

	If UCase(WORKING_LIST_CASES_ARRAY(add_to_daily_worklist, case_entry)) = "TRUE" Then WORKING_LIST_CASES_ARRAY(add_to_daily_worklist, case_entry) = True
	If UCase(WORKING_LIST_CASES_ARRAY(priv_case, case_entry)) = "TRUE" Then WORKING_LIST_CASES_ARRAY(priv_case, case_entry) = True
	If UCase(WORKING_LIST_CASES_ARRAY(case_in_other_co, case_entry)) = "TRUE" Then WORKING_LIST_CASES_ARRAY(case_in_other_co, case_entry) = True
	If UCase(WORKING_LIST_CASES_ARRAY(case_closed_in_30, case_entry)) = "TRUE" Then WORKING_LIST_CASES_ARRAY(case_closed_in_30, case_entry) = True

	If UCase(WORKING_LIST_CASES_ARRAY(add_to_daily_worklist, case_entry)) = "FALSE" Then WORKING_LIST_CASES_ARRAY(add_to_daily_worklist, case_entry) = False
	If UCase(WORKING_LIST_CASES_ARRAY(priv_case, case_entry)) = "FALSE" Then WORKING_LIST_CASES_ARRAY(priv_case, case_entry) = False
	If UCase(WORKING_LIST_CASES_ARRAY(case_in_other_co, case_entry)) = "FALSE" Then WORKING_LIST_CASES_ARRAY(case_in_other_co, case_entry) = False
	If UCase(WORKING_LIST_CASES_ARRAY(case_closed_in_30, case_entry)) = "FALSE" Then WORKING_LIST_CASES_ARRAY(case_closed_in_30, case_entry) = False

	' If WORKING_LIST_CASES_ARRAY(line_update_date, case_entry) <> date Then

	MAXIS_case_number = WORKING_LIST_CASES_ARRAY(case_number, case_entry)        'setting this so that nav functionality works
	WORKING_LIST_CASES_ARRAY(script_action_taken, case_entry) = False

	' If WORKING_LIST_CASES_ARRAY(next_action_needed, case_entry) = "REVIEW QUESTIONABLE INTERVIEW DATE(S)" Then WORKING_LIST_CASES_ARRAY(next_action_needed, case_entry) = ""
	WORKING_LIST_CASES_ARRAY(error_notes, case_entry) = Replace(WORKING_LIST_CASES_ARRAY(error_notes, case_entry), "Display Limit", "")
	WORKING_LIST_CASES_ARRAY(error_notes, case_entry) = Replace(WORKING_LIST_CASES_ARRAY(error_notes, case_entry), "PRIVILEGED CASE.", "")
	WORKING_LIST_CASES_ARRAY(error_notes, case_entry) = Replace(WORKING_LIST_CASES_ARRAY(error_notes, case_entry), "Cash interview incomplete.", "")
	WORKING_LIST_CASES_ARRAY(error_notes, case_entry) = Replace(WORKING_LIST_CASES_ARRAY(error_notes, case_entry), "SCRIPT DENIAL ALREADY NOTED", "")
	WORKING_LIST_CASES_ARRAY(error_notes, case_entry) = trim(WORKING_LIST_CASES_ARRAY(error_notes, case_entry))

	CALL back_to_SELF
	Call navigate_to_MAXIS_screen_review_PRIV("CASE", "CURR", is_this_priv)
	EMReadScreen WORKING_LIST_CASES_ARRAY(worker_ID, case_entry), 7, 21, 14
	EMReadScreen county_check, 2, 21, 16            'Looking to see if case has Hennepin County worker
    EMReadScreen case_removed_in_MAXIS, 19, 24, 2   'There was one case that was removed from MX and it got a little weird.

	If is_this_priv = True Then
		WORKING_LIST_CASES_ARRAY(priv_case, case_entry) = TRUE
		EMReadScreen WORKING_LIST_CASES_ARRAY(worker_ID, case_entry), 7, 24, 65
		WORKING_LIST_CASES_ARRAY(add_to_daily_worklist, case_entry) = True
		If Instr(WORKING_LIST_CASES_ARRAY(error_notes, case_entry), "PRIVILEGED CASE") = 0 Then WORKING_LIST_CASES_ARRAY(error_notes, case_entry) = "PRIVILEGED CASE. " & WORKING_LIST_CASES_ARRAY(error_notes, case_entry)
	Elseif county_check <> "27" Then
		WORKING_LIST_CASES_ARRAY(out_of_co, case_entry) = "OUT OF COUNTY - " & county_check
	ElseIf case_removed_in_MAXIS = "INVALID CASE NUMBER" Then
		WORKING_LIST_CASES_ARRAY(next_action_needed, case_entry) = "CASE HAS BEEN DELETED"
	Else
		WORKING_LIST_CASES_ARRAY(priv_case, case_entry) = FALSE




		' If WORKING_LIST_CASES_ARRAY(case_closed_in_30, case_entry) = "" Then
		'     'POSSIBLE NEW FUNCTION TO ADD '
		'     WORKING_LIST_CASES_ARRAY(case_closed_in_30, case_entry) = False
		'     EMWriteScreen "X", 4, 9
		'     transmit
		'     prog_hist_row = 8
		'     Do
		'         EMReadScreen prog_name, 4, prog_hist_row, 4
		'         EMReadScreen prog_status, 8, prog_hist_row, 38
		'         ' MsgBox "PROGRAM - " & prog_name & vbCr & "STATUS - *" & prog_status & "*"
		'         If last_prog_name <> prog_name AND prog_status = "INACTIVE" Then
		'             EMReadScreen inactive_date, 8, prog_hist_row, 18
		'             inactive_date = DateAdd("d", 0, inactive_date)
		'
		'             ' MsgBox "PROGRAM - " & prog_name & vbCr & "Inactive Date - " & inactive_date & vbCr & "DATE DIFF - " & DateDiff("m", inactive_date, date)
		'             If DateDiff("m", inactive_date, date) < 1 OR DateDiff("d", inactive_date, date) < 31 Then
		'                 If prog_name <> "  MD" and prog_name <> " QI1" and prog_name <> "SLMB" and prog_name <> " QMB" and prog_name <> "  MA" and prog_name <> "EMER" Then
		'                     WORKING_LIST_CASES_ARRAY(case_closed_in_30, case_entry) = True
		'                     WORKING_LIST_CASES_ARRAY(next_action_needed, case_entry) = "REVIEW RECENT CLOSURE/DENIAL"
		'                     WORKING_LIST_CASES_ARRAY(add_to_daily_worklist, case_entry) = True
		'                     Exit Do
		'                 End If
		'             End If
		'         End If
		'
		'         last_prog_name = prog_name
		'         prog_hist_row = prog_hist_row + 1
		'         If prog_hist_row = 18 Then
		'             PF8
		'             prog_hist_row = 8
		'             EMReadScreen end_of_list, 9, 24, 14
		'         End If
		'     Loop until end_of_list = "LAST PAGE"
		' End If



		Call back_to_SELF
		Call navigate_to_MAXIS_screen("REPT", "PND2")
		EMReadScreen pnd2_disp_limit, 13, 6, 35             'functionality to bypass the display limit warning if it appears.
		If pnd2_disp_limit = "Display Limit" Then
		    TRANSMIT
		    If InStr(WORKING_LIST_CASES_ARRAY(error_notes, case_entry), "Display Limit") = 0 Then WORKING_LIST_CASES_ARRAY(error_notes, case_entry) = WORKING_LIST_CASES_ARRAY(error_notes, case_entry) & " Display Limit"
		    If Instr(list_of_baskets_at_display_limit, WORKING_LIST_CASES_ARRAY(worker_ID, case_entry)) = 0 Then list_of_baskets_at_display_limit = list_of_baskets_at_display_limit & ", " & WORKING_LIST_CASES_ARRAY(worker_ID, case_entry)
		End If
		row = 1                                             'searching for the CASE NUMBER to read from the right row
		col = 1
		EMSearch MAXIS_case_number, row, col
		If row <> 24 and row <> 0 Then
		    EMReadScreen nbr_days_pending, 3, row, 50
		    WORKING_LIST_CASES_ARRAY(rept_pnd2_listed_days, case_entry) = trim(nbr_days_pending)
		    EMReadScreen additional_application_check, 14, row + 1, 17                 'looking to see if this case has a secondary application date entered
		    IF additional_application_check = "ADDITIONAL APP" THEN                         'If it does this string will be at that location and we need to do some handling around the application date to use.
		        multiple_app_dates = True           'identifying that this case has multiple application dates - this is not used specifically yet but is in place so we can output information for managment of case handling in the future.
		        EMReadScreen original_application_date, 8, row, 38               'reading the app date from the other application line
		        EMReadScreen original_cash_code, 1, row, 54
		        EMReadScreen original_snap_code, 1, row, 62
		        EMReadScreen original_emer_code, 1, row, 68
		        EMReadScreen original_grh_code, 1, row, 72
		        use_original_app_date = False
		        If original_cash_code <> "_" Then use_original_app_date = True
		        If original_snap_code <> "_" Then use_original_app_date = True
		        If original_emer_code <> "_" Then use_original_app_date = True
		        If original_grh_code <> "_" Then use_original_app_date = True

		        EMReadScreen additional_application_date, 8, row + 1, 38               'reading the app date from the other application line
		        EMReadScreen additional_cash_code, 1, row + 1, 54
		        EMReadScreen additional_snap_code, 1, row + 1, 62
		        EMReadScreen additional_emer_code, 1, row + 1, 68
		        EMReadScreen additional_grh_code, 1, row + 1, 72
		        use_additional_app_date = False
		        If additional_cash_code <> "_" Then use_additional_app_date = True
		        If additional_snap_code <> "_" Then use_additional_app_date = True
		        If additional_emer_code <> "_" Then use_additional_app_date = True
		        If additional_grh_code <> "_" Then use_additional_app_date = True

		        If use_original_app_date = True AND use_additional_app_date = True Then
		            original_application_date = replace(original_application_date, " ", "/")
		            WORKING_LIST_CASES_ARRAY(application_date, case_entry) = original_application_date
		            additional_application_date = replace(additional_application_date, " ", "/")
		            WORKING_LIST_CASES_ARRAY(additional_app_date, case_entry) = additional_application_date
		            WORKING_LIST_CASES_ARRAY(next_action_needed, case_entry) = "RESOLVE SUBSEQUENT APPLICATION DATE"
		        End If
		        ' WORKING_LIST_CASES_ARRAY(error_notes, case_entry) = additional_application_date & " Please review,  " & WORKING_LIST_CASES_ARRAY(error_notes, case_entry)
		    END IF
		End If



		If WORKING_LIST_CASES_ARRAY(client_name, case_entry) = "XXXXX" Then
			Call navigate_to_MAXIS_screen("STAT", "MEMB")       'go to MEMB - do not need to chose a different memb number because we are looking for the case name
			EMReadScreen last_name, 25, 6, 30       'read each name
			EMReadScreen first_name, 12, 6, 63
			EMReadScreen middle_initial, 1, 6, 79
			last_name = replace(last_name, "_", "") 'format so there are no underscores
			first_name = replace(first_name, "_", "")
			middle_initial = replace(middle_initial, "_", "")

			WORKING_LIST_CASES_ARRAY(client_name, case_entry) = last_name & ", " & first_name & " " & middle_initial     'this is how the BOBI lists names so we want them to match
		End If
		Call back_to_SELF



	End If


Next


Call navigate_to_MAXIS_screen("CASE", "NOTE")       'First to case note to find what has ahppened'
For case_entry = 0 to UBOUND(WORKING_LIST_CASES_ARRAY, 2)
	' If WORKING_LIST_CASES_ARRAY(line_update_date, case_entry) <> date and WORKING_LIST_CASES_ARRAY(priv_case, case_entry) = False Then
	If WORKING_LIST_CASES_ARRAY(priv_case, case_entry) = False Then

		MAXIS_case_number	= WORKING_LIST_CASES_ARRAY(case_number, case_entry)        'setting this so that nav functionality wor
		day_before_app = DateAdd("d", -1, WORKING_LIST_CASES_ARRAY(application_date, case_entry)) 'will set the date one day prior to app date'

		WORKING_LIST_CASES_ARRAY(questionable_intv, case_entry) = trim(WORKING_LIST_CASES_ARRAY(questionable_intv, case_entry))
		If InStr(WORKING_LIST_CASES_ARRAY(questionable_intv, case_entry), "~") <> 0 Then
			start_dates = WORKING_LIST_CASES_ARRAY(questionable_intv, case_entry)            '
		ElseIf WORKING_LIST_CASES_ARRAY(questionable_intv, case_entry) <> "" Then
			Call convert_to_mainframe_date(WORKING_LIST_CASES_ARRAY(questionable_intv, case_entry), 2)
			start_dates = WORKING_LIST_CASES_ARRAY(questionable_intv, case_entry)
		End If


		If WORKING_LIST_CASES_ARRAY(appt_notc_sent, case_entry) = "" OR WORKING_LIST_CASES_ARRAY(nomi_sent, case_entry) = "" Then 'WORKING_LIST_CASES_ARRAY(case_in_other_co, case_entry) = "" OR
			EMWriteScreen "        ", 20, 38
			Call write_value_and_transmit(MAXIS_case_number, 20, 38)	'navigating to the case number and staying in NOTES will bring the view back to the top of the NOTES'
            note_row = 5            'resetting the variables on the loop
            note_date = ""
            note_title = ""
            appt_date = ""
			note_worker = ""
            Do
                EMReadScreen note_date, 8, note_row, 6      'reading the note date
                EMReadScreen note_title, 55, note_row, 25   'reading the note header
				EMReadScreen note_worker, 4, note_row, 16
                note_title = trim(note_title)

				If WORKING_LIST_CASES_ARRAY(appt_notc_sent, case_entry) = "" or WORKING_LIST_CASES_ARRAY(appointment_date, case_entry) = "" Then        'if the ARRAY and Working Excel does not have a date listed for  when the appt notice was sent, the script will go to case ntoes to look for one
                    IF left(note_title, 35) = "~ Appointment letter sent in MEMO ~" then
                        WORKING_LIST_CASES_ARRAY(appt_notc_sent, case_entry) = note_date
    				ElseIF left(note_title, 42) = "~ Appointment letter sent in MEMO for SNAP" then
                        WORKING_LIST_CASES_ARRAY(appt_notc_sent, case_entry) = note_date
    				ElseIF left(note_title, 37) = "~ Appointment letter sent in MEMO for" then
    					EMReadScreen appt_date, 10, note_row, 63
    					appt_date = replace(appt_date, "~", "")
    				 	appt_date = trim(appt_date)
    					WORKING_LIST_CASES_ARRAY(appointment_date, case_entry) = appt_date
                        WORKING_LIST_CASES_ARRAY(appt_notc_sent, case_entry) = note_date
                        'MsgBox WORKING_LIST_CASES_ARRAY(appointment_date, case_entry)
    				END IF
				End If

				' If WORKING_LIST_CASES_ARRAY(case_in_other_co, case_entry) = "" Then
				' 	If note_worker <> "X127" and note_worker <> "MONY" Then WORKING_LIST_CASES_ARRAY(case_in_other_co, case_entry) = True
				' 	WORKING_LIST_CASES_ARRAY(next_action_needed, case_entry) = "REVIEW OTHER COUNTY CASE"
				' End If

				If WORKING_LIST_CASES_ARRAY(nomi_sent, case_entry) = "" Then     'if the date the NOMI was sent is blank in the ARRAY/Working Excel - then we are going to check CASE NOTES for information
					IF note_title = "~ Client missed application interview, NOMI sent via sc" then WORKING_LIST_CASES_ARRAY(nomi_sent, case_entry) = note_date
					IF left(note_title, 32) = "**Client missed SNAP interview**" then WORKING_LIST_CASES_ARRAY(nomi_sent, case_entry) = note_date
					IF left(note_title, 32) = "**Client missed CASH interview**" then WORKING_LIST_CASES_ARRAY(nomi_sent, case_entry) = note_date
					IF left(note_title, 37) = "**Client missed SNAP/CASH interview**" then WORKING_LIST_CASES_ARRAY(nomi_sent, case_entry) = note_date
					IF note_title = "~ Client has not completed application interview, NOMI" then WORKING_LIST_CASES_ARRAY(nomi_sent, case_entry) = note_date
					IF note_title = "~ Client has not completed CASH APP interview, NOMI sen" then WORKING_LIST_CASES_ARRAY(nomi_sent, case_entry) = note_date
					IF note_title = "* A notice was previously sent to client with detail ab" then WORKING_LIST_CASES_ARRAY(nomi_sent, case_entry) = note_date
				End If



                note_row = note_row + 1
                IF note_row = 19 THEN
                    PF8
                    note_row = 5
                END IF
                EMReadScreen next_note_date, 8, note_row, 6
                IF next_note_date = "        " then Exit Do
            Loop until datevalue(next_note_date) < day_before_app 'looking ahead at the next case note kicking out the dates before app'
        End If

		'resetting the action needed based on what is going on with the case if the action needed is not defined
		If WORKING_LIST_CASES_ARRAY(next_action_needed, case_entry) = "" OR WORKING_LIST_CASES_ARRAY(next_action_needed, case_entry) = "SEND APPOINTMENT NOTICE" OR WORKING_LIST_CASES_ARRAY(next_action_needed, case_entry) = "SEND NOMI" Then
			WORKING_LIST_CASES_ARRAY(next_action_needed, case_entry) = "DENY AT DAY 30"
			If WORKING_LIST_CASES_ARRAY(nomi_sent, case_entry) = "" Then WORKING_LIST_CASES_ARRAY(next_action_needed, case_entry) = "SEND NOMI"
			If WORKING_LIST_CASES_ARRAY(appt_notc_sent, case_entry) = "" THen WORKING_LIST_CASES_ARRAY(next_action_needed, case_entry) = "SEND APPOINTMENT NOTICE"
		End If
	End If
Next



For case_entry = 0 to UBOUND(WORKING_LIST_CASES_ARRAY, 2)

	' If WORKING_LIST_CASES_ARRAY(line_update_date, case_entry) <> date and WORKING_LIST_CASES_ARRAY(priv_case, case_entry) = False Then
	If WORKING_LIST_CASES_ARRAY(priv_case, case_entry) = False and WORKING_LIST_CASES_ARRAY(deleted_today, case_entry) = False Then

		MAXIS_case_number	= WORKING_LIST_CASES_ARRAY(case_number, case_entry)        'setting this so that nav functionality wor
		day_before_app = DateAdd("d", -1, WORKING_LIST_CASES_ARRAY(application_date, case_entry)) 'will set the date one day prior to app date'


		'these are for cases where the appointemnt notice sent date is found but the actual appointment date was not found
		'the script will go in to MEMO to read the appointment date from the actual memo.
		If WORKING_LIST_CASES_ARRAY(appt_notc_sent, case_entry) <> "" AND WORKING_LIST_CASES_ARRAY(appointment_date, case_entry) = "" Then
			Call navigate_to_MAXIS_screen ("SPEC", "MEMO")

			'defining the right month to look for the MEMO for as this doesn't work with the NAV functions
			memo_mo = DatePart("m", WORKING_LIST_CASES_ARRAY(appt_notc_sent, case_entry))
			memo_mo = right("00"&memo_mo, 2)
			memo_yr = DatePart("yyyy", WORKING_LIST_CASES_ARRAY(appt_notc_sent, case_entry))
			memo_yr = right(memo_yr, 2)

			EmWriteScreen memo_mo, 3, 48        'writing in the correct footer month and year and going there
			EmWriteScreen memo_yr, 3, 53
			transmit

			'creating a variable in the MM/DD/YY format to compare with date read from MAXIS
			look_date = WORKING_LIST_CASES_ARRAY(appt_notc_sent, case_entry)
			CAll convert_to_mainframe_date(look_date, 2)

			'Loop through all the lines
			Do
				EMReadScreen create_date, 8, memo_row, 19                 'Reading the date of each memo and the status
				EMReadScreen print_status, 7, memo_row, 67
				'MsgBox print_status
				IF create_date = look_date AND print_status = "Printed" Then   'MEMOs created the date the appointment notice was noted and has been printed is likely out memo
					EmWriteScreen "X", memo_row, 16         'opening the memo
					transmit
					PF8                                     'going to the next page

					EMReadScreen start_of_msg, 35, 13, 12    'reading the first line of the message to see if it is the right one
					If start_of_msg = "You applied for assistance in Henne"	Then		'pin County on " & WORKING_LIST_CASES_ARRAY(application_date, case_entry) & "")
					' If start_of_msg = "You recently applied for assistance" Then    'this is how the appt notices start
						EMReadScreen date_in_memo, 10, 16, 50                       'reading the date that was listed in the memo
						date_in_memo = trim(date_in_memo)                           'this formats the date because sometimes dates are 10 chacters and sometimges they are 8
						date_in_memo = replace(date_in_memo, ".", "")
						date_in_memo = replace(date_in_memo, "*", "")
						date_in_memo = trim(date_in_memo)                           'this formats the date because sometimes dates are 10 chacters and sometimges they are 8
						WORKING_LIST_CASES_ARRAY(appointment_date, case_entry) = date_in_memo
						If IsDate(WORKING_LIST_CASES_ARRAY(appointment_date, case_entry)) = False Then WORKING_LIST_CASES_ARRAY(appointment_date, case_entry) =  ""
						WORKING_LIST_CASES_ARRAY(appointment_date, case_entry)  = DateAdd("d", 0, WORKING_LIST_CASES_ARRAY(appointment_date, case_entry))
						Pf3                     'leaving the message and the loop
						Exit Do
					End If
					PF3
				End If
				memo_row = memo_row + 1           'Looking at next row'
			Loop Until create_date = "        "
		End If


        WORKING_LIST_CASES_ARRAY(take_action_today, case_entry) = FALSE      'default this for all cases so that there is no carryover from the previous loop
        If WORKING_LIST_CASES_ARRAY(next_action_needed, case_entry) = "SEND NOMI" AND WORKING_LIST_CASES_ARRAY(appointment_date, case_entry) = "" Then PENDING_CASES_ARRAY(next_action_needed, case_entry) = "Send Manual NOMI"

		'this bit of logic determines if we need to continue looking at the case in STAT
		If WORKING_LIST_CASES_ARRAY(next_action_needed, case_entry) = "" Then
			' WORKING_LIST_CASES_ARRAY(take_action_today, case_entry) = TRUE      'cases where the script doesn't know the next action always needs more information from STAT
			WORKING_LIST_CASES_ARRAY(next_action_needed, case_entry) = "REVIEW NOTICE ACTIONS"
		End If
		If WORKING_LIST_CASES_ARRAY(next_action_needed, case_entry) = "SEND APPOINTMENT NOTICE" Then WORKING_LIST_CASES_ARRAY(take_action_today, case_entry) = TRUE       'Cases where we need to send an appointment notice ALWAYS need further action
		If WORKING_LIST_CASES_ARRAY(next_action_needed, case_entry) = "SEND NOMI" AND IsDate(WORKING_LIST_CASES_ARRAY(appointment_date, case_entry)) = False Then
			WORKING_LIST_CASES_ARRAY(next_action_needed, case_entry) = "Send Manual NOMI"     'If we have to send a NOMI and it is the day before the appointment date - we need to get some additional informaion
		ElseIf IsDate(WORKING_LIST_CASES_ARRAY(appointment_date, case_entry)) = True Then
			If WORKING_LIST_CASES_ARRAY(next_action_needed, case_entry) = "SEND NOMI" AND DateDiff("d", date, WORKING_LIST_CASES_ARRAY(appointment_date, case_entry)) <= 0 Then WORKING_LIST_CASES_ARRAY(take_action_today, case_entry) = TRUE     'If we have to send a NOMI and it is the day before the appointment date - we need to get some additional informaion
		End If
		If WORKING_LIST_CASES_ARRAY(next_action_needed, case_entry) = "DENY AT DAY 30" and DateDiff("d", next_working_day, WORKING_LIST_CASES_ARRAY(data_day_30, case_entry)) = 0 Then
			WORKING_LIST_CASES_ARRAY(take_action_today, case_entry) = TRUE   'If we are going to be denying tomorrow, we need some additional information
			WORKING_LIST_CASES_ARRAY(next_action_needed, case_entry) = "PREP FOR DENIAL"
		End If
		' If WORKING_LIST_CASES_ARRAY(next_action_needed, case_entry) = "" Then MsgBox "Case Number: " & WORKING_LIST_CASES_ARRAY(case_number, case_entry) & vbNewLine & "Does not have an action to take!!!"           'This is here for testing but has never come up

	End If
Next




Call navigate_to_MAXIS_screen("CASE", "NOTE")       'First to case note to find what has ahppened'
For case_entry = 0 to UBOUND(WORKING_LIST_CASES_ARRAY, 2)
	' If WORKING_LIST_CASES_ARRAY(line_update_date, case_entry) <> date and WORKING_LIST_CASES_ARRAY(priv_case, case_entry) = False Then
	If WORKING_LIST_CASES_ARRAY(priv_case, case_entry) = False and WORKING_LIST_CASES_ARRAY(deleted_today, case_entry) = False Then

		MAXIS_case_number	= WORKING_LIST_CASES_ARRAY(case_number, case_entry)        'setting this so that nav functionality wor
		day_before_app = DateAdd("d", -1, WORKING_LIST_CASES_ARRAY(application_date, case_entry)) 'will set the date one day prior to app date'

		'For cases that need an action taken and we do not know an interview date - we will check the case notes for a note that indicates an interview may have happened
		If WORKING_LIST_CASES_ARRAY(take_action_today, case_entry) = TRUE and WORKING_LIST_CASES_ARRAY(interview_date, case_entry) = "" Then
			note_row = 5                                        'setting these for the beginning of the loop to look through all the notes
			start_dates = ""
			day_before_app = DateAdd("d", -1, WORKING_LIST_CASES_ARRAY(application_date, case_entry)) 'will set the date one day prior to app date
			'setting a variable of previously known questionable interview date(s) - this will be used to determine if anything changed
			WORKING_LIST_CASES_ARRAY(questionable_intv, case_entry) = trim(WORKING_LIST_CASES_ARRAY(questionable_intv, case_entry))
			If InStr(WORKING_LIST_CASES_ARRAY(questionable_intv, case_entry), "~") <> 0 Then
				start_dates = WORKING_LIST_CASES_ARRAY(questionable_intv, case_entry)            '
			ElseIf WORKING_LIST_CASES_ARRAY(questionable_intv, case_entry) <> "" Then
				Call convert_to_mainframe_date(WORKING_LIST_CASES_ARRAY(questionable_intv, case_entry), 2)
				start_dates = WORKING_LIST_CASES_ARRAY(questionable_intv, case_entry)
			End If
			Do
				EMWriteScreen "        ", 20, 38
				Call write_value_and_transmit(MAXIS_case_number, 20, 38)	'navigating to the case number and staying in NOTES will bring the view back to the top of the NOTES'

				EMReadScreen note_date, 8, note_row, 6          'read date of the note
				EMReadScreen note_title, 55, note_row, 25       'read the title of the note
				note_title = trim(note_title)
				check_this_date = TRUE                          'setting this as the default.
				IF note_date = "        " THEN EXIT DO
				array_of_dates = ""                             'clearing the array from previous loops
				If InStr(WORKING_LIST_CASES_ARRAY(questionable_intv, case_entry), "~") <> 0 Then             'if there is a ~ that means there is a list of dates
					array_of_dates = split(WORKING_LIST_CASES_ARRAY(questionable_intv, case_entry), "~")     'if there is a list then it should be split in to an array
					If array_of_dates(0) <> "" Then
						For each dates in array_of_dates
							'MsgBox MAXIS_case_number & " - Date 2"
							Call convert_to_mainframe_date(dates, 2)        'Excel always turns dates into m/d/yyyy but MAXIS always displays them as mm/dd/yy and they don't match if they are in these different formats
							'MsgBox "Already known questionable date: " & dates & vbNewLine & "Note Date: " & note_date
							if DateValue(dates) = DateValue(note_date) Then check_this_date = FALSE     'if the date of the note is already known to be a questionable interview then we won't even LOOK at the note title because it has already been reviewed.
						Next
					End If
				Else            'If there is no ~ then it isn't a list - either blank or a single date
					'MsgBox "Already known questionable date: " & dates & vbNewLine & "Note Date: " & note_date
					If WORKING_LIST_CASES_ARRAY(questionable_intv, case_entry) <> "" Then            'If the questionable interview date is not blank
						Call convert_to_mainframe_date(WORKING_LIST_CASES_ARRAY(questionable_intv, case_entry), 2)   'making it mm/dd/yy for comparison
						'MsgBox "Already known questionable date: " & WORKING_LIST_CASES_ARRAY(questionable_intv, case_entry) & vbNewLine & "Note Date: " & note_date
						if DateValue(WORKING_LIST_CASES_ARRAY(questionable_intv, case_entry)) = DateValue(note_date) Then check_this_date = FALSE        'IF the already known questionable interview matches the date of the case notes then we will not assess the note
					End If
				End If

				If check_this_date = TRUE Then 'if a questionable interview date is left on the spreadsheet - that means it has been reviewed and is NOT an interview.
					'All of these notes are used when interviews are done HOWEVER sometimes these notes are made when there is NO interview so we cannot assume the interview has happened - a worker must actually review these questionable interviews
					'We will also add the note date to the list of questionable interviews
					IF left(note_title, 15) = "***Add program:" then
						WORKING_LIST_CASES_ARRAY(questionable_intv, case_entry) = WORKING_LIST_CASES_ARRAY(questionable_intv, case_entry) & "~" & note_date
						WORKING_LIST_CASES_ARRAY(error_notes, case_entry) = WORKING_LIST_CASES_ARRAY(error_notes, case_entry) & ", Case Note suggests interview completed but interview not listed on PROG."
					END IF
					IF left(note_title, 33) = "***Intake Interview Completed ***" then
						WORKING_LIST_CASES_ARRAY(questionable_intv, case_entry) = WORKING_LIST_CASES_ARRAY(questionable_intv, case_entry) & "~" & note_date
						WORKING_LIST_CASES_ARRAY(error_notes, case_entry) = WORKING_LIST_CASES_ARRAY(error_notes, case_entry) & ", Case Note suggests interview completed but interview not listed on PROG."
					END IF
					IF left(note_title, 40) = "***Reapplication Interview Completed ***" then
						WORKING_LIST_CASES_ARRAY(questionable_intv, case_entry) = WORKING_LIST_CASES_ARRAY(questionable_intv, case_entry) & "~" & note_date
						WORKING_LIST_CASES_ARRAY(error_notes, case_entry) = WORKING_LIST_CASES_ARRAY(error_notes, case_entry) & ", Case Note suggests interview completed but interview not listed on PROG."
					END IF
					IF left(note_title, 42) = "~ Interview Completed for SNAP ~" then
						WORKING_LIST_CASES_ARRAY(questionable_intv, case_entry) = WORKING_LIST_CASES_ARRAY(questionable_intv, case_entry) & "~" & note_date
						WORKING_LIST_CASES_ARRAY(error_notes, case_entry) = WORKING_LIST_CASES_ARRAY(error_notes, case_entry) & ", Case Note suggests interview completed but interview not listed on PROG."
					END IF
					IF left(note_title, 42) = "*client interviewed* onboarding processing" then
						WORKING_LIST_CASES_ARRAY(questionable_intv, case_entry) = WORKING_LIST_CASES_ARRAY(questionable_intv, case_entry) & "~" & note_date
						WORKING_LIST_CASES_ARRAY(error_notes, case_entry) = WORKING_LIST_CASES_ARRAY(error_notes, case_entry) & ", Case Note suggests interview completed but interview not listed on PROG."
					END IF
					IF left(note_title, 34) = "***Intake: pending mentor review**" then
						WORKING_LIST_CASES_ARRAY(questionable_intv, case_entry) = WORKING_LIST_CASES_ARRAY(questionable_intv, case_entry) & "~" & note_date
						WORKING_LIST_CASES_ARRAY(error_notes, case_entry) = WORKING_LIST_CASES_ARRAY(error_notes, case_entry) & ", Case Note suggests interview completed but interview not listed on PROG."
					END IF
					IF left(note_title, 23) = "~ Interview Completed ~" then
						WORKING_LIST_CASES_ARRAY(questionable_intv, case_entry) = WORKING_LIST_CASES_ARRAY(questionable_intv, case_entry) & "~" & note_date
						WORKING_LIST_CASES_ARRAY(error_notes, case_entry) = WORKING_LIST_CASES_ARRAY(error_notes, case_entry) & ", Case Note suggests interview completed but interview not listed on PROG."
					END IF
					IF left(note_title, 10) = "***Intake:" then
						WORKING_LIST_CASES_ARRAY(questionable_intv, case_entry) = WORKING_LIST_CASES_ARRAY(questionable_intv, case_entry) & "~" & note_date
						WORKING_LIST_CASES_ARRAY(error_notes, case_entry) = WORKING_LIST_CASES_ARRAY(error_notes, case_entry) & ", Case Note suggests interview completed but interview not listed on PROG."
					END IF
					IF left(note_title, 24) = "~ Application interview ~" then
						WORKING_LIST_CASES_ARRAY(questionable_intv, case_entry) = WORKING_LIST_CASES_ARRAY(questionable_intv, case_entry) & "~" & note_date
						WORKING_LIST_CASES_ARRAY(error_notes, case_entry) = WORKING_LIST_CASES_ARRAY(error_notes, case_entry) & ", This case may not require an interview."
					END IF
					IF left(note_title, 33) = "***Intake Interview Completed ***" then
						WORKING_LIST_CASES_ARRAY(questionable_intv, case_entry) = WORKING_LIST_CASES_ARRAY(questionable_intv, case_entry) & "~" & note_date
						WORKING_LIST_CASES_ARRAY(error_notes, case_entry) = WORKING_LIST_CASES_ARRAY(error_notes, case_entry) & ", Case Note suggests interview completed but interview not listed on PROG."
					END IF
					IF left(UCase(note_title), 51) = "Phone call from client re: Phone interview Complete" then
						WORKING_LIST_CASES_ARRAY(questionable_intv, case_entry) = WORKING_LIST_CASES_ARRAY(questionable_intv, case_entry) & "~" & note_date
						WORKING_LIST_CASES_ARRAY(error_notes, case_entry) = WORKING_LIST_CASES_ARRAY(error_notes, case_entry) & ", Possible case note indicating an interview. If an interview happened, PROG was not updated and an incorrect script was used."
					END IF
					IF left(UCase(note_title), 41) = "Phone call from client re: SNAP interview" then
						WORKING_LIST_CASES_ARRAY(questionable_intv, case_entry) = WORKING_LIST_CASES_ARRAY(questionable_intv, case_entry) & "~" & note_date
						WORKING_LIST_CASES_ARRAY(error_notes, case_entry) = WORKING_LIST_CASES_ARRAY(error_notes, case_entry) & ", Possible case note indicating an interview. If an interview happened, PROG was not updated and an incorrect script was used."
					END IF
				End If

				IF note_date = "        " then Exit Do      'for newer cases we might meet the end of the case notes before the date is prior to the app date - this accounts for that
				note_row = note_row + 1                     'go to the next row
				IF note_row = 19 THEN                       'go to the next page if the end of the page has been reached
					PF8
					note_row = 5
				END IF
				EMReadScreen next_note_date, 8, note_row, 6     'read what note is next to know when to exit
				IF next_note_date = "        " then Exit Do
			Loop until datevalue(next_note_date) < day_before_app 'looking ahead at the next case note kicking out the dates before app'

			If left(WORKING_LIST_CASES_ARRAY(questionable_intv, case_entry), 1) = "~" Then WORKING_LIST_CASES_ARRAY(questionable_intv, case_entry) = right(WORKING_LIST_CASES_ARRAY(questionable_intv, case_entry), len(WORKING_LIST_CASES_ARRAY(questionable_intv, case_entry))-1)     'triming off the left most ~ of the questionale interview dates
			if WORKING_LIST_CASES_ARRAY(questionable_intv, case_entry) <> start_dates Then WORKING_LIST_CASES_ARRAY(next_action_needed, case_entry) = "REVIEW QUESTIONABLE INTERVIEW DATE(S)"     'if a new date was added it needs to be reviewed but if they are the same then we know they have been reviewed and we can continue with the correct action

		End If
	End If
Next



For case_entry = 0 to UBOUND(WORKING_LIST_CASES_ARRAY, 2)
	' If WORKING_LIST_CASES_ARRAY(line_update_date, case_entry) <> date and WORKING_LIST_CASES_ARRAY(priv_case, case_entry) = False Then
	If WORKING_LIST_CASES_ARRAY(priv_case, case_entry) = False and WORKING_LIST_CASES_ARRAY(deleted_today, case_entry) = False Then

		MAXIS_case_number	= WORKING_LIST_CASES_ARRAY(case_number, case_entry)        'setting this so that nav functionality wor
		CALL back_to_SELF

		forms_to_swkr = ""      'setting these for the start a memo function
		forms_to_arep = ""
		memo_started = TRUE

		If WORKING_LIST_CASES_ARRAY(take_action_today, case_entry) = TRUE Then       'only the cases that we have determined need something today
			'TODO add MEMB for written language information

			' Call navigate_to_MAXIS_screen("STAT", "MEMB")
			' EMReadScreen language_code, 2, 13, 42
			' WORKING_LIST_CASES_ARRAY(written_lang, case_entry) = language_code
			WORKING_LIST_CASES_ARRAY(written_lang, case_entry) = "99"

			IF WORKING_LIST_CASES_ARRAY(CASH_status, case_entry) = "Pending" then           'setting the language for the notices - Cash or SNAP or both
				if WORKING_LIST_CASES_ARRAY(SNAP_status, case_entry) = "Pending" then
					programs = "CASH/SNAP"
				else
					programs = "CASH"
				end if
			else
				programs = "SNAP"
			end if


			'Cases needing an Appointment Notice
			If WORKING_LIST_CASES_ARRAY(next_action_needed, case_entry) = "SEND APPOINTMENT NOTICE" Then
				need_intv_date = dateadd("d", 5, WORKING_LIST_CASES_ARRAY(application_date, case_entry))       'setting the appointment date - it should be 7 days from the date of application
				If need_intv_date <= date then need_intv_date = dateadd("d", 5, date)         'if this is today or in the past then we reset this for 7 days from today
				Call change_date_to_soonest_working_day(need_intv_date, "FORWARD")
				last_contact_day = dateadd("d", 30, WORKING_LIST_CASES_ARRAY(application_date, case_entry))       'setting the date to enter on the NOMI of the day of denial
				'ensuring that we have given the client an additional10days fromt he day nomi sent'
				IF DateDiff("d", need_intv_date, last_contact_day) < 1 then last_contact_day = need_intv_date

				need_intv_date = need_intv_date & ""		'turns interview date into string for variable

				' Call start_a_new_spec_memo(memo_started, search_for_arep_and_swkr, forms_to_arep, forms_to_swkr, send_to_other, other_name, other_street, other_city, other_state, other_zip, end_script)
				Call start_a_new_spec_memo(memo_started, True, forms_to_arep, forms_to_swkr, "N", other_name, other_street, other_city, other_state, other_zip, False)
				IF memo_started = True THEN
					'TODO - add languages in when we can'
					Call write_variable_in_SPEC_MEMO("You applied for assistance in Hennepin County on " & WORKING_LIST_CASES_ARRAY(application_date, case_entry) & "")
					Call write_variable_in_SPEC_MEMO("and an interview is required to process your application.")
					Call write_variable_in_SPEC_MEMO(" ")
					Call write_variable_in_SPEC_MEMO("** The interview must be completed by " & need_intv_date & ". **")
					Call write_variable_in_SPEC_MEMO("To complete a phone interview, call the EZ Info Line at")
					Call write_variable_in_SPEC_MEMO("612-596-1300 between 8:00am and 4:30pm Monday thru Friday.")
					Call write_variable_in_SPEC_MEMO(" ")
					Call write_variable_in_SPEC_MEMO("* You may be able to have SNAP benefits issued within 24 hours of the interview.")
					Call write_variable_in_SPEC_MEMO(" ")
					Call write_variable_in_SPEC_MEMO("  ** If we do not hear from you by " & last_contact_day & " **")
					Call write_variable_in_SPEC_MEMO("  **    your application will be denied.     **")
					Call write_variable_in_SPEC_MEMO(" ")
					CALL write_variable_in_SPEC_MEMO("All interviews are completed via phone. If you do not have a phone, go to one of our Digital Access Spaces at any Hennepin County Library or Service Center. No processing, no interviews are completed at these sites. Some Options:")
					CALL write_variable_in_SPEC_MEMO(" - 7051 Brooklyn Blvd Brooklyn Center 55429")
					CALL write_variable_in_SPEC_MEMO(" - 1011 1st St S Hopkins 55343")
					CALL write_variable_in_SPEC_MEMO(" - 1001 Plymouth Ave N Minneapolis 55411")
					CALL write_variable_in_SPEC_MEMO(" - 2215 East Lake Street Minneapolis 55407")
					CALL write_variable_in_SPEC_MEMO(" (Hours are 8 - 4:30 Monday - Friday)")
					CALL write_variable_in_SPEC_MEMO("*** Submitting Documents:")
					CALL write_variable_in_SPEC_MEMO("- Online at infokeep.hennepin.us or MNBenefits.mn.gov")
					CALL write_variable_in_SPEC_MEMO("  Use InfoKeep to upload documents directly to your case.")
					CALL write_variable_in_SPEC_MEMO("- Mail, Fax, or Drop Boxes at service centers(listed above)")
					Call write_variable_in_SPEC_MEMO(" ")
					CALL write_variable_in_SPEC_MEMO("Domestic violence brochures are available at this website: https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3477-ENG. You can always request a paper copy via phone.")

					PF4
				ELSE
					WORKING_LIST_CASES_ARRAY(appt_notc_confirm, case_entry) = "N" 'Setting this as N if the MEMO failed
				END IF

				If WORKING_LIST_CASES_ARRAY(appt_notc_confirm, case_entry) <> "N" Then Call confirm_memo_waiting(WORKING_LIST_CASES_ARRAY(appt_notc_confirm, case_entry))       'reading that a MEMO exists to confirm the notice went
				If WORKING_LIST_CASES_ARRAY(appt_notc_confirm, case_entry) = "N" Then              'if the MEMO failed we need to send it manually
					WORKING_LIST_CASES_ARRAY(next_action_needed, case_entry) = "Send Manual Appt Notice"
				ElseIf WORKING_LIST_CASES_ARRAY(appt_notc_confirm, case_entry) = "Y" Then
					'if the memo was successful then we will changed the next action needed and we will create a case note
					WORKING_LIST_CASES_ARRAY(next_action_needed, case_entry) = "SEND NOMI"
					WORKING_LIST_CASES_ARRAY(appt_notc_sent, case_entry) = date
					WORKING_LIST_CASES_ARRAY(appointment_date, case_entry) = need_intv_date        'adding this date to the appointment date in the ARRAY

					Call start_a_blank_case_note
					Call write_variable_in_CASE_NOTE("~ Appointment letter sent in MEMO for " & need_intv_date & "~")
					Call write_variable_in_CASE_NOTE("* A notice has been sent via SPEC/MEMO informing the client of needed interview.")
					Call write_variable_in_CASE_NOTE("* Households failing to complete the interview within 30 days of the date they file an application will receive a denial notice")
					Call write_variable_in_CASE_NOTE("* A link to the domestic violence brochure sent to client in SPEC/MEMO as a part of interview notice.")
					Call write_variable_in_CASE_NOTE("---")
					Call write_variable_in_CASE_NOTE(worker_signature)
					'MsgBox "What casenote was sent?"
					PF3
				Else
					WORKING_LIST_CASES_ARRAY(next_action_needed, case_entry) = "???"       'if the memo confirm is not N or Y then this next action holder is here for testing
				End If
				Call back_to_SELF

				'Adding the notice to the array of cases taken action on today
				ReDim Preserve ACTION_TODAY_CASES_ARRAY(error_notes, todays_cases)
				ACTION_TODAY_CASES_ARRAY(case_number, todays_cases)         = WORKING_LIST_CASES_ARRAY(case_number, case_entry)
				ACTION_TODAY_CASES_ARRAY(client_name, todays_cases)         = WORKING_LIST_CASES_ARRAY(client_name, case_entry)
				ACTION_TODAY_CASES_ARRAY(worker_ID, todays_cases)           = WORKING_LIST_CASES_ARRAY(worker_ID, case_entry)
				ACTION_TODAY_CASES_ARRAY(SNAP_status, todays_cases)         = WORKING_LIST_CASES_ARRAY(SNAP_status, case_entry)
				ACTION_TODAY_CASES_ARRAY(CASH_status, todays_cases)         = WORKING_LIST_CASES_ARRAY(CASH_status, case_entry)
				ACTION_TODAY_CASES_ARRAY(application_date, todays_cases)    = WORKING_LIST_CASES_ARRAY(application_date, case_entry)
				ACTION_TODAY_CASES_ARRAY(interview_date, todays_cases)      = WORKING_LIST_CASES_ARRAY(interview_date, case_entry)
				ACTION_TODAY_CASES_ARRAY(questionable_intv, todays_cases)   = WORKING_LIST_CASES_ARRAY(questionable_intv, case_entry)
				ACTION_TODAY_CASES_ARRAY(appt_notc_sent, todays_cases)      = WORKING_LIST_CASES_ARRAY(appt_notc_sent, case_entry)
				ACTION_TODAY_CASES_ARRAY(appt_notc_confirm, todays_cases)   = WORKING_LIST_CASES_ARRAY(appt_notc_confirm, case_entry)
				ACTION_TODAY_CASES_ARRAY(appointment_date, todays_cases)    = WORKING_LIST_CASES_ARRAY(appointment_date, case_entry)
				ACTION_TODAY_CASES_ARRAY(nomi_sent, todays_cases)           = WORKING_LIST_CASES_ARRAY(nomi_sent, case_entry)
				ACTION_TODAY_CASES_ARRAY(nomi_confirm, todays_cases)        = WORKING_LIST_CASES_ARRAY(nomi_confirm, case_entry)
				ACTION_TODAY_CASES_ARRAY(deny_day30, todays_cases)          = WORKING_LIST_CASES_ARRAY(deny_day30, case_entry)
				ACTION_TODAY_CASES_ARRAY(deny_memo_confirm, todays_cases)   = WORKING_LIST_CASES_ARRAY(deny_memo_confirm, case_entry)
				ACTION_TODAY_CASES_ARRAY(next_action_needed, todays_cases)  = WORKING_LIST_CASES_ARRAY(next_action_needed, case_entry)
				ACTION_TODAY_CASES_ARRAY(error_notes, todays_cases)         = "Appointment Notice Sent today"
				todays_cases = todays_cases + 1       'increasing the counter for the array

				WORKING_LIST_CASES_ARRAY(script_action_taken, case_entry) = True

			ElseIf WORKING_LIST_CASES_ARRAY(next_action_needed, case_entry) = "SEND NOMI" Then       'These cases need NOMIs


				If IsDate(WORKING_LIST_CASES_ARRAY(application_date, case_entry)) = False Then
					WORKING_LIST_CASES_ARRAY(next_action_needed, case_entry) = "Send Manual NOMI"
				Else
					nomi_last_contact_day = dateadd("d", 30, WORKING_LIST_CASES_ARRAY(application_date, case_entry))       'setting the date to enter on the NOMI of the day of denial
					'ensuring that we have given the client an additional10days fromt he day nomi sent'
					IF DateDiff("d", date, nomi_last_contact_day) < 1 then nomi_last_contact_day = dateadd("d", 10, date)
					nomi_last_contact_day = nomi_last_contact_day & ""		'turns interview date into string for variable

					Call start_a_new_spec_memo(memo_started, True, forms_to_arep, forms_to_swkr, "N", other_name, other_street, other_city, other_state, other_zip, False)
					IF memo_started = True THEN
						'TODO - add languages in when we can'

						Call write_variable_in_SPEC_MEMO("You recently applied for assistance on " & WORKING_LIST_CASES_ARRAY(application_date, case_entry) & ".")
						Call write_variable_in_SPEC_MEMO("Your interview should have been completed by " & WORKING_LIST_CASES_ARRAY(appointment_date, case_entry) & ".")
						Call write_variable_in_SPEC_MEMO("An interview is required to process your application.")
						Call write_variable_in_SPEC_MEMO(" ")
						Call write_variable_in_SPEC_MEMO("To complete a phone interview, call the EZ Info Line at")
						Call write_variable_in_SPEC_MEMO("612-596-1300 between 8:00am and 4:30pm Monday thru Friday.")
						Call write_variable_in_SPEC_MEMO(" ")
						Call write_variable_in_SPEC_MEMO("* You may be able to have SNAP benefits issued within 24 hours of the interview.")
						Call write_variable_in_SPEC_MEMO(" ")
						Call write_variable_in_SPEC_MEMO("  ** If we do not hear from you by " & nomi_last_contact_day & " **")
						Call write_variable_in_SPEC_MEMO("  **    your application will be denied.     **") 'add 30 days
						Call write_variable_in_SPEC_MEMO(" ")
						CALL write_variable_in_SPEC_MEMO("All interviews are completed via phone. If you do not have a phone, go to one of our Digital Access Spaces at any Hennepin County Library or Service Center. No processing, no interviews are completed at these sites. Some Options:")
						CALL write_variable_in_SPEC_MEMO(" - 7051 Brooklyn Blvd Brooklyn Center 55429")
						CALL write_variable_in_SPEC_MEMO(" - 1011 1st St S Hopkins 55343")
						CALL write_variable_in_SPEC_MEMO(" - 1001 Plymouth Ave N Minneapolis 55411")
						CALL write_variable_in_SPEC_MEMO(" - 2215 East Lake Street Minneapolis 55407")
						CALL write_variable_in_SPEC_MEMO(" (Hours are 8 - 4:30 Monday - Friday)")
						CALL write_variable_in_SPEC_MEMO(" More detail can be found at hennepin.us/economic-supports")
						CALL write_variable_in_SPEC_MEMO("")
						CALL write_variable_in_SPEC_MEMO("*** Submitting Documents:")
						CALL write_variable_in_SPEC_MEMO("- Online at infokeep.hennepin.us or MNBenefits.mn.gov")
						CALL write_variable_in_SPEC_MEMO("  Use InfoKeep to upload documents directly to your case.")
						CALL write_variable_in_SPEC_MEMO("- Mail, Fax, or Drop Boxes at service centers(listed above)")

						PF4
					Else
						WORKING_LIST_CASES_ARRAY(nomi_confirm, case_entry) = "N"   'if the MEMO didn't start then setting this for the ARRAY and Working Excel.
					End If

					If WORKING_LIST_CASES_ARRAY(nomi_confirm, case_entry) <> "N" Then Call confirm_memo_waiting(WORKING_LIST_CASES_ARRAY(nomi_confirm, case_entry))     'reading the SPEC/MEMO page to see that a MEMO for today is waiting.

					'Resetting the next action needed based on message success and writing the case note if successful
					If WORKING_LIST_CASES_ARRAY(nomi_confirm, case_entry) = "N" Then
						WORKING_LIST_CASES_ARRAY(next_action_needed, case_entry) = "Send Manual NOMI"
						WORKING_LIST_CASES_ARRAY(nomi_sent, case_entry) = ""
					ElseIf WORKING_LIST_CASES_ARRAY(nomi_confirm, case_entry) = "Y" Then
						WORKING_LIST_CASES_ARRAY(next_action_needed, case_entry) = "DENY AT DAY 30"
						WORKING_LIST_CASES_ARRAY(nomi_sent, case_entry) = date

						Call start_a_blank_case_note
						Call write_variable_in_CASE_NOTE("~ Client has not completed application interview, NOMI sent ~ ")
						Call write_variable_in_CASE_NOTE("* A notice was previously sent to client with detail about completing an interview. ")
						Call write_variable_in_CASE_NOTE("* Households failing to complete the interview within 30 days of the date they file an application will receive a denial notice")
						Call write_variable_in_CASE_NOTE("* A link to the domestic violence brochure sent to client in SPEC/MEMO as a part of interview notice.")
						Call write_variable_in_CASE_NOTE("---")
						Call write_variable_in_CASE_NOTE(worker_signature)
						'MsgBox "What casenote was sent?"
						PF3
					Else
						WORKING_LIST_CASES_ARRAY(next_action_needed, case_entry) = "???"           'this is for testing - this has never come up
					End If
					Call back_to_SELF

					'Adding this case to the list of cases that we took action on today
					ReDim Preserve ACTION_TODAY_CASES_ARRAY(error_notes, todays_cases)
					ACTION_TODAY_CASES_ARRAY(case_number, todays_cases)         = WORKING_LIST_CASES_ARRAY(case_number, case_entry)
					ACTION_TODAY_CASES_ARRAY(client_name, todays_cases)         = WORKING_LIST_CASES_ARRAY(client_name, case_entry)
					ACTION_TODAY_CASES_ARRAY(worker_ID, todays_cases)           = WORKING_LIST_CASES_ARRAY(worker_ID, case_entry)
					ACTION_TODAY_CASES_ARRAY(SNAP_status, todays_cases)         = WORKING_LIST_CASES_ARRAY(SNAP_status, case_entry)
					ACTION_TODAY_CASES_ARRAY(CASH_status, todays_cases)         = WORKING_LIST_CASES_ARRAY(CASH_status, case_entry)
					ACTION_TODAY_CASES_ARRAY(application_date, todays_cases)    = WORKING_LIST_CASES_ARRAY(application_date, case_entry)
					ACTION_TODAY_CASES_ARRAY(interview_date, todays_cases)      = WORKING_LIST_CASES_ARRAY(interview_date, case_entry)
					ACTION_TODAY_CASES_ARRAY(questionable_intv, todays_cases)   = WORKING_LIST_CASES_ARRAY(questionable_intv, case_entry)
					ACTION_TODAY_CASES_ARRAY(appt_notc_sent, todays_cases)      = WORKING_LIST_CASES_ARRAY(appt_notc_sent, case_entry)
					ACTION_TODAY_CASES_ARRAY(appt_notc_confirm, todays_cases)   = WORKING_LIST_CASES_ARRAY(appt_notc_confirm, case_entry)
					ACTION_TODAY_CASES_ARRAY(appointment_date, todays_cases)    = WORKING_LIST_CASES_ARRAY(appointment_date, case_entry)
					ACTION_TODAY_CASES_ARRAY(nomi_sent, todays_cases)           = WORKING_LIST_CASES_ARRAY(nomi_sent, case_entry)
					ACTION_TODAY_CASES_ARRAY(nomi_confirm, todays_cases)        = WORKING_LIST_CASES_ARRAY(nomi_confirm, case_entry)
					ACTION_TODAY_CASES_ARRAY(deny_day30, todays_cases)          = WORKING_LIST_CASES_ARRAY(deny_day30, case_entry)
					ACTION_TODAY_CASES_ARRAY(deny_memo_confirm, todays_cases)   = WORKING_LIST_CASES_ARRAY(deny_memo_confirm, case_entry)
					ACTION_TODAY_CASES_ARRAY(next_action_needed, todays_cases)  = WORKING_LIST_CASES_ARRAY(next_action_needed, case_entry)
					ACTION_TODAY_CASES_ARRAY(error_notes, todays_cases)         = "NOMI Sent today"
					todays_cases = todays_cases + 1

					WORKING_LIST_CASES_ARRAY(script_action_taken, case_entry) = True
				End If

			ElseIf WORKING_LIST_CASES_ARRAY(next_action_needed, case_entry) = "DENY AT DAY 30" Then
				IF datediff("d", WORKING_LIST_CASES_ARRAY(application_date, case_entry), date) >= 30 and WORKING_LIST_CASES_ARRAY(interview_date, case_entry) = "" Then
					WORKING_LIST_CASES_ARRAY(case_over_30_days, case_entry) = True
					day_30 = dateadd("d", 30, WORKING_LIST_CASES_ARRAY(application_date, case_entry))
					If WORKING_LIST_CASES_ARRAY(appt_notc_sent, case_entry) = "" Then
						reason_cannot_deny = "The Appointment Notice has not been sent."
						WORKING_LIST_CASES_ARRAY(next_action_needed, case_entry) = "REVIEW CANNOT DENY - No Appt Notc"
					ElseIf WORKING_LIST_CASES_ARRAY(nomi_sent, case_entry) = "" Then
						reason_cannot_deny = "The NOMI has not been sent."
						WORKING_LIST_CASES_ARRAY(next_action_needed, case_entry) = "REVIEW CANNOT DENY - No NOMI"
					ElseIf DateDiff("d", day_30, WORKING_LIST_CASES_ARRAY(nomi_sent, case_entry)) >= 0 AND datediff("d", WORKING_LIST_CASES_ARRAY(nomi_sent, case_entry), date) < 10 Then
						reason_cannot_deny = "NOMIs sent on or after Day 30 cannot be denied until 10 days from the date the NOMI is sent."
						WORKING_LIST_CASES_ARRAY(next_action_needed, case_entry) = "REVIEW CANNOT DENY - NOMI after Day 30"
					End If
					'TODO - add a CNOTE about not being able to deny yet.
				End If
			End If
		End If
	End If

	If IsNumeric(WORKING_LIST_CASES_ARRAY(rept_pnd2_listed_days, case_entry)) = True Then
		days_pending_nbr = WORKING_LIST_CASES_ARRAY(rept_pnd2_listed_days, case_entry) * 1
		For next_day = 0 to number_of_days_until_next_working_day
			If days_pending_nbr + next_day = 30 Then
				WORKING_LIST_CASES_ARRAY(add_to_daily_worklist, case_entry) = True
				WORKING_LIST_CASES_ARRAY(next_action_needed, case_entry) = "PREP FOR DENIAL"
			End If
		Next
	Else
		' WORKING_LIST_CASES_ARRAY(add_to_daily_worklist, case_entry) = True
	End If
	IF WORKING_LIST_CASES_ARRAY(next_action_needed, case_entry) = "REVIEW QUESTIONABLE INTERVIEW DATE(S)" THEN WORKING_LIST_CASES_ARRAY(add_to_daily_worklist, case_entry) = True
	IF WORKING_LIST_CASES_ARRAY(next_action_needed, case_entry) = "REVIEW OTHER COUNTY CASE"	Then WORKING_LIST_CASES_ARRAY(add_to_daily_worklist, case_entry) = True
	IF WORKING_LIST_CASES_ARRAY(next_action_needed, case_entry) = "RESOLVE SUBSEQUENT APPLICATION DATE" Then WORKING_LIST_CASES_ARRAY(add_to_daily_worklist, case_entry) = True
	IF WORKING_LIST_CASES_ARRAY(next_action_needed, case_entry) = "ALIGN INTERVIEW DATES" Then WORKING_LIST_CASES_ARRAY(add_to_daily_worklist, case_entry) = True
	If WORKING_LIST_CASES_ARRAY(next_action_needed, case_entry) = "REVIEW RECENT CLOSURE/DENIAL" Then WORKING_LIST_CASES_ARRAY(add_to_daily_worklist, case_entry) = True
	If WORKING_LIST_CASES_ARRAY(next_action_needed, case_entry) = "REVIEW NOTICE ACTIONS" Then WORKING_LIST_CASES_ARRAY(add_to_daily_worklist, case_entry) = True

	IF WORKING_LIST_CASES_ARRAY(next_action_needed, case_entry) = "PREP FOR DENIAL" Then WORKING_LIST_CASES_ARRAY(add_to_daily_worklist, case_entry) = True
	IF WORKING_LIST_CASES_ARRAY(next_action_needed, case_entry) = "REVIEW CANNOT DENY - No Appt Notc" Then WORKING_LIST_CASES_ARRAY(add_to_daily_worklist, case_entry) = True
	IF WORKING_LIST_CASES_ARRAY(next_action_needed, case_entry) = "REVIEW CANNOT DENY - No NOMI" Then WORKING_LIST_CASES_ARRAY(add_to_daily_worklist, case_entry) = True
	IF WORKING_LIST_CASES_ARRAY(next_action_needed, case_entry) = "REVIEW CANNOT DENY - NOMI after Day 30" Then WORKING_LIST_CASES_ARRAY(add_to_daily_worklist, case_entry) = True
	IF WORKING_LIST_CASES_ARRAY(next_action_needed, case_entry) = "Send Manual Appt Notice" THEN WORKING_LIST_CASES_ARRAY(add_to_daily_worklist, case_entry) = True
	IF WORKING_LIST_CASES_ARRAY(next_action_needed, case_entry) = "Send Manual NOMI" THEN WORKING_LIST_CASES_ARRAY(add_to_daily_worklist, case_entry) = True
	If WORKING_LIST_CASES_ARRAY(priv_case, case_entry) = True Then WORKING_LIST_CASES_ARRAY(add_to_daily_worklist, case_entry) = True

	IF WORKING_LIST_CASES_ARRAY(next_action_needed, case_entry) = "???" THEN
		WORKING_LIST_CASES_ARRAY(add_to_daily_worklist, case_entry) = True
		cases_to_alert_BZST = cases_to_alert_BZST & ", " & MAXIS_case_number
	End If
	If WORKING_LIST_CASES_ARRAY(deleted_today, case_entry) = True Then WORKING_LIST_CASES_ARRAY(add_to_daily_worklist, case_entry) = False
	If WORKING_LIST_CASES_ARRAY(add_to_daily_worklist, case_entry) = True Then WORKING_LIST_CASES_ARRAY(last_wl_date, case_entry) = date
	' objWorkRecordSet.Open "DELETE FROM ES.ES_OnDemanCashAndSnapBZProcessed WHERE CaseNumber = '" & case_number_to_review & "'", objWorkConnection
	' ' objWorkRecordSet.Close
	' ' objWorkConnection.Close
	' objWorkRecordSet.Open "INSERT INTO ES.ES_OnDemanCashAndSnapBZProcessed (CaseNumber, CaseName, ApplDate, InterviewDate, Day_30, DaysPending, SnapStatus, CashStatus, SecondApplicationDate, REPT_PND2Days, QuestionableInterview, Resolved, ApptNoticeDate, ApptDate, Confirmation, NOMIDate, Confirmation2, DenialNeeded, NextActionNeeded, AddedtoWorkList)" & _
    '                   "VALUES ('" & MAXIS_case_number &  "', '" & _
    '                                 WORKING_LIST_CASES_ARRAY(client_name, case_entry) &  "', '" & _
    '                                 WORKING_LIST_CASES_ARRAY(application_date, case_entry) &  "', '" & _
    '                                 WORKING_LIST_CASES_ARRAY(additional_app_date, case_entry) &  "', '" & _
    '                                 WORKING_LIST_CASES_ARRAY(data_day_30, case_entry) &  "', '" & _
    '                                 WORKING_LIST_CASES_ARRAY(data_days_pend, case_entry) &  "', '" & _
    '                                 WORKING_LIST_CASES_ARRAY(SNAP_status, case_entry) &  "', '" & _
    '                                 WORKING_LIST_CASES_ARRAY(CASH_status, case_entry) &  "', '" & _
    '                                 WORKING_LIST_CASES_ARRAY(additional_app_date, case_entry) &  "', '" & _
    '                                 WORKING_LIST_CASES_ARRAY(rept_pnd2_listed_days, case_entry) &  "', '" & _
    '                                 WORKING_LIST_CASES_ARRAY(questionable_intv, case_entry) &  "', '" & _
    '                                 WORKING_LIST_CASES_ARRAY(intvw_quest_resolve, case_entry) &  "', '" & _
    '                                 WORKING_LIST_CASES_ARRAY(appt_notc_sent, case_entry) &  "', '" & _
    '                                 WORKING_LIST_CASES_ARRAY(appointment_date, case_entry) &  "', '" & _
    '                                 WORKING_LIST_CASES_ARRAY(appt_notc_confirm, case_entry) &  "', '" & _
    '                                 WORKING_LIST_CASES_ARRAY(nomi_sent, case_entry) &  "', '" & _
    '                                 WORKING_LIST_CASES_ARRAY(nomi_confirm, case_entry) &  "', '" & _
    '                                 WORKING_LIST_CASES_ARRAY(case_over_30_days, case_entry) &  "', '" & _
    '                                 WORKING_LIST_CASES_ARRAY(next_action_needed, case_entry) &  "', '" & _
    '                                 WORKING_LIST_CASES_ARRAY(last_wl_date, case_entry) & "')", objWorkConnection, adOpenStatic, adLockOptimistic
	' objWorkRecordSet.Close
	' objWorkConnection.Close

Next

On Error Resume Next
'declare the SQL statement that will query the database
objWorkSQL = "SELECT * FROM ES.ES_OnDemanCashAndSnapBZProcessed"

'Creating objects for Access
Set objWorkConnection = CreateObject("ADODB.Connection")
Set objWorkRecordSet = CreateObject("ADODB.Recordset")

'This is the file path for the statistics Access database.
objWorkConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
' objWorkRecordSet.Open objWorkSQL, objWorkConnection

objWorkRecordSet.Open "DELETE FROM ES.ES_OnDemanCashAndSnapBZProcessed", objWorkConnection, adOpenStatic, adLockOptimistic

objWorkRecordSet.Close
objWorkConnection.Close

Set objWorkRecordSet=nothing
Set objWorkConnection=nothing
Set objWorkSQL=nothing


'declare the SQL statement that will query the database
objWorkSQL = "SELECT * FROM ES.ES_OnDemanCashAndSnapBZProcessed"

'Creating objects for Access
Set objWorkConnection = CreateObject("ADODB.Connection")
Set objWorkRecordSet = CreateObject("ADODB.Recordset")

'This is the file path for the statistics Access database.
objWorkConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""

For case_entry = 0 to UBOUND(WORKING_LIST_CASES_ARRAY, 2)
	' If WORKING_LIST_CASES_ARRAY(line_update_date, case_entry) <> date and WORKING_LIST_CASES_ARRAY(priv_case, case_entry) = False Then
	If WORKING_LIST_CASES_ARRAY(deleted_today, case_entry) = False Then

		objWorkRecordSet.Open "INSERT INTO ES.ES_OnDemanCashAndSnapBZProcessed (CaseNumber, CaseName, ApplDate, InterviewDate, Day_30, DaysPending, SnapStatus, CashStatus, SecondApplicationDate, REPT_PND2Days, QuestionableInterview, Resolved, ApptNoticeDate, ApptDate, Confirmation, NOMIDate, Confirmation2, DenialNeeded, NextActionNeeded, AddedtoWorkList, SecondApplicationDateNotes, ClosedInPast30Days, ClosedInPast30DaysNotes, StartedOutOfCounty, StartedOutOfCountyNotes, TrackingNotes)" & _
						  "VALUES ('" & WORKING_LIST_CASES_ARRAY(case_number, case_entry) &  "', '" & _
										WORKING_LIST_CASES_ARRAY(client_name, case_entry) &  "', '" & _
										WORKING_LIST_CASES_ARRAY(application_date, case_entry) &  "', '" & _
										WORKING_LIST_CASES_ARRAY(additional_app_date, case_entry) &  "', '" & _
										WORKING_LIST_CASES_ARRAY(data_day_30, case_entry) &  "', '" & _
										WORKING_LIST_CASES_ARRAY(data_days_pend, case_entry) &  "', '" & _
										WORKING_LIST_CASES_ARRAY(SNAP_status, case_entry) &  "', '" & _
										WORKING_LIST_CASES_ARRAY(CASH_status, case_entry) &  "', '" & _
										WORKING_LIST_CASES_ARRAY(additional_app_date, case_entry) &  "', '" & _
										WORKING_LIST_CASES_ARRAY(rept_pnd2_listed_days, case_entry) &  "', '" & _
										WORKING_LIST_CASES_ARRAY(questionable_intv, case_entry) &  "', '" & _
										WORKING_LIST_CASES_ARRAY(intvw_quest_resolve, case_entry) &  "', '" & _
										WORKING_LIST_CASES_ARRAY(appt_notc_sent, case_entry) &  "', '" & _
										WORKING_LIST_CASES_ARRAY(appointment_date, case_entry) &  "', '" & _
										WORKING_LIST_CASES_ARRAY(appt_notc_confirm, case_entry) &  "', '" & _
										WORKING_LIST_CASES_ARRAY(nomi_sent, case_entry) &  "', '" & _
										WORKING_LIST_CASES_ARRAY(nomi_confirm, case_entry) &  "', '" & _
										WORKING_LIST_CASES_ARRAY(case_over_30_days, case_entry) &  "', '" & _
										WORKING_LIST_CASES_ARRAY(next_action_needed, case_entry) &  "', '" & _
										WORKING_LIST_CASES_ARRAY(last_wl_date, case_entry) &  "', '" & _
										WORKING_LIST_CASES_ARRAY(subsqt_appl_resolve, case_entry) &  "', '" & _
										WORKING_LIST_CASES_ARRAY(case_closed_in_30, case_entry) &  "', '" & _
										WORKING_LIST_CASES_ARRAY(closed_in_30_resolve, case_entry) &  "', '" & _
										WORKING_LIST_CASES_ARRAY(case_in_other_co, case_entry) &  "', '" & _
										WORKING_LIST_CASES_ARRAY(out_of_co_resolve, case_entry) &  "', '" & _
										WORKING_LIST_CASES_ARRAY(script_notes_info, case_entry) & "')", objWorkConnection, adOpenStatic, adLockOptimistic

	End If
Next
objWorkRecordSet.Close
objWorkConnection.Close

Set objWorkRecordSet=nothing
Set objWorkConnection=nothing
Set objWorkSQL=nothing


If warning_checkbox = 1 then MsgBox "Do not engage with any applications while Excel is outputing the Excel lists for the On Demand Waiver Applications Assignment." & vbcr & vbcr & "Press OK when you're ready to continue.",64, "Excel Output is Ready"   'Warning to staff re: Excel Output


date_month = DatePart("m", date)
date_day = DatePart("d", date)
date_year = DatePart("yyyy", date)
date_header = date_month & "-" & date_day & "-" & date_year
worksheet_header = "Work List for " & date_header

daily_worklist_template_path = t_drive & "/Eligibility Support/Restricted/QI - Quality Improvement/REPORTS/On Demand Waiver/QI On Demand Daily Assignment/Archive/Worklist Template.xlsx"
daily_worklist_path = t_drive & "/Eligibility Support/Restricted/QI - Quality Improvement/REPORTS/On Demand Waiver/QI On Demand Daily Assignment/QI " & date_header & " Worklist.xlsx"

call excel_open(daily_worklist_template_path, True, True, ObjDailyWorkListExcel, objWorkListWorkbook)

ObjDailyWorkListExcel.ActiveWorkbook.SaveAs daily_worklist_path

ObjDailyWorkListExcel.worksheets("CASE LIST").Activate
ObjDailyWorkListExcel.ActiveSheet.Name = worksheet_header

xl_row = 2
count_cases_on_wl = 0
For case_entry = 0 to UBound(WORKING_LIST_CASES_ARRAY, 2)
	If WORKING_LIST_CASES_ARRAY(add_to_daily_worklist, case_entry) = True Then
		count_cases_on_wl = count_cases_on_wl + 1

		ObjDailyWorkListExcel.Cells(xl_row, worker_id_col).Value 				= WORKING_LIST_CASES_ARRAY(worker_ID, case_entry)
		ObjDailyWorkListExcel.Cells(xl_row, case_nbr_col).Value 				= WORKING_LIST_CASES_ARRAY(case_number, case_entry)
		ObjDailyWorkListExcel.Cells(xl_row, case_name_col).Value 				= WORKING_LIST_CASES_ARRAY(client_name, case_entry)
		ObjDailyWorkListExcel.Cells(xl_row, snap_stat_col).Value 				= WORKING_LIST_CASES_ARRAY(SNAP_status, case_entry)
		ObjDailyWorkListExcel.Cells(xl_row, cash_stat_col).Value 				= WORKING_LIST_CASES_ARRAY(CASH_status, case_entry)
		ObjDailyWorkListExcel.Cells(xl_row, wl_rept_pnd2_days_col).Value 		= WORKING_LIST_CASES_ARRAY(rept_pnd2_listed_days, case_entry)
		ObjDailyWorkListExcel.Cells(xl_row, wl_app_date_col).Value 				= WORKING_LIST_CASES_ARRAY(application_date, case_entry)
		ObjDailyWorkListExcel.Cells(xl_row, wl_second_app_date_col).Value 		= WORKING_LIST_CASES_ARRAY(additional_app_date, case_entry)
		ObjDailyWorkListExcel.Cells(xl_row, wl_resolve_2nd_app_date_col).Value	= WORKING_LIST_CASES_ARRAY(subsqt_appl_resolve, case_entry)

		ObjDailyWorkListExcel.Cells(xl_row, wl_intvw_date_col).Value 			= WORKING_LIST_CASES_ARRAY(interview_date, case_entry)
		ObjDailyWorkListExcel.Cells(xl_row, wl_quest_intvw_date_col).Value 		= WORKING_LIST_CASES_ARRAY(questionable_intv, case_entry)
		ObjDailyWorkListExcel.Cells(xl_row, wl_resolve_quest_intvw_col).Value	= WORKING_LIST_CASES_ARRAY(intvw_quest_resolve, case_entry)

		ObjDailyWorkListExcel.Cells(xl_row, wl_other_county_col).Value 			= WORKING_LIST_CASES_ARRAY(case_in_other_co, case_entry)
		ObjDailyWorkListExcel.Cells(xl_row, wl_resolve_othr_co_col).Value		= WORKING_LIST_CASES_ARRAY(out_of_co_resolve, case_entry)
		ObjDailyWorkListExcel.Cells(xl_row, wl_closed_in_30_col).Value 			= WORKING_LIST_CASES_ARRAY(case_closed_in_30, case_entry)
		ObjDailyWorkListExcel.Cells(xl_row, wl_resolve_closed_in_30_col).Value	= WORKING_LIST_CASES_ARRAY(closed_in_30_resolve, case_entry)
		ObjDailyWorkListExcel.Cells(xl_row, wl_appt_notc_date_col).Value 		= WORKING_LIST_CASES_ARRAY(appt_notc_sent, case_entry)
		ObjDailyWorkListExcel.Cells(xl_row, wl_appt_date_col).Value 			= WORKING_LIST_CASES_ARRAY(appointment_date, case_entry)
		ObjDailyWorkListExcel.Cells(xl_row, wl_nomi_date_col).Value 			= WORKING_LIST_CASES_ARRAY(nomi_sent, case_entry)
		ObjDailyWorkListExcel.Cells(xl_row, wl_day_30_col).Value				= WORKING_LIST_CASES_ARRAY(data_day_30, case_entry)
		' If left(WORKING_LIST_CASES_ARRAY(next_action_needed, case_entry), 18) = "REVIEW CANNOT DENY"  Then ObjDailyWorkListExcel.Cells(xl_row, wl_cannot_deny_col).Value = "True"

		ObjDailyWorkListExcel.Cells(xl_row, wl_action_taken_col).Value			= "FOLLOW UP NEEDED"
		ObjDailyWorkListExcel.Cells(xl_row, wl_work_notes_col).Value = WORKING_LIST_CASES_ARRAY(next_action_needed, case_entry) & " - " & WORKING_LIST_CASES_ARRAY(error_notes, case_entry)


		xl_row = xl_row + 1
	End If
Next


' For yest_entry = 0 to UBound(YESTERDAYS_PENDING_CASES_ARRAY, 2)
' 	case_found_on_working_list = False
' 	For case_entry = 0 to UBOUND(WORKING_LIST_CASES_ARRAY, 2)
' 		If WORKING_LIST_CASES_ARRAY(case_number, case_entry) = YESTERDAYS_PENDING_CASES_ARRAY(case_number, yest_entry) Then
' 			If WORKING_LIST_CASES_ARRAY(deleted_today, case_entry) = False Then case_found_on_working_list = True
' 		End If
' 	Next
' 	If case_found_on_working_list = False AND Instr(YESTERDAYS_PENDING_CASES_ARRAY(error_notes, yest_entry), "REVIEW CANNOT DENY") <> 0 Then
' 		ObjDailyWorkListExcel.Cells(xl_row, worker_id_col).Value 				= YESTERDAYS_PENDING_CASES_ARRAY(worker_ID, yest_entry)
' 		ObjDailyWorkListExcel.Cells(xl_row, case_nbr_col).Value 				= YESTERDAYS_PENDING_CASES_ARRAY(case_number, yest_entry)
' 		ObjDailyWorkListExcel.Cells(xl_row, case_name_col).Value 				= YESTERDAYS_PENDING_CASES_ARRAY(client_name, yest_entry)
' 		ObjDailyWorkListExcel.Cells(xl_row, snap_stat_col).Value 				= YESTERDAYS_PENDING_CASES_ARRAY(SNAP_status, yest_entry)
' 		ObjDailyWorkListExcel.Cells(xl_row, cash_stat_col).Value 				= YESTERDAYS_PENDING_CASES_ARRAY(CASH_status, yest_entry)
' 		ObjDailyWorkListExcel.Cells(xl_row, wl_rept_pnd2_days_col).Value 		= YESTERDAYS_PENDING_CASES_ARRAY(rept_pnd2_listed_days, yest_entry)
' 		ObjDailyWorkListExcel.Cells(xl_row, wl_app_date_col).Value 				= YESTERDAYS_PENDING_CASES_ARRAY(application_date, yest_entry)
' 		ObjDailyWorkListExcel.Cells(xl_row, wl_second_app_date_col).Value 		= YESTERDAYS_PENDING_CASES_ARRAY(additional_app_date, yest_entry)
' 		ObjDailyWorkListExcel.Cells(xl_row, wl_resolve_2nd_app_date_col).Value	= YESTERDAYS_PENDING_CASES_ARRAY(subsqt_appl_resolve, yest_entry)
'
' 		ObjDailyWorkListExcel.Cells(xl_row, wl_intvw_date_col).Value 			= YESTERDAYS_PENDING_CASES_ARRAY(interview_date, yest_entry)
' 		ObjDailyWorkListExcel.Cells(xl_row, wl_quest_intvw_date_col).Value 		= YESTERDAYS_PENDING_CASES_ARRAY(questionable_intv, yest_entry)
' 		ObjDailyWorkListExcel.Cells(xl_row, wl_resolve_quest_intvw_col).Value	= YESTERDAYS_PENDING_CASES_ARRAY(intvw_quest_resolve, yest_entry)
'
' 		ObjDailyWorkListExcel.Cells(xl_row, wl_other_county_col).Value 			= YESTERDAYS_PENDING_CASES_ARRAY(case_in_other_co, yest_entry)
' 		ObjDailyWorkListExcel.Cells(xl_row, wl_resolve_othr_co_col).Value		= YESTERDAYS_PENDING_CASES_ARRAY(out_of_co_resolve, yest_entry)
' 		ObjDailyWorkListExcel.Cells(xl_row, wl_closed_in_30_col).Value 			= YESTERDAYS_PENDING_CASES_ARRAY(case_closed_in_30, yest_entry)
' 		ObjDailyWorkListExcel.Cells(xl_row, wl_resolve_closed_in_30_col).Value	= YESTERDAYS_PENDING_CASES_ARRAY(closed_in_30_resolve, yest_entry)
' 		ObjDailyWorkListExcel.Cells(xl_row, wl_appt_notc_date_col).Value 		= YESTERDAYS_PENDING_CASES_ARRAY(appt_notc_sent, yest_entry)
' 		ObjDailyWorkListExcel.Cells(xl_row, wl_appt_date_col).Value 			= YESTERDAYS_PENDING_CASES_ARRAY(appointment_date, yest_entry)
' 		ObjDailyWorkListExcel.Cells(xl_row, wl_nomi_date_col).Value 			= YESTERDAYS_PENDING_CASES_ARRAY(nomi_sent, yest_entry)
' 		ObjDailyWorkListExcel.Cells(xl_row, wl_day_30_col).Value				= YESTERDAYS_PENDING_CASES_ARRAY(data_day_30, yest_entry)
'
' 		ObjDailyWorkListExcel.Cells(xl_row, wl_work_notes_col).Value = "CHECK DENIAL - " & YESTERDAYS_PENDING_CASES_ARRAY(error_notes, yest_entry)
' 	End If
' Next


qi_worklist_threshold_reached = False
' If count_cases_on_wl > 99 Then qi_worklist_threshold_reached = True

ObjDailyWorkListExcel.Worksheets("Statistics").visible = True
ObjDailyWorkListExcel.worksheets("Statistics").Activate
ObjDailyWorkListExcel.Cells(2, 2).Value = qi_member_on_ONDEMAND
ObjDailyWorkListExcel.Cells(3, 2).Value = date
ObjDailyWorkListExcel.Cells(4, 2).Value = time

ObjDailyWorkListExcel.worksheets(worksheet_header).Activate
ObjDailyWorkListExcel.Worksheets("Statistics").visible = False

objWorkListWorkbook.Save
ObjDailyWorkListExcel.Quit

' objWorkbook.Save
' MsgBox "Step Five - going to do the stats"


this_year = DatePart("yyyy", date)
this_month = MonthName(Month(date))

statistics_excel_file_path = "T:\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\On Demand Waiver\Applications Statistics\" & this_year & " Statistics Tracking.xlsx"
call excel_open(statistics_excel_file_path, True,  True, ObjStatsExcel, objStatsWorkbook)

'Now we need to open the right worksheet
'Select Case MonthName(Month(#2/15/21#)) 'will need to be updated for future dates and tracking
sheet_selection = this_month & " " & this_year

'Activates worksheet based on user selection
ObjStatsExcel.worksheets(sheet_selection).Activate   'activates the stat worksheet.'

stats_excel_nomi_row = 3
Do
    this_entry = ObjStatsExcel.Cells(stats_excel_nomi_row, 1).Value
    this_entry = trim(this_entry)
    If this_entry <> "" Then stats_excel_nomi_row = stats_excel_nomi_row + 1
Loop until this_entry = ""

stats_excel_email_row = 3
Do
    this_entry = ObjStatsExcel.Cells(stats_excel_email_row, 8).Value
    this_entry = trim(this_entry)
    If this_entry <> "" Then stats_excel_email_row = stats_excel_email_row + 1
Loop until this_entry = ""

' 'For testing'
' ObjStatsExcel.Visible = TRUE
' MsgBox "NOMI Row - " & stats_excel_nomi_row & vbNewLine & "Email row - " & stats_excel_email_row
For action_case = 0 to UBOUND(ACTION_TODAY_CASES_ARRAY, 2)      'looping through the ARRAY created when we took actions on the cases on the Working Excel
    If InStr(ACTION_TODAY_CASES_ARRAY(error_notes, action_case), "NOMI Sent today") <> 0 Then
        'Here we add the NOMI to the statistics
        ObjStatsExcel.Cells(stats_excel_nomi_row, 1).Value = ACTION_TODAY_CASES_ARRAY(case_number, action_case)        'Adding the case number to the statistics sheet
        ObjStatsExcel.Cells(stats_excel_nomi_row, 2).Value = ACTION_TODAY_CASES_ARRAY(application_date, action_case)   'Adding the date of application to the statistics sheet
        ObjStatsExcel.Cells(stats_excel_nomi_row, 3).Value = date                                                    'Adding today's date of the NOMI date for the stats sheet
        stats_excel_nomi_row = stats_excel_nomi_row + 1
    End If
Next

For case_removed = 0 to UBOUND(CASES_NO_LONGER_WORKING, 2)      'looping through each of the cases in the ARRAY from the beginning of cases that were taken off of the Working Excel
    If CASES_NO_LONGER_WORKING(worker_name_one, case_removed) <> "" OR CASES_NO_LONGER_WORKING(issue_item_one, case_removed) <> "" OR CASES_NO_LONGER_WORKING(qi_worker_one, case_removed) <> "" Then
        ObjStatsExcel.Cells(stats_excel_email_row, 8).Value = CASES_NO_LONGER_WORKING(case_number, case_removed)        'Adding all information to the stats excel
        ObjStatsExcel.Cells(stats_excel_email_row, 9).Value = CASES_NO_LONGER_WORKING(worker_name_one, case_removed)
        ObjStatsExcel.Cells(stats_excel_email_row, 10).Value = CASES_NO_LONGER_WORKING(sup_name_one, case_removed)
        ObjStatsExcel.Cells(stats_excel_email_row, 11).Value = CASES_NO_LONGER_WORKING(issue_item_one, case_removed)
        ObjStatsExcel.Cells(stats_excel_email_row, 12).Value = CASES_NO_LONGER_WORKING(email_ym_one, case_removed)
        ObjStatsExcel.Cells(stats_excel_email_row, 13).Value = CASES_NO_LONGER_WORKING(qi_worker_one, case_removed)
        ObjStatsExcel.Cells(stats_excel_email_row, 14).Value = CASES_NO_LONGER_WORKING(worker_name_two, case_removed)
        ObjStatsExcel.Cells(stats_excel_email_row, 15).Value = CASES_NO_LONGER_WORKING(sup_name_two, case_removed)
        ObjStatsExcel.Cells(stats_excel_email_row, 16).Value = CASES_NO_LONGER_WORKING(issue_item_two, case_removed)
        ObjStatsExcel.Cells(stats_excel_email_row, 17).Value = CASES_NO_LONGER_WORKING(email_ym_two, case_removed)
        ObjStatsExcel.Cells(stats_excel_email_row, 18).Value = CASES_NO_LONGER_WORKING(qi_worker_two, case_removed)
        ObjStatsExcel.Cells(stats_excel_email_row, 19).Value = CASES_NO_LONGER_WORKING(worker_name_three, case_removed)
        ObjStatsExcel.Cells(stats_excel_email_row, 20).Value = CASES_NO_LONGER_WORKING(sup_name_three, case_removed)
        ObjStatsExcel.Cells(stats_excel_email_row, 21).Value = CASES_NO_LONGER_WORKING(issue_item_three, case_removed)
        ObjStatsExcel.Cells(stats_excel_email_row, 22).Value = CASES_NO_LONGER_WORKING(email_ym_three, case_removed)
        ObjStatsExcel.Cells(stats_excel_email_row, 23).Value = CASES_NO_LONGER_WORKING(qi_worker_three, case_removed)
        stats_excel_email_row = stats_excel_email_row + 1
    End If
Next


'Now the script reopens the daily list that was identified in the beginning
file_date = replace(current_date, "/", "-")   'Changing the format of the date to use as file path selection default
daily_case_list_folder = right("0" & DatePart("m", file_date), 2) & "-" & DatePart("yyyy", file_date)
file_selection_path = t_drive & "/Eligibility Support/Restricted/QI - Quality Improvement/REPORTS/On Demand Waiver/Daily case lists/" & daily_case_list_folder & "/" & file_date & ".xlsx" 'single assignment file

call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)



'On the same Excel file the script creates a new sheet and names it
ObjExcel.Worksheets.Add().Name = "Actions Today"

'Header row is added
ObjExcel.Cells(1, worker_id_col)         = "Worker ID"
ObjExcel.Cells(1, case_nbr_col)          = "Case Number"
ObjExcel.Cells(1, case_name_col)         = "Case Name"
ObjExcel.Cells(1, snap_stat_col)         = "SNAP"
ObjExcel.Cells(1, cash_stat_col)         = "CASH"
ObjExcel.Cells(1, app_date_col)          = "Application Date"
ObjExcel.Cells(1, second_app_date_col)	= "Second App Date"
ObjExcel.Cells(1, rept_pnd2_days_col)	= "REPT/PND2 Days"
ObjExcel.Cells(1, intvw_date_col)        = "Interview Date"
ObjExcel.Cells(1, quest_intvw_date_col)  = "Questionable Interview Date"
ObjExcel.Cells(1, appt_notc_date_col)    = "Appt Notice Sent"
ObjExcel.Cells(1, appt_date_col)         = "Appointment Date"
ObjExcel.Cells(1, appt_notc_confirm_col) = "Confirm"
ObjExcel.Cells(1, nomi_date_col)         = "NOMI Sent"
ObjExcel.Cells(1, nomi_confirm_col)      = "Confirm"
ObjExcel.Cells(1, need_deny_col)         = "Denial"
ObjExcel.Cells(1, next_action_col)       = "Next Action"
ObjExcel.Cells(1, worker_notes_col)      = "Detail"
' ObjExcel.Cells(1, action_worker_col)    =
' ObjExcel.Cells(1, action_sup_col)       =
' ObjExcel.Cells(1, email_sent_col)       =

ObjExcel.Rows(1).Font.Bold = TRUE       'header row is bold

action_row = 2      'setting the first row


For action_case = 0 to UBOUND(ACTION_TODAY_CASES_ARRAY, 2)      'looping through the ARRAY created when we took actions on the cases on the Working Excel

    'removing leading separators
    IF ACTION_TODAY_CASES_ARRAY(error_notes, action_case) <> "" AND left(ACTION_TODAY_CASES_ARRAY(error_notes, action_case), 3) = " - " THEN ACTION_TODAY_CASES_ARRAY(error_notes, action_case) = right(ACTION_TODAY_CASES_ARRAY(error_notes, action_case), len(ACTION_TODAY_CASES_ARRAY(error_notes, action_case))- 3)

    'adding the information from the ARRAY to the spreadsheet
    ObjExcel.Cells(action_row, worker_id_col)        	= ACTION_TODAY_CASES_ARRAY(worker_ID, action_case)
    ObjExcel.Cells(action_row, case_nbr_col)         	= ACTION_TODAY_CASES_ARRAY(case_number, action_case)
    ObjExcel.Cells(action_row, case_name_col)        	= ACTION_TODAY_CASES_ARRAY(client_name, action_case)
    ObjExcel.Cells(action_row, snap_stat_col)        	= ACTION_TODAY_CASES_ARRAY(SNAP_status, action_case)
    ObjExcel.Cells(action_row, cash_stat_col)        	= ACTION_TODAY_CASES_ARRAY(CASH_status, action_case)
	ObjExcel.Cells(action_row, app_date_col)         	= ACTION_TODAY_CASES_ARRAY(application_date, action_case)
	ObjExcel.Cells(action_row, second_app_date_col)		= ACTION_TODAY_CASES_ARRAY(additional_app_date, action_case)
    ObjExcel.Cells(action_row, rept_pnd2_days_col)		= ACTION_TODAY_CASES_ARRAY(rept_pnd2_listed_days, action_case)
    ObjExcel.Cells(action_row, intvw_date_col)       	= ACTION_TODAY_CASES_ARRAY(interview_date, action_case)
	ObjExcel.Cells(action_row, quest_intvw_date_col) 	= ACTION_TODAY_CASES_ARRAY(questionable_intv, action_case)

	ObjExcel.Cells(action_row, other_county_col)		= ACTION_TODAY_CASES_ARRAY(case_in_other_co, action_case)
	ObjExcel.Cells(action_row, closed_in_30_col)		= ACTION_TODAY_CASES_ARRAY(case_closed_in_30, action_case)

	' ObjExcel.Cells(action_row, resolve_quest_intvw_col)	= ACTION_TODAY_CASES_ARRAY(intvw_quest_resolve, action_case)

	ObjExcel.Cells(action_row, appt_notc_date_col)   	= ACTION_TODAY_CASES_ARRAY(appt_notc_sent, action_case)
    ObjExcel.Cells(action_row, appt_notc_confirm_col)	= ACTION_TODAY_CASES_ARRAY(appt_notc_confirm, action_case)
    ObjExcel.Cells(action_row, appt_date_col)        	= ACTION_TODAY_CASES_ARRAY(appointment_date, action_case)
    ObjExcel.Cells(action_row, nomi_date_col)        	= ACTION_TODAY_CASES_ARRAY(nomi_sent, action_case)
    ObjExcel.Cells(action_row, nomi_confirm_col)     	= ACTION_TODAY_CASES_ARRAY(nomi_confirm, action_case)
    ObjExcel.Cells(action_row, need_deny_col)        	= ACTION_TODAY_CASES_ARRAY(deny_day30, action_case)
    ObjExcel.Cells(action_row, next_action_col)      	= ACTION_TODAY_CASES_ARRAY(next_action_needed, action_case)
    ObjExcel.Cells(action_row, worker_notes_col)     	= ACTION_TODAY_CASES_ARRAY(error_notes, action_case)
    action_row = action_row + 1     'go to the next row
Next

For col_to_autofit =1 to  worker_notes_col      'formatting the sheet
    ObjExcel.Columns(col_to_autofit).AutoFit()
Next

'Saving the Daily List
objWorkbook.Save
ObjExcel.Quit

objStatsWorkbook.Save
ObjStatsExcel.Quit





' MsgBox "Step Six - The emails, the emails, what what, the emails"
qi_member_email = replace(qi_member_on_ONDEMAND, " ", ".") & "@hennepin.us"
cc_email = "tanya.payne@hennepin.us; hsph.ews.bluezonescripts@hennepin.us"
cc_email = "hsph.ews.bluezonescripts@hennepin.us"
If qi_worklist_threshold_reached = True Then cc_email = "HSPH.EWS.QUALITYIMPROVEMENT@hennepin.us; tanya.payne@hennepin.us"

email_subject = "On Demand List is Ready"
If qi_worklist_threshold_reached = True Then email_subject = email_subject & " - HELP NEEDED"
email_body = "Hello " & qi_member_on_ONDEMAND & "," & vbCr & vbCr
email_body = email_body & "The worklist is completed and ready to be worked. All cases on the list should be reveiwed." & vbCr
email_body = email_body & "There are " & count_cases_on_wl & " cases on the worklist." & vbCr
' email_body = email_body & "There are " & count_denials & " DENIALS on the worklist." & vbCr
If qi_worklist_threshold_reached = True Then email_body = email_body & "As the list is so large, help has been requested via email to the QUALITY IMPROVEMENT email. If you are NOT on the assignment today and have capacity to assist, contact " & qi_member_on_ONDEMAND & "." & vbCr
email_body = email_body & "Access the Worklist here: "
email_body = email_body & vbCr & "<" & daily_worklist_path & ">" & vbCr
email_body = email_body & "Please contact Tanya if you have issues with the list or questions about the assignment." & vbCr & vbCr
email_body = email_body & "Thank you!"

Call create_outlook_email(qi_member_email, cc_email, email_subject, email_body, "", True)

If list_of_baskets_at_display_limit <> "" Then
	If left(list_of_baskets_at_display_limit, 1) = "," Then list_of_baskets_at_display_limit = right(list_of_baskets_at_display_limit, len(list_of_baskets_at_display_limit)-1)
	list_of_baskets_at_display_limit = trim(list_of_baskets_at_display_limit)
	basket_email_subject = "Baskets at Display Limit for PND2"
	basket_email_body = "Good morning," & vbCr & vbCr
	basket_email_body = basket_email_body & "It appears there are some baskets in which the REPT/PND2 Display limit has been reached." & vbCr
	basket_email_body = basket_email_body & "The following baskets were identified during the On Demand Application process:" & vbCr
	basket_email_body = basket_email_body & list_of_baskets_at_display_limit & vbCr & vbCr
	' basket_email_body = basket_email_body & "" & vbCr
	basket_email_body = basket_email_body & "Thank you!" & vbCr
	Call create_outlook_email("Faughn.Ramisch-Church@hennepin.us", "hsph.ews.bluezonescripts@hennepin.us", basket_email_subject, basket_email_body, "", True)
End If

If cases_to_alert_BZST <> "" Then
	If left(cases_to_alert_BZST, 1) = "," Then cases_to_alert_BZST = right(cases_to_alert_BZST, len(cases_to_alert_BZST)-1)
	cases_to_alert_BZST = trim(cases_to_alert_BZST)
	Call create_outlook_email("hsph.ews.bluezonescripts@hennepin.us", "", "ON DEMAND Could not determine next action needed", "These cases have an unknown issue... " & cases_to_alert_BZST, "", True)
End If

If does_file_exist = True then objFSO.MoveFile previous_list_file_selection_path , archive_files & "\QI " & previous_date_header & " Worklist.xlsx"    'moving each file to the archive file

end_msg = "The Daily On Demand Assignment has been created. Emails have been sent regarding the case discovery and work to be completed." & vbCr & vbCr & "The worklist generated today has " & count_cases_on_wl & " cases."
script_end_procedure_with_error_report(end_msg)  'WE'RE DONE!
