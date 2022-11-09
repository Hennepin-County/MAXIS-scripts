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

'
' objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
' objRecordSet.Open objSQL, objConnection

'NOTE about using the SQL - I am adding information to the 'NextActionNeeded' column in the table -- this will be set as an array for now
	'information will be seperated by '~-*-~'
	' 0 -- Next Action Needed
	' 1 -- Worker Notes
	' 2 -- Script Notes
	' 3 -- Case was in other county
	' 4 -- Case closed in past 30 days
	' 5 -- PRIV Case
	' 6 -- Out of county resolved
	' 7 -- closed in 30 days resolved
	' 8 -- Subsequent Application resolved
	'


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
			' ElseIf anything_number = current_number Then    'this is if we are looking at the same case still
			'     'Checking to see if one of the later lines for the case indicates no interview = this will make the array show no interview if EITHER Cash or SNAP have no interview indicated in PROG
			'     If trim(objExcel.cells(row, 8).value) = "" Then TODAYS_CASES_ARRAY(interview_date, case_entry-1) = ""
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

		            'MsgBox "Excel case number: " & case_number_to_assess & vbNewLine & "Array case number: " & TODAYS_CASES_ARRAY(case_number, each_case)
		            ' If ObjWorkExcel.Cells(row, next_action_col) = "REMOVE FROM LIST" Then       'These cases were flagged on the Working Excel to be removed - usually because neither CASH or SNAP are pending any more.
		            '     'MsgBox "REMOVE FROM LIST"
		            '     ReDim Preserve CASES_NO_LONGER_WORKING(error_notes, case_removed)           'It is removed from the working list and added to an ARRAY of all the cases removed from the working list that day.
		            '     CASES_NO_LONGER_WORKING(worker_ID, case_removed) = ObjWorkExcel.Cells(row, worker_id_col)
		            '     CASES_NO_LONGER_WORKING(case_number, case_removed) = ObjWorkExcel.Cells(row, case_nbr_col)
		            '     CASES_NO_LONGER_WORKING(excel_row, case_removed) = row
		            '     CASES_NO_LONGER_WORKING(client_name, case_removed) = ObjWorkExcel.Cells(row, case_name_col)
		            '     CASES_NO_LONGER_WORKING(application_date, case_removed) = ObjWorkExcel.Cells(row, app_date_col)
		            '     'CASES_NO_LONGER_WORKING(interview_date, case_removed) = ObjWorkExcel.Cells(row, intvw_date_col)
		            '     CASES_NO_LONGER_WORKING(interview_date, case_removed) = ObjWorkExcel.Cells(row, intvw_date_col)
		            '     CASES_NO_LONGER_WORKING(CASH_status, case_removed) = ObjWorkExcel.Cells(row, cash_stat_col)
		            '     CASES_NO_LONGER_WORKING(SNAP_status, case_removed) = ObjWorkExcel.Cells(row, snap_stat_col)
		            ' 	CASES_NO_LONGER_WORKING(appt_notc_sent, case_removed) = ObjWorkExcel.Cells(row, appt_notc_date_col)
		            '     CASES_NO_LONGER_WORKING(appt_notc_confirm, case_removed) = ObjWorkExcel.Cells(row, appt_notc_confirm_col).Value
		            '     CASES_NO_LONGER_WORKING(appointment_date, case_removed) = ObjWorkExcel.Cells(row, appt_date_col)
					' 	CASES_NO_LONGER_WORKING(additional_app_date, case_removed) = ObjWorkExcel.Cells(row, second_app_date_col)
					' 	CASES_NO_LONGER_WORKING(rept_pnd2_listed_days, case_removed) = ObjWorkExcel.Cells(row, rept_pnd2_days_col)
		            '     CASES_NO_LONGER_WORKING(nomi_sent, case_removed) = ObjWorkExcel.Cells(row, nomi_date_col)
		            '     CASES_NO_LONGER_WORKING(nomi_confirm, case_removed) = ObjWorkExcel.Cells(row, nomi_confirm_col)
		            '     CASES_NO_LONGER_WORKING(next_action_needed, case_removed) = ObjWorkExcel.Cells(row, next_action_col)
		            '     CASES_NO_LONGER_WORKING(questionable_intv, case_removed) = ObjWorkExcel.Cells(row, quest_intvw_date_col)
					'
					' 	CASES_NO_LONGER_WORKING(case_in_other_co, case_removed) = ObjWorkExcel.Cells(row, other_county_col)
					' 	CASES_NO_LONGER_WORKING(case_closed_in_30, case_removed) = ObjWorkExcel.Cells(row, closed_in_30_col)
					'
					' 	' CASES_NO_LONGER_WORKING(intvw_quest_resolve, case_removed) = ObjWorkExcel.Cells(row, resolve_quest_intvw_col)
					'
		            '     CASES_NO_LONGER_WORKING(worker_name_one, case_removed) = ObjWorkExcel.Cells(row, worker_name_one_col)
		            '     CASES_NO_LONGER_WORKING(sup_name_one, case_removed) = ObjWorkExcel.Cells(row, sup_name_one_col)
		            '     CASES_NO_LONGER_WORKING(issue_item_one, case_removed) = ObjWorkExcel.Cells(row, issue_item_one_col)
		            '     CASES_NO_LONGER_WORKING(email_ym_one, case_removed) = ObjWorkExcel.Cells(row, email_ym_one_col)
		            '     CASES_NO_LONGER_WORKING(qi_worker_one, case_removed) = ObjWorkExcel.Cells(row, qi_worker_one_col)
					'
		            '     CASES_NO_LONGER_WORKING(worker_name_two, case_removed) = ObjWorkExcel.Cells(row, worker_name_two_col)
		            '     CASES_NO_LONGER_WORKING(sup_name_two, case_removed) = ObjWorkExcel.Cells(row, sup_name_two_col)
		            '     CASES_NO_LONGER_WORKING(issue_item_two, case_removed) = ObjWorkExcel.Cells(row, issue_item_two_col)
		            '     CASES_NO_LONGER_WORKING(email_ym_two, case_removed) = ObjWorkExcel.Cells(row, email_ym_two_col)
		            '     CASES_NO_LONGER_WORKING(qi_worker_two, case_removed) = ObjWorkExcel.Cells(row, qi_worker_two_col)
					'
		            '     CASES_NO_LONGER_WORKING(worker_name_three, case_removed) = ObjWorkExcel.Cells(row, worker_name_three_col)
		            '     CASES_NO_LONGER_WORKING(sup_name_three, case_removed) = ObjWorkExcel.Cells(row, sup_name_three_col)
		            '     CASES_NO_LONGER_WORKING(issue_item_three, case_removed) = ObjWorkExcel.Cells(row, issue_item_three_col)
		            '     CASES_NO_LONGER_WORKING(email_ym_three, case_removed) = ObjWorkExcel.Cells(row, email_ym_three_col)
		            '     CASES_NO_LONGER_WORKING(qi_worker_three, case_removed) = ObjWorkExcel.Cells(row, qi_worker_three_col)
					'
		            '     CASES_NO_LONGER_WORKING(error_notes, case_removed) = "No programs pending."     'This field is used on the removed list to indicate WHY it is no longer on the Working Excel
					'
		            '     'CASES_NO_LONGER_WORKING(error_notes, case_removed) = "Interview Completed on " & TODAYS_CASES_ARRAY(interview_date, case_entry)
		            '     'MsgBox row
		            '     case_removed = case_removed + 1             'adding to the incrementer for the removed cases ARRAY
		            '     'DELETING THE ROW FOR THIS CASE FROM THE WORKING LIST- notice that ROW does not increase as the curent row is now new
		            '     SET objRange = ObjWorkExcel.Cells(row, 1).EntireRow
		            '     objRange.Delete
		            ' Else        'Any case that does not have an interview completed or was previously inidcated as no longer pending is still potentially in need of a notice or denial - and is already listed on the Working Excel
		                ReDim Preserve ALL_PENDING_CASES_ARRAY(error_notes, case_entry)     'resizing the WORKING CASES ARRAY
		                'Now basically the Excel sheet is transcriped row by row to the script ARRAY so we can work with it.
						' actions_detail_var = objWorkRecordSet("NextActionNeeded") 'ObjWorkExcel.Cells(row, script_notes_col).Value objWorkRecordSet("AddedtoWorkList")
						' ' 0 -- Next Action Needed
						' ' 1 -- Worker Notes
						' ' 2 -- Script Notes
						' ' 3 -- Case was in other county
						' ' 4 -- Case closed in past 30 days
						' ' 5 -- PRIV Case
						' ' 6 -- Out of county resolved
						' ' 7 -- closed in 30 days resolved
						' ' 8 -- Subsequent Application resolved
						' array_of_script_notes = split(actions_detail_var, "~-*-~")
						' script_notes_var = trim(array_of_script_notes(2))
						' script_notes_var = replace(script_notes_var, "ADD TO ACTION TODAY EXCEL", "")
						' script_notes_var = replace(script_notes_var, "ADD TO TODAY'S WORKLIST", "")
						' script_notes_var = replace(script_notes_var, "--", "-")

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
						' ALL_PENDING_CASES_ARRAY(error_notes, case_entry)			= trim(array_of_script_notes(1))
						' ALL_PENDING_CASES_ARRAY(script_notes_info, case_entry)		= script_notes_var
						' ALL_PENDING_CASES_ARRAY(case_in_other_co, case_entry)		= trim(array_of_script_notes(3))
						' ALL_PENDING_CASES_ARRAY(case_closed_in_30, case_entry)		= trim(array_of_script_notes(4))
						' ALL_PENDING_CASES_ARRAY(priv_case, case_entry)				= trim(array_of_script_notes(5))
						' ALL_PENDING_CASES_ARRAY(out_of_co_resolve, case_entry)		= trim(array_of_script_notes(6))
						' ALL_PENDING_CASES_ARRAY(closed_in_30_resolve, case_entry)	= trim(array_of_script_notes(7))
						' ALL_PENDING_CASES_ARRAY(subsqt_appl_resolve, case_entry)	= trim(array_of_script_notes(8))

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
		                row = row + 1                   'moving to the next row
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
		yesterday_worker = ObjYestExcel.Cells(2, 2).Value

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
					yesterdays_notes = YESTERDAYS_PENDING_CASES_ARRAY(error_notes, yest_entry)
					If ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) <> "" Then YESTERDAYS_PENDING_CASES_ARRAY(error_notes, yest_entry) = replace(YESTERDAYS_PENDING_CASES_ARRAY(error_notes, yest_entry), ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry),"")
					yesterdays_action_info = YESTERDAYS_PENDING_CASES_ARRAY(yesterday_action_taken, yest_entry)
					If yesterday_worker = qi_member_on_ONDEMAND Then ALL_PENDING_CASES_ARRAY(error_notes, case_entry) = yesterdays_action_info & " - " & yesterdays_notes
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
					IF ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "REVIEW OTHER COUNTY CASE"	and YESTERDAYS_PENDING_CASES_ARRAY(out_of_co_resolve, yest_entry) <> "" Then ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = ""
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
' Else
' 	case_entry = 0      'incrementor to add a case to ALL_PENDING_CASES_ARRAY
'
' 	'This do loops through all of the cases that are already on the working sheet to see if we can find them in today's array
' 	'Reading through each item on the Workking SQP table'
' 	Do While NOT objWorkRecordSet.Eof
' 		ReDim Preserve ALL_PENDING_CASES_ARRAY(error_notes, case_entry)     'resizing the WORKING CASES ARRAY
'
' 		change_date_time = objWorkRecordSet("AuditChangeDate")
' 		change_date_time_array = split(change_date_time, " ")
' 		change_date = change_date_time_array(0)
' 		ALL_PENDING_CASES_ARRAY(line_update_date, case_entry) 		= DateAdd("d", 0, change_date)
'
'
' 		' actions_detail_var = objWorkRecordSet("NextActionNeeded") 'ObjWorkExcel.Cells(row, script_notes_col).Value objWorkRecordSet("AddedtoWorkList")
' 		' ' 0 -- Next Action Needed
' 		' ' 1 -- Worker Notes
' 		' ' 2 -- Script Notes
' 		' ' 3 -- Case was in other county
' 		' ' 4 -- Case closed in past 30 days
' 		' ' 5 -- PRIV Case
' 		' ' 6 -- Out of county resolved
' 		' ' 7 -- closed in 30 days resolved
' 		' ' 8 -- Subsequent Application resolved
' 		' array_of_script_notes = split(actions_detail_var, "~-*-~")
' 		' script_notes_var = trim(array_of_script_notes(2))
' 		' If ALL_PENDING_CASES_ARRAY(line_update_date, case_entry) <> date Then
' 		' 	script_notes_var = replace(script_notes_var, "ADD TO ACTION TODAY EXCEL", "")
' 		' 	script_notes_var = replace(script_notes_var, "ADD TO TODAY'S WORKLIST", "")
' 		' 	script_notes_var = replace(script_notes_var, "--", "-")
' 		' End if
'
' 		' ObjWorkExcel.Cells(row, script_notes_col).Value = script_notes_var
' 		ALL_PENDING_CASES_ARRAY(worker_ID, case_entry) 				= TODAYS_CASES_ARRAY(worker_ID, each_case)
' 		ALL_PENDING_CASES_ARRAY(case_number, case_entry) 			= TODAYS_CASES_ARRAY(case_number, each_case)
' 		' ALL_PENDING_CASES_ARRAY(excel_row, case_entry) = row
' 		ALL_PENDING_CASES_ARRAY(client_name, case_entry) 			= objWorkRecordSet("CaseName") 'ObjWorkExcel.Cells(row, case_name_col)       'This is gathered from the Working Excel instead of the BOBI list because we may have populated a priv case with an actual name
' 		ALL_PENDING_CASES_ARRAY(application_date, case_entry) 		= TODAYS_CASES_ARRAY(application_date, each_case)
' 		ALL_PENDING_CASES_ARRAY(data_day_30, case_entry) 			= objWorkRecordSet("Day_30")
' 		ALL_PENDING_CASES_ARRAY(interview_date, case_entry) 		= objWorkRecordSet("InterviewDate") 		'ObjWorkExcel.Cells(row, intvw_date_col)   'This is gathered from the Working Excel as we may have found an interview date that is NOT in PROG
' 		ALL_PENDING_CASES_ARRAY(CASH_status, case_entry) 			= objWorkRecordSet("CashStatus") 			'ObjWorkExcel.Cells(row, cash_stat_col)
' 		ALL_PENDING_CASES_ARRAY(SNAP_status, case_entry) 			= objWorkRecordSet("SnapStatus") 			'ObjWorkExcel.Cells(row, snap_stat_col)
'
' 		ALL_PENDING_CASES_ARRAY(appt_notc_sent, case_entry) 		= objWorkRecordSet("ApptNoticeDate") 		'ObjWorkExcel.Cells(row, appt_notc_date_col)
' 		ALL_PENDING_CASES_ARRAY(appt_notc_confirm, case_entry) 		= objWorkRecordSet("Confirmation") 			'ObjWorkExcel.Cells(row, appt_notc_confirm_col).Value
' 		ALL_PENDING_CASES_ARRAY(appointment_date, case_entry) 		= objWorkRecordSet("ApptDate") 				'ObjWorkExcel.Cells(row, appt_date_col)
' 		ALL_PENDING_CASES_ARRAY(additional_app_date, case_entry) 	= objWorkRecordSet("SecondApplicationDate") 'ObjWorkExcel.Cells(row, second_app_date_col).Value
' 		ALL_PENDING_CASES_ARRAY(rept_pnd2_listed_days, case_entry) 	= objWorkRecordSet("REPT_PND2Days") 		'ObjWorkExcel.Cells(row, rept_pnd2_days_col).Value
' 		ALL_PENDING_CASES_ARRAY(data_days_pend, case_entry) 		= TODAYS_CASES_ARRAY(data_days_pend, each_case)
' 		ALL_PENDING_CASES_ARRAY(nomi_sent, case_entry) 				= objWorkRecordSet("NOMIDate") 				'ObjWorkExcel.Cells(row, nomi_date_col)
' 		ALL_PENDING_CASES_ARRAY(nomi_confirm, case_entry) 			= objWorkRecordSet("Confirmation2") 		'ObjWorkExcel.Cells(row, nomi_confirm_col)
' 		ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) 	= objWorkRecordSet("NextActionNeeded")			'ObjWorkExcel.Cells(row, next_action_col)
' 		' ALL_PENDING_CASES_ARRAY(error_notes, case_entry)			= trim(array_of_script_notes(1))
' 		' ALL_PENDING_CASES_ARRAY(script_notes_info, case_entry)		= script_notes_var
' 		' ALL_PENDING_CASES_ARRAY(case_in_other_co, case_entry)		= trim(array_of_script_notes(3))
' 		' ALL_PENDING_CASES_ARRAY(case_closed_in_30, case_entry)		= trim(array_of_script_notes(4))
' 		' ALL_PENDING_CASES_ARRAY(priv_case, case_entry)				= trim(array_of_script_notes(5))
' 		' ALL_PENDING_CASES_ARRAY(out_of_co_resolve, case_entry)		= trim(array_of_script_notes(6))
' 		' ALL_PENDING_CASES_ARRAY(closed_in_30_resolve, case_entry)	= trim(array_of_script_notes(7))
' 		' ALL_PENDING_CASES_ARRAY(subsqt_appl_resolve, case_entry)	= trim(array_of_script_notes(8))
'
' 		ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry) 		= objWorkRecordSet("QuestionableInterview") 'ObjWorkExcel.Cells(row, quest_intvw_date_col)
' 		ALL_PENDING_CASES_ARRAY(intvw_quest_resolve, case_entry)	= objWorkRecordSet("Resolved")
'
' 		ALL_PENDING_CASES_ARRAY(add_to_daily_worklist, case_entry) = False
' 		ALL_PENDING_CASES_ARRAY(last_wl_date, case_entry) 			= objWorkRecordSet("AddedtoWorkList")
' 		If IsDate(ALL_PENDING_CASES_ARRAY(last_wl_date, case_entry)) = True Then
' 			ALL_PENDING_CASES_ARRAY(last_wl_date, case_entry) = DateAdd("d", 0, ALL_PENDING_CASES_ARRAY(last_wl_date, case_entry))
' 			If ALL_PENDING_CASES_ARRAY(last_wl_date, case_entry) = date Then ALL_PENDING_CASES_ARRAY(add_to_daily_worklist, case_entry) = True
' 		End If
' 		change_date_time = objWorkRecordSet("AuditChangeDate")
' 		change_date_time_array = split(change_date_time, " ")
' 		change_date = change_date_time_array(0)
' 		ALL_PENDING_CASES_ARRAY(line_update_date, case_entry) 		= DateAdd("d", 0, change_date)		' "DenialNeeded"
' 		' ALL_PENDING_CASES_ARRAY(error_notes, case_entry) 			= objWorkRecordSet("AddedtoWorkList") 'ObjWorkExcel.Cells(row, worker_notes_col)
' 		' ALL_PENDING_CASES_ARRAY(line_update_date, case_entry) 		= objWorkRecordSet("AddedtoWorkList") 'ObjWorkExcel.Cells(row, script_revw_date_col)
' 		' ALL_PENDING_CASES_ARRAY(line_update_date, case_entry) = dateAdd("d", 0, ALL_PENDING_CASES_ARRAY(line_update_date, case_entry))
'
' 		'Defaulting this values at this time as we will determine them to be different as the script proceeds.
' 		ALL_PENDING_CASES_ARRAY(take_action_today, case_entry) = FALSE
'
' 		ALL_PENDING_CASES_ARRAY(script_action_taken, case_entry) = False
' 		If DateDiff("d", ALL_PENDING_CASES_ARRAY(appt_notc_sent, case_entry), date) = 0 AND ALL_PENDING_CASES_ARRAY(appt_notc_confirm, case_entry) = "Y" Then ALL_PENDING_CASES_ARRAY(script_action_taken, case_entry) = True
' 		If DateDiff("d", ALL_PENDING_CASES_ARRAY(nomi_sent, case_entry), date) = 0 AND ALL_PENDING_CASES_ARRAY(nomi_confirm, case_entry) = "Y" Then ALL_PENDING_CASES_ARRAY(script_action_taken, case_entry) = True
' 		case_entry = case_entry + 1     'increasing the count for '
' 		objWorkRecordSet.MoveNext
' 	Loop
' 	objWorkRecordSet.Close
' 	' objWorkConnection.Close
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

		objWorkRecordSet.Open "INSERT INTO ES.ES_OnDemanCashAndSnapBZProcessed (CaseNumber, CaseName, ApplDate, InterviewDate, Day_30, DaysPending, SnapStatus, CashStatus, SecondApplicationDate, REPT_PND2Days, QuestionableInterview, Resolved, ApptNoticeDate, ApptDate, Confirmation, NOMIDate, Confirmation2, DenialNeeded, NextActionNeeded, AddedtoWorkList)" & _
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
										ALL_PENDING_CASES_ARRAY(last_wl_date, case_entry) & "')", objWorkConnection, adOpenStatic, adLockOptimistic
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
ObjExcel.ActiveWorkbook.SaveAs file_selection_path
ObjExcel.Quit

objStatsWorkbook.Save
ObjStatsExcel.Quit





Call script_end_procedure("Working List is Updated")













'END'
