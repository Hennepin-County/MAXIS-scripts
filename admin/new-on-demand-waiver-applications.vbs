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
const deny_memo_confirm     = 18
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

const error_notes 			= 52

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

const wl_rept_pnd2_days_col		= 6		'worklist'
const wl_app_date_col 			= 7		'worklist'
const wl_second_app_date_col	= 8		'worklist'
const wl_intvw_date_col        	= 9		'worklist'
const wl_quest_intvw_date_col  	= 10	'worklist'
' const wl_resolve_quest_intvw_col	= 11	'worklist'
const wl_other_county_col		= 11	'worklist'
const wl_closed_in_30_col		= 12	'worklist'

const wl_appt_notc_date_col   	= 13	'worklist'
const wl_appt_date_col         	= 14	'worklist'
const wl_nomi_date_col         	= 15	'worklist'
const wl_day_30_col 			= 16	'worklist'
const wl_deny_col 				= 17	'worklist'
const wl_ecf_doc_accepted_col	= 18	'worklist'
const wl_action_taken_col 		= 19	'worklist'
const wl_work_notes_col			= 20	'worklist'
const wl_email_worker_col		= 21	'worklist'
const wl_email_issue_col		= 22	'worklist'

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
objWorkSQL = "SELECT * FROM ES.ES_OnDemandCashAndSnapBZProcessed"

'Creating objects for Access
Set objWorkConnection = CreateObject("ADODB.Connection")
Set objWorkRecordSet = CreateObject("ADODB.Recordset")

'This is the file path for the statistics Access database.
objWorkConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
objWorkRecordSet.Open objWorkSQL, objWorkConnection

'pulling the date changed from the first record in the working list.
'This it to identify if this is a restart or not.
first_item_date = objWorkRecordSet("AuditChangeDate")
first_item_date = DateAdd("d", 0, first_item_date)

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

	'Reading through each item on the Workking SQP table'
	Do While NOT objWorkRecordSet.Eof
		case_number_to_assess = objWorkRecordSet("CaseNumber")  			'getting the case number in the Working Excel sheet
		found_case_on_todays_list = FALSE                               	'this Boolean is used to determine if the case number is on the BOBI run today

		For each_case = 0 to UBound(TODAYS_CASES_ARRAY, 2)              'This loops through each case that was on the BOBI today
	        'MsgBox "Excel case number: " & case_number_to_assess & vbNewLine & "Array case number: " & TODAYS_CASES_ARRAY(case_number, each_case)
	        If case_number_to_assess = TODAYS_CASES_ARRAY(case_number, each_case) Then  'If a matching case number is found this means the case was on the working excel AND is on the BOBI
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
					script_notes_var = objWorkRecordSet("AddedtoWorkList") 'ObjWorkExcel.Cells(row, script_notes_col).Value objWorkRecordSet("AddedtoWorkList")
					script_notes_var = replace(script_notes_var, "ADD TO ACTION TODAY EXCEL", "")
					script_notes_var = replace(script_notes_var, "ADD TO TODAY'S WORKLIST", "")
					script_notes_var = replace(script_notes_var, "--", "-")
					array_of_script_notes = split(script_notes_var, "~-*-~")
					script_notes_var = trim(array_of_script_notes(1))
					array_of_script_notes = ""

					' ObjWorkExcel.Cells(row, script_notes_col).Value = script_notes_var
	                ALL_PENDING_CASES_ARRAY(worker_ID, case_entry) 				= TODAYS_CASES_ARRAY(worker_ID, each_case)
	                ALL_PENDING_CASES_ARRAY(case_number, case_entry) 			= TODAYS_CASES_ARRAY(case_number, each_case)
	                ' ALL_PENDING_CASES_ARRAY(excel_row, case_entry) = row
	                ALL_PENDING_CASES_ARRAY(client_name, case_entry) 			= objWorkRecordSet("CaseName") 'ObjWorkExcel.Cells(row, case_name_col)       'This is gathered from the Working Excel instead of the BOBI list because we may have populated a priv case with an actual name
	                ALL_PENDING_CASES_ARRAY(application_date, case_entry) 		= TODAYS_CASES_ARRAY(application_date, each_case)
					ALL_PENDING_CASES_ARRAY(data_day_30, case_entry) 			= TODAYS_CASES_ARRAY(data_day_30, each_case)
	                ALL_PENDING_CASES_ARRAY(interview_date, case_entry) 		= objWorkRecordSet("InterviewDate") 		'ObjWorkExcel.Cells(row, intvw_date_col)   'This is gathered from the Working Excel as we may have found an interview date that is NOT in PROG
	                ALL_PENDING_CASES_ARRAY(CASH_status, case_entry) 			= objWorkRecordSet("CashStatus") 			'ObjWorkExcel.Cells(row, cash_stat_col)
	                ALL_PENDING_CASES_ARRAY(SNAP_status, case_entry) 			= objWorkRecordSet("SnapStatus") 			'ObjWorkExcel.Cells(row, snap_stat_col)

	                ALL_PENDING_CASES_ARRAY(appt_notc_sent, case_entry) 		= objWorkRecordSet("ApptNoticeDate") 		'ObjWorkExcel.Cells(row, appt_notc_date_col)
	                ALL_PENDING_CASES_ARRAY(appt_notc_confirm, case_entry) 		= objWorkRecordSet("Confirmation") 			'ObjWorkExcel.Cells(row, appt_notc_confirm_col).Value
	                ALL_PENDING_CASES_ARRAY(appointment_date, case_entry) 		= objWorkRecordSet("ApptDate") 				'ObjWorkExcel.Cells(row, appt_date_col)
					ALL_PENDING_CASES_ARRAY(additional_app_date, case_entry) 	= objWorkRecordSet("SecondApplicationDate") 'ObjWorkExcel.Cells(row, second_app_date_col).Value
					ALL_PENDING_CASES_ARRAY(rept_pnd2_listed_days, case_entry) 	= objWorkRecordSet("REPT_PND2Days") 		'ObjWorkExcel.Cells(row, rept_pnd2_days_col).Value
	                ALL_PENDING_CASES_ARRAY(nomi_sent, case_entry) 				= objWorkRecordSet("NOMIDate") 				'ObjWorkExcel.Cells(row, nomi_date_col)
	                ALL_PENDING_CASES_ARRAY(nomi_confirm, case_entry) 			= objWorkRecordSet("Confirmation2") 		'ObjWorkExcel.Cells(row, nomi_confirm_col)
	                ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) 	= objWorkRecordSet("NextActionNeeded") 		'ObjWorkExcel.Cells(row, next_action_col)
	                ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry) 		= objWorkRecordSet("QuestionableInterview") 'ObjWorkExcel.Cells(row, quest_intvw_date_col)
					'TODO - MISSING - ALL_PENDING_CASES_ARRAY(case_in_other_co, case_entry)		= objWorkRecordSet("")		'Call read_boolean_from_excel(ObjWorkExcel.Cells(row, other_county_col).Value, ALL_PENDING_CASES_ARRAY(case_in_other_co, case_entry))
					'TODO - MISSING - ALL_PENDING_CASES_ARRAY(case_closed_in_30, case_entry)		= objWorkRecordSet("")		'Call read_boolean_from_excel(ObjWorkExcel.Cells(row, closed_in_30_col).Value, ALL_PENDING_CASES_ARRAY(case_closed_in_30, case_entry))
	                ' ALL_PENDING_CASES_ARRAY(error_notes, case_entry) 			= objWorkRecordSet("AddedtoWorkList") 'ObjWorkExcel.Cells(row, worker_notes_col)
					' ALL_PENDING_CASES_ARRAY(line_update_date, case_entry) 		= objWorkRecordSet("AddedtoWorkList") 'ObjWorkExcel.Cells(row, script_revw_date_col)
					' ALL_PENDING_CASES_ARRAY(line_update_date, case_entry) = dateAdd("d", 0, ALL_PENDING_CASES_ARRAY(line_update_date, case_entry))
	                'ALL_PENDING_CASES_ARRAY(, case_entry) = ObjWorkExcel.Cells(row, )

	                'Defaulting this values at this time as we will determine them to be different as the script proceeds.
	                ALL_PENDING_CASES_ARRAY(take_action_today, case_entry) = FALSE
					ALL_PENDING_CASES_ARRAY(add_to_daily_worklist, case_entry) = False
	                case_entry = case_entry + 1     'increasing the count for '
	                row = row + 1                   'moving to the next row
	            End If
	            Exit For                            'This is to leave the loop of looking through all of the cases in the BOBI list ARRAY because we found the match - and there should never be duplicates
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

		objWorkRecordSet.MoveNext
	Loop

	'Actually deleting the row in the Working Excel - notice that ROW does not increase as the curent row is now new
	For case_removed = 0 to UBound(CASES_NO_LONGER_WORKING, 2)
		case_number_to_review = CASES_NO_LONGER_WORKING(case_number, case_removed)
		objWorkRecordSet.Open "DELETE FROM ES.ES_OnDemandCashAndSnap WHERE CaseNumber = '" & case_number_to_review & "'", objWorkConnection
	Next



	'BE SURE TO ALWAYS LEAVE THE row VARIABLE ALONE HERE AS WE USE IT IN THIS FOR NEXT TO ADD TO THE END OF THE WORKING EXCEL
	add_a_case = case_entry     'creating an incrementer that starts where the last one ended for the ALL PENDING CASES ARRAY
	For case_entry = 0 to UBOUND(TODAYS_CASES_ARRAY, 2)     'now we are going to look at each of the cases in the ARRAY for today's BOBI list
	    'MsgBox TODAYS_CASES_ARRAY(on_working_list, case_entry)
	    'MsgBox TODAYS_CASES_ARRAY(interview_date, case_entry)
	    If TODAYS_CASES_ARRAY(on_working_list, case_entry) = FALSE AND TODAYS_CASES_ARRAY(interview_date, case_entry) = "" Then
	        'These are all the cases on todays list that were NOT on the Working Excel AND have not already had an interview
	        'adding the information known from the BOBI to the Working Excel
	        ObjWorkExcel.Cells(row, worker_id_col) = TODAYS_CASES_ARRAY(worker_ID, case_entry)
	        ObjWorkExcel.Cells(row, case_nbr_col) = TODAYS_CASES_ARRAY(case_number, case_entry)
	        ObjWorkExcel.Cells(row, case_name_col) = TODAYS_CASES_ARRAY(client_name, case_entry)
	        ObjWorkExcel.Cells(row, app_date_col) = TODAYS_CASES_ARRAY(application_date, case_entry)
	        ObjWorkExcel.Cells(row, intvw_date_col) = TODAYS_CASES_ARRAY(interview_date, case_entry)
			ObjWorkExcel.Cells(row, rept_pnd2_days_col) = TODAYS_CASES_ARRAY(data_days_pend, case_entry)
			ObjWorkExcel.Cells(row, day_30_col) = TODAYS_CASES_ARRAY(data_day_30, case_entry)
	        'ObjWorkExcel.Cells(row, ) = TODAYS_CASES_ARRAY(, case_entry)

	        ReDim Preserve ALL_PENDING_CASES_ARRAY(error_notes, add_a_case)         'resizing the array of the Working Excel
	        'Now all the information needs to be added to the ARRAY from the Working Excel
			script_notes_var = ObjWorkExcel.Cells(row, script_notes_col).Value
			script_notes_var = replace(script_notes_var, "ADD TO ACTION TODAY EXCEL", "")
			script_notes_var = replace(script_notes_var, "ADD TO TODAY'S WORKLIST", "")
			script_notes_var = replace(script_notes_var, "--", "-")
			script_notes_var = trim(script_notes_var)
			ObjWorkExcel.Cells(row, script_notes_col).Value = script_notes_var

	        ALL_PENDING_CASES_ARRAY(worker_ID, add_a_case) = TODAYS_CASES_ARRAY(worker_ID, case_entry)
	        ALL_PENDING_CASES_ARRAY(case_number, add_a_case) = TODAYS_CASES_ARRAY(case_number, case_entry)
	        ALL_PENDING_CASES_ARRAY(excel_row, add_a_case) = row
	        ALL_PENDING_CASES_ARRAY(client_name, add_a_case) = TODAYS_CASES_ARRAY(client_name, case_entry)
	        ALL_PENDING_CASES_ARRAY(application_date, add_a_case) 		= ObjWorkExcel.Cells(row, app_date_col)
	        ALL_PENDING_CASES_ARRAY(interview_date, add_a_case) 		= ObjWorkExcel.Cells(row, intvw_date_col)
	        ALL_PENDING_CASES_ARRAY(CASH_status, add_a_case) 			= ObjWorkExcel.Cells(row, cash_stat_col)
	        ALL_PENDING_CASES_ARRAY(SNAP_status, add_a_case) 			= ObjWorkExcel.Cells(row, snap_stat_col)

	        ALL_PENDING_CASES_ARRAY(appt_notc_sent, add_a_case) 		= ObjWorkExcel.Cells(row, appt_notc_date_col)
	        ALL_PENDING_CASES_ARRAY(appt_notc_confirm, add_a_case) 		= ObjWorkExcel.Cells(row, appt_notc_confirm_col).Value
	        ALL_PENDING_CASES_ARRAY(appointment_date, add_a_case) 		= ObjWorkExcel.Cells(row, appt_date_col)
			ALL_PENDING_CASES_ARRAY(additional_app_date, add_a_case) 	= ObjWorkExcel.Cells(row, second_app_date_col).Value
			ALL_PENDING_CASES_ARRAY(rept_pnd2_listed_days, add_a_case) 	= ObjWorkExcel.Cells(row, rept_pnd2_days_col).Value
	        ALL_PENDING_CASES_ARRAY(nomi_sent, add_a_case) 				= ObjWorkExcel.Cells(row, nomi_date_col)
	        ALL_PENDING_CASES_ARRAY(nomi_confirm, add_a_case) 			= ObjWorkExcel.Cells(row, nomi_confirm_col)
	        ALL_PENDING_CASES_ARRAY(next_action_needed, add_a_case) 	= ObjWorkExcel.Cells(row, next_action_col)
	        ALL_PENDING_CASES_ARRAY(questionable_intv, add_a_case) 		= ObjWorkExcel.Cells(row, quest_intvw_date_col)
	        ALL_PENDING_CASES_ARRAY(questionable_intv, add_a_case) 		= trim(ALL_PENDING_CASES_ARRAY(questionable_intv, add_a_case))
			ObjWorkExcel.Cells(row, other_county_col).Value = trim(ObjWorkExcel.Cells(row, other_county_col).Value)
			ObjWorkExcel.Cells(row, closed_in_30_col).Value = trim(ObjWorkExcel.Cells(row, closed_in_30_col).Value)
			Call read_boolean_from_excel(ObjWorkExcel.Cells(row, other_county_col).Value, ALL_PENDING_CASES_ARRAY(case_in_other_co, add_a_case))
			Call read_boolean_from_excel(ObjWorkExcel.Cells(row, closed_in_30_col).Value, ALL_PENDING_CASES_ARRAY(case_closed_in_30, add_a_case))
			ALL_PENDING_CASES_ARRAY(error_notes, add_a_case) 			= ObjWorkExcel.Cells(row, worker_notes_col)
			ALL_PENDING_CASES_ARRAY(line_update_date, add_a_case) 		= ObjWorkExcel.Cells(row, script_revw_date_col)
			ALL_PENDING_CASES_ARRAY(line_update_date, add_a_case) = dateAdd("d", 0, ALL_PENDING_CASES_ARRAY(line_update_date, add_a_case))
			'ALL_PENDING_CASES_ARRAY(, add_a_case) = ObjWorkExcel.Cells(row, )
	        'defaulting this variable as we will determine if it is true later
	        ALL_PENDING_CASES_ARRAY(take_action_today, add_a_case) = FALSE
			ALL_PENDING_CASES_ARRAY(add_to_daily_worklist, add_a_case) = False
	        add_a_case = add_a_case + 1     'incrementing the counter for this ARRAY
	        row = row + 1                   'going to the next row so that we don't overwrite the information we just added
	    End If
	Next








End If




'END'
