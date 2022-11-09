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


'Checking the working list to see when last updated
'declare the SQL statement that will query the database
objWorkSQL = "SELECT * FROM ES.ES_OnDemanCashAndSnapBZProcessed"

'Creating objects for Access
Set objWorkConnection = CreateObject("ADODB.Connection")
Set objWorkRecordSet = CreateObject("ADODB.Recordset")

'This is the file path for the statistics Access database.
objWorkConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
objWorkRecordSet.Open objWorkSQL, objWorkConnection



case_entry = 0      'incrementor to add a case to ALL_PENDING_CASES_ARRAY
list_of_all_cases = ""

'Reading through each item on the Workking SQP table'
Do While NOT objWorkRecordSet.Eof
	case_number_to_assess = objWorkRecordSet("CaseNumber")  			'getting the case number in the Working Excel sheet
	' case_name_to_assess = objWorkRecordSet("CaseName")
	' found_case_on_todays_list = FALSE                               	'this Boolean is used to determine if the case number is on the BOBI run today
	If InStr(list_of_all_cases, "*" & case_number_to_assess & "*") = 0 Then 		'making sure we don't have repeat case numbers
		list_of_all_cases = list_of_all_cases & case_number_to_assess & "*"
  		ReDim Preserve ALL_PENDING_CASES_ARRAY(error_notes, case_entry)     'resizing the WORKING CASES ARRAY

        ' ALL_PENDING_CASES_ARRAY(worker_ID, case_entry) 				= TODAYS_CASES_ARRAY(worker_ID, each_case)
        ALL_PENDING_CASES_ARRAY(case_number, case_entry) 			= objWorkRecordSet("CaseNumber")
        ' ALL_PENDING_CASES_ARRAY(excel_row, case_entry) = row
        ALL_PENDING_CASES_ARRAY(client_name, case_entry) 			= objWorkRecordSet("CaseName") 'ObjWorkExcel.Cells(row, case_name_col)       'This is gathered from the Working Excel instead of the BOBI list because we may have populated a priv case with an actual name
        ALL_PENDING_CASES_ARRAY(application_date, case_entry) 		= objWorkRecordSet("ApplDate")	'TODAYS_CASES_ARRAY(application_date, each_case)
		ALL_PENDING_CASES_ARRAY(data_day_30, case_entry) 			= objWorkRecordSet("Day_30")
        ALL_PENDING_CASES_ARRAY(interview_date, case_entry) 		= objWorkRecordSet("InterviewDate") 		'ObjWorkExcel.Cells(row, intvw_date_col)   'This is gathered from the Working Excel as we may have found an interview date that is NOT in PROG
        ALL_PENDING_CASES_ARRAY(CASH_status, case_entry) 			= objWorkRecordSet("CashStatus") 			'ObjWorkExcel.Cells(row, cash_stat_col)
        ALL_PENDING_CASES_ARRAY(SNAP_status, case_entry) 			= objWorkRecordSet("SnapStatus") 			'ObjWorkExcel.Cells(row, snap_stat_col)

        ALL_PENDING_CASES_ARRAY(appt_notc_sent, case_entry) 		= objWorkRecordSet("ApptNoticeDate") 		'ObjWorkExcel.Cells(row, appt_notc_date_col)
        ALL_PENDING_CASES_ARRAY(appt_notc_confirm, case_entry) 		= objWorkRecordSet("Confirmation") 			'ObjWorkExcel.Cells(row, appt_notc_confirm_col).Value
        ALL_PENDING_CASES_ARRAY(appointment_date, case_entry) 		= objWorkRecordSet("ApptDate") 				'ObjWorkExcel.Cells(row, appt_date_col)
		ALL_PENDING_CASES_ARRAY(additional_app_date, case_entry) 	= objWorkRecordSet("SecondApplicationDate") 'ObjWorkExcel.Cells(row, second_app_date_col).Value
		ALL_PENDING_CASES_ARRAY(rept_pnd2_listed_days, case_entry) 	= objWorkRecordSet("REPT_PND2Days") 		'ObjWorkExcel.Cells(row, rept_pnd2_days_col).Value
		ALL_PENDING_CASES_ARRAY(data_days_pend, case_entry) 		= objWorkRecordSet("DaysPending") 		'TODAYS_CASES_ARRAY(data_days_pend, each_case)
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
	End If

	objWorkRecordSet.MoveNext
Loop

objWorkRecordSet.Close
objWorkConnection.Close

Set objWorkRecordSet=nothing
Set objWorkConnection=nothing
Set objWorkSQL=nothing


For case_entry = 0 to UBOUND(ALL_PENDING_CASES_ARRAY, 2)
	' If ALL_PENDING_CASES_ARRAY(line_update_date, case_entry) <> date and ALL_PENDING_CASES_ARRAY(priv_case, case_entry) = False Then
	' If ALL_PENDING_CASES_ARRAY(priv_case, case_entry) = False and ALL_PENDING_CASES_ARRAY(deleted_today, case_entry) = False Then

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


	If ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "Send Manual NOMI" Then

		MAXIS_case_number	= ALL_PENDING_CASES_ARRAY(case_number, case_entry)        'setting this so that nav functionality wor
		CALL back_to_SELF

		forms_to_swkr = ""      'setting these for the start a memo function
		forms_to_arep = ""
		memo_started = TRUE


		If IsDate(ALL_PENDING_CASES_ARRAY(application_date, case_entry)) = False Then
			ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "Send Manual NOMI"
		Else
			nomi_last_contact_day = dateadd("d", 30, ALL_PENDING_CASES_ARRAY(application_date, case_entry))       'setting the date to enter on the NOMI of the day of denial
			'ensuring that we have given the client an additional10days fromt he day nomi sent'
			IF DateDiff("d", date, nomi_last_contact_day) < 1 then nomi_last_contact_day = dateadd("d", 10, date)

			Call start_a_new_spec_memo(memo_started, True, forms_to_arep, forms_to_swkr, "N", other_name, other_street, other_city, other_state, other_zip, False)
			IF memo_started = True THEN
				'TODO - add languages in when we can'

				Call write_variable_in_SPEC_MEMO("You recently applied for assistance on " & ALL_PENDING_CASES_ARRAY(application_date, case_entry) & ".")
				Call write_variable_in_SPEC_MEMO("Your interview should have been completed by " & ALL_PENDING_CASES_ARRAY(appointment_date, case_entry) & ".")
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
				ALL_PENDING_CASES_ARRAY(nomi_confirm, case_entry) = "N"   'if the MEMO didn't start then setting this for the ARRAY and Working Excel.
			End If

			If ALL_PENDING_CASES_ARRAY(nomi_confirm, case_entry) <> "N" Then Call confirm_memo_waiting(ALL_PENDING_CASES_ARRAY(nomi_confirm, case_entry))     'reading the SPEC/MEMO page to see that a MEMO for today is waiting.

			'Resetting the next action needed based on message success and writing the case note if successful
			If ALL_PENDING_CASES_ARRAY(nomi_confirm, case_entry) = "N" Then
				ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "Send Manual NOMI"
				ALL_PENDING_CASES_ARRAY(nomi_sent, case_entry) = ""
			ElseIf ALL_PENDING_CASES_ARRAY(nomi_confirm, case_entry) = "Y" Then
				ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "DENY AT DAY 30"
				ALL_PENDING_CASES_ARRAY(nomi_sent, case_entry) = date

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
				ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry) = "???"           'this is for testing - this has never come up
			End If
			Call back_to_SELF

			'Adding this case to the list of cases that we took action on today
			ReDim Preserve ACTION_TODAY_CASES_ARRAY(error_notes, todays_cases)
			ACTION_TODAY_CASES_ARRAY(case_number, todays_cases)         = ALL_PENDING_CASES_ARRAY(case_number, case_entry)
			ACTION_TODAY_CASES_ARRAY(client_name, todays_cases)         = ALL_PENDING_CASES_ARRAY(client_name, case_entry)
			ACTION_TODAY_CASES_ARRAY(worker_ID, todays_cases)           = ALL_PENDING_CASES_ARRAY(worker_ID, case_entry)
			ACTION_TODAY_CASES_ARRAY(SNAP_status, todays_cases)         = ALL_PENDING_CASES_ARRAY(SNAP_status, case_entry)
			ACTION_TODAY_CASES_ARRAY(CASH_status, todays_cases)         = ALL_PENDING_CASES_ARRAY(CASH_status, case_entry)
			ACTION_TODAY_CASES_ARRAY(application_date, todays_cases)    = ALL_PENDING_CASES_ARRAY(application_date, case_entry)
			ACTION_TODAY_CASES_ARRAY(interview_date, todays_cases)      = ALL_PENDING_CASES_ARRAY(interview_date, case_entry)
			ACTION_TODAY_CASES_ARRAY(questionable_intv, todays_cases)   = ALL_PENDING_CASES_ARRAY(questionable_intv, case_entry)
			ACTION_TODAY_CASES_ARRAY(appt_notc_sent, todays_cases)      = ALL_PENDING_CASES_ARRAY(appt_notc_sent, case_entry)
			ACTION_TODAY_CASES_ARRAY(appt_notc_confirm, todays_cases)   = ALL_PENDING_CASES_ARRAY(appt_notc_confirm, case_entry)
			ACTION_TODAY_CASES_ARRAY(appointment_date, todays_cases)    = ALL_PENDING_CASES_ARRAY(appointment_date, case_entry)
			ACTION_TODAY_CASES_ARRAY(nomi_sent, todays_cases)           = ALL_PENDING_CASES_ARRAY(nomi_sent, case_entry)
			ACTION_TODAY_CASES_ARRAY(nomi_confirm, todays_cases)        = ALL_PENDING_CASES_ARRAY(nomi_confirm, case_entry)
			ACTION_TODAY_CASES_ARRAY(deny_day30, todays_cases)          = ALL_PENDING_CASES_ARRAY(deny_day30, case_entry)
			ACTION_TODAY_CASES_ARRAY(deny_memo_confirm, todays_cases)   = ALL_PENDING_CASES_ARRAY(deny_memo_confirm, case_entry)
			ACTION_TODAY_CASES_ARRAY(next_action_needed, todays_cases)  = ALL_PENDING_CASES_ARRAY(next_action_needed, case_entry)
			ACTION_TODAY_CASES_ARRAY(error_notes, todays_cases)         = "NOMI Sent today"
			todays_cases = todays_cases + 1

			ALL_PENDING_CASES_ARRAY(script_action_taken, case_entry) = True
		End If


	End If

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

'Now the script reopens the daily list that was identified in the beginning
file_date = replace(current_date, "/", "-")   'Changing the format of the date to use as file path selection default
daily_case_list_folder = right("0" & DatePart("m", file_date), 2) & "-" & DatePart("yyyy", file_date)
file_selection_path = t_drive & "/Eligibility Support/Restricted/QI - Quality Improvement/REPORTS/On Demand Waiver/Daily case lists/" & daily_case_list_folder & "/" & file_date & ".xlsx" 'single assignment file

call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)

'On the same Excel file the script creates a new sheet and names it
objWorkbook.worksheets("Actions Today").Activate


action_row = 1      'setting the first row
Do
	action_row = action_row + 1     'go to the next row
	listed_case_number = trim(ObjExcel.Cells(action_row, case_nbr_col))
Loop until listed_case_number = ""
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
