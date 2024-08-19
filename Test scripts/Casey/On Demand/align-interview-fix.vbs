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
const align_intv_revw_dt	= 58

const error_notes 			= 59

Dim WORKING_LIST_CASES_ARRAY()
ReDim WORKING_LIST_CASES_ARRAY(error_notes, 0)

MsgBox "Working"


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
	If InStr(list_of_all_cases, "*" & case_number_to_assess & "*") = 0 and case_number_to_assess <> "2536665" Then 		'making sure we don't have repeat case numbers

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

		If WORKING_LIST_CASES_ARRAY(next_action_needed, case_entry) = "PREP FOR DENIAL" Then WORKING_LIST_CASES_ARRAY(next_action_needed, case_entry) = ""
		If WORKING_LIST_CASES_ARRAY(next_action_needed, case_entry) = "PENDING MORE THAN 30 DAYS" Then WORKING_LIST_CASES_ARRAY(next_action_needed, case_entry) = ""

		case_review_notes = "FOLLOW UP NEEDED - " & case_review_notes
		' MsgBox "script_notes_info" & vbCr & WORKING_LIST_CASES_ARRAY(script_notes_info, case_entry)
		edit_notes = false
		If WORKING_LIST_CASES_ARRAY(script_notes_info, case_entry) <> NULL Then edit_notes = True
		If WORKING_LIST_CASES_ARRAY(script_notes_info, case_entry) <> "" Then edit_notes = True
		If edit_notes = True Then
			WORKING_LIST_CASES_ARRAY(script_notes_info, case_entry) = replace(WORKING_LIST_CASES_ARRAY(script_notes_info, case_entry), "STS-NR", "")
			WORKING_LIST_CASES_ARRAY(script_notes_info, case_entry) = trim(WORKING_LIST_CASES_ARRAY(script_notes_info, case_entry))

			If InStr(WORKING_LIST_CASES_ARRAY(script_notes_info, case_entry), "%$#@") Then
				beg_of_intv_revw = InStr(WORKING_LIST_CASES_ARRAY(script_notes_info, case_entry), "@#$%")
				beg_of_intv_revw = beg_of_intv_revw+17
				end_of_intv_revw = InStr(WORKING_LIST_CASES_ARRAY(script_notes_info, case_entry), "@%$#@")
				If beg_of_intv_revw = end_of_intv_revw Then
					WORKING_LIST_CASES_ARRAY(align_intv_revw_dt, case_entry) = False
					WORKING_LIST_CASES_ARRAY(script_notes_info, case_entry) = replace(WORKING_LIST_CASES_ARRAY(script_notes_info, case_entry), "@#$%REVIEWED ON: @%$#@", "")
				Else
					WORKING_LIST_CASES_ARRAY(align_intv_revw_dt, case_entry) = Mid(WORKING_LIST_CASES_ARRAY(script_notes_info, case_entry), beg_of_intv_revw, end_of_intv_revw-beg_of_intv_revw)
					WORKING_LIST_CASES_ARRAY(script_notes_info, case_entry) = replace(WORKING_LIST_CASES_ARRAY(script_notes_info, case_entry), "@#$%REVIEWED ON: " & WORKING_LIST_CASES_ARRAY(align_intv_revw_dt, case_entry) & "@%$#@", "")
				End If
				WORKING_LIST_CASES_ARRAY(align_intv_revw_dt, case_entry) = trim(WORKING_LIST_CASES_ARRAY(align_intv_revw_dt, case_entry))
			End If
			' align_intv_revw_dt
			' "@#$%REVIEWED ON: @%$#@"

			' MsgBox "Align Interview Date - " & WORKING_LIST_CASES_ARRAY(align_intv_revw_dt, case_entry) & vbCr & vbCr & "SQL Notes:" & vbCr & objWorkRecordSet("TrackingNotes") & vbCr & "Array Notes:" & WORKING_LIST_CASES_ARRAY(script_notes_info, case_entry)

		End If
		' "DenialNeeded"
		' WORKING_LIST_CASES_ARRAY(error_notes, case_entry) 			= objWorkRecordSet("AddedtoWorkList") 'ObjWorkExcel.Cells(row, worker_notes_col)
		' WORKING_LIST_CASES_ARRAY(line_update_date, case_entry) 		= objWorkRecordSet("AddedtoWorkList") 'ObjWorkExcel.Cells(row, script_revw_date_col)
		' WORKING_LIST_CASES_ARRAY(line_update_date, case_entry) = dateAdd("d", 0, WORKING_LIST_CASES_ARRAY(line_update_date, case_entry))

		'Defaulting this values at this time as we will determine them to be different as the script proceeds.
		WORKING_LIST_CASES_ARRAY(add_to_daily_worklist, case_entry) = False

		' For case_info = 0 to UBOUND(ALL_PENDING_CASES_ARRAY, 2)
		' 	If ALL_PENDING_CASES_ARRAY(case_number, case_info) = WORKING_LIST_CASES_ARRAY(case_number, case_entry) Then
		' 		WORKING_LIST_CASES_ARRAY(error_notes, case_entry) = ALL_PENDING_CASES_ARRAY(error_notes, case_info)
		' 	End If
		' Next

		case_entry = case_entry + 1     'increasing the count for '
	End If

	objWorkRecordSet.MoveNext
Loop

objWorkRecordSet.Close
objWorkConnection.Close

Set objWorkRecordSet=nothing
Set objWorkConnection=nothing
Set objWorkSQL=nothing

count_cases_on_wl = 0
For case_entry = 0 to UBOUND(WORKING_LIST_CASES_ARRAY, 2)
	MAXIS_case_number = WORKING_LIST_CASES_ARRAY(case_number, case_entry)

	update_sql = False

	IF WORKING_LIST_CASES_ARRAY(next_action_needed, case_entry) = "ALIGN INTERVIEW DATES" Then
		prev_add_to_list = WORKING_LIST_CASES_ARRAY(add_to_daily_worklist, case_entry)
		WORKING_LIST_CASES_ARRAY(add_to_daily_worklist, case_entry) = True
		If IsDate(WORKING_LIST_CASES_ARRAY(align_intv_revw_dt, case_entry)) = True Then
			If DateDiff("d", WORKING_LIST_CASES_ARRAY(align_intv_revw_dt, case_entry), date) < 7 Then WORKING_LIST_CASES_ARRAY(add_to_daily_worklist, case_entry) = prev_add_to_list
		End If

		update_sql = True
	End If

	If IsDate(WORKING_LIST_CASES_ARRAY(align_intv_revw_dt, case_entry)) = True Then
		WORKING_LIST_CASES_ARRAY(script_notes_info, case_entry) = WORKING_LIST_CASES_ARRAY(script_notes_info, case_entry) & "@#$%REVIEWED ON: " & WORKING_LIST_CASES_ARRAY(align_intv_revw_dt, case_entry) & "@%$#@"
			' align_intv_revw_dt
			' "@#$%REVIEWED ON: @%$#@"
	End If
	WORKING_LIST_CASES_ARRAY(last_wl_date, case_entry) = #1/1/1900#
	If WORKING_LIST_CASES_ARRAY(add_to_daily_worklist, case_entry) = True Then
		WORKING_LIST_CASES_ARRAY(last_wl_date, case_entry) = date
		WORKING_LIST_CASES_ARRAY(script_notes_info, case_entry) = "STS-NR " & WORKING_LIST_CASES_ARRAY(script_notes_info, case_entry)
		count_cases_on_wl = count_cases_on_wl + 1
	End If

	If update_sql = True Then
		' MsgBox "GOING TO UPDATE" & vbCr & vbCr & MAXIS_case_number

		On Error Resume Next
		'declare the SQL statement that will query the database
		objWorkSQL = "SELECT * FROM ES.ES_OnDemanCashAndSnapBZProcessed"

		'Creating objects for Access
		Set objWorkConnection = CreateObject("ADODB.Connection")
		Set objWorkRecordSet = CreateObject("ADODB.Recordset")

		'This is the file path for the statistics Access database.
		objWorkConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
		' objWorkRecordSet.Open objWorkSQL, objWorkConnection

		objWorkRecordSet.Open "DELETE FROM ES.ES_OnDemanCashAndSnapBZProcessed WHERE CaseNumber = '" & MAXIS_case_number & "'", objWorkConnection, adOpenStatic, adLockOptimistic

		objWorkRecordSet.Close
		objWorkConnection.Close

		Set objWorkRecordSet=nothing
		Set objWorkConnection=nothing
		Set objWorkSQL=nothing

		'Creating objects for Access
		Set objWorkConnection = CreateObject("ADODB.Connection")
		Set objWorkRecordSet = CreateObject("ADODB.Recordset")

		'This is the file path for the statistics Access database.
		objWorkConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""

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

		objWorkRecordSet.Close
		objWorkConnection.Close

		Set objWorkRecordSet=nothing
		Set objWorkConnection=nothing
		Set objWorkSQL=nothing
		On Error Goto 0
		' MsgBox "UPDATED!!!" & vbCr & vbCr & MAXIS_case_number

	End If
Next

end_msg =  "SQL updated for Interview Align issue. " & vbCr & "Total updated: " & count_cases_on_wl
Call script_end_procedure(end_msg)