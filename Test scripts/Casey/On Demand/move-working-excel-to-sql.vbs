
'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "TEST.vbs"
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

working_excel_file_path = t_drive & "/Eligibility Support/Restricted/QI - Quality Improvement/REPORTS/On Demand Waiver/Working Excel.xlsx"   'THIS IS THE REAL ONE

'Opens Excel file here, as it needs to populate the dialog with the details from the spreadsheet.
call excel_open(working_excel_file_path, True, True, ObjWorkExcel, objWorkWorkbook)

'Setting constants
Const adOpenStatic = 3
Const adLockOptimistic = 3

objSQL = "SELECT * FROM ES.ES_OnDemanCashAndSnapBZProcessed"

'Creating objects for Access
Set objConnection = CreateObject("ADODB.Connection")
Set objRecordSet = CreateObject("ADODB.Recordset")

'This is the file path for the statistics Access database.
objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" & "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""

' objRecordSet.Open "DELETE FROM ES.ES_OnDemanCashAndSnapBZProcessed", objConnection, adOpenStatic, adLockOptimistic
' call script_end_procedure("DELETED")
' objRecordSet.Open objSQL, objConnection
'
' objRecordSet.Open "INSERT INTO ES.ES_OnDemanCashAndSnapBZProcessed (CaseNumber, CaseName, ApplDate, InterviewDate, Day_30, REPT_PND2Days)" & _
'                   "VALUES ('" & ALL_PENDING_CASES_ARRAY(case_number, add_a_case) &  "', '" & _
'                                 ALL_PENDING_CASES_ARRAY(client_name, add_a_case) &  "', '" & _
'                                 ALL_PENDING_CASES_ARRAY(application_date, add_a_case) &  "', '" & _
'                                 ALL_PENDING_CASES_ARRAY(interview_date, add_a_case) &  "', '" & _
'                                 ALL_PENDING_CASES_ARRAY(data_day_30, add_a_case) &  "', '" & _
'                                 ALL_PENDING_CASES_ARRAY(rept_pnd2_listed_days, add_a_case) & "')"
'
'

row = 2
Do
    case_number_to_assess = trim(objWorkExcel.Cells(row, 2).Value)  'getting the case number in the Working Excel sheet

    next_action_array = ObjWorkExcel.Cells(row, next_action_col) & "|"
    next_action_array = next_action_array & ObjWorkExcel.Cells(row, script_notes_col)
    next_action_array = replace(next_action_array, "'", "")


    quest_intvw_array = ObjWorkExcel.Cells(row, quest_intvw_date_col) & "|"
    quest_intvw_array = quest_intvw_array & ObjWorkExcel.Cells(row, other_county_col) & "|"
    quest_intvw_array = quest_intvw_array & ObjWorkExcel.Cells(row, closed_in_30_col)
    quest_intvw_array = replace(quest_intvw_array, "'", "")

    resolved_array = "|" & "|" & "|"    'questionable interview resolved | day 30 resolved | out of county resolved | second app resolved'

    If Instr(ObjWorkExcel.Cells(row, worker_notes_col), "PRIV") = 0 Then priv_case = False
    If Instr(ObjWorkExcel.Cells(row, worker_notes_col), "PRIV") <> 0 Then priv_case = True
    denied_array = ObjWorkExcel.Cells(row, need_deny_col) & "|"
    denied_array = denied_array & ObjWorkExcel.Cells(row, worker_notes_col) & "|"
    denied_array = denied_array & priv_case
    denied_array = replace(denied_array, "'", "")

    ' next_action_array = next_action_array & " "
    wl_case_name = replace(ObjWorkExcel.Cells(row, case_name_col), "'", "")

    wl_app_date = ObjWorkExcel.Cells(row, app_date_col)
    wl_app_date = DateAdd("d", 0, wl_app_date)

    wl_intvw_date = ObjWorkExcel.Cells(row, intvw_date_col)
    If IsDate(wl_intvw_date) = True Then
        wl_intvw_date = DateAdd("d", 0, wl_intvw_date)
    Else
        wl_intvw_date = "''"
        wl_intvw_date = Null
    End If

    wl_day_30 = ObjWorkExcel.Cells(row, day_30_col)
    wl_day_30 = DateAdd("d", 0, wl_day_30)

    wl_second_app_date = ObjWorkExcel.Cells(row, second_app_date_col)
    If IsDate(wl_second_app_date) = True Then
        ' MsgBox wl_second_app_date
        wl_second_app_date = DateAdd("d", 0, wl_second_app_date)
    Else
        wl_second_app_date = "''"
        wl_second_app_date = Null
    End If

    wl_rept_pnd2_days = ObjWorkExcel.Cells(row, rept_pnd2_days_col)
    If trim(wl_rept_pnd2_days) = "" Then
        wl_rept_pnd2_days = "''"
        wl_rept_pnd2_days = Null
    Else
        wl_rept_pnd2_days = wl_rept_pnd2_days*1
    End If

    wl_appt_notc_date = ObjWorkExcel.Cells(row, appt_notc_date_col)
    If IsDate(wl_appt_notc_date) = True Then
        wl_appt_notc_date = DateAdd("d", 0, wl_appt_notc_date)
    Else
        wl_appt_notc_date = "''"
        wl_appt_notc_date = Null
    End If

    wl_appt_date = ObjWorkExcel.Cells(row, appt_date_col)
    If IsDate(wl_appt_date) = True Then
        wl_appt_date = DateAdd("d", 0, wl_appt_date)
    Else
        wl_appt_date = "''"
        wl_appt_date = Null
    End If

    wl_nomi_date = ObjWorkExcel.Cells(row, nomi_date_col)
    If IsDate(wl_nomi_date) = True Then
        wl_nomi_date = DateAdd("d", 0, wl_nomi_date)
    Else
        wl_nomi_date = "''"
        wl_nomi_date = Null
    End If

    wl_recent_wl_date = ObjWorkExcel.Cells(row, recent_wl_date_col)
    If IsDate(wl_recent_wl_date) = True Then
        ' MsgBox wl_recent_wl_date
        wl_recent_wl_date = DateAdd("d", 0, wl_recent_wl_date)
    Else
        wl_recent_wl_date = "''"
        wl_recent_wl_date = Null
    End If

    this_is_aCase_number = ObjWorkExcel.Cells(row, case_nbr_col)

    '
    objRecordSet.Open "INSERT INTO ES.ES_OnDemanCashAndSnapBZProcessed (CaseNumber, CaseName, ApplDate, InterviewDate, Day_30, DaysPending, SnapStatus, CashStatus, SecondApplicationDate, REPT_PND2Days, QuestionableInterview, Resolved, ApptNoticeDate, ApptDate, Confirmation, NOMIDate, Confirmation2, DenialNeeded, NextActionNeeded, AddedtoWorkList)" & _
                      "VALUES ('" & ObjWorkExcel.Cells(row, case_nbr_col) &  "', '" & _
                                    wl_case_name &  "', '" & _
                                    wl_app_date &  "', '" & _
                                    wl_intvw_date &  "', '" & _
                                    wl_day_30 &  "', '" & _
                                    "" &  "', '" & _
                                    ObjWorkExcel.Cells(row, snap_stat_col) &  "', '" & _
                                    ObjWorkExcel.Cells(row, cash_stat_col) &  "', '" & _
                                    wl_second_app_date &  "', '" & _
                                    wl_rept_pnd2_days &  "', '" & _
                                    ObjWorkExcel.Cells(row, quest_intvw_date_col) &  "', '" & _
                                    "" &  "', '" & _
                                    wl_appt_notc_date &  "', '" & _
                                    wl_appt_date &  "', '" & _
                                    ObjWorkExcel.Cells(row, appt_notc_confirm_col) &  "', '" & _
                                    wl_nomi_date &  "', '" & _
                                    ObjWorkExcel.Cells(row, nomi_confirm_col) &  "', '" & _
                                    ObjWorkExcel.Cells(row, need_deny_col) &  "', '" & _
                                    ObjWorkExcel.Cells(row, next_action_col) &  "', '" & _
                                    wl_recent_wl_date & "')", objConnection, adOpenStatic, adLockOptimistic




    ' If IsDate(ObjWorkExcel.Cells(row, recent_wl_date_col)) = True Then
    '     objRecordSet.Open "UPDATE ES.ES_CasesPending SET AddedtoWorkList = ''" & ObjWorkExcel.Cells(row, recent_wl_date_col) & "'' " &_
    ' End If
    ' AddedtoWorkList - wl_recent_wl_date

    ' If IsDate(ObjWorkExcel.Cells(row, second_app_date_col)) = True Then
    '     objRecordSet.Open "UPDATE ES.ES_CasesPending SET SecondApplicationDate = ''" & ObjWorkExcel.Cells(row, second_app_date_col) & "'' " &_
    ' End If
    ' SecondApplicationDate - wl_second_app_date

    ' NextActionNeeded - next_action_array

    ' 0 -- Next Action Needed
    ' 1 -- Worker Notes
    ' 2 -- Script Notes
    ' 3 -- Case was in other county
    ' 4 -- Case closed in past 30 days
    ' 5 -- PRIV Case
    ' 6 -- Out of county resolved
    ' 7 -- closed in 30 days resolved
    ' 8 -- Subsequent Application resolved
    ' array_of_script_notes = split(actions_detail_var, "|")

    ' CASES_NO_LONGER_WORKING(case_in_other_co, case_removed) = ObjWorkExcel.Cells(row, other_county_col)
    ' CASES_NO_LONGER_WORKING(case_closed_in_30, case_removed) = ObjWorkExcel.Cells(row, closed_in_30_col)




    row = row + 1

    next_case_number = trim(objWorkExcel.Cells(row, 1).Value)       'looking for when to exit the loop - when we reach the end of the Working Excel
Loop Until next_case_number = ""

call script_end_procedure("WE MADE IT TO THE END")
