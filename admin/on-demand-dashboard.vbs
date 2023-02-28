'GATHERING STATS===========================================================================================
name_of_script = "ADMIN - On Demand Applications Dashboard.vbs"
start_time = timer
STATS_counter = 1
STATS_manualtime = 300
STATS_denominatinon = "C"
'END OF STATS BLOCK===========================================================================================

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
    IF on_the_desert_island = TRUE Then
        FuncLib_URL = "\\hcgg.fr.co.hennepin.mn.us\lobroot\hsph\team\Eligibility Support\Scripts\Script Files\desert-island\MASTER FUNCTIONS LIBRARY.vbs"
        Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
        Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
        text_from_the_other_script = fso_command.ReadAll
        fso_command.Close
        Execute text_from_the_other_script
    ELSEIF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
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
call changelog_update("12/06/2022", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'Information about this script and how it works
' STS-NR            'Status - Needs Review'
' STS-RC            'Status - Review Completed'
' STS-IP-WFXXXX     'Status - In Progress - Worker number'
' STS-HD-WFXXXX     'Status - HolD - Worker number'

'DECLARATIONS BLOCK ========================================================================================================

MAXIS_case_number = ""
assigned_case_name = ""
table_application_date = ""
assigned_application_date = ""
table_interview_date = ""
assigned_interview_date = ""
table_day_30 = ""
assigned_day_30 = ""
assigned_days_pending = ""
assigned_snap_status = ""
assigned_cash_status = ""
table_2nd_application_date = ""
assigned_2nd_application_date = ""
assigned_rept_pnd2_days = ""
assigned_questionable_interview = ""
assigned_questionable_interview_resolve = ""
table_appt_notice_date = ""
assigned_appt_notice_date = ""
table_appt_date = ""
assigned_appt_date = ""
assigned_appt_notc_confirmation = ""
table_nomi_date = ""
assigned_nomi_date = ""
assigned_nomi_confirmation = ""
assigned_denial_needed = ""
assigned_next_action_needed = ""
table_added_to_work_list = ""
assigned_added_to_work_list = ""
assigned_2nd_application_date_resolve = ""
assigned_closed_recently = ""
assigned_closed_recently_resolve = ""
assigned_out_of_county = ""
assigned_out_of_county_resolve = ""
assigned_tracking_notes = ""
case_review_notes = ""
case_on_hold = False
case_in_progress = False
completed_reviews = 0
reviews_completed_by_me = 0
reviews_still_needed = 0

assigned_worker  = ""
assigned_date  = ""
assigned_start_time  = ""
assigned_end_time  = ""
assigned_hold_1_start_time  = ""
assigned_hold_1_end_time  = ""
assigned_hold_2_start_time  = ""
assigned_hold_2_end_time  = ""
assigned_hold_3_start_time  = ""
assigned_hold_3_end_time  = ""


ADMIN_run = False
BULK_Run_completed = False
worker_on_task = False
total_cases_for_review = 0
cases_with_review_completed = 0
cases_waiting_for_review = 0
cases_on_hold = 0
case_nbr_in_progress = ""
admin_count_NR = 0
admin_count_RC = 0
admin_count_IP = 0
admin_count_HD = 0
ADMIN_list_workers_RC = "~"
ADMIN_list_workers_IP = "~"
ADMIN_list_workers_HD = "~"

current_day_work_tracking_folder = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\On Demand Waiver\QI On Demand Daily Assignment\Archive\WIP\"
'txt file name format
' USERID-MM-DD-YY-CASENUMBER
curr_day = DatePart("d", date)
curr_day = right("00" & curr_day, 2)
file_date = CM_mo & "-" & curr_day & "-" & CM_yr

txt_file_name = user_ID_for_validation & "_" & MAXIS_case_number & "_" & file_date & ".txt"
od_revw_tracking_file_path = current_day_work_tracking_folder  & txt_file_name

' txt_file_name = "expedited_determination_detail_" & MAXIS_case_number & "_" & replace(replace(replace(now, "/", "_"),":", "_")," ", "_") & ".txt"
' MsgBox exp_info_file_path
end_msg = "Script Run is completed."

'BUTTONS'
complete_bulk_run_btn   = 1001
bulk_run_details_btn    = 1002
bulk_run_incomplete_btn = 1003
get_new_case_btn        = 2001
work_list_details_btn   = 2002
resume_hold_case_btn    = 2003
complete_review_btn     = 2004
hold_case_btn           = 2005
finish_work_day_btn     = 3001
test_access_btn         = 4001
close_dialog_btn        = 5001
release_IP_btn          = 6001
release_HD_btn          = 6002
finish_day_btn          = 6003


const case_nbr_const    = 0
const radio_btn_const   = 1
const case_notes_const  = 2
const last_const        = 10

Dim CASES_ON_HOLD_ARRAY()
ReDim CASES_ON_HOLD_ARRAY(last_const, 0)

const case_number_const                             = 00
const assigned_worker_const                         = 01
const assigned_date_const                           = 02
const assigned_start_time_const                     = 03
const assigned_end_time_const                       = 04
const assigned_hold_1_start_time_const              = 05
const assigned_hold_1_end_time_const                = 06
const assigned_hold_2_start_time_const              = 07
const assigned_hold_2_end_time_const                = 08
const assigned_hold_3_start_time_const              = 09
const assigned_hold_3_end_time_const                = 10
const assigned_case_name_const                      = 11
const assigned_application_date_const               = 12
const assigned_interview_date_const                 = 13
const assigned_day_30_const                         = 14
const assigned_days_pending_const                   = 15
const assigned_snap_status_const                    = 16
const assigned_cash_status_const                    = 17
const assigned_2nd_application_date_const           = 18
const assigned_rept_pnd2_days_const                 = 19
const assigned_questionable_interview_const         = 20
const assigned_questionable_interview_resolve_const = 21
const assigned_appt_notice_date_const               = 22
const assigned_appt_date_const                      = 23
const assigned_appt_notc_confirmation_const         = 24
const assigned_nomi_date_const                      = 25
const assigned_nomi_confirmation_const              = 26
const assigned_denial_needed_const                  = 27
const assigned_next_action_needed_const             = 28
const assigned_added_to_work_list_const             = 29
const assigned_2nd_application_date_resolve_const   = 30
const assigned_closed_recently_const                = 31
const assigned_closed_recently_resolve_const        = 32
const assigned_out_of_county_const                  = 33
const assigned_out_of_county_resolve_const          = 34
const case_review_notes_const                       = 35
const final_const                                   = 50

Dim COMPLETED_REVIEWS_ARRAY()
ReDim COMPLETED_REVIEWS_ARRAY(final_const, 0)

const wrkr_id_const         = 0
const wrkr_name_const       = 1
const case_status_const     = 2
const admin_radio_btn_const = 3
const admin_wrkr_last_const = 4

Dim ADMIN_worker_list_array()
ReDim ADMIN_worker_list_array(admin_wrkr_last_const, 0)

const Worker_col                        = 1
const AssignedDate_col                  = 2
const CaseNumber_col                    = 3
const CaseName_col                      = 4
const ApplDate_col                      = 5
const InterviewDate_col                 = 6
const Day_30_dash_col                   = 7
const DaysPending_col                   = 8
const SnapStatus_col                    = 9
const CashStatus_col                    = 10
const SecondApplicationDate_col         = 11
const REPT_PND2Days_col                 = 12
const QuestionableInterview_col         = 13
const Resolved_col                      = 14
const ApptNoticeDate_col                = 15
const ApptDate_col                      = 16
const Confirmation_col                  = 17
const NOMIDate_col                      = 18
const Confirmation2_col                 = 19
const DenialNeeded_col                  = 20
const NextActionNeeded_col              = 21
const AddedtoWorkList_col               = 22
const SecondApplicationDateNotes_col    = 23
const ClosedInPast30Days_col            = 24
const ClosedInPast30DaysNotes_col       = 25
const StartedOutOfCounty_col            = 26
const StartedOutOfCountyNotes_col       = 27
const TrackingNotes_col                 = 28
const CaseSelectedTime_col              = 29
const Hold1Start_col                    = 30
const Hold1End_col                      = 31
const Hold2Start_col                    = 32
const Hold2End_col                      = 33
const Hold3Start_col                    = 34
const Hold3End_col                      = 35
const CaseCompletedTime_col             = 36

'END DECLARATIONS ==========================================================================================================

'FUNCTIONS BLOCK ===========================================================================================================

function assess_worklist_to_finish_day()
    case_on_hold = False
    case_in_progress = False
    completed_reviews = 0
    reviews_completed_by_me = 0
    reviews_still_needed = 0

    'Access the SQL Table
    'declare the SQL statement that will query the database
    objSQL = "SELECT * FROM ES.ES_OnDemanCashAndSnapBZProcessed"

    'Creating objects for Access
    Set objConnection = CreateObject("ADODB.Connection")
    Set objRecordSet = CreateObject("ADODB.Recordset")

    'This is the file path for the statistics Access database.
    objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
    objRecordSet.Open objSQL, objConnection

    Do While NOT objRecordSet.Eof
        case_tracking_notes = objRecordSet("TrackingNotes")

        If Instr(case_tracking_notes, "STS-RC") <> 0 Then
            completed_reviews = completed_reviews + 1
            If Instr(case_tracking_notes, user_ID_for_validation) <> 0 Then reviews_completed_by_me = reviews_completed_by_me + 1
        End If
        If Instr(case_tracking_notes, "STS-NR") <> 0 Then reviews_still_needed = reviews_still_needed + 1

        If Instr(case_tracking_notes, "STS-HD") Then
            If Instr(case_tracking_notes, user_ID_for_validation) <> 0 Then case_on_hold = True
        End If

        If Instr(case_tracking_notes, "STS-IP") Then
            If Instr(case_tracking_notes, user_ID_for_validation) <> 0 Then case_in_progress = True
        End If
        objRecordSet.MoveNext
    Loop

    'close the connection and recordset objects to free up resources
    objRecordSet.Close
    objConnection.Close
    Set objRecordSet=nothing
    Set objConnection=nothing
end function

function assign_a_case()
    If local_demo = False Then
        If ButtonPressed = get_new_case_btn Then
            objSQL = "SELECT * FROM ES.ES_OnDemanCashAndSnapBZProcessed"

            'Creating objects for Access
            Set objConnection = CreateObject("ADODB.Connection")
            Set objRecordSet = CreateObject("ADODB.Recordset")

            'This is the file path for the statistics Access database.
            objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
            objRecordSet.Open objSQL, objConnection

            'Read the whole table
            Do While NOT objRecordSet.Eof    'Read the whole table
                case_tracking_notes = objRecordSet("TrackingNotes")
                If Instr(case_tracking_notes, "STS-NR") <> 0 Then
                    set_variables_from_SQL
                    Exit Do
                End If
                objRecordSet.MoveNext
            Loop

            'close the connection and recordset objects to free up resources
            objRecordSet.Close
            objConnection.Close
            Set objRecordSet=nothing
            Set objConnection=nothing

            txt_file_name = user_ID_for_validation & "_" & MAXIS_case_number & "_" & file_date & ".txt"
            od_revw_tracking_file_path = current_day_work_tracking_folder  & txt_file_name
            assigned_start_time = time
            assigned_date = date
            create_tracking_cookie
        End If

        If ButtonPressed = resume_hold_case_btn Then
            'Creating objects for Access
            Set objConnection = CreateObject("ADODB.Connection")
            Set objRecordSet = CreateObject("ADODB.Recordset")

            'This is the file path for the statistics Access database.
            objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
            objRecordSet.Open "SELECT * FROM ES.ES_OnDemanCashAndSnapBZProcessed WHERE CaseNumber = '" & resume_case_number & "'", objConnection

            set_variables_from_SQL

            'close the connection and recordset objects to free up resources
            objRecordSet.Close
            objConnection.Close
            Set objRecordSet=nothing
            Set objConnection=nothing
        End If
    Else
        MAXIS_case_number                       = "318040"
        ' assigned_ = objRecordSet("CaseNumber")
        assigned_case_name                      = "ZORO, RORONOA"
        assigned_application_date               = "11/23/2022"
        assigned_interview_date                 = ""
        assigned_day_30                         = "12/23/2022"
        assigned_days_pending                   = "12"
        assigned_snap_status                    = "Pending"
        assigned_cash_status                    = ""
        assigned_2nd_application_date           = ""
        assigned_rept_pnd2_days                 = "12"
        assigned_questionable_interview         = ""
        assigned_questionable_interview_resolve = ""
        assigned_appt_notice_date               = "11/25/2022"
        assigned_appt_date                      = "12/5/2022"
        assigned_appt_notc_confirmation         = "Y"
        assigned_nomi_date                      = "12/5/2022"
        assigned_nomi_confirmation              = "Y"
        assigned_denial_needed                  = ""
        assigned_next_action_needed             = "DO THIS NEXT"
        assigned_added_to_work_list             = date
        assigned_2nd_application_date_resolve   = ""
        assigned_closed_recently                = ""
        assigned_closed_recently_resolve        = ""
        assigned_out_of_county                  = ""
        assigned_out_of_county_resolve          = ""
        case_review_notes                 = "TrackingNotes"
    End If
    end_msg = end_msg & vbCr & vbCr & "You have a case selected for review: " & MAXIS_case_number
    assigned_tracking_notes = "STS-IP-"&user_ID_for_validation & " " & case_review_notes

    If local_demo = False Then
        'Creating objects for Access
        Set objConnection = CreateObject("ADODB.Connection")
        Set objRecordSet = CreateObject("ADODB.Recordset")

        'This is the BZST connection to SQL Database'
        objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""

        'delete a record if the case number matches
        objRecordSet.Open "UPDATE ES.ES_OnDemanCashAndSnapBZProcessed SET CaseNumber = '" & MAXIS_case_number & "', " &_
                                                                          "CaseName = '" & assigned_case_name & "', " &_
                                                                          "ApplDate = '" & table_application_date & "', " &_
                                                                          "InterviewDate = '" & table_interview_date & "', " &_
                                                                          "Day_30 = '" & table_day_30 & "', " &_
                                                                          "DaysPending = '" & assigned_days_pending & "', " &_
                                                                          "SnapStatus = '" & assigned_snap_status & "', " &_
                                                                          "CashStatus = '" & assigned_cash_status & "', " &_
                                                                          "SecondApplicationDate = '" & table_2nd_application_date & "', " &_
                                                                          "REPT_PND2Days = '" & assigned_rept_pnd2_days & "', " &_
                                                                          "QuestionableInterview = '" & assigned_questionable_interview & "', " &_
                                                                          "Resolved = '" & assigned_questionable_interview_resolve & "', " &_
                                                                          "ApptNoticeDate = '" & table_appt_notice_date & "', " &_
                                                                          "ApptDate = '" & table_appt_date & "', " &_
                                                                          "Confirmation = '" & assigned_appt_notc_confirmation & "', " &_
                                                                          "NOMIDate = '" & table_nomi_date & "', " &_
                                                                          "Confirmation2 = '" & assigned_nomi_confirmation & "', " &_
                                                                          "DenialNeeded = '" & assigned_denial_needed & "', " &_
                                                                          "NextActionNeeded = '" & assigned_next_action_needed & "', " &_
                                                                          "AddedtoWorkList = '" & table_added_to_work_list & "', " &_
                                                                          "SecondApplicationDateNotes = '" & assigned_2nd_application_date_resolve & "', " &_
                                                                          "ClosedInPast30Days = '" & assigned_closed_recently & "', " &_
                                                                          "ClosedInPast30DaysNotes = '" & assigned_closed_recently_resolve & "', " &_
                                                                          "StartedOutOfCounty = '" & assigned_out_of_county & "', " &_
                                                                          "StartedOutOfCountyNotes = '" & assigned_out_of_county_resolve & "', " &_
                                                                          "TrackingNotes = '" & assigned_tracking_notes & "' WHERE CaseNumber = '" & MAXIS_case_number & "'", objConnection
        'close the connection and recordset objects to free up resources
        objConnection.Close
        Set objRecordSet=nothing
        Set objConnection=nothing
    End If

    txt_file_name = user_ID_for_validation & "_" & MAXIS_case_number & "_" & file_date & ".txt"
    od_revw_tracking_file_path = current_day_work_tracking_folder  & txt_file_name
    If ButtonPressed = get_new_case_btn Then
        assigned_start_time = time
        call create_tracking_cookie
    End If
    If ButtonPressed = resume_hold_case_btn Then
        call read_tracking_cookie

        If assigned_hold_1_start_time <> "" AND assigned_hold_1_end_time = "" Then
            assigned_hold_1_end_time = time
        ElseIf assigned_hold_2_start_time <> "" AND assigned_hold_2_end_time = "" Then
            assigned_hold_2_end_time = time
        ElseIf assigned_hold_3_start_time <> "" AND assigned_hold_3_end_time = "" Then
            assigned_hold_3_end_time = time
        End If

        call update_tracking_cookie("RESUME")
    End If
    worker_on_task = True

    Call Back_to_SELF
    Call navigate_to_MAXIS_screen("CASE", "CURR")
    'Once this function ends, the script will move to the final fucnctionality
    'that dissplays a dialog to allow the review to be completed
end function

function complete_admin_functions()
    Do
        Do
            err_msg = ""

            dlg_len = 165 + 10 * (UBound(ADMIN_worker_list_array, 2)+1)
            grp_len = 110 + 10 * (UBound(ADMIN_worker_list_array, 2)+1)

            Dialog1 = ""
            BeginDialog Dialog1, 0, 0, 451, 180, "On Demand Applications Dashboard"
              ButtonGroup ButtonPressed
                Text 170, 10, 135, 10, "On Demand Applications Dashboard"
                GroupBox 10, 25, 430, 130, "On Demand Case List Information"
                Text 20, 40, 230, 10, "Today " & total_cases_for_review & " cases require review. These cases are currently:"
                Text 40, 55, 110, 10, "Reviews still needed: " & admin_count_NR
                Text 40, 70, 115, 10, "Reviews In Progress: " & admin_count_IP
                PushButton 190, 70, 130, 13, "Release In Progress Assignments", release_IP_btn
                y_pos = 85
                If IsArray(IP_id_ARRAY) = True then
                    OptionGroup RadioGroupIP
                    For worker_indc = 0 to UBound(ADMIN_worker_list_array, 2)
                        If ADMIN_worker_list_array(case_status_const, worker_indc) = "IP" Then
                            RadioButton 50, y_pos, 300, 10, ADMIN_worker_list_array(wrkr_name_const, worker_indc) & " - " & ADMIN_worker_list_array(wrkr_id_const, worker_indc), btn_hold'', ADMIN_worker_list_array(admin_radio_btn_const, worker_indc)
                            y_pos = y_pos + 10
                        End If
                    Next
                End If
                y_pos = y_pos + 5

                Text 40, y_pos, 140, 10, "Reviews on Hold: " & admin_count_HD
                PushButton 190, y_pos, 130, 13, "Release Hold Assignments", release_HD_btn
                y_pos = y_pos + 15
                If IsArray(HD_id_ARRAY) = True then
                    OptionGroup RadioGroupHD
                    For worker_indc = 0 to UBound(ADMIN_worker_list_array, 2)
                        If ADMIN_worker_list_array(case_status_const, worker_indc) = "HD" Then
                            RadioButton 50, y_pos, 300, 10, ADMIN_worker_list_array(wrkr_name_const, worker_indc) & " - " & ADMIN_worker_list_array(wrkr_id_const, worker_indc), btn_hold', ADMIN_worker_list_array(admin_radio_btn_const, worker_indc)
                            y_pos = y_pos + 10
                        End If
                    Next
                End If
                y_pos = y_pos + 5

                Text 40, y_pos, 145, 10, "Reviews Completed: " & admin_count_RC
                PushButton 190, y_pos, 130, 13, "Run Finish Day", finish_day_btn
                y_pos = y_pos + 15
                If IsArray(RC_id_ARRAY) = True then
                    OptionGroup RadioGroupRC
                    For worker_indc = 0 to UBound(ADMIN_worker_list_array, 2)
                        If ADMIN_worker_list_array(case_status_const, worker_indc) = "RC" Then
                            RadioButton 50, y_pos, 300, 10, ADMIN_worker_list_array(wrkr_name_const, worker_indc) & " - " & ADMIN_worker_list_array(wrkr_id_const, worker_indc), btn_hold', ADMIN_worker_list_array(admin_radio_btn_const, worker_indc)
                            y_pos = y_pos + 10
                        End If
                    Next
                End If
                y_pos = y_pos + 5

                OkButton 335, 160, 50, 15
                CancelButton 390, 160, 50, 15
                EditBox 500, 300, 50, 15, fake_edit_box
            EndDialog

            Dialog Dialog1
            cancel_without_confirmation

            worker_number_to_resolve = ""
            If ButtonPressed = release_IP_btn Then
                MsgBox "RadioGroupIP - " & RadioGroupIP
                For worker_indc = 0 to UBound(ADMIN_worker_list_array, 2)
                    If RadioGroupIP = ADMIN_worker_list_array(admin_radio_btn_const, worker_indc) Then worker_number_to_resolve = ADMIN_worker_list_array(wrkr_id_const, worker_indc)
                Next
            ElseIf ButtonPressed = release_HD_btn Then
                MsgBox "RadioGroupHD - " & RadioGroupHD
                For worker_indc = 0 to UBound(ADMIN_worker_list_array, 2)
                    If RadioGroupHD = ADMIN_worker_list_array(admin_radio_btn_const, worker_indc) Then worker_number_to_resolve = ADMIN_worker_list_array(wrkr_id_const, worker_indc)
                Next
            ElseIf ButtonPressed = finish_day_btn Then
                MsgBox "RadioGroupRC - " & RadioGroupRC
                For worker_indc = 0 to UBound(ADMIN_worker_list_array, 2)
                    If RadioGroupRC = ADMIN_worker_list_array(admin_radio_btn_const, worker_indc) Then worker_number_to_resolve = ADMIN_worker_list_array(wrkr_id_const, worker_indc)
                Next
            End If

            MsgBox "worker_number_to_resolve - " & worker_number_to_resolve
            For tester = 0 to UBound(tester_array)                         'looping through all of the testers
                ' pulling QI members by supervisor from the Complete List of Testers
                If tester_array(tester).tester_id_number = worker_number_to_resolve Then
                    worker_full_name_to_resolve = tester_array(tester).tester_full_name
                    qi_worker_supervisor_email = tester_array(tester).tester_supervisor_email
                    qi_worker_first_name = tester_array(tester).tester_first_name
                    If tester_array(tester).tester_supervisor_name = "Tanya Payne" Then qi_member_identified = True
                    If tester_array(tester).tester_population = "BZ" Then qi_member_identified = True
                    assigned_worker = tester_array(tester).tester_full_name
                    ' MsgBox "user_ID_for_validation - " & user_ID_for_validation & vbCr & "tester_array(tester).tester_id_number - " & tester_array(tester).tester_id_number & vbCr & "qi_member_identified - " & qi_member_identified
                End If
            Next

            case_to_fix_found = False
            Do
                'declare the SQL statement that will query the database
                objSQL = "SELECT * FROM ES.ES_OnDemanCashAndSnapBZProcessed"

                'Creating objects for Access
                Set objConnection = CreateObject("ADODB.Connection")
                Set objRecordSet = CreateObject("ADODB.Recordset")

                'This is the file path for the statistics Access database.
                objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
                objRecordSet.Open objSQL, objConnection

                Do While NOT objRecordSet.Eof
                    case_tracking_notes = objRecordSet("TrackingNotes")
                    worklist_case_number = objRecordSet("CaseNumber")

                    If ButtonPressed = release_IP_btn Then
                        If Instr(case_tracking_notes, "STS-IP") Then
                            If Instr(case_tracking_notes, worker_number_to_resolve) <> 0 Then
                                case_to_fix_found = True
                                case_tracking_notes = replace(case_tracking_notes, "STS-IP-"&worker_number_to_resolve, "")
                                case_tracking_notes = replace(case_tracking_notes, "STS-HD-"&worker_number_to_resolve, "")
                                case_tracking_notes = replace(case_tracking_notes, "STS-RC-"&worker_number_to_resolve, "")
                                case_tracking_notes = replace(case_tracking_notes, "STS-RC", "")
                                case_tracking_notes = trim(case_tracking_notes)
                                case_tracking_notes = "STS-NR " & case_tracking_notes
                                case_tracking_notes = trim(case_tracking_notes)
                                Exit Do
                                ' objRecordSet.Update "UPDATE ES.ES_OnDemanCashAndSnapBZProcessed SET TrackingNotes = '" & case_review_notes & "'", objConnection

                            End If
                        End if
                    ElseIf ButtonPressed = release_HD_btn Then
                        If Instr(case_tracking_notes, "STS-HD") Then
                            If Instr(case_tracking_notes, worker_number_to_resolve) <> 0 Then
                                case_to_fix_found = True
                                case_tracking_notes = replace(case_tracking_notes, "STS-IP-"&worker_number_to_resolve, "")
                                case_tracking_notes = replace(case_tracking_notes, "STS-HD-"&worker_number_to_resolve, "")
                                case_tracking_notes = replace(case_tracking_notes, "STS-RC-"&worker_number_to_resolve, "")
                                case_tracking_notes = replace(case_tracking_notes, "STS-RC", "")
                                case_tracking_notescase_tracking_notescase_tracking_notes = trim(case_tracking_notes)
                                case_tracking_notes = "STS-NR " & case_tracking_notes
                                case_tracking_notes = trim(case_tracking_notes)
                                Exit Do
                                ' objRecordSet.Update "UPDATE ES.ES_OnDemanCashAndSnapBZProcessed SET TrackingNotes = '" & case_review_notes & "'", objConnection

                            End If
                        End if
                    End If

                    objRecordSet.MoveNext
                Loop

                'close the connection and recordset objects to free up resources
                objRecordSet.Close
                objConnection.Close
                Set objRecordSet=nothing
                Set objConnection=nothing


                Set objConnection = CreateObject("ADODB.Connection")
                Set objRecordSet = CreateObject("ADODB.Recordset")

                'This is the BZST connection to SQL Database'
                objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""

                'delete a record if the case number matches
                objRecordSet.Open "UPDATE ES.ES_OnDemanCashAndSnapBZProcessed SET TrackingNotes = '" & case_tracking_notes & "' WHERE CaseNumber = '" & worklist_case_number & "'", objConnection

                'close the connection and recordset objects to free up resources
                objConnection.Close
                Set objRecordSet=nothing
                Set objConnection=nothing

                If ButtonPressed = finish_day_btn Then
                    actual_user_ID_for_validation = user_ID_for_validation
                    user_ID_for_validation = worker_number_to_resolve
                    actual_assigned_worker = assigned_worker
                    assigned_worker = worker_full_name_to_resolve

                    Call assess_worklist_to_finish_day
                    If case_on_hold = False and case_in_progress = False Then
                        ' Call update_qi_worklist_at_end_of_working_day
                        Call create_assignment_report
                        end_msg = "Tracking log has been updated with work completed by " & assigned_worker & "."
                        call script_end_procedure(end_msg)
                    Else
                        loop_dlg_msg = "You cannot finish the work day with cases in progress or on hold." & vbCr
                        loop_dlg_msg = loop_dlg_msg & "The dialog will reappear, finish all reviews that have been started first." & vbCr & vbCr
                        loop_dlg_msg = loop_dlg_msg & "Once there are no cases on the worklist on hold or in progress the finish work day functionality will operate."
                        ' ButtonPressed = work_list_details_btn
                        MsgBox loop_dlg_msg
                        err_msg = "LOOP"
                    End If
                End If
            Loop until case_to_fix_found = False
	    Loop until err_msg = ""
	    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
	Loop until are_we_passworded_out = false					'loops until user passwords back in
end function

function create_assignment_report()
    'Access the SQL Table
    'declare the SQL statement that will query the database
    objSQL = "SELECT * FROM ES.ES_OnDemanCashAndSnapBZProcessed"

    'Creating objects for Access
    Set objConnection = CreateObject("ADODB.Connection")
    Set objRecordSet = CreateObject("ADODB.Recordset")

    'This is the file path for the statistics Access database.
    objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
    objRecordSet.Open objSQL, objConnection

    cases_completed_by_me = 0
    Do While NOT objRecordSet.Eof
        case_tracking_notes = objRecordSet("TrackingNotes")

        If Instr(case_tracking_notes, "STS-RC-"&user_ID_for_validation) <> 0 Then

            ReDim Preserve COMPLETED_REVIEWS_ARRAY(final_const, cases_completed_by_me)

            COMPLETED_REVIEWS_ARRAY(case_number_const, cases_completed_by_me)                             = objRecordSet("CaseNumber")
            COMPLETED_REVIEWS_ARRAY(assigned_case_name_const, cases_completed_by_me)                      = objRecordSet("CaseName")
            COMPLETED_REVIEWS_ARRAY(assigned_application_date_const, cases_completed_by_me)               = objRecordSet("ApplDate")
            COMPLETED_REVIEWS_ARRAY(assigned_interview_date_const, cases_completed_by_me)                 = objRecordSet("InterviewDate")
            COMPLETED_REVIEWS_ARRAY(assigned_day_30_const, cases_completed_by_me)                         = objRecordSet("Day_30")
            COMPLETED_REVIEWS_ARRAY(assigned_days_pending_const, cases_completed_by_me)                   = objRecordSet("DaysPending")
            COMPLETED_REVIEWS_ARRAY(assigned_snap_status_const, cases_completed_by_me)                    = objRecordSet("SnapStatus")
            COMPLETED_REVIEWS_ARRAY(assigned_cash_status_const, cases_completed_by_me)                    = objRecordSet("CashStatus")
            COMPLETED_REVIEWS_ARRAY(assigned_2nd_application_date_const, cases_completed_by_me)           = objRecordSet("SecondApplicationDate")
            COMPLETED_REVIEWS_ARRAY(assigned_rept_pnd2_days_const, cases_completed_by_me)                 = objRecordSet("REPT_PND2Days")
            COMPLETED_REVIEWS_ARRAY(assigned_questionable_interview_const, cases_completed_by_me)         = objRecordSet("QuestionableInterview")
            COMPLETED_REVIEWS_ARRAY(assigned_questionable_interview_resolve_const, cases_completed_by_me) = objRecordSet("Resolved")
            COMPLETED_REVIEWS_ARRAY(assigned_appt_notice_date_const, cases_completed_by_me)               = objRecordSet("ApptNoticeDate")
            COMPLETED_REVIEWS_ARRAY(assigned_appt_date_const, cases_completed_by_me)                      = objRecordSet("ApptDate")
            COMPLETED_REVIEWS_ARRAY(assigned_appt_notc_confirmation_const, cases_completed_by_me)         = objRecordSet("Confirmation")
            COMPLETED_REVIEWS_ARRAY(assigned_nomi_date_const, cases_completed_by_me)                      = objRecordSet("NOMIDate")
            COMPLETED_REVIEWS_ARRAY(assigned_nomi_confirmation_const, cases_completed_by_me)              = objRecordSet("Confirmation2")
            COMPLETED_REVIEWS_ARRAY(assigned_denial_needed_const, cases_completed_by_me)                  = objRecordSet("DenialNeeded")
            COMPLETED_REVIEWS_ARRAY(assigned_next_action_needed_const, cases_completed_by_me)             = objRecordSet("NextActionNeeded")
            COMPLETED_REVIEWS_ARRAY(assigned_added_to_work_list_const, cases_completed_by_me)             = objRecordSet("AddedtoWorkList")
            COMPLETED_REVIEWS_ARRAY(assigned_2nd_application_date_resolve_const, cases_completed_by_me)   = objRecordSet("SecondApplicationDateNotes")
            COMPLETED_REVIEWS_ARRAY(assigned_closed_recently_const, cases_completed_by_me)                = objRecordSet("ClosedInPast30Days")
            COMPLETED_REVIEWS_ARRAY(assigned_closed_recently_resolve_const, cases_completed_by_me)        = objRecordSet("ClosedInPast30DaysNotes")
            COMPLETED_REVIEWS_ARRAY(assigned_out_of_county_const, cases_completed_by_me)                  = objRecordSet("StartedOutOfCounty")
            COMPLETED_REVIEWS_ARRAY(assigned_out_of_county_resolve_const, cases_completed_by_me)          = objRecordSet("StartedOutOfCountyNotes")
            COMPLETED_REVIEWS_ARRAY(case_review_notes_const, cases_completed_by_me)                       = objRecordSet("TrackingNotes")

            cases_completed_by_me = cases_completed_by_me + 1
        End If
        objRecordSet.MoveNext
    Loop

    'close the connection and recordset objects to free up resources
    objRecordSet.Close
    objConnection.Close
    Set objRecordSet=nothing
    Set objConnection=nothing

    file_url = "T:\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\On Demand Waiver\QI On Demand Daily Assignment\QI On Demand Work Log.xlsx"
    Call excel_open(file_url, True, False, ObjExcel, objWorkbook)

    excel_row = 2

    Do While trim(ObjExcel.Cells(excel_row, Worker_col).value) <> ""
	    excel_row = excel_row + 1
    Loop

    For review = 0 to UBound(COMPLETED_REVIEWS_ARRAY, 2)
        MAXIS_case_number = COMPLETED_REVIEWS_ARRAY(case_number_const, review)
        assigned_date  = ""
        assigned_start_time  = ""
        assigned_end_time  = ""
        assigned_hold_1_start_time  = ""
        assigned_hold_1_end_time  = ""
        assigned_hold_2_start_time  = ""
        assigned_hold_2_end_time  = ""
        assigned_hold_3_start_time  = ""
        assigned_hold_3_end_time  = ""
        od_revw_tracking_file_path = ""

        txt_file_name = user_ID_for_validation & "_" & MAXIS_case_number & "_" & file_date & ".txt"
        od_revw_tracking_file_path = current_day_work_tracking_folder  & txt_file_name

        Call read_tracking_cookie

        case_review_notes = COMPLETED_REVIEWS_ARRAY(case_review_notes_const, review)
        case_review_notes = replace(case_review_notes, "STS-IP-"&user_ID_for_validation, "")
        case_review_notes = replace(case_review_notes, "STS-HD-"&user_ID_for_validation, "")
        case_review_notes = replace(case_review_notes, "STS-RC-"&user_ID_for_validation, "")
        case_review_notes = replace(case_review_notes, "STS-RC", "")
        case_review_notes = trim(case_review_notes)

        ObjExcel.Cells(excel_row, Worker_col).value = assigned_worker
        ObjExcel.Cells(excel_row, AssignedDate_col).value = assigned_date
        ObjExcel.Cells(excel_row, CaseNumber_col).value = COMPLETED_REVIEWS_ARRAY(case_number_const, review)
        ObjExcel.Cells(excel_row, CaseName_col).value = COMPLETED_REVIEWS_ARRAY(assigned_case_name_const, review)
        ObjExcel.Cells(excel_row, ApplDate_col).value = COMPLETED_REVIEWS_ARRAY(assigned_application_date_const, review)
        ObjExcel.Cells(excel_row, InterviewDate_col).value = COMPLETED_REVIEWS_ARRAY(assigned_interview_date_const, review)
        ObjExcel.Cells(excel_row, Day_30_dash_col).value = COMPLETED_REVIEWS_ARRAY(assigned_day_30_const, review)
        ObjExcel.Cells(excel_row, DaysPending_col).value = COMPLETED_REVIEWS_ARRAY(assigned_days_pending_const, review)
        ObjExcel.Cells(excel_row, SnapStatus_col).value = COMPLETED_REVIEWS_ARRAY(assigned_snap_status_const, review)
        ObjExcel.Cells(excel_row, CashStatus_col).value = COMPLETED_REVIEWS_ARRAY(assigned_cash_status_const, review)
        ObjExcel.Cells(excel_row, SecondApplicationDate_col).value = COMPLETED_REVIEWS_ARRAY(assigned_2nd_application_date_const, review)
        ObjExcel.Cells(excel_row, REPT_PND2Days_col).value = COMPLETED_REVIEWS_ARRAY(assigned_rept_pnd2_days_const, review)
        ObjExcel.Cells(excel_row, QuestionableInterview_col).value = COMPLETED_REVIEWS_ARRAY(assigned_questionable_interview_const, review)
        ObjExcel.Cells(excel_row, Resolved_col).value = COMPLETED_REVIEWS_ARRAY(assigned_questionable_interview_resolve_const, review)
        ObjExcel.Cells(excel_row, ApptNoticeDate_col).value = COMPLETED_REVIEWS_ARRAY(assigned_appt_notice_date_const, review)
        ObjExcel.Cells(excel_row, ApptDate_col).value = COMPLETED_REVIEWS_ARRAY(assigned_appt_date_const, review)
        ObjExcel.Cells(excel_row, Confirmation_col).value = COMPLETED_REVIEWS_ARRAY(assigned_appt_notc_confirmation_const, review)
        ObjExcel.Cells(excel_row, NOMIDate_col).value = COMPLETED_REVIEWS_ARRAY(assigned_nomi_date_const, review)
        ObjExcel.Cells(excel_row, Confirmation2_col).value = COMPLETED_REVIEWS_ARRAY(assigned_nomi_confirmation_const, review)
        ObjExcel.Cells(excel_row, DenialNeeded_col).value = COMPLETED_REVIEWS_ARRAY(assigned_denial_needed_const, review)
        ObjExcel.Cells(excel_row, NextActionNeeded_col).value = COMPLETED_REVIEWS_ARRAY(assigned_next_action_needed_const, review)
        ObjExcel.Cells(excel_row, AddedtoWorkList_col).value = COMPLETED_REVIEWS_ARRAY(assigned_added_to_work_list_const, review)
        ObjExcel.Cells(excel_row, SecondApplicationDateNotes_col).value = COMPLETED_REVIEWS_ARRAY(assigned_2nd_application_date_resolve_const, review)
        ObjExcel.Cells(excel_row, ClosedInPast30Days_col).value = COMPLETED_REVIEWS_ARRAY(assigned_closed_recently_const, review)
        ObjExcel.Cells(excel_row, ClosedInPast30DaysNotes_col).value = COMPLETED_REVIEWS_ARRAY(assigned_closed_recently_resolve_const, review)
        ObjExcel.Cells(excel_row, StartedOutOfCounty_col).value = COMPLETED_REVIEWS_ARRAY(assigned_out_of_county_const, review)
        ObjExcel.Cells(excel_row, StartedOutOfCountyNotes_col).value = COMPLETED_REVIEWS_ARRAY(assigned_out_of_county_resolve_const, review)
        ObjExcel.Cells(excel_row, TrackingNotes_col).value = case_review_notes
        ObjExcel.Cells(excel_row, CaseSelectedTime_col).value = assigned_start_time
        ObjExcel.Cells(excel_row, Hold1Start_col).value = assigned_hold_1_start_time
        ObjExcel.Cells(excel_row, Hold1End_col).value = assigned_hold_1_end_time
        ObjExcel.Cells(excel_row, Hold2Start_col).value = assigned_hold_2_start_time
        ObjExcel.Cells(excel_row, Hold2End_col).value = assigned_hold_2_end_time
        ObjExcel.Cells(excel_row, Hold3Start_col).value = assigned_hold_3_start_time
        ObjExcel.Cells(excel_row, Hold3End_col).value = assigned_hold_3_end_time
        ObjExcel.Cells(excel_row, CaseCompletedTime_col).value = assigned_end_time

        excel_row = excel_row + 1

        If local_demo = False Then
            'Creating objects for Access
            Set objConnection = CreateObject("ADODB.Connection")
            Set objRecordSet = CreateObject("ADODB.Recordset")

            'This is the BZST connection to SQL Database'
            objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""

            'delete a record if the case number matches
            objRecordSet.Open "UPDATE ES.ES_OnDemanCashAndSnapBZProcessed SET TrackingNotes = '" & case_review_notes & "' WHERE CaseNumber = '" & MAXIS_case_number & "'", objConnection

            'close the connection and recordset objects to free up resources
            objConnection.Close
            Set objRecordSet=nothing
            Set objConnection=nothing
        End If

        With objFSO
            If .FileExists(od_revw_tracking_file_path) = True then
                .DeleteFile(od_revw_tracking_file_path)
            End If
        End With
    Next

    objWorkbook.Save()		'saving the excel
    ObjExcel.ActiveWorkbook.Close

    ObjExcel.Application.Quit
    ObjExcel.Quit

	main_email_subject = "On Demand Daily Case Review Completed"

    main_email_body = "The On Demand Appplication Reviews have been completed for " & date & "."
    main_email_body = main_email_body & vbCr & "Completed by: " & assigned_worker
    main_email_body = main_email_body & vbCr & vbCr & "Number of cases reviewed: " & UBound(COMPLETED_REVIEWS_ARRAY, 2)+1
    main_email_body = main_email_body & vbCr & "On Demand Case Reviews completed for the day."
    main_email_body = main_email_body & vbCr & vbCr & "Details of the reviews are logged here:"
    main_email_body = main_email_body & vbCr & "<" & file_url & ">"

    'KEEPING FOR FUTURE - there was more information in the reports for work assignment completed and we might want to add them back in
    ' main_email_body = main_email_body & vbCr & vbCr & "Work assignment assesment: " & assignment_assesment
    ' main_email_body = main_email_body & vbCr & vbCr & "Length of assignment: " & assignment_hours & " hours and " & assignment_minutes & " minutes."

    ' assignment_case_numbers_to_save = trim(assignment_case_numbers_to_save)
    ' assignment_new_ideas = trim(assignment_new_ideas)
    ' If assignment_case_numbers_to_save <> "" Then
    '     main_email_body = main_email_body & vbCr & vbCr & "Case numbers to discuss sent to QI email to be added to meeting agenda. Case numbers:"
    '     main_email_body = main_email_body & vbCr & assignment_case_numbers_to_save

    '     case_numbers_email_body = case_numbers_email_body & vbCr & assignment_case_numbers_to_save
    '     case_numbers_email_body = case_numbers_email_body & vbCr & vbCR & "These cases should be reviewed by the whole QI team and follow up decisions made."
    '     case_numbers_email_body = case_numbers_email_body & vbCr & vbCr & "------"
    '     case_numbers_email_body = case_numbers_email_body & vbCr & email_signature
    '     STATS_manualtime = STATS_manualtime + 120
    ' End If
    ' If assignment_new_ideas <> "" Then
    '     main_email_body = main_email_body & vbCr & vbCr & "New ideas for statistics to gather sent to the BZST. Ideas:"
    '     main_email_body = main_email_body & vbCr & assignment_new_ideas

    '     ideas_email_body = ideas_email_body & vbCr & assignment_new_ideas
    '     ideas_email_body = ideas_email_body & vbCr & vbCr & "------"
    '     ideas_email_body = ideas_email_body & vbCr & email_signature
    '     STATS_manualtime = STATS_manualtime + 120
    ' End If
    main_email_body = main_email_body & vbCr & vbCr & "------"
    main_email_body = main_email_body & vbCr & qi_worker_first_name

    CALL create_outlook_email(qi_worker_supervisor_email, "HSPH.EWS.BlueZoneScripts@hennepin.us", main_email_subject, main_email_body, "", TRUE)
end function

function create_tracking_cookie()
    With objFSO
        Dim objTextStream
        Set objTextStream = .OpenTextFile(od_revw_tracking_file_path, ForWriting, true)

        objTextStream.WriteLine "CASE NUMBER ^*^*^" & MAXIS_case_number
        objTextStream.WriteLine "ASSIGNED WORKER ^*^*^" & assigned_worker
        objTextStream.WriteLine "WINDOWS USER ID ^*^*^" & user_ID_for_validation
        objTextStream.WriteLine "ASSIGNED DATE ^*^*^" & assigned_date
        objTextStream.WriteLine "START TIME ^*^*^" & assigned_start_time

        'Close the object so it can be opened again shortly
        objTextStream.Close
    End With
end function

function merge_worklist_to_SQL()
    'setting up information and variables for accessing yesterday's worklist
    previous_date = dateadd("d", -1, date)
    Call change_date_to_soonest_working_day(previous_date, "back")       'finds the most recent previous working day
    archive_folder = DatePart("yyyy", previous_date) & "-" & right("0" & DatePart("m", previous_date), 2)

    previous_date_month = DatePart("m", previous_date)
    previous_date_day = DatePart("d", previous_date)
    previous_date_year = DatePart("yyyy", previous_date)
    previous_date_header = previous_date_month & "-" & previous_date_day & "-" & previous_date_year

    'setting up file paths for accessing yesterday's worklist
    archive_files = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\On Demand Waiver\QI On Demand Daily Assignment\Archive\" & archive_folder

    previous_list_file_selection_path = t_drive & "/Eligibility Support/Restricted/QI - Quality Improvement/REPORTS/On Demand Waiver/QI On Demand Daily Assignment/QI " & previous_date_header & " Worklist.xlsx"
    Call File_Exists(previous_list_file_selection_path, does_file_exist)
    previous_worksheet_header = "Work List for " & previous_date_month & "-" & previous_date_day & "-" & previous_date_year

	yesterday_case_list = 0

	If does_file_exist = True Then
		'open the file
		call excel_open(previous_list_file_selection_path, True, False, ObjYestExcel, objYestWorkbook)

		objYestWorkbook.worksheets(previous_worksheet_header).Activate

        'Pull info into a NEW array of prevvious day work.
		xl_row = 2
		Do
			this_case = trim(ObjYestExcel.Cells(xl_row, 2).Value)
			If this_case <> "" Then
                worklist_notes = ""
                yesterday_list_case_number = trim(ObjYestExcel.Cells(xl_row, 2).Value)
                yesterday_notes = replace(ObjYestExcel.Cells(xl_row, 23).Value, "FOLLOW UP NEEDED", "")
                yesterday_notes = replace(yesterday_notes, "  ", " ")
                yesterday_notes = replace(yesterday_notes, "'", "")
                yesterday_notes = trim(yesterday_notes)
                yesterday_action = ObjYestExcel.Cells(xl_row, 22).Value
                yesterday_action = trim(yesterday_action)
                If yesterday_action = "FOLLOW UP NEEDED" Then worklist_notes = trim(yesterday_action) & " - " & yesterday_notes
                If yesterday_action <> "FOLLOW UP NEEDED" Then worklist_notes = yesterday_notes

                'Creating objects for Access
                Set objConnection = CreateObject("ADODB.Connection")
                Set objRecordSet = CreateObject("ADODB.Recordset")

                'This is the BZST connection to SQL Database'
                objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""

                'delete a record if the case number matches
                objRecordSet.Open "UPDATE ES.ES_OnDemanCashAndSnapBZProcessed SET TrackingNotes = '" & worklist_notes & "' WHERE CaseNumber = '" & yesterday_list_case_number & "'", objConnection

                'close the connection and recordset objects to free up resources
                objConnection.Close
                Set objRecordSet=nothing
                Set objConnection=nothing
			End If
            xl_row = xl_row + 1
		Loop until this_case = ""

		'close the file
		ObjYestExcel.ActiveWorkbook.Close
		ObjYestExcel.Application.Quit
		ObjYestExcel.Quit
	End If
    Call script_end_procedure("Clean Up Completed")
end function

function read_tracking_cookie()
    With objFSO
        'Creating an object for the stream of text which we'll use frequently
        If .FileExists(od_revw_tracking_file_path) = True then

            Set objTextStream = .OpenTextFile(od_revw_tracking_file_path, ForReading)

            'Reading the entire text file into a string
            every_line_in_text_file = objTextStream.ReadAll

            'Splitting the text file contents into an array which will be sorted
            od_revw_case_details = split(every_line_in_text_file, vbNewLine)

            For Each text_line in od_revw_case_details										'read each line in the file
                If Instr(text_line, "^*^*^") <> 0 Then
                    line_info = split(text_line, "^*^*^")								'creating a small array for each line. 0 has the header and 1 has the information
                    line_info(0) = trim(line_info(0))
                    'here we add the information from TXT to Excel

                    If line_info(0) = "CASE NUMBER" Then MAXIS_case_number  = line_info(1)
                    If line_info(0) = "ASSIGNED WORKER" Then cookie_worker  = line_info(1)
                    If line_info(0) = "WINDOWS USER ID" Then user_ID_for_validation  = line_info(1)
                    If line_info(0) = "ASSIGNED DATE" Then assigned_date  = line_info(1)
                    If line_info(0) = "START TIME" Then assigned_start_time  = line_info(1)
                    If line_info(0) = "END TIME" Then assigned_end_time  = line_info(1)
                    If line_info(0) = "HOLD 1 START" Then assigned_hold_1_start_time  = line_info(1)
                    If line_info(0) = "HOLD 1 END" Then assigned_hold_1_end_time  = line_info(1)
                    If line_info(0) = "HOLD 2 START" Then assigned_hold_2_start_time  = line_info(1)
                    If line_info(0) = "HOLD 2 END" Then assigned_hold_2_end_time  = line_info(1)
                    If line_info(0) = "HOLD 3 START" Then assigned_hold_3_start_time  = line_info(1)
                    If line_info(0) = "HOLD 3 END" Then assigned_hold_3_end_time  = line_info(1)
                End If
            Next

            objTextStream.Close
        End If
    End With
end function

function set_variables_from_SQL()
    MAXIS_case_number                       = objRecordSet("CaseNumber")
    assigned_case_name                      = objRecordSet("CaseName")
    table_application_date                  = objRecordSet("ApplDate")
    table_interview_date                    = objRecordSet("InterviewDate")
    table_day_30                            = objRecordSet("Day_30")
    assigned_days_pending                   = objRecordSet("DaysPending")
    assigned_snap_status                    = objRecordSet("SnapStatus")
    assigned_cash_status                    = objRecordSet("CashStatus")
    table_2nd_application_date              = objRecordSet("SecondApplicationDate")
    assigned_rept_pnd2_days                 = objRecordSet("REPT_PND2Days")
    assigned_questionable_interview         = objRecordSet("QuestionableInterview")
    assigned_questionable_interview_resolve = objRecordSet("Resolved")
    table_appt_notice_date                  = objRecordSet("ApptNoticeDate")
    table_appt_date                         = objRecordSet("ApptDate")
    assigned_appt_notc_confirmation         = objRecordSet("Confirmation")
    table_nomi_date                         = objRecordSet("NOMIDate")
    assigned_nomi_confirmation              = objRecordSet("Confirmation2")
    assigned_denial_needed                  = objRecordSet("DenialNeeded")
    assigned_next_action_needed             = objRecordSet("NextActionNeeded")
    table_added_to_work_list                = objRecordSet("AddedtoWorkList")
    assigned_2nd_application_date_resolve   = objRecordSet("SecondApplicationDateNotes")
    assigned_closed_recently                = objRecordSet("ClosedInPast30Days")
    assigned_closed_recently_resolve        = objRecordSet("ClosedInPast30DaysNotes")
    assigned_out_of_county                  = objRecordSet("StartedOutOfCounty")
    assigned_out_of_county_resolve          = objRecordSet("StartedOutOfCountyNotes")
    assigned_tracking_notes                 = objRecordSet("TrackingNotes")

    case_review_notes = replace(assigned_tracking_notes, "STS-NR", "")
    case_review_notes = replace(case_review_notes, "STS-RC-"&user_ID_for_validation, "")
    case_review_notes = replace(case_review_notes, "STS-IP-"&user_ID_for_validation, "")
    case_review_notes = replace(case_review_notes, "STS-HD-"&user_ID_for_validation, "")
    case_review_notes = replace(case_review_notes, "STS-RC", "")
    case_review_notes = replace(case_review_notes, "STS-NL", "")
    case_review_notes = trim(case_review_notes)

    date_zero =  #1/1/1900#
    If IsDate(table_application_date) = True Then
        assigned_application_date = FormatDateTime(table_application_date, 2)
        If DateDiff("d", assigned_application_date, date_zero) = 0 Then assigned_application_date = ""
    End If
    If IsDate(table_day_30) = True Then
        assigned_day_30 = FormatDateTime(table_day_30, 2)
        If DateDiff("d", assigned_day_30, date_zero) = 0 Then assigned_day_30 = ""
    End If

    If IsDate(table_interview_date) = True Then
        assigned_interview_date = FormatDateTime(table_interview_date, 2)
        If DateDiff("d", assigned_interview_date, date_zero) = 0 Then assigned_interview_date = ""
    End If
    If IsDate(table_2nd_application_date) = True Then
        assigned_2nd_application_date = FormatDateTime(table_2nd_application_date, 2)
        If DateDiff("d", assigned_2nd_application_date, date_zero) = 0 Then assigned_2nd_application_date = ""
    End If
    If IsDate(table_appt_notice_date) = True Then
        assigned_appt_notice_date = FormatDateTime(table_appt_notice_date, 2)
        If DateDiff("d", assigned_appt_notice_date, date_zero) = 0 Then assigned_appt_notice_date = ""
    End If
    If IsDate(table_appt_date) = True Then
        assigned_appt_date = FormatDateTime(table_appt_date, 2)
        If DateDiff("d", assigned_appt_date, date_zero) = 0 Then assigned_appt_date = ""
    End If
    If IsDate(table_nomi_date) = True Then
        assigned_nomi_date = FormatDateTime(table_nomi_date, 2)
        If DateDiff("d", assigned_nomi_date, date_zero) = 0 Then assigned_nomi_date = ""
    End If
    If IsDate(table_added_to_work_list) = True Then
        assigned_added_to_work_list = FormatDateTime(table_added_to_work_list, 2)
        If DateDiff("d", assigned_added_to_work_list, date_zero) = 0 Then assigned_added_to_work_list = ""
    End If
end function

function test_sql_access()
    'Access the pending cases TABLE - ES_OnDemandCashAndSnap'
    'declare the SQL statement that will query the database
    objSQL = "SELECT * FROM ES.ES_OnDemandCashAndSnap"

    'Creating objects for Access
    Set objConnection = CreateObject("ADODB.Connection")
    Set objRecordSet = CreateObject("ADODB.Recordset")

    'This is the file path for the statistics Access database.
    objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
    objRecordSet.Open objSQL, objConnection

    Do While NOT objRecordSet.Eof
		anything_number = objRecordSet("CaseNumber")
		case_basket = objRecordSet("WorkerID")
        objRecordSet.MoveNext
    Loop

    'close the connection and recordset objects to free up resources
    objRecordSet.Close
    objConnection.Close
    Set objRecordSet=nothing
    Set objConnection=nothing
    Set objSQL=nothing

    'Access the working list TABLE - ES_OnDemanCashAndSnapBZProcessed'
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

    objWorkRecordSet.Close
    objWorkConnection.Close

    Set objWorkRecordSet=nothing
    Set objWorkConnection=nothing
    Set objWorkSQL=nothing

    email_subject = "ON Demand Access Test Completed for " & assigned_worker
    email_body = "VARIABLE DETAILS: " & vbCr

    email_body = email_body & "anything_number - " & anything_number & vbCr
    email_body = email_body & "case_basket - " & case_basket & vbCr
    email_body = email_body & "first_item_change - " & first_item_change & vbCr
    email_body = email_body & "first_item_date - " & first_item_date & vbCr
    email_body = email_body & "This test is complete and this worker has read access."

    Call create_outlook_email("hsph.ews.bluezonescripts@hennepin.us", "", email_subject, email_body, "", True)

    end_msg = "The test is complete an an email has been sent to the BlueZone Script Team regarding access to the tables for the new On Demand functionality."
    Call script_end_procedure(end_msg)
end function

function update_tracking_cookie(update_reason)
    With objFSO
        'If the file doesn't exist, it needs to create it here and initialize it here! After this, it can just exit as the file will now be initialized
        If .FileExists(od_revw_tracking_file_path) = True then
            .DeleteFile(od_revw_tracking_file_path)
        End If

        If .FileExists(od_revw_tracking_file_path) = False then
            Dim objTextStream
            Set objTextStream = .OpenTextFile(od_revw_tracking_file_path, ForWriting, true)

            ' objTextStream.WriteLine ""
            objTextStream.WriteLine "CASE NUMBER ^*^*^" & MAXIS_case_number
            objTextStream.WriteLine "ASSIGNED WORKER ^*^*^" & assigned_worker
            objTextStream.WriteLine "WINDOWS USER ID ^*^*^" & user_ID_for_validation
            objTextStream.WriteLine "ASSIGNED DATE ^*^*^" & assigned_date
            objTextStream.WriteLine "START TIME ^*^*^" & assigned_start_time
            objTextStream.WriteLine "END TIME ^*^*^" & assigned_end_time
            objTextStream.WriteLine "HOLD 1 START ^*^*^" & assigned_hold_1_start_time
            objTextStream.WriteLine "HOLD 1 END ^*^*^" & assigned_hold_1_end_time
            objTextStream.WriteLine "HOLD 2 START ^*^*^" & assigned_hold_2_start_time
            objTextStream.WriteLine "HOLD 2 END ^*^*^" & assigned_hold_2_end_time
            objTextStream.WriteLine "HOLD 3 START ^*^*^" & assigned_hold_3_start_time
            objTextStream.WriteLine "HOLD 3 END ^*^*^" & assigned_hold_3_end_time

            If update_reason = "HOLD" Then
                objTextStream.WriteLine "CASE NAME ^*^*^" & assigned_case_name
                objTextStream.WriteLine "APPLICATION DATE ^*^*^" & assigned_application_date
                objTextStream.WriteLine "INTERVIEW DATE ^*^*^" & assigned_interview_date
                objTextStream.WriteLine "DAY 30 ^*^*^" & assigned_day_30
                objTextStream.WriteLine "DAYS PENDING ^*^*^" & assigned_days_pending
                objTextStream.WriteLine "SNAP STATUS ^*^*^" & assigned_snap_status
                objTextStream.WriteLine "CASH STATUS ^*^*^" & assigned_cash_status
                objTextStream.WriteLine "SUBSEQUENT APP DATE ^*^*^" & assigned_2nd_application_date
                objTextStream.WriteLine "REPT PND2 DAYS ^*^*^" & assigned_rept_pnd2_days
                objTextStream.WriteLine "QUESTIONABLE INTERVIEW ^*^*^" & assigned_questionable_interview
                objTextStream.WriteLine "QUESTIONABLE INTERVIEW RESOLVE ^*^*^" & assigned_questionable_interview_resolve
                objTextStream.WriteLine "APPT NOTC DATE ^*^*^" & assigned_appt_notice_date
                objTextStream.WriteLine "APPT DATE ^*^*^" & assigned_appt_date
                objTextStream.WriteLine "APPT NOTC CONFIRMATION ^*^*^" & assigned_appt_notc_confirmation
                objTextStream.WriteLine "NOMI DATE ^*^*^" & assigned_nomi_date
                objTextStream.WriteLine "NOMI CONFIRMATION ^*^*^" & assigned_nomi_confirmation
                objTextStream.WriteLine "DENIAL NEEDED ^*^*^" & assigned_denial_needed
                objTextStream.WriteLine "NEXT ACTION NEEDED ^*^*^" & assigned_next_action_needed
                objTextStream.WriteLine "ADDED TO WORKLIST ^*^*^" & assigned_added_to_work_list
                objTextStream.WriteLine "SUBSEQUENT APP DATE RESOLVE ^*^*^" & assigned_2nd_application_date_resolve
                objTextStream.WriteLine "CLOSED RECENTLY ^*^*^" & assigned_closed_recently
                objTextStream.WriteLine "CLOSED RECENTLY RESOLVE ^*^*^" & assigned_closed_recently_resolve
                objTextStream.WriteLine "OUT OF COUNTY ^*^*^" & assigned_out_of_county
                objTextStream.WriteLine "OUT OF COUNTY RESOLVE ^*^*^" & assigned_out_of_county_resolve
                objTextStream.WriteLine "CASE REVIEW NOTES ^*^*^" & case_review_notes
            End If

            'Close the object so it can be opened again shortly
            objTextStream.Close
        End If
    End With
end function

'END FUNCTIONS =============================================================================================================

EMConnect ""                'connecting to MAXIS
Call check_for_MAXIS(True)  'If we are not in MAXIS or not passworded into MAXIS the script will end.

'If a BZST worker is running this script, there is functionality for running a cleanup or running in DEMMO mode
If user_ID_for_validation = "CALO001" or user_ID_for_validation = "ILFE001" Then
    Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 271, 70, "BZST ScriptWriter Options"                     'dialog to select demo or clean up options
        ButtonGroup ButtonPressed
            OkButton 215, 50, 50, 15
        Text 10, 15, 195, 10, "Do you need to run the On Demand Dashboard in DEMO?"
        DropListBox 205, 10, 60, 45, "No"+chr(9)+"Yes", run_demo
        Text 90, 35, 115, 10, "Do you need to clean up the list?"
        DropListBox 205, 30, 60, 45, "No"+chr(9)+"Yes", clean_up_list
    EndDialog

    dialog Dialog1

    If run_demo = "Yes" Then local_demo = True
    If clean_up_list = "Yes" Then Call merge_worklist_to_SQL

    ADMIN_run = True
End If
'this defines workers that have access to the Admmin functions along with the BZST writers
If user_ID_for_validation = "TAPA002" or user_ID_for_validation = "WFX901" or user_ID_for_validation = "WFU851" Then ADMIN_run = True   'TP, FRC, JF

'the scripts should have loaded the tester array from GlobVar but if it did not, this will load it
If IsArray(tester_array) = False Then
    Dim tester_array()
    ReDim tester_array(0)
    tester_list_URL = "\\hcgg.fr.co.hennepin.mn.us\lobroot\hsph\team\Eligibility Support\Scripts\Script Files\COMPLETE LIST OF TESTERS.vbs"        'Opening the list of testers - which is saved locally for security
    Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
    Set fso_command = run_another_script_fso.OpenTextFile(tester_list_URL)
    text_from_the_other_script = fso_command.ReadAll
    fso_command.Close
    Execute text_from_the_other_script
End If

'confirm QI Member
'This script can only be run by members of QI for access reasons. The script is held in the QI section but this is a double check.
qi_member_identified = False
For tester = 0 to UBound(tester_array)                         'looping through all of the testers
    ' pulling QI members by supervisor from the Complete List of Testers
    If tester_array(tester).tester_id_number = user_ID_for_validation Then
        qi_worker_supervisor_email = tester_array(tester).tester_supervisor_email
        qi_worker_first_name = tester_array(tester).tester_first_name
        If tester_array(tester).tester_supervisor_name = "Tanya Payne" Then qi_member_identified = True
        If tester_array(tester).tester_population = "BZ" Then qi_member_identified = True
        assigned_worker = tester_array(tester).tester_full_name
    End If
Next
'cancelling the script run if a QI member does not run this script
If qi_member_identified = False Then script_end_procedure("This script can only be operated by a member of core QI due to access restrictions. The script will now end.")

If local_demo = True then end_msg = end_msg & vbCr & "This script run was completed as a DEMO."     'setting the end mmessage to indicate that it was run as a demo

'if not a DEMO the script will read informmation fromm SQL to determmine the functionality of the dashboard.
If local_demo = False Then
    'Access the SQL Table
    'declare the SQL statement that will query the database
    objSQL = "SELECT * FROM ES.ES_OnDemanCashAndSnapBZProcessed"

    'Creating objects for Access
    Set objConnection = CreateObject("ADODB.Connection")
    Set objRecordSet = CreateObject("ADODB.Recordset")

    'This is the file path for the statistics Access database.
    objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
    objRecordSet.Open objSQL, objConnection

    'Determine the LOAD Date to identify if the BULK Run was completed
    first_item_change = objRecordSet("AuditLoadDate")
    first_item_array = split(first_item_change, " ")
    first_item_date = first_item_array(0)
    first_item_time = first_item_array(1)
    first_item_time_array = split(first_item_change, ".")
    first_item_time = first_item_time_array(0)
    first_item_date = DateAdd("d", 0, first_item_date)
    If DateDiff("d", first_item_date, date) = 0 Then BULK_Run_completed = True
    first_item_time = FormatDateTime(first_item_time, 3)

    cases_on_hold_count = 0
    If BULK_Run_completed = True Then
        'Read the whole table
        Do While NOT objRecordSet.Eof
            'count all of today's cases using added to worklist
            case_worklist_date = objRecordSet("AddedtoWorkList")
            case_worklist_date = DateAdd("d", 0, case_worklist_date)
            If DateDiff("d", case_worklist_date, date) = 0 Then total_cases_for_review = total_cases_for_review + 1

            case_tracking_notes = objRecordSet("TrackingNotes")

            'count completed reviews using info in tracking notes
            If Instr(case_tracking_notes, "STS-RC") <> 0 Then cases_with_review_completed =cases_with_review_completed + 1
            If DateDiff("d", case_worklist_date, date) = 0 AND Instr(case_tracking_notes, "STS") = 0 Then cases_with_review_completed =cases_with_review_completed + 1

            'count waiting using info in tracking notes
            If Instr(case_tracking_notes, "STS-NR") <> 0 Then cases_waiting_for_review =cases_waiting_for_review + 1

            'count cases on hold
            If Instr(case_tracking_notes, "STS-HD") Then
                If Instr(case_tracking_notes, user_ID_for_validation) <> 0 Then
                    cases_on_hold = cases_on_hold + 1
                    If Instr(case_tracking_notes, user_ID_for_validation) <> 0 Then
                        ReDim preserve CASES_ON_HOLD_ARRAY(last_const, cases_on_hold_count)
                        CASES_ON_HOLD_ARRAY(case_nbr_const, cases_on_hold_count) = objRecordSet("CaseNumber")
                        CASES_ON_HOLD_ARRAY(case_notes_const, cases_on_hold_count) = objRecordSet("TrackingNotes")
                        CASES_ON_HOLD_ARRAY(case_notes_const, cases_on_hold_count) = replace(CASES_ON_HOLD_ARRAY(case_notes_const, cases_on_hold_count), "STS-RC", "")
                        CASES_ON_HOLD_ARRAY(case_notes_const, cases_on_hold_count) = replace(CASES_ON_HOLD_ARRAY(case_notes_const, cases_on_hold_count), "STS-RC-"&user_ID_for_validation, "")
                        CASES_ON_HOLD_ARRAY(case_notes_const, cases_on_hold_count) = replace(CASES_ON_HOLD_ARRAY(case_notes_const, cases_on_hold_count), "STS-IP-"&user_ID_for_validation, "")
                        CASES_ON_HOLD_ARRAY(case_notes_const, cases_on_hold_count) = replace(CASES_ON_HOLD_ARRAY(case_notes_const, cases_on_hold_count), "STS-HD-"&user_ID_for_validation, "")
                        CASES_ON_HOLD_ARRAY(case_notes_const, cases_on_hold_count) = trim(CASES_ON_HOLD_ARRAY(case_notes_const, cases_on_hold_count))
                        cases_on_hold_count = cases_on_hold_count + 1
                    End If
                End If
            End If

            'find if there is a case 'checked out'
            If Instr(case_tracking_notes, "STS-IP") Then
                If Instr(case_tracking_notes, user_ID_for_validation) <> 0 Then
                    worker_on_task = True
                    case_nbr_in_progress = objRecordSet("CaseNumber")
                    set_variables_from_SQL
                    'NOTE - if another worker has a case in progress, it will not show for the current worker, but it will show in the ADMIN function
                End If
            End If

            'If we are running as an ADMIN, the script needs to collect some additional information about the status of the cases.
            If ADMIN_run = True Then
                RC_id_start = ""
                HD_id_start = ""
                IP_id_start = ""
                worker_id_RC = ""
                worker_id_HD = ""
                worker_id_IP = ""
                If Instr(case_tracking_notes, "STS-RC") <> 0 Then
                    admin_count_RC = admin_count_RC + 1
                    RC_id_start = Instr(case_tracking_notes, "STS-RC") + 7
                    worker_id_RC = MID(case_tracking_notes, RC_id_start, 7)
					worker_id_RC = trim(worker_id_RC)
                    If InStr(ADMIN_list_workers_RC, "~" & worker_id_RC & "~") = 0 then ADMIN_list_workers_RC = ADMIN_list_workers_RC & worker_id_RC & "~"
                End If
                If Instr(case_tracking_notes, "STS-NR") <> 0 Then admin_count_NR = admin_count_NR + 1
                If Instr(case_tracking_notes, "STS-HD") <> 0 Then
                    admin_count_HD = admin_count_HD + 1
                    HD_id_start = Instr(case_tracking_notes, "STS-HD") + 7
                    worker_id_HD = MID(case_tracking_notes, HD_id_start, 7)
					worker_id_HD = trim(worker_id_HD)
                    If InStr(ADMIN_list_workers_HD, "~" & worker_id_HD & "~") = 0 then ADMIN_list_workers_HD = ADMIN_list_workers_HD & worker_id_HD & "~"
                End If
                If Instr(case_tracking_notes, "STS-IP") <> 0 Then
                    admin_count_IP = admin_count_IP + 1
                    IP_id_start = Instr(case_tracking_notes, "STS-IP") + 7
                    worker_id_IP = MID(case_tracking_notes, IP_id_start, 7)
					worker_id_IP = trim(worker_id_IP)
                    If InStr(ADMIN_list_workers_IP, "~" & worker_id_IP & "~") = 0 then ADMIN_list_workers_IP = ADMIN_list_workers_IP & worker_id_IP & "~"
                End If
            End If
            objRecordSet.MoveNext
        Loop
    End If

    'close the connection and recordset objects to free up resources
    objRecordSet.Close
    objConnection.Close
    Set objRecordSet=nothing
    Set objConnection=nothing

    'If we are running ADMIN functionality we need to capture additional details from information read from case statuses.
    worker_count = 0
    RCcount = 0
    IPCount = 0
    HDCount = 0
    If ADMIN_run = True Then
        If len(ADMIN_list_workers_RC) > 1 Then
            ADMIN_list_workers_RC = right(ADMIN_list_workers_RC, len(ADMIN_list_workers_RC)-1)
            ADMIN_list_workers_RC = left(ADMIN_list_workers_RC, len(ADMIN_list_workers_RC)-1)
            If InStr(ADMIN_list_workers_RC, "~") <> 0 Then RC_id_ARRAY = split(ADMIN_list_workers_RC, "~")
            If InStr(ADMIN_list_workers_RC, "~") = 0 Then RC_id_ARRAY = array(ADMIN_list_workers_RC)
            For each RC_worker_Id in RC_id_ARRAY
                For tester = 0 to UBound(tester_array)                         'looping through all of the testers
                    ' pulling QI members by supervisor from the Complete List of Testers
                    If tester_array(tester).tester_id_number = RC_worker_Id Then
                        ReDim Preserve ADMIN_worker_list_array(admin_wrkr_last_const, worker_count)

                        ADMIN_worker_list_array(wrkr_id_const, worker_count) = RC_worker_Id
                        ADMIN_worker_list_array(wrkr_name_const, worker_count) = tester_array(tester).tester_first_name
                        ADMIN_worker_list_array(case_status_const, worker_count) = "RC"
                        ADMIN_worker_list_array(admin_radio_btn_const, worker_count) = RCcount

                        worker_count = worker_count + 1
                        RCcount = RCcount + 1
                    End If
                Next
            Next
        End If
        If len(ADMIN_list_workers_HD) > 1 Then
            ADMIN_list_workers_HD = right(ADMIN_list_workers_HD, len(ADMIN_list_workers_HD)-1)
            ADMIN_list_workers_HD = left(ADMIN_list_workers_HD, len(ADMIN_list_workers_HD)-1)
            If InStr(ADMIN_list_workers_HD, "~") <> 0 Then HD_id_ARRAY = split(ADMIN_list_workers_HD, "~")
            If InStr(ADMIN_list_workers_HD, "~") = 0 Then HD_id_ARRAY = array(ADMIN_list_workers_HD)
            For each HD_worker_Id in HD_id_ARRAY
                For tester = 0 to UBound(tester_array)                         'looping through all of the testers
                    ' pulling QI members by supervisor from the Complete List of Testers
                    If tester_array(tester).tester_id_number = HD_worker_Id Then
                        ReDim Preserve ADMIN_worker_list_array(admin_wrkr_last_const, worker_count)

                        ADMIN_worker_list_array(wrkr_id_const, worker_count) = HD_worker_Id
                        ADMIN_worker_list_array(wrkr_name_const, worker_count) = tester_array(tester).tester_first_name
                        ADMIN_worker_list_array(case_status_const, worker_count) = "HD"
                        ADMIN_worker_list_array(admin_radio_btn_const, worker_count) = HDCount

                        worker_count = worker_count + 1
                        HDCount = HDCount + 1
                    End If
                Next
            Next
        End If
        If len(ADMIN_list_workers_IP) > 1 Then
            ADMIN_list_workers_IP = right(ADMIN_list_workers_IP, len(ADMIN_list_workers_IP)-1)
            ADMIN_list_workers_IP = left(ADMIN_list_workers_IP, len(ADMIN_list_workers_IP)-1)
            If InStr(ADMIN_list_workers_IP, "~") <> 0 Then IP_id_ARRAY = split(ADMIN_list_workers_IP, "~")
            If InStr(ADMIN_list_workers_IP, "~") = 0 Then IP_id_ARRAY = array(ADMIN_list_workers_IP)
            For each IP_worker_Id in IP_id_ARRAY
                For tester = 0 to UBound(tester_array)                         'looping through all of the testers
                    ' pulling QI members by supervisor from the Complete List of Testers
                    If tester_array(tester).tester_id_number = IP_worker_Id Then
                        ReDim Preserve ADMIN_worker_list_array(admin_wrkr_last_const, worker_count)

                        ADMIN_worker_list_array(wrkr_id_const, worker_count) = IP_worker_Id
                        ADMIN_worker_list_array(wrkr_name_const, worker_count) = tester_array(tester).tester_first_name
                        ADMIN_worker_list_array(case_status_const, worker_count) = "IP"
                        ADMIN_worker_list_array(admin_radio_btn_const, worker_count) = IPCount

                        worker_count = worker_count + 1
                        IPCount = IPCount + 1
                    End If
                Next
            Next
        End If
    End If
Else                            'if we are running in DEMO mode, we don't read the table - we have a dialog to select the process to view.
    Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 191, 85, "On Demand Demo"
      DropListBox 10, 25, 175, 45, "Select One..."+chr(9)+"Complete BULK Run"+chr(9)+"Select New Case"+chr(9)+"Case in Progress", demo_process
      EditBox 85, 45, 25, 15, cases_on_hold
      EditBox 85, 65, 25, 15, reviews_complete
      ButtonGroup ButtonPressed
        OkButton 135, 45, 50, 15
        CancelButton 135, 65, 50, 15
      Text 10, 15, 65, 10, "Process to Demo"
      Text 15, 50, 55, 10, "Cases on Hold"
      Text 15, 70, 65, 10, "Reviews Complete"
      Text 90, 5, 95, 10, "Total Cases to Review: 124"
    EndDialog

    dialog Dialog1
    cancel_without_confirmation

    If trim(cases_on_hold) = "" Then cases_on_hold = 0                                  'formatting some of the information from this dialog.
    If trim(reviews_complete) = "" Then reviews_complete = 0

    cases_with_review_completed = reviews_complete * 1
    cases_on_hold = cases_on_hold * 1

    'now we have to make some case numbers to list fake cases on hold.
    If cases_on_hold <> 0 Then
        case_numb = 202151
        For case_hold_count = 1 to cases_on_hold
            ReDim preserve CASES_ON_HOLD_ARRAY(last_const, case_hold_count-1)
            CASES_ON_HOLD_ARRAY(case_nbr_const, case_hold_count-1) = case_numb
            CASES_ON_HOLD_ARRAY(case_notes_const, case_hold_count-1) = "Info here."
            randmon_nbr = Rnd
            randmon_nbr = randmon_nbr * 12300
            randmon_nbr = Int(randmon_nbr)
            case_numb = case_numb + randmon_nbr
        Next
    End If

    cases_waiting_for_review = total_cases_for_review - cases_with_review_completed         'doing some math

    'setting the booleans for the rest of the script run.
    If demo_process = "Complete BULK Run" Then
        BULK_Run_completed = False
        worker_on_task = False
        end_msg = end_msg & vbcr & "BULK run completion mock up completed."
    ElseIf demo_process = "Select New Case" Then
        BULK_Run_completed = True
        worker_on_task = False
        end_msg = end_msg & vbcr & "Selecting a case mock up completed."
    ElseIf demo_process = "Case in Progress" Then
        BULK_Run_completed = True
        worker_on_task = True
        end_msg = end_msg & vbcr & "Displaying a case in progress mock up completed."
    End If
End If

'Here is where the script will decide which dialog to display in the process step for the day.
If BULK_Run_completed = False Then                  'if the main run has not happened yet, we start here.
    Do
        Do
            Dialog1 = ""
            BeginDialog Dialog1, 0, 0, 451, 155, "On Demand Applications Dashboard"
                EditBox 500, 300, 50, 15, fake_edit_box
                ButtonGroup ButtonPressed
                PushButton 310, 55, 125, 15, "Start On Demand BULK Run", complete_bulk_run_btn
                PushButton 50, 105, 170, 10, "More information about the BULK Run", bulk_run_details_btn
                PushButton 375, 5, 65, 15, "Test Access", test_access_btn
                If ADMIN_run = True Then PushButton 10, 135, 70, 15, "Admin Functions", admin_btn
                ' OkButton 335, 130, 50, 15
                CancelButton 390, 130, 50, 15
                Text 170, 10, 135, 10, "On Demand Applications Dashboard"
                GroupBox 10, 25, 430, 100, "Applications BULK Run"
                Text 20, 40, 170, 10, "The BULK run was last completed on " & first_item_date & "."
                Text 20, 55, 285, 20, "The BULK Run has not been completed today. You must complete the BULK run before any other On Demand work can be completed."
                Text 45, 75, 235, 10, "- The BULK run takes 2 - 3 hour and uses PRODUCTION the entire time."
                Text 45, 85, 260, 10, "- You can complete other work in other sessions while the BULK Run happens."
                Text 45, 95, 325, 10, "- The BULK Run can be unattended (you can walk away) but this is not paid time if you walk away."
            EndDialog

            dialog Dialog1
            cancel_without_confirmation

            'each button takes us to a different process step
            If ButtonPressed = bulk_run_details_btn Then MsgBox "More details will be here" 'TODO - add BULK Run Explanation'
            If ButtonPressed = test_access_btn Then Call test_sql_access()      'quick access functionality to make sure the worker has the correct permissions
            If ButtonPressed = complete_bulk_run_btn Then                       'if it is selected to start the BULK run, this will check for production and open the On Demand Applications script.
                If local_demo = True Then call script_end_procedure("The script would now run the the On Demand Applications script")
                If local_demo = False Then
                    Call back_to_SELF
                    EMReadScreen MX_region, 10, 22, 48
                    MX_region = trim(MX_region)
                    If MX_region <> "PRODUCTION" Then Call script_end_procedure("You have selected to complete the BULK Run for On Demand but you are not in production. The script will now end. Move to PRODUCTION and run On Demand Dashboard again.")
                    Call run_from_GitHub(script_repository & "admin\" & "on-demand-waiver-applications.vbs")
                End If
            End If
            If ButtonPressed = admin_btn Then call complete_admin_functions      'this opens the admin information and details
        Loop
        CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    Loop until are_we_passworded_out = false					'loops until user passwords back in
End If

If worker_on_task = False Then
    If cases_on_hold = 0 Then
        Do
            Do
                Dialog1 = ""
                BeginDialog Dialog1, 0, 0, 451, 235, "On Demand Applications Dashboard"
                  EditBox 500, 600, 50, 15, fake_edit_box
                  If total_cases_for_review = cases_with_review_completed Then Text 20, 170, 155, 10, "*** All Cases have been Pulled for Review ***'"
                  ButtonGroup ButtonPressed
                    If total_cases_for_review <> cases_with_review_completed Then PushButton 15, 165, 110, 15, "Pull a case to Review", get_new_case_btn
                    PushButton 285, 70, 140, 10, "More information about the Work List", work_list_details_btn
                    PushButton 10, 215, 105, 15, "Finish Work Day", finish_work_day_btn
                    If ADMIN_run = True Then PushButton 120, 215, 70, 15, "Admin Functions", admin_btn
                    PushButton 375, 20, 65, 15, "Test Access", test_access_btn
                    PushButton 330, 85, 105, 15, "Restart the BULK Run", bulk_run_incomplete_btn
                    CancelButton 390, 215, 50, 15
                  Text 170, 10, 135, 10, "On Demand Applications Dashboard"
                  GroupBox 10, 20, 430, 85, "Applications BULK Run"
                  Text 20, 35, 170, 10, "The BULK run was last completed on " & first_item_date &" ."
                  Text 190, 35, 205, 10, "The BULK Run was completed today around " & first_item_time & "."
                  Text 20, 50, 305, 10, "The BULK Run can only be completed once per day. The Work List is ready to be reviewed."
                  Text 40, 60, 305, 10, "- The worklist is held in a SQL Table and can only be viewed through this Dashboard script. "
                  Text 40, 70, 245, 10, "- Use this script to pull a case from the Work List and complete the review."
                  Text 140, 90, 190, 10, "If the BULK Run was not completed, you can restart here:"
                  GroupBox 10, 110, 430, 35, "Work List Overview"
                  Text 20, 125, 115, 10, "Total cases on the worklist: " & total_cases_for_review
                  Text 220, 120, 130, 10, "Cases with Review Completed: " & cases_with_review_completed
                  Text 215, 130, 135, 10, "Cases with Reviews In Progress: " & cases_on_hold
                  GroupBox 10, 150, 430, 55, "Reviews"
                  Text 20, 190, 190, 10, "--- There are no cases with reviews already started. ---"
                EndDialog

                dialog Dialog1
                cancel_without_confirmation

                If ButtonPressed = work_list_details_btn Then MsgBox "More details will be here" 'TODO - add worklist Explanation'
                If ButtonPressed = test_access_btn Then Call test_sql_access()

                If ButtonPressed = finish_work_day_btn Then
                    Call assess_worklist_to_finish_day
                    If case_on_hold = False and case_in_progress = False Then
                        Call create_assignment_report
                    Else
                        loop_dlg_msg = "You cannot finish the work day with cases in progress or on hold." & vbCr
                        loop_dlg_msg = loop_dlg_msg & "The dialog will reappear, finish all reviews that have been started first." & vbCr & vbCr
                        loop_dlg_msg = loop_dlg_msg & "Once there are no cases on the worklist on hold or in progress the finish work day functionality will operate."
                        ButtonPressed = work_list_details_btn
                        MsgBox loop_dlg_msg
                    End If
                End If
                If ButtonPressed = admin_btn Then call complete_admin_functions
            Loop until ButtonPressed <> work_list_details_btn
            CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
        Loop until are_we_passworded_out = false					'loops until user passwords back in

        If ButtonPressed = get_new_case_btn Then Call assign_a_case
        If ButtonPressed = bulk_run_incomplete_btn Then
            Call back_to_SELF
            EMReadScreen MX_region, 10, 22, 48
            MX_region = trim(MX_region)
            If MX_region <> "PRODUCTION" Then Call script_end_procedure("You have selected to complete the BULK Run for On Demand but you are not in production. The script will now end. Move to PRODUCTION and run On Demand Dashboard again.")
            Call run_from_GitHub(script_repository & "admin\" & "on-demand-waiver-applications.vbs")
        End If
    Else
        Do
            Do
                Dialog1 = ""
                grp_len = 45 + (UBound(CASES_ON_HOLD_ARRAY, 2)+1) * 10      '85'
                dlg_len = 205 + (UBound(CASES_ON_HOLD_ARRAY, 2)+1) * 10     '245'

                BeginDialog Dialog1, 0, 0, 451, dlg_len, "On Demand Applications Dashboard"
                  EditBox 500, 600, 50, 15, fake_edit_box
                  ButtonGroup ButtonPressed
                    If total_cases_for_review <> cases_with_review_completed Then PushButton 20, 145, 110, 15, "Pull a case to Review", get_new_case_btn
                    PushButton 375, 20, 65, 15, "Test Access", test_access_btn
                    If total_cases_for_review = cases_with_review_completed Then Text 20, 150, 155, 10, "*** All Cases have been Pulled for Review ***"'
                  Text 320, 5, 135, 10, "On Demand Applications Dashboard"
                  GroupBox 10, 10, 430, 75, "Applications BULK Run"
                  Text 20, 25, 170, 10, "The BULK run was last completed on " & first_item_date & "."
                  Text 190, 25, 180, 10, "The BULK Run was completed today around " & first_item_time & "."
                  Text 20, 40, 305, 10, "The BULK Run can only be completed once per day. The Work List is ready to be reviewed."
                  Text 40, 50, 305, 10, "- The worklist is held in a SQL Table and can only be viewed through this Dashboard script. "
                  Text 40, 60, 245, 10, "- Use this script to pull a case from the Work List and complete the review."
                  ButtonGroup ButtonPressed
                    PushButton 45, 70, 170, 10, "More information about the Work List", work_list_details_btn
                  GroupBox 10, 90, 430, 40, "Work List Overview"
                  Text 20, 105, 115, 10, "Total cases on the worklist: " & total_cases_for_review
                  Text 220, 105, 130, 10, "Cases with Review Completed: " & cases_with_review_completed
                  Text 215, 115, 135, 10, "Cases with Reviews In Progress: " & cases_on_hold
                  GroupBox 10, 135, 430, grp_len, "Reviews"
                  Text 20, 165, 125, 10, "Reviews Started and put on Hold:"
                  OptionGroup RadioGroup1
                    y_pos = 175
                    For fold_case = 0 to UBound(CASES_ON_HOLD_ARRAY, 2)
                        RadioButton 30, y_pos, 300, 10, "CASE # " & CASES_ON_HOLD_ARRAY(case_nbr_const, fold_case) & " - " & CASES_ON_HOLD_ARRAY(case_notes_const, fold_case), CASES_ON_HOLD_ARRAY(radio_btn_const, fold_case)
                        y_pos = y_pos + 10
                    Next
                  ButtonGroup ButtonPressed
                    PushButton 330, dlg_len - 45, 105, 15, "Resume selected Hold Case", resume_hold_case_btn
                    PushButton 10, dlg_len - 20, 105, 15, "Finish Work Day", finish_work_day_btn
                    CancelButton 390, dlg_len - 20, 50, 15
                EndDialog

                dialog Dialog1
                cancel_confirmation

                If ButtonPressed = work_list_details_btn Then MsgBox "More details will be here" 'TODO - add worklist Explanation'
                If ButtonPressed = test_access_btn Then Call test_sql_access()

                If ButtonPressed = finish_work_day_btn Then
                    Call assess_worklist_to_finish_day
                    If case_on_hold = False and case_in_progress = False Then
                        Call create_assignment_report
                    Else
                        loop_dlg_msg = "You cannot finish the work day with cases in progress or on hold." & vbCr
                        loop_dlg_msg = loop_dlg_msg & "The dialog will reappear, finish all reviews that have been started first." & vbCr & vbCr
                        loop_dlg_msg = loop_dlg_msg & "Once there are no cases on the worklist on hold or in progress the finish work day functionality will operate."
                        ButtonPressed = work_list_details_btn
                        MsgBox loop_dlg_msg
                    End If
                End If
            Loop until ButtonPressed <> work_list_details_btn
            CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
        Loop until are_we_passworded_out = false					'loops until user passwords back in
        If ButtonPressed = resume_hold_case_btn Then
            resume_case_number = CASES_ON_HOLD_ARRAY(case_nbr_const, RadioGroup1)
            MAXIS_case_number = resume_case_number
            txt_file_name = user_ID_for_validation & "_" & MAXIS_case_number & "_" & file_date & ".txt"
            od_revw_tracking_file_path = current_day_work_tracking_folder  & txt_file_name

            Call read_tracking_cookie
            If assigned_hold_1_start_time = "" Then
                assigned_hold_1_start_time = time
            ElseIf assigned_hold_2_start_time = "" Then
                assigned_hold_2_start_time = time
            ElseIf assigned_hold_3_start_time = "" Then
                assigned_hold_3_start_time = time
            End If
            assign_a_case
        End If
        If ButtonPressed = get_new_case_btn Then Call assign_a_case
    End If
End If

If worker_on_task = True Then

    If local_demo = True Then
        MAXIS_case_number                       = "318040"
        ' assigned_ = objRecordSet("CaseNumber")
        assigned_case_name                      = "ZORO, RORONOA"
        assigned_application_date               = "11/23/2022"
        assigned_interview_date                 = ""
        assigned_day_30                         = "12/23/2022"
        assigned_days_pending                   = "12"
        assigned_snap_status                    = "Pending"
        assigned_cash_status                    = ""
        assigned_2nd_application_date           = ""
        assigned_rept_pnd2_days                 = "12"
        assigned_questionable_interview         = ""
        assigned_questionable_interview_resolve = ""
        assigned_appt_notice_date               = "11/25/2022"
        assigned_appt_date                      = "12/5/2022"
        assigned_appt_notc_confirmation         = "Y"
        assigned_nomi_date                      = "12/5/2022"
        assigned_nomi_confirmation              = "Y"
        assigned_denial_needed                  = ""

        assigned_next_action_needed             = "DO THIS NEXT"


        assigned_added_to_work_list             = date
        assigned_2nd_application_date_resolve   = ""
        assigned_closed_recently                = ""
        assigned_closed_recently_resolve        = ""
        assigned_out_of_county                  = ""
        assigned_out_of_county_resolve          = ""
        case_review_notes                 = "Saved notes are here."

        ' assigned_next_action_needed = "ALIGN INTERVIEW DATES"
        ' assigned_interview_date = "12/02/22"

        ' assigned_next_action_needed = "REVIEW QUESTIONABLE INTERVIEW DATE(S)"
        ' assigned_questionable_interview = "12/01/22"

        assigned_next_action_needed = "RESOLVE SUBSEQUENT APPLICATION DATE"
        assigned_2nd_application_date = "12/5/22"

        ' assigned_next_action_needed = "REVIEW RECENT CLOSURE/DENIAL"

        ' assigned_next_action_needed = "REVIEW OTHER COUNTY CASE"

        ' assigned_next_action_needed = "REVIEW CANNOT DENY - NOMI after Day 30"
        ' assigned_day_30 = date
        ' assigned_application_date = DateAdd("d", -30, date)
        ' assigned_appt_notice_date = DateAdd("d", 22, assigned_application_date)
        ' assigned_appt_date = DateAdd("d", 5, assigned_appt_notice_date)
        ' assigned_nomi_date = date
        ' assigned_days_pending = "30"
        ' assigned_rept_pnd2_days = "30"

        ' assigned_next_action_needed = "PREP FOR DENIAL"
        ' assigned_day_30 = DateAdd("d", 1, date)
        ' assigned_application_date = DateAdd("d", -29, date)
        ' assigned_appt_notice_date = DateAdd("d", 3, assigned_application_date)
        ' assigned_appt_date = DateAdd("d", 5, assigned_appt_notice_date)
        ' assigned_nomi_date = assigned_appt_date
        ' assigned_days_pending = "29"
        ' assigned_rept_pnd2_days = "29"
    End If

	txt_file_name = user_ID_for_validation & "_" & MAXIS_case_number & "_" & file_date & ".txt"
    od_revw_tracking_file_path = current_day_work_tracking_folder  & txt_file_name
	call read_tracking_cookie

	Do
	    Do
	        err_msg = ""

            Dialog1 = ""
	        BeginDialog Dialog1, 0, 0, 451, 310, "On Demand Applications Case Review"
	          Text 185, 10, 60, 10, "Case in Review"
	          GroupBox 10, 20, 230, 75, "Case Information"
	          Text 20, 35, 85, 10, " Case Number: " & MAXIS_case_number
	          Text 30, 45, 210, 10, "Case Name: " & assigned_case_name
	          Text 15, 60, 120, 10, "Application Date: " & assigned_application_date
	          Text 20, 70, 75, 10, " Days Pending: " & assigned_rept_pnd2_days
	          Text 45, 80, 80, 10, "Day 30: " & assigned_day_30
	          If assigned_snap_status = "Pending" Then Text 145, 60, 50, 10, "SNAP: Pending"
	          If assigned_cash_status = "Pending" Then Text 145, 70, 50, 10, "CASH: Pending"

	          GroupBox 10, 100, 230, 95, "Interview"
	          If assigned_interview_date = "" Then Text 20, 115, 140, 10, "No interview entered on PROG"
              If assigned_interview_date <> "" Then Text 20, 115, 140, 10, "Interview Date from PROG: " & assigned_interview_date

	          If assigned_next_action_needed = "ALIGN INTERVIEW DATES" Then
	        	  Text 20, 130, 180, 10, "*** Interview Dates on PROG need to be ALIGNED ***"
	        	  Text 20, 145, 40, 10, "Resolution: "
	        	  EditBox 60, 140, 170, 15, align_interview_dates_resolution
	          End If
	          If assigned_next_action_needed = "REVIEW QUESTIONABLE INTERVIEW DATE(S)" Then
	        	  Text 20, 165, 180, 10, "*** Questionable Interview Date Found: " & assigned_questionable_interview & " ***"
	        	  Text 20, 180, 40, 10, "Resolution: "
	        	  EditBox 60, 175, 170, 15, assigned_questionable_interview_resolve
	          End If
	          GroupBox 10, 200, 230, 55, "Notices"
	          If assigned_appt_notice_date = "" Then
	          	  Text 20, 215, 110, 10, "NO APPOINTMENT NOTICE FOUND"
	          Else
	        	  Text 20, 215, 110, 10, "Appt Notice Sent on " & assigned_appt_notice_date
	        	  Text 30, 225, 80, 10, "Appt Date: " & assigned_appt_date
	          End If
	          If assigned_nomi_date = "" Then
	        	  Text 20, 240, 110, 10, "NO NOMI FOUND"
	          Else
	        	  Text 20, 240, 110, 10, "NOMI Sent on " & assigned_nomi_date
	          End If
	          GroupBox 250, 20, 195, 235, "Actions"
	          Text 260, 35, 70, 10, "Next Action Needed:"
	          Text 265, 45, 165, 10, assigned_next_action_needed
	          If assigned_next_action_needed = "RESOLVE SUBSEQUENT APPLICATION DATE" Then
	        	  Text 260, 65, 145, 10, "Subsequent Appliction Date: " & assigned_2nd_application_date
	        	  Text 260, 75, 50, 10, "Resolution:"
	        	  EditBox 260, 85, 175, 15, subseuent_application_resolution
	          End If
	          If assigned_next_action_needed = "REVIEW RECENT CLOSURE/DENIAL" Then
	        	  Text 260, 65, 145, 10, "Case closed in past 30 Days"
	        	  Text 260, 75, 50, 10, "Review Notes:"
	        	  EditBox 260, 85, 175, 15, assigned_closed_recently_resolve
	          End If
	          If assigned_next_action_needed = "REVIEW OTHER COUNTY CASE" Then
	        	  Text 260, 65, 145, 10, "Case was in Another County"
	        	  Text 260, 75, 50, 10, "Review Notes:"
	        	  EditBox 260, 85, 175, 15, assigned_out_of_county_resolve
	          End If
	          If left(assigned_next_action_needed, 18) = "REVIEW CANNOT DENY" Then
	        	  Text 260, 65, 175, 10, "Case Cannot be Denied"
	        	  ' Text 260, 210, 175, 10, "REASON"
	        	  If assigned_next_action_needed = "REVIEW CANNOT DENY - No Appt Notc" Then Text 260, 75, 175, 10, "No Appointment Notice"
	        	  If assigned_next_action_needed = "REVIEW CANNOT DENY - No NOMI" Then Text 260, 75, 175, 10, "No NOMI"
	        	  If assigned_next_action_needed = "REVIEW CANNOT DENY - NOMI after Day 30" Then Text 260, 75, 175, 10, "NOMI after Day 30"
	        	  Text 260, 85, 50, 10, "Notes:"
	        	  EditBox 260, 95, 175, 15, cannot_deny_resolution
	          End If
              CheckBox 260, 240, 180, 10, "Check here if this case requires follow up tomorrow.", follow_up_tomorrow_checkbox
	          Text 10, 260, 70, 10, "Additional Notes:"
	          EditBox 10, 270, 435, 15, case_review_notes
	          ButtonGroup ButtonPressed
	            PushButton 280, 290, 110, 15, "Complete Review", complete_review_btn
	            PushButton 335, 5, 110, 15, "Put Case on Hold", hold_case_btn
                PushButton 160, 290, 110, 15, "Close Review Dialog", close_dialog_btn
                PushButton 10, 290, 65, 15, "Test Access", test_access_btn
                If ADMIN_run = True Then PushButton 80, 290, 70, 15, "Admin Functions", admin_btn
	            CancelButton 395, 290, 50, 15
	        EndDialog

	        dialog Dialog1
	        cancel_confirmation

            If ButtonPressed = -1 or ButtonPressed = close_dialog_btn Then script_end_procedure("")
            If ButtonPressed = test_access_btn Then Call test_sql_access()
            If ButtonPressed = admin_btn Then call complete_admin_functions
        Loop until err_msg = ""
	    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
	Loop until are_we_passworded_out = false					'loops until user passwords back in

    If ButtonPressed = complete_review_btn Then
        txt_file_name = user_ID_for_validation & "_" & MAXIS_case_number & "_" & file_date & ".txt"
        od_revw_tracking_file_path = current_day_work_tracking_folder  & txt_file_name
		Call read_tracking_cookie

        case_review_notes = replace(case_review_notes, "STS-IP-"&user_ID_for_validation, "")
        case_review_notes = replace(case_review_notes, "STS-HD-"&user_ID_for_validation, "")
        case_review_notes = replace(case_review_notes, "STS-RC-"&user_ID_for_validation, "")
        case_review_notes = replace(case_review_notes, "STS-RC", "")
        case_review_notes = trim(case_review_notes)
        case_review_notes = "STS-RC-"&user_ID_for_validation & " " & case_review_notes
        If follow_up_tomorrow_checkbox = checked Then case_review_notes = "FOLLOW UP NEEDED - " & case_review_notes
        end_msg = end_msg & vbCr & vbCr & "The review for Case # " & MAXIS_case_number & " has been completed."

        assigned_end_time = time
        Call update_tracking_cookie("END")
    End If

	If ButtonPressed = hold_case_btn Then
        txt_file_name = user_ID_for_validation & "_" & MAXIS_case_number & "_" & file_date & ".txt"
        od_revw_tracking_file_path = current_day_work_tracking_folder  & txt_file_name
		Call read_tracking_cookie

        If assigned_hold_1_start_time = "" Then
			assigned_hold_1_start_time = time
		ElseIf assigned_hold_2_start_time = "" Then
			assigned_hold_2_start_time = time
		ElseIf assigned_hold_3_start_time = "" Then
			assigned_hold_3_start_time = time
		End If
        ' "STS-HD"
        end_msg = end_msg & vbCr & vbCr & "Case # " & MAXIS_case_number & " has been put on hold to be reviewed later today."
        case_review_notes = replace(case_review_notes, "STS-IP-"&user_ID_for_validation, "")
        case_review_notes = replace(case_review_notes, "STS-HD-"&user_ID_for_validation, "")
        case_review_notes = replace(case_review_notes, "STS-RC-"&user_ID_for_validation, "")
        case_review_notes = replace(case_review_notes, "STS-RC", "")
        case_review_notes = trim(case_review_notes)
        case_review_notes = "STS-HD-"&user_ID_for_validation & " " & case_review_notes
        Call update_tracking_cookie("HOLD")
	End If

    If local_demo = False Then
        'Creating objects for Access
        Set objConnection = CreateObject("ADODB.Connection")
        Set objRecordSet = CreateObject("ADODB.Recordset")

        'This is the BZST connection to SQL Database'
        objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""

        'delete a record if the case number matches
        objRecordSet.Open "UPDATE ES.ES_OnDemanCashAndSnapBZProcessed SET CaseNumber = '" & MAXIS_case_number & "', " &_
                                                                          "CaseName = '" & assigned_case_name & "', " &_
                                                                          "ApplDate = '" & table_application_date & "', " &_
                                                                          "InterviewDate = '" & table_interview_date & "', " &_
                                                                          "Day_30 = '" & table_day_30 & "', " &_
                                                                          "DaysPending = '" & assigned_days_pending & "', " &_
                                                                          "SnapStatus = '" & assigned_snap_status & "', " &_
                                                                          "CashStatus = '" & assigned_cash_status & "', " &_
                                                                          "SecondApplicationDate = '" & table_2nd_application_date & "', " &_
                                                                          "REPT_PND2Days = '" & assigned_rept_pnd2_days & "', " &_
                                                                          "QuestionableInterview = '" & assigned_questionable_interview & "', " &_
                                                                          "Resolved = '" & assigned_questionable_interview_resolve & "', " &_
                                                                          "ApptNoticeDate = '" & table_appt_notice_date & "', " &_
                                                                          "ApptDate = '" & table_appt_date & "', " &_
                                                                          "Confirmation = '" & assigned_appt_notc_confirmation & "', " &_
                                                                          "NOMIDate = '" & table_nomi_date & "', " &_
                                                                          "Confirmation2 = '" & assigned_nomi_confirmation & "', " &_
                                                                          "DenialNeeded = '" & assigned_denial_needed & "', " &_
                                                                          "NextActionNeeded = '" & assigned_next_action_needed & "', " &_
                                                                          "AddedtoWorkList = '" & table_added_to_work_list & "', " &_
                                                                          "SecondApplicationDateNotes = '" & assigned_2nd_application_date_resolve & "', " &_
                                                                          "ClosedInPast30Days = '" & assigned_closed_recently & "', " &_
                                                                          "ClosedInPast30DaysNotes = '" & assigned_closed_recently_resolve & "', " &_
                                                                          "StartedOutOfCounty = '" & assigned_out_of_county & "', " &_
                                                                          "StartedOutOfCountyNotes = '" & assigned_out_of_county_resolve & "', " &_
                                                                          "TrackingNotes = '" & case_review_notes & "' WHERE CaseNumber = '" & MAXIS_case_number & "'", objConnection

        'close the connection and recordset objects to free up resources
        objConnection.Close
        Set objRecordSet=nothing
        Set objConnection=nothing
    End If
End If

Call script_end_procedure(end_msg)
