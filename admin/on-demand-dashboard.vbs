'GATHERING STATS===========================================================================================
name_of_script = "ADMIN - On Demand Dashboard.vbs"
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
call changelog_update("01/15/2024", "Since the Interview Waiver has ended, the functionality for allowing interview date misalignment has been removed.", "Casey Love, Hennepin County")
call changelog_update("07/24/2024", "Added an option for Interviews that do not have Aligned dates to indicate the case does not have to be reviewed for a week. This is intending to reduce the number of cases on the list to review on a daily basis.", "Casey Love, Hennepin County")
call changelog_update("07/21/2023", "Updated function that sends an email through Outlook", "Mark Riegel, Hennepin County")
call changelog_update("12/06/2022", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'Information about this script and how it works
' STS-NR            'Status - Needs Review'
' STS-RC-WFXXXX     'Status - Review Completed - Worker number'
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
assigned_case_is_priv = ""
case_review_notes = ""
case_on_hold = False
case_in_progress = False
completed_reviews = 0
reviews_completed_by_me = 0
reviews_still_needed = 0

end_msg = ""
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

local_demo = False
ADMIN_run = False
BULK_Run_completed = False
worker_on_task = False
finish_day_completed_yesterday = True
workers_first_task_pulled_for_review = False
total_cases_for_review = 0
admin_cases_for_review = 0
cases_with_review_completed = 0
cases_waiting_for_review = 0
cases_on_hold = 0
cases_completed_by_current_worker = 0
case_nbr_in_progress = ""
admin_count_NR = 0
admin_count_RC = 0
admin_count_IP = 0
admin_count_HD = 0
ADMIN_list_workers_RC = "~"
ADMIN_list_workers_IP = "~"
ADMIN_list_workers_HD = "~"

script_instructions_file_path = "T:\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\On Demand Waiver\QI On Demand Daily Assignment\On Demand Dashboard Script Instructions.docx"
worklist_instructions_file_path = "T:\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\On Demand Waiver\QI On Demand Daily Assignment\QI On Demand Case Review Processing.docx"
current_day_work_tracking_folder = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\On Demand Waiver\QI On Demand Daily Assignment\Archive\WIP\"
'txt file name format
' USERID-MM-DD-YY-CASENUMBER
curr_day = DatePart("d", date)
curr_day = right("00" & curr_day, 2)
file_date = CM_mo & "-" & curr_day & "-" & CM_yr

txt_file_name = user_ID_for_validation & "_" & MAXIS_case_number & "_" & file_date & ".txt"
od_revw_tracking_file_path = current_day_work_tracking_folder  & txt_file_name


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
script_instructions_btn	= 7001
worklist_process_doc_btn = 7002

const case_nbr_const    = 0
const radio_btn_const   = 1
const case_notes_const  = 2
const last_const        = 10

Dim CASES_ON_HOLD_ARRAY()
ReDim CASES_ON_HOLD_ARRAY(last_const, 0)

const wrkr_id_const         = 0
const wrkr_name_const       = 1
const case_status_const     = 2
const admin_radio_btn_const = 3
const case_count_const		= 4
const admin_wrkr_last_const = 5

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


bulk_run_information_msg = "     * * * * *  The BULK On Demand Application Run Information  * * * * *"
bulk_run_information_msg = bulk_run_information_msg & vbCr & vbCr & "The BULK run is the original functionality of the On Demand Applications script. This is the core of the work for On Demand."
bulk_run_information_msg = bulk_run_information_msg & vbCr & vbCr & "This functionality uses data tables of all cases pending SNAP or Cash in the agency that do not have an interview date entered into MAXIS. The primary purpose of the On Demand process is to ensure proper notification is sent to the residents of the interview requirement for new applications. This includes the initial interview information and a NOMI."
bulk_run_information_msg = bulk_run_information_msg & vbCr & vbCr & "The BULK Run automates the evaluation and sending of these notices for most cases."
bulk_run_information_msg = bulk_run_information_msg & vbCr & vbCr & "In the course of running the BULK Run, the script may identify some instances in which it cannot fully evaluate the status of the case, or the case has hit a point that a manual review is advised. The On Demand worklist is generated from these cases during the BULK run."
bulk_run_information_msg = bulk_run_information_msg & vbCr & vbCr & "--------------------------------------------------------------------------------------------"
bulk_run_information_msg = bulk_run_information_msg & vbCr & vbCr & vbCR & "The BULK run should only be run once per day and requires no input during the run."
bulk_run_information_msg = bulk_run_information_msg & vbCr & vbCr & "The run takes a couple hours and ties up production, though Inquiry or Training region can be used in a different session."
bulk_run_information_msg = bulk_run_information_msg & vbCr & vbCr & "Any notices sent or CASE/NOTEs entered during the BULK run are in the X-Number of the worker who started the run."
bulk_run_information_msg = bulk_run_information_msg & vbCr & vbCr & "The BULK run should be completed on every working day."
bulk_run_information_msg = bulk_run_information_msg & vbCr & vbCr & "The BULK run works best when run in the morning, as fewer people are working on cases."

worklist_information_msg = "     * * * * *  The On Demand QI Daily Worklist Information  * * * * *"
worklist_information_msg = worklist_information_msg & vbCr & vbCr & "The Worklist is the list of cases that require manual review or action from the On Demand BULK Run. This is specifically for cases at application that have not yet had an interview completed."
worklist_information_msg = worklist_information_msg & vbCr & vbCr & "This worklist used to be stored in daily Excel files. These files became to unweildy to manage and had limited capabilities for data collection and review, so the list has been movved to a SQL table in the server. This will increase data integrity and ability to manage the list."
worklist_information_msg = worklist_information_msg & vbCr & vbCr & "Changing the methodology for the list changes how we can interact with it, mostly for the better. The cases on the list must be accessed through the On Demand Dashboard script as the scripts are able to connect to the data tables."
worklist_information_msg = worklist_information_msg & vbCr & vbCr & "The benefits of moving to this method of data storage:"
worklist_information_msg = worklist_information_msg & vbCr & " - Cases can now be reviewed one at a time."
worklist_information_msg = worklist_information_msg & vbCr & " - Additional workers can be assigned to support the completion of the worklist, even in the same day"
worklist_information_msg = worklist_information_msg & vbCr & " - Skill in using Excel is no longer needed."
worklist_information_msg = worklist_information_msg & vbCr & " - The BULK Run is better able to pull information from the worklist."
worklist_information_msg = worklist_information_msg & vbCr & vbCr & "--------------------------------------------------------------------------------------------"
worklist_information_msg = worklist_information_msg & vbCr & vbCr & "If you have any questions on the script functionality for pulling cases for review from the worklist, connect with the BlueZone Script team."
worklist_information_msg = worklist_information_msg & vbCr & vbCr & "If you need support with how to complete a review on a case, connect with Tanya or the QI Worklsit Process documenntation."
'END DECLARATIONS ==========================================================================================================

'FUNCTIONS BLOCK ===========================================================================================================
'this function checks the SQL worklist and determines if there are still cases on hold or in progress that would prevent the completion of the work day
function assess_worklist_to_finish_day()
    case_on_hold = False			'these are the booleans that come out of this function to indicate cases are still being worked
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
    objConnection.Open db_full_string
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

'this function will assign the case informaiton to a specific worker.
function assign_a_case()
    If local_demo = False Then					'if not a DEMO run, pull case information from the SQL table
		'if we need to get a new case, we need to find a case with the status of 'STS-NR' since that is a case that needs review
        If ButtonPressed = get_new_case_btn Then
            objSQL = "SELECT * FROM ES.ES_OnDemanCashAndSnapBZProcessed"

            'Creating objects for Access
            Set objConnection = CreateObject("ADODB.Connection")
            Set objRecordSet = CreateObject("ADODB.Recordset")

            'This is the file path for the statistics Access database.
            objConnection.Open db_full_string
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

		'if we are resuming a hold case, we already know the case number and can resume it directly using the case number
        If ButtonPressed = resume_hold_case_btn Then
            'Creating objects for Access
            Set objConnection = CreateObject("ADODB.Connection")
            Set objRecordSet = CreateObject("ADODB.Recordset")

            'This is the file path for the statistics Access database.
            objConnection.Open db_full_string
            objRecordSet.Open "SELECT * FROM ES.ES_OnDemanCashAndSnapBZProcessed WHERE CaseNumber = '" & resume_case_number & "'", objConnection

            set_variables_from_SQL

            'close the connection and recordset objects to free up resources
            objRecordSet.Close
            objConnection.Close
            Set objRecordSet=nothing
            Set objConnection=nothing
        End If
    Else		'these are for demo cases
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
    end_msg = end_msg & vbCr & vbCr & "You have a case selected for review: " & MAXIS_case_number		'saving and formatting the information
	case_review_notes = replace(case_review_notes, "'", "")								'remove any single quote from the string because it is a reserved character in SQL
    assigned_tracking_notes = "STS-IP-"&user_ID_for_validation & " " & case_review_notes

	'if not a demo case, the updated information should be saved to SQL
    If local_demo = False Then
        'Creating objects for Access
        Set objConnection = CreateObject("ADODB.Connection")
        Set objRecordSet = CreateObject("ADODB.Recordset")

        'This is the BZST connection to SQL Database'
        objConnection.Open db_full_string

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
	If workers_first_task_pulled_for_review = True Then
		appt_body = "Once you are completed with the work for On Demand be sure to run the On Demand Dashboard Finish Day functionality."
		appt_header = "REMINDER - Run ON DEMAND Finish Day"

		call create_outlook_appointment(date, "03:00 PM", "03:30 PM", appt_header, appt_body, "", True, 15, "")

		end_msg = end_msg & vbCr & vbCr & "THE SCRIPT HAS SET AN OUTLOOK REMINDER TO RUN FINISH DAY."
		end_msg = end_msg & vbCr & "It is set for 3:00 PM but you can change it best match your work day."
		end_msg = end_msg & vbCr & "The On Demand process requires 'Finish Day' to be run by every worker that completes a case review in the day."
	End if

	'this part is to document some time information when assigning the case
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

	'going to CASE/CURR for the case
    Call Back_to_SELF
    Call navigate_to_MAXIS_screen("CASE", "CURR")
	'after here, the script will go to the functionality where a worker is indicated on task to display the dialog with detail about the case.
end function

'this is the function to display information about all the cases and any that are being worked or has been worked
function complete_admin_functions()
	'show the information read from the SQL table
    Do
        Do
            err_msg = ""

            dlg_len = 165 + 10 * (UBound(ADMIN_worker_list_array, 2)+1)
            grp_len = 120 + 10 * (UBound(ADMIN_worker_list_array, 2)+1)

            Dialog1 = ""
            BeginDialog Dialog1, 0, 0, 451, dlg_len, "On Demand Applications Dashboard"
              ButtonGroup ButtonPressed
                Text 170, 10, 135, 10, "On Demand Applications Dashboard"
                GroupBox 10, 15, 430, grp_len, "On Demand Case List Information"
				Text 20, 30, 230, 10, "Information from list run on: " & first_item_date
                Text 20, 40, 230, 10, "On " & first_item_date & ", " & admin_cases_for_review & " cases require review. These cases are currently:"
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
                            RadioButton 50, y_pos, 300, 10, ADMIN_worker_list_array(wrkr_name_const, worker_indc) & " - " & ADMIN_worker_list_array(wrkr_id_const, worker_indc) & ": " & ADMIN_worker_list_array(case_count_const, worker_indc), btn_hold', ADMIN_worker_list_array(admin_radio_btn_const, worker_indc)
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
                            RadioButton 50, y_pos, 300, 10, ADMIN_worker_list_array(wrkr_name_const, worker_indc) & " - " & ADMIN_worker_list_array(wrkr_id_const, worker_indc) & ": " & ADMIN_worker_list_array(case_count_const, worker_indc), btn_hold', ADMIN_worker_list_array(admin_radio_btn_const, worker_indc)
                            y_pos = y_pos + 10
                        End If
                    Next
                End If
                y_pos = y_pos + 5

                OkButton 335, dlg_len-20, 50, 15
                CancelButton 390, dlg_len-20, 50, 15
                EditBox 500, 300, 50, 15, fake_edit_box
            EndDialog

            Dialog Dialog1
            cancel_without_confirmation

			'this part will determine the worker number of the bselected group if a button to release assignment or finish the day
            worker_number_to_resolve = ""
            If ButtonPressed = release_IP_btn Then
                For worker_indc = 0 to UBound(ADMIN_worker_list_array, 2)
                    If RadioGroupIP = ADMIN_worker_list_array(admin_radio_btn_const, worker_indc) Then worker_number_to_resolve = ADMIN_worker_list_array(wrkr_id_const, worker_indc)
                Next
            ElseIf ButtonPressed = release_HD_btn Then
                For worker_indc = 0 to UBound(ADMIN_worker_list_array, 2)
                    If RadioGroupHD = ADMIN_worker_list_array(admin_radio_btn_const, worker_indc) Then worker_number_to_resolve = ADMIN_worker_list_array(wrkr_id_const, worker_indc)
                Next
            ElseIf ButtonPressed = finish_day_btn Then
                For worker_indc = 0 to UBound(ADMIN_worker_list_array, 2)
                    If RadioGroupRC = ADMIN_worker_list_array(admin_radio_btn_const, worker_indc) Then worker_number_to_resolve = ADMIN_worker_list_array(wrkr_id_const, worker_indc)
                Next
            End If

			'getting some detail about the worker that is selected
            For tester = 0 to UBound(tester_array)                         'looping through all of the testers
                ' pulling QI members by supervisor from the Complete List of Testers
                If tester_array(tester).tester_id_number = worker_number_to_resolve Then
                    worker_full_name_to_resolve = tester_array(tester).tester_full_name
                    qi_worker_supervisor_email = tester_array(tester).tester_supervisor_email
					qi_worker_email = tester_array(tester).tester_email
                    qi_worker_first_name = tester_array(tester).tester_first_name
					If tester_array(tester).tester_population = "QI" Then qi_member_identified = True
                    If tester_array(tester).tester_population = "BZ" Then qi_member_identified = True
                    assigned_worker = tester_array(tester).tester_full_name
                    ' MsgBox "user_ID_for_validation - " & user_ID_for_validation & vbCr & "tester_array(tester).tester_id_number - " & tester_array(tester).tester_id_number & vbCr & "qi_member_identified - " & qi_member_identified
                End If
            Next
			'end message details
			end_msg = "ADMIN Function completed. " & vbCr & vbCr & "Worker selected: " & worker_number_to_resolve & vbCr & "This worker's task"
            If ButtonPressed = release_IP_btn Then
				end_msg = end_msg & " IN PROGRESS was released."
            ElseIf ButtonPressed = release_HD_btn Then
				end_msg = end_msg & "s ON HOLD were released."
            ElseIf ButtonPressed = finish_day_btn Then
				end_msg = end_msg & "s that were COMPLETED today were logged and FINISH DAY was run."
            End If

			'This part will loop through all of the cases on the SQL table to find if the notes indicate the status meets the requirements of the selection
			Do
	            case_to_fix_found = False
                'declare the SQL statement that will query the database
                objSQL = "SELECT * FROM ES.ES_OnDemanCashAndSnapBZProcessed"

                'Creating objects for Access
                Set objConnection = CreateObject("ADODB.Connection")
                Set objRecordSet = CreateObject("ADODB.Recordset")

                'This is the file path for the statistics Access database.
                objConnection.Open db_full_string
                objRecordSet.Open objSQL, objConnection

                Do While NOT objRecordSet.Eof
                    case_tracking_notes = objRecordSet("TrackingNotes")
                    worklist_case_number = objRecordSet("CaseNumber")

                    If ButtonPressed = release_IP_btn Then			'if the button was pressed to release a case in progress, looking for STS-IP with the worker ID listed
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
                            End If
                        End if
                    ElseIf ButtonPressed = release_HD_btn Then			'if the button was pressed to release cases on hold, looking for STS-HD with the worker ID listed
                        If Instr(case_tracking_notes, "STS-HD") Then
                            If Instr(case_tracking_notes, worker_number_to_resolve) <> 0 Then
                                case_to_fix_found = True
                                case_tracking_notes = replace(case_tracking_notes, "STS-IP-"&worker_number_to_resolve, "")
                                case_tracking_notes = replace(case_tracking_notes, "STS-HD-"&worker_number_to_resolve, "")
                                case_tracking_notes = replace(case_tracking_notes, "STS-RC-"&worker_number_to_resolve, "")
                                case_tracking_notes = replace(case_tracking_notes, "STS-RC", "")
                                case_tracking_notes = trim(case_tracking_notes)
								case_tracking_notes = replace(case_tracking_notes, "'", "")								'remove any single quote from the string because it is a reserved character in SQL
                                case_tracking_notes = "STS-NR " & case_tracking_notes
                                case_tracking_notes = trim(case_tracking_notes)
                                Exit Do
                            End If
                        End if
                    End If

                    objRecordSet.MoveNext		'going to the next case
                Loop

                'close the connection and recordset objects to free up resources
                objRecordSet.Close
                objConnection.Close
                Set objRecordSet=nothing
                Set objConnection=nothing

				If case_to_fix_found = True Then		'if a case was found to fix, this will pull the case and update the tracking notes and remove the hold/in progress
					Set objConnection = CreateObject("ADODB.Connection")
					Set objRecordSet = CreateObject("ADODB.Recordset")

					'This is the BZST connection to SQL Database'
					objConnection.Open db_full_string

					'delete a record if the case number matches
					objRecordSet.Open "UPDATE ES.ES_OnDemanCashAndSnapBZProcessed SET TrackingNotes = '" & case_tracking_notes & "' WHERE CaseNumber = '" & worklist_case_number & "'", objConnection

					'close the connection and recordset objects to free up resources
					objConnection.Close
					Set objRecordSet=nothing
					Set objConnection=nothing
				End If

                If ButtonPressed = finish_day_btn Then		'if the button was pressed to finish the day, it will pull the right worker and call the finish day function
                    actual_user_ID_for_validation = user_ID_for_validation
                    user_ID_for_validation = worker_number_to_resolve
                    actual_assigned_worker = assigned_worker
                    assigned_worker = worker_full_name_to_resolve
					actual_file_date = file_date
					curr_day = DatePart("d", first_item_date)
					curr_day = right("00" & curr_day, 2)
					curr_month = DatePart("m", first_item_date)
					curr_month = right("00" & curr_month, 2)
					curr_year = DatePart("yyyy", first_item_date)
					curr_year = right(curr_year, 2)
					file_date = curr_month & "-" & curr_day & "-" & curr_year

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
            Loop until case_to_fix_found = False				'we are going to go through the list until no cases are found that meet the criteria to fix
	    Loop until err_msg = ""
	    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
	Loop until are_we_passworded_out = false					'loops until user passwords back in
end function

'this function will log the completed reviews for a specific worker as a part of the Finish Day process
function create_assignment_report()
    cases_completed_by_me = 0

	'opening task log
    file_url = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\On Demand Waiver\QI On Demand Daily Assignment\QI On Demand Work Log.xlsx"
    Call excel_open(file_url, False, False, ObjExcel, objWorkbook)

	ObjExcel.worksheets("QI Review Work").Activate

	''finding the first empty row in the log
    excel_row = 2
	assigned_date = date
    Do While trim(ObjExcel.Cells(excel_row, Worker_col).value) <> ""
		If ObjExcel.Cells(excel_row, Worker_col).value = assigned_worker and ObjExcel.Cells(excel_row, AssignedDate_col).value = assigned_date Then
			cases_completed_by_me = cases_completed_by_me + 1
		End If
		excel_row = excel_row + 1
	Loop

    'Access the SQL Table
    'declare the SQL statement that will query the database
    objSQL = "SELECT * FROM ES.ES_OnDemanCashAndSnapBZProcessed"

    'Creating objects for Access
    Set objConnection = CreateObject("ADODB.Connection")
    Set objRecordSet = CreateObject("ADODB.Recordset")

    'This is the file path for the statistics Access database.
    objConnection.Open db_full_string
    objRecordSet.Open objSQL, objConnection

	'we are going to add each case on the TABLE that was completed by the worker running the script and adding the information to the array
    Do While NOT objRecordSet.Eof
        case_tracking_notes = objRecordSet("TrackingNotes")

        If Instr(case_tracking_notes, "STS-RC-"&user_ID_for_validation) <> 0 Then
       		txt_file_name = ""
			od_revw_tracking_file_path = ""

			assigned_date  = ""
			assigned_start_time  = ""
			assigned_end_time  = ""
			assigned_hold_1_start_time  = ""
			assigned_hold_1_end_time  = ""
			assigned_hold_2_start_time  = ""
			assigned_hold_2_end_time  = ""
			assigned_hold_3_start_time  = ""
			assigned_hold_3_end_time  = ""

	        MAXIS_case_number = objRecordSet("CaseNumber")

			txt_file_name = user_ID_for_validation & "_" & MAXIS_case_number & "_" & file_date & ".txt"
			od_revw_tracking_file_path = current_day_work_tracking_folder  & txt_file_name

			Call read_tracking_cookie			'read details from the tracking cookie to save it to the log
			If assigned_date = "" Then assigned_date = date

			'removing the status tracking information from the notes since the review is now done and logged
			case_review_notes = objRecordSet("TrackingNotes")
			case_review_notes = replace(case_review_notes, "STS-IP-"&user_ID_for_validation, "")
			case_review_notes = replace(case_review_notes, "STS-HD-"&user_ID_for_validation, "")
			case_review_notes = replace(case_review_notes, "STS-RC-"&user_ID_for_validation, "")
			case_review_notes = replace(case_review_notes, "STS-RC", "")
			case_review_notes = replace(case_review_notes, "'", "")								'remove any single quote from the string because it is a reserved character in SQL
			case_review_notes = trim(case_review_notes)

			ObjExcel.Cells(excel_row, Worker_col).value = assigned_worker
			ObjExcel.Cells(excel_row, AssignedDate_col).value = assigned_date
			ObjExcel.Cells(excel_row, CaseNumber_col).value = objRecordSet("CaseNumber")
			ObjExcel.Cells(excel_row, CaseName_col).value =  objRecordSet("CaseName")
			ObjExcel.Cells(excel_row, ApplDate_col).value =  DateAdd("d", 0, objRecordSet("ApplDate"))
			ObjExcel.Cells(excel_row, InterviewDate_col).value =  DateAdd("d", 0, objRecordSet("InterviewDate"))
			ObjExcel.Cells(excel_row, Day_30_dash_col).value =  DateAdd("d", 0, objRecordSet("Day_30"))
			ObjExcel.Cells(excel_row, DaysPending_col).value =  objRecordSet("DaysPending")
			ObjExcel.Cells(excel_row, SnapStatus_col).value =  objRecordSet("SnapStatus")
			ObjExcel.Cells(excel_row, CashStatus_col).value =  objRecordSet("CashStatus")
			ObjExcel.Cells(excel_row, SecondApplicationDate_col).value =  DateAdd("d", 0, objRecordSet("SecondApplicationDate"))
			ObjExcel.Cells(excel_row, REPT_PND2Days_col).value =  objRecordSet("REPT_PND2Days")
			ObjExcel.Cells(excel_row, QuestionableInterview_col).value =  objRecordSet("QuestionableInterview")
			ObjExcel.Cells(excel_row, Resolved_col).value =  objRecordSet("Resolved")
			ObjExcel.Cells(excel_row, ApptNoticeDate_col).value =  DateAdd("d", 0, objRecordSet("ApptNoticeDate"))
			ObjExcel.Cells(excel_row, ApptDate_col).value =  DateAdd("d", 0, objRecordSet("ApptDate"))
			ObjExcel.Cells(excel_row, Confirmation_col).value =  objRecordSet("Confirmation")
			ObjExcel.Cells(excel_row, NOMIDate_col).value =  DateAdd("d", 0, objRecordSet("NOMIDate"))
			ObjExcel.Cells(excel_row, Confirmation2_col).value =  objRecordSet("Confirmation2")
			ObjExcel.Cells(excel_row, DenialNeeded_col).value =  objRecordSet("DenialNeeded")
			ObjExcel.Cells(excel_row, NextActionNeeded_col).value =  objRecordSet("NextActionNeeded")
			ObjExcel.Cells(excel_row, AddedtoWorkList_col).value =  DateAdd("d", 0, objRecordSet("AddedtoWorkList"))
			ObjExcel.Cells(excel_row, SecondApplicationDateNotes_col).value =  objRecordSet("SecondApplicationDateNotes")
			ObjExcel.Cells(excel_row, ClosedInPast30Days_col).value =  objRecordSet("ClosedInPast30Days")
			ObjExcel.Cells(excel_row, ClosedInPast30DaysNotes_col).value =  objRecordSet("ClosedInPast30DaysNotes")
			ObjExcel.Cells(excel_row, StartedOutOfCounty_col).value =  objRecordSet("StartedOutOfCounty")
			ObjExcel.Cells(excel_row, StartedOutOfCountyNotes_col).value =  objRecordSet("StartedOutOfCountyNotes")
			ObjExcel.Cells(excel_row, TrackingNotes_col).value = case_review_notes
			ObjExcel.Cells(excel_row, CaseSelectedTime_col).value = assigned_start_time
			ObjExcel.Cells(excel_row, Hold1Start_col).value = assigned_hold_1_start_time
			ObjExcel.Cells(excel_row, Hold1End_col).value = assigned_hold_1_end_time
			ObjExcel.Cells(excel_row, Hold2Start_col).value = assigned_hold_2_start_time
			ObjExcel.Cells(excel_row, Hold2End_col).value = assigned_hold_2_end_time
			ObjExcel.Cells(excel_row, Hold3Start_col).value = assigned_hold_3_start_time
			ObjExcel.Cells(excel_row, Hold3End_col).value = assigned_hold_3_end_time
			ObjExcel.Cells(excel_row, CaseCompletedTime_col).value = assigned_end_time

			'update the information on SQL table with the notes updates
			If local_demo = False Then
				'Creating objects for Access
				Set objUpdateConnection = CreateObject("ADODB.Connection")
				Set objUpdateRecordSet = CreateObject("ADODB.Recordset")

				'This is the BZST connection to SQL Database'
				objUpdateConnection.Open db_full_string

				'delete a record if the case number matches
				objUpdateRecordSet.Open "UPDATE ES.ES_OnDemanCashAndSnapBZProcessed SET TrackingNotes = '" & case_review_notes & "' WHERE CaseNumber = '" & MAXIS_case_number & "'", objUpdateConnection

				'close the connection and recordset objects to free up resources
				objUpdateConnection.Close
				Set objUpdateRecordSet=nothing
				Set objUpdateConnection=nothing
			End If

			'removing the tracking cookie
			With objFSO
				If .FileExists(od_revw_tracking_file_path) = True then
					.DeleteFile(od_revw_tracking_file_path)
				End If
			End With

            excel_row = excel_row + 1
            cases_completed_by_me = cases_completed_by_me + 1
        End If
        objRecordSet.MoveNext
    Loop

    'close the connection and recordset objects to free up resources
    objRecordSet.Close
    objConnection.Close
    Set objRecordSet=nothing
    Set objConnection=nothing

    objWorkbook.Save()		'saving the excel
    ObjExcel.ActiveWorkbook.Close

    ObjExcel.Application.Quit
    ObjExcel.Quit

	' MsgBox "All case details should now be saved in the Excel and the file should be closed." & vbCr & "Excel last line: " & excel_row & vbCr & vbCr & "TABLE information should also be updated."
	'creating the email to report completion of work
	main_email_subject = "On Demand Daily Case Review Completed"

    main_email_body = "The On Demand Appplication Reviews have been completed for " & date & "."
    main_email_body = main_email_body & vbCr & "Completed by: " & assigned_worker
    main_email_body = main_email_body & vbCr & vbCr & "Number of cases reviewed: " & cases_completed_by_me
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
	' cc_email = "HSPH.EWS.BlueZoneScripts@hennepin.us"
	cc_email = qi_worker_email

	''sending the email
    Call create_outlook_email("", qi_worker_supervisor_email, cc_email, "", main_email_subject, 1, False, "", "", False, "", main_email_body, False, "", True)

	' MsgBox "Now the Email should have been sent and you should have a copy." & vbCr & "qi_worker_supervisor_email - " & qi_worker_supervisor_email & vbCr & "cc_email - " & cc_email

	'this part will review the cookie folder to remove any that are more than a week old. This is a clean up effort
	Set objFolder = objFSO.GetFolder(current_day_work_tracking_folder)							'Creates an oject of the whole my documents folder
	Set colFiles = objFolder.Files																'Creates an array/collection of all the files in the folder
	For Each objFile in colFiles																'looping through each file
		delete_this_file = False																'Default to NOT delete the file
		this_file_type = objFile.Type															'Grabing the file type
		this_file_created_date = objFile.DateCreated											'Reading the date created
		this_file_path = objFile.Path															'Grabing the path for the file

		If this_file_type <> "Text Document" then delete_this_file = False						'We do NOT want to delete files that are NOT TXT file types
		If DateDiff("d", this_file_created_date, date) > 7 Then delete_this_file = True		'We do NOT want to delete files that are 7 days old or less - we may need to reference the saved work in these files.

		If delete_this_file = True Then objFSO.DeleteFile(this_file_path)						'If we have determined that we need to delete the file - here we delete it
	Next
end function

'function to create a txt file to save information about the work as a case is reviewed
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

'this is a cleanup functionality to move the previous worklist into SQL
function merge_worklist_to_SQL()
	continue_msg = MSGBOX ("THIS WILL DELETE THE SQL TABLE", vbYesNo + 48 + 256, "DELETE?")
	If continue_msg = vbNo then StopScript
	'Creating objects for Access
	Set objWorkConnection = CreateObject("ADODB.Connection")
	Set objWorkRecordSet = CreateObject("ADODB.Recordset")

	'This is the file path for the statistics Access database.
	objWorkConnection.Open db_full_string
	' objWorkRecordSet.Open objWorkSQL, objWorkConnection

	objWorkRecordSet.Open "DELETE FROM ES.ES_OnDemanCashAndSnapBZProcessed", objWorkConnection, 3, 3

	' objWorkRecordSet.Close
	objWorkConnection.Close

	Set objWorkRecordSet=nothing
	Set objWorkConnection=nothing
	Set objWorkSQL=nothing

    'setting up information and variables for accessing yesterday's worklist
	previous_list_file_selection_path = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\On Demand Waiver\QI On Demand Daily Assignment\Archive\Old QI Worklist.xlsx"
    Call File_Exists(previous_list_file_selection_path, does_file_exist)
    previous_worksheet_header = "Work List"

	yesterday_case_list = 0

	If does_file_exist = True Then
		'open the file
		call excel_open(previous_list_file_selection_path, True, False, ObjYestExcel, objYestWorkbook)

		objYestWorkbook.worksheets(previous_worksheet_header).Activate

        'Pull info into a NEW array of prevvious day work.
		xl_row = 2
		Do
			this_case = trim(ObjYestExcel.Cells(xl_row, 2).Value)
			the_name =  trim(ObjYestExcel.Cells(xl_row, 3).Value)
			the_snap_status =  trim(ObjYestExcel.Cells(xl_row, 4).Value)
			the_cash_status =  trim(ObjYestExcel.Cells(xl_row, 5).Value)
			the_pnd2_days =  trim(ObjYestExcel.Cells(xl_row, 6).Value)
			the_app_date =  trim(ObjYestExcel.Cells(xl_row, 7).Value)
			If this_case <> "" Then

                'Creating objects for Access
                Set objConnection = CreateObject("ADODB.Connection")
                Set objRecordSet = CreateObject("ADODB.Recordset")

                'This is the BZST connection to SQL Database'
                objConnection.Open db_full_string

                'delete a record if the case number matches
                ' objRecordSet.Open "INSERT INTO ES.ES_OnDemanCashAndSnapBZProcessed SET TrackingNotes = '" & worklist_notes & "' WHERE CaseNumber = '" & yesterday_list_case_number & "'", objConnection
				objRecordSet.Open "INSERT INTO ES.ES_OnDemanCashAndSnapBZProcessed (CaseNumber, CaseName, ApplDate, SnapStatus, CashStatus, REPT_PND2Days)" & _
								  "VALUES ('" & this_case &  "', '" & _
										the_name &  "', '" & _
										the_app_date &  "', '" & _
										the_snap_status &  "', '" & _
										the_cash_status &  "', '" & _
										the_pnd2_days & "')", objConnection, 3, 3


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

'this function will pull details from the tracking cookie, this has time specifics and notes from a case on hold
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

'this reads all the information from SQL and saves it to reserved variables
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
	assigned_case_is_priv					= False

    case_review_notes = replace(assigned_tracking_notes, "STS-NR", "")
    case_review_notes = replace(case_review_notes, "STS-RC-"&user_ID_for_validation, "")
    case_review_notes = replace(case_review_notes, "STS-IP-"&user_ID_for_validation, "")
    case_review_notes = replace(case_review_notes, "STS-HD-"&user_ID_for_validation, "")
    case_review_notes = replace(case_review_notes, "STS-RC", "")
    case_review_notes = replace(case_review_notes, "STS-NL", "")
    case_review_notes = trim(case_review_notes)

	If Instr(case_review_notes,"PRIVILEGED CASE.") <> 0 Then
		' case_review_notes = replace(case_review_notes, "PRIVILEGED CASE.", "")
		assigned_case_is_priv = True
	End If

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

'function to make sure the worker running the script has access to the SQL table
function test_sql_access()
    'Access the pending cases TABLE - ES_OnDemandCashAndSnap'
    'declare the SQL statement that will query the database
    objSQL = "SELECT * FROM ES.ES_OnDemandCashAndSnap"

    'Creating objects for Access
    Set objConnection = CreateObject("ADODB.Connection")
    Set objRecordSet = CreateObject("ADODB.Recordset")

    'This is the file path for the statistics Access database.
    objConnection.Open db_full_string
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
    objWorkConnection.Open db_full_string
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

    Call create_outlook_email("", "hsph.ews.bluezonescripts@hennepin.us", "", "", email_subject, 1, False, "", "", False, "", email_body, False, "", True)

    end_msg = "The test is complete an an email has been sent to the BlueZone Script Team regarding access to the tables for the new On Demand functionality."
    Call script_end_procedure(end_msg)
end function

'this saves the information into the txt cookie, it has a passthrough variable the sets certain options for where in the process the case is at.
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
If user_ID_for_validation = "CALO001" OR user_ID_for_validation = "ILFE001" OR user_ID_for_validation = "MEGE001" OR user_ID_for_validation = "MARI001" OR user_ID_for_validation = "DACO003" Then
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
    If clean_up_list = "Yes" Then
		MsgBox "THIS WILL CLEAN THE SQL TABLE AND DELETE IT."
		Call merge_worklist_to_SQL
	End If

    If local_demo = False Then ADMIN_run = True
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
		qi_worker_email = tester_array(tester).tester_email
		If tester_array(tester).tester_population = "QI" Then qi_member_identified = True
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
    objConnection.Open db_full_string
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

	'reading all of the cases on the SQL table to look for the status of each case, reviews and assignmments
    cases_on_hold_count = 0						'start at 0
    If BULK_Run_completed = True Then			'we only review the SQL list if the BULK run has been completed.
        'Read the whole table
		review_work_started = False
		all_cases_on_hold = 0
		all_cases_in_progress = 0
        Do While NOT objRecordSet.Eof			'this is the loop
            'count all of today's cases using added to worklist
            case_worklist_date = objRecordSet("AddedtoWorkList")			'reading the date the case was last added to the work list and make it a date
            case_worklist_date = DateAdd("d", 0, case_worklist_date)
            If DateDiff("d", case_worklist_date, date) = 0 Then total_cases_for_review = total_cases_for_review + 1		'if the worklist date is today

			case_tracking_notes = objRecordSet("TrackingNotes")				'reading the notes from the SQL table as this where case status informaiton is held

            'count completed reviews using info in tracking notes and for cases that have been completed and recorded in the log
            If Instr(case_tracking_notes, "STS-RC") <> 0 Then
				cases_with_review_completed =cases_with_review_completed + 1
				If Instr(case_tracking_notes, user_ID_for_validation) <> 0 Then cases_completed_by_current_worker = cases_completed_by_current_worker + 1
			End If
			If DateDiff("d", case_worklist_date, date) = 0 AND Instr(case_tracking_notes, "STS") = 0 Then cases_with_review_completed =cases_with_review_completed + 1

            'count cases that are waiting for review using info in tracking notes
            If Instr(case_tracking_notes, "STS-NR") <> 0 Then cases_waiting_for_review =cases_waiting_for_review + 1

            'count cases on hold
            If Instr(case_tracking_notes, "STS-HD") Then
				all_cases_on_hold = all_cases_on_hold + 1
                If Instr(case_tracking_notes, user_ID_for_validation) <> 0 Then
                    cases_on_hold = cases_on_hold + 1
                    If Instr(case_tracking_notes, user_ID_for_validation) <> 0 Then
                        ReDim preserve CASES_ON_HOLD_ARRAY(last_const, cases_on_hold_count)						'save to an array of hold cases
                        CASES_ON_HOLD_ARRAY(case_nbr_const, cases_on_hold_count) = objRecordSet("CaseNumber")
                        CASES_ON_HOLD_ARRAY(case_notes_const, cases_on_hold_count) = objRecordSet("TrackingNotes")		'remove status information from the notes in the array so they do not display
                        CASES_ON_HOLD_ARRAY(case_notes_const, cases_on_hold_count) = replace(CASES_ON_HOLD_ARRAY(case_notes_const, cases_on_hold_count), "STS-RC", "")
                        CASES_ON_HOLD_ARRAY(case_notes_const, cases_on_hold_count) = replace(CASES_ON_HOLD_ARRAY(case_notes_const, cases_on_hold_count), "STS-RC-"&user_ID_for_validation, "")
                        CASES_ON_HOLD_ARRAY(case_notes_const, cases_on_hold_count) = replace(CASES_ON_HOLD_ARRAY(case_notes_const, cases_on_hold_count), "STS-IP-"&user_ID_for_validation, "")
                        CASES_ON_HOLD_ARRAY(case_notes_const, cases_on_hold_count) = replace(CASES_ON_HOLD_ARRAY(case_notes_const, cases_on_hold_count), "STS-HD-"&user_ID_for_validation, "")
                        CASES_ON_HOLD_ARRAY(case_notes_const, cases_on_hold_count) = trim(CASES_ON_HOLD_ARRAY(case_notes_const, cases_on_hold_count))
                        cases_on_hold_count = cases_on_hold_count + 1
                    End If
                End If
            End If

            'find if there is a case in progress
            If Instr(case_tracking_notes, "STS-IP") Then
				all_cases_in_progress = all_cases_in_progress + 1
                If Instr(case_tracking_notes, user_ID_for_validation) <> 0 Then	'and assigned to the worker running the script
                    worker_on_task = True
                    case_nbr_in_progress = objRecordSet("CaseNumber")
                    set_variables_from_SQL										'this is a function to read information from the SQL table
                    'NOTE - if another worker has a case in progress, it will not show for the current worker, but it will show in the ADMIN function
                End If
            End If


            objRecordSet.MoveNext		'going to the next case
        Loop
		'identifying if the worker has already pulled a case for review or not
		If cases_on_hold = 0 and cases_completed_by_current_worker = 0 and worker_on_task = False Then workers_first_task_pulled_for_review = True
		'using the counts to determine if work has been started
		If cases_completed_by_current_worker <> 0 Then review_work_started = True
		If all_cases_on_hold <> 0 Then review_work_started = True
		If all_cases_in_progress <> 0 Then review_work_started = True
	Else 'BULK run has not been completed for the day
        Do While NOT objRecordSet.Eof			'this is the loop
			case_tracking_notes = objRecordSet("TrackingNotes")				'reading the notes from the SQL table as this where case status informaiton is held
            If Instr(case_tracking_notes, "STS-RC") <> 0 Then finish_day_completed_yesterday = False
            If Instr(case_tracking_notes, "STS-HD") <> 0 Then finish_day_completed_yesterday = False
            If Instr(case_tracking_notes, "STS-IP") <> 0 Then finish_day_completed_yesterday = False
            objRecordSet.MoveNext		'going to the next case
        Loop
    End If

    'close the connection and recordset objects to free up resources
    objRecordSet.Close
    objConnection.Close
    Set objRecordSet=nothing
    Set objConnection=nothing

    'If we are running as an ADMIN, the script needs to collect some additional information about the status of the cases.
	If ADMIN_run = True Then
		'Access the SQL Table
		'declare the SQL statement that will query the database
		objSQL = "SELECT * FROM ES.ES_OnDemanCashAndSnapBZProcessed"

		'Creating objects for Access
		Set objConnection = CreateObject("ADODB.Connection")
		Set objRecordSet = CreateObject("ADODB.Recordset")

		'This is the file path for the statistics Access database.
		objConnection.Open db_full_string
		objRecordSet.Open objSQL, objConnection

		Do While NOT objRecordSet.Eof			'this is the loop
			RC_id_start = ""			'default variables to start
			HD_id_start = ""
			IP_id_start = ""
			worker_id_RC = ""
			worker_id_HD = ""
			worker_id_IP = ""
            case_worklist_date = objRecordSet("AddedtoWorkList")			'reading the date the case was last added to the work list and make it a date
            case_tracking_notes = objRecordSet("TrackingNotes")				'reading the notes from the SQL table as this where case status informaiton is held
            If DateDiff("d", case_worklist_date, first_item_date) = 0 Then admin_cases_for_review = admin_cases_for_review+ 1
			If Instr(case_tracking_notes, "STS-RC") <> 0 Then		'reviews completed
				admin_count_RC = admin_count_RC + 1					'counting all cases with review completed
				RC_id_start = Instr(case_tracking_notes, "STS-RC") + 7		'finding the where the id information is positioned in the notes
				worker_id_RC = MID(case_tracking_notes, RC_id_start, 7)		'reading the worker id from the notes
				worker_id_RC = trim(worker_id_RC)							'the worker id lengths are different - trim to remove spaces
				'If the worker ID has not already been read on another case, it will add it the worker id to a list of all workers with completed reviews
				If InStr(ADMIN_list_workers_RC, "~" & worker_id_RC & "~") = 0 then ADMIN_list_workers_RC = ADMIN_list_workers_RC & worker_id_RC & "~"
			End If
			If Instr(case_tracking_notes, "STS-NR") <> 0 Then admin_count_NR = admin_count_NR + 1		'counting all the cases that still need a review
			If Instr(case_tracking_notes, "STS-HD") <> 0 Then		'reviews on hold
				admin_count_HD = admin_count_HD + 1					'counting all cases that are on hold
				HD_id_start = Instr(case_tracking_notes, "STS-HD") + 7		'finding the where the id information is positioned in the notes
				worker_id_HD = MID(case_tracking_notes, HD_id_start, 7)		'reading the worker id from the notes
				worker_id_HD = trim(worker_id_HD)							'the worker id lengths are different - trim to remove spaces
				'If the worker ID has not already been read on another case, it will add it the worker id to a list of all workers with cases on hold
				If InStr(ADMIN_list_workers_HD, "~" & worker_id_HD & "~") = 0 then ADMIN_list_workers_HD = ADMIN_list_workers_HD & worker_id_HD & "~"
			End If
			If Instr(case_tracking_notes, "STS-IP") <> 0 Then		'reviews in progress
				admin_count_IP = admin_count_IP + 1					'counting all cases that are in progress
				IP_id_start = Instr(case_tracking_notes, "STS-IP") + 7		'finding the where the id information is positioned in the notes
				worker_id_IP = MID(case_tracking_notes, IP_id_start, 7)		'reading the worker id from the notes
				worker_id_IP = trim(worker_id_IP)							'the worker id lengths are different - trim to remove spaces
				'If the worker ID has not already been read on another case, it will add it the worker id to a list of all workers with cases in progress
				If InStr(ADMIN_list_workers_IP, "~" & worker_id_IP & "~") = 0 then ADMIN_list_workers_IP = ADMIN_list_workers_IP & worker_id_IP & "~"
			End If
            objRecordSet.MoveNext		'going to the next case
        Loop
	    'close the connection and recordset objects to free up resources
		objRecordSet.Close
		objConnection.Close
		Set objRecordSet=nothing
		Set objConnection=nothing
	End If

    'If we are running ADMIN functionality we need to get the case information and worker to add to arrays that are easier to work with in dialogs
    worker_count = 0
    RCcount = 0
    IPCount = 0
    HDCount = 0
    If ADMIN_run = True Then
        If len(ADMIN_list_workers_RC) > 1 Then
            ADMIN_list_workers_RC = right(ADMIN_list_workers_RC, len(ADMIN_list_workers_RC)-1)					'formatting the list
            ADMIN_list_workers_RC = left(ADMIN_list_workers_RC, len(ADMIN_list_workers_RC)-1)
            If InStr(ADMIN_list_workers_RC, "~") <> 0 Then RC_id_ARRAY = split(ADMIN_list_workers_RC, "~")		'making the list an array
            If InStr(ADMIN_list_workers_RC, "~") = 0 Then RC_id_ARRAY = array(ADMIN_list_workers_RC)
            For each RC_worker_Id in RC_id_ARRAY																'going through each completed worker and adding details to a main array
                For tester = 0 to UBound(tester_array)                         'looping through all of the testers
                    ' pulling QI members by supervisor from the Complete List of Testers
                    If tester_array(tester).tester_id_number = RC_worker_Id Then
                        ReDim Preserve ADMIN_worker_list_array(admin_wrkr_last_const, worker_count)

                        ADMIN_worker_list_array(wrkr_id_const, worker_count) = RC_worker_Id
                        ADMIN_worker_list_array(wrkr_name_const, worker_count) = tester_array(tester).tester_first_name
                        ADMIN_worker_list_array(case_status_const, worker_count) = "RC"
                        ADMIN_worker_list_array(admin_radio_btn_const, worker_count) = RCcount
						ADMIN_worker_list_array(case_count_const, worker_count) = 0

						worker_count = worker_count + 1
                        RCcount = RCcount + 1
                    End If
                Next
            Next
        End If
        If len(ADMIN_list_workers_HD) > 1 Then
            ADMIN_list_workers_HD = right(ADMIN_list_workers_HD, len(ADMIN_list_workers_HD)-1)					'formatting the list
            ADMIN_list_workers_HD = left(ADMIN_list_workers_HD, len(ADMIN_list_workers_HD)-1)
            If InStr(ADMIN_list_workers_HD, "~") <> 0 Then HD_id_ARRAY = split(ADMIN_list_workers_HD, "~")		'making the list an array
            If InStr(ADMIN_list_workers_HD, "~") = 0 Then HD_id_ARRAY = array(ADMIN_list_workers_HD)
            For each HD_worker_Id in HD_id_ARRAY																'going through each completed worker and adding details to a main array
                For tester = 0 to UBound(tester_array)                         'looping through all of the testers
                    ' pulling QI members by supervisor from the Complete List of Testers
                    If tester_array(tester).tester_id_number = HD_worker_Id Then
                        ReDim Preserve ADMIN_worker_list_array(admin_wrkr_last_const, worker_count)

                        ADMIN_worker_list_array(wrkr_id_const, worker_count) = HD_worker_Id
                        ADMIN_worker_list_array(wrkr_name_const, worker_count) = tester_array(tester).tester_first_name
                        ADMIN_worker_list_array(case_status_const, worker_count) = "HD"
                        ADMIN_worker_list_array(admin_radio_btn_const, worker_count) = HDCount
						ADMIN_worker_list_array(case_count_const, worker_count) = 0

                        worker_count = worker_count + 1
                        HDCount = HDCount + 1
                    End If
                Next
            Next
        End If
        If len(ADMIN_list_workers_IP) > 1 Then
            ADMIN_list_workers_IP = right(ADMIN_list_workers_IP, len(ADMIN_list_workers_IP)-1)					'formatting the list
            ADMIN_list_workers_IP = left(ADMIN_list_workers_IP, len(ADMIN_list_workers_IP)-1)
            If InStr(ADMIN_list_workers_IP, "~") <> 0 Then IP_id_ARRAY = split(ADMIN_list_workers_IP, "~")		'making the list an array
            If InStr(ADMIN_list_workers_IP, "~") = 0 Then IP_id_ARRAY = array(ADMIN_list_workers_IP)
            For each IP_worker_Id in IP_id_ARRAY																'going through each completed worker and adding details to a main array
                For tester = 0 to UBound(tester_array)                         'looping through all of the testers
                    ' pulling QI members by supervisor from the Complete List of Testers
                    If tester_array(tester).tester_id_number = IP_worker_Id Then
                        ReDim Preserve ADMIN_worker_list_array(admin_wrkr_last_const, worker_count)

                        ADMIN_worker_list_array(wrkr_id_const, worker_count) = IP_worker_Id
                        ADMIN_worker_list_array(wrkr_name_const, worker_count) = tester_array(tester).tester_first_name
                        ADMIN_worker_list_array(case_status_const, worker_count) = "IP"
                        ADMIN_worker_list_array(admin_radio_btn_const, worker_count) = IPCount
						ADMIN_worker_list_array(case_count_const, worker_count) = 1

                        worker_count = worker_count + 1
                        IPCount = IPCount + 1
                    End If
                Next
            Next
        End If

		'Read the whole table
		'declare the SQL statement that will query the database
		objSQL = "SELECT * FROM ES.ES_OnDemanCashAndSnapBZProcessed"

		'Creating objects for Access
		Set objConnection = CreateObject("ADODB.Connection")
		Set objRecordSet = CreateObject("ADODB.Recordset")

		'This is the file path for the statistics Access database.
		objConnection.Open db_full_string
		objRecordSet.Open objSQL, objConnection

		Do While NOT objRecordSet.Eof
			For qi_worker = 0 to UBound(ADMIN_worker_list_array, 2)
				case_info_notes = objRecordSet("TrackingNotes")
				If Instr(case_info_notes, "STS-RC") <> 0 AND ADMIN_worker_list_array(case_status_const, qi_worker) = "RC" Then
					If InStr(case_info_notes, ADMIN_worker_list_array(wrkr_id_const, qi_worker))<> 0 Then ADMIN_worker_list_array(case_count_const, qi_worker) = ADMIN_worker_list_array(case_count_const, qi_worker) + 1
				End if
				If Instr(case_info_notes, "STS-HD") <> 0 AND ADMIN_worker_list_array(case_status_const, qi_worker) = "HD" Then
					If InStr(case_info_notes, ADMIN_worker_list_array(wrkr_id_const, qi_worker))<> 0 Then ADMIN_worker_list_array(case_count_const, qi_worker) = ADMIN_worker_list_array(case_count_const, qi_worker) + 1
				End If
			Next
			objRecordSet.MoveNext
		Loop
		'close the connection and recordset objects to free up resources
		objRecordSet.Close
		objConnection.Close
		Set objRecordSet=nothing
		Set objConnection=nothing
    End If
Else                            'if we are running in DEMO mode, we don't read the table - we have a dialog to select the process to view.
	total_cases_for_review = 124
	review_work_started = False
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
      Text 90, 5, 95, 10, "Total Cases to Review: " & total_cases_for_review
    EndDialog

    dialog Dialog1
    cancel_without_confirmation

    If trim(cases_on_hold) = "" Then cases_on_hold = 0                                  'formatting some of the information from this dialog.
    If trim(reviews_complete) = "" Then reviews_complete = 0

    cases_with_review_completed = reviews_complete * 1
    cases_on_hold = cases_on_hold * 1
	If reviews_complete <> 0 Then review_work_started = True
	If cases_on_hold <> 0 Then review_work_started = True

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

	workers_first_task_pulled_for_review = True
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

'If the work from yesterday has anything in progress, on hold, or listed as review completed, the script will stop and advise worker to get support from supervisor
If BULK_Run_completed = False and finish_day_completed_yesterday = False and ADMIN_run = False Then
	end_msg = "The On Demand Dashboard cannot be accessed as work from yesterday was not finished."
	end_msg = end_msg & vbCr & vbCr & "The worklist still has cases indicated that are either:"
	end_msg = end_msg & vbCr & " - In progress"
	end_msg = end_msg & vbCr & " - On Hold"
	end_msg = end_msg & vbCr & " - Review Completed and 'Finish Day' has not been run"
	end_msg = end_msg & vbCr & vbCr & "In order to resolve this issue and run the On Demand Dashboard, you will need to contact Tanya Payne (or her coverage) to have these statuses cleared or the 'Finish Day' functionality for a different day run."
	end_msg = end_msg & vbCr & vbCr & "The script will now end"
	call script_end_procedure_with_error_report(end_msg)
End If

'Here is where the script will decide which dialog to display in the process step for the day.
If BULK_Run_completed = False Then                  'if the main run has not happened yet, we start here.
    Do
        Do
            Dialog1 = ""
            BeginDialog Dialog1, 0, 0, 451, 155, "On Demand Applications Dashboard"
                EditBox 500, 300, 50, 15, fake_edit_box
                ButtonGroup ButtonPressed
                If finish_day_completed_yesterday = False then Text 310, 60, 125, 10, "BULK RUN CANNOT BE STARTED"
				If finish_day_completed_yesterday = True then PushButton 310, 55, 125, 15, "Start On Demand BULK Run", complete_bulk_run_btn
                PushButton 50, 105, 170, 13, "More information about the BULK Run", bulk_run_details_btn
				PushButton 230, 105, 150, 13, "Script Instructions", script_instructions_btn
                PushButton 375, 5, 65, 15, "Test Access", test_access_btn
                If ADMIN_run = True Then PushButton 10, 135, 70, 15, "Admin Functions", admin_btn
                ' OkButton 335, 130, 50, 15
                CancelButton 390, 135, 50, 15
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
            If ButtonPressed = bulk_run_details_btn Then MsgBox bulk_run_information_msg
			If ButtonPressed = script_instructions_btn Then call word_doc_open(script_instructions_file_path, objWord, objDoc)
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

If worker_on_task = False Then			'if the worker is currently NOT on a task, the script shows a dialog to select a case to review
	Do
		Do
			Dialog1 = ""
			'setting the dialog sizes
			grp_len = 45
			dlg_len = 210
			If review_work_started = False then dlg_len = 220
			If cases_on_hold <> 0 Then
				grp_len = 50 + (UBound(CASES_ON_HOLD_ARRAY, 2)+1) * 10      '85'
				dlg_len = 215 + (UBound(CASES_ON_HOLD_ARRAY, 2)+1) * 10     '245'
			End If
			If review_work_started = True then first_grp_len = 75
			If review_work_started = False then first_grp_len = 85
			y_pos = first_grp_len + 25

			BeginDialog Dialog1, 0, 0, 451, dlg_len, "On Demand Applications Dashboard"
				ButtonGroup ButtonPressed
				EditBox 500, 600, 50, 15, fake_edit_box
				Text 170, 10, 135, 10, "On Demand Applications Dashboard"
				GroupBox 10, 20, 430, first_grp_len, "Applications BULK Run"
				Text 20, 35, 170, 10, "The BULK run was last completed on " & first_item_date &" ."
				Text 190, 35, 240, 10, "The BULK Run was completed today around " & first_item_time & "."
				Text 20, 50, 305, 10, "The BULK Run can only be completed once per day. The Work List is ready to be reviewed."
				Text 40, 60, 305, 10, "- The worklist is held in a SQL Table and can only be viewed through this Dashboard script. "
				Text 40, 70, 245, 10, "- Use this script to pull a case from the Work List and complete the review."
				If review_work_started = False then
					Text 140, 90, 190, 10, "If the BULK Run was not completed, you can restart here:"
					PushButton 330, 85, 105, 15, "Restart the BULK Run", bulk_run_incomplete_btn
				End If
				GroupBox 10, y_pos, 430, 35, "Work List Overview"
				y_pos = y_pos + 10
				Text 20, y_pos, 115, 10, "Total cases on the worklist: " & total_cases_for_review
				Text 220, y_pos, 130, 10, "Cases with Review Completed: " & cases_with_review_completed
				worklist_btn_pos = y_pos + 10
				y_pos = y_pos + 10
				Text 215, y_pos, 135, 10, "Cases with Reviews In Progress: " & cases_on_hold
				y_pos = y_pos + 20
				GroupBox 10, y_pos, 430, grp_len, "Reviews"
				y_pos = y_pos + 10
				If total_cases_for_review <> cases_with_review_completed Then
					PushButton 15, y_pos, 110, 15, "Pull a case to Review", get_new_case_btn
				Else
					y_pos = y_pos + 5
					Text 20, y_pos, 155, 10, "*** All Cases have been Pulled for Review ***'"
				End If
				y_pos = y_pos + 20
				If cases_on_hold = 0 Then Text 20, y_pos, 190, 10, "--- There are no cases with reviews already started. ---"
				If cases_on_hold <> 0 Then
					Text 20, y_pos, 125, 10, "Reviews Started and put on Hold:"
					OptionGroup RadioGroup1
						hld_y_pos = first_grp_len + 110
						For fold_case = 0 to UBound(CASES_ON_HOLD_ARRAY, 2)
							RadioButton 30, hld_y_pos, 300, 10, "CASE # " & CASES_ON_HOLD_ARRAY(case_nbr_const, fold_case) & " - " & CASES_ON_HOLD_ARRAY(case_notes_const, fold_case), CASES_ON_HOLD_ARRAY(radio_btn_const, fold_case)
							hld_y_pos = hld_y_pos + 10
						Next
				End If
				If cases_on_hold <> 0 Then PushButton 330, dlg_len - 45, 105, 15, "Resume selected Hold Case", resume_hold_case_btn
				PushButton 375, 20, 65, 15, "Test Access", test_access_btn
				PushButton 285, 70, 140, 13, "More information about the Work List", work_list_details_btn
				PushButton 20, worklist_btn_pos, 110, 12, "Worklist Process Information", worklist_process_doc_btn
				PushButton 10, dlg_len - 20, 105, 15, "Finish Work Day", finish_work_day_btn
				If ADMIN_run = True Then PushButton 120, dlg_len - 20, 70, 15, "Admin Functions", admin_btn
				PushButton 305, dlg_len - 18, 75, 12, "Script Instructions", script_instructions_btn
				CancelButton 390, dlg_len - 20, 50, 15
			EndDialog

			dialog Dialog1
			cancel_without_confirmation

			'each button will run a different functionality.
			If ButtonPressed = work_list_details_btn Then MsgBox worklist_information_msg
			If ButtonPressed = test_access_btn Then Call test_sql_access()
			If ButtonPressed = worklist_process_doc_btn Then call word_doc_open(worklist_instructions_file_path, objWord, objDoc)
			If ButtonPressed = script_instructions_btn Then call word_doc_open(script_instructions_file_path, objWord, objDoc)
			If ButtonPressed = finish_work_day_btn Then
				Call assess_worklist_to_finish_day
				If case_on_hold = False and case_in_progress = False Then
					Call create_assignment_report
					end_msg = "Tracking log has been updated with work completed by " & assigned_worker & "."
                    call script_end_procedure(end_msg)
				Else
					loop_dlg_msg = "You cannot finish the work day with cases in progress or on hold." & vbCr
					loop_dlg_msg = loop_dlg_msg & "The dialog will reappear, finish all reviews that have been started first." & vbCr & vbCr
					loop_dlg_msg = loop_dlg_msg & "Once there are no cases on the worklist on hold or in progress the finish work day functionality will operate."
					ButtonPressed = work_list_details_btn
					MsgBox loop_dlg_msg
				End If
			End If
			If ButtonPressed = admin_btn Then call complete_admin_functions
		Loop until ButtonPressed <> work_list_details_btn and ButtonPressed <> worklist_process_doc_btn and ButtonPressed <> script_instructions_btn		'loop until we haven't hit an info function
		CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
	Loop until are_we_passworded_out = false					'loops until user passwords back in

	If ButtonPressed = get_new_case_btn Then Call assign_a_case			'pulling a function if the button is pressed to get a new case
	If ButtonPressed = bulk_run_incomplete_btn Then						'If the BULK run was incomplete - this will restart it
		Call back_to_SELF
		EMReadScreen MX_region, 10, 22, 48
		MX_region = trim(MX_region)
		If MX_region <> "PRODUCTION" Then Call script_end_procedure("You have selected to complete the BULK Run for On Demand but you are not in production. The script will now end. Move to PRODUCTION and run On Demand Dashboard again.")
		Call run_from_GitHub(script_repository & "admin\" & "on-demand-waiver-applications.vbs")
	End If

	If ButtonPressed = resume_hold_case_btn Then						'setting information for a case on hold and assigning it if set to resume the case
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
End If

If worker_on_task = True Then
    If local_demo = True Then
        MAXIS_case_number                       = "318040"					'these are hard copy set for DEMO cases
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

	txt_file_name = user_ID_for_validation & "_" & MAXIS_case_number & "_" & file_date & ".txt"			'setting the file name for the tracking cookie
    od_revw_tracking_file_path = current_day_work_tracking_folder  & txt_file_name
	call read_tracking_cookie			'this pulls information saved in the tracking cookie - where we store notes about the case reviews

	'this dialog will allow information from the case review to be recorded.
	Do
	    Do
	        err_msg = ""

            Dialog1 = ""
	        BeginDialog Dialog1, 0, 0, 451, 310, "On Demand Applications Case Review"
	          Text 185, 10, 60, 10, "Case in Review"
	          GroupBox 10, 20, 230, 75, "Case Information"
	          Text 20, 35, 85, 10, " Case Number: " & MAXIS_case_number
			  If assigned_case_is_priv = True Then Text 20, 105, 85, 10, "CASE IS PRIVILEGED"
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
				PushButton 200, 255, 115, 13, "Script Instructions", script_instructions_btn
				PushButton 315, 255, 130, 13, "Worklist Process Information", worklist_process_doc_btn
	            CancelButton 395, 290, 50, 15
	        EndDialog

	        dialog Dialog1
	        cancel_confirmation

			'these buttons will call a different functionality
            If ButtonPressed = -1 or ButtonPressed = close_dialog_btn Then script_end_procedure(end_msg)
            If ButtonPressed = test_access_btn Then Call test_sql_access()
            If ButtonPressed = admin_btn Then call complete_admin_functions
			If ButtonPressed = worklist_process_doc_btn Then
				call word_doc_open(worklist_instructions_file_path, objWord, objDoc)
				err_msg = "LOOP"
			End If
        Loop until err_msg = ""
	    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
	Loop until are_we_passworded_out = false					'loops until user passwords back in

	'if the review is indicated as done, we will format some information and save that information to the tracking cookie
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

	'if the case is selected to be put on hold, the information will be saved and recorded to the cookie
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
		case_review_notes = replace(case_review_notes, "'", "")								'remove any single quote from the string because it is a reserved character in SQL
        case_review_notes = "STS-HD-"&user_ID_for_validation & " " & case_review_notes
        Call update_tracking_cookie("HOLD")
	End If

	'if this is not running in a demo, the information will be saved to the SQL table
    If local_demo = False Then
        'Creating objects for Access
        Set objConnection = CreateObject("ADODB.Connection")
        Set objRecordSet = CreateObject("ADODB.Recordset")

        'This is the BZST connection to SQL Database'
        objConnection.Open db_full_string

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

'----------------------------------------------------------------------------------------------------Closing Project Documentation - Version date 03/03/2023
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------03/03/2023
'--Tab orders reviewed & confirmed----------------------------------------------03/03/2023		This is actually really funky with this script because there is some prioritization concerns
'--Mandatory fields all present & Reviewed--------------------------------------03/03/2023
'--All variables in dialog match mandatory fields-------------------------------03/03/2023
'Review dialog names for content and content fit in dialog----------------------03/03/2023
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------N/A				There is not CASE/NOTE in this script
'--CASE:NOTE Header doesn't look funky------------------------------------------N/A
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------N/A
'--write_variable_in_CASE_NOTE function: confirm that proper punctuation is used -----------------------------------N/A
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------03/03/2023
'--MAXIS_background_check reviewed (if applicable)------------------------------N/A
'--PRIV Case handling reviewed -------------------------------------------------N/A
'--Out-of-County handling reviewed----------------------------------------------N/A
'--script_end_procedures (w/ or w/o error messaging)----------------------------03/03/2023
'--BULK - review output of statistics and run time/count (if applicable)--------N/A				This script doesn't speak to time savings
'--All strings for MAXIS entry are uppercase vs. lower case (Ex: "X")-----------03/03/2023
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------N/A				This script doesn't speak to time savings
'--Incrementors reviewed (if necessary)-----------------------------------------N/A				This script doesn't speak to time savings
'--Denomination reviewed -------------------------------------------------------N/A				This script doesn't speak to time savings
'--Script name reviewed---------------------------------------------------------03/03/223
'--BULK - remove 1 incrementor at end of script reviewed------------------------N/A

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------03/03/2023
'--comment Code-----------------------------------------------------------------03/03/2023
'--Update Changelog for release/update------------------------------------------03/03/2023
'--Remove testing message boxes-------------------------------------------------03/03/2023
'--Remove testing code/unnecessary code-----------------------------------------03/03/2023
'--Review/update SharePoint instructions----------------------------------------03/03/2023		Instructions are not held on SharePoint
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------03/03/2023
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------03/03/2023
'--COMPLETE LIST OF SCRIPTS update policy references----------------------------N/A				Supports an internal process, not policy
'--Complete misc. documentation (if applicable)---------------------------------03/03/2023
'--Update project team/issue contact (if applicable)----------------------------03/03/2023