'LOADING GLOBAL VARIABLES
'Find who is running
' Set objNet = CreateObject("WScript.NetWork")                                    'getting the users windows ID
' windows_user_ID = objNet.UserName
' user_ID_for_validation = ucase(windows_user_ID)
'
' Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
' If user_ID_for_validation = "CALO001" OR user_ID_for_validation = "ILFE001" OR user_ID_for_validation = "WFS395"Then
' 	Set fso_command = run_another_script_fso.OpenTextFile("C:\MAXIS-Scripts\locally-installed-files\SETTINGS - GLOBAL VARIABLES.vbs")
' Else
' 	Set fso_command = run_another_script_fso.OpenTextFile("\\hcgg.fr.co.hennepin.mn.us\lobroot\hsph\team\Eligibility Support\Scripts\Script Files\SETTINGS - GLOBAL VARIABLES.vbs")
' End If
' text_from_the_other_script = fso_command.ReadAll
' fso_command.Close
' Execute text_from_the_other_script

' 'LOADING SCRIPT
' script_url = script_repository & "/admin/admin-main-menu.vbs"
' IF run_locally = False THEN
' 	SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a script_URL
' 	req.open "GET", script_URL, FALSE									'Attempts to open the script_URL
' 	req.send													'Sends request
' 	IF req.Status = 200 THEN									'200 means great success
' 		Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
' 		Execute req.responseText								'Executes the script code
' 	ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
' 		MsgBox 	"Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_
' 		vbCr & _
' 		"Before contacting the BlueZone script team at HSPH.EWS.BlueZoneScripts@Hennepin.us, please check to make sure you can load the main page at www.GitHub.com." & vbCr &_
' 		vbCr & _
' 		"If you can reach GitHub.com, but this script still does not work, contact the BlueZone script team at HSPH.EWS.BlueZoneScripts@Hennepin.us and provide the following information:" & vbCr &_
' 		vbTab & "- The name of the script you are running." & vbCr &_
' 		vbTab & "- Whether or not the script is ""erroring out"" for any other users." & vbCr &_
' 		vbTab & "- The URL indicated below (a screenshot should suffice)." & vbCr &_
' 		vbCr & _
' 		"We will work with your IT department to try and solve this issue, if needed." & vbCr &_
' 		vbCr &_
' 		"URL: " & url
' 		StopScript
' 	END IF
' ELSE
' 	Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
' 	Set fso_command = run_another_script_fso.OpenTextFile(script_url)
' 	text_from_the_other_script = fso_command.ReadAll
' 	fso_command.Close
' 	Execute text_from_the_other_script
' END IF

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

function assign_a_case()
    If local_demo = False Then
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

                MAXIS_case_number                       = objRecordSet("CaseNumber")
                ' assigned_ = objRecordSet("CaseNumber")
                assigned_case_name                      = objRecordSet("CaseName")
                assigned_application_date               = objRecordSet("ApplDate")
                assigned_interview_date                 = objRecordSet("InterviewDate")
                assigned_day_30                         = objRecordSet("Day_30")
                assigned_days_pending                   = objRecordSet("DaysPending")
                assigned_snap_status                    = objRecordSet("SnapStatus")
                assigned_cash_status                    = objRecordSet("CashStatus")
                assigned_2nd_application_date           = objRecordSet("SecondApplicationDate")
                assigned_rept_pnd2_days                 = objRecordSet("REPT_PND2Days")
                assigned_questionable_interview         = objRecordSet("QuestionableInterview")
                assigned_questionable_interview_resolve = objRecordSet("Resolved")
                assigned_appt_notice_date               = objRecordSet("ApptNoticeDate")
                assigned_appt_date                      = objRecordSet("ApptDate")
                assigned_appt_notc_confirmation         = objRecordSet("Confirmation")
                assigned_nomi_date                      = objRecordSet("NOMIDate")
                assigned_nomi_confirmation              = objRecordSet("Confirmation2")
                assigned_denial_needed                  = objRecordSet("DenialNeeded")
                assigned_next_action_needed             = objRecordSet("NextActionNeeded")
                assigned_added_to_work_list             = objRecordSet("AddedtoWorkList")
                assigned_2nd_application_date_resolve   = objRecordSet("SecondApplicationDateNotes")
                assigned_closed_recently                = objRecordSet("ClosedInPast30Days")
                assigned_closed_recently_resolve        = objRecordSet("ClosedInPast30DaysNotes")
                assigned_out_of_county                  = objRecordSet("StartedOutOfCounty")
                assigned_out_of_county_resolve          = objRecordSet("StartedOutOfCountyNotes")
                assigned_tracking_notes                 = objRecordSet("TrackingNotes")

                case_review_notes = replace(assigned_tracking_notes, "STS-NR", "")
                ' case_review_notes = replace(case_review_notes, "STS-RC", "")
                ' case_review_notes = replace(case_review_notes, "STS-IP-"&user_ID_for_validation, "")
                ' case_review_notes = replace(case_review_notes, "STS-HD-"&user_ID_for_validation, "")
                ' case_review_notes = replace(case_review_notes, "STS-NL", "")
                case_review_notes = trim(case_review_notes)

                date_zero =  #1/1/2010#
            	If IsDate(assigned_interview_date) = True Then
            		If DateDiff("d", assigned_interview_date, date_zero) > 0 Then assigned_interview_date = ""
            	ElseIf IsDate(assigned_2nd_application_date) = True Then
            		If DateDiff("d", assigned_2nd_application_date, date_zero) > 0 Then assigned_2nd_application_date = ""
            	ElseIf IsDate(assigned_appt_notice_date) = True Then
            		If DateDiff("d", assigned_appt_notice_date, date_zero) > 0 Then assigned_appt_notice_date = ""
            	ElseIf IsDate(assigned_appt_date) = True Then
            		If DateDiff("d", assigned_appt_date, date_zero) > 0 Then assigned_appt_date = ""
            	ElseIf IsDate(assigned_nomi_date) = True Then
            		If DateDiff("d", assigned_nomi_date, date_zero) > 0 Then assigned_nomi_date = ""
            	ElseIf IsDate(assigned_added_to_work_list) = True Then
            		If DateDiff("d", assigned_added_to_work_list, date_zero) > 0 Then assigned_added_to_work_list = ""
            	End If


                Exit Do
            End If
            objRecordSet.MoveNext
        Loop

        'close the connection and recordset objects to free up resources
        objRecordSet.Close
        objConnection.Close
        Set objRecordSet=nothing
        Set objConnection=nothing

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
    assigned_tracking_notes = "STS-IP-"&user_ID_for_validation & " " & case_review_notes
    If assigned_interview_date = "" = "" Then assigned_interview_date = #1/1/1900#
    If assigned_2nd_application_date = "" = "" Then assigned_2nd_application_date = #1/1/1900#
    If assigned_appt_notice_date = "" = "" Then assigned_appt_notice_date = #1/1/1900#
    If assigned_appt_date = "" = "" Then assigned_appt_date = #1/1/1900#
    If assigned_nomi_date = "" = "" Then assigned_nomi_date = #1/1/1900#
    If assigned_added_to_work_list = "" = "" Then assigned_added_to_work_list = #1/1/1900#

    If local_demo = False Then
        WScript.Sleep 10000
        'Creating objects for Access
        Set objConnection = CreateObject("ADODB.Connection")
        Set objRecordSet = CreateObject("ADODB.Recordset")

        'This is the BZST connection to SQL Database'
        objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
        objSQL = "SELECT * FROM ES.ES_OnDemanCashAndSnapBZProcessed"
        objRecordSet.Open objSQL, objConnection

        'delete a record if the case number matches
        objRecordSet.Open "UPDATE ES.ES_OnDemanCashAndSnapBZProcessed SET CaseNumber = '" & MAXIS_case_number & "', " &_
                                                                          "CaseName = '" & assigned_case_name & "', " &_
                                                                          "ApplDate = '" & assigned_application_date & "', " &_
                                                                          "InterviewDate = '" & assigned_interview_date & "', " &_
                                                                          "Day_30 = '" & assigned_day_30 & "', " &_
                                                                          "DaysPending = '" & assigned_days_pending & "', " &_
                                                                          "SnapStatus = '" & assigned_snap_status & "', " &_
                                                                          "CashStatus = '" & assigned_cash_status & "', " &_
                                                                          "SecondApplicationDate = '" & assigned_2nd_application_date & "', " &_
                                                                          "REPT_PND2Days = '" & assigned_rept_pnd2_days & "', " &_
                                                                          "QuestionableInterview = '" & assigned_questionable_interview & "', " &_
                                                                          "Resolved = '" & assigned_questionable_interview_resolve & "', " &_
                                                                          "ApptNoticeDate = '" & assigned_appt_notice_date & "', " &_
                                                                          "ApptDate = '" & assigned_appt_date & "', " &_
                                                                          "Confirmation = '" & assigned_appt_notc_confirmation & "', " &_
                                                                          "NOMIDate = '" & assigned_nomi_date & "', " &_
                                                                          "Confirmation2 = '" & assigned_nomi_confirmation & "', " &_
                                                                          "DenialNeeded = '" & assigned_denial_needed & "', " &_
                                                                          "NextActionNeeded = '" & assigned_next_action_needed & "', " &_
                                                                          "AddedtoWorkList = '" & assigned_added_to_work_list & "', " &_
                                                                          "SecondApplicationDateNotes = '" & assigned_2nd_application_date_resolve & "', " &_
                                                                          "ClosedInPast30Days = '" & assigned_closed_recently & "', " &_
                                                                          "ClosedInPast30DaysNotes = '" & assigned_closed_recently_resolve & "', " &_
                                                                          "StartedOutOfCounty = '" & assigned_out_of_county & "', " &_
                                                                          "StartedOutOfCountyNotes = '" & assigned_out_of_county_resolve & "', " &_
                                                                          "TrackingNotes = '" & assigned_tracking_notes & "' WHERE CaseNumber = '" & MAXIS_case_number & "'", objConnection



        ' objRecordSet.Open "UPDATE ES.ES_OnDemanCashAndSnapBZProcessed SET CaseNumber = '" & MAXIS_case_number & "', CaseName = '" & assigned_case_name & "', ApplDate = '" & assigned_application_date & "', InterviewDate = '" & assigned_interview_date & "', "Day_30" = '" & assigned_day_30 & "', "DaysPending" = '" & assigned_days_pending & "', SnapStatus = '" & assigned_snap_status & "', CashStatus = '" & assigned_cash_status & "', SecondApplicationDate = '" & assigned_2nd_application_date & "', REPT_PND2Days = '" & assigned_rept_pnd2_days & "', QuestionableInterview = '" & assigned_questionable_interview & "', Resolved = '" & assigned_questionable_interview_resolve & "', ApptNoticeDate = '" & assigned_appt_notice_date & "', ApptDate = '" & assigned_appt_date & "', Confirmation = '" & assigned_appt_notc_confirmation & "', NOMIDate = '" & assigned_nomi_date & "', Confirmation2 = '" & assigned_nomi_confirmation & "', DenialNeeded = '" & assigned_denial_needed & "', NextActionNeeded = '" & assigned_next_action_needed & "', AddedtoWorkList = '" & assigned_added_to_work_list & "', SecondApplicationDateNotes = '" & assigned_2nd_application_date_resolve & "', ClosedInPast30Days = '" & assigned_closed_recently & "', ClosedInPast30DaysNotes = '" & assigned_closed_recently_resolve & "', StartedOutOfCounty = '" & assigned_out_of_county & "', StartedOutOfCountyNotes = '" & assigned_out_of_county_resolve & "', TrackingNotes = '" & assigned_tracking_notes & "' WHERE CaseNumber = '" & MAXIS_case_number & "'", objConnection

        'close the connection and recordset objects to free up resources
        objRecordSet.Close
        objConnection.Close
        Set objRecordSet=nothing
        Set objConnection=nothing
    End If

    txt_file_name = user_ID_for_validation & "_" & MAXIS_case_number & "_" & file_date & ".txt"
    od_revw_tracking_file_path = current_day_work_tracking_folder  & txt_file_name
    call update_tracking_cookie

    Call Back_to_SELF
    Call navigate_to_MAXIS_screen("CASE", "CURR")
    Do
        Dialog1 = ""
        BeginDialog Dialog1, 0, 0, 246, 335, "On Demand Applications Case Review"
          Text 95, 10, 60, 10, "Case to Review"
          GroupBox 10, 20, 230, 75, "Case Information"
          Text 20, 35, 85, 10, " Case Number: " & MAXIS_case_number
          Text 30, 45, 210, 10, "Case Name: " & assigned_case_name
          Text 15, 60, 120, 10, "Application Date: " & assigned_application_date
          Text 20, 70, 75, 10, " Days Pending: " & assigned_rept_pnd2_days
          Text 45, 80, 80, 10, "Day 30: " & assigned_day_30
          If assigned_snap_status = "Pending" Then Text 145, 60, 50, 10, "SNAP: Pending"
          If assigned_cash_status = "Pending" Then Text 145, 70, 50, 10, "CASH: Pending"
          GroupBox 10, 100, 230, 30, "Interview"
          ' Text 20, 115, 140, 10, "Interview Date from PROG: " & assigned_interview_date
          If assigned_interview_date = "" Then Text 20, 115, 140, 10, "No interview entered on PROG"
          If assigned_interview_date <> "" Then Text 20, 115, 140, 10, "Interview Date from PROG: " & assigned_interview_date
          GroupBox 10, 135, 230, 55, "Notices"
          If assigned_appt_notice_date = "" Then
            Text 20, 150, 110, 10, "NO APPOINTMENT NOTICE FOUND"
          Else
            Text 20, 150, 110, 10, "Appt Notice Sent on " & assigned_appt_notice_date
            Text 30, 160, 80, 10, "Appt Date: " & assigned_appt_date
          End If
          ' Text 20, 150, 110, 10, "Appt Notice Sent on MM/DD/YY"
          ' Text 30, 160, 80, 10, "Appt Date: MM/DD/YY"
          If assigned_nomi_date = "" Then
            Text 20, 175, 110, 10, "NO NOMI FOUND"
          Else
            Text 20, 175, 110, 10, "NOMI Sent on " & assigned_nomi_date
          End If
          ' Text 20, 175, 110, 10, "NOMI Sent on MM/DD/YY"
          GroupBox 10, 195, 230, 60, "Actions"
          Text 20, 210, 70, 10, "Next Action Needed:"
          Text 25, 220, 165, 10, assigned_next_action_needed
          If assigned_next_action_needed = "RESOLVE SUBSEQUENT APPLICATION DATE" Then Text 20, 235, 145, 10, "Subsequent Appliction Date: " & assigned_2nd_application_date
          If assigned_next_action_needed = "ALIGN INTERVIEW DATES" Then Text 25, 235, 180, 10, "*** Interview Dates on PROG need to be ALIGNED ***"
          If assigned_next_action_needed = "REVIEW QUESTIONABLE INTERVIEW DATE(S)" Then Text 20, 235, 180, 10, "*** Questionable Interview Date Found: " & assigned_questionable_interview & " ***"
          Text 10, 260, 70, 10, "Additional Notes:"
          Text 10, 270, 230, 40, case_review_notes
          EditBox 500, 600, 50, 15, fake_edit_box
          ButtonGroup ButtonPressed
            OkButton 190, 315, 50, 15
        EndDialog

        dialog Dialog1


        CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    Loop until are_we_passworded_out = false					'loops until user passwords back in

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
                    If line_info(0) = "ASSIGNED WORKER" Then saved_worker  = line_info(1)
                    If line_info(0) = "WINDOWS USER ID" Then user_ID_for_validation  = line_info(1)
                    If line_info(0) = "ASSIGNED DATE" Then saved_date  = line_info(1)
                    If line_info(0) = "START TIME" Then saved_start_time  = line_info(1)
                    If line_info(0) = "END TIME" Then saved_end_time  = line_info(1)
                    If line_info(0) = "HOLD 1 START" Then saved_hold_1_start_time  = line_info(1)
                    If line_info(0) = "HOLD 1 END" Then saved_hold_1_end_time  = line_info(1)
                    If line_info(0) = "HOLD 2 START" Then saved_hold_2_start_time  = line_info(1)
                    If line_info(0) = "HOLD 2 END" Then saved_hold_2_end_time  = line_info(1)
                    If line_info(0) = "HOLD 3 START" Then saved_hold_3_start_time  = line_info(1)
                    If line_info(0) = "HOLD 3 END" Then saved_hold_3_end_time  = line_info(1)

                End If
            Next

            ' .DeleteFile(od_revw_tracking_file_path)
            objTextStream.Close
        End If

    End With

end function

function test_sql_access()
    'Access the pending cases TABLE - ES_OnDemandCashAndSnap'
    'declare the SQL statement that will query the database
    objSQL = "SELECT * FROM ES.ES_OnDemandCashAndSnap"

    'Creating objects for Access
    Set objConnection = CreateObject("ADODB.Connection")
    Set objRecordSet = CreateObject("ADODB.Recordset")

    'This is the file path for the statistics Access database.
    ' stats_database_path = "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;"
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

function update_tracking_cookie()

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

            'Close the object so it can be opened again shortly
            objTextStream.Close
        End If

    End With
end function


'BULK NOT COMPLETED TODAY'
BeginDialog Dialog1, 0, 0, 451, 155, "On Demand Applications Dashboard"
  EditBox 500, 300, 50, 15, fake_edit_box
  ButtonGroup ButtonPressed
    PushButton 310, 55, 125, 15, "Start On Demand BULK Run", complete_bulk_run_btn
    PushButton 50, 105, 170, 10, "More information about the BULK Run", bulk_run_details_btn
    OkButton 335, 130, 50, 15
    CancelButton 390, 130, 50, 15
  Text 170, 10, 135, 10, "On Demand Applications Dashboard"
  GroupBox 10, 25, 430, 100, "Applications BULK Run"
  Text 20, 40, 170, 10, "The BULK run was last completed on MM/DD/YY."
  Text 20, 55, 285, 20, "The BULK Run has not been completed today. You must complete the BULK run before any other On Demand work can be completed."
  Text 45, 75, 235, 10, "- The BULK run takes 2 - 3 hour and uses PRODUCTION the entire time."
  Text 45, 85, 260, 10, "- You can complete other work in other sessions while the BULK Run happens."
  Text 45, 95, 325, 10, "- The BULK Run can be unattended (you can walk away) but this is not paid time if you walk away."
EndDialog


'BULK COMPLETED - NO REVIEWS IN HOLD
BeginDialog Dialog1, 0, 0, 451, 235, "On Demand Applications Dashboard"
  EditBox 500, 600, 50, 15, fake_edit_box
  ButtonGroup ButtonPressed
    PushButton 15, 165, 110, 15, "Pull a case to Review", get_new_case_btn
    'Text 20, 170, 155, 10, "*** All Cases have been Pulled for Review ***"'
    PushButton 45, 85, 170, 10, "More information about the Work List", work_list_details_btn
    PushButton 10, 215, 105, 15, "Finish Work Day", finish_work_day_btn
    CancelButton 390, 215, 50, 15
  Text 170, 10, 135, 10, "On Demand Applications Dashboard"
  GroupBox 10, 25, 430, 75, "Applications BULK Run"
  Text 20, 40, 170, 10, "The BULK run was last completed on MM/DD/YY."
  Text 190, 40, 180, 10, "The BULK Run was completed today around HH:MM."
  Text 20, 55, 305, 10, "The BULK Run can only be completed once per day. The Work List is ready to be reviewed."
  Text 40, 65, 305, 10, "- The worklist is held in a SQL Table and can only be viewed through this Dashboard script. "
  Text 40, 75, 245, 10, "- Use this script to pull a case from the Work List and complete the review."
  GroupBox 10, 105, 430, 40, "Work List Overview"
  Text 20, 120, 115, 10, "Total cases on the worklist: XXX"
  Text 220, 120, 130, 10, "Cases with Review Completed: XXX"
  Text 215, 130, 135, 10, "Cases with Reviews In Progress: XXX"
  GroupBox 10, 150, 430, 55, "Reviews"
  Text 20, 190, 190, 10, "--- There are no cases with reviews already started. ---"
EndDialog


'BULK COMPLETED with REVIEWS IN HOLD
BeginDialog Dialog1, 0, 0, 451, 245, "On Demand Applications Dashboard"
  EditBox 500, 600, 50, 15, fake_edit_box
  ButtonGroup ButtonPressed
    PushButton 20, 145, 110, 15, "Pull a case to Review", get_new_case_btn
    'Text 20, 170, 155, 10, "*** All Cases have been Pulled for Review ***"'
  Text 320, 5, 135, 10, "On Demand Applications Dashboard"
  GroupBox 10, 10, 430, 75, "Applications BULK Run"
  Text 20, 25, 170, 10, "The BULK run was last completed on MM/DD/YY."
  Text 190, 25, 180, 10, "The BULK Run was completed today around HH:MM."
  Text 20, 40, 305, 10, "The BULK Run can only be completed once per day. The Work List is ready to be reviewed."
  Text 40, 50, 305, 10, "- The worklist is held in a SQL Table and can only be viewed through this Dashboard script. "
  Text 40, 60, 245, 10, "- Use this script to pull a case from the Work List and complete the review."
  ButtonGroup ButtonPressed
    PushButton 45, 70, 170, 10, "More information about the Work List", work_list_details_btn
  GroupBox 10, 90, 430, 40, "Work List Overview"
  Text 20, 105, 115, 10, "Total cases on the worklist: XXX"
  Text 220, 105, 130, 10, "Cases with Review Completed: XXX"
  Text 215, 115, 135, 10, "Cases with Reviews In Progress: XXX"
  GroupBox 10, 135, 430, 85, "Reviews"
  Text 20, 165, 125, 10, "Reviews Started and put on Hold:"
  OptionGroup RadioGroup1
    RadioButton 30, 175, 120, 10, "CASE # XXXXXXX", case_info_array
    RadioButton 30, 185, 120, 10, "CASE # XXXXXXX", Radio2
    RadioButton 30, 195, 120, 10, "CASE # XXXXXXX", Radio3
    RadioButton 30, 205, 120, 10, "CASE # XXXXXXX", Radio4
  ButtonGroup ButtonPressed
    PushButton 330, 200, 105, 15, "Resume selected Hold Case", resume_hold_case_btn
    PushButton 10, 225, 105, 15, "Finish Work Day", finish_work_day_btn
    CancelButton 390, 225, 50, 15
EndDialog


'CASE INFORMATION'
BeginDialog Dialog1, 0, 0, 451, 310, "On Demand Applications Case Review"
  Text 185, 10, 60, 10, "Case in Review"
  GroupBox 10, 20, 230, 75, "Case Information"
  Text 20, 35, 85, 10, " Case Number: XXXXXX"
  Text 30, 45, 210, 10, "Case Name: THIIS IS A LONG NAME WITH TOO MANY IIIIIS"
  Text 15, 60, 120, 10, "Application Date: MM/DD/YY"
  Text 20, 70, 75, 10, " Days Pending: XX"
  Text 45, 80, 80, 10, "Day 30: MM/DD/YY"
  Text 145, 60, 50, 10, "SNAP: Pending"
  Text 145, 70, 50, 10, "CASH: Pending"
  GroupBox 10, 100, 230, 95, "Interview"
  Text 20, 115, 140, 10, "Interview Date from PROG: MM/DD/YY"
  Text 20, 130, 180, 10, "*** Interview Dates on PROG need to be ALIGNED ***"
  Text 20, 145, 40, 10, "Resolution: "
  EditBox 60, 140, 170, 15, align_interview_dates_resolution
  Text 20, 165, 180, 10, "*** Questionable Interview Date Found: MM/DD/YY ***"
  Text 20, 180, 40, 10, "Resolution: "
  EditBox 60, 175, 170, 15, questionable_interview_datesresolution
  GroupBox 10, 200, 230, 55, "Notices"
  Text 20, 215, 110, 10, "Appt Notice Sent on MM/DD/YY"
  Text 30, 225, 80, 10, "Appt Date: MM/DD/YY"
  Text 20, 240, 110, 10, "NOMI Sent on MM/DD/YY"
  GroupBox 250, 20, 195, 235, "Group7"
  Text 260, 35, 70, 10, "Next Action Needed:"
  Text 265, 45, 165, 10, "NEXT ACTION"
  Text 260, 65, 145, 10, "Subsequent Appliction Date: MM/DD/YY"
  Text 260, 75, 50, 10, "Resolution:"
  EditBox 260, 85, 175, 15, subseuent_application_resolution
  Text 260, 110, 145, 10, "Case closed in past 30 Days"
  Text 260, 120, 50, 10, "Review Notes:"
  EditBox 260, 130, 175, 15, case_recently_closed_resolution
  Text 260, 155, 145, 10, "Case was in Another County"
  Text 260, 165, 50, 10, "Review Notes:"
  EditBox 260, 175, 175, 15, case_in_another_county_resolution
  Text 260, 200, 175, 10, "Case Cannot be Denied"
  Text 260, 210, 175, 10, "REASON"
  Text 260, 220, 50, 10, "Resolution:"
  EditBox 260, 230, 175, 15, cannot_deny_resolution
  Text 10, 260, 70, 10, "Additional Notes:"
  EditBox 10, 270, 435, 15, case_review_notes
  ButtonGroup ButtonPressed
    PushButton 280, 290, 110, 15, "Complete Review", complete_review_btn
    PushButton 335, 5, 110, 15, "Put Case on Hold", hold_case_btn
    CancelButton 395, 290, 50, 15
EndDialog

'ASSIGNMENT OF CASE'
BeginDialog Dialog1, 0, 0, 246, 335, "On Demand Applications Case Review"
  Text 95, 10, 60, 10, "Case to Review"
  GroupBox 10, 20, 230, 75, "Case Information"
  Text 20, 35, 85, 10, " Case Number: XXXXXX"
  Text 30, 45, 210, 10, "Case Name: THIIS IS A LONG NAME WITH TOO MANY IIIIIS"
  Text 15, 60, 120, 10, "Application Date: MM/DD/YY"
  Text 20, 70, 75, 10, " Days Pending: XX"
  Text 45, 80, 80, 10, "Day 30: MM/DD/YY"
  Text 145, 60, 50, 10, "SNAP: Pending"
  Text 145, 70, 50, 10, "CASH: Pending"
  GroupBox 10, 100, 230, 30, "Interview"
  Text 20, 115, 140, 10, "Interview Date from PROG: MM/DD/YY"
  Text 25, 245, 180, 10, "*** Interview Dates on PROG need to be ALIGNED ***"
  Text 20, 245, 180, 10, "*** Questionable Interview Date Found: MM/DD/YY ***"
  GroupBox 10, 135, 230, 55, "Notices"
  Text 20, 150, 110, 10, "Appt Notice Sent on MM/DD/YY"
  Text 30, 160, 80, 10, "Appt Date: MM/DD/YY"
  Text 20, 175, 110, 10, "NOMI Sent on MM/DD/YY"
  GroupBox 10, 195, 230, 60, "Group7"
  Text 20, 210, 70, 10, "Next Action Needed:"
  Text 25, 220, 165, 10, "NEXT ACTION"
  Text 20, 235, 145, 10, "Subsequent Appliction Date: MM/DD/YY"
  Text 10, 260, 70, 10, "Additional Notes:"
  EditBox 500, 600, 50, 15, fake_edit_box
  ButtonGroup ButtonPressed
    OkButton 190, 315, 50, 15
  Text 10, 270, 230, 40, "These are NOTES"
EndDialog

' assigned_date
' assigned_start_time
' assigned_end_time
' assigned_hold_1_start_time
' assigned_hold_1_end_time
' assigned_hold_2_start_time
' assigned_hold_2_end_time
' assigned_hold_3_start_time
' assigned_hold_3_end_time

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


'BUTTONS'
complete_bulk_run_btn   = 1001
bulk_run_details_btn    = 1002
get_new_case_btn        = 2001
work_list_details_btn   = 2002
resume_hold_case_btn    = 2003
complete_review_btn     = 2004
hold_case_btn           = 2005
finish_work_day_btn     = 3001

' STS-NR            'Status - Needs Review'
' STS-RC            'Status - Review Completed'
' STS-IP-WFXXXX     'Status - In Progress - Worker number'
' STS-HD-WFXXXX     'Status - HolD - Worker number'
' STS-NL            'Status - NULL'

BULK_Run_completed = False
worker_on_task = False
total_cases_for_review = 0
cases_with_review_completed = 0
cases_waiting_for_review = 0
cases_on_hold = 0
case_nbr_in_progress = ""

const case_nbr_const    = 0
const radio_btn_const   = 1
const case_notes_const  = 2
const last_const        = 10

Dim CASES_ON_HOLD_ARRAY()
ReDim CASES_ON_HOLD_ARRAY(last_const, 0)

If user_ID_for_validation = "CALO001" or user_ID_for_validation = "ILFE001" Then
    run_in_demo_mode = MsgBox("Do you need to run this script as a DEMO?", vbQuestion + VBYesNo, "Scriptwriter DEMO?")
    If run_in_demo_mode = vbYes Then local_demo = True
End If

EMConnect ""
Call check_for_MAXIS(True)

'TODO - add functionality for QI leadership
'confirm QI Member'
qi_member_identified = False
For tester = 0 to UBound(tester_array)                         'looping through all of the testers
    ' pulling QI members by supervisor from the Complete List of Testers
    If tester_array(tester).tester_id_number = user_ID_for_validation Then
        If tester_array(tester).tester_supervisor_name = "Tanya Payne" Then qi_member_identified = True
        If tester_array(tester).tester_population = "BZ" Then qi_member_identified = True
        assigned_worker = tester_array(tester).tester_full_name
        ' MsgBox "user_ID_for_validation - " & user_ID_for_validation & vbCr & "tester_array(tester).tester_id_number - " & tester_array(tester).tester_id_number & vbCr & "qi_member_identified - " & qi_member_identified
    End If
Next
If qi_member_identified = False Then script_end_procedure("This script can only be operated by a member of core QI due to access restrictions. The script will now end.")

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
    ' MsgBox "first_item_change - " & first_item_change & vbCr & "first_item_date - " & first_item_date & vbCr & "first_item_time - " & first_item_time

    cases_on_hold_count = 0
    'BULK Run has been completed'
    If BULK_Run_completed = True Then


        'Read the whole table
        Do While NOT objRecordSet.Eof
            'count all of today's cases using added to worklist
            case_worklist_date = objRecordSet("AddedtoWorkList")
            case_worklist_date = DateAdd("d", 0, case_worklist_date)
            If DateDiff("d", case_worklist_date, date) = 0 Then
                total_cases_for_review = total_cases_for_review + 1

                case_tracking_notes = objRecordSet("TrackingNotes")

                'count completed reviews using info in tracking notes
                If Instr(case_tracking_notes, "STS-RC") <> 0 Then cases_with_review_completed =cases_with_review_completed + 1

                'count waiting using info in tracking notes
                If Instr(case_tracking_notes, "STS-NR") <> 0 Then cases_waiting_for_review =cases_waiting_for_review + 1

                'count cases on hold
                If Instr(case_tracking_notes, "STS-HD") Then            'TODO - add worker specific holds'
                    If Instr(case_tracking_notes, user_ID_for_validation) <> 0 Then
                        cases_on_hold = cases_on_hold + 1
                        If Instr(case_tracking_notes, user_ID_for_validation) <> 0 Then
                            ReDim preserve CASES_ON_HOLD_ARRAY(last_const, cases_on_hold_count)
                            CASES_ON_HOLD_ARRAY(case_nbr_const, cases_on_hold_count) = objRecordSet("CaseNumber")
                            CASES_ON_HOLD_ARRAY(case_notes_const, cases_on_hold_count) = objRecordSet("TrackingNotes")

                            cases_on_hold_count = cases_on_hold_count + 1
                        End If
                    End If
                End If

                'find if there is a case 'checked out'
                If Instr(case_tracking_notes, "STS-IP") Then
                    If Instr(case_tracking_notes, user_ID_for_validation) <> 0 Then
                        worker_on_task = True
                        case_nbr_in_progress = objRecordSet("CaseNumber")
                        'TODO - read all on task information

                        MAXIS_case_number                       = objRecordSet("CaseNumber")
                        ' assigned_ = objRecordSet("CaseNumber")
                        assigned_case_name                      = objRecordSet("CaseName")
                        assigned_application_date               = objRecordSet("ApplDate")
                        assigned_interview_date                 = objRecordSet("InterviewDate")
                        assigned_day_30                         = objRecordSet("Day_30")
                        assigned_days_pending                   = objRecordSet("DaysPending")
                        assigned_snap_status                    = objRecordSet("SnapStatus")
                        assigned_cash_status                    = objRecordSet("CashStatus")
                        assigned_2nd_application_date           = objRecordSet("SecondApplicationDate")
                        assigned_rept_pnd2_days                 = objRecordSet("REPT_PND2Days")
                        assigned_questionable_interview         = objRecordSet("QuestionableInterview")
                        assigned_questionable_interview_resolve = objRecordSet("Resolved")
                        assigned_appt_notice_date               = objRecordSet("ApptNoticeDate")
                        assigned_appt_date                      = objRecordSet("ApptDate")
                        assigned_appt_notc_confirmation         = objRecordSet("Confirmation")
                        assigned_nomi_date                      = objRecordSet("NOMIDate")
                        assigned_nomi_confirmation              = objRecordSet("Confirmation2")
                        assigned_denial_needed                  = objRecordSet("DenialNeeded")
                        assigned_next_action_needed             = objRecordSet("NextActionNeeded")
                        assigned_added_to_work_list             = objRecordSet("AddedtoWorkList")
                        assigned_2nd_application_date_resolve   = objRecordSet("SecondApplicationDateNotes")
                        assigned_closed_recently                = objRecordSet("ClosedInPast30Days")
                        assigned_closed_recently_resolve        = objRecordSet("ClosedInPast30DaysNotes")
                        assigned_out_of_county                  = objRecordSet("StartedOutOfCounty")
                        assigned_out_of_county_resolve          = objRecordSet("StartedOutOfCountyNotes")
                        assigned_tracking_notes                 = objRecordSet("TrackingNotes")

                        case_review_notes = replace(assigned_tracking_notes, "STS-NR", "")
                    Else
                        'TODO - handling for another worker'
                    End If
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
Else
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

    If trim(cases_on_hold) = "" Then cases_on_hold = 0
    If trim(reviews_complete) = "" Then reviews_complete = 0

    cases_with_review_completed = reviews_complete * 1
    cases_on_hold = cases_on_hold * 1

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

    cases_waiting_for_review = total_cases_for_review - cases_with_review_completed

    If demo_process = "Complete BULK Run" Then
        BULK_Run_completed = False
        worker_on_task = False
    ElseIf demo_process = "Select New Case" Then
        BULK_Run_completed = True
        worker_on_task = False
    ElseIf demo_process = "Case in Progress" Then
        BULK_Run_completed = True
        worker_on_task = True
    End If
End If

If BULK_Run_completed = False Then
    Do
        Do
            Dialog1 = ""
            BeginDialog Dialog1, 0, 0, 451, 155, "On Demand Applications Dashboard"
              EditBox 500, 300, 50, 15, fake_edit_box
              ButtonGroup ButtonPressed
                PushButton 310, 55, 125, 15, "Start On Demand BULK Run", complete_bulk_run_btn
                PushButton 50, 105, 170, 10, "More information about the BULK Run", bulk_run_details_btn
                PushButton 375, 5, 65, 15, "Test Access", test_access_btn
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

            If ButtonPressed = bulk_run_details_btn Then MsgBox "More details will be here" 'TODO - add BULK Run Explanation'
            If ButtonPressed = test_access_btn Then Call test_sql_access()
            If ButtonPressed = complete_bulk_run_btn Then
                If local_demo = True Then call script_end_procedure("The script would now run the the On Demand Applications script")
                If local_demo = False Then
                    Call back_to_SELF
                    EMReadScreen MX_region, 10, 22, 48
                    MX_region = trim(MX_region)
                    If MX_region <> "PRODUCTION" Then Call script_end_procedure("You have selected to complete the BULK Run for On Demand but you are not in production. The script will now end. Move to PRODUCTION and run On Demand Dashboard again.")
                    Call run_from_GitHub(script_repository & "admin\" & "on-demand-waiver-applications.vbs")
                End If
            End If
        Loop
        CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    Loop until are_we_passworded_out = false					'loops until user passwords back in
End If
If local_demo = False and BULK_Run_completed = True Then
    Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 451, 235, "On Demand Applications Dashboard"
      EditBox 500, 600, 50, 15, fake_edit_box
      Text 20, 170, 155, 10, "Review Functionality is not yet supported."
      ButtonGroup ButtonPressed
        PushButton 375, 5, 65, 15, "Test Access", test_access_btn
        CancelButton 390, 215, 50, 15
      Text 170, 10, 135, 10, "On Demand Applications Dashboard"
      GroupBox 10, 25, 430, 75, "Applications BULK Run"
      Text 20, 40, 170, 10, "The BULK run was last completed on " & first_item_date &" ."
      Text 190, 40, 200, 10, "The BULK Run was completed today around " & first_item_time &" ."
      Text 20, 55, 305, 10, "The BULK Run can only be completed once per day. The Work List is ready to be reviewed."
      Text 40, 65, 305, 10, "- The worklist is held in a SQL Table and can only be viewed through this Dashboard script. "
      Text 40, 75, 245, 10, "- Use this script to pull a case from the Work List and complete the review."
      GroupBox 10, 105, 430, 40, "Work List Overview"
      Text 20, 120, 115, 10, "Total cases on the worklist: " & total_cases_for_review
      Text 220, 120, 130, 10, "Cases with Review Completed: " & cases_with_review_completed
      Text 215, 130, 135, 10, "Cases with Reviews In Progress: " & cases_on_hold
      GroupBox 10, 150, 430, 55, "Reviews"
    EndDialog

    dialog Dialog1
    cancel_without_confirmation

    If ButtonPressed = test_access_btn Then Call test_sql_access()
    end_early_msg = "The BULK Run for On Demand appears to have been completed today. If this is not true, contact the BlueZone Script Team." & vbCr & vbCr & "Eventually this script will support additional functionality, handling the processing of the worklist. This funcitonality is not yet ready. The script will now end."
    Call script_end_procedure(end_early_msg)
End If

If worker_on_task = False Then
    If cases_on_hold = 0 Then
        Do
            Do
                Dialog1 = ""
                BeginDialog Dialog1, 0, 0, 451, 235, "On Demand Applications Dashboard"
                  EditBox 500, 600, 50, 15, fake_edit_box
                  ButtonGroup ButtonPressed
                    PushButton 15, 165, 110, 15, "Pull a case to Review", get_new_case_btn
                    'Text 20, 170, 155, 10, "*** All Cases have been Pulled for Review ***"'
                    PushButton 45, 85, 170, 10, "More information about the Work List", work_list_details_btn
                    PushButton 10, 215, 105, 15, "Finish Work Day", finish_work_day_btn
                    PushButton 375, 5, 65, 15, "Test Access", test_access_btn
                    CancelButton 390, 215, 50, 15
                  Text 170, 10, 135, 10, "On Demand Applications Dashboard"
                  GroupBox 10, 25, 430, 75, "Applications BULK Run"
                  Text 20, 40, 170, 10, "The BULK run was last completed on " & first_item_date & "."
                  Text 190, 40, 200, 10, "The BULK Run was completed today around " & first_item_time & "."
                  Text 20, 55, 305, 10, "The BULK Run can only be completed once per day. The Work List is ready to be reviewed."
                  Text 40, 65, 305, 10, "- The worklist is held in a SQL Table and can only be viewed through this Dashboard script. "
                  Text 40, 75, 245, 10, "- Use this script to pull a case from the Work List and complete the review."
                  GroupBox 10, 105, 430, 40, "Work List Overview"
                  Text 20, 120, 115, 10, "Total cases on the worklist: " & total_cases_for_review
                  Text 220, 120, 130, 10, "Cases with Review Completed: " & cases_with_review_completed
                  Text 215, 130, 135, 10, "Cases with Reviews In Progress: " & cases_on_hold
                  GroupBox 10, 150, 430, 55, "Reviews"
                  Text 20, 190, 190, 10, "--- There are no cases with reviews already started. ---"
                EndDialog

                dialog Dialog1
                cancel_without_confirmation

                If ButtonPressed = work_list_details_btn Then MsgBox "More details will be here" 'TODO - add worklist Explanation'
                If ButtonPressed = test_access_btn Then Call test_sql_access()
            Loop until ButtonPressed <> work_list_details_btn
            CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
        Loop until are_we_passworded_out = false					'loops until user passwords back in

        If ButtonPressed = get_new_case_btn Then Call assign_a_case

    Else

        Do
            Do
                Dialog1 = ""
                grp_len = 45 + (UBound(CASES_ON_HOLD_ARRAY, 2)+1) * 10      '85'
                dlg_len = 205 + (UBound(CASES_ON_HOLD_ARRAY, 2)+1) * 10     '245'

                BeginDialog Dialog1, 0, 0, 451, dlg_len, "On Demand Applications Dashboard"
                  EditBox 500, 600, 50, 15, fake_edit_box
                  ButtonGroup ButtonPressed
                    PushButton 20, 145, 110, 15, "Pull a case to Review", get_new_case_btn
                    PushButton 375, 5, 65, 15, "Test Access", test_access_btn
                    'Text 20, 170, 155, 10, "*** All Cases have been Pulled for Review ***"'
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
                        RadioButton 30, y_pos, 120, 10, "CASE # " & CASES_ON_HOLD_ARRAY(case_nbr_const, fold_case) & " - " & CASES_ON_HOLD_ARRAY(case_notes_const, fold_case), CASES_ON_HOLD_ARRAY(radio_btn_const, fold_case)
                        y_pos = y_pos + 10
                    Next
                    ' RadioButton 30, 175, 120, 10, "CASE # XXXXXXX", case_info_array
                    ' RadioButton 30, 185, 120, 10, "CASE # XXXXXXX", Radio2
                    ' RadioButton 30, 195, 120, 10, "CASE # XXXXXXX", Radio3
                    ' RadioButton 30, 205, 120, 10, "CASE # XXXXXXX", Radio4
                  ButtonGroup ButtonPressed
                    PushButton 330, dlg_len - 45, 105, 15, "Resume selected Hold Case", resume_hold_case_btn
                    PushButton 10, dlg_len - 20, 105, 15, "Finish Work Day", finish_work_day_btn
                    CancelButton 390, dlg_len - 20, 50, 15
                EndDialog

                dialog Dialog1
                cancel_confirmation

                If ButtonPressed = work_list_details_btn Then MsgBox "More details will be here" 'TODO - add worklist Explanation'
                If ButtonPressed = test_access_btn Then Call test_sql_access()

                If ButtonPressed = resume_hold_case_btn Then
                    MsgBox "Case Information will be displayed here"            'TODO make functionality for reselecting a HOLD case'
                End If
            Loop until ButtonPressed <> work_list_details_btn
            CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
        Loop until are_we_passworded_out = false					'loops until user passwords back in

        If ButtonPressed = get_new_case_btn Then Call assign_a_case

    End If
Else


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
	          Text 10, 260, 70, 10, "Additional Notes:"
	          EditBox 10, 270, 435, 15, case_review_notes
	          ButtonGroup ButtonPressed
	            PushButton 280, 290, 110, 15, "Complete Review", complete_review_btn
	            PushButton 335, 5, 110, 15, "Put Case on Hold", hold_case_btn
                PushButton 375, 5, 65, 15, "Test Access", test_access_btn
	            CancelButton 395, 290, 50, 15
	        EndDialog

	        dialog Dialog1
	        cancel_confirmation
            If ButtonPressed = test_access_btn Then Call test_sql_access()

	        Loop until err_msg = ""
	    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
	Loop until are_we_passworded_out = false					'loops until user passwords back in

	If ButtonPressed = hold_case_btn Then
		If saved_hold_1_start_time = "" Then
			saved_hold_1_start_time = time
		ElseIf saved_hold_2_start_time = "" Then
			saved_hold_2_start_time = time
		ElseIf saved_hold_3_start_time = "" Then
			saved_hold_3_start_time = time
		End If

	End If

    If local_demo = False Then
        'Creating objects for Access
        Set objConnection = CreateObject("ADODB.Connection")
        Set objRecordSet = CreateObject("ADODB.Recordset")

        'This is the BZST connection to SQL Database'
        objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
        objSQL = "SELECT * FROM ES.ES_OnDemanCashAndSnapBZProcessed"
        objRecordSet.Open objSQL, objConnection

        'delete a record if the case number matches
        objRecordSet.Open "UPDATE ES.ES_OnDemanCashAndSnapBZProcessed SET CaseNumber = '" & MAXIS_case_number & "', " &_
                                                                          "CaseName = '" & assigned_case_name & "', " &_
                                                                          "ApplDate = '" & assigned_application_date & "', " &_
                                                                          "InterviewDate = '" & assigned_interview_date & "', " &_
                                                                          "Day_30 = '" & assigned_day_30 & "', " &_
                                                                          "DaysPending = '" & assigned_days_pending & "', " &_
                                                                          "SnapStatus = '" & assigned_snap_status & "', " &_
                                                                          "CashStatus = '" & assigned_cash_status & "', " &_
                                                                          "SecondApplicationDate = '" & assigned_2nd_application_date & "', " &_
                                                                          "REPT_PND2Days = '" & assigned_rept_pnd2_days & "', " &_
                                                                          "QuestionableInterview = '" & assigned_questionable_interview & "', " &_
                                                                          "Resolved = '" & assigned_questionable_interview_resolve & "', " &_
                                                                          "ApptNoticeDate = '" & assigned_appt_notice_date & "', " &_
                                                                          "ApptDate = '" & assigned_appt_date & "', " &_
                                                                          "Confirmation = '" & assigned_appt_notc_confirmation & "', " &_
                                                                          "NOMIDate = '" & assigned_nomi_date & "', " &_
                                                                          "Confirmation2 = '" & assigned_nomi_confirmation & "', " &_
                                                                          "DenialNeeded = '" & assigned_denial_needed & "', " &_
                                                                          "NextActionNeeded = '" & assigned_next_action_needed & "', " &_
                                                                          "AddedtoWorkList = '" & assigned_added_to_work_list & "', " &_
                                                                          "SecondApplicationDateNotes = '" & assigned_2nd_application_date_resolve & "', " &_
                                                                          "ClosedInPast30Days = '" & assigned_closed_recently & "', " &_
                                                                          "ClosedInPast30DaysNotes = '" & assigned_closed_recently_resolve & "', " &_
                                                                          "StartedOutOfCounty = '" & assigned_out_of_county & "', " &_
                                                                          "StartedOutOfCountyNotes = '" & assigned_out_of_county_resolve & "', " &_
                                                                          "TrackingNotes = '" & "STS-RC " & case_review_notes & "' WHERE CaseNumber = '" & MAXIS_case_number & "'", objConnection



        ' objRecordSet.Open "UPDATE ES.ES_OnDemanCashAndSnapBZProcessed SET CaseNumber = '" & MAXIS_case_number & "', CaseName = '" & assigned_case_name & "', ApplDate = '" & assigned_application_date & "', InterviewDate = '" & assigned_interview_date & "', "Day_30" = '" & assigned_day_30 & "', "DaysPending" = '" & assigned_days_pending & "', SnapStatus = '" & assigned_snap_status & "', CashStatus = '" & assigned_cash_status & "', SecondApplicationDate = '" & assigned_2nd_application_date & "', REPT_PND2Days = '" & assigned_rept_pnd2_days & "', QuestionableInterview = '" & assigned_questionable_interview & "', Resolved = '" & assigned_questionable_interview_resolve & "', ApptNoticeDate = '" & assigned_appt_notice_date & "', ApptDate = '" & assigned_appt_date & "', Confirmation = '" & assigned_appt_notc_confirmation & "', NOMIDate = '" & assigned_nomi_date & "', Confirmation2 = '" & assigned_nomi_confirmation & "', DenialNeeded = '" & assigned_denial_needed & "', NextActionNeeded = '" & assigned_next_action_needed & "', AddedtoWorkList = '" & assigned_added_to_work_list & "', SecondApplicationDateNotes = '" & assigned_2nd_application_date_resolve & "', ClosedInPast30Days = '" & assigned_closed_recently & "', ClosedInPast30DaysNotes = '" & assigned_closed_recently_resolve & "', StartedOutOfCounty = '" & assigned_out_of_county & "', StartedOutOfCountyNotes = '" & assigned_out_of_county_resolve & "', TrackingNotes = '" & assigned_tracking_notes & "' WHERE CaseNumber = '" & MAXIS_case_number & "'", objConnection

        'close the connection and recordset objects to free up resources
        objRecordSet.Close
        objConnection.Close
        Set objRecordSet=nothing
        Set objConnection=nothing
    End If
End If

'TODO add all the updates to Work Assignment Completed Tracking'

end_msg = "Information here of action requested"
Call script_end_procedure(end_msg)
