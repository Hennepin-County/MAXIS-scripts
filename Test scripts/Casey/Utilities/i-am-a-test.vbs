'Required for statistical purposes==========================================================================================
name_of_script = "I am a test.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 0                	'manual run time in seconds
STATS_denomination = "C"       		'C is for each CASE
'END OF stats block=========================================================================================================
run_locally = TRUE
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

case_number_list = " "
notes_list = "~~~"

worker_number_to_resolve = "YEYA001"

'declare the SQL statement that will query the database
objSQL = "SELECT * FROM ES.ES_OnDemanCashAndSnapBZProcessed"

'Creating objects for Access
Set objConnection = CreateObject("ADODB.Connection")
Set objRecordSet = CreateObject("ADODB.Recordset")

'This is the file path for the statistics Access database.
objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
objRecordSet.Open objSQL, objConnection


Do While NOT objRecordSet.Eof
	'count all of today's cases using added to worklist
	worklist_case_number = objRecordSet("CaseNumber")
	case_worklist_date = objRecordSet("AddedtoWorkList")
	case_worklist_date = DateAdd("d", 0, case_worklist_date)
	case_tracking_notes = objRecordSet("TrackingNotes")
	' case_worklist_date = date
	If DateDiff("d", case_worklist_date, date) = 0  and InStr(case_tracking_notes, "STS") = 0 Then
		case_number_list = case_number_list & worklist_case_number & " "
		case_tracking_notes = "STS-RC-"& worker_number_to_resolve & " " & case_tracking_notes
		case_tracking_notes = trim(case_tracking_notes)
		notes_list = notes_list & case_tracking_notes & "~~~"

	End If
	objRecordSet.MoveNext
Loop

'close the connection and recordset objects to free up resources
objRecordSet.Close
objConnection.Close
Set objRecordSet=nothing
Set objConnection=nothing

case_number_list = trim(case_number_list)
The_case_number_array = split(case_number_list)

If right(notes_list, 3) = "~~~" Then notes_list = left(notes_list, len(notes_list)-3)
If left(notes_list, 3) = "~~~" Then notes_list = right(notes_list, len(notes_list)-3)
notes_array = split(notes_list, "~~~")
' case_review_notes = replace(case_review_notes, "STS-IP-"&worker_number_to_resolve, "")
' case_review_notes = replace(case_review_notes, "STS-HD-"&worker_number_to_resolve, "")
' case_review_notes = replace(case_review_notes, "STS-RC-"&worker_number_to_resolve, "")
' case_review_notes = replace(case_review_notes, "STS-RC", "")
' case_review_notes = trim(case_review_notes)


For each_item = 0 to UBOUND(The_case_number_array)

	Set objConnection = CreateObject("ADODB.Connection")
	Set objRecordSet = CreateObject("ADODB.Recordset")

	'This is the BZST connection to SQL Database'
	objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
	' objSQL = "SELECT * FROM ES.ES_OnDemanCashAndSnapBZProcessed"
	' objRecordSet.Open objSQL, objConnection

	'delete a record if the case number matches
	objRecordSet.Open "UPDATE ES.ES_OnDemanCashAndSnapBZProcessed SET TrackingNotes = '" & notes_array(each_item) & "' WHERE CaseNumber = '" & The_case_number_array(each_item) & "'", objConnection



	' objRecordSet.Open "UPDATE ES.ES_OnDemanCashAndSnapBZProcessed SET CaseNumber = '" & MAXIS_case_number & "', CaseName = '" & assigned_case_name & "', ApplDate = '" & assigned_application_date & "', InterviewDate = '" & assigned_interview_date & "', "Day_30" = '" & assigned_day_30 & "', "DaysPending" = '" & assigned_days_pending & "', SnapStatus = '" & assigned_snap_status & "', CashStatus = '" & assigned_cash_status & "', SecondApplicationDate = '" & assigned_2nd_application_date & "', REPT_PND2Days = '" & assigned_rept_pnd2_days & "', QuestionableInterview = '" & assigned_questionable_interview & "', Resolved = '" & assigned_questionable_interview_resolve & "', ApptNoticeDate = '" & assigned_appt_notice_date & "', ApptDate = '" & assigned_appt_date & "', Confirmation = '" & assigned_appt_notc_confirmation & "', NOMIDate = '" & assigned_nomi_date & "', Confirmation2 = '" & assigned_nomi_confirmation & "', DenialNeeded = '" & assigned_denial_needed & "', NextActionNeeded = '" & assigned_next_action_needed & "', AddedtoWorkList = '" & assigned_added_to_work_list & "', SecondApplicationDateNotes = '" & assigned_2nd_application_date_resolve & "', ClosedInPast30Days = '" & assigned_closed_recently & "', ClosedInPast30DaysNotes = '" & assigned_closed_recently_resolve & "', StartedOutOfCounty = '" & assigned_out_of_county & "', StartedOutOfCountyNotes = '" & assigned_out_of_county_resolve & "', TrackingNotes = '" & assigned_tracking_notes & "' WHERE CaseNumber = '" & MAXIS_case_number & "'", objConnection

	'close the connection and recordset objects to free up resources
	' objRecordSet.Close
	objConnection.Close
	Set objRecordSet=nothing
	Set objConnection=nothing
Next


'this is time zone functionality
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
For Each objItem In colItems
	MsgBox "Current Time Zone (Hours Offset From GMT): " & (objItem.CurrentTimeZone / 60)
	MsgBox "Daylight Saving In Effect: " & objItem.DaylightInEffect
Next

MsgBox "STOP"

const memb_ref_numb_const 	= 0
const memb_name_const 		= 1
const memb_age_const		= 2
const memb_is_caregiver 	= 3
const cash_request_const	= 4
const hours_per_week_const	= 5
const exempt_from_ed_const	= 6
const comply_with_ed_const	= 7
const orientation_needed_const	= 8
const orientation_done_const	= 9
const orientation_exempt_const	= 10
const exemption_reason_const	= 11
const emps_exemption_code_const	= 12
const choice_form_done_const 	= 13
const orientation_notes			= 14


Dim CAREGIVER_ARRAY()
ReDim CAREGIVER_ARRAY(orientation_notes, 3)

CAREGIVER_ARRAY(memb_ref_numb_const, 0) = ""
CAREGIVER_ARRAY(memb_name_const, 0) = "Beverly Miller"
CAREGIVER_ARRAY(memb_age_const, 0) = 25
CAREGIVER_ARRAY(cash_request_const, 0) = False
CAREGIVER_ARRAY(memb_is_caregiver, 0) = False
CAREGIVER_ARRAY(hours_per_week_const, 0) = ""
CAREGIVER_ARRAY(exempt_from_ed_const, 0) = False
CAREGIVER_ARRAY(comply_with_ed_const, 0) = False
CAREGIVER_ARRAY(orientation_needed_const, 0) = False
CAREGIVER_ARRAY(orientation_exempt_const, 0) = False
' CAREGIVER_ARRAY(orientation_done_const, 0) = False
' CAREGIVER_ARRAY(choice_form_done_const, 0) = False

CAREGIVER_ARRAY(memb_ref_numb_const, 1) = ""
CAREGIVER_ARRAY(memb_name_const, 1) = "Corbin Miller"
CAREGIVER_ARRAY(memb_age_const, 1) = 25
CAREGIVER_ARRAY(cash_request_const, 1) = False
CAREGIVER_ARRAY(memb_is_caregiver, 1) = False
CAREGIVER_ARRAY(hours_per_week_const, 1) = ""
CAREGIVER_ARRAY(exempt_from_ed_const, 1) = False
CAREGIVER_ARRAY(comply_with_ed_const, 1) = False
CAREGIVER_ARRAY(orientation_needed_const, 1) = False
CAREGIVER_ARRAY(orientation_exempt_const, 1) = False
' CAREGIVER_ARRAY(orientation_done_const, 1) = False
' CAREGIVER_ARRAY(choice_form_done_const, 1) = False

CAREGIVER_ARRAY(memb_ref_numb_const, 2) = ""
CAREGIVER_ARRAY(memb_name_const, 2) = "Ava Miller"
CAREGIVER_ARRAY(memb_age_const, 2) = 1
CAREGIVER_ARRAY(cash_request_const, 2) = False
CAREGIVER_ARRAY(memb_is_caregiver, 2) = False
CAREGIVER_ARRAY(hours_per_week_const, 2) = ""
CAREGIVER_ARRAY(exempt_from_ed_const, 2) = False
CAREGIVER_ARRAY(comply_with_ed_const, 2) = False
CAREGIVER_ARRAY(orientation_needed_const, 2) = False
CAREGIVER_ARRAY(orientation_exempt_const, 2) = False
' CAREGIVER_ARRAY(orientation_done_const, 2) = False
' CAREGIVER_ARRAY(choice_form_done_const, 2) = False

CAREGIVER_ARRAY(memb_ref_numb_const, 3) = ""
CAREGIVER_ARRAY(memb_name_const, 3) = "Benny Miller"
CAREGIVER_ARRAY(memb_age_const, 3) = 1
CAREGIVER_ARRAY(cash_request_const, 3) = False
CAREGIVER_ARRAY(memb_is_caregiver, 3) = False
CAREGIVER_ARRAY(hours_per_week_const, 3) = ""
CAREGIVER_ARRAY(exempt_from_ed_const, 3) = False
CAREGIVER_ARRAY(comply_with_ed_const, 3) = False
CAREGIVER_ARRAY(orientation_needed_const, 3) = False
CAREGIVER_ARRAY(orientation_exempt_const, 3) = False
' CAREGIVER_ARRAY(orientation_done_const, 3) = False
' CAREGIVER_ARRAY(choice_form_done_const, 3) = False

MAXIS_case_number = "311021"


function complete_MFIP_orientation(CAREGIVER_ARRAY, memb_ref_numb_const, memb_name_const, memb_age_const, memb_is_caregiver, cash_request_const, hours_per_week_const, exempt_from_ed_const, comply_with_ed_const, orientation_needed_const, orientation_done_const, orientation_exempt_const, exemption_reason_const, emps_exemption_code_const, choice_form_done_const, orientation_notes)


	'first - assess if caregiver meets an exemption
		'- Single parent household employed at least 35 hours per week
		'- 2 Parent household where the 1st parent is employed at least 35 hours per week
		'- 2 Parened household where the 2nd parent is employed at least 20 hours per week and the 1st is employed 35
		'- Pregnant or parenting minor under 20 who is coplying with the educational requirements
		'- Caregiver is not receiving MFIP

	'Identify the caregivers
	'Identify if they are requesting Cash
	'Indicate if this will be DWP or MFIP
	'Identify if the caregiver is a minor
	'List the hours employed for each caregiver
	'
	person_list = "Select One..."+chr(9)+"No Caregiver"
	second_person_list = "Select One..."+chr(9)+"No Second Caregiver"

	For person = 0 to UBound(CAREGIVER_ARRAY, 2)
		person_list = person_list+chr(9)+CAREGIVER_ARRAY(memb_name_const, person)
		second_person_list = second_person_list+chr(9)+CAREGIVER_ARRAY(memb_name_const, person)
	Next
	caregiver_one = CAREGIVER_ARRAY(memb_name_const, 0)

	Do
		err_msg = ""
		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 551, 150, "Assess for Caregiver MFIP Orientation Requirement"
		  DropListBox 185, 10, 60, 45, "MFIP"+chr(9)+"DWP", family_cash_program
		  EditBox 110, 30, 430, 15, famliy_cash_notes
		  DropListBox 65, 65, 140, 45, person_list, caregiver_one
		  DropListBox 330, 65, 45, 45, "Yes"+chr(9)+"No"+chr(9)+"Not Elig", caregiver_one_req_cash
		  EditBox 430, 65, 30, 15, caregiver_one_hours_per_week
		  DropListBox 65, 85, 140, 45, second_person_list, caregiver_two
		  DropListBox 330, 85, 45, 45, "Yes"+chr(9)+"No"+chr(9)+"Not Elig", caregiver_two_req_cash
		  EditBox 430, 85, 30, 15, caregiver_two_hours_per_week
		  Text 15, 125, 450, 20, "These questions will identify if these caregivers need an MFIP orientation. See CM 05.12.12.06   to see the reasons that a caregiver would not need an MFIP Orientation. The script will use this information to determine if the MFIP Orientation Functionality should be run."
		  ButtonGroup ButtonPressed
			OkButton 490, 125, 50, 15
			PushButton 260, 123, 55, 10, "CM05.12.12.06", cm_05_12_12_06_btn
		  Text 10, 15, 170, 10, "Which Family Cash Program is this Application for?"
		  Text 10, 35, 100, 10, "Notes on Program Selection:"
		  GroupBox 10, 50, 530, 55, "Who are the Caregivers"
		  Text 20, 70, 40, 10, "Caregiver:"
		  Text 215, 70, 115, 10, "Is this caregiver requesting cash?"
		  Text 385, 70, 40, 10, "Employed: "
		  Text 465, 70, 50, 10, "hours/week"
		  Text 20, 90, 40, 10, "Caregiver:"
		  Text 215, 90, 115, 10, "Is this caregiver requesting cash?"
		  Text 385, 90, 40, 10, "Employed: "
		  Text 465, 90, 50, 10, "hours/week"
		  Text 15, 110, 100, 10, "Why is this being asked?"
		EndDialog

		dialog Dialog1
		cancel_confirmation

		If caregiver_one = "Select One..." Then err_msg = err_msg & vbCr & "* Indicate the First Caregiver or clarify that there is no caregiver"
		If caregiver_two = "Select One..." Then err_msg = err_msg & vbCr & "* Indicate the Second Caregiver or clarify that there is no second caregiver"
		If caregiver_one = caregiver_two Then err_msg = err_msg & vbCr & "* Select two different caregivers"
		If IsNumeric(caregiver_one_hours_per_week) = False AND trim(caregiver_one_hours_per_week) <> "" Then err_msg = err_msg & vbCr & "* Hours per week should be a number, or left blank."
		If IsNumeric(caregiver_two_hours_per_week) = False AND trim(caregiver_two_hours_per_week) <> "" Then err_msg = err_msg & vbCr & "* Hours per week should be a number, or left blank."

		If family_cash_program = "DWP" Then err_msg = ""

		If ButtonPressed <> -1 Then err_msg = "LOOP"
		If err_msg <> "" And ButtonPressed = -1 Then MsgBox err_msg

		If ButtonPressed = cm_05_12_12_06_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=CM_0005121206"

	Loop until err_msg = ""

	If family_cash_program = "MFIP" Then
		If IsNumeric(caregiver_one_hours_per_week) = True Then caregiver_one_hours_per_week = caregiver_one_hours_per_week * 1
		If trim(caregiver_one_hours_per_week) = "" Then caregiver_one_hours_per_week = 0

		If IsNumeric(caregiver_two_hours_per_week) = True Then caregiver_two_hours_per_week = caregiver_two_hours_per_week * 1
		If trim(caregiver_two_hours_per_week) = "" Then caregiver_two_hours_per_week = 0

		minor_caregiver_on_case = 0

		For person = 0 to UBound(CAREGIVER_ARRAY, 2)
			If CAREGIVER_ARRAY(memb_name_const, person) = caregiver_one Then
				CAREGIVER_ARRAY(memb_is_caregiver, person) = True
				CAREGIVER_ARRAY(orientation_needed_const, person) = True

				If caregiver_one_req_cash = "Yes" Then CAREGIVER_ARRAY(cash_request_const, person) = True
				If caregiver_one_req_cash <> "Yes" Then
					CAREGIVER_ARRAY(cash_request_const, person) = False
					CAREGIVER_ARRAY(orientation_needed_const, person) = False
					CAREGIVER_ARRAY(orientation_exempt_const, person) = True
					CAREGIVER_ARRAY(exemption_reason_const, person) = "Caregiver Not on MFIP"
					CAREGIVER_ARRAY(emps_exemption_code_const, person) = "NO"
				End If
				CAREGIVER_ARRAY(hours_per_week_const, person) = caregiver_one_hours_per_week

				If CAREGIVER_ARRAY(hours_per_week_const, person) > 34 Then
					CAREGIVER_ARRAY(orientation_needed_const, person) = False
					CAREGIVER_ARRAY(orientation_exempt_const, person) = True
					CAREGIVER_ARRAY(exemption_reason_const, person) = "Employed 35+ hours per week"
					CAREGIVER_ARRAY(emps_exemption_code_const, person) = "20"
				ElseIf CAREGIVER_ARRAY(hours_per_week_const, person) > 19 Then
					If caregiver_two <> "No Second Caregiver" AND caregiver_two_req_cash = "Yes" AND caregiver_two_hours_per_week > 34 Then
						CAREGIVER_ARRAY(orientation_needed_const, person) = False
						CAREGIVER_ARRAY(orientation_exempt_const, person) = True
						CAREGIVER_ARRAY(exemption_reason_const, person) = "2nd Caregiver Employed 20+ hours per week"
						CAREGIVER_ARRAY(emps_exemption_code_const, person) = "21"
					End If
				End If
				If CAREGIVER_ARRAY(memb_age_const, person) < 20 Then
					minor_caregiver_on_case = minor_caregiver_on_case + 1
					CAREGIVER_ARRAY(exempt_from_ed_const, person) = "No"
					CAREGIVER_ARRAY(comply_with_ed_const, person) = "Yes"
				End If

			End If

			If CAREGIVER_ARRAY(memb_name_const, person) = caregiver_two Then
				CAREGIVER_ARRAY(memb_is_caregiver, person) = True
				CAREGIVER_ARRAY(orientation_needed_const, person) = True

				If caregiver_two_req_cash = "Yes" Then CAREGIVER_ARRAY(cash_request_const, person) = True
				If caregiver_two_req_cash <> "Yes" Then
					CAREGIVER_ARRAY(cash_request_const, person) = False
					CAREGIVER_ARRAY(orientation_needed_const, person) = False
					CAREGIVER_ARRAY(orientation_exempt_const, person) = True
					CAREGIVER_ARRAY(exemption_reason_const, person) = "Caregiver Not on MFIP"
					CAREGIVER_ARRAY(emps_exemption_code_const, person) = "NO"
				End If
				CAREGIVER_ARRAY(hours_per_week_const, person) = caregiver_two_hours_per_week

				If CAREGIVER_ARRAY(hours_per_week_const, person) > 34 Then
					CAREGIVER_ARRAY(orientation_needed_const, person) = False
					CAREGIVER_ARRAY(orientation_exempt_const, person) = True
					CAREGIVER_ARRAY(exemption_reason_const, person) = "Employed 35+ hours per week"
					CAREGIVER_ARRAY(emps_exemption_code_const, person) = "20"
				ElseIf CAREGIVER_ARRAY(hours_per_week_const, person) > 19 Then
					If caregiver_one <> "No Second Caregiver" AND caregiver_one_req_cash = "Yes" AND caregiver_one_hours_per_week > 34 Then
						CAREGIVER_ARRAY(orientation_needed_const, person) = False
						CAREGIVER_ARRAY(orientation_exempt_const, person) = True
						CAREGIVER_ARRAY(exemption_reason_const, person) = "2nd Caregiver Employed 20+ hours per week"
						CAREGIVER_ARRAY(emps_exemption_code_const, person) = "21"
					End If
				End If
				If CAREGIVER_ARRAY(memb_age_const, person) < 20 Then
					minor_caregiver_on_case = minor_caregiver_on_case + 1
					CAREGIVER_ARRAY(exempt_from_ed_const, person) = "No"
					CAREGIVER_ARRAY(comply_with_ed_const, person) = "Yes"
				End If
			End If



		Next

		'IF A MINOR IS FOUND
		If minor_caregiver_on_case > 0 Then
			Do
				err_msg = ""
				dlg_len = 210
				If minor_caregiver_on_case = 2 Then dlg_len = 290

				Dialog1 = ""
				BeginDialog Dialog1, 0, 0, 551, dlg_len, "Assess for Caregiver MFIP Orientation Requirement"
				  Text 10, 15, 200, 10, "Which Family Cash Program is this Application for? " & family_cash_program
				  Text 10, 25, 500, 20, "Notes on Program Selection: " & famliy_cash_notes
				  GroupBox 10, 50, 530, 40, "Who are the Caregivers"
				  Text 20, 60, 190, 10, "Caregiver: " & caregiver_one
				  Text 215, 60, 165, 10, "Is this caregiver requesting cash? " & caregiver_one_req_cash
				  Text 385, 60, 90, 10, "Employed: " & caregiver_one_hours_per_week
				  Text 465, 60, 50, 10, "hours/week"
				  Text 20, 75, 190, 10, "Caregiver: " & caregiver_two
				  Text 215, 75, 165, 10, "Is this caregiver requesting cash? " & caregiver_two_req_cash
				  Text 385, 75, 90, 10, "Employed: " & caregiver_two_hours_per_week
				  Text 465, 75, 50, 10, "hours/week"
				  y_pos = 30
				  For caregiver = 0 to UBound(CAREGIVER_ARRAY, 2)
					  If CAREGIVER_ARRAY(memb_is_caregiver, caregiver) = True and CAREGIVER_ARRAY(memb_age_const, caregiver) < 20 Then
						  y_pos = y_pos + 70
						  GroupBox 10, y_pos, 530, 65, CAREGIVER_ARRAY(memb_name_const, caregiver)
						  Text 20, y_pos + 10, 270, 10, "This caregiver appears to be a minor by MFIP program rules (under 20 years old)."
						  Text 20, y_pos + 30, 195, 10, "Is this caregiver exempt from the Educational Requirement?"
						  DropListBox 230, y_pos + 25, 40, 45, "No"+chr(9)+"Yes", CAREGIVER_ARRAY(exempt_from_ed_const, caregiver)
						  Text 20, y_pos + 50, 205, 10, "Is this caregiver complying with the Educational Requirement?"
						  DropListBox 230, y_pos + 45, 40, 45, "No"+chr(9)+"Yes", CAREGIVER_ARRAY(comply_with_ed_const, caregiver)
					  End If
				  Next
				  Text 15, y_pos + 90, 450, 20, "These questions will identify if these caregivers need an MFIP orientation. See CM 05.12.12.06 to see the reasons that a caregiver would not need an MFIP Orientation. The script will use this information to determine if the MFIP Orientation Functionality should be run."
				  ButtonGroup ButtonPressed
					OkButton 490, y_pos + 90, 50, 15
					PushButton 485, y_pos + 45, 50, 15, "CM 28.12", cm_28_12_btn
					PushButton 260, y_pos + 87, 55, 10, "CM05.12.12.06", cm_05_12_12_06_btn
				  Text 355, y_pos + 45, 125, 20, "See details about the educational requirement in the Combined Manual "
				  Text 15, y_pos + 75, 100, 10, "Why is this being asked?"
				EndDialog

				dialog Dialog1
				cancel_confirmation

				If err_msg <> "" Then MsgBox err_msg

			Loop until err_msg = ""

			For caregiver = 0 to UBound(CAREGIVER_ARRAY, 2)
				If CAREGIVER_ARRAY(memb_is_caregiver, caregiver) = True and CAREGIVER_ARRAY(memb_age_const, caregiver) < 20 Then
					If CAREGIVER_ARRAY(exempt_from_ed_const, caregiver) = "No" Then CAREGIVER_ARRAY(exempt_from_ed_const, caregiver) = False
					If CAREGIVER_ARRAY(exempt_from_ed_const, caregiver) = "Yes" Then CAREGIVER_ARRAY(exempt_from_ed_const, caregiver) = True
					If CAREGIVER_ARRAY(comply_with_ed_const, caregiver) = "No" Then CAREGIVER_ARRAY(comply_with_ed_const, caregiver) = False
					If CAREGIVER_ARRAY(comply_with_ed_const, caregiver) = "Yes" Then CAREGIVER_ARRAY(comply_with_ed_const, caregiver) = True

					If CAREGIVER_ARRAY(exempt_from_ed_const, caregiver) = False and CAREGIVER_ARRAY(comply_with_ed_const, caregiver) = True Then
						CAREGIVER_ARRAY(orientation_needed_const, caregiver) = False
						CAREGIVER_ARRAY(orientation_exempt_const, caregiver) = True
						CAREGIVER_ARRAY(exemption_reason_const, caregiver) = "Minor Caregiver meeting Educational Requirements"
						CAREGIVER_ARRAY(emps_exemption_code_const, caregiver) = "22"
					End If
				Else
					CAREGIVER_ARRAY(exempt_from_ed_const, caregiver) = False
					CAREGIVER_ARRAY(comply_with_ed_const, caregiver) = False
				End If
			Next

		End If

		const mf_step_rights_resp 	= 1
		const mf_step_time_limits	= 2
		const mf_step_extension		= 3
		const mf_step_dv			= 4
		const mf_step_expectations	= 5
		const mf_step_esp			= 6
		const mf_step_compliance	= 7
		const mf_step_ep			= 8
		const mf_step_ccap			= 9
		const mf_step_incentives	= 10
		const mf_step_hc			= 11
		const mf_completion			= 12

		' mf_step_rights_resp_viewed = False
		' mf_step_time_limits_viewed = False
		' mf_step_extension_viewed = False
		' mf_step_dv_viewed = False
		' mf_step_expectations_viewed = False
		' mf_step_esp_viewed = False
		' mf_step_compliance_viewed = False
		' mf_step_ep_viewed = False
		' mf_step_ccap_viewed = False
		' mf_step_incentives_viewed = False
		' mf_step_hc_viewed = False
		' mf_completion_viewed = False
		'
		' orientation_script_document_viewed = False
		'
		'FIRST - Participant Responsibilities and Rights'
		'SECOND - MFIP Time Limits'
		'THIRD - MFIp Extension Eligibility'
		'FOURTH - Family Violence'
		'FIFTH - Expectations'
		'SIXTH - Choosing ESP'
		'SEVENTH - Assignment and Compliance'
		'EIGHTH - Developing an EP'
		'NINTH - CCAP'
		'TENTH - Incentives'
		'ELEVENTH - Health Care'

		' all_mfip_orientation_info_viewed = False
		For caregiver = 0 to UBound(CAREGIVER_ARRAY, 2)
			' Call Navigate_to_MAXIS_screen("STAT", "EMPS")
			' If CAREGIVER_ARRAY(memb_ref_numb_const, caregiver) <> "" Then
			' 	EMWriteScreen CAREGIVER_ARRAY(memb_ref_numb_const, caregiver), 20, 76
			' 	transmit
			' End If

			If CAREGIVER_ARRAY(orientation_needed_const, caregiver) = True Then
				MFIP_orientation_step = mf_step_rights_resp

				mf_step_rights_resp_viewed = False
				mf_step_time_limits_viewed = False
				mf_step_extension_viewed = False
				mf_step_dv_viewed = False
				mf_step_expectations_viewed = False
				mf_step_esp_viewed = False
				mf_step_compliance_viewed = False
				mf_step_ep_viewed = False
				mf_step_ccap_viewed = False
				mf_step_incentives_viewed = False
				mf_step_hc_viewed = False
				mf_completion_viewed = False

				orientation_script_document_viewed = False

				all_mfip_orientation_info_viewed = False

				Do
					err_msg = ""

					Dialog1 = ""
					BeginDialog Dialog1, 0, 0, 551, 385, "MFIP Orientation"
					  ' GroupBox 10, 10, 450, 45, "Group1"
					  ButtonGroup ButtonPressed
					  	If MFIP_orientation_step <> mf_completion Then PushButton 495, 365, 50, 15, "NEXT", next_btn

						'FIRST - Participant Responsibilities and Rights'
						If MFIP_orientation_step = mf_step_rights_resp Then
						  Text 10, 10, 145, 10, "Caregiver: " & CAREGIVER_ARRAY(memb_name_const, caregiver)
						  GroupBox 10, 30, 450, 130, "Participant Responsibilities and Rights"
						  Text 20, 45, 370, 10, "As a program participant you have responsibilities and rights that were discussed during your intake interview."
						  Text 20, 60, 430, 10, "Please keep a copy of the Client Responsibilities and Rights (DHS-4163) for your reference. Let us know if you have any questions."
						  Text 20, 80, 335, 10, "Please remember it's important to report ANY changes that could affect your eligibility within 10 days."
						  GroupBox 10, 75, 450, 15, ""
						  Text 20, 100, 420, 20, "If your income decreases by at least 50% contact your financial worker right away!  You may be eligible for a significant change meaning a recalculation of your income which may result in an increase of your cash and/or food benefits."
						  Text 20, 125, 335, 20, "If you do not meet program eligibility such as cash assistance, your financial worker will assess other program eligibility such as SNAP."
						  PushButton 385, 160, 75, 15, "DHS - 4163", open_dhs_4163_btn

						  mf_step_rights_resp_viewed = True
						  'ADD BUTTON DHS 4163'
						End If

						'SECOND - MFIP Time Limits'
						If MFIP_orientation_step = mf_step_time_limits Then
						  GroupBox 10, 10, 450, 160, "MFIP Time-Limits"
						  Text 20, 25, 430, 30, "The MFIP program is available to you for up to 60 months in your lifetime.  If you have used cash assistance in another state those months must be reported and may count toward your lifetime limit. There are some instances the months you use may be exempt, meaning the months do not count towards the 60-month lifetime limit."
						  Text 20, 55, 55, 10, "These Include:"
						  Text 30, 70, 125, 10, "1. Months you are over 60 years old"
						  Text 30, 80, 310, 10, "2. Months you are living on a reservation where at least 50% of the adults were not employed"
						  Text 30, 90, 360, 10, "3. Months when you are a victim of family violence AND have an approved family violence waiver plan"
						  Text 30, 100, 335, 10, "4. Months you don't receive the cash portion of MFIP (*talk to your financial worker for more details)"
						  Text 30, 110, 350, 10, "5. Months you are a parent under 18 years of age and complying with your school or social service plan"
						  Text 30, 120, 395, 10, "6. Months you are 18 or 19 years old and do not have a high school diploma/GED AND complying with a school plan"
						  Text 40, 135, 355, 25, "Note: If you are eligible for an exemption but you are not complying with program requirements and do not meet a good cause reason, those months will count toward the lifetime limit. If you have questions about possible good cause reasons, talk to a worker."
							  Text 10, 175, 145, 10, "Caregiver: " & CAREGIVER_ARRAY(memb_name_const, caregiver)

						  mf_step_time_limits_viewed = True
						End If

						'THIRD - MFIp Extension Eligibility'
						If MFIP_orientation_step = mf_step_extension Then
						  GroupBox 10, 10, 450, 295, "MFIP Extension Eligibility"
						  Text 20, 30, 165, 10, "You may be eligible for an MFIP Extension if:"
						  Text 30, 45, 375, 10, "- You are a single or two parent household working the required number of hours that meet extension eligibility"
						  Text 30, 55, 365, 10, "- Your health care provider states you are only able to work 20 hours per week due to an illness or disability"
						  Text 20, 75, 380, 20, "A qualified professional verifies you have one or more of the conditions below that severely limits your ability to obtain or maintain suitable employment for 20 or more hours per week:"
						  Text 30, 100, 165, 10, "- Developmentally Disabled or Mentally Ill"
						  Text 30, 110, 95, 10, "- Learning Disability"
						  Text 30, 120, 60, 10, "- IQ Below 80"
						  Text 30, 130, 260, 10, "- You are ill/injured or incapacitated that's expected to last more than 30 days"
						  Text 20, 145, 125, 10, "A qualified professional verifies:"
						  Text 35, 160, 280, 15, "You are needed in the home to provide care for a family member or foster child in the household that is expected to continue for more than 30 days "
						  Text 35, 185, 285, 35, "A child or adult in the home meets the Special Medical Criteria for home care services or a home and community-based waiver services program, severe emotional disturbance (SED diagnosed child) or serious and persistent mental illness (SPMI diagnosed adult)"
						  Text 35, 225, 275, 20, "You have significant barriers to employment and determined Unemployable by a vocational specialist or other qualified professional designated by the county"
						  Text 35, 250, 165, 10, "You are a victim of family violence"
						  Text 20, 265, 415, 30, "If you believe you meet any of the criteria's above it's important to discuss with your financial worker AND your employment counselor. You may qualify for a modified employment plan prior to reaching your 60-month as well as receive an extension of your cash benefits."
							  Text 10, 310, 145, 10, "Caregiver: " & CAREGIVER_ARRAY(memb_name_const, caregiver)

						  mf_step_extension_viewed = True
						End If

						'FOURTH - Family Violence'
						If MFIP_orientation_step = mf_step_dv Then
						  GroupBox 10, 10, 450, 75, "Family Violence Resources/Supports"
						  Text 20, 30, 390, 10, "Your financial worker discussed and provided information regarding resources if you are a victim of family violence."
						  Text 20, 45, 375, 35, " Please review that brochure if you need assistance with shelter and/or supports Domestic Violence Information (DHS 3477) and Family Violence Referral (DHS 3323). If you are a victim of domestic violence, you may choose to work with your assigned Employment Counselor to determine if you are eligible for a Family Violence Waiver to allow your family time and flexibility to focus on safety issues."
							  Text 10, 90, 145, 10, "Caregiver: " & CAREGIVER_ARRAY(memb_name_const, caregiver)
						  PushButton 385, 90, 75, 15, "DHS - 3477", open_dhs_3477_btn
						  PushButton 385, 105, 75, 15, "DHS - 3323", open_dhs_3323_btn

						  mf_step_dv_viewed = True
						  'ADD BUTTON DHS 3477
						  'ADD BUTTON DHS 3323
						End If

						'FIFTH - Expectations'
						If MFIP_orientation_step = mf_step_expectations Then
						  GroupBox 10, 10, 450, 110, "Expectations of Participants Approved for the MFIP Program"
						  Text 20, 30, 360, 20, "MFIP services focus on putting you on the most direct path to employment and other related steps that will support long-term economic stability."
						  Text 20, 55, 375, 20, "While you are expected to work, look for work, or participate in activities to prepare for work, the steps toward economic stability look different for all families and participants."
						  Text 20, 80, 405, 20, "Employment Services have a variety of tools to address the unique needs of each family. You will hear more about these tools and resources during your Employment Services Overview."
							  Text 10, 125, 145, 10, "Caregiver: " & CAREGIVER_ARRAY(memb_name_const, caregiver)

						  mf_step_expectations_viewed = True
						End If

						'SIXTH - Choosing ESP'
						If MFIP_orientation_step = mf_step_esp Then
						  GroupBox 10, 10, 450, 155, "Choosing an MFIP Employment Service Provider (ESP)"
						  Text 25, 25, 345, 10, "As part of the MFIP program you are required to work with an MFIP Employment Service Provider (ESP)."
						  Text 25, 40, 410, 20, "There's a variety of providers available to help support your employment goals. On the MFIP ESP Choice Sheet, choose the top three providers you'd like to work with listing your top three choices with 1 being the provider you most want to work with."
						  Text 25, 65, 330, 10, "We will do our best to refer you to one of your top three choices depending on available openings."
						  Text 25, 85, 195, 10, "There are a few exceptions in choosing your provider:"
						  Text 40, 100, 345, 10, "If you have worked with an MFIP ESP in the past ninety (90) days, you may be referred to that provider."
						  Text 40, 115, 350, 20, "If you are under 18 and do not have a HS diploma/GED, you will be referred to Minnesota Visiting Nurse Association to discuss your education and employment options"
						  Text 40, 140, 345, 20, "If you have used 60 months or more of your TANF time limit and granted an extension under a specific category you will be referred to an agency that specializes in that type of extension."
							  Text 10, 170, 145, 10, "Caregiver: " & CAREGIVER_ARRAY(memb_name_const, caregiver)
						  ' PushButton 385, 170, 75, 15, "Choice Sheet", open_choice_sheet_btn

						  mf_step_esp_viewed = True
						  'ADD BUTTON - CHOICE SHEET ???'
						End If

						'SEVENTH - Assignment and Compliance'
						If MFIP_orientation_step = mf_step_compliance Then
						  GroupBox 10, 10, 450, 110, "Assignment and Compliance with MFIP Employment Services"
						  Text 25, 30, 295, 10, "Once you are approved for MFIP you will be referred to an Employment Service Provider."
						  Text 25, 45, 375, 20, "In Hennepin County, many of the Employment Services Providers are community based nonprofit organizations who partner with Hennepin County to deliver services."
						  Text 25, 70, 410, 20, "The provider will send you a notice to attend an MFIP Employment Service Overview. You are required to attend the overview and work with your assigned employment service counselor."
						  Text 25, 95, 400, 20, "If you choose not to comply with program requirements, your case may be sanctioned resulting in a reduction of your cash and/or food benefits."
							  Text 10, 125, 145, 10, "Caregiver: " & CAREGIVER_ARRAY(memb_name_const, caregiver)

						  mf_step_compliance_viewed = True
						End If

						'EIGHTH - Developing an EP'
						If MFIP_orientation_step = mf_step_ep Then
						  GroupBox 10, 10, 450, 270, "Developing an Employment Plan (EP) with your MFIP Employment Counselor"
						  Text 20, 25, 430, 25, "Program participants will work with their assigned Employment Counselor to develop an Employment Plan. Your Employment Plan will be based on your goals and will include activities that are intended to lead to employment and financial stability. On the path to stable employment, many different types of activities are available."
						  Text 20, 55, 140, 10, "Some of the allowable activities include:"
						  Text 30, 70, 260, 10, "- Job search (including participation in job clubs, workshops, and hiring events)"
						  Text 30, 80, 260, 10, "- Employment"
						  Text 30, 90, 260, 10, "- Self-employment"
						  Text 30, 100, 260, 10, "- Community work experience and/or volunteer work"
						  Text 30, 110, 260, 10, "- On the job training"
						  Text 30, 120, 260, 10, "- English Language Learning (ELL and ESL) or Functional Work Literacy (FWL)"
						  Text 30, 130, 260, 10, "- Adult Basic Education, GED preparation and Adult High School Diploma"
						  Text 30, 140, 260, 10, "- Job skills training directly related to employment"
						  Text 30, 150, 260, 10, "- Post-Secondary Training and Education"
						  Text 30, 160, 415, 10, "- Other activities that are critical to your family's success in reaching your employment goals such as chemical dependency"
						  Text 35, 170, 260, 10, "treatment, mental health services, social services, and parenting education."
						  Text 20, 190, 430, 25, "You are required to follow through with the activities in your employment plan. If you are unable to complete the activities, contact your Employment Counselor right away to determine if your plan need to be updated. Good communication with your employment counselor can help prevent reduction in your grant (sanctions)."
						  Text 20, 220, 425, 30, "Your Employment Counselor may conduct assessments with you to support you in selecting an education and training path that creates opportunities for long term economic stability. If you have more questions about education and training options, you can also see the Education and Training Brochure (DHS 3366)."
						  Text 20, 255, 420, 20, "Work study programs under the higher education systems may also be available.  Your assigned employment counselor will discuss this opportunity in more detail."
							  Text 10, 285, 145, 10, "Caregiver: " & CAREGIVER_ARRAY(memb_name_const, caregiver)
						  PushButton 385, 285, 75, 15, "DHS - 3366", open_dhs_3366_btn

						  mf_step_ep_viewed = True
						  'ADD BUTTON DHS 3366'
						ENd If

						'NINTH - CCAP'
						If MFIP_orientation_step = mf_step_ccap Then
						  GroupBox 10, 10, 450, 150, "Availability of Childcare Assistance"
						  Text 25, 25, 425, 20, "There are several Childcare Assistance programs (CCAP) available to support your participation in employment, pre-employment activities, training, and/or educational programs"
						  Text 40, 45, 120, 10, "- MFIP/DWP Childcare assistance"
						  Text 40, 55, 120, 10, "- Transition Year Childcare assistance"
						  Text 60, 65, 385, 25, "Many families continue to be eligible for childcare assistance when their MFIP case closes.  It's highly recommended that you speak to your assigned childcare worker to discuss eligibility details specific to your continued needs for assistance when MFIP closes"
						  Text 40, 90, 170, 10, "- Transition Year Extension Childcare assistance"
						  Text 40, 100, 170, 10, "- Basic sliding fee Childcare assistance"
						  Text 60, 110, 215, 10, "If funds are not available, you may be put on a waiting list"
						  Text 25, 125, 430, 10, "Contact your assigned Employment Counselor or Childcare Assistance Worker to discuss eligibility requirements in more detail."
						  Text 25, 140, 395, 10, "If you need help locating childcare provider options, here's a great resource to contact Think Small or (651-641-0332)"
						  GroupBox 10, 165, 450, 65, "Who to Contact about Childcare Assistance?"
						  Text 25, 180, 420, 20, "If you are receiving MFIP your assigned Employment Counselor will work with you to determine how many childcare hours need to be approved based on the activities in your Employment Plan"
						  Text 25, 205, 420, 20, "If you are receiving MFIP but have not been assigned to an Employment Counselor or if your MFIP has closed contact the childcare assistance line directly at 612-348-5937"
						  GroupBox 10, 235, 450, 115, "Program Compliance and Unavailability of Childcare Assistance"
						  Text 25, 250, 425, 20, "The county may NOT impose a sanction for failure to comply with program requirements if you have good cause because of the unavailability of childcare. The inability to obtain childcare does not exempt or extend your TANF time limit."
						  Text 25, 275, 105, 10, "Some good cause reasons are:"
						  Text 35, 285, 135, 10, "- Unavailability of appropriate childcare"
						  Text 35, 295, 135, 10, "- Unreasonable distance to childcare provider"
						  Text 35, 305, 235, 10, "- Provider does not meet health and safety standards for the child(ren)"
						  Text 35, 315, 275, 10, "- The provider charges an excess amount above the maximum the county can pay"
						  Text 25, 330, 335, 10, "Your Childcare Worker or Employment Counselor can discuss good cause reasons in more detail"
							  Text 10, 365, 145, 10, "Caregiver: " & CAREGIVER_ARRAY(memb_name_const, caregiver)

						  mf_step_ccap_viewed = True
						End If

						'TENTH - Incentives'
						If MFIP_orientation_step = mf_step_incentives Then
						  GroupBox 10, 10, 450, 150, "Incentives and Tax Credits"
						  Text 25, 30, 230, 10, "The MFIP program is designed to benefit you when you are working."
						  Text 25, 45, 420, 35, "For example, your financial worker will not budget all your earned income when they calculate the amount of cash and food benefits you are eligible for. When determining your benefit amount, they will not count the first $65 of income you earn AND beyond that, they will only count half of your remaining gross earned income. Here is a link to explain how this works: Bulletin 21-11-01 - DHS Reissues 'Work Will Always Pay ... With MFIP'"
						  Text 25, 85, 425, 10, "If you are working, when you file your taxes apply for the Earned Income Credit and the Minnesota Working Family Credit."
						  Text 25, 100, 225, 10, "Getting a tax refund will NOT affect your eligibility for MFIP."
						  Text 25, 115, 425, 35, "Have your taxes done for FREE! For a list of free tax preparation sites call the Minnesota Department of Revenue at 651-296-3781 or 1-800-652-9094. Neighborhood Volunteer Income Tax Assistance (VITA) sites are available throughout the state. They are open from February 1 through April 15. Some sites are open year around to help you file back taxes. Search for free tax preparation sites at Department of Revenue."
							  Text 10, 165, 145, 10, "Caregiver: " & CAREGIVER_ARRAY(memb_name_const, caregiver)
						  PushButton 385, 165, 75, 15, "DHS Bulletin 21-11-01", open_dhs_bulletin_21_11_01_btn

						  mf_step_incentives_viewed = True
						  'ADD BUTTON BULLETIN 21-11-01'
						End If

						'ELEVENTH - Health Care'
						If MFIP_orientation_step = mf_step_hc Then
						  GroupBox 10, 10, 450, 90, "Health Care"
						  Text 25, 30, 230, 10, "You may qualify for Minnesota Health Care programs."
						  Text 25, 45, 410, 20, "You can apply for health care online at www.mnsure.org (for assistance completing an online application call 1-855-366-7873) or we can mail you a paper application (DHS 6696)."
						  Text 25, 70, 425, 20, "For help with age-appropriate preventive health services check out the Child and Teen Checkup program at: http://edocs.dhs.state.mn.us/lfserver/public/DHS-1826-ENG"
							  Text 10, 105, 145, 10, "Caregiver: " & CAREGIVER_ARRAY(memb_name_const, caregiver)
						  PushButton 385, 105, 75, 15, "DHS - 1826", open_dhs_1826_btn

						  mf_step_hc_viewed = True
						  'ADD BUTTON DHS 1826'
						End If

						If MFIP_orientation_step = mf_completion Then
						  GroupBox 10, 10, 450, 140, "Document MFIP Orientation Completion"
						  Text 20, 30, 135, 10, "For CAREGIVER NAME:"
						  Text 25, 50, 215, 10, "Did you verbally review all information in the MFIP Oreientation?"
						  DropListBox 240, 45, 210, 45, "Select One..."+chr(9)+"Yes - all information has been reviewed"+chr(9)+"No - could not complete", CAREGIVER_ARRAY(orientation_done_const, caregiver)
						  Text 25, 65, 240, 10, "Notes from any questions/conversation during the MFIP Orientation:"
						  EditBox 25, 75, 425, 15, CAREGIVER_ARRAY(orientation_notes, caregiver)
						  Text 25, 105, 125, 10, "IF COMPLETE - OPEN ECF NOW"
						  Text 35, 120, 220, 10, "Complete the ESP Choice Sheet (D387) with the resident now."
						  Text 35, 135, 175, 10, "Confirm Choice Sheet Completed and saved to ECF:"
						  DropListBox 205, 130, 140, 45, "Select One..."+chr(9)+"Yes - Choice Sheet Saved to ECF"+chr(9)+"No - could not complete", CAREGIVER_ARRAY(choice_form_done_const, caregiver)
						  Text 210, 155, 250, 25, "MFIP Orientation is now complete for this resident. If this case has a second caregiver that requires the MFIP Orientation, this dialog will reappear for the next caregiver as this is a person based process."
						  PushButton 385, 180, 75, 15, "HSR Manual", open_hsr_manual_btn
						  PushButton 385, 195, 75, 15, "CM 05.12.12.06", cm_05_12_12_06_btn

						  mf_completion_viewed = True
						End If

						Text 470, 5, 80, 10, "MFIP Orientation Topics"
						Text 10, 360, 190, 20, "The entire MFIP Orientation to Financial Serviews Script can be viewed on Sharepoint - Open Word Document here:"

						If MFIP_orientation_step = mf_step_rights_resp Then 	Text 500, 18, 55, 10, "Rights / Resp"
						If MFIP_orientation_step = mf_step_time_limits Then 	Text 504, 33, 55, 10, "Time Limits"
						If MFIP_orientation_step = mf_step_extension Then 	Text 509, 48, 55, 10, "Extention"
						If MFIP_orientation_step = mf_step_dv Then 			Text 497, 63, 55, 10, "Family Violence"
						If MFIP_orientation_step = mf_step_expectations Then 	Text 503, 78, 55, 10, "Expectations"
						If MFIP_orientation_step = mf_step_esp Then 			Text 508, 93, 55, 10, "MFIP ESP"
						If MFIP_orientation_step = mf_step_compliance Then 	Text 497, 108, 55, 10, "ES Compliance"
						If MFIP_orientation_step = mf_step_ep Then 			Text 505, 123, 55, 10, "Emplmt Plan"
						If MFIP_orientation_step = mf_step_ccap Then 			Text 512, 138, 55, 10, "CCAP"
						If MFIP_orientation_step = mf_step_incentives Then 	Text 506, 153, 55, 10, "Incentives"
						If MFIP_orientation_step = mf_step_hc Then 			Text 505, 168, 55, 10, "Health Care"
						If MFIP_orientation_step = mf_completion Then 		Text 502, 188, 55, 10, "Confirmation"


					    If MFIP_orientation_step = mf_completion Then PushButton 495, 365, 50, 15, "DONE", done_btn


					    If MFIP_orientation_step <> mf_step_rights_resp Then 	PushButton 495, 15, 55, 15, "Rights / Resp", button_one
					    If MFIP_orientation_step <> mf_step_time_limits Then 	PushButton 495, 30, 55, 15, "Time Limits", button_two
						If MFIP_orientation_step <> mf_step_extension Then 		PushButton 495, 45, 55, 15, "Extention", button_three
					    If MFIP_orientation_step <> mf_step_dv Then 			PushButton 495, 60, 55, 15, "Family Violence", button_four
					    If MFIP_orientation_step <> mf_step_expectations Then 	PushButton 495, 75, 55, 15, "Expectations", button_five
					    If MFIP_orientation_step <> mf_step_esp Then 			PushButton 495, 90, 55, 15, "MFIP ESP", button_six
					    If MFIP_orientation_step <> mf_step_compliance Then 	PushButton 495, 105, 55, 15, "ES Compliance", button_seven
					    If MFIP_orientation_step <> mf_step_ep Then 			PushButton 495, 120, 55, 15, "Emplmt Plan", button_eight
					    If MFIP_orientation_step <> mf_step_ccap Then 			PushButton 495, 135, 55, 15, "CCAP", button_nine
					    If MFIP_orientation_step <> mf_step_incentives Then 	PushButton 495, 150, 55, 15, "Incentives", button_ten
					    If MFIP_orientation_step <> mf_step_hc Then 			PushButton 495, 165, 55, 15, "Health Care", button_eleven
					    If MFIP_orientation_step <> mf_completion Then 			PushButton 495, 185, 55, 15, "Confirmation", button_twelve

					    ' PushButton 495, 195, 55, 15, "Button 2", Button13
					    ' PushButton 495, 210, 55, 15, "Button 2", Button14
					    ' PushButton 495, 225, 55, 15, "Button 2", Button15
					    ' PushButton 495, 240, 55, 15, "Button 2", Button16
						PushButton 205, 360, 135, 15, "MFIP Oriendation Document", mfip_orientation_word_doc_btn
						' OkButton 495, 365, 50, 15

					EndDialog

					dialog Dialog1
					cancel_confirmation

					err_msg = ""

					If ButtonPressed = next_btn Then MFIP_orientation_step = MFIP_orientation_step + 1
					If ButtonPressed = button_one Then MFIP_orientation_step = mf_step_rights_resp
					If ButtonPressed = button_two Then MFIP_orientation_step = mf_step_time_limits
					If ButtonPressed = button_three Then MFIP_orientation_step = mf_step_extension
					If ButtonPressed = button_four Then MFIP_orientation_step = mf_step_dv
					If ButtonPressed = button_five Then MFIP_orientation_step = mf_step_expectations
					If ButtonPressed = button_six Then MFIP_orientation_step = mf_step_esp
					If ButtonPressed = button_seven Then MFIP_orientation_step = mf_step_compliance
					If ButtonPressed = button_eight Then MFIP_orientation_step = mf_step_ep
					If ButtonPressed = button_nine Then MFIP_orientation_step = mf_step_ccap
					If ButtonPressed = button_ten Then MFIP_orientation_step = mf_step_incentives
					If ButtonPressed = button_eleven Then MFIP_orientation_step = mf_step_hc
					If ButtonPressed = button_twelve Then MFIP_orientation_step = mf_completion


					If ButtonPressed = mfip_orientation_word_doc_btn Then
						run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/:w:/r/teams/hs-es-manual/_layouts/15/Doc.aspx?sourcedoc=%7BCB2C8281-95F1-45EE-84D8-B2DF617AA62C%7D&file=MFIP%20Orientation%20to%20Financial%20Services.docx"
						MFIP_orientation_step = mf_completion
						orientation_script_document_viewed = True
					End If
					If ButtonPressed = open_dhs_4163_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://edocs.dhs.state.mn.us/lfserver/Public/DHS-4163-ENG"
					If ButtonPressed = open_dhs_3477_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3477-ENG"
					If ButtonPressed = open_dhs_3323_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3323-ENG"
					If ButtonPressed = open_dhs_3366_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3366-ENG"
					If ButtonPressed = open_dhs_bulletin_21_11_01_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_FILE&RevisionSelectionMethod=LatestReleased&Rendition=Primary&allowInterrupt=1&noSaveAs=1&dDocName=dhs-328254"
					If ButtonPressed = open_dhs_1826_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://edocs.dhs.state.mn.us/lfserver/Public/DHS-1826-ENG"

					If ButtonPressed = open_hsr_manual_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/MFIP_Orientation.aspx"
					If ButtonPressed = cm_05_12_12_06_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=CM_0005121206"
					' If ButtonPressed = cm_28_12_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=CM_002812"




					If mf_step_rights_resp_viewed = True and mf_step_time_limits_viewed = True and mf_step_extension_viewed = True and mf_step_dv_viewed = True and mf_step_expectations_viewed = True and mf_step_esp_viewed = True and mf_step_compliance_viewed = True and mf_step_ep_viewed = True and mf_step_ccap_viewed = True and mf_step_incentives_viewed = True and mf_step_hc_viewed = True and mf_completion_viewed = True Then all_mfip_orientation_info_viewed = True
					If orientation_script_document_viewed = True and mf_completion_viewed = True Then all_mfip_orientation_info_viewed = True


					' MsgBox "DONE? - " & CAREGIVER_ARRAY(orientation_done_const, caregiver) & vbCr & "CHOICE SHEET? - " & CAREGIVER_ARRAY(choice_form_done_const, caregiver)
					If all_mfip_orientation_info_viewed = False Then err_msg = err_msg & vbCr & "* You must review the entire MFIP Orientation before continuing."
					If CAREGIVER_ARRAY(orientation_done_const, caregiver) = "Select One..." Then err_msg = err_msg & vbCr & "* Indicate if the MFIP Orientation has been completed."
					If CAREGIVER_ARRAY(orientation_done_const, caregiver) = "Yes - all information has been reviewed" and CAREGIVER_ARRAY(choice_form_done_const, caregiver) = "Select One..." Then err_msg = err_msg & vbCr & "* Indicate if the MFIP ESP Choice Sheet has been completed in ECF."

					If ButtonPressed = done_btn and err_msg <> "" Then MsgBox err_msg
					' If ButtonPressed = done_btn Then MsgBox err_msg
					If ButtonPressed <> done_btn Then err_msg = "HOLD"

				Loop Until all_mfip_orientation_info_viewed = True and err_msg = ""
			End If
			If CAREGIVER_ARRAY(orientation_done_const, caregiver) = "Yes - all information has been reviewed" Then CAREGIVER_ARRAY(orientation_done_const, caregiver) = True
			If CAREGIVER_ARRAY(orientation_done_const, caregiver) = "No - could not complete" Then CAREGIVER_ARRAY(orientation_done_const, caregiver) = False
			If CAREGIVER_ARRAY(choice_form_done_const, caregiver) = "Yes - Choice Sheet Saved to ECF" Then CAREGIVER_ARRAY(choice_form_done_const, caregiver) = True
			If CAREGIVER_ARRAY(choice_form_done_const, caregiver) = "No - could not complete" Then CAREGIVER_ARRAY(choice_form_done_const, caregiver) = False

			'HERE WE HAVE A DIALOG TO GO TO EMPS AND GIVE INSTRUCTION ON HOW TO COMPLETE IT
			If (CAREGIVER_ARRAY(orientation_needed_const, caregiver) = True and CAREGIVER_ARRAY(orientation_done_const, caregiver) = True) or CAREGIVER_ARRAY(orientation_exempt_const, caregiver) = True Then
				Dialog1 = ""
				BeginDialog Dialog1, 0, 0, 281, 185, "Update EMPS Panel"
				  ButtonGroup ButtonPressed
				    PushButton 125, 135, 145, 15, "The EMPS Panel Updates is Complete", emps_update_complete_btn
				  Text 15, 10, 125, 10, "Caregiver: " & CAREGIVER_ARRAY(memb_name_const, caregiver)
				  If CAREGIVER_ARRAY(orientation_needed_const, caregiver) = True Then Text 35, 20, 205, 10, "NEEDS an MFIP Orientation"
				  If CAREGIVER_ARRAY(orientation_exempt_const, caregiver) = True Then Text 35, 20, 205, 10, "Is Exempt from having an MFIP Orientation"
				  If CAREGIVER_ARRAY(orientation_done_const, caregiver) = True Then Text 15, 35, 255, 10, "The MFIP Orientation to Financial Services is Completed"
				  If CAREGIVER_ARRAY(orientation_done_const, caregiver) = False Then Text 15, 35, 255, 10, "The MFIP Orientation to Financial Services is NOT Completed"
				  If CAREGIVER_ARRAY(choice_form_done_const, caregiver) = True Then Text 15, 45, 255, 10, "The ESP Choice Sheet in ECF is Completed"
				  If CAREGIVER_ARRAY(choice_form_done_const, caregiver) = False Then Text 15, 45, 255, 10, "The ESP Choice Sheet in ECF is NOT Completed"

				  Text 15, 65, 260, 10, "This person has met the requirement for the MFIP Orientation."
				  If CAREGIVER_ARRAY(orientation_exempt_const, caregiver) = True Then Text 20, 75, 260, 10, "Exemption Reason: " & CAREGIVER_ARRAY(exemption_reason_const, caregiver)
				  GroupBox 15, 90, 255, 65, "Update EMPS Panel Now"
				  Text 25, 105, 210, 10, "Update panel: EMPS for " & CAREGIVER_ARRAY(memb_name_const, caregiver)
				  Text 30, 115, 45, 10, "Fin Orient Dt: "
				  If CAREGIVER_ARRAY(orientation_done_const, caregiver) = True Then Text 85, 115, 40, 10, date
				  If CAREGIVER_ARRAY(orientation_done_const, caregiver) = False Then Text 85, 115, 40, 10, "__ __ __"
				  Text 45, 125, 35, 10, "Attended: "
				  If CAREGIVER_ARRAY(orientation_done_const, caregiver) = True Then Text 85, 125, 20, 10, "Y"
				  If CAREGIVER_ARRAY(orientation_done_const, caregiver) = False Then Text 85, 125, 20, 10, "N"
				  Text 30, 135, 45, 10, "Good Cause:"
				  If CAREGIVER_ARRAY(orientation_exempt_const, caregiver) = False Then Text 85, 135, 20, 10, "__"
				  If CAREGIVER_ARRAY(orientation_exempt_const, caregiver) = True Then Text 85, 135, 20, 10, CAREGIVER_ARRAY(emps_exemption_code_const, caregiver)
				EndDialog

				dialog Dialog1

				Call start_a_blank_CASE_NOTE

				If CAREGIVER_ARRAY(orientation_done_const, caregiver) = True Then

					Call write_variable_in_CASE_NOTE("MFIP Orientation completed with " & CAREGIVER_ARRAY(memb_name_const, caregiver))
					Call write_bullet_and_variable_in_CASE_NOTE("Orientation Completed on", date)
					Call write_bullet_and_variable_in_CASE_NOTE("Orientation Notes", CAREGIVER_ARRAY(orientation_notes, caregiver))
					If CAREGIVER_ARRAY(choice_form_done_const, caregiver) = True Then Call write_variable_in_CASE_NOTE("* ESP Choice Sheet: Completed in Case File ")
					Call write_variable_in_CASE_NOTE("---")
					Call write_variable_in_CASE_NOTE(CAREGIVER_ARRAY(memb_name_const, caregiver) & " did not meet an exemption from completing an MFIP Orientation")
					Call write_variable_in_CASE_NOTE("---")
					Call write_variable_in_CASE_NOTE(worker_signature)

				ElseIf CAREGIVER_ARRAY(orientation_exempt_const, caregiver) = True Then

					Call write_variable_in_CASE_NOTE(CAREGIVER_ARRAY(memb_name_const, caregiver) & " is Exempt from MFIP Orientation")
					Call write_bullet_and_variable_in_CASE_NOTE("Assessment Completed", date)
					Call write_bullet_and_variable_in_CASE_NOTE("Exemption Reason", CAREGIVER_ARRAY(exemption_reason_const, caregiver))
					Call write_variable_in_CASE_NOTE("---")
					Call write_variable_in_CASE_NOTE(worker_signature)

				End If
				PF3

				call back_to_SELF

			End If



			MsgBox CAREGIVER_ARRAY(memb_name_const, caregiver) & " - DONE"
		Next

	End If


	MsgBox "STOP HERE"

end function






Call complete_MFIP_orientation(CAREGIVER_ARRAY, memb_ref_numb_const, memb_name_const, memb_age_const, memb_is_caregiver, cash_request_const, hours_per_week_const, exempt_from_ed_const, comply_with_ed_const, orientation_needed_const, orientation_done_const, orientation_exempt_const, exemption_reason_const, emps_exemption_code_const, choice_form_done_const, orientation_notes)






































































































' function script_end_procedure_with_error_report(closing_message)
' '--- This function is how all user stats are collected when a script ends.
' '~~~~~ closing_message: message to user in a MsgBox that appears once the script is complete. Example: "Success! Your actions are complete."
' '===== Keywords: MAXIS, MMIS, PRISM, end, script, statistics, stopscript
' 	stop_time = timer
'     send_error_message = ""
' 	If closing_message <> "" AND left(closing_message, 3) <> "~PT" then        '"~PT" forces the message to "pass through", i.e. not create a pop-up, but to continue without further diversion to the database, where it will write a record with the message
'         send_error_message = MsgBox(closing_message & vbNewLine & vbNewLine & "Do you need to send an error report about this script run?", vbSystemModal + vbDefaultButton2 + vbYesNo, "Script Run Completed")
'     End If
'     script_run_time = stop_time - start_time
' 	If is_county_collecting_stats  = True then
' 		'Getting user name
' 		Set objNet = CreateObject("WScript.NetWork")
' 		user_ID = objNet.UserName
'
' 		'Setting constants
' 		Const adOpenStatic = 3
' 		Const adLockOptimistic = 3
'
'         'Determining if the script was successful
'         If closing_message = "" or left(ucase(closing_message), 7) = "SUCCESS" THEN
'             SCRIPT_success = -1
'         else
'             SCRIPT_success = 0
'         end if
'
' 		'Determines if the value of the MAXIS case number - BULK and UTILITIES scripts will not have case number informaiton input into the database
' 		IF left(name_of_script, 4) = "BULK" or left(name_of_script, 4) = "UTIL" then
' 			MAXIS_CASE_NUMBER = ""
' 		End if
'
' 		'Creating objects for Access
' 		Set objConnection = CreateObject("ADODB.Connection")
' 		Set objRecordSet = CreateObject("ADODB.Recordset")
'
' 		'Fixing a bug when the script_end_procedure has an apostrophe (this interferes with Access)
' 		closing_message = replace(closing_message, "'", "")
'
' 		'Opening DB
' 		IF using_SQL_database = TRUE then
'     		objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" & stats_database_path & ""
' 		ELSE
' 			objConnection.Open "Provider = Microsoft.ACE.OLEDB.12.0; Data Source = " & "" & stats_database_path & ""
' 		END IF
'
'         'Adds some data for users of the old database, but adds lots more data for users of the new.
'         If STATS_enhanced_db = false or STATS_enhanced_db = "" then     'For users of the old db
'     		'Opening usage_log and adding a record
'     		objRecordSet.Open "INSERT INTO usage_log (USERNAME, SDATE, STIME, SCRIPT_NAME, SRUNTIME, CLOSING_MSGBOX)" &  _
'     		"VALUES ('" & user_ID & "', '" & date & "', '" & time & "', '" & name_of_script & "', " & script_run_time & ", '" & closing_message & "')", objConnection, adOpenStatic, adLockOptimistic
' 		'collecting case numbers counties
' 		Elseif collect_MAXIS_case_number = true then
' 			objRecordSet.Open "INSERT INTO usage_log (USERNAME, SDATE, STIME, SCRIPT_NAME, SRUNTIME, CLOSING_MSGBOX, STATS_COUNTER, STATS_MANUALTIME, STATS_DENOMINATION, WORKER_COUNTY_CODE, SCRIPT_SUCCESS, CASE_NUMBER)" &  _
' 			"VALUES ('" & user_ID & "', '" & date & "', '" & time & "', '" & name_of_script & "', " & abs(script_run_time) & ", '" & closing_message & "', " & abs(STATS_counter) & ", " & abs(STATS_manualtime) & ", '" & STATS_denomination & "', '" & worker_county_code & "', " & SCRIPT_success & ", '" & MAXIS_CASE_NUMBER & "')", objConnection, adOpenStatic, adLockOptimistic
' 		 'for users of the new db
' 		Else
'             objRecordSet.Open "INSERT INTO usage_log (USERNAME, SDATE, STIME, SCRIPT_NAME, SRUNTIME, CLOSING_MSGBOX, STATS_COUNTER, STATS_MANUALTIME, STATS_DENOMINATION, WORKER_COUNTY_CODE, SCRIPT_SUCCESS)" &  _
'             "VALUES ('" & user_ID & "', '" & date & "', '" & time & "', '" & name_of_script & "', " & abs(script_run_time) & ", '" & closing_message & "', " & abs(STATS_counter) & ", " & abs(STATS_manualtime) & ", '" & STATS_denomination & "', '" & worker_county_code & "', " & SCRIPT_success & ")", objConnection, adOpenStatic, adLockOptimistic
'         End if
'
' 		'Closing the connection
' 		objConnection.Close
' 	End if
'
'     If send_error_message = vbYes Then
'         'dialog here to gather more detail
'         Dialog1 = ""
'         BeginDialog Dialog1, 0, 0, 401, 175, "Report Error Detail"
'           Text 60, 35, 55, 10, MAXIS_case_number
'           ComboBox 220, 30, 175, 45, ""+chr(9)+"BUG - somethng happened that was wrong"+chr(9)+"ENHANCEMENT - somthing could be done better"+chr(9)+"TYPO - gramatical/spelling type errors", error_type
'           EditBox 65, 50, 330, 15, error_detail
'           CheckBox 20, 100, 65, 10, "CASE/NOTE", case_note_checkbox
'           CheckBox 95, 100, 65, 10, "Update in STAT", stat_update_checkbox
'           CheckBox 170, 100, 75, 10, "Problems with Dates", date_checkbox
'           CheckBox 265, 100, 65, 10, "Math is incorrect", math_checkbox
'           CheckBox 20, 115, 65, 10, "TIKL is incorrect", tikl_checkbox
'           CheckBox 95, 115, 65, 10, "MEMO or WCOM", memo_wcom_checkbox
'           CheckBox 170, 115, 75, 10, "Created Document", document_checkbox
'           CheckBox 265, 115, 115, 10, "Missing a place for Information", missing_spot_checkbox
'           EditBox 60, 140, 165, 15, worker_signature
'           ButtonGroup ButtonPressed
'             OkButton 290, 140, 50, 15
'             CancelButton 345, 140, 50, 15
'           Text 10, 10, 300, 10, "Information is needed about the error for our scriptwriters to review and resolve the issue. "
'           Text 5, 35, 50, 10, "Case Number:"
'           Text 125, 35, 95, 10, "What type of error occured?"
'           Text 5, 55, 60, 10, "Explain in detail:"
'           GroupBox 10, 75, 380, 60, "Common areas of issue"
'           Text 20, 85, 200, 10, "Check any that were impacted by the error you are reporting."
'           Text 10, 145, 50, 10, "Worker Name:"
'           Text 25, 160, 335, 10, "*** Remember to leave the case as is if possible. We can resolve error better when in a live case. ***"
'         EndDialog
'
'         Dialog Dialog1
'
'         'sent email here
'         If ButtonPressed = -1 Then
'             bzt_email = "HSPH.EWS.BlueZoneScripts@hennepin.us"
'             subject_of_email = "Script Error -- " & name_of_script & " (Automated Report)"
'
'             full_text = "Error occured on " & date & " at " & time
'             full_text = full_text & vbCr & "Error type - " & error_type
'             full_text = full_text & vbCr & "Script name - " & name_of_script & " was run on Case #" & MAXIS_case_number & " with a runtime of " & script_run_time & " seconds."
'             full_text = full_text & vbCr & "Information: " & error_detail
'             If case_note_checkbox = checked OR stat_update_checkbox = checked OR date_checkbox = checked OR math_checkbox = checked OR tikl_checkbox = checked OR memo_wcom_checkbox = checked OR document_checkbox = checked OR missing_spot_checkbox = checked Then full_text = full_text & vbCr & vbCr & "Script has issues/concerns in the following areas:"
'
'             If case_note_checkbox = checked Then full_text = full_text & vbCr & " - CASE/NOTE"
'             If stat_update_checkbox = checked Then full_text = full_text & vbCr & " - Update in STAT"
'             If date_checkbox = checked Then full_text = full_text & vbCr & " - Dates are incorrect"
'             If math_checkbox = checked Then full_text = full_text & vbCr & " - Math is incorrect"
'             If tikl_checkbox = checked Then full_text = full_text & vbCr & " - TIKL"
'             If memo_wcom_checkbox = checked Then full_text = full_text & vbCr & " - NOTICES (WCOM/MEMO)"
'             If document_checkbox = checked Then full_text = full_text & vbCr & " - The Excel or Word Document"
'             If missing_spot_checkbox = checked Then full_text = full_text & vbCr & " - There is no space to enter particular information"
'
'             full_text = full_text & vbCr & vbCr & "Sent by: " & worker_signature
'
'             If script_run_lowdown <> "" Then full_text = full_text & vbCr & vbCr & "All Script Run Details:" & vbCr & script_run_lowdown
'
'             Call create_outlook_email(bzt_email, "", subject_of_email, full_text, "", true)
'
'             MsgBox "Error Report completed!" & vbNewLine & vbNewLine & "Thank you for working with us for Continuous Improvement."
'         Else
'             MsgBox "Your error report has been cancelled and has NOT been sent to the BlueZone Script Team"
'         End If
'     End If
' 	If disable_StopScript = FALSE or disable_StopScript = "" then stopscript
' end function


' Call MAXIS_case_number_finder(MAXIS_case_number)
'
' Dialog1 = ""
' BeginDialog Dialog1, 0, 0, 126, 55, "Dialog"
'   ButtonGroup ButtonPressed
'     OkButton 70, 30, 50, 15
'   Text 10, 15, 50, 10, "Case Number"
'   EditBox 60, 10, 60, 15, MAXIS_case_number
' EndDialog
'
' Do
'     err_msg = ""
'
'     dialog Dialog1
'     Call validate_MAXIS_case_number(err_msg, "-")
'     If err_msg <> "" Then MsgBox("Please review the foloowing in order for the script to continue:" & vbNewLine & err_msg)
'
' Loop until err_msg = ""
'
' Call start_a_blank_CASE_NOTE
' notes_variable = "03/19 for 01 is BANKED MONTH - Banked Month: 3.; 04/19 for 01 is BANKED MONTH - Banked Month: 4.;"
' bullet_variable = "This is where the bullet would be all the things."
' time_variable = "Now is the time and this is the place."
' order_variable = "Everything in it's place."
'
' Call write_variable_in_CASE_NOTE("*** SNAP approved starting in 03/19 ***")
' Call write_variable_in_CASE_NOTE("* SNAP approved for 03/19")
' Call write_variable_in_CASE_NOTE("    Eligible Household Members: 01,")
' Call write_variable_in_CASE_NOTE("    Income: Earned: $522.00 Unearned: $0.00")
' Call write_variable_in_CASE_NOTE("    Shelter Costs: $0.00")
' Call write_variable_in_CASE_NOTE("    SNAP BENEFTIT: $115.00 Reporting Status: NON-HRF")
' Call write_variable_in_CASE_NOTE("* SNAP approved for 04/19")
' Call write_variable_in_CASE_NOTE("    Eligible Household Members: 01,")
' Call write_variable_in_CASE_NOTE("    Income: Earned: $522.00 Unearned: $0.00")
' Call write_variable_in_CASE_NOTE("    Shelter Costs: $0.00")
' Call write_variable_in_CASE_NOTE("    SNAP BENEFTIT: $115.00 Reporting Status: NON-HRF")
' Call write_bullet_and_variable_in_CASE_NOTE("Notes", notes_variable)
' Call write_variable_in_CASE_NOTE("This is a thing")
' Call write_variable_in_CASE_NOTE("   this is another thing")
' Call write_variable_in_CASE_NOTE("How now brown cow")
' Call write_variable_in_CASE_NOTE("the thing and thing and stuff")
' Call write_variable_in_CASE_NOTE("all the writing")
' Call write_variable_in_CASE_NOTE("blah blah blah")
' Call write_bullet_and_variable_in_CASE_NOTE("BULLET", bullet_variable)
' Call write_bullet_and_variable_in_CASE_NOTE("Time", time_variable)
' Call write_bullet_and_variable_in_CASE_NOTE("Order", order_variable)
'
' Call write_variable_in_CASE_NOTE("H.Lamb/QI")

' script_list_URL = "C:\MAXIS-scripts\Test scripts\Casey\User Group\COMPLETE LIST OF TESTERS.vbs"
' Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
' Set fso_command = run_another_script_fso.OpenTextFile(script_list_URL)
' text_from_the_other_script = fso_command.ReadAll
' fso_command.Close
' Execute text_from_the_other_script

' Call confirm_tester_information


' function MFIP_cert_length_details(verbal_attestation, attestation_verif_array)
' ' This script requires
' '~~~~~ verbal_attestation: BOOLEAN - that idetifies if ANY verbal attestation was used
'
' 	Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, unknown_cash_pending)
' 	ReDim attestation_verif_array(0)
' 	If mfip_case = TRUE Then
' 		Call navigate_to_MAXIS_screen("STAT", "REVW")
' 		EMReadScreen ER_Month, 2, 9, 37
' 		EMReadScreen ER_Year, 2, 9, 43
' 		If snap_case = TRUE Then
' 			EmWriteScreen "X", 5, 58
' 			transmit
' 			EMReadScreen SNAP_ER_Month, 2, 9, 64
' 			EMReadScreen SNAP_ER_Year, 2, 9, 70
' 			transmit
' 		End If
'
' 		Do
' 			err_msg = ""
' 			If verif_by_attestation_yn = "" Then
' 				BeginDialog Dialog1, 0, 0, 256, 30, "Verification by Attestation Details"
' 			Else
' 				dlg_len = 155
' 				If attestation_verif_array(0) <> "" Then dlg_len = dlg_len + ((UBound(attestation_verif_array)+1) * 15)
' 				If verbal_attestation = FALSE Then dlg_len = 100
'
' 				BeginDialog Dialog1, 0, 0, 256, dlg_len, "Verification by Attestation Details"
' 			End If
'
' 			  ButtonGroup ButtonPressed
' 			    Text 10, 15, 140, 10, "Were ANY verifications by ATTESTATION?"
' 			    DropListBox 155, 10, 40, 45, "?"+chr(9)+"Yes"+chr(9)+"No", verif_by_attestation_yn
' 			    If verif_by_attestation_yn = "" Then PushButton 210, 10, 40, 15, "Enter", enter_btn
' 				If verbal_attestation = TRUE Then
' 					Text 10, 35, 240, 20, "Since verbal attestation has been used to approve MFIP for this case, list all the verifications that we did NOT receive full required documentation:"
' 				    ' Text 20, 60, 235, 10, "VERIF DETIL HERE"
'
' 					y_pos = 60
' 					If attestation_verif_array(0) <> "" Then
' 						For each item in attestation_verif_array
' 							Text 20, y_pos, 235, 10, "- " & item
' 							y_pos = y_pos + 15
' 						Next
' 					End If
'
' 					Text 15, y_pos, 90, 10, "Enter a single verification:"
' 					y_pos = y_pos + 10
' 					EditBox 15, y_pos, 230, 15, verif_entry
' 					y_pos = y_pos + 20
' 				    PushButton 145, y_pos, 100, 10, "SAVE THIS VERIFICATION", save_verif_button
' 					y_pos = y_pos + 25
' 				    Text 10, y_pos, 125, 10, "Current or Upcoming Renewal Month:"
' 				    EditBox 140, y_pos - 5, 15, 15, ER_Month
' 				    EditBox 160, y_pos - 5, 15, 15, ER_Year
' 				    Text 180, y_pos, 30, 10, "(MM YY)"
' 					y_pos = y_pos + 20
' 					PushButton 120, y_pos, 130, 15, "Finish Saving Attestation Information", finish_btn
' 				End If
'
' 				If verbal_attestation = FALSE Then
' 					Text 10, 35, 225, 20, "Since all verifications have been received according to 'non-waiver' policy, confirm the renewal month."
' 				    Text 10, 60, 125, 10, "Current or Upcoming Renewal Month:"
' 				    EditBox 140, 55, 15, 15, ER_Month
' 				    EditBox 160, 55, 15, 15, ER_Year
' 				    Text 180, 60, 30, 10, "(MM YY)"
' 				    PushButton 120, 80, 130, 15, "Finish Cert Period Assesment", finish_btn
' 				End If
' 			EndDialog
'
' 			Dialog Dialog1
'
' 			If verif_by_attestation_yn = "Yes" Then verbal_attestation = TRUE
' 			If verif_by_attestation_yn = "No" Then verbal_attestation = FALSE
'
' 			verif_entry = trim(verif_entry)
' 			If verif_entry <> "" Then
' 				If attestation_verif_array(0) = "" Then
' 					next_verif = 0
' 				Else
' 					next_verif = UBound(attestation_verif_array) + 1
' 					ReDim Preserve attestation_verif_array(next_verif)
' 				End If
' 				attestation_verif_array(next_verif) = verif_entry
' 				verif_entry = ""
' 			End If
'
' 			If ButtonPressed <> finish_btn Then err_msg = "LOOP"
'
' 		Loop until err_msg = ""
'
' 		how_far_away_is_the_next_REVW = ""
' 		next_REVW_date = ER_Month & "/1/" & ER_Year
' 		next_REVW_date = DateAdd("d", 0, next_REVW_date)
'
' 		how_far_away_is_the_next_REVW = DateDiff("m", date, next_REVW_date)
' 		MsgBox how_far_away_is_the_next_REVW
'
' 		before_er_cutoff = FALSE
' 		current_day = DatePart("d", date)
' 		If current_day < 16 Then before_er_cutoff = TRUE
'
' 		If verbal_attestation = FALSE AND how_far_away_is_the_next_REVW < 8 Then
'
' 			Call Navigate_to_MAXIS_screen("CASE", "NOTE")
'
'
' 			first_day_of_this_process = #4/15/2021#
'
'             note_row = 5        'these always need to be reset when looking at Case note
'             note_date = ""
'             note_title = ""
' 			previously_set_to_6_months = FALSE
'             Do                  'this do-loop moves down the list of case notes - looking at each row in MAXIS
'                 EMReadScreen note_date, 8, note_row, 6      'reading the date of the row
'                 EMReadScreen note_title, 55, note_row, 25   'reading the header of the note
'                 note_title = trim(note_title)               'trim it down
'
'                 'if the note headers match any of the following then we can know if a face to face is needed or not - then we add that detail to the ARRAY
'                 If trim(note_title) = "MFIP Certification Period set for 6 MONTHS due Verification by Attestation" Then
' 					previously_set_to_6_months = TRUE
' 					Exit Do
' 				End If
' 				If trim(note_title) = "MFIP Certification Period set for 12 MONTHS since Verifs have been Received" Then Exit Do
'
'                 IF note_date = "        " then Exit Do      'if the case is new, we will hit blank note dates and we don't need to read any further
'                 note_row = note_row + 1                     'going to the next row to look at the next notws
'                 IF note_row = 19 THEN                       'if we have reached the end of the list of case notes then we will go to the enxt page of notes
'                     PF8
'                     note_row = 5
'                 END IF
'                 EMReadScreen next_note_date, 8, note_row, 6 'looking at the next note date
'                 IF next_note_date = "        " then Exit Do
'             Loop until datevalue(next_note_date) < first_day_of_this_process 'looking ahead at the next case note kicking out the dates before app'
'
' 		End If
'
' 		' Call start_a_blank_CASE_NOTE
' 		' If verbal_attestation = TRUE Then
' 		' 	Call write_variable_in_CASE_NOTE("MFIP Certification Period set for 6 MONTHS due Verification by Attestation")
' 		' 	Call write_variable_in_CASE_NOTE()
' 		' 	Call write_variable_in_CASE_NOTE()
' 		' 	Call write_variable_in_CASE_NOTE("---")
' 		' 	Call write_variable_in_CASE_NOTE(worker_signature)
' 		' End If
' 		'
' 		' If verbal_attestation = FALSE Then
' 		' 	Call write_variable_in_CASE_NOTE("MFIP Certification Period set for 12 MONTHS since Verifs have been Received")
' 		' 	Call write_variable_in_CASE_NOTE()
' 		' 	Call write_variable_in_CASE_NOTE()
' 		' 	Call write_variable_in_CASE_NOTE("---")
' 		' 	Call write_variable_in_CASE_NOTE(worker_signature)
' 		'
' 		' End If
'
' 	End If
'
'
' end function
'
'
' BeginDialog Dialog1, 0, 0, 256, 170, "Verbal Attestation Details"
'   Text 10, 15, 140, 10, "Were ANY verifications by ATTESTATION?"
'   DropListBox 150, 10, 40, 45, "", verif_by_attestation_yn
'   Text 10, 35, 240, 20, "Since verbal attestation has been used to approve MFIP for this case, list all the verifications that we did NOT receive full required documentation:"
'   Text 20, 60, 235, 10, "VERIF DETIL HERE"
'   Text 15, 75, 90, 10, "Enter a single verification:"
'   EditBox 15, 85, 230, 15, Edit1
'   ButtonGroup ButtonPressed
'     PushButton 145, 105, 100, 10, "SAVE THIS VERIFICATION", save_verif_button
'   Text 10, 130, 125, 10, "Current or Upcoming Renewal Month:"
'   EditBox 140, 125, 15, 15, Edit2
'   EditBox 160, 125, 15, 15, Edit3
'   Text 180, 130, 30, 10, "(MM YY)"
'   ButtonGroup ButtonPressed
'     PushButton 95, 150, 155, 15, "Finish Saving Verbal Attestation Information", finish_btn
' EndDialog
'
' BeginDialog Dialog1, 0, 0, 256, 155, "Verbal Attestation Details"
'   Text 10, 15, 140, 10, "Were ANY verifications by ATTESTATION?"
'   DropListBox 150, 10, 40, 45, "", verif_by_attestation_yn
'   Text 10, 35, 240, 20, "Since verbal attestation has been used to approve MFIP for this case, list all the verifications that we did NOT receive full required documentation:"
'   Text 15, 60, 90, 10, "Enter a single verification:"
'   EditBox 15, 70, 230, 15, Edit1
'   ButtonGroup ButtonPressed
'     PushButton 145, 90, 100, 10, "SAVE THIS VERIFICATION", save_verif_button
'   Text 10, 115, 125, 10, "Current or Upcoming Renewal Month:"
'   EditBox 140, 110, 15, 15, Edit2
'   EditBox 160, 110, 15, 15, Edit3
'   Text 180, 115, 30, 10, "(MM YY)"
'   ButtonGroup ButtonPressed
'     PushButton 95, 135, 155, 15, "Finish Saving Verbal Attestation Information", finish_btn
' EndDialog
'
' MAXIS_case_number = "1529051"
' MAXIS_footer_month = "04"
' MAXIS_footer_year = "21"
'
' Call MFIP_cert_length_details(verbal_attestation, attestation_verif_array)
'
'





























'
'
'
' MY_STANDARD_ARRAY = Array("Chris", "Casey", "Aurelia", "Ronin")
' ' all_the_people_in_my_house = all_the_people_in_my_house & "Casey" & "~"
' all_the_people_in_my_house = "Chris~Casey~Aurelia~Ronin~"
' ' all_the_people_in_my_house = left(all_the_people_in_my_house, len(all_the_people_in_my_house)-1)
' MY_STANDARD_ARRAY = Split(all_the_people_in_my_house, "~")
' 	' Dim CLIENT_ARRAY()
' 	' ReDIm CLIENT_ARRAY(0)
' 	'
' 	' Dim CLIENT_ARRAY_WITH_MORE()
' 	' Dim CLIENT_ARRAY_WITH_MORE(0)
' Const ref_numb_const 	= 0
' Const first_name_const 	= 1
' Const last_name_const	= 2
' Const clt_dob_const 	= 3
' Const clt_ssn_last_four_const 	= 4
'
' Dim ALL_CLT_INFO_ARRAY()
' ReDim ALL_CLT_INFO_ARRAY(clt_ssn_last_four_const, 0)
'
' the_incrementer = 0
' Do
' 	EmReadscreen MEMB_first_name, 15, 6, 65
' 	EMReadScreen MEMB_last_name, 25, 6, 35
' 	EMReadScreen ref_numb
' 	EMReadScreen dob_mo
' 	EMReadScreen dob_day
' 	EMReadScreen dob_yr
' 	EMReadScreen ssn_last_four
' 	' client_string = client_string & MEMB_first_name & " " & MEMB_last_name & "~"
' 		' ReDim Preserve CLIENT_ARRAY(the_incrementer)
' 		' ReDim Preserve CLIENT_ARRAY_WITH_MORE(the_incrementer)
' 		' CLIENT_ARRAY(the_incrementer) = MEMB_first_name & " " & MEMB_last_name
' 		' CLIENT_ARRAY_WITH_MORE(the_incrementer) = "MEMB " & ref_numb & " - " &  CLIENT_ARRAY(the_incrementer) & " DOB: " & dob_mo & "/" & dob_day & "/" & dob_yr & " SSN: xxx-xx-" & ssn_last_four
' 	ReDim Preserve ALL_CLT_INFO_ARRAY(clt_ssn_last_four_const, the_incrementer)
'
' 	ALL_CLT_INFO_ARRAY(0, the_incrementer) = ref_numb
' 	ALL_CLT_INFO_ARRAY(first_name_const, the_incrementer) = MEMB_first_name
' 	ALL_CLT_INFO_ARRAY(last_name_const, the_incrementer) = MEMB_last_name
' 	ALL_CLT_INFO_ARRAY(clt_dob_const, the_incrementer) = dob_mo & "/" & dob_day & "/" & dob_yr
' 	ALL_CLT_INFO_ARRAY(clt_ssn_last_four_const, the_incrementer) = ssn_last_four
'
' 	the_incrementer = the_incrementer + 1
' 	transmit
' 	EmReadscreen memb_check, 7, 24, 2
' Loop until memb_check = "ENTER A"
' ' CLIENT_ARRAY = split(client_string, "~")
'
' MsgBox Join(MY_STANDARD_ARRAY, ", ")
' For each person in MY_STANDARD_ARRAY
' 	MsgBOx person
' Next
' For the_pers = 0 to UBound(MY_STANDARD_ARRAY)
' 	MsgBOx the_pers
' 	MsgBox MY_STANDARD_ARRAY(the_pers)
' Next
' the_pers = 0
' Do
' 	MsgBox MY_STANDARD_ARRAY(the_pers)
' 	the_pers = the_pers + 1
' Loop until the_pers = UBound(MY_STANDARD_ARRAY)
'
'
'
'
'
'
'
'
'
'
'
' For the_pers = 0 to UBound(ALL_CLT_INFO_ARRAY, 2)				'YOU ALWAYS INCREMENT THE 2nd Parameter of the ARRAY because that is the one that has the different information'
' 	MsgBox "MEMB " & ALL_CLT_INFO_ARRAY(ref_numb_const, the_pers)
'
' 	' last name, first name - dob
' 	MsgBox ALL_CLT_INFO_ARRAY(last_name_const, the_pers) & ", " & ALL_CLT_INFO_ARRAY(first_name_const, the_pers) & " - DOB: " & ALL_CLT_INFO_ARRAY(clt_dob_const, the_pers)
'
' 	' dob for MEMB XX - SSN: xxx-xx-____
' 	MsgBox "DOB: " & ALL_CLT_INFO_ARRAY(clt_dob_const, the_pers) & " for MEMB " & ALL_CLT_INFO_ARRAY(ref_numb_const, the_pers) & "- SSN: xxx-xx-" & ALL_CLT_INFO_ARRAY(clt_ssn_last_four_const, the_pers)
' Next
'
'
'
'
'
' function pause_at_certificate_of_understanding()
'     region_known = FALSE        'setting this to start
'     Do
'         EMReadScreen check_for_cert_of_understanding, 28, 2, 28
'         If check_for_cert_of_understanding = "Certificate Of Understanding" Then
'             'go to training region because that is where this thing happens
'             attn            'getting to the primary menu
'             Do
'                 EMReadScreen MAI_check, 3, 1, 33
'                 If MAI_check <> "MAI" then EMWaitReady 1, 1
'             Loop until MAI_check = "MAI"
'
'             If region_known = FALSE Then                        'We only want to look for the region one time - otherwise it would always be Training
'                 region_known = TRUE
'                 EMReadScreen production_status, 7, 6, 15        'looking to see which session was opened
'                 EMReadScreen inquiry_status, 7, 7, 15
'                 EMReadScreen training_status, 7, 8, 15
'                 If production_status = "RUNNING" Then           'Setting a boolean to know which one was opened originally so we can go back to it.
'                     use_prod = TRUE
'                     EMWriteScreen "C", 6, 2     'here we close because otherwise the agreement stays up
'                     transmit
'                 ElseIf inquiry_status = "RUNNING" Then
'                     use_inq = TRUE
'                     EMWriteScreen "C", 7, 2     'here we close because otherwise the agreement stays up
'                     transmit
'                 ElseIf training_status = "RUNNING" Then
'                     use_trn = TRUE
'                 End If
'             End If
'
'             EMWriteScreen "3", 2, 15                        'actually going into training region'
'             transmit
'
'             'Now we stop the script with a dialog so that the user can still interact with MAXIS
'             Dialog1 = ""
'             BeginDialog Dialog1, 0, 0, 211, 155, "MAXIS Certificate of Understanding"
'               ButtonGroup ButtonPressed
'                 OkButton 155, 135, 50, 15
'               Text 5, 5, 135, 15, "It appears it is time for you to review your MAXIS agreement to maintain access."
'               Text 5, 25, 125, 25, "This annual agreement details of using this system in line with privacy and confidentiality requirements."
'               Text 5, 60, 200, 10, "*** YOU MUST READ AND REVIEW THIS INFORMATION ***"
'               GroupBox 5, 75, 200, 55, "Instructions"
'               Text 15, 90, 175, 35, "Leave this dialog up and read the MAXIS screen currently displayed. Enter your agreement selection. Once this is completed, press 'OK' on this dialog and the script will continue. "
'             EndDialog
'
'             Dialog Dialog1                                  'showing the dialog here
'             cancel_without_confirmation
'             'If ButtonPressed = 0 Then stopscript
'         End If
'     Loop until check_for_cert_of_understanding <> "Certificate Of Understanding"    'we keep showing the dialog until this is done
'     If region_known = TRUE Then
'         'Now we are going back to the region we started in.
'         attn
'         Do
'             EMReadScreen MAI_check, 3, 1, 33
'             If MAI_check <> "MAI" then EMWaitReady 1, 1
'         Loop until MAI_check = "MAI"
'         EMWriteScreen "C", 8, 2
'         transmit
'
'         If use_prod = TRUE Then EMWriteScreen "1", 2, 15
'         If use_inq = TRUE Then EMWriteScreen "2", 2, 15
'         If use_trn = TRUE Then EMWriteScreen "3", 2, 15
'         transmit
'     End If
' end function
' EMConnect ""
'
' Call pause_at_certificate_of_understanding
' MsgBox "Moving On"
'
' ' employer_check = MsgBox("Do you have income verification for this job? Employer name: " & "FAMILY DOLLAR", vbYesNo + vbQuestion, "Select Income Panel")
' '
' ' employer_ended_msg = MsgBox("This job has an income end date." & vbNewLine & vbNewLine & "The employer name: FAMILY DOLLAR" & vbNewLine & "End Date: 12/31/19" & vbNewLine & vbNewLine & "The script can update this job with information provided BUT it will remove the 'End Date' field on JOBS." & vbNewLine & vbNewLine & "Would you like to continue with the update of this job?", vbquestion + vbOkCancel, "Income Panel Ended - Cannot Update")
' ' 'Initial Dialog which requests a file path for the excel file
' ' Dialog1 = ""
' ' BeginDialog Dialog1, 0, 0, 361, 105, "On Demand Recertifications"
' '   EditBox 130, 60, 175, 15, recertification_cases_excel_file_path
' '   ButtonGroup ButtonPressed
' '     PushButton 310, 60, 45, 15, "Browse...", select_a_file_button
' '   EditBox 75, 85, 140, 15, worker_signature
' '   ButtonGroup ButtonPressed
' '     OkButton 250, 85, 50, 15
' '     CancelButton 305, 85, 50, 15
' '   Text 10, 10, 170, 10, "Welcome to the On Demand Recertification Notifier."
' '   Text 10, 25, 340, 30, "This script will send an Appointment Notice or NOMI for recertification for a list of cases in a county that currently has an On Demand Waiver in effect for interviews. If your county does not have this waiver, this script should not be used."
' '   Text 10, 65, 120, 10, "Select an Excel file for recert cases:"
' '   Text 10, 90, 60, 10, "Worker Signature"
' ' EndDialog
'
'
' 'Confirmation Diaglog will require worker to afirm the appointment notices/NOMIs should actually be sent
'
' 'END DIALOGS ===============================================================================================================
'
' 'SCRIPT ====================================================================================================================
' 'Connects to BlueZone
EMConnect ""

Call MAXIS_case_number_finder(MAXIS_case_number)
'
' Call navigate_to_MAXIS_screen_review_PRIV("STAT", "MEMB", memb_priv)
'
' Call navigate_to_MAXIS_screen_review_PRIV("STAT", "JOBS", jobs_priv)
'
' Call navigate_to_MAXIS_screen_review_PRIV("CASE", "CURR", curr_priv)
'
' Call navigate_to_MAXIS_screen_review_PRIV("ELIG", "FS  ", fs_priv)
'
' call script_end_procedure("DONE")
'
Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type, case_status, list_active_programs, list_pending_programs)
Call script_end_procedure("Case Information" & vbNewLine & vbNewLine & "Case Active - " & case_active & vbNewLine & "Case Pending - " & case_pending & vbNewLine & "Case REIN - " & case_rein & vbNewLine & "Family Cash - " & family_cash_case & vbNewLine &_
       "MFIP - " & mfip_case & vbNewLine & "DWP - " & dwp_case & vbNewLine & "Adult Cash - " & adult_cash_case & vbNewLine & "GA - " & ga_case & vbNewLine & "MSA - " & msa_case & vbNewLine & "GRH - " & grh_case & vbNewLine &_
       "SNAP - " & snap_case & vbNewLine & "MA - " & ma_case & vbNewLine & "MSP - " & msp_case  & vbNewLine & "EMER - " & emer_case & vbNewLine & "CASH Pend - " & unknown_cash_pending & vbNewLine & "HC Pend - " & unknown_hc_pending & vbNewLine & "GA Status - " & ga_status & vbNewLine &_
	   "MSA Status - " & msa_status & vbNewLine & "MFIP Status - " & mfip_status & vbNewLine & "DWP Status - " & dwp_status & vbNewLine & "GRH Status - " & grh_status & vbNewLine & "SNAP Status - " & snap_status & vbNewLine &_
	   "MA Status - " & ma_status & vbNewLine & "MSP Status - " & msp_status & vbNewLine & "MSP Type - " & msp_type & vbNewLine & "EMER Status - " & emer_status & vbNewLine & "EMER Type - " & emer_type & vbNewLine & "Case Status - " & case_status & vbCr & vbCr &_
	   "ACTIVE Programs: " & list_active_programs & vbCr & vbCr & "PENDING Programs: " & list_pending_programs)
' 'Initial Dialog to determine the excel file to use, column with case numbers, and which process should be run
' 'Show initial dialog
' Do
' 	Dialog Dialog1
' 	If ButtonPressed = cancel then stopscript
' 	If ButtonPressed = select_a_file_button then call file_selection_system_dialog(recertification_cases_excel_file_path, ".xlsx")
' Loop until ButtonPressed = OK and recertification_cases_excel_file_path <> "" and worker_signature <> ""
'
' 'Opens Excel file here, as it needs to populate the dialog with the details from the spreadsheet.
' call excel_open(recertification_cases_excel_file_path, True, True, ObjExcel, objWorkbook)
'
' 'Set objWorkSheet = objWorkbook.Worksheet
' For Each objWorkSheet In objWorkbook.Worksheets
' 	If instr(objWorkSheet.Name, "Sheet") = 0 and objWorkSheet.Name <> "controls" then scenario_list = scenario_list & chr(9) & objWorkSheet.Name
' Next
'
' 'Dialog to select worksheet
' 'DIALOG is defined here so that the dropdown can be populated with the above code
' Dialog1 = ""
' BeginDialog Dialog1, 0, 0, 151, 75, "On Demand Recertifications"
'   DropListBox 5, 35, 140, 15, "Select One..." & scenario_list, scenario_dropdown
'   ButtonGroup ButtonPressed
'     OkButton 40, 55, 50, 15
'     CancelButton 95, 55, 50, 15
'   Text 5, 10, 130, 20, "Select the correct worksheet to run for recertification interview notifications:"
' EndDialog
'
' 'Shows the dialog to select the correct worksheet
' Do
'     Dialog Dialog1
'     If ButtonPressed = cancel then stopscript
' Loop until scenario_dropdown <> "Select One..."
'
' objExcel.worksheets(scenario_dropdown).Activate
'
' excel_row = 2
' leave_loop = FALSE
' Do
'     Call back_to_SELF
'     MAXIS_case_number = trim(ObjExcel.Cells(excel_row, 2).value)
'
'     If MAXIS_case_number <> "" Then
'         Call navigate_to_MAXIS_screen("STAT", "REVW")
'
'         EMReadScreen cash_revw_status, 1, 7, 40
'         EMReadScreen snap_revw_status, 1, 7, 60
'         EMReadScreen hc_revw_status, 1, 7, 73
'
'         If cash_revw_status = "U" Then leave_loop = TRUE
'         If snap_revw_status = "U" Then leave_loop = TRUE
'         If hc_revw_status = "U" Then leave_loop = TRUE
'         If cash_revw_status = "A" Then leave_loop = TRUE
'         If snap_revw_status = "A" Then leave_loop = TRUE
'         If hc_revw_status = "A" Then leave_loop = TRUE
'
'     Else
'         leave_loop = TRUE
'     End If
'     MAXIS_case_number = ""
'     excel_row = excel_row + 1
'
' Loop until leave_loop = TRUE

Call script_end_procedure_with_error_report("The End")
