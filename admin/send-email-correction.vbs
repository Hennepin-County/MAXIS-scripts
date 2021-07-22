'Required for statistical purposes==========================================================================================
name_of_script = "ADMIN - Send Correction Email.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 0          'manual run time in seconds
STATS_denomination = "I"        'C is for each case
'END OF stats block=========================================================================================================

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

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("03/01/2021", "Added option to Expedited Review for insufficient annotations/interview notes in ECF.", "Casey Love, Hennepin County")
call changelog_update("01/28/2021", "Added the Expedited Retrain Your Brain Video links as resources to be sent on email corrections.", "Casey Love, Hennepin County")
call changelog_update("11/12/2020", "Updated all SharePoint hyperlinks due to SharePoint Online Migration.", "Ilse Ferris, Hennepin County")
call changelog_update("09/21/2020", "Added a checkbox for situations where the interview has been completed but the worker did not address the Adult Cash programs requested. This option is now available for both Expedited SNAP and On Demand options.", "Casey Love, Hennepin County")
call changelog_update("07/29/2020", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'Find who is running
Set objNet = CreateObject("WScript.NetWork")                                    'getting the users windows ID
windows_user_ID = objNet.UserName
user_ID_for_validation = ucase(windows_user_ID)

For each tester in tester_array                                                 'Loop through all the testers in the array to see if the user is in the list of testers.
    If user_ID_for_validation = tester.tester_id_number Then
        qi_worker_full_name            = tester.tester_full_name
        qi_worker_first_name           = tester.tester_first_name
        qi_worker_last_name            = tester.tester_last_name
        qi_worker_email                = tester.tester_email
        qi_worker_id_number            = tester.tester_id_number
        qi_worker_x_number             = tester.tester_x_number
        qi_worker_supervisor           = tester.tester_supervisor_name
        qi_worker_supervisor_email     = tester.tester_supervisor_email
        qi_worker_test_groups          = tester.tester_groups
        qi_staff = FALSE
        For each group in qi_worker_test_groups                                 'looking at all of the groups this tester is a part of to see if QI or BZ
            If group = "QI" Then qi_staff = TRUE
            If group = "BZ" Then qi_staff = TRUE
        Next
    End If
Next
'If this did not find the user is a tester for QI the script will end as this is only for QI staff - access to the files and folders will be restricted and the script will fail
If qi_staff = FALSE Then script_end_procedure_with_error_report("This script is for QI specific processes and only for QI staff. You are not listed as QI staff and running this script could cause errors in data recording and QI processes. Please contact the BlueZone script team or press 'Yes' below if you believe this to be in error.")

EMConnect ""											'connecting to MAXIS
Call MAXIS_case_number_finder(MAXIS_case_number)		'Grabbing the case number if it can find one
'No additional action is taken in MAXIS and password/login is not checked or validated.

email_subject = ""										'setting some variables
email_recipient = ""
email_recipient_cc = ""
email_signature = qi_worker_full_name
send_email = FALSE


'Dialog to select which type of correction needs to be sent
BeginDialog Dialog1, 0, 0, 231, 170, "What type of correction?"
  DropListBox 10, 40, 215, 45, "Select One ..."+chr(9)+"Expedited Review"+chr(9)+"On Demand Applications", correction_process
  EditBox 10, 70, 160, 15, email_recipient
  EditBox 10, 105, 160, 15, email_recipient_cc
  ButtonGroup ButtonPressed
    OkButton 125, 150, 50, 15
    CancelButton 175, 150, 50, 15
    PushButton 5, 155, 70, 10, "INSTRUCTIONS", instructions_btn
  Text 10, 10, 210, 10, "This script can be used for different types of email corrections."
  Text 10, 25, 195, 10, "Select the process you are sending the correction email on:"
  Text 10, 60, 100, 10, "Enter the email of the worker:"
  Text 175, 75, 50, 10, "@hennepin.us"
  Text 10, 95, 120, 10, "Enter the worker's supervisor email:"
  Text 175, 110, 50, 10, "@hennepin.us"
  Text 10, 130, 160, 20, "You do not need to include '@hennepin.us', the script will add that automatically)"
EndDialog

Do
    err_msg = ""
    dialog Dialog1
    cancel_without_confirmation
	email_recipient = trim(email_recipient)
    'Everything is required in this dialog.
	If email_recipient = "" Then
		err_msg = err_msg & vbNewLine & "* Enter the email address of the worker you need to send the correction to."
	ElseIf InStr(email_recipient, ".") = 0 Then
		err_msg = err_msg & vbNewLine & "* The email address entered for the worker does not appear to be a valid email address."
	End If
	If email_recipient_cc = "" Then
		err_msg = err_msg & vbNewLine & "* Enter the email address of the supervisor you need to send the correction to."
	ElseIf InStr(email_recipient_cc, ".") = 0 Then
		err_msg = err_msg & vbNewLine & "* The email address of the supervisor entered does not appear to be a valid email address."
	End If
    If correction_process = "Select One ..." Then err_msg = err_msg & vbNewLine & "* Select which process the correction email is regarding."
	If ButtonPressed = instructions_btn Then
		Call open_URL_in_browser("https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/ADMIN/QI%20-%20SEND%20EMAIL%20CORRECTION.docx")
		err_msg = "LOOP" & err_msg
    ElseIf err_msg <> "" Then
		MsgBox "Please resolve to continue:" & vbNewLine & err_msg
	End If
Loop until err_msg = ""

'formatting the email addresses.
If InStr(email_recipient, "@hennepin.us") = 0 Then email_recipient = email_recipient & "@hennepin.us"
If InStr(email_recipient_cc, "@hennepin.us") = 0 Then email_recipient_cc = email_recipient_cc & "@hennepin.us"

Noon = TimeValue("12:00:00 PM")				'Setting the greeting for the correct time of day
MidAfternoon = TimeValue("4:00:00 PM")
time_of_day = ""

If time < MidAfternoon Then time_of_day = "Afternoon"
If time < Noon Then time_of_day = "Morning"
If time > MidAfternoon Then time_of_day = "Evening"

email_body = "<p>" & "Good " & time_of_day & ", " & "</p>"			'start of the email boday

STATS_manualtime = STATS_manualtime + 45
'Dialog of the email options
Select Case correction_process
	Case "Expedited Review"
		end_msg = "Expedited Review Correction Email sent and tracking updated. Thank you!"

		BeginDialog Dialog1, 0, 0, 550, 395, "Expedited Corrections"
		  EditBox 130, 5, 50, 15, MAXIS_case_number
		  CheckBox 20, 40, 235, 10, "Case was not approved timely. Case should have been approved on ", not_approved_timely_checkbox
		  EditBox 260, 35, 50, 15, should_have_approved_date
		  CheckBox 20, 55, 150, 10, "Identity for more than MEMB 01 requested.", identity_for_more_than_MEMB_01_checkbox_01
		  CheckBox 20, 70, 280, 10, "MEMB 01 Identity was available for SOL-Q or on file. Identity verifcation was found ", identity_was_available_checkbox
		  EditBox 305, 65, 235, 15, identity_verif_found
		  CheckBox 20, 85, 180, 10, "Delayed for verification other than proof of identity.", delayed_for_other_verifs_checkbox
		  CheckBox 20, 100, 260, 10, "Out-of-state benefits reported closing, 2nd month eligibility not determined.", out_of_state_month_two_checkbox
		  CheckBox 20, 130, 235, 10, "The expedited determination was not created or clear in case notes.", expedited_determination_not_done_checkbox
		  CheckBox 20, 145, 155, 10, "The expedited determination was incorrect.", expedited_determination_incorrect_checkbox
		  CheckBox 20, 175, 305, 10, "The Verification Request is incomplete, not sent, or specific information was blank. Details:", verif_request_incomplete_checkbox
		  EditBox 330, 170, 210, 15, verif_request_details
		  CheckBox 20, 190, 150, 10, "Identity for more than MEMB 01 requested.", identity_for_more_than_MEMB_01_checkbox_02
		  CheckBox 20, 205, 240, 10, "Verification requested was not needed for SNAP. What was requested:", verif_request_not_needed_checkbox
		  EditBox 265, 200, 275, 15, what_was_requested
		  CheckBox 20, 220, 195, 10, "Verifications were not postponed for expedited SNAP.", verif_not_delayed_checkbox
		  CheckBox 160, 245, 70, 10, "assets incorrect.", maxis_coded_incorectly_assets_checkbox
		  CheckBox 255, 245, 95, 10, "postponed verifications.", maxis_coded_incorectly_postponed_verif_checkbox
		  CheckBox 150, 260, 130, 10, " processing has not been completed.", interview_complete_processing_not_complete_checkbox
		  CheckBox 285, 260, 120, 10, "CASE/NOTE has not been added.", interview_complete_case_note_missing_checkbox
		  CheckBox 415, 260, 125, 10, "Adult Cash Programs not addressed.", interview_complete_adult_cash_not_addressed_checkbox
		  CheckBox 20, 275, 205, 10, "Case was pended, but expedited screening not conducted.", screening_not_done_checkbox
		  CheckBox 20, 290, 305, 10, "SNAP is pending as MFIP is closing, does not appear to need any mandatory verifications.", snap_pending_after_mfip_closed_checkbox
		  CheckBox 20, 305, 185, 10, "Discrepancy between case notes and MAXIS coding.", maxis_coding_case_note_discrepancy_checkbox
		  CheckBox 20, 320, 110, 10, "The CASE/NOTE is Insufficient.", insufficient_case_note_checkbox
		  CheckBox 20, 335, 225, 10, "Missing or Insufficient annotations or notes about the interview.", insufficient_intv_notes_in_ecf_checkbox
		  EditBox 105, 355, 440, 15, email_notes
		  DropListBox 15, 375, 195, 45, "Indicate if case has been or needs to be updated."+chr(9)+"Case has been updated already - no action needed."+chr(9)+"Please update the case to resolve these issues.", action_needed
		  EditBox 285, 375, 125, 15, email_signature
		  ButtonGroup ButtonPressed
		    OkButton 445, 375, 50, 15
		    CancelButton 495, 375, 50, 15
		  Text 10, 10, 120, 10, "Case number the error occurred on: "
		  Text 205, 10, 230, 10, "Select the Issues/Errors to alert the client about. Check all that apply:"
		  Text 20, 245, 135, 10, "MAXIS panels were not coded correctly - "
		  GroupBox 10, 25, 535, 95, "Case Appears Expedited after Interview but was Not Approved"
		  GroupBox 10, 115, 535, 50, "Expedited to be Determined"
		  Text 15, 360, 90, 10, "Additional Notes for Email:"
		  GroupBox 10, 160, 535, 80, "Verification Request Issues"
		  GroupBox 10, 235, 535, 115, "MAXIS Actions"
		  Text 220, 380, 60, 10, "Sign your Email:"
		  Text 20, 260, 115, 10, "Interview completed, still needed - "
		EndDialog

		Do
			err_msg = ""

			dialog Dialog1
			cancel_confirmation

			email_notes = trim(email_notes)
			email_signature = trim(email_signature)

			'All the counts are required.
			call validate_MAXIS_case_number(err_msg, "*")
			If action_needed = "Indicate if case has been or needs to be updated." Then err_msg = err_msg & vbNewLine & "* Enter if the worker needs to take action or not."
			If email_signature = "" Then err_msg = err_msg & vbNewLine & "* Sign your email."

			If err_msg <> "" Then MsgBox "Please resolve to continue:" & vbNewLine & err_msg

		Loop until err_msg = ""

		'This is in two areas but has the same information. '
		If identity_for_more_than_MEMB_01_checkbox_01 = checked OR identity_for_more_than_MEMB_01_checkbox_02 = checked Then identity_for_more_than_MEMB_01_checkbox = checked

		email_subject = "EXP SNAP Correction on Case # " & MAXIS_case_number			'Setting the subject for the email

		'NOW WE ARE CREATING THE EMAIL BODY - this is done using HTML tags as a part of the string.'
		counter = 1					'This will add a progressing number to each correction indicated

		'start of the email - it will start the same for each correction options
		email_body = email_body & "<p>" & "Quality Improvement staff have been targeting expedited SNAP cases as part of our ongoing effort to reduce errors, improve our timeliness and standardize the application process." & "</p>"
		email_body = email_body & "<p>" & "It appears the expedited SNAP guidelines have not been followed when you processed this case. This is not a performance measurement, our goal is to ensure that every case is able to be given the highest quality care. Please review the following areas and let us know if you have any questions about our assessment or if you feel there is something we may not be understanding. We appreciate your time and any feedback you might have." & "</p>"

		'This adds verbiage to the email that indicates if action is needed or not based on the entry in the dialog.
		If action_needed = "Case has been updated already - no action needed." Then email_body = email_body & "<p>" & "This case has already been updated with correct information and actions. No additional action is needed from you at this time." & "</p>"
		If action_needed = "Please update the case to resolve these issues." Then email_body = email_body & "<p style=" & chr(34) & "color:red" & chr(34) & "><b><u>" & "Updates needed on this case. Please take appropriate action to correct the errors." & "</u></b></p>"

		'Here starts the logic for adding the correction specific verbiage and identifying which resources are needed.
		email_body = email_body & "<p style=" & chr(34) & "font-size:20px" & chr(34) & "><b>" & "Issues/Errors Found on Case # " & MAXIS_case_number & "</b></p>"

		If not_approved_timely_checkbox = checked OR identity_for_more_than_MEMB_01_checkbox = checked OR identity_was_available_checkbox = checked OR delayed_for_other_verifs_checkbox = checked OR out_of_state_month_two_checkbox = checked Then
			email_body = email_body & "<p><i>" & "&emsp;" & counter &  ". This case appears Expedited after completion of the interview. Processing was incorrect because:" & "</i><br>"
			If not_approved_timely_checkbox = checked Then
				If trim(should_have_approved_date) = "" Then
					email_body = email_body &  "&emsp;&ensp;" & "- Case was not approved timely." & "<br>"
				Else
					email_body = email_body & "&emsp;&ensp;" & "- Case was not approved timely. Case should have been approved on " & should_have_approved_date & "." & "<br>"
				End If
				email_body = email_body & "&emsp;&emsp;" & "* Case should have been approved immediately following the Interview if determined expedited." & "<br>"
				add_exp_timeliness_video_resource = TRUE
				STATS_manualtime = STATS_manualtime + 30
			End If
			If identity_for_more_than_MEMB_01_checkbox = checked Then
				email_body = email_body & "&emsp;&ensp;" & "      - Identity for more than MEMB 01 requested" & "<br>"
				add_snap_id_doc_resource = TRUE
				add_cm_10_18_02_resource = TRUE
				add_script_cit_id_resource = TRUE
				STATS_manualtime = STATS_manualtime + 30
			End If
			If identity_was_available_checkbox = checked Then
				If trim(identity_verif_found) = "" Then
					email_body = email_body & "&emsp;&ensp;" & "- Delayed due to Proof of Identity for MEMB 01. Proof of Identity was available through SOLQ-1 or on file." & "<br>"
				Else
					email_body = email_body & "&emsp;&ensp;" & "- Delayed due to Proof of Identity for MEMB 01. Proof of Identity was available through SOLQ-1 or on file. Identity was found: " & identity_verif_found & "<br>"
				End If
				email_body = email_body & "&emsp;&emsp;" & "* When searching for ID in the ECF file, expand your filter to see all contents of the ECF folder." & "<br>"
				add_snap_id_doc_resource = TRUE
				add_cm_10_18_02_resource = TRUE
				add_script_cit_id_resource = TRUE
				add_exp_identity_video_resource = TRUE
				STATS_manualtime = STATS_manualtime + 30
			End If
			If delayed_for_other_verifs_checkbox = checked Then
				email_body = email_body & "&emsp;&ensp;" & "- Delayed for verifications other than proof of identity." & "<br>"
				add_snap_id_doc_resource = TRUE
				add_script_cit_id_resource = TRUE
				add_exp_identity_video_resource = TRUE
				STATS_manualtime = STATS_manualtime + 30
			End If
			If out_of_state_month_two_checkbox = checked Then
				email_body = email_body & "&emsp;&ensp;" & "- Resident reports out-of-state benefits have closed/are closing for the end of the month, and 2nd month eligibility has not been determined." & "<br>"
				add_temp_02_10_79_resource = TRUE
				STATS_manualtime = STATS_manualtime + 30
			End If

			add_cm_04_04_resource = TRUE
			add_cm_04_06_resource = TRUE

			add_script_app_progs_resource = TRUE
			add_script_caf_resource = TRUE
			add_script_exp_det_resource = TRUE
			add_script_add_wcom_resource = TRUE

			add_temp_02_10_01_resource = TRUE

			email_body = email_body & "</p>"
			counter = counter + 1
		End If

		If expedited_determination_not_done_checkbox = checked Then
			email_body = email_body & "<p><i>" & "&emsp;" & counter & ". An expedited determination was either not created or was unclear in case notes after the interview was completed. " & "</i></p>"
			add_cm_04_04_resource = TRUE
			add_cm_04_06_resource = TRUE
			add_hsr_case_note_guidelines_resource = TRUE
			add_script_caf_resource = TRUE
			add_script_exp_det_resource = TRUE
			counter = counter + 1
			STATS_manualtime = STATS_manualtime + 30
		End If
		If expedited_determination_incorrect_checkbox = checked Then
			email_body = email_body & "<p><i>" & "&emsp;" & counter & ". The expedited determination was incorrect in case notes after the interview was completed. " & "</i></p>"
			add_cm_04_04_resource = TRUE
			add_cm_04_06_resource = TRUE
			add_hsr_case_note_guidelines_resource = TRUE
			add_script_caf_resource = TRUE
			add_script_exp_det_resource = TRUE
			counter = counter + 1
			STATS_manualtime = STATS_manualtime + 30
		End If

		If verif_request_incomplete_checkbox = checked OR verif_request_not_needed_checkbox = checked OR verif_not_delayed_checkbox = checked OR identity_for_more_than_MEMB_01_checkbox = checked Then
			email_body = email_body & "<p><i>" & "&emsp;" & counter & ". Verification Request issues:" & "</i><br>"
			If verif_request_incomplete_checkbox = checked Then
				If trim(verif_request_details) = "" Then
					email_body = email_body & "&emsp;&ensp;" & "- Verification request is incomplete, not sent, or specific information was left blank." & "<br>"
				Else
					email_body = email_body & "&emsp;&ensp;" & "- Verification request is incomplete, not sent, or specific information was left blank. " & verif_request_details & "<br>"
				End If
				email_body = email_body & "&emsp;&emsp;" & "* Verification Requests must be sent to clients through ECF the same day processing is completed." & "<br>"
				STATS_manualtime = STATS_manualtime + 30
			End If
			If identity_for_more_than_MEMB_01_checkbox = checked Then
				email_body = email_body & "&emsp;&ensp;" & "- Identity for more than MEMB 01 requested" & "<br>"
				STATS_manualtime = STATS_manualtime + 30
			End If
			If verif_request_not_needed_checkbox = checked Then
				If trim(what_was_requested) = "" Then
					email_body = email_body & "&emsp;&ensp;" & "- Verification requested was not needed for SNAP program." & "<br>"
				Else
					email_body = email_body & "&emsp;&ensp;" & "- Verification requested was not needed for SNAP program. What was requested: " & what_was_requested & "<br>"
				End If
				show_imediate_app_resource = TRUE
				STATS_manualtime = STATS_manualtime + 30
			End If
			If verif_not_delayed_checkbox = checked Then
				email_body = email_body & "&emsp;&ensp;" & "- Verifications not postponed for expedited SNAP case " & "<br>"
				show_imediate_app_resource = TRUE
				STATS_manualtime = STATS_manualtime + 30
			End If
			If show_imediate_app_resource = TRUE Then email_body = email_body & "&emsp;&emsp;" & "* Case should have been approved immediately following the Interview if determined expedited." & "<br>"

			add_cm_04_06_resource = TRUE
			add_cm_10_resource = TRUE

			add_script_caf_resource = TRUE
			add_script_exp_det_resource = TRUE
			add_script_add_wcom_resource = TRUE
			add_script_verif_needed_resource = TRUE

			add_temp_02_10_01_resource = TRUE

			email_body = email_body & "</p>"
			counter = counter + 1
		End If

		If maxis_coded_incorectly_assets_checkbox = checked OR maxis_coded_incorectly_postponed_verif_checkbox = checked Then
			email_body = email_body & "<p><i>" & "&emsp;" & counter & ". MAXIS panels in STAT were not updated correctly." & "</i><br>"
			If maxis_coded_incorectly_assets_checkbox = checked Then
				email_body = email_body & "&emsp;&ensp;" & "- Asset Panels were not coded correctly" & "<br>"
				add_exp_assets_video_resource = TRUE
				STATS_manualtime = STATS_manualtime + 30
			End If
			If maxis_coded_incorectly_postponed_verif_checkbox = checked Then
				email_body = email_body & "&emsp;&ensp;" & "- Panels for postponed verifications were not coded correctly with a '?' for the verification code." & "<br>"
				add_temp_02_10_01_resource = TRUE
				STATS_manualtime = STATS_manualtime + 30
			End If

			add_script_add_wcom_resource = TRUE

			add_temp_19_152_resource = TRUE

			email_body = email_body & "</p>"
			counter = counter + 1
		End If

		If interview_complete_processing_not_complete_checkbox = checked OR interview_complete_case_note_missing_checkbox = checked OR insufficient_case_note_checkbox = checked OR insufficient_intv_notes_in_ecf_checkbox = checked Then
			email_body = email_body & "<p><i>" & "&emsp;" & counter & ". Interview was complete but follow up processing was not." & "</i><br>"
			If interview_complete_processing_not_complete_checkbox = checked Then
				email_body = email_body & "&emsp;&ensp;" & "- MAXIS processing has not been completed." & "<br>"
				add_exp_timeliness_video_resource = TRUE
			End If
			If interview_complete_case_note_missing_checkbox = checked Then
				email_body = email_body & "&emsp;&ensp;" & "- CASE NOTE is missing." & "<br>"
				add_hsr_case_note_guidelines_resource = TRUE
			End If
			If insufficient_case_note_checkbox = checked Then
				email_body = email_body & "&emsp;&ensp;" &  "- CASE NOTE is insufficient." & "<br>"
				add_hsr_case_note_guidelines_resource = TRUE
			End If
			If insufficient_intv_notes_in_ecf_checkbox = checked Then
				email_body = email_body & "&emsp;&ensp;" &  "- Annotations are missing or incomplete, please complete annotations for the interview on the CAF.." & "<br>"
				add_cm_05_12_03_resource = TRUE
				add_ew_guide_to_CAF_p64_resource = TRUE
				add_annotaction_video_resource = TRUE
			End If
			STATS_manualtime = STATS_manualtime + 45

			add_cm_04_04_resource = TRUE
			add_cm_04_06_resource = TRUE

			add_script_caf_resource = TRUE
			add_script_exp_det_resource = TRUE

			email_body = email_body & "</p>"
			counter = counter + 1
		End If

		If interview_complete_adult_cash_not_addressed_checkbox = checked Then
			email_body = email_body & "<p><i>" & "&emsp;" & counter & ". The interview should cover all programs requested. Questions and details about adult cash programs (GA/MSA) were not covered in your interview or CASE:NOTE." & "<br>"
			add_client_contact_interview_qs_resource = TRUE
			add_cm_05_12_12_resource = TRUE

			STATS_manualtime = STATS_manualtime + 30

			email_body = email_body & "</p>"
			counter = counter + 1
		End If

		If screening_not_done_checkbox = checked Then
			email_body = email_body & "<p><i>" & "&emsp;" & counter & ". Case was pended, but an expedited screening was not conducted." & "</i></p>"
			add_cm_04_06_resource = TRUE
			add_hsr_case_note_guidelines_resource = TRUE
			add_script_app_recvd_resource = TRUE
			add_script_exp_screen_resource = TRUE
			add_temp_16_09_resource = TRUE
			STATS_manualtime = STATS_manualtime + 30
			counter = counter + 1
		End If

		If snap_pending_after_mfip_closed_checkbox = checked Then
			email_body = email_body & "<p><i>" & "&emsp;" & counter & ". SNAP is pending as MFIP is closing, and the case is transferring to stand-alone SNAP. SNAP does not appear to be pending for mandatory verifications per case notes." & "</i><br>"
			email_body = email_body & "&emsp;&ensp;" & "* Same day processing should be completed if mandatory verifs are not needed." & "</p>"
			add_script_mf_to_fs_resource = TRUE
			add_temp_02_08_143_resource = TRUE
			STATS_manualtime = STATS_manualtime + 30
			counter = counter + 1
		End If

		If maxis_coding_case_note_discrepancy_checkbox = checked Then
			email_body = email_body & "<p><i>" & "&emsp;" & counter & ". Discrepancy between case notes and MAXIS coding or case note was insufficient." & "</i></p>"
			add_hsr_case_note_guidelines_resource = TRUE
			add_script_caf_resource = TRUE
			add_script_exp_det_resource = TRUE
			STATS_manualtime = STATS_manualtime + 30
			counter = counter + 1
		End If

		If trim(email_notes) <> "" Then
			email_body = email_body & "<p>" & "Additional Information:" & "<br>"
			email_body = email_body & email_notes & "</p>"
			STATS_manualtime = STATS_manualtime + 15
		End If

		'here is where we actually add the resource informtation
		email_body = email_body & "<p><b><i>" & "&ensp;" & "RESOURCES:" & "</b></i><br>"

		If add_cm_10_18_02_resource = TRUE OR add_cm_04_04_resource = TRUE OR add_cm_04_06_resource = TRUE OR add_cm_05_12_12_resource = TRUE OR add_cm_10_resource = TRUE OR add_cm_05_12_03_resource = TRUE Then email_body = email_body & "<i>" & "&emsp;" & "Combined Manual:" & "</i><br>"

		If add_cm_10_18_02_resource = TRUE Then email_body = email_body & "&emsp;&ensp;" & "- " & "<a href=" & chr(34) & "https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=CM_00101802" & chr(34) & ">" & "CM 0010.18.02" & "</a>" & " Mandatory Verifications - SNAP" & "<br>"
		If add_cm_04_04_resource = TRUE Then email_body = email_body & "&emsp;&ensp;" & "- " & "<a href=" & chr(34) & "http://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=cm_000404" & chr(34) & ">" & "CM 0004.04" & "</a>" & " Emergency Aid Eligibility - SNAP/Expedited Processing" & "<br>"
		If add_cm_04_06_resource = TRUE Then email_body = email_body & "&emsp;&ensp;" & "- " & "<a href=" & chr(34) & "http://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=cm_000406" & chr(34) & ">" & "CM 0004.06" & "</a>" & " Emergencies - 1st Month Processing" & "<br>"
		If add_cm_05_12_12_resource = TRUE Then email_body = email_body & "&emsp;&ensp;" & "- " & "<a href=" & chr(34) & "https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=CM_00051212" & chr(34) & ">" & "CM 0005.12.12" & "</a>" & " Application Interviews - 'Offer applicants or their authorized representatives a single interview that covers all the programs for which they apply.'" & "<br>"
		If add_cm_05_12_03_resource = TRUE Then email_body = email_body & "&emsp;&ensp;" & "- " & "<a href=" & chr(34) & "https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=CM_00051203" & chr(34) & ">" & "CM 0005.12.03" & "</a>" & " What is a Complete Application" & "<br>"
		If add_cm_10_resource = TRUE Then  email_body = email_body & "&emsp;&ensp;" & "- " & "<a href=" & chr(34) & "https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=CM_0010" & chr(34) & ">" & "CM 0010" & "</a>" & " Verification" & "<br>"

		If add_ew_guide_to_CAF_p64_resource = TRUE Then email_body = email_body & "<i>" & "&emsp;" & "Other DHS Resources:" & "</i><br>"

		If add_ew_guide_to_CAF_p64_resource = TRUE Then email_body = email_body & "&emsp;&ensp;" & "- " & "<a href=" & chr(34) & "https://www.dhssir.cty.dhs.state.mn.us/MAXIS/trntl/_layouts/15/WopiFrame2.aspx?sourcedoc=%7B3230AF4F-4FA7-448C-BAA7-506671E03A49%7D&file=An%20Eligibility%20Workers%20Guide%20to%20the%20Combined%20Application%20Form%20(With%20Answers).pdf&action=default&IsList=1&ListId=%7B032C9304-E9F4-4ED6-90A0-92F9CC18CD31%7D&ListItemId=2" & chr(34) & ">" & "An Eligibility Workers Guide to the CAF" & "</a>" & " Page 64 - Need SIR signed in to access." & "<br>"

		If add_snap_id_doc_resource = TRUE OR add_hsr_case_note_guidelines_resource = TRUE OR add_client_contact_interview_qs_resource = TRUE Then email_body = email_body & "<i>" & "&emsp;" & "Internal:" & "</i><br>"

		If add_snap_id_doc_resource = TRUE Then email_body = email_body & "&emsp;&ensp;" & "- " & "<a href=" & chr(34) & "https://www.dhssir.cty.dhs.state.mn.us/MAXIS/Documents/SNAP Identity Verification.pdf" & chr(34) & ">" & " SNAP Identity Verification" & "</a><br>"
		If add_hsr_case_note_guidelines_resource = TRUE Then email_body = email_body & "&emsp;&ensp;" & "- HSR Manual: " & "<a href=" & chr(34) & "https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/Guidelines_and_Format.aspx" & chr(34) & ">" & "Case Notes Guidelines and Format" & "</a><br>"
		If add_client_contact_interview_qs_resource = TRUE Then email_body = email_body & "&emsp;&ensp;" & "- Client Contact Documents: " & "<a href=" & chr(34) & "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/Adults%20and%20Families%20Eligibility%20Documents/Cash%20and%20EGA%20interview%20questions.docx" & chr(34) & ">" & "Cash and EGA Interview Questions" & "</a><br>"

		If add_exp_timeliness_video_resource = TRUE OR add_exp_identity_video_resource = TRUE OR add_exp_assets_video_resource = TRUE OR add_annotaction_video_resource = TRUE Then email_body = email_body & "<i>" & "&emsp;" & "Retrain Your Brain Videos:" & "</i><br>"

		If add_exp_identity_video_resource = TRUE Then email_body = email_body & "&emsp;&ensp;" & "- " & "<a href=" & chr(34) & "https://web.microsoftstream.com/video/5639ccc5-02ba-47b9-a99a-fdcc321b20d6?channelId=a52b0e9b-200d-4e2c-bf78-697fd7d5cf2e" & chr(34) & ">" & "Watch 'SNAP EXP 1 - ID' | Microsoft Stream" & "</a><br>"
		If add_exp_timeliness_video_resource = TRUE Then email_body = email_body & "&emsp;&ensp;" & "- " & "<a href=" & chr(34) & "https://web.microsoftstream.com/video/ef749c68-873c-4719-a6ae-ea18748efc04?channelId=a52b0e9b-200d-4e2c-bf78-697fd7d5cf2e" & chr(34) & ">" & "Watch 'SNAP EXP 2 - Timeliness' | Microsoft Stream" & "</a><br>"
		If add_exp_assets_video_resource = TRUE Then email_body = email_body & "&emsp;&ensp;" & "- " & "<a href=" & chr(34) & "https://web.microsoftstream.com/video/be6bd005-b1ad-4f7c-96a6-05868e296fea?channelId=a52b0e9b-200d-4e2c-bf78-697fd7d5cf2e" & chr(34) & ">" & "Watch 'SNAP EXP 3 - Assets' | Microsoft Stream" & "</a><br>"
		If add_annotaction_video_resource = TRUE Then email_body = email_body & "&emsp;&ensp;" & "- " & "<a href=" & chr(34) & "https://web.microsoftstream.com/video/4fb56128-a9ea-4dec-859a-35f8f96c7c8d" & chr(34) & ">" & "Watch 'MN.benefits.org Application Annotations' | Microsoft Stream" & "</a><br>"

		If add_temp_02_10_01_resource = TRUE OR add_temp_02_10_79_resource = TRUE OR add_temp_19_152_resource = TRUE OR add_temp_16_09_resource = TRUE OR add_temp_02_08_143_resource = TRUE Then email_body = email_body & "<i>" & "&emsp;" & "TEMP Manual:" & "</i><br>"

		If add_temp_02_10_01_resource = TRUE Then email_body = email_body & "&emsp;&ensp;" & "- POLI TEMP - TE02.10.01 Expedited SNAP with Pending Verifications" & "<br>"
		If add_temp_02_10_79_resource = TRUE Then email_body = email_body & "&emsp;&ensp;" & "- POLI TEMP - TE.02.10.79 Expedited FS 2nd Month Eligibility" & "<br>"
		If add_temp_19_152_resource = TRUE Then email_body = email_body & "&emsp;&ensp;" & "- POLI TEMP - TE19.152 - QTIP #152 Expedited Food Support" & "<br>"
		If add_temp_16_09_resource = TRUE Then email_body = email_body & "&emsp;&ensp;" & "- POLI TEMP - TE16.09 EBT - Expedited Food Support" & "<br>"
		If add_temp_02_08_143_resource = TRUE Then email_body = email_body & "&emsp;&ensp;" & "- POLI TEMP - TE02.08.143 FOOD SUPPORT WHEN MFIP IS ENDING" & "<br>"

		If add_script_app_progs_resource = TRUE OR add_script_caf_resource = TRUE OR add_script_cit_id_resource = TRUE OR add_script_exp_det_resource = TRUE OR add_script_add_wcom_resource = TRUE OR add_script_verif_needed_resource = TRUE OR add_script_app_recvd_resource = TRUE OR add_script_exp_screen_resource = TRUE OR add_script_mf_to_fs_resource = TRUE Then email_body = email_body & "<i>" & "&emsp;" & "Scripts: (Instructions)" & "</i><br>"

		If add_script_app_progs_resource = TRUE Then email_body = email_body & "&emsp;&ensp;" & "- " & "<a href=" & chr(34) & "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/NOTES/NOTES%20-%20APPROVED%20PROGRAMS.docx" & chr(34) & ">" & "NOTES - APPROVED PROGRAMS" & "</a><br>"
		If add_script_caf_resource = TRUE Then email_body = email_body & "&emsp;&ensp;" & "- " & "<a href=" & chr(34) & "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/NOTES/NOTES%20-%20CAF.docx" & chr(34) & ">" & "NOTES - CAF" & "</a><br>"
		If add_script_cit_id_resource = TRUE Then email_body = email_body & "&emsp;&ensp;" & "- " & "<a href=" & chr(34) & "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/NOTES/NOTES%20-%20CITIZENSHIP%20IDENTITY%20VERIFIED.docx" & chr(34) & ">" & "NOTES - CITIZENSHIP IDENTITY VERIFIED" & "</a><br>"
		If add_script_exp_det_resource = TRUE Then email_body = email_body & "&emsp;&ensp;" & "- " & "<a href=" & chr(34) & "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/NOTES/NOTES%20-%20EXPEDITED%20DETERMINATION.docx" & chr(34) & ">" & "NOTES - EXPEDITED DETERMINATION" & "</a><br>"
		If add_script_add_wcom_resource = TRUE Then email_body = email_body & "&emsp;&ensp;" & "- " & "<a href=" & chr(34) & "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/NOTICES/NOTICES%20-%20ADD%20WCOM.docx" & chr(34) & ">" & "NOTICES - ADD WCOM" & "</a><br>"
		If add_script_verif_needed_resource = TRUE Then email_body = email_body & "&emsp;&ensp;" & "- " & "<a href=" & chr(34) & "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/NOTES/NOTES%20-%20VERIFICATIONS%20NEEDED.docx" & chr(34) & ">" & "NOTES - VERIFICATIONS NEEDED" & "</a><br>"
		If add_script_app_recvd_resource = TRUE Then email_body = email_body & "&emsp;&ensp;" & "- " & "<a href=" & chr(34) & "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/NOTES/NOTES%20-%20APPLICATION%20RECEIVED.docx" & chr(34) & ">" & "NOTES - APPLICATION RECEIVED" & "</a><br>"
		If add_script_exp_screen_resource = TRUE Then email_body = email_body & "&emsp;&ensp;" & "- " & "<a href=" & chr(34) & "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/NOTES/NOTES%20-%20EXPEDITED%20SCREENING.docx" & chr(34) & ">" & "NOTES - EXPEDITED SCREENING" & "</a><br>"
		If add_script_mf_to_fs_resource = TRUE Then email_body = email_body & "&emsp;&ensp;" & "- " & "<a href=" & chr(34) & "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/NOTES/NOTES%20-%20MFIP%20TO%20SNAP%20TRANSITION.docx" & chr(34) & ">" & "NOTES - MFIP TO SNAP TRANSITION" & "</a><br>"
		email_body = email_body & "</p>"

		'End of the message with email llinks
		email_body = email_body & "<p>" & "Thank you for taking the time to review this information. If you have any additional questions or want additional directions and resources, please contact the QI team. For Script questions contact the BlueZone Script Team." & "<br>"
		email_body = email_body & "<a href=" & chr(34) & "mailto:HSPH.EWS.QUALITYIMPROVEMENT@hennepin.us?subject=SNAP%20EXP%20Questions" & chr(34) & ">" & "Email Quality Improvement" & "</a><br>"
		email_body = email_body & "<a href=" & chr(34) & "mailto:HSPH.EWS.BlueZoneScripts@hennepin.us?subject=SNAP%20EXP%20Questions" & chr(34) & ">" & "Email the BlueZone Script Team" & "</a><br>"
		email_body = email_body & "</p>"

		email_body = email_body & "<p style=" & chr(34) & "font-size:20px" & chr(34) & ">" & "From, " & email_signature & "<br>"
		email_body = email_body & "Member of the ES Quality Improvement Team" & "</p>"

		'NOW WE SEND THE EMAIL
		'Setting up the Outlook application
	    Set objOutlook = CreateObject("Outlook.Application")
	    Set objMail = objOutlook.CreateItem(0)
	    objMail.Display                                 'To display message

	    'Adds the information to the email
	    objMail.to = email_recipient                        'email recipient
	    objMail.cc = email_recipient_cc                     'cc recipient
		objMail.SentOnBehalfOfName = "HSPH.EWS.QUALITYIMPROVEMENT@hennepin.us"
        objMail.Importance = 2      'Sending with high importance
		' objMail.SentOnBehalfOfName = "HSPH.EWS.BlueZoneScripts@hennepin.us"
	    objMail.Subject = email_subject                 'email subject
	    objMail.HTMLBody = email_body                       'email body
	    If email_attachment <> "" then objMail.Attachments.Add(email_attachment)       'email attachement (can only support one for now)
	    'Sends email
	    If send_email = true then objMail.Send	                   'Sends the email
	    Set objMail =   Nothing
	    Set objOutlook = Nothing

		'EXPEDITED SNAP Column CONSTANTS
		recip_col 							= 1
		date_col 							= 2
		case_numb_col 						= 3
		sup_email_col 						= 4
		not_app_timely_col 					= 5
		not_app_date_col 					= 6
		id_for_more_than_01_col 			= 7
		id_was_on_file_col 					= 8
		id_found_location_col 				= 9
		delayed_non_id_verif_col 			= 10
		out_of_state_delayed_col 			= 11
		exp_det_note_concern_col 			= 12
		exp_det_incorrect_col 				= 13
		verif_req_incomplete_col 			= 14
		verif_req_incp_details_col 			= 15
		verif_req_not_needed_col 			= 16
		verif_req_not_needed_detail_col 	= 17
		verif_not_postponed_col 			= 18
		maxis_assets_wrong_col 				= 19
		maxis_verif_codes_wrong_col 		= 20
		maxis_not_processed_col 			= 21
		case_note_missing_col 				= 22
		no_intv_adult_progs_col				= 23
		no_exp_screening_col 				= 24
		mf_to_FS_col 						= 25
		case_note_maxis_mismatch_col 		= 26
		case_note_insufficient_col 			= 27
		intv_notes_insufficient_col			= 28
		other_notes_col 					= 29
		worker_to_follow_up_col 			= 30
		qi_worker_col 						= 31

		'HERE IS THE TRAKING PART - so we can save the information about what and who we emailed for corrections
		excel_file_path = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\SNAP\EXP SNAP Project\EXP SNAP Email Corrections Data.xlsx"
		call excel_open(excel_file_path, False,  False, ObjExcel, objWorkbook)  'opening the Excel'
		ObjExcel.worksheets("Data").Activate

		excel_row = 2                                                           'finding the first empty excel row
		Do
			this_entry = ObjExcel.Cells(excel_row, 1).Value
			this_entry = trim(this_entry)
			If this_entry <> "" Then excel_row = excel_row + 1
		Loop until this_entry = ""

		'Adding all the information to the Excel
		ObjExcel.Cells(excel_row, recip_col).Value 							= email_recipient
		ObjExcel.Cells(excel_row, date_col).Value 							= date
		ObjExcel.Cells(excel_row, case_numb_col).Value 						= MAXIS_case_number
		ObjExcel.Cells(excel_row, sup_email_col).Value 						= email_recipient_cc
		If not_approved_timely_checkbox = checked 							Then ObjExcel.Cells(excel_row, not_app_timely_col).Value 				= "X"
		ObjExcel.Cells(excel_row, not_app_date_col).Value 					= should_have_approved_date
		If identity_for_more_than_MEMB_01_checkbox = checked 				Then ObjExcel.Cells(excel_row, id_for_more_than_01_col).Value 			= "X"
		If identity_was_available_checkbox = checked 						Then ObjExcel.Cells(excel_row, id_was_on_file_col).Value 				= "X"
		ObjExcel.Cells(excel_row, id_found_location_col).Value				= identity_verif_found
		If delayed_for_other_verifs_checkbox = checked 						Then ObjExcel.Cells(excel_row, delayed_non_id_verif_col).Value 			= "X"
		If out_of_state_month_two_checkbox = checked 						Then ObjExcel.Cells(excel_row, out_of_state_delayed_col).Value 			= "X"
		If expedited_determination_not_done_checkbox = checked 				Then ObjExcel.Cells(excel_row, exp_det_note_concern_col).Value 			= "X"
		If expedited_determination_incorrect_checkbox = checked 			Then ObjExcel.Cells(excel_row, exp_det_incorrect_col).Value 			= "X"
		If verif_request_incomplete_checkbox = checked 						Then ObjExcel.Cells(excel_row, verif_req_incomplete_col).Value 			= "X"
		ObjExcel.Cells(excel_row, verif_req_incp_details_col).Value 		= verif_request_details
		If verif_request_not_needed_checkbox = checked 						Then ObjExcel.Cells(excel_row, verif_req_not_needed_col).Value 			= "X"
		ObjExcel.Cells(excel_row, verif_req_not_needed_detail_col).Value 	= what_was_requested
		If verif_not_delayed_checkbox = checked 							Then ObjExcel.Cells(excel_row, verif_not_postponed_col).Value 			= "X"
		If maxis_coded_incorectly_assets_checkbox = checked 				Then ObjExcel.Cells(excel_row, maxis_assets_wrong_col).Value 			= "X"
		If maxis_coded_incorectly_postponed_verif_checkbox = checked 		Then ObjExcel.Cells(excel_row, maxis_verif_codes_wrong_col).Value 		= "X"
		If interview_complete_processing_not_complete_checkbox = checked 	Then ObjExcel.Cells(excel_row, maxis_not_processed_col).Value 			= "X"
		If interview_complete_case_note_missing_checkbox = checked 			Then ObjExcel.Cells(excel_row, case_note_missing_col).Value 			= "X"
		If interview_complete_adult_cash_not_addressed_checkbox = checked 	Then ObjExcel.Cells(excel_row, no_intv_adult_progs_col).Value 			= "X"
		If screening_not_done_checkbox = checked 							Then ObjExcel.Cells(excel_row, no_exp_screening_col).Value 				= "X"
		If snap_pending_after_mfip_closed_checkbox = checked 				Then ObjExcel.Cells(excel_row, mf_to_FS_col).Value 						= "X"
		If maxis_coding_case_note_discrepancy_checkbox = checked			Then ObjExcel.Cells(excel_row, case_note_maxis_mismatch_col).Value 		= "X"
		If insufficient_case_note_checkbox = checked 						Then ObjExcel.Cells(excel_row, case_note_insufficient_col).Value 		= "X"
		If insufficient_intv_notes_in_ecf_checkbox = checked 				Then ObjExcel.Cells(excel_row, intv_notes_insufficient_col).Value 		= "X"
		ObjExcel.Cells(excel_row, other_notes_col).Value 					= email_notes
		If action_needed = "Case has been updated already - no action needed." Then ObjExcel.Cells(excel_row, worker_to_follow_up_col).Value 		= "NO"
		If action_needed = "Please update the case to resolve these issues." Then ObjExcel.Cells(excel_row, worker_to_follow_up_col).Value 			= "YES"
		ObjExcel.Cells(excel_row, qi_worker_col).Value 						= qi_worker_full_name

		ObjExcel.ActiveWorkbook.Save                                            'saving and closing the Excel spreadsheet
        ObjExcel.ActiveWorkbook.Close
        ObjExcel.Application.Quit

	Case "On Demand Applications"
		end_msg = "On Demand Application Correction Email sent and tracking updated. Thank you!"

		add_hsr_appling_resource = FALSE
		add_hsr_app_guide_resource = FALSE
		add_hsr_interview_process_resource = FALSE
		add_hsr_nomi_resource = FALSE
		add_hsr_on_demand_resource = FALSE
		add_script_app_recvd_resource = FALSE
		add_script_app_check_resource = FALSE
		add_script_caf_resource = FALSE
		add_script_appr_progs_resource = FALSE
		add_script_clt_contact_resource = FALSE
		add_script_interview_comp_resource = FALSE

		BeginDialog Dialog1, 0, 0, 536, 265, "On Demand Corrections"
		  EditBox 130, 5, 50, 15, MAXIS_case_number
		  CheckBox 20, 25, 295, 10, "The script NOTES - APPLICATION RECEIVED was not used when the case was APPL'd.", app_recvd_script_not_used_checkbox
		  CheckBox 30, 55, 340, 10, "The interview date field was not completed on STAT/PROG, though interview appears to be completed.", prog_not_updated_interview_completed_checkbox
		  CheckBox 30, 70, 130, 10, "Adult Cash Programs not assessed.", intv_adult_cash_not_completed
		  CheckBox 30, 85, 495, 10, "The client applied for Cash and SNAP, Cash needs a face to face but this is not clear in notes, cash information should be reviewed in phone interview.", cash_detail_not_covered_checkbox
		  CheckBox 30, 100, 370, 10, "Cash and SNAP were applied for, Cash needs a face to face, SNAP should have had a phone interview offered.", snap_phone_interview_should_have_been_offered_checkbox
		  CheckBox 30, 135, 260, 10, "Case was denied for no interview from PND2, QI should be doing all of these.", denied_for_no_interview_on_PND2_checkbox
		  CheckBox 30, 150, 170, 10, "Case was denied for no interview NOT on PND2.", denied_for_no_interview_NOT_on_PND2_checkbox
		  CheckBox 20, 170, 305, 10, "The NOTES - CAF script was used on a case when the interview has not been completed.", caf_script_used_no_interview_checkbox
		  CheckBox 20, 185, 360, 10, "Client contacted the agency and should have had an interview offered, even with only Page 1 of the CAF.", client_contacted_should_have_been_offered_interview_checkbox
		  CheckBox 20, 200, 295, 10, "An interview was completed the same day after a NOMI was sent, notice not cancelled.", interview_completed_same_day_NOMI_sent_checkbox
		  EditBox 105, 220, 425, 15, email_notes
		  DropListBox 15, 240, 195, 45, "Indicate if case has been or needs to be updated."+chr(9)+"Case has been updated already - no action needed."+chr(9)+"Please update the case to resolve these issues.", action_needed
		  EditBox 285, 240, 125, 15, email_signature
		  ButtonGroup ButtonPressed
		    OkButton 430, 240, 50, 15
		    CancelButton 480, 240, 50, 15
		  Text 10, 10, 120, 10, "Case Number the Error happend on: "
		  Text 205, 10, 230, 10, "Select the Issues/Errors to alert the client about. Check all that apply:"
		  GroupBox 20, 40, 510, 75, "Interview Issues"
		  GroupBox 20, 120, 505, 45, "Denial for No Interview Issue"
		  Text 15, 225, 90, 10, "Additional Notes for Email:"
		  Text 220, 245, 60, 10, "Sign your Email:"
		EndDialog

		Do
			err_msg = ""

			dialog Dialog1
			cancel_confirmation

			email_notes = trim(email_notes)
			email_signature = trim(email_signature)

			'All the counts are required.
			call validate_MAXIS_case_number(err_msg, "*")
			If action_needed = "Indicate if case has been or needs to be updated." Then err_msg = err_msg & vbNewLine & "* Enter if the worker needs to take action or not."
			If email_signature = "" Then err_msg = err_msg & vbNewLine & "* Sign your email."

			If err_msg <> "" Then MsgBox "Please resolve to continue:" & vbNewLine & err_msg

		Loop until err_msg = ""

		email_subject = "Application Process Correction on Case # " & MAXIS_case_number					'Setting the subject for the email

		'NOW WE ARE CREATING THE EMAIL BODY - this is done using HTML tags as a part of the string.'
		counter = 1 				'This will add a progressing number to each correction indicated

		'start of the email - it will start the same for each correction options
		email_body = email_body & "<p>" & "Quality Improvement staff have been targeting application cases as part of our ongoing effort to reduce errors, improve our timeliness, and standardize the application process." & "</p>"

		'This adds verbiage to the email that indicates if action is needed or not based on the entry in the dialog.
		If action_needed = "Case has been updated already - no action needed." Then email_body = email_body & "<p>" & "This case has already been updated with correct information and actions. No additional action is needed from you at this time." & "</p>"
		If action_needed = "Please update the case to resolve these issues." Then email_body = email_body & "<p style=" & chr(34) & "color:red" & chr(34) & "><b><u>" & "Updates needed on this case. Please take appropriate action to correct the errors." & "</u></b></p>"

		'Here starts the logic for adding the correction specific verbiage and identifying which resources are needed.
		email_body = email_body & "<p>" & "The reason you are getting this email is to inform you that this case was not processed within the application guidelines for the following reason(s):" & "</p>"
		email_body = email_body & "<p style=" & chr(34) & "font-size:20px" & chr(34) & "><b>" & "Issues/Errors Found on Case # " & MAXIS_case_number & "</b></p>"

		If app_recvd_script_not_used_checkbox = checked Then
			email_body = email_body & "<p><i>" & "&emsp;" & counter &  ". The " & "<a href=" & chr(34) & "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/NOTES/NOTES%20-%20APPLICATION%20RECEIVED.docx" & chr(34) & ">" & "NOTES - APPLICATION RECEIVED" & "</a>" & " script was not used at the time the case was APPL'd. " & "</i>" & " This script case notes details about the application, screens for expedited SNAP (if applicable), sends the appointment letter for SNAP and/or CASH cases and transfers the case (if applicable)." & "</p>"
			add_hsr_appling_resource = TRUE
			add_script_app_recvd_resource = TRUE
			STATS_manualtime = STATS_manualtime + 30
			counter = counter + 1
		End If
		If prog_not_updated_interview_completed_checkbox = checked Then
			email_body = email_body & "<p><i>" & "&emsp;" & counter &  ". The interview date field was not completed on STAT/PROG though an interview was indicated in your case note. " & "</i></p>"
			add_hsr_interview_process_resource = TRUE
			add_script_caf_resource = TRUE
			add_script_interview_comp_resource = TRUE
			STATS_manualtime = STATS_manualtime + 30
			counter = counter + 1
		End If
		If intv_adult_cash_not_completed = checked Then
			email_body = email_body & "<p><i>" & "&emsp;" & counter & ". The interview should cover all programs requested. Questions and details about adult cash programs (GA/MSA) were not covered in your interview or CASE:NOTE." & "<br>"
			add_client_contact_interview_qs_resource = TRUE
			add_cm_05_12_12_resource = TRUE
			add_hsr_interview_process_resource = TRUE
			add_script_caf_resource = TRUE
			add_script_interview_comp_resource = TRUE
			STATS_manualtime = STATS_manualtime + 30
			counter = counter + 1
		End If
		If cash_detail_not_covered_checkbox = checked Then
			email_body = email_body & "<p><i>" & "&emsp;" & counter &  ". The client applied for cash and snap and both should have been addressed at the time of the interview." & "</i>" & " We cannot require the applicant to complete a second interview, but an in-person interview was required for cash and we need to ensure that the cash is addressed in the case note. We also want to make sure the client is informed of the need for a face to face interview." & "</p>"
			add_script_clt_contact_resource = TRUE
			add_hsr_interview_process_resource = TRUE
			add_script_caf_resource = TRUE
			add_script_interview_comp_resource = TRUE
			STATS_manualtime = STATS_manualtime + 30
			counter = counter + 1
		End If
		If snap_phone_interview_should_have_been_offered_checkbox = checked Then
			email_body = email_body & "<p><i>" & "&emsp;" & counter &  ". The client applied for cash and snap, both should have been addressed at the time of the interview." & "</i>" & " While you addressed a need for a face to face interview for cash a snap interview should have been offered at that time. Your case note indicated that the client understood, however DHS ME reviews require a case note that indicates the client declined to do the OFFERED interview for food support." & "</p>"
			add_script_clt_contact_resource = TRUE
			add_hsr_interview_process_resource = TRUE
			add_script_caf_resource = TRUE
			add_script_interview_comp_resource = TRUE
			STATS_manualtime = STATS_manualtime + 30
			counter = counter + 1
		End If
		If denied_for_no_interview_on_PND2_checkbox = checked Then
			email_body = email_body & "<p><i>" & "&emsp;" & counter &  ". This case was denied for no interview from PND2." & "</i>" & " This process is now completed by Quality Improvement staff only as part of the On-Demand waiver process. Do not deny cases at day 30 without an interview any longer." & "</p>"
			add_hsr_on_demand_resource = TRUE
			STATS_manualtime = STATS_manualtime + 30
			counter = counter + 1
		End If
		If denied_for_no_interview_NOT_on_PND2_checkbox = checked Then
			email_body = email_body & "<p><i>" & "&emsp;" & counter &  ". This case was denied for no interview; however, it was not denied from PND2." & "</i>" & " This process is now completed by Quality Improvement staff as part of the On-Demand waiver process. Do not deny cases at day 30 without an interview any longer." & "</p>"
			add_hsr_on_demand_resource = TRUE
			STATS_manualtime = STATS_manualtime + 30
			counter = counter + 1
		End If
		If caf_script_used_no_interview_checkbox = checked Then
			email_body = email_body & "<p><i>" & "&emsp;" & counter &  ". The " & "<a href=" & chr(34) & "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/NOTES/NOTES%20-%20CAF.docx" & chr(34) & ">" & "NOTES - CAF" & "</a>" & " script was used on the case when an interview has been not been completed." & "</i>" & " This script is for cases that have completed the interview process, and all STAT panels have been reviewed and updated with the information provided in the interview. If you are reviewing the case and no action has been taken in MAXIS please use "
			email_body = email_body & "<a href=" & chr(34) & "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/NOTES/NOTES%20-%20APPLICATION%20CHECK.docx" & chr(34) & ">" & "NOTES - APPLICATION CHECK" & "</a>" & ". If you cannot update MAXIS panels, and want to case note information about the interview use "
			email_body = email_body & "<a href=" & chr(34) & "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/NOTES/NOTES%20-%20INTERVIEW%20COMPLETED.docx" & chr(34) & ">" & "NOTES - INTERVIEW COMPLETED." & "</a>" & "</p>"
			add_script_app_check_resource = TRUE
			add_script_caf_resource = TRUE
			add_script_interview_comp_resource = TRUE
			STATS_manualtime = STATS_manualtime + 30
			counter = counter + 1
		End If
		If client_contacted_should_have_been_offered_interview_checkbox = checked Then
			email_body = email_body & "<p><i>" & "&emsp;" & counter &  ". Client contacted the agency and should have had an interview offered to them;" & "<b>" & " per DHS PQ the interview can be completed with just page 1 of the CAF." & "</i></b>" & " The worker will note all the client's responses (sending a copy to ECF and clearly case noting) and mail out the signature page of the CAF (along with any other needed verifications)." & "</p>"
			add_hsr_interview_process_resource = TRUE
			add_script_caf_resource = TRUE
			add_script_interview_comp_resource = TRUE
			STATS_manualtime = STATS_manualtime + 30
			counter = counter + 1
		End If
		If interview_completed_same_day_NOMI_sent_checkbox = checked Then
			email_body = email_body & "<p><i>" & "&emsp;" & counter &  ". An interview was completed the same day after a NOMI was sent, but notice was not canceled." & "</i>" & " Please cancel NOMI notices in the future if the interview is completed the day the NOMI is sent to the client." & "</p>"
			add_hsr_nomi_resource = TRUE
			STATS_manualtime = STATS_manualtime + 30
			counter = counter + 1
		End If

		If trim(email_notes) <> "" Then
			email_body = email_body & "<p>" & "Additional Information:" & "<br>"
			email_body = email_body & email_notes & "</p>"
			STATS_manualtime = STATS_manualtime + 15
		End If
		add_hsr_on_demand_resource = TRUE

		'here is where we actually add the resource informtation
		email_body = email_body & "<p><b><i>" & "&ensp;" & "RESOURCES:" & "</b></i></p>"

		If add_cm_05_12_12_resource = TRUE Then email_body = email_body & "<i>" & "&emsp;" & "Combined Manual:" & "</i><br>"

		If add_cm_05_12_12_resource = TRUE Then email_body = email_body & "&emsp;&ensp;" & "- " & "<a href=" & chr(34) & "https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=CM_00051212" & chr(34) & ">" & "CM 0005.12.12" & "</a>" & " Application Interviews - 'Offer applicants or their authorized representatives a single interview that covers all the programs for which they apply.'" & "<br>"

		If add_hsr_app_guide_resource = TRUE OR add_hsr_interview_process_resource = TRUE OR add_hsr_appling_resource = TRUE OR add_client_contact_interview_qs_resource = TRUE OR add_hsr_nomi_resource = TRUE OR add_hsr_on_demand_resource = TRUE Then email_body = email_body & "<i>" & "&emsp;" & "Internal:" & "</i><br>"

		If add_hsr_appling_resource = TRUE Then email_body = email_body & "&emsp;&ensp;" & "- HSR Manual:" & "<a href=" & chr(34) & "https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/APPLing.aspx" & chr(34) & ">" & " APPLing" & "</a><br>"
		If add_hsr_app_guide_resource = TRUE Then email_body = email_body & "&emsp;&ensp;" & "- HSR Manual:" & "<a href=" & chr(34) & "https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/Application_Guide.aspx" & chr(34) & ">" & " Application Guide" & "</a><br>"
		If add_hsr_interview_process_resource = TRUE Then email_body = email_body & "&emsp;&ensp;" & "- HSR Manual:" & "<a href=" & chr(34) & "https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/Application_Guide_Interview_Process.aspx" & chr(34) & ">" & " Application Guide: Interview Process" & "</a><br>"
		If add_client_contact_interview_qs_resource = TRUE Then email_body = email_body & "&emsp;&ensp;" & "- Client Contact Documents: " & "<a href=" & chr(34) & "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/Adults%20and%20Families%20Eligibility%20Documents/Cash%20and%20EGA%20interview%20questions.docx" & chr(34) & ">" & "Cash and EGA Interview Questions" & "</a><br>"
		If add_hsr_nomi_resource = TRUE Then email_body = email_body & "&emsp;&ensp;" & "- HSR Manual:" & "<a href=" & chr(34) & "https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/NOMI.aspx" & chr(34) & ">" & " NOMI" & "</a><br>"
		If add_hsr_on_demand_resource = TRUE Then email_body = email_body & "&emsp;&ensp;" & "- HSR Manual:" & "<a href=" & chr(34) & "https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/On_Demand_Waiver.aspx" & chr(34) & ">" & " On Demand" & "</a><br>"


		If add_script_app_recvd_resource = TRUE OR add_script_app_check_resource = TRUE OR add_script_caf_resource = TRUE OR add_script_appr_progs_resource = TRUE OR add_script_clt_contact_resource = TRUE OR add_script_interview_comp_resource = TRUE Then email_body = email_body & "<i>" & "&emsp;" & "Scripts: (Instructions)" & "</i><br>"

		If add_script_app_recvd_resource = TRUE Then email_body = email_body & "&emsp;&ensp;" & "- " & "<a href=" & chr(34) & "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/NOTES/NOTES%20-%20APPLICATION%20RECEIVED.docx" & chr(34) & ">" & "NOTES - APPLICATION RECEIVED" & "</a><br>"
		If add_script_app_check_resource = TRUE Then email_body = email_body & "&emsp;&ensp;" & "- " & "<a href=" & chr(34) & "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/NOTES/NOTES%20-%20APPLICATION%20CHECK.docx" & chr(34) & ">" & "NOTES - APPLICATION CHECK" & "</a><br>"
		If add_script_caf_resource = TRUE Then email_body = email_body & "&emsp;&ensp;" & "- " & "<a href=" & chr(34) & "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/NOTES/NOTES%20-%20CAF.docx" & chr(34) & ">" & "NOTES - CAF" & "</a><br>"
		If add_script_appr_progs_resource = TRUE Then email_body = email_body & "&emsp;&ensp;" & "- " & "<a href=" & chr(34) & "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/NOTES/NOTES%20-%20APPROVED%20PROGRAMS.docx" & chr(34) & ">" & "NOTES - APPROVED PROGRAMS" & "</a><br>"
		If add_script_clt_contact_resource = TRUE Then email_body = email_body & "&emsp;&ensp;" & "- " & "<a href=" & chr(34) & "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/NOTES/NOTES%20-%20CLIENT%20CONTACT.docx" & chr(34) & ">" & "NOTES - CLIENT CONTACT" & "</a><br>"
		If add_script_interview_comp_resource = TRUE Then email_body = email_body & "&emsp;&ensp;" & "- " & "<a href=" & chr(34) & "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/NOTES/NOTES%20-%20INTERVIEW%20COMPLETED.docx" & chr(34) & ">" & "NOTES - INTERVIEW COMPLETED" & "</a><br>"

		'End of the message with email llinks
		email_body = email_body & "<p>" & "Thank you for taking the time to review this information. If you have any additional questions or want additional directions and resources, please contact the QI team. For Script questions contact the BlueZone Script Team." & "<br>"
		email_body = email_body & "<a href=" & chr(34) & "mailto:HSPH.EWS.QUALITYIMPROVEMENT@hennepin.us?subject=SNAP%20EXP%20Questions" & chr(34) & ">" & "Email Quality Improvement" & "</a><br>"
		email_body = email_body & "<a href=" & chr(34) & "mailto:HSPH.EWS.BlueZoneScripts@hennepin.us?subject=SNAP%20EXP%20Questions" & chr(34) & ">" & "Email the BlueZone Script Team" & "</a><br>"
		email_body = email_body & "</p>"

		email_body = email_body & "<p style=" & chr(34) & "font-size:20px" & chr(34) & ">" & "From, " & email_signature & "<br>"
		email_body = email_body & "Member of the ES Quality Improvement Team" & "</p>"

		'NOW WE SEND THE EMAIL
		'Setting up the Outlook application
		Set objOutlook = CreateObject("Outlook.Application")
		Set objMail = objOutlook.CreateItem(0)
		objMail.Display                                 'To display message

		'Adds the information to the email
		objMail.to = email_recipient                        'email recipient
		objMail.cc = email_recipient_cc                     'cc recipient
		objMail.SentOnBehalfOfName = "HSPH.EWS.QUALITYIMPROVEMENT@hennepin.us"
		' objMail.SentOnBehalfOfName = "HSPH.EWS.BlueZoneScripts@hennepin.us"
		objMail.Subject = email_subject                 'email subject
		objMail.HTMLBody = email_body                       'email body
		If email_attachment <> "" then objMail.Attachments.Add(email_attachment)       'email attachement (can only support one for now)
		'Sends email
		If send_email = true then objMail.Send	                   'Sends the email
		Set objMail =   Nothing
		Set objOutlook = Nothing

		'On Demand Column constants'
		recip_col = 1
		date_col = 2
		case_numb_col = 3
		sup_email_col = 4
		app_recvd_script_not_used_col 								= 5
		prog_not_updated_interview_completed_col 					= 6
		adult_cash_interview_not_completed_col						= 7
		cash_detail_not_covered_col 								= 8
		snap_phone_interview_should_have_been_offered_col 			= 9
		denied_for_no_interview_on_PND2_col 						= 10
		denied_for_no_interview_NOT_on_PND2_col						= 11
		caf_script_used_no_interview_col 							= 12
		client_contacted_should_have_been_offered_interview_col 	= 13
		interview_completed_same_day_NOMI_sent_col 					= 14
		other_notes_col 											= 15
		worker_to_follow_up_col										= 16
		qi_worker_col 												= 17

		'HERE IS THE TRAKING PART - so we can save the information about what and who we emailed for corrections
		excel_file_path = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\On Demand Waiver\Applications Statistics\ON DEMAND Email Corrections Data.xlsx"
		call excel_open(excel_file_path, False,  False, ObjExcel, objWorkbook)  'opening the Excel'
		ObjExcel.worksheets("Data").Activate

		excel_row = 2                                                           'finding the first empty excel row
		Do
			this_entry = ObjExcel.Cells(excel_row, 1).Value
			this_entry = trim(this_entry)
			If this_entry <> "" Then excel_row = excel_row + 1
		Loop until this_entry = ""

		'Adding all the information to the Excel
		ObjExcel.Cells(excel_row, recip_col).Value 							= email_recipient
		ObjExcel.Cells(excel_row, date_col).Value 							= date
		ObjExcel.Cells(excel_row, case_numb_col).Value 						= MAXIS_case_number
		ObjExcel.Cells(excel_row, sup_email_col).Value 						= email_recipient_cc
		If app_recvd_script_not_used_checkbox = checked						Then ObjExcel.Cells(excel_row, app_recvd_script_not_used_col).Value 								= "X"
		If prog_not_updated_interview_completed_checkbox = checked			Then ObjExcel.Cells(excel_row, prog_not_updated_interview_completed_col).Value 						= "X"
		If intv_adult_cash_not_completed = checked							Then ObjExcel.Cells(excel_row, adult_cash_interview_not_completed_col).Value 						= "X"
		If cash_detail_not_covered_checkbox = checked						Then ObjExcel.Cells(excel_row, cash_detail_not_covered_col).Value 									= "X"
		If snap_phone_interview_should_have_been_offered_checkbox = checked Then ObjExcel.Cells(excel_row, snap_phone_interview_should_have_been_offered_col).Value 			= "X"
		If denied_for_no_interview_on_PND2_checkbox = checked				Then ObjExcel.Cells(excel_row, denied_for_no_interview_on_PND2_col).Value 							= "X"
		If denied_for_no_interview_NOT_on_PND2_checkbox = checked			Then ObjExcel.Cells(excel_row, denied_for_no_interview_NOT_on_PND2_col).Value 						= "X"
		If caf_script_used_no_interview_checkbox = checked					Then ObjExcel.Cells(excel_row, caf_script_used_no_interview_col).Value 								= "X"
		If client_contacted_should_have_been_offered_interview_checkbox = checked Then ObjExcel.Cells(excel_row, client_contacted_should_have_been_offered_interview_col).Value = "X"
		If interview_completed_same_day_NOMI_sent_checkbox = checked		Then ObjExcel.Cells(excel_row, interview_completed_same_day_NOMI_sent_col).Value 					= "X"
		ObjExcel.Cells(excel_row, other_notes_col).Value 					= email_notes
		If action_needed = "Case has been updated already - no action needed." Then ObjExcel.Cells(excel_row, worker_to_follow_up_col).Value 		= "NO"
		If action_needed = "Please update the case to resolve these issues." Then ObjExcel.Cells(excel_row, worker_to_follow_up_col).Value 			= "YES"
		ObjExcel.Cells(excel_row, qi_worker_col).Value 						= qi_worker_full_name

		ObjExcel.ActiveWorkbook.Save                                            'saving and closing the Excel spreadsheet
		ObjExcel.ActiveWorkbook.Close
		ObjExcel.Application.Quit

End Select

Call script_end_procedure_with_error_report(end_msg)
