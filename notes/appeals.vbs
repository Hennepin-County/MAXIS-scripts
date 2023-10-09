'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - APPEALS.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 500                     'manual run time in seconds
STATS_denomination = "C"                   'C is for each CASE
'END OF stats block================================================================================

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
' call run_from_GitHub(script_repository & "application-received.vbs")

'END FUNCTIONS LIBRARYBLOCK================================================================================================

'CHANGELOG BLOCK===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County
call changelog_update("01/26/2023", "Removed term 'ECF' from the case note per DHS guidance, and referencing the case file instead.", "Ilse Ferris, Hennepin County")
CALL changelog_update("08/10/2022", "Removed anticipated date results as a mandatory field.", "Ilse Ferris, Hennepin County")
CALL changelog_update("04/22/2022", "Update to resolve bug with docket number writing in case notes.", "MiKayla Handley, Hennepin County")
CALL changelog_update("05/14/2021", "Update to resolve bug with N/A and case notes.", "MiKayla Handley, Hennepin County")
CALL changelog_update("02/22/2021", "Update to dialog for received handling. Added mandatory explanation for continuation of pre-appeal benefits.", "MiKayla Handley, Hennepin County")
CALL changelog_update("01/20/2021", "Initial version.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'---------------------------------------------------------------------------------------The script
'Grabs the case number
EMConnect ""
CALL MAXIS_case_number_finder(MAXIS_case_number)

'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 176, 105, "Appeals"
  EditBox 75, 5, 45, 15, MAXIS_case_number
  EditBox 75, 25, 45, 15, docket_number
  DropListBox 75, 45, 95, 15, "Select One:"+chr(9)+"Received"+chr(9)+"Summary Completed"+chr(9)+"Hearing Information"+chr(9)+"Decision Received"+chr(9)+"Pending Request"+chr(9)+"Reconsideration"+chr(9)+"Resolution", appeal_actions
  EditBox 75, 65, 95, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 75, 85, 45, 15
    CancelButton 125, 85, 45, 15
  Text 5, 10, 50, 10, "Case Number:"
  Text 5, 30, 55, 10, "Docket Number:"
  Text 5, 50, 50, 10, "Appeal Action:"
  Text 5, 70, 60, 10, "Worker Signature:"
EndDialog

'Runs the first dialog - which confirms the case number
Do
	Do
		err_msg = ""
		Dialog Dialog1
		cancel_without_confirmation
      	Call validate_MAXIS_case_number(err_msg, "*")
        IF docket_number = "" then err_msg = err_msg & vbNewLine & "* Please enter a docket number or enter N/A if unknown."
        IF appeal_actions = "Select One:" then err_msg = err_msg & vbNewLine & "* Please select what type of appeal action the client is claiming."
        IF worker_signature = "" THEN err_msg = err_msg & vbCr & "* Please sign your case note."
        IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
	Loop until err_msg = ""
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

IF appeal_actions = "Received" THEN
    '---------------------------------------------------------------------------------------------'pending & active programs information
    CALL navigate_to_MAXIS_screen_review_PRIV("CASE", "CURR", is_this_priv) ' need discovery on priv cases for xfer handling'
    IF is_this_priv = TRUE THEN script_end_procedure("This case is privileged, the script will now end.")
    CALL determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type, case_status, list_active_programs, list_pending_programs)

    '-------------------------------------------------------------------------------------------------DIALOG
    Dialog1 = "" 'Blanking out previous dialog detail
    BeginDialog Dialog1, 0, 0, 286, 210, "Received - App pend: "  & list_pending_programs & " Active on: "  & list_active_programs
      DropListBox 60, 5, 85, 15, "Select One:"+chr(9)+"Case File"+chr(9)+"DHS"+chr(9)+"Phone-Verbal Request", how_appeal_rcvd_dropdown
      EditBox 235, 5, 45, 15, client_request_date
      EditBox 140, 25, 45, 15, effective_date
      DropListBox 90, 45, 55, 15, "Select One:"+chr(9)+"Denial"+chr(9)+"Overpayment"+chr(9)+"Reduction"+chr(9)+"Termination"+chr(9)+"Other", appeal_action_dropdown
      DropListBox 135, 65, 55, 15, "Select:"+chr(9)+"YES"+chr(9)+"NO", benefits_continuing_dropdown
      EditBox 235, 45, 45, 15, claim_number
      EditBox 165, 85, 115, 15, benefits_continuing_explanation
      EditBox 70, 105, 210, 15, proofs_attachments
      CheckBox 10, 135, 35, 10, "CASH", cash_appeal_checkbox
      CheckBox 45, 135, 30, 10, "SNAP", snap_appeal_checkbox
      CheckBox 80, 135, 30, 10, "GRH", grh_appeal_checkbox
      CheckBox 115, 135, 25, 10, "HC", hc_appeal_checkbox
      CheckBox 140, 135, 35, 10, "EMER", emer_appeal_checkbox
      CheckBox 175, 135, 40, 10, "OTHER", ot_appeal_checkbox
      CheckBox 215, 135, 25, 10, "CCA", cca_appeal_checkbox
      CheckBox 245, 135, 25, 10, "IVE", ive_appeal_checkbox
      CheckBox 10, 150, 65, 10, "BURIAL ASSIST", BURIAL_ASSIST_checkbox
      CheckBox 80, 150, 75, 10, "REVENUE RECAP", REVENUE_RECAP_checkbox
      CheckBox 160, 150, 50, 10, "SANCTION", SANCTION_checkbox
      CheckBox 215, 150, 55, 10, "TRANSPORT", TRANSPORT_checkbox
      EditBox 50, 170, 230, 15, other_notes
      ButtonGroup ButtonPressed
        OkButton 185, 190, 45, 15
        CancelButton 235, 190, 45, 15
      Text 5, 175, 45, 10, "Other Notes:"
      Text 5, 30, 130, 10, "Effective date of action being appealed:"
      Text 5, 110, 65, 10, "Proof/Attachments:"
      Text 175, 10, 55, 10, "Date Received:"
      GroupBox 5, 125, 275, 40, "Appealed Programs/Decisions"
      Text 5, 70, 130, 10, "Benefits continuing at pre-appeal level:"
      Text 5, 10, 50, 10, "How Received:"
      Text 5, 50, 85, 10, "Action client is appealing:"
      Text 5, 90, 155, 10, "How was determination made for cont benefits:"
      Text 200, 30, 105, 10, "Docket #: " & docket_number
      Text 200, 50, 35, 10, "Claim(s)#:"
    EndDialog
    '------------------------------------------------------------------------------------DIALOG
    Do
    	Do
            err_msg = ""
            Dialog Dialog1
            cancel_confirmation
            IF isdate(client_request_date) = FALSE THEN err_msg = err_msg & vbNewLine & "* Please enter the date the appeal form was received."
            IF isdate(effective_date) = "" THEN err_msg = err_msg & vbNewLine & "* Please enter the effective date of the appeal."
            IF appeal_actions = "Select One:" then err_msg = err_msg & vbNewLine & "* Please select what type of appeal was received."
    	    IF how_appeal_rcvd_dropdown = "Select One:" THEN  err_msg = err_msg & vbNewLine & "* Please select how the appeal was received."
            IF appeal_action_dropdown = "Other" and other_notes = "" THEN  err_msg = err_msg & vbNewLine & "* Please advise what the appeal action was in other notes."
    	    IF docket_number = "" THEN err_msg = err_msg & vbNewLine & "* Please enter the docket number, if unknown enter N/A."
            IF benefits_continuing_dropdown = "Select:" then err_msg = err_msg & vbNewLine & "* Please select if the benefits will be continuing and explain your decision."
            IF benefits_continuing_explanation = "" then err_msg = err_msg & vbNewLine & "* Please advise why the benefits will or will not be continuing at pre-appeal level."
            IF proofs_attachments = "" then err_msg = err_msg & vbNewLine & "* Please advise what proofs or information has been provided."
            If ot_appeal_checkbox = CHECKED and other_notes = "" THEN err_msg = err_msg & vbNewLine & "* Please advise what other program  or decision is being appealed."
            IF cash_appeal_checkbox = UNCHECKED AND snap_appeal_checkbox = UNCHECKED AND grh_appeal_checkbox = UNCHECKED AND hc_appeal_checkbox = UNCHECKED AND emer_appeal_checkbox = UNCHECKED AND ot_appeal_checkbox = UNCHECKED AND ca_appeal_checkbox  = UNCHECKED AND ive_appeal_checkbox = UNCHECKED AND BURIAL_ASSIST_checkbox = UNCHECKED AND REVENUE_RECAP_checkbox = UNCHECKED AND SANCTION_checkbox = UNCHECKED AND TRANSPORT_checkbox = UNCHECKED THEN err_msg = err_msg & vbCr & "* Please select the appealed program or decision."
    	    IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
    	LOOP UNTIL err_msg = ""
        CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    Loop until are_we_passworded_out = false					'loops until user passwords back in

    appeal_programs = ""
    IF cash_appeal_checkbox = CHECKED THEN appeal_programs =  appeal_programs & "CASH, "
    IF snap_appeal_checkbox = CHECKED THEN appeal_programs = appeal_programs & "SNAP, "
    IF grh_appeal_checkbox = CHECKED THEN appeal_programs = appeal_programs & "GRH, "
    IF hc_appeal_checkbox = CHECKED THEN appeal_programs = appeal_programs & "HC, "
    IF emer_appeal_checkbox = CHECKED THEN appeal_programs = appeal_programs & "EMERGENCY, "
    IF cca_appeal_checkbox = CHECKED THEN appeal_programs = appeal_programs & "CCA, "
    IF ive_appeal_checkbox = CHECKED THEN appeal_programs = appeal_programs & "IV-E, "
    IF BURIAL_ASSIST_checkbox = CHECKED THEN appeal_programs =  appeal_programs & "BURIAL ASSISTANCE, "
    IF REVENUE_RECAP_checkbox = CHECKED THEN appeal_programs =  appeal_programs & "REVENUE RECAPTURE, "
    IF SANCTION_checkbox = CHECKED THEN appeal_programs =  appeal_programs & "SANCTION, "
    IF TRANSPORT_checkbox = CHECKED THEN appeal_programs =  appeal_programs & "TRANSPORT, "
    IF ot_appeal_checkbox = CHECKED THEN appeal_programs = appeal_programs & "Other:, "
    appeal_programs = trim(appeal_programs)  'trims excess spaces of appeal_programs
    If right(appeal_programs, 1) = "," THEN appeal_programs = left(appeal_programs, len(appeal_programs) - 1)

    start_a_blank_case_note      'navigates to case/note and puts case/note into edit mode
    CALL write_variable_in_CASE_NOTE("-----Appeal " & appeal_actions & "-----")
    CALL write_bullet_and_variable_in_CASE_NOTE("Docket Number", docket_number)
    CALL write_bullet_and_variable_in_CASE_NOTE("Date appeal request received", client_request_date)
    CALL write_bullet_and_variable_in_CASE_NOTE("Claim(s) Number", claim_number)
    CALL write_bullet_and_variable_in_CASE_NOTE("How appeal request received", how_appeal_rcvd_dropdown)
    CALL write_bullet_and_variable_in_CASE_NOTE("Effective date of action being appealed", effective_date)
    CALL write_bullet_and_variable_in_CASE_NOTE("Action client is appealing", appeal_action_dropdown)
    CALL write_bullet_and_variable_in_CASE_NOTE ("Program(s) client appealing", appeal_programs)
    CALL write_bullet_and_variable_in_CASE_NOTE("Benefits continuing at pre-appeal level", benefits_continuing_dropdown)
    CALL write_bullet_and_variable_in_CASE_NOTE("Explanation", benefits_continuing_explanation)
    CALL write_bullet_and_variable_in_CASE_NOTE("Proofs/attachments", proofs_attachments)
    CALL write_bullet_and_variable_in_CASE_NOTE("Other Notes", other_notes)
    CALL write_bullet_and_variable_in_CASE_NOTE("Active Programs", list_active_programs)
    CALL write_bullet_and_variable_in_CASE_NOTE("Pending Programs", list_pending_programs)
    CALL write_variable_in_CASE_NOTE("---")
    CALL write_variable_in_CASE_NOTE(worker_signature)
END IF 'appeal_actions = "Received"'
IF appeal_actions = "Pending Request"  THEN
    '-------------------------------------------------------------------------------------------------DIALOG
    Dialog1 = "" 'Blanking out previous dialog detail
    BeginDialog Dialog1, 0, 0, 231, 85, "Pending Request"
      EditBox 55, 5, 50, 15, date_requested
      EditBox 170, 5, 55, 15, docket_number
      EditBox 95, 25, 130, 15, verification_needed
      EditBox 60, 45, 165, 15, other_notes
      ButtonGroup ButtonPressed
        OkButton 130, 65, 45, 15
        CancelButton 180, 65, 45, 15
      Text 5, 30, 85, 10, "Verification(s) Requested:"
      Text 5, 50, 45, 10, "Other Notes:"
      Text 5, 10, 50, 10, "Request Date:"
      Text 130, 10, 35, 10, "Docket #:"
    EndDialog

    DO
        Do
            err_msg = ""
            Dialog Dialog1
            cancel_confirmation
            IF docket_number = "" THEN err_msg = err_msg & vbNewLine & "* Please enter a docket number or enter N/A if unknown."
            IF isdate(date_requested) = false THEN err_msg = err_msg & vbNewLine & "* Please complete date of hearing."
            IF verification_needed = "" THEN err_msg = err_msg & vbNewLine & "* Please enter the pending verifications"
            IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
        Loop until err_msg = ""
        CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    Loop until are_we_passworded_out = false					'loops until user passwords back in

    start_a_blank_case_note      'navigates to case/note and puts case/note into edit mode
    Call write_variable_in_CASE_NOTE("-----Appeal Pending Request-----")
    CALL write_bullet_and_variable_in_CASE_NOTE("Docket Number", docket_number)
    Call write_bullet_and_variable_in_CASE_NOTE("Request date", date_requested)
    Call write_bullet_and_variable_in_CASE_NOTE("Requested Verification(s)", verification_needed)
    Call write_bullet_and_variable_in_CASE_NOTE("Other Notes", other_notes)
    Call write_variable_in_CASE_NOTE ("---")
    Call write_variable_in_CASE_NOTE(worker_signature)
END IF

IF appeal_actions = "Summary Completed"  THEN
    '-------------------------------------------------------------------------------------------------DIALOG
    Dialog1 = "" 'Blanking out previous dialog detail
    BeginDialog Dialog1, 0, 0, 276, 85, "Appeal Summary Completed"
      EditBox 65, 5, 50, 15, docket_number
      EditBox 210, 5, 60, 15, date_appeal_rcvd
      EditBox 65, 25, 50, 15, claim_number
      EditBox 210, 25, 60, 15, effective_date
      EditBox 95, 45, 175, 15, action_client_is_appealing
      ButtonGroup ButtonPressed
        OkButton 175, 65, 45, 15
        CancelButton 225, 65, 45, 15
      Text 135, 10, 75, 10, "Date appeal received:"
      Text 130, 30, 80, 10, "Effective date of action:"
      Text 5, 30, 60, 10, "Claim number(s):"
      Text 5, 50, 85, 10, "Action client is appealing:"
      Text 5, 10, 55, 10, "Docket number:"
    EndDialog

    Do
        Do
            err_msg = ""
            Dialog Dialog1
            cancel_confirmation
            IF docket_number = "" THEN err_msg = err_msg & vbNewLine & "* Please enter a valid docket number or enter N/A if unknown."
            IF Isdate(date_appeal_rcvd) = false THEN err_msg = err_msg & vbNewLine & "* Please enter a date for the appeal."
            IF Isdate(effective_date) = false THEN err_msg = err_msg & vbNewLine & "* Please enter the effective date."
            IF action_client_is_appealing = "" THEN err_msg = err_msg & vbNewLine & "* Please enter action that client is appealing."
            IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
        Loop until err_msg = ""
        CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    Loop until are_we_passworded_out = false					'loops until user passwords back in

    start_a_blank_case_note      'navigates to case/note and puts case/note into edit mode
    Call write_variable_in_CASE_NOTE("-----Appeal Summary Completed-----")
    Call write_bullet_and_variable_in_CASE_NOTE("Docket number", docket_number)
    CALL write_bullet_and_variable_in_CASE_NOTE("Claim(s) number", claim_number)
    Call write_bullet_and_variable_in_CASE_NOTE("Date appeal request received", date_appeal_rcvd)
    Call write_bullet_and_variable_in_CASE_NOTE("Effective date of action being appealed", effective_date)
    Call write_bullet_and_variable_in_CASE_NOTE("Action client is appealing", action_client_is_appealing)
    Call write_variable_in_CASE_NOTE ("---")
    Call write_variable_in_CASE_NOTE(worker_signature)
END IF

IF appeal_actions = "Reconsideration" THEN
    '-------------------------------------------------------------------------------------------------DIALOG
    Dialog1 = "" 'Blanking out previous dialog detail
    BeginDialog Dialog1, 0, 0, 286, 105, "Reconsideration"
      EditBox 65, 5, 55, 15, hearing_date
      EditBox 225, 5, 55, 15, docket_number
      DropListBox 65, 25, 60, 15, "Select One:"+chr(9)+"Yes, in person"+chr(9)+"Yes, by phone"+chr(9)+"Did not attend", appeal_attendence
      EditBox 225, 25, 55, 15, anticipated_date_result
      EditBox 65, 45, 215, 15, hearing_details
      EditBox 65, 65, 215, 15, other_notes
      ButtonGroup ButtonPressed
        OkButton 175, 85, 50, 15
        CancelButton 230, 85, 50, 15
      Text 5, 30, 55, 10, "Client Attended:"
      Text 5, 50, 55, 10, "Hearing Details:"
      Text 5, 70, 45, 10, "Other Notes:"
      Text 5, 10, 60, 10, "Date Of Hearing:"
      Text 185, 10, 35, 10, "Docket #:"
      Text 135, 30, 85, 10, "Anticipated Decision Date:"
    EndDialog

    DO
        Do
            err_msg = ""
            Dialog Dialog1
            cancel_confirmation
            IF isdate(hearing_date) = false THEN err_msg = err_msg & vbNewLine & "* Please complete date of hearing."
            If appeal_attendence = "Select One:" THEN err_msg = err_msg & vbNewLine & "* Please select if the client attended appeal, or if appeal was held by phone"
            IF hearing_details = "" THEN err_msg = err_msg & vbNewLine & "* Please enter hearing details"
            IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
        Loop until err_msg = ""
        CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    Loop until are_we_passworded_out = false					'loops until user passwords back in

    start_a_blank_case_note      'navigates to case/note and puts case/note into edit mode
    Call write_variable_in_CASE_NOTE("-----Reconsideration-----")
    CALL write_bullet_and_variable_in_CASE_NOTE("Docket Number", docket_number)
    Call write_bullet_and_variable_in_CASE_NOTE("Date Of Hearing", hearing_date)
    Call write_bullet_and_variable_in_CASE_NOTE("Did client attend the appeal", appeal_attendence)
    Call write_bullet_and_variable_in_CASE_NOTE("Hearing details", hearing_details)
    Call write_bullet_and_variable_in_CASE_NOTE("Anticipated date of decision", anticipated_date_result)
    Call write_bullet_and_variable_in_CASE_NOTE("Other Notes", other_notes)
    Call write_variable_in_CASE_NOTE ("---")
    Call write_variable_in_CASE_NOTE(worker_signature)
END IF

IF appeal_actions = "Hearing Information" THEN
    '-------------------------------------------------------------------------------------------------DIALOG
    Dialog1 = "" 'Blanking out previous dialog detail
    BeginDialog Dialog1, 0, 0, 286, 105, "Hearing Information"
      EditBox 65, 5, 55, 15, hearing_date
      EditBox 225, 5, 55, 15, docket_number
      DropListBox 65, 25, 60, 15, "Select One:"+chr(9)+"Yes, in person"+chr(9)+"Yes, by phone"+chr(9)+"Did not attend", appeal_attendence
      EditBox 225, 25, 55, 15, anticipated_date_result
      EditBox 65, 45, 215, 15, hearing_details
      EditBox 65, 65, 215, 15, other_notes
      ButtonGroup ButtonPressed
        OkButton 175, 85, 50, 15
        CancelButton 230, 85, 50, 15
      Text 5, 10, 60, 10, "Date Of Hearing:"
      Text 185, 10, 35, 10, "Docket #:"
      Text 5, 30, 55, 10, "Client Attended:"
      Text 135, 30, 85, 10, "Anticipated Decision Date:"
      Text 5, 50, 55, 10, "Hearing Details:"
      Text 5, 70, 45, 10, "Other Notes:"
    EndDialog

    DO
        Do
            err_msg = ""
            Dialog Dialog1
            cancel_confirmation
            IF isdate(hearing_date) = false THEN err_msg = err_msg & vbNewLine & "* Please complete date of hearing."
            If appeal_attendence = "Select One:" THEN err_msg = err_msg & vbNewLine & "* Please select if the client attended appeal, or if appeal was held by phone"
            IF hearing_details = "" THEN err_msg = err_msg & vbNewLine & "* Please enter hearing details"
            IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
        Loop until err_msg = ""
        CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    Loop until are_we_passworded_out = false					'loops until user passwords back in

    start_a_blank_case_note      'navigates to case/note and puts case/note into edit mode
    Call write_variable_in_CASE_NOTE("-----Appeal Hearing Info-----")
    CALL write_bullet_and_variable_in_CASE_NOTE("Docket Number", docket_number)
    Call write_bullet_and_variable_in_CASE_NOTE("Date Of Hearing", hearing_date)
    Call write_bullet_and_variable_in_CASE_NOTE("Did Client Attend The Appeal", appeal_attendence)
    Call write_bullet_and_variable_in_CASE_NOTE("Hearing Details", hearing_details)
    Call write_bullet_and_variable_in_CASE_NOTE("Anticipated date of decision", anticipated_date_result)
    Call write_bullet_and_variable_in_CASE_NOTE("Other Notes", other_notes)
    Call write_variable_in_CASE_NOTE ("---")
    Call write_variable_in_CASE_NOTE(worker_signature)
END IF

IF appeal_actions = "Decision Received" THEN
'-------------------------------------------------------------------------------------------------DIALOG
    Dialog1 = "" 'Blanking out previous dialog detail
    BeginDialog Dialog1, 0, 0, 336, 85, "Appeal Decision Received"
      EditBox 85, 5, 60, 15, docket_number
      EditBox 270, 5, 60, 15, date_signed_by_judge
      EditBox 85, 25, 245, 15, disposition_of_appeal
      EditBox 85, 45, 245, 15, actions_needed
      DropListBox 125, 65, 55, 15, "Select One:"+chr(9)+"YES"+chr(9)+"NO"+chr(9)+"N/A", compliance_form_needed
      ButtonGroup ButtonPressed
        OkButton 225, 65, 50, 15
        CancelButton 280, 65, 50, 15
      Text 5, 10, 40, 10, "Docket #:"
      Text 5, 30, 75, 10, "Disposition of appeal:"
      Text 5, 50, 55, 10, "Actions needed:"
      Text 185, 10, 75, 10, "Date signed by judge:"
      Text 5, 70, 115, 10, "SNAP compliance form completed:"
    EndDialog

    'Shows dialog and creates and displays an error message if worker completes things incorrectly.
    Do
        Do
            err_msg = ""
            Dialog Dialog1
            cancel_confirmation
            IF disposition_of_appeal = "" THEN err_msg = err_msg & vbNewLine & "* Please enter the disposition of the appeal"
            IF actions_needed = "" THEN err_msg = err_msg & vbNewLine & "* Please enter actions needed"
            If compliance_form_needed = "Select One:" THEN err_msg = err_msg & vbNewLine & "* Please select whether a compliance form is needed"
            IF isdate(date_signed_by_judge) = false THEN err_msg = err_msg & vbNewLine & "* Please enter a valid date the appeal findings were signed by the Judge"
            IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
        Loop until err_msg = ""
        CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    Loop until are_we_passworded_out = false					'loops until user passwords back in

     start_a_blank_case_note      'navigates to case/note and puts case/note into edit mode
     Call write_variable_in_CASE_NOTE("-----Appeal Decision Received-----")
     CALL write_bullet_and_variable_in_CASE_NOTE("Docket Number", docket_number)
     Call write_bullet_and_variable_in_CASE_NOTE("Disposition of Appeal", disposition_of_appeal)
     Call write_bullet_and_variable_in_CASE_NOTE("Actions Required", actions_needed)
     Call write_bullet_and_variable_in_CASE_NOTE("SNAP compliance form completed", compliance_form_needed)
     Call write_bullet_and_variable_in_CASE_NOTE("Date signed by judge", date_signed_by_judge)
     Call write_variable_in_CASE_NOTE ("---")
     Call write_variable_in_CASE_NOTE(worker_signature)
END IF

IF appeal_actions = "Resolution" THEN
    '-------------------------------------------------------------------------------------------------DIALOG
    Dialog1 = "" 'Blanking out previous dialog detail
    BeginDialog Dialog1, 0, 0, 231, 195, "Appeal Resolution"
      DropListBox 180, 15, 40, 15, "Select:"+chr(9)+"YES"+chr(9)+"NO", Further_Action_Required_dropdown
      DropListBox 180, 30, 40, 15, "Select:"+chr(9)+"YES"+chr(9)+"NO", Overpayments_Required_dropdown
      EditBox 95, 45, 40, 15, claim_number
      EditBox 180, 45, 40, 15, overpayment_amount
      DropListBox 180, 65, 40, 15, "Select:"+chr(9)+"YES"+chr(9)+"NO", Confirmed_Resolution_dropdown
      DropListBox 180, 80, 40, 15, "Select:"+chr(9)+"YES"+chr(9)+"NO", Withdrawn_with_appellant_dropdown
      DropListBox 180, 95, 40, 15, "Select:"+chr(9)+"YES"+chr(9)+"NO", Notified_EWS_Withdraw_dropdown
      DropListBox 180, 110, 40, 15, "Select:"+chr(9)+"YES"+chr(9)+"NO", Referred_Appellant_dropdown
      EditBox 95, 135, 130, 15, actions_taken_required
      EditBox 50, 155, 175, 15, other_notes
      ButtonGroup ButtonPressed
        OkButton 130, 175, 45, 15
        CancelButton 180, 175, 45, 15
      Text 10, 85, 85, 10, "Withdrawn with Appellant:"
      Text 10, 100, 125, 10, "Notified Hennepin EWS of Withdrawal:"
      Text 10, 115, 170, 10, "Referred Appellant to DHS Appeals 651-431-3600:"
      Text 10, 20, 80, 10, "Further Action Required:"
      Text 5, 160, 45, 10, "Other Notes:"
      Text 150, 50, 30, 10, "Amount:"
      Text 10, 35, 85, 10, "Overpayments Required:"
      Text 5, 140, 85, 10, "Actions Taken/Required:"
      GroupBox 5, 5, 220, 125, "Select to Confirm:"
      Text 10, 70, 80, 10, "Confirmed Resolution:"
      Text 35, 50, 60, 10, "Claim(s) Number:"
    EndDialog
    Do
        DO
            err_msg = ""
            Dialog Dialog1
            cancel_confirmation
            IF Further_Action_Required_dropdown = "Select:" THEN err_msg = err_msg & vbNewLine & "* Please select YES or NO what action is needed by caseworker."
            IF Overpayments_Required_dropdown = "Select:" THEN err_msg = err_msg & vbNewLine & "* Please select YES or NO if there is an overpayment required."
            IF Confirmed_Resolution_dropdown = "Select:" THEN err_msg = err_msg & vbNewLine & "* Please select YES or NO if the resolution has been confirmed."
            IF Withdrawn_with_appellant_dropdown = "Select:" THEN err_msg = err_msg & vbNewLine & "* Please select YES or NO if the withdrawl was done with the appellant."
            IF Notified_EWS_Withdraw_dropdown = "Select:" THEN err_msg = err_msg & vbNewLine & "* Please select YES or NO if ES has been notified."
            IF Referred_Appellant_dropdown = "Select:" THEN err_msg = err_msg & vbNewLine & "* Please select YES or NO if appellant has been referred."
            IF Further_Action_Required_dropdown = "YES" and actions_taken_required = "" THEN err_msg = err_msg & vbNewLine & "* Please enter what action is needed by caseworker."
            IF Overpayments_Required_dropdown = "YES" and overpayment_amount = "" THEN err_msg = err_msg & vbNewLine & "* Please enter the amount of the overpayment, if unknown enter N/A."
            IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
        Loop until err_msg = ""
        CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    Loop until are_we_passworded_out = false					'loops until user passwords back in

    start_a_blank_case_note      'navigates to case/note and puts case/note into edit mode
    Call write_variable_in_CASE_NOTE("-----Appeal Resolution-----")
    IF Overpayments_Required_dropdown = "YES" THEN
        Call write_variable_in_CASE_NOTE("* Overpayments Required")
        CALL write_bullet_and_variable_in_CASE_NOTE("Claim(s) Number", claim_number)
        Call write_bullet_and_variable_in_CASE_NOTE("Overpayment Amount", overpayment_amount)
    ELSE
        Call write_variable_in_CASE_NOTE("* No Overpayments Established")
    END IF
    Call write_bullet_and_variable_in_CASE_NOTE("Confirmed Resolution", Confirmed_Resolution_dropdown)
    Call write_bullet_and_variable_in_CASE_NOTE("Withdrawn with Appellant", Withdrawn_with_appellant_dropdown)
    Call write_bullet_and_variable_in_CASE_NOTE("Notified Hennepin EWS of Withdrawal", Notified_EWS_Withdraw_dropdown )
    Call write_bullet_and_variable_in_CASE_NOTE("Referred Appellant to DHS Appeals 651-431-3600", Referred_Appellant_dropdown)
    Call write_bullet_and_variable_in_CASE_NOTE("Action is needed by caseworker.", Further_Action_Required_dropdown)
    Call write_bullet_and_variable_in_CASE_NOTE("Actions Taken/Required", actions_taken_required)
    Call write_bullet_and_variable_in_CASE_NOTE("Other Notes", other_notes)
    Call write_variable_in_CASE_NOTE ("---")
    call write_variable_in_CASE_NOTE(worker_signature)
END IF
script_end_procedure_with_error_report("Success! CASE/NOTE has been updated please review to ensure information was noted correctly.")
