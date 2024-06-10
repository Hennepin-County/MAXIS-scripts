'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - EBT BENEFITS STOLEN.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 150                	'manual run time in seconds - INCLUDES A POLICY LOOKUP
STATS_denomination = "C"       		'C is for each CASE
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
call changelog_update("04/22/2024", "Initial version.", "Mark Riegel, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================
'Connecting to BlueZone

'Connecting to BlueZone
EMConnect ""

'Gather case details as applicable
get_county_code
Call check_for_MAXIS(False)
CALL MAXIS_case_number_finder(MAXIS_case_number)

'Initial dialog to gather case details
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 176, 65, "Case Number Dialog"
  EditBox 75, 5, 45, 15, MAXIS_case_number
  EditBox 75, 25, 95, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 75, 45, 45, 15
    CancelButton 125, 45, 45, 15
  Text 20, 10, 50, 10, "Case Number:"
  Text 10, 30, 60, 10, "Worker Signature:"
EndDialog

'Runs the first dialog - which confirms the case number
Do
	Do
		err_msg = ""
		Dialog Dialog1
		cancel_without_confirmation
      	Call validate_MAXIS_case_number(err_msg, "*")
        If trim(worker_signature) = "" THEN err_msg = err_msg & vbCr & "* Sign your case note."
        IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
	Loop until err_msg = ""
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

Call check_for_MAXIS(False)

'Initial dialog to determine food benefit replacement option and action
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 246, 90, "EBT Stolen Benefits"
  ButtonGroup ButtonPressed
    OkButton 75, 70, 45, 15
    CancelButton 125, 70, 45, 15
  EditBox 75, 5, 45, 15, MAXIS_case_number
  EditBox 75, 45, 95, 15, worker_signature
  Text 10, 10, 50, 10, "Case Number:"
  Text 10, 50, 60, 10, "Worker Signature:"
  Text 10, 30, 65, 10, "Action to complete:"
  DropListBox 75, 25, 165, 20, "Select one:"+chr(9)+"CASE/NOTE Information about Request"+chr(9)+"Send SPEC/MEMO regarding Request", action_step
EndDialog

'Dialog validation
Do
  Do
    err_msg = ""
    DIALOG dialog1
    Cancel_confirmation
    If action_step = "Select one:" Then err_msg = err_msg & vbCr & "* You must select the action to complete for EBT stolen benefits replacement request." 
    Call validate_MAXIS_case_number(err_msg, "*")
    If trim(worker_signature) = "" then err_msg = err_msg & vbNewLine & "* You must provide a worker signature."
    IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
  LOOP UNTIL err_msg = ""									'loops until all errors are resolved
  CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'If CASE/NOTE option selected
If action_step = "CASE/NOTE Information about Request" Then

  Dialog1 = "" 'Blanking out previous dialog detail
  BeginDialog Dialog1, 0, 0, 346, 400, "EBT Stolen Benefits"
  EditBox 100, 30, 50, 15, MAXIS_case_number
  DropListBox 100, 45, 105, 15, "Select one:"+chr(9)+"Pending DHS-8557"+chr(9)+"Pending DHS-3335A"+chr(9)+"Request Approved"+chr(9)+"Request Denied", request_status
  EditBox 135, 70, 45, 15, date_client_reported
  EditBox 135, 85, 45, 15, stolen_benefit_type
  EditBox 135, 100, 45, 15, date_client_discovered_stolen_benefit
  EditBox 135, 115, 45, 15, dollar_value_stolen_benefit
  EditBox 135, 130, 200, 15, illegal_method_used
  EditBox 135, 170, 45, 15, date_dhs_8557_sent
  CheckBox 185, 170, 140, 15, "Check here to create TIKL for DHS-8557", dhs_8557_tikl_check
  EditBox 135, 185, 45, 15, date_dhs_8557_signed
  EditBox 135, 200, 45, 15, date_dhs_8557_returned
  DropListBox 135, 215, 200, 15, "Select one:"+chr(9)+"Awaiting return of DHS-8557"+chr(9)+"DHS-8557 never returned"+chr(9)+"DHS-8557 returned - doesn't meet digital theft definition"+chr(9)+"DHS-8557 returned - fraud findings don't support digital theft definition"+chr(9)+"DHS-8557 returned - fraud findings do support digital theft definition", dhs_8557_action_taken
  EditBox 160, 250, 45, 15, date_dhs_3335A_sent_to_fraud_investigator
  EditBox 160, 265, 175, 15, results_of_dhs_3335A
  EditBox 160, 280, 175, 15, stolen_benefit_validation
  DropListBox 10, 310, 95, 40, "Select one:"+chr(9)+"Yes"+chr(9)+"No"+chr(9)+"N/A", previous_benefit_replacement
  EditBox 10, 350, 325, 15, past_replacement_benefits_details
  EditBox 75, 380, 165, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 250, 380, 35, 15
    CancelButton 290, 380, 35, 15
  Text 5, 5, 265, 10, "When a client reports that they believe their benefits have been digitally stolen"
  Text 5, 15, 130, 10, "(see CM0024.06.03.10 or TE02.11.127)"
  Text 10, 35, 50, 10, "Case number:"
  Text 10, 45, 75, 10, "Decision on Request:"
  GroupBox 5, 60, 335, 90, "EBT Stolen Benefit Details"
  Text 10, 75, 115, 10, "Date client reported stolen benefit:"
  Text 10, 90, 75, 10, "Type of benefit stolen:"
  Text 10, 105, 125, 10, "Date client discovered stolen benefit:"
  Text 10, 120, 100, 10, "Dollar value of stolen benefit:"
  Text 10, 135, 125, 10, "Illegal method used to steal benefits:"
  GroupBox 5, 155, 335, 80, "Replacement of Stolen EBT Benefits (DHS-8557)"
  Text 10, 170, 115, 10, "Date DHS-8557 mailed to the client:"
  Text 10, 185, 120, 10, "Date DHS-8557 signed by the client:"
  Text 10, 200, 115, 10, "Date client returned DHS-8557:"
  Text 10, 215, 115, 10, "Action Taken on DHS-8557:"
  GroupBox 5, 240, 335, 130, "Fraud Prevention Investigation (FPI) Referral (DHS-3335A)"
  Text 10, 255, 145, 10, "Date DHS-3335A sent to Fraud Investigator:"
  Text 10, 270, 90, 10, "Results of the DHS-3335A:"
  Text 10, 285, 140, 10, "How stolen benefit replacement validated:"
  Text 10, 300, 330, 10, "Has client requested and received previous benefits replacement between 10/1/22 through present?"
  Text 10, 330, 295, 20, "If yes, how many other replacements due to stolen EBT benefits has this client received? (Indicate # and and benefit type replaced)"
  Text 10, 385, 60, 10, "Worker Signature: "
  EndDialog

  'Validation for dialog fields
  Do
    Do
      err_msg = ""
      DIALOG dialog1
      Cancel_confirmation
      'Validation for Loss Information group box
      Call validate_MAXIS_case_number(err_msg, "*")
      ' If IsDate(trim(loss_date)) <> True Then err_msg = err_msg & vbCr & "* You must enter the date the loss occurred."
      ' If trim(amount_loss) = "" Then err_msg = err_msg & vbCr & "* You must enter the dollar amount of food loss."
      ' If trim(disaster_type) = "" Then err_msg = err_msg & vbCr & "* You must enter the type of disaster - power outage, fire, etc."
      ' If trim(how_verif) = "" Then err_msg = err_msg & vbCr & "* You must explain how the disaster was verified - news reports, Social Worker, Red Cross, etc."
      
      ' 'Validation for Nonreceipt/Replacement Affidavit (DHS-1609) group box
      ' If request_status = "Pending Complete DHS-1609" and (dhs1609_sig_date <> "N/A" or dhs1609_done_date <> "N/A" or dhs1609_rcvd_date <> "N/A") Then err_msg = err_msg & vbCr & "* You indicated that the request is Pending a Complete DHS-1609. Enter N/A for each of the DHS-1609 Dates."
      ' If (request_status = "Pending Verification(s)" or request_status = "All Required Request Info Received") and (IsDate(trim(dhs1609_sig_date)) <> True or IsDate(trim(dhs1609_done_date)) <> True or IsDate(trim(dhs1609_rcvd_date)) <> True) Then err_msg = err_msg & vbCr & "* You must enter dates in the for each of the DHS-1609 fields."

      ' 'Validation for County review group box
      ' If IsDate(trim(report_date)) <> TRUE Then err_msg = err_msg & vbCr & "* You must enter the date the resident reported the loss to the county."
      ' If request_status = "All Required Request Info Received" and IsDate(trim(loss_verified_date)) <> TRUE Then err_msg = err_msg & vbCr & "* You must enter the date the county verified the loss."
      ' If request_status = "Select One:" THEN err_msg = err_msg & vbCr & "* You must select the status of the request."
      ' If request_decision = "Select One:" THEN err_msg = err_msg & vbCr & "* You must select the status of the replacement."
      ' IF rei_replacement = "Select One:" THEN err_msg = err_msg & vbCr & "* You must make a selection for the Replace as REI field."

      ' If request_status = "Pending Verification(s)" and trim(verif_needed) = "" THEN err_msg = err_msg & vbCr & "* You must fill out the verifications needed field."
      ' If request_status = "All Required Request Info Received" and (request_decision = "Pending Additional Information" or request_decision = "Select One:") THEN err_msg = err_msg & vbCr & "* You indicated that all required request information has been received so you must indicate the decision made on the decision on request dropdown."
      ' If request_status <> "All Required Request Info Received" and request_decision <> "Pending Additional Information" THEN err_msg = err_msg & vbCr & "* You indicated that information regarding the request is pending so you cannot select the request approved option for the decision on request dropdown."
      ' If request_decision = "Request Denied" and trim(denial_reason) = "" THEN err_msg = err_msg & vbCr & "* You indicated the request is denied so you must provide a explanation for the denial."
      ' If request_decision = "Request Approved" and TSS_BENE_sent_checkbox = UNCHECKED THEN err_msg = err_msg & vbCr & "* You indicated the request is approved so you must submit the request to the TSS BENE Unit and check the box to confirm the request has been sent."
      ' If request_decision <> "Request Approved" and TSS_BENE_sent_checkbox = CHECKED THEN err_msg = err_msg & vbCr & "* The TSS BENE checkbox should only be checked if the request has been approved."

      IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
      IF worker_signature = "" THEN err_msg = err_msg & vbCr & "* Please sign your case note."
    LOOP UNTIL err_msg = ""									'loops until all errors are resolved
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
  Loop until are_we_passworded_out = false					'loops until user passwords back in

  'Navigate to CASE/NOTE to document food benefits replacement request
  start_a_blank_case_note

  If request_status = "Pending DHS-8557" or request_status = "Pending DHS-3335A" Then
    CALL write_variable_in_Case_Note("--EBT Stolen Benefits - " & request_status & "--")
    CALL write_bullet_and_variable_in_Case_Note("Decision on Request", request_status)
    CALL write_bullet_and_variable_in_Case_Note("Date client reported the stolen benefit to the county", date_client_reported)
    CALL write_bullet_and_variable_in_Case_Note("Type of benefit stolen", stolen_benefit_type)
    CALL write_bullet_and_variable_in_Case_Note("Decision on Request", request_status)
    CALL write_bullet_and_variable_in_Case_Note("Date client discovered stolen benefit", date_client_discovered_stolen_benefit)
    CALL write_bullet_and_variable_in_Case_Note("Dollar value of the stolen benefit", dollar_value_stolen_benefit)
    CALL write_bullet_and_variable_in_Case_Note("Type of illegal method used to steal benefits", illegal_method_used)
    CALL write_bullet_and_variable_in_Case_Note("Date DHS-8557 mailed to client", date_dhs_8557_sent)
    If dhs_8557_tikl_check = CHECKED Then CALL write_variable_in_Case_Note("TIKL created for return of DHS-8557 form.")
    CALL write_bullet_and_variable_in_Case_Note("Date DHS-8557 signed by client", date_dhs_8557_signed)
    CALL write_bullet_and_variable_in_Case_Note("Date DHS-3335A sent to Fraud Investigator", date_dhs_3335A_sent_to_fraud_investigator)
    CALL write_bullet_and_variable_in_Case_Note("Result of the DHS-3335A", results_of_dhs_3335A)
    CALL write_bullet_and_variable_in_Case_Note("How the stolen benefit replacement was validated", stolen_benefit_validation)
    CALL write_bullet_and_variable_in_Case_Note("Requested & received previous benefits replacement (10/1/22 to present)", previous_benefit_replacement)
    If previous_benefit_replacement = "Yes" Then CALL write_bullet_and_variable_in_Case_Note("Details of previous benefits replacements", past_replacement_benefits_details)

    script_end_procedure_with_error_report("The case note has been created. Please run this script again when a decision is made on the request.")

  ElseIf request_status = "Request Approved" Then
    CALL write_variable_in_Case_Note("--EBT Stolen Benefits - " & request_status & "--")
    CALL write_bullet_and_variable_in_Case_Note("Decision on Request", request_status)
    CALL write_bullet_and_variable_in_Case_Note("Date client reported the stolen benefit to the county", date_client_reported)
    CALL write_bullet_and_variable_in_Case_Note("Type of benefit stolen", stolen_benefit_type)
    CALL write_bullet_and_variable_in_Case_Note("Decision on Request", request_status)
    CALL write_bullet_and_variable_in_Case_Note("Date client discovered stolen benefit", date_client_discovered_stolen_benefit)
    CALL write_bullet_and_variable_in_Case_Note("Dollar value of the stolen benefit", dollar_value_stolen_benefit)
    CALL write_bullet_and_variable_in_Case_Note("Type of illegal method used to steal benefits", illegal_method_used)
    CALL write_bullet_and_variable_in_Case_Note("Date DHS-8557 mailed to client", date_dhs_8557_sent)
    If dhs_8557_tikl_check = CHECKED Then CALL write_variable_in_Case_Note("TIKL created for return of DHS-8557 form.")
    CALL write_bullet_and_variable_in_Case_Note("Date DHS-8557 signed by client", date_dhs_8557_signed)
    CALL write_bullet_and_variable_in_Case_Note("Date DHS-3335A sent to Fraud Investigator", date_dhs_3335A_sent_to_fraud_investigator)
    CALL write_bullet_and_variable_in_Case_Note("Result of the DHS-3335A", results_of_dhs_3335A)
    CALL write_bullet_and_variable_in_Case_Note("How the stolen benefit replacement was validated", stolen_benefit_validation)
    CALL write_bullet_and_variable_in_Case_Note("Requested & received previous benefits replacement (10/1/22 to present)", previous_benefit_replacement)
    If previous_benefit_replacement = "Yes" Then CALL write_bullet_and_variable_in_Case_Note("Details of previous benefits replacements", past_replacement_benefits_details)
    CALL write_variable_in_Case_Note("----")
    CALL write_variable_in_Case_Note(worker_signature)

    script_end_procedure_with_error_report("The case note has been created. You must complete the SIR webform 'Replacement of Stolen EBT Benefits Request Form' reporting the required information to DHS.")

  ElseIf request_status = "Request Denied" Then
    CALL write_variable_in_Case_Note("--EBT Stolen Benefits - " & request_status & "--")
    CALL write_bullet_and_variable_in_Case_Note("Decision on Request", request_status)
    CALL write_bullet_and_variable_in_Case_Note("Date client reported the stolen benefit to the county", date_client_reported)
    CALL write_bullet_and_variable_in_Case_Note("Type of benefit stolen", stolen_benefit_type)
    CALL write_bullet_and_variable_in_Case_Note("Decision on Request", request_status)
    CALL write_bullet_and_variable_in_Case_Note("Date client discovered stolen benefit", date_client_discovered_stolen_benefit)
    CALL write_bullet_and_variable_in_Case_Note("Dollar value of the stolen benefit", dollar_value_stolen_benefit)
    CALL write_bullet_and_variable_in_Case_Note("Type of illegal method used to steal benefits", illegal_method_used)
    CALL write_bullet_and_variable_in_Case_Note("Date DHS-8557 mailed to client", date_dhs_8557_sent)
    If dhs_8557_tikl_check = CHECKED Then CALL write_variable_in_Case_Note("TIKL created for return of DHS-8557 form.")
    CALL write_bullet_and_variable_in_Case_Note("Date DHS-8557 signed by client", date_dhs_8557_signed)
    CALL write_bullet_and_variable_in_Case_Note("Date DHS-3335A sent to Fraud Investigator", date_dhs_3335A_sent_to_fraud_investigator)
    CALL write_bullet_and_variable_in_Case_Note("Result of the DHS-3335A", results_of_dhs_3335A)
    CALL write_bullet_and_variable_in_Case_Note("How the stolen benefit replacement was validated", stolen_benefit_validation)
    CALL write_bullet_and_variable_in_Case_Note("Requested & received previous benefits replacement (10/1/22 to present)", previous_benefit_replacement)
    If previous_benefit_replacement = "Yes" Then CALL write_bullet_and_variable_in_Case_Note("Details of previous benefits replacements", past_replacement_benefits_details)
    CALL write_variable_in_Case_Note("----")
    CALL write_variable_in_Case_Note(worker_signature)
    
    'Add dialog to allow for edits to CASE/NOTE before going to SPEC/MEMO

    Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 246, 75, "EBT Stolen Benefits"
    ButtonGroup ButtonPressed
      CancelButton 195, 55, 45, 15
    Text 10, 5, 230, 35, "Please review the generated CASE/NOTE and make any needed changes. Once the CASE/NOTE is finalized, the script will generate a SPEC/MEMO. Please click the button to save the CASE/NOTE and navigate to SPEC/MEMO. "
    ButtonGroup ButtonPressed
      PushButton 125, 55, 65, 15, "Save CASE/NOTE", save_case_note_btn
    EndDialog

    'Dialog validation
    Do
      DIALOG dialog1
      Cancel_confirmation
      CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    Loop until are_we_passworded_out = false					'loops until user passwords back in

    'Add handling to save CASE/NOTE (if needed) and then get to SPEC/MEMO
    EMReadScreen case_note_panel_check, 8, 1, 72
    If case_note_panel_check = "FMCAMAM2" Then
      'CASE/NOTE still in edit mode, need to PF3 to save
      PF3
    End If
    
    Call navigate_to_MAXIS_screen("SPEC", "MEMO")
    'Add dialog where worker selects approval or denial reason

    BeginDialog Dialog1, 0, 0, 286, 245, "EBT Stolen Benefits"
    CheckBox 5, 25, 10, 15, "", denial_two_replacements
    CheckBox 5, 50, 10, 15, "", denial_not_reported_30_days
    CheckBox 5, 75, 10, 15, "", denial_dhs_8557_not_returned
    CheckBox 5, 100, 10, 15, "", denial_additional_info_not_provided
    EditBox 20, 120, 260, 15, denial_specific_additional_info_not_provided
    CheckBox 5, 140, 10, 15, "", denial_not_stolen_determination
    CheckBox 5, 165, 10, 15, "", denial_stolen_outside_time_period
    EditBox 5, 200, 275, 15, denial_additional_reasons
    ButtonGroup ButtonPressed
      OkButton 175, 225, 50, 15
      CancelButton 230, 225, 50, 15
    Text 5, 10, 230, 10, "Select the reason(s) for the denial of the benefits replacement request:"
    Text 20, 25, 265, 20, "You received two replacements of stolen electronic benefits in the Federal Fiscal Year (FFY). The FFY runs from October 1 through September 30"
    Text 20, 50, 265, 20, "You did not report your stolen benefits to your county or Tribal Nation worker within 30 business days of discovering the stolen benefits"
    Text 20, 75, 265, 20, "You did not provide the signed Replacement of Stolen EBT Benefits (DHS-8557) form"
    Text 20, 100, 265, 20, "You did not provide additional information to validate the claim as requested by the county/Tribal Nation or DHS staff (list the information not provided below)"
    Text 20, 140, 265, 20, "Hennepin County has determined that the EBT benefits were not stolen because of card skimming, cloning, or similar illegal methods"
    Text 20, 165, 265, 20, "The stolen EBT benefits replacement were stolen outside of the 10/01/2022 - 09/30/2024 time period"
    Text 5, 190, 205, 10, "Provide additional reason(s) for denial besides options above:"
    EndDialog

    Do
        DIALOG dialog1
        Cancel_confirmation
        CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    Loop until are_we_passworded_out = false					'loops until user passwords back in
  End If

  CALL start_a_new_spec_memo(memo_opened, search_for_arep_and_swkr, forms_to_arep, forms_to_swkr, send_to_other, other_name, other_street, other_city, other_state, other_zip, False)

  Call write_variable_in_SPEC_MEMO("Your request for the replacement of stolen EBT benefits has been denied because:")
  If denial_two_replacements = 1 then Call write_variable_in_SPEC_MEMO("> You received two replacements of stolen electronic benefits in the Federal Fiscal Year (FFY). The FFY runs from October 1 through September 30")
  If denial_not_reported_30_days = 1 then Call write_variable_in_SPEC_MEMO("> You did not report your stolen benefits to your county or Tribal Nation worker within 30 business days of discovering the stolen benefits")
  If denial_dhs_8557_not_returned = 1 then Call write_variable_in_SPEC_MEMO("> You did not provide the signed Replacement of Stolen EBT Benefits (DHS-8557) form")
  If denial_additional_info_not_provided = 1 then Call write_variable_in_SPEC_MEMO("You did not provide additional information to validate the claim as requested by the county/Tribal Nation or DHS staff:")
  If denial_additional_info_not_provided = 1 then Call write_variable_in_SPEC_MEMO("> " & denial_specific_additional_info_not_provided)
  If denial_not_stolen_determination = 1 then Call write_variable_in_SPEC_MEMO("> Hennepin County has determined that the EBT benefits were not stolen because of card skimming, cloning, or similar illegal methods")
  If denial_stolen_outside_time_period = 1 then Call write_variable_in_SPEC_MEMO("> The stolen EBT benefits replacement were stolen outside of the 10/01/2022 - 09/30/2024 time period")
  If trim(denial_additional_reasons) <> "" then Call write_variable_in_SPEC_MEMO("> " & denial_additional_reasons)

End If