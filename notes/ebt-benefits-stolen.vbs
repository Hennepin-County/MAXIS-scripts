'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - EBT BENEFITS STOLEN.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 0                	'manual run time in seconds
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

'Initial dialog to determine EBT stolen benefit action
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 231, 135, "EBT Stolen Benefits - Action on Replacement Request"
  Text 70, 10, 85, 15, MAXIS_case_number
  DropListBox 20, 35, 165, 20, "Select one:"+chr(9)+"CASE/NOTE Information about Request"+chr(9)+"Send SPEC/MEMO regarding Request", action_step
  DropListBox 20, 65, 200, 20, "Select one:"+chr(9)+"Request Denied - Hennepin County Determination"+chr(9)+"Benefits Replaced - DHS TSS BENE Unit Determination", spec_memo_type
  EditBox 85, 90, 120, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 110, 115, 45, 15
    CancelButton 160, 115, 45, 15
  Text 20, 10, 50, 10, "Case Number:"
  Text 20, 25, 65, 10, "Action to complete:"
  Text 20, 55, 125, 10, "If sending SPEC/MEMO, indicate type:"
  Text 20, 95, 60, 10, "Worker Signature:"
EndDialog

'Dialog validation
Do
  Do
    err_msg = ""
    DIALOG dialog1
    Cancel_confirmation
    If action_step = "Select one:" Then err_msg = err_msg & vbCr & "* You must select the action to complete for EBT stolen benefits replacement request." 
    If action_step = "Send SPEC/MEMO regarding Request" and spec_memo_type = "Select one:" Then err_msg = err_msg & vbCr & "* You must select the type of SPEC/MEMO to send." 
    Call validate_MAXIS_case_number(err_msg, "*")
    If action_step <> "Send SPEC/MEMO regarding Request" and spec_memo_type <> "Select one:" then err_msg = err_msg & vbCr & "* You can only select the SPEC/MEMO type if you select the 'Send SPEC/MEMO regarding Request' option for the action to complete." 
    If trim(worker_signature) = "" then err_msg = err_msg & vbNewLine & "* You must provide a worker signature."
    IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
  LOOP UNTIL err_msg = ""									'loops until all errors are resolved
  CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'If CASE/NOTE option selected
If action_step = "CASE/NOTE Information about Request" Then

  Dialog1 = "" 'Blanking out previous dialog detail
  BeginDialog Dialog1, 0, 0, 386, 400, "EBT Stolen Benefits - Replacement Request Details"
    EditBox 100, 25, 50, 15, MAXIS_case_number
    DropListBox 100, 40, 105, 15, "Select one:"+chr(9)+"Pending DHS-8557"+chr(9)+"Pending DHS-3335A"+chr(9)+"Request Approved"+chr(9)+"Request Denied", request_status
    EditBox 135, 70, 45, 15, date_client_reported
    EditBox 135, 85, 240, 15, stolen_benefit_type
    EditBox 135, 100, 45, 15, date_client_discovered_stolen_benefit
    EditBox 135, 115, 45, 15, dollar_value_stolen_benefit
    EditBox 135, 130, 240, 15, illegal_method_used
    EditBox 135, 170, 45, 15, date_dhs_8557_sent
    CheckBox 185, 170, 140, 15, "Check here to create TIKL for DHS-8557", dhs_8557_tikl_check
    EditBox 135, 185, 45, 15, date_dhs_8557_signed
    EditBox 135, 200, 45, 15, date_dhs_8557_returned
    DropListBox 135, 215, 240, 15, "Select one:"+chr(9)+"Awaiting return of DHS-8557"+chr(9)+"DHS-8557 never returned"+chr(9)+"DHS-8557 returned - doesn't meet digital theft definition"+chr(9)+"DHS-8557 returned - fraud findings don't support digital theft definition"+chr(9)+"DHS-8557 returned - fraud findings do support digital theft definition", dhs_8557_action_taken
    EditBox 160, 250, 45, 15, date_dhs_3335A_sent_to_fraud_investigator
    EditBox 160, 265, 215, 15, results_of_dhs_3335A
    EditBox 160, 280, 215, 15, stolen_benefit_validation
    DropListBox 10, 310, 50, 40, "Select one:"+chr(9)+"Yes"+chr(9)+"No"+chr(9)+"N/A", previous_benefit_replacement
    EditBox 10, 350, 365, 15, past_replacement_benefits_details
    EditBox 75, 380, 165, 15, worker_signature
    ButtonGroup ButtonPressed
      OkButton 305, 380, 35, 15
      CancelButton 345, 380, 35, 15
    Text 10, 10, 360, 10, "Client reports that they believe their benefits have been digitally stolen (see CM0024.06.03.10 or TE02.11.127)"
    Text 10, 30, 50, 10, "Case number:"
    Text 10, 45, 75, 10, "Decision on Request:"
    GroupBox 5, 60, 375, 90, "EBT Stolen Benefit Details"
    Text 10, 75, 115, 10, "Date client reported stolen benefit:"
    Text 10, 90, 75, 10, "Type of benefit stolen:"
    Text 10, 105, 125, 10, "Date client discovered stolen benefit:"
    Text 10, 120, 100, 10, "Dollar value of stolen benefit:"
    Text 10, 135, 125, 10, "Illegal method used to steal benefits:"
    GroupBox 5, 155, 375, 80, "Replacement of Stolen EBT Benefits (DHS-8557)"
    Text 10, 170, 115, 10, "Date DHS-8557 mailed to the client:"
    Text 10, 185, 120, 10, "Date DHS-8557 signed by the client:"
    Text 10, 200, 115, 10, "Date client returned DHS-8557:"
    Text 10, 215, 115, 10, "Action Taken on DHS-8557:"
    GroupBox 5, 240, 375, 130, "Fraud Prevention Investigation (FPI) Referral (DHS-3335A)"
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

      'Error handling for case number and decision
      Call validate_MAXIS_case_number(err_msg, "*")
      If request_status = "Select one:" Then err_msg = err_msg & vbCr & "* You must select an option for 'Decision on Request'."
      If request_status = "Pending DHS-8557" and dhs_8557_action_taken <> "Awaiting return of DHS-8557" Then err_msg = err_msg & vbCr & "* You indicated that you are awaiting the return of the DHS-8557 and therefore must select the 'Awaiting return of DHS-8557' option for the Action Taken on DHS-8557 dropdown."
      
      'Error handling for EBT Stolen Benefit Details groupbox
      If IsDate(trim(date_client_reported)) <> True or len(trim(date_client_reported)) > 10 or len(trim(date_client_reported)) < 6 Then err_msg = err_msg & vbCr & "* You must enter a date in the MM/DD/YYYY format in the 'Date client reported stolen benefit' field."
      If trim(stolen_benefit_type) = "" Then err_msg = err_msg & vbCr & "* You must fill out the 'Type of benefit stolen' field."
      If IsDate(trim(date_client_discovered_stolen_benefit)) <> True or len(trim(date_client_discovered_stolen_benefit)) > 10 or len(trim(date_client_discovered_stolen_benefit)) < 6 Then err_msg = err_msg & vbCr & "* You must enter a date in the MM/DD/YYYY format in the 'Date client discovered stolen benefit' field."
      If trim(dollar_value_stolen_benefit) = "" Then err_msg = err_msg & vbCr & "* You must fill out the 'Dollar value of the stolen benefit' field."
      If trim(illegal_method_used) = "" Then err_msg = err_msg & vbCr & "* You must fill out the 'Illegal method used to steal benefits' field."
      
      'Validation for DHS-8557 details groupbox
      If dhs_8557_action_taken = "Select one:" Then 
        err_msg = err_msg & vbCr & "* You must select an option for the 'Action Taken on DHS-8557' dropdown."
      ElseIf request_status = "Request Denied" Then
        If dhs_8557_action_taken = "Awaiting return of DHS-8557" Then 
          err_msg = err_msg & vbCr & "* You indicated that you are denying the request and therefore need to select an option other than 'Awaiting return of DHS-8557' option for the 'Action Taken on DHS-8557' dropdown."
        ElseIf dhs_8557_action_taken = "DHS-8557 returned - fraud findings do support digital theft definition" Then 
          err_msg = err_msg & vbCr & "* You indicated that you are denying the request and therefore need to select an option other than 'DHS-8557 returned - fraud findings do support digital theft definition' option for the 'Action Taken on DHS-8557' dropdown."
        ElseIf dhs_8557_action_taken = "DHS-8557 never returned" Then
          If trim(date_dhs_8557_signed) <> "" Then err_msg = err_msg & vbCr & "* You indicated that the client never returned the DHS-8557. Therefore the 'Date DHS-8557 signed by the client' field should be blank."
          If trim(date_dhs_8557_returned) <> "" Then err_msg = err_msg & vbCr & "* You indicated that the client never returned the DHS-8557. Therefore the 'Date client returned DHS-8557' field should be blank."
        ElseIf Instr(dhs_8557_action_taken, "DHS-8557 returned -") Then
          If trim(date_dhs_8557_signed) = "" or (IsDate(date_dhs_8557_signed) <> True or len(trim(date_dhs_8557_signed)) > 10 or len(trim(date_dhs_8557_signed)) < 6) Then err_msg = err_msg & vbCr & "* You indicated that the client returned the DHS-8557. Therefore the 'Date DHS-8557 signed by the client' field cannot be blank and must be in the format MM/DD/YYYY."
          If trim(date_dhs_8557_returned) = "" or (IsDate(date_dhs_8557_returned) <> True or len(trim(date_dhs_8557_returned)) > 10 or len(trim(date_dhs_8557_returned)) < 6) Then err_msg = err_msg & vbCr & "* You indicated that the client returned the DHS-8557. Therefore the 'Date client returned DHS-8557' field cannot be blank and must be in the format MM/DD/YYYY."
        End If
      ElseIf request_status = "Request Approved" Then
        If dhs_8557_action_taken <> "DHS-8557 returned - fraud findings do support digital theft definition" Then err_msg = err_msg & vbCr & "* You indicated that you are approving the request and therefore need to select the option 'DHS-8557 returned - fraud findings do support digital theft definition' option for the 'Action Taken on DHS-8557' dropdown."
        If trim(date_dhs_8557_signed) = "" or (IsDate(date_dhs_8557_signed) <> True or len(trim(date_dhs_8557_signed)) > 10 or len(trim(date_dhs_8557_signed)) < 6) Then err_msg = err_msg & vbCr & "* You indicated that the client returned the DHS-8557. Therefore the 'Date DHS-8557 signed by the client' field cannot be blank and must be in the format MM/DD/YYYY."
        If trim(date_dhs_8557_returned) = "" or (IsDate(date_dhs_8557_returned) <> True or len(trim(date_dhs_8557_returned)) > 10 or len(trim(date_dhs_8557_returned)) < 6) Then err_msg = err_msg & vbCr & "* You indicated that the client returned the DHS-8557. Therefore the 'Date client returned DHS-8557' field cannot be blank and must be in the format MM/DD/YYYY."
      ElseIf request_status = "Pending DHS-8557" Then
        If dhs_8557_action_taken = "Awaiting return of DHS-8557" Then
          If trim(date_dhs_8557_signed) <> "" Then err_msg = err_msg & vbCr & "* You indicated that you are awaiting the return of the DHS-8557. Therefore the 'Date DHS-8557 signed by the client' field should be blank."
          If trim(date_dhs_8557_returned) <> "" Then err_msg = err_msg & vbCr & "* You indicated that you are awaiting the return of the DHS-8557. Therefore the 'Date client returned DHS-8557' field should be blank."
        ElseIf dhs_8557_action_taken <> "Awaiting return of DHS-8557" Then
          err_msg = err_msg & vbCr & "* You indicated that the decision on the benefits replacement request is pending the DHS-8557. Therefore need to select the 'Awaiting return of DHS-8557' option for the 'Action Taken on DHS-8557' dropdown."
        End If
      End If

      'Validation for FPI details groupbox
      If trim(date_dhs_3335A_sent_to_fraud_investigator) <> "" and (IsDate(date_dhs_3335A_sent_to_fraud_investigator) <> True or len(trim(date_dhs_3335A_sent_to_fraud_investigator)) > 10 or len(trim(date_dhs_3335A_sent_to_fraud_investigator)) < 6) Then err_msg = err_msg & vbCr & "* You must enter a date in the format MM/DD/YYYY in the 'Date the DHS-3335A sent to Fraud Investigator' field."

      If request_status = "Pending DHS-3335A" Then 
        If (trim(results_of_dhs_3335A) <> "" or trim(stolen_benefit_validation) <> "" ) Then err_msg = err_msg & vbCr & "* You indicated that the DHS-3335A is pending so the following fields should be blank: 'Results of the DHS-3335A' and 'How stolen benefit replacement validated'."
      ElseIf request_status = "Request Approved" Then
        If trim(date_dhs_3335A_sent_to_fraud_investigator) = "" or trim(results_of_dhs_3335A) = "" or trim(stolen_benefit_validation) = "" or previous_benefit_replacement = "Select one:" then err_msg = err_msg & vbCr & "* You indicated that the benefits replacement request has been approved. Therefore all fields in the 'Fraud Prevention Investigation (FPI) Referral (DHS-3335A) groupbox must be filled out."
        If trim(date_dhs_3335A_sent_to_fraud_investigator) <> "" and (IsDate(date_dhs_3335A_sent_to_fraud_investigator) <> True or len(trim(date_dhs_3335A_sent_to_fraud_investigator)) > 10 or len(trim(date_dhs_3335A_sent_to_fraud_investigator)) < 6) Then err_msg = err_msg & vbCr & "* You must enter a date in the format MM/DD/YYYY in the 'Date the DHS-3335A sent to Fraud Investigator' field."
      ElseIf request_status = "Request Denied" Then
        If dhs_8557_action_taken = "DHS-8557 returned - fraud findings don't support digital theft definition" Then
          If trim(date_dhs_3335A_sent_to_fraud_investigator) = "" or trim(results_of_dhs_3335A) = "" or trim(stolen_benefit_validation) = "" or previous_benefit_replacement = "Select one:" then err_msg = err_msg & vbCr & "* You indicated that the benefits replacement request was denied as the fraud findings did not support the digital theft definition. Therefore all fields in the 'Fraud Prevention Investigation (FPI) Referral (DHS-3335A) groupbox must be filled out with details about the fraud findings."
        End If
      End If

      If previous_benefit_replacement = "Yes" and trim(past_replacement_benefits_details) = "" Then err_msg = err_msg & vbCr & "* You must provide details about the number of previous stolen EBT benefits replacements." 
      If previous_benefit_replacement <> "Yes" and trim(past_replacement_benefits_details) <> "" Then err_msg = err_msg & vbCr & "* The previous benefits details field should only be filled out if the client received a previous replacement of benefits." 

      IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
      IF worker_signature = "" THEN err_msg = err_msg & vbCr & "* Please sign your case note."
    LOOP UNTIL err_msg = ""									'loops until all errors are resolved
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
  Loop until are_we_passworded_out = false					'loops until user passwords back in

  'Navigate to CASE/NOTE to document food benefits replacement request
  start_a_blank_case_note

  If request_status = "Pending DHS-8557" or request_status = "Pending DHS-3335A" Then
    'Worker is awaiting additional information from forms, will CASE/NOTE and then end
    CALL write_variable_in_Case_Note("--EBT Stolen Benefits - " & request_status & "--")
    CALL write_bullet_and_variable_in_Case_Note("Decision on Request", request_status)
    CALL write_bullet_and_variable_in_Case_Note("Date client reported stolen benefit to county", date_client_reported)
    CALL write_bullet_and_variable_in_Case_Note("Type of benefit stolen", stolen_benefit_type)
    CALL write_bullet_and_variable_in_Case_Note("Date client discovered stolen benefit", date_client_discovered_stolen_benefit)
    CALL write_bullet_and_variable_in_Case_Note("Dollar value of stolen benefit", dollar_value_stolen_benefit)
    CALL write_bullet_and_variable_in_Case_Note("Type of illegal method used to steal benefits", illegal_method_used)
    CALL write_bullet_and_variable_in_Case_Note("Date DHS-8557 mailed to client", date_dhs_8557_sent)
    If dhs_8557_tikl_check = CHECKED Then CALL write_variable_in_Case_Note("* TIKL created for return of DHS-8557 form.")
    If trim(date_dhs_8557_signed) <> "" Then CALL write_bullet_and_variable_in_Case_Note("Date DHS-8557 signed by client", date_dhs_8557_signed)
    If trim(date_dhs_8557_returned) <> "" Then CALL write_bullet_and_variable_in_Case_Note("Date DHS-8557 returned", date_dhs_8557_returned)
    CALL write_bullet_and_variable_in_Case_Note("Action Taken on DHS-3335A", dhs_8557_action_taken)
    If trim(date_dhs_3335A_sent_to_fraud_investigator) <> "" Then CALL write_bullet_and_variable_in_Case_Note("Date DHS-3335A sent to Fraud Investigator", date_dhs_3335A_sent_to_fraud_investigator)
    If trim(results_of_dhs_3335A) <> "" Then CALL write_bullet_and_variable_in_Case_Note("Result of DHS-3335A", results_of_dhs_3335A)
    If trim(stolen_benefit_validation) <> "" Then CALL write_bullet_and_variable_in_Case_Note("How stolen benefit replacement validated", stolen_benefit_validation)
    If previous_benefit_replacement <> "Select one:" Then CALL write_bullet_and_variable_in_Case_Note("Requested & received prev. benefits replacement (10/1/22 to present)", previous_benefit_replacement)
    If trim(previous_benefit_replacement) <> "" Then CALL write_bullet_and_variable_in_Case_Note("Details of previous benefits replacements", past_replacement_benefits_details)
    CALL write_variable_in_Case_Note("----")
    CALL write_variable_in_Case_Note(worker_signature)

    'Manual time calculation
    STATS_manualtime = STATS_manualtime + 120

    script_end_procedure_with_error_report("The case note has been created. Please run this script again when a decision is made on the request.")

  ElseIf request_status = "Request Approved" Then
    'Worker has approved the request, will CASE/NOTE and then end
    CALL write_variable_in_Case_Note("--EBT Stolen Benefits - " & request_status & "--")
    CALL write_bullet_and_variable_in_Case_Note("Decision on Request", request_status)
    CALL write_bullet_and_variable_in_Case_Note("Date client reported stolen benefit to county", date_client_reported)
    CALL write_bullet_and_variable_in_Case_Note("Type of benefit stolen", stolen_benefit_type)
    CALL write_bullet_and_variable_in_Case_Note("Date client discovered stolen benefit", date_client_discovered_stolen_benefit)
    CALL write_bullet_and_variable_in_Case_Note("Dollar value of stolen benefit", dollar_value_stolen_benefit)
    CALL write_bullet_and_variable_in_Case_Note("Type of illegal method used to steal benefits", illegal_method_used)
    CALL write_bullet_and_variable_in_Case_Note("Date DHS-8557 mailed to client", date_dhs_8557_sent)
    If dhs_8557_tikl_check = CHECKED Then CALL write_variable_in_Case_Note("* TIKL created for return of DHS-8557 form.")
    If trim(date_dhs_8557_signed) <> "" Then CALL write_bullet_and_variable_in_Case_Note("Date DHS-8557 signed by client", date_dhs_8557_signed)
    If trim(date_dhs_8557_returned) <> "" Then CALL write_bullet_and_variable_in_Case_Note("Date DHS-8557 returned", date_dhs_8557_returned)
    CALL write_bullet_and_variable_in_Case_Note("Action Taken on DHS-3335A", dhs_8557_action_taken)
    If trim(date_dhs_3335A_sent_to_fraud_investigator) <> "" Then CALL write_bullet_and_variable_in_Case_Note("Date DHS-3335A sent to Fraud Investigator", date_dhs_3335A_sent_to_fraud_investigator)
    If trim(results_of_dhs_3335A) <> "" Then CALL write_bullet_and_variable_in_Case_Note("Result of DHS-3335A", results_of_dhs_3335A)
    If trim(stolen_benefit_validation) <> "" Then CALL write_bullet_and_variable_in_Case_Note("How stolen benefit replacement validated", stolen_benefit_validation)
    If previous_benefit_replacement <> "Select one:" Then CALL write_bullet_and_variable_in_Case_Note("Requested & received prev. benefits replacement (10/1/22 to present)", previous_benefit_replacement)
    If trim(previous_benefit_replacement) <> "" Then CALL write_bullet_and_variable_in_Case_Note("Details of previous benefits replacements", past_replacement_benefits_details)
    CALL write_variable_in_Case_Note("----")
    CALL write_variable_in_Case_Note(worker_signature)

    'Manual time calculation
    STATS_manualtime = STATS_manualtime + 240

    script_end_procedure_with_error_report("The case note has been created. You must complete the SIR webform 'Replacement of Stolen EBT Benefits Request Form' reporting the required information to DHS.")

  ElseIf request_status = "Request Denied" Then
    'Worker has denied the request, will CASE/NOTE details, and then generate a SPEC/MEMO
    CALL write_variable_in_Case_Note("--EBT Stolen Benefits - " & request_status & "--")
    CALL write_bullet_and_variable_in_Case_Note("Decision on Request", request_status)
    CALL write_bullet_and_variable_in_Case_Note("Date client reported stolen benefit to county", date_client_reported)
    CALL write_bullet_and_variable_in_Case_Note("Type of benefit stolen", stolen_benefit_type)
    CALL write_bullet_and_variable_in_Case_Note("Date client discovered stolen benefit", date_client_discovered_stolen_benefit)
    CALL write_bullet_and_variable_in_Case_Note("Dollar value of stolen benefit", dollar_value_stolen_benefit)
    CALL write_bullet_and_variable_in_Case_Note("Type of illegal method used to steal benefits", illegal_method_used)
    CALL write_bullet_and_variable_in_Case_Note("Date DHS-8557 mailed to client", date_dhs_8557_sent)
    If dhs_8557_tikl_check = CHECKED Then CALL write_variable_in_Case_Note("* TIKL created for return of DHS-8557 form.")
    If trim(date_dhs_8557_signed) <> "" Then CALL write_bullet_and_variable_in_Case_Note("Date DHS-8557 signed by client", date_dhs_8557_signed)
    If trim(date_dhs_8557_returned) <> "" Then CALL write_bullet_and_variable_in_Case_Note("Date DHS-8557 returned", date_dhs_8557_returned)
    CALL write_bullet_and_variable_in_Case_Note("Action Taken on DHS-3335A", dhs_8557_action_taken)
    If trim(date_dhs_3335A_sent_to_fraud_investigator) <> "" Then CALL write_bullet_and_variable_in_Case_Note("Date DHS-3335A sent to Fraud Investigator", date_dhs_3335A_sent_to_fraud_investigator)
    If trim(results_of_dhs_3335A) <> "" Then CALL write_bullet_and_variable_in_Case_Note("Result of DHS-3335A", results_of_dhs_3335A)
    If trim(stolen_benefit_validation) <> "" Then CALL write_bullet_and_variable_in_Case_Note("How stolen benefit replacement validated", stolen_benefit_validation)
    If previous_benefit_replacement <> "Select one:" Then CALL write_bullet_and_variable_in_Case_Note("Requested & received prev. benefits replacement (10/1/22 to present)", previous_benefit_replacement)
    If trim(previous_benefit_replacement) <> "" Then CALL write_bullet_and_variable_in_Case_Note("Details of previous benefits replacements", past_replacement_benefits_details)
    CALL write_variable_in_Case_Note("SPEC/MEMO sent to client notifying them of denial of benefits replacement request.")
    CALL write_variable_in_Case_Note("----")
    CALL write_variable_in_Case_Note(worker_signature)

    'Manual time calculation
    STATS_manualtime = STATS_manualtime + 240
    
    'Dialog to allow for edits to CASE/NOTE before going to SPEC/MEMO
    Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 246, 75, "EBT Stolen Benefits - Save CASE/NOTE"
      ButtonGroup ButtonPressed
        PushButton 125, 55, 65, 15, "Save CASE/NOTE", save_case_note_btn
        CancelButton 195, 55, 45, 15
      Text 10, 5, 230, 35, "Please review the generated CASE/NOTE and make any needed changes. Once the CASE/NOTE is finalized, the script will generate a SPEC/MEMO. Please click the 'Save CASE/NOTE' and navigate to SPEC/MEMO."
    EndDialog

    'Dialog validation
    Do
      DIALOG dialog1
      Cancel_confirmation
      CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    Loop until are_we_passworded_out = false					'loops until user passwords back in

    'Handling to save CASE/NOTE (if needed) and then navigate to SPEC/MEMO
    EMReadScreen case_note_panel_check, 8, 1, 72
    If case_note_panel_check = "FMCAMAM2" Then
      'CASE/NOTE still in edit mode, need to PF3 to save
      PF3
    End If
    
    'Navigate to SPEC/MEMO to provide details of denial
    Call navigate_to_MAXIS_screen("SPEC", "MEMO")

    Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 286, 245, "EBT Stolen Benefits - Denial Reasons"
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
      Text 20, 50, 265, 20, "You did not report your stolen benefits to your county worker within 30 business days of discovering the stolen benefits"
      Text 20, 75, 265, 20, "You did not provide the signed Replacement of Stolen EBT Benefits (DHS-8557) form"
      Text 20, 100, 265, 20, "You did not provide additional information to validate the claim as requested by the county or DHS staff (list the information not provided below)"
      Text 20, 140, 265, 20, "Hennepin County has determined that the EBT benefits were not stolen because of card skimming, cloning, or similar illegal methods"
      Text 20, 165, 265, 20, "The stolen EBT benefits replacement were stolen outside of the 10/01/2022 - 09/30/2024 time period"
      Text 5, 190, 205, 10, "Provide additional reason(s) for denial besides options above:"
    EndDialog

    DO
      Do
        err_msg = ""    'This is the error message handling
        Dialog Dialog1
        cancel_confirmation
        If (denial_two_replacements + denial_not_reported_30_days + denial_dhs_8557_not_returned + denial_additional_info_not_provided + denial_not_stolen_determination + denial_stolen_outside_time_period = 0) and trim(denial_additional_reasons) = "" Then err_msg = err_msg & vbCr & "* You must check at least one of the reasons for denial."
        If denial_additional_info_not_provided = 1 and (trim(denial_specific_additional_info_not_provided) = "" or len(trim(denial_specific_additional_info_not_provided)) < 5) Then err_msg = err_msg & vbCr & "* You must provide enough details to explain the specific information requested to validate the claim that was not provided."
        If trim(denial_additional_reasons) <> "" and len(trim(denial_additional_reasons)) < 5 Then err_msg = err_msg & vbCr & "* You must provide sufficient details for the additional reasons for the denial."
        IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
      Loop until err_msg = ""
      CALL check_for_password(are_we_passworded_out)
    LOOP UNTIL are_we_passworded_out = false

    CALL start_a_new_spec_memo(memo_opened, search_for_arep_and_swkr, forms_to_arep, forms_to_swkr, send_to_other, other_name, other_street, other_city, other_state, other_zip, False)

    Call write_variable_in_SPEC_MEMO("Your request for the replacement of stolen EBT benefits has been denied because:")
    If denial_two_replacements = 1 then Call write_variable_in_SPEC_MEMO("> You received two replacements of stolen electronic benefits in the Federal Fiscal Year (FFY). The FFY runs from October 1 through September 30")
    If denial_not_reported_30_days = 1 then Call write_variable_in_SPEC_MEMO("> You did not report your stolen benefits to your county worker within 30 business days of discovering the stolen benefits")
    If denial_dhs_8557_not_returned = 1 then Call write_variable_in_SPEC_MEMO("> You did not provide the signed Replacement of Stolen EBT Benefits (DHS-8557) form")
    If denial_additional_info_not_provided = 1 then Call write_variable_in_SPEC_MEMO("> You did not provide additional information to validate the claim as requested by the county or DHS staff:")
    If denial_additional_info_not_provided = 1 then Call write_variable_in_SPEC_MEMO("  > " & denial_specific_additional_info_not_provided)
    If denial_not_stolen_determination = 1 then Call write_variable_in_SPEC_MEMO("> Hennepin County has determined that the EBT benefits were not stolen because of card skimming, cloning, or similar illegal methods")
    If denial_stolen_outside_time_period = 1 then Call write_variable_in_SPEC_MEMO("> The stolen EBT benefits replacement were stolen outside of the 10/01/2022 - 09/30/2024 time period")
    If trim(denial_additional_reasons) <> "" then Call write_variable_in_SPEC_MEMO("> " & denial_additional_reasons)

    'Manual time calculation
    STATS_manualtime = STATS_manualtime + 150

    script_end_procedure_with_error_report("The SPEC/MEMO has been created. You must complete the SIR webform 'Replacement of Stolen EBT Benefits Request Form' reporting the required information to DHS.")
  End If
End If

If action_step = "Send SPEC/MEMO regarding Request" Then
  If spec_memo_type = "Request Denied - Hennepin County Determination" Then
    'Handling for sending a SPEC/MEMO after denying request
    Call navigate_to_MAXIS_screen("SPEC", "MEMO")

    Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 286, 245, "EBT Stolen Benefits - Denial Reasons"
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
      Text 20, 50, 265, 20, "You did not report your stolen benefits to your county worker within 30 business days of discovering the stolen benefits"
      Text 20, 75, 265, 20, "You did not provide the signed Replacement of Stolen EBT Benefits (DHS-8557) form"
      Text 20, 100, 265, 20, "You did not provide additional information to validate the claim as requested by the county or DHS staff (list the information not provided below)"
      Text 20, 140, 265, 20, "Hennepin County has determined that the EBT benefits were not stolen because of card skimming, cloning, or similar illegal methods"
      Text 20, 165, 265, 20, "The stolen EBT benefits replacement were stolen outside of the 10/01/2022 - 09/30/2024 time period"
      Text 5, 190, 205, 10, "Provide additional reason(s) for denial besides options above:"
    EndDialog

    DO
      Do
        err_msg = ""    'This is the error message handling
        Dialog Dialog1
        cancel_confirmation
        If (denial_two_replacements + denial_not_reported_30_days + denial_dhs_8557_not_returned + denial_additional_info_not_provided + denial_not_stolen_determination + denial_stolen_outside_time_period = 0) and trim(denial_additional_reasons) = "" Then err_msg = err_msg & vbCr & "* You must check at least one of the reasons for denial."
        If denial_additional_info_not_provided = 1 and (trim(denial_specific_additional_info_not_provided) = "" or len(trim(denial_specific_additional_info_not_provided)) < 5) Then err_msg = err_msg & vbCr & "* You must provide enough details to explain the specific information requested to validate the claim that was not provided."
        If trim(denial_additional_reasons) <> "" and len(trim(denial_additional_reasons)) < 5 Then err_msg = err_msg & vbCr & "* You must provide sufficient details for the additional reasons for the denial."
        IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
      Loop until err_msg = ""
      CALL check_for_password(are_we_passworded_out)
    LOOP UNTIL are_we_passworded_out = false

    CALL start_a_new_spec_memo(memo_opened, search_for_arep_and_swkr, forms_to_arep, forms_to_swkr, send_to_other, other_name, other_street, other_city, other_state, other_zip, False)

    Call write_variable_in_SPEC_MEMO("Your request for the replacement of stolen EBT benefits has been denied because:")
    If denial_two_replacements = 1 then Call write_variable_in_SPEC_MEMO("> You received two replacements of stolen electronic benefits in the Federal Fiscal Year (FFY). The FFY runs from October 1 through September 30")
    If denial_not_reported_30_days = 1 then Call write_variable_in_SPEC_MEMO("> You did not report your stolen benefits to your county worker within 30 business days of discovering the stolen benefits")
    If denial_dhs_8557_not_returned = 1 then Call write_variable_in_SPEC_MEMO("> You did not provide the signed Replacement of Stolen EBT Benefits (DHS-8557) form")
    If denial_additional_info_not_provided = 1 then Call write_variable_in_SPEC_MEMO("> You did not provide additional information to validate the claim as requested by the county or DHS staff:")
    If trim(denial_specific_additional_info_not_provided) <> "" then Call write_variable_in_SPEC_MEMO("  > " & denial_specific_additional_info_not_provided)
    If denial_not_stolen_determination = 1 then Call write_variable_in_SPEC_MEMO("> Hennepin County has determined that the EBT benefits were not stolen because of card skimming, cloning, or similar illegal methods")
    If denial_stolen_outside_time_period = 1 then Call write_variable_in_SPEC_MEMO("> The stolen EBT benefits replacement were stolen outside of the 10/01/2022 - 09/30/2024 time period")
    If trim(denial_additional_reasons) <> "" then Call write_variable_in_SPEC_MEMO("> " & denial_additional_reasons)

    'Save the SPEC/MEMO
    PF4

    'CASE/NOTE that SPEC/MEMO sent
    start_a_blank_case_note

    'Worker is awaiting additional information from forms, will CASE/NOTE and then end
    CALL write_variable_in_Case_Note("--EBT Stolen Benefits - Denial SPEC/MEMO Sent--")
    CALL write_variable_in_Case_Note("SPEC/MEMO sent to client informing them of denial of benefits replacement request. Information provided in SPEC/MEMO provided below:")
    CALL write_variable_in_Case_Note("----")
    Call write_variable_in_Case_Note("Your request for the replacement of stolen EBT benefits has been denied because:")
    If denial_two_replacements = 1 then Call write_variable_in_Case_Note("> You received two replacements of stolen electronic benefits in the Federal Fiscal Year (FFY). The FFY runs from October 1 through September 30")
    If denial_not_reported_30_days = 1 then Call write_variable_in_Case_Note("> You did not report your stolen benefits to your county worker within 30 business days of discovering the stolen benefits")
    If denial_dhs_8557_not_returned = 1 then Call write_variable_in_Case_Note("> You did not provide the signed Replacement of Stolen EBT Benefits (DHS-8557) form")
    If denial_additional_info_not_provided = 1 then Call write_variable_in_Case_Note("> You did not provide additional information to validate the claim as requested by the county or DHS staff:")
    If trim(denial_specific_additional_info_not_provided) <> "" then Call write_variable_in_Case_Note("  > " & denial_specific_additional_info_not_provided)
    If denial_not_stolen_determination = 1 then Call write_variable_in_Case_Note("> Hennepin County has determined that the EBT benefits were not stolen because of card skimming, cloning, or similar illegal methods")
    If denial_stolen_outside_time_period = 1 then Call write_variable_in_Case_Note("> The stolen EBT benefits replacement were stolen outside of the 10/01/2022 - 09/30/2024 time period")
    If trim(denial_additional_reasons) <> "" then Call write_variable_in_Case_Note(denial_additional_reasons)
    CALL write_variable_in_Case_Note("----")
    CALL write_variable_in_Case_Note(worker_signature)    

    'Manual time calculation
    STATS_manualtime = STATS_manualtime + 270

    script_end_procedure_with_error_report("The SPEC/MEMO has been created. A CASE/NOTE regarding the SPEC/MEMO was also created. You must complete the SIR webform 'Replacement of Stolen EBT Benefits Request Form' reporting the required information to DHS.")

  ElseIf spec_memo_type = "Benefits Replaced - DHS TSS BENE Unit Determination" Then
    'Handling for sending a SPEC/MEMO after TSS BENE Unit has approved request
    Call navigate_to_MAXIS_screen("SPEC", "MEMO")

    Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 291, 165, "EBT Stolen Benefits - Approval Details"
      EditBox 220, 35, 55, 15, stolen_benefits_replaced_amount
      EditBox 220, 55, 55, 15, stolen_benefits_not_replaced_amount
      EditBox 5, 90, 270, 15, stolen_benefit_not_replaced_explanation
      EditBox 5, 120, 270, 15, additional_info_spec_memo
      ButtonGroup ButtonPressed
        OkButton 170, 145, 50, 15
        CancelButton 225, 145, 50, 15
      Text 5, 10, 265, 20, "For requests approved and replaced by the TSS BENE Team, a SPEC:MEMO must be sent to the client. Enter information about the replacement request: "
      Text 5, 40, 190, 10, "Dollar amount of stolen benefits returned to EBT account:"
      Text 5, 60, 215, 10, "If applicable, enter dollar amount of stolen benefits NOT replaced:"
      Text 5, 80, 265, 10, "If the full amount of the stolen benefit cannot be replaced, provide explanation:"
      Text 5, 110, 265, 10, "Enter additional information to include in SPEC:MEMO:"
    EndDialog  

    DO
      Do
        err_msg = ""
        Dialog Dialog1
        cancel_confirmation
        
        If trim(stolen_benefits_replaced_amount) = "" Then err_msg = err_msg & vbCr & "* You must enter the dollar amount of stolen benefits returned to EBT account."
        If IsNumeric(trim(stolen_benefits_replaced_amount)) = False Then err_msg = err_msg & vbCr & "* You must enter the dollar amount of stolen benefits returned to EBT account in a number format."
        If IsNumeric(trim(stolen_benefits_not_replaced_amount)) = False and trim(stolen_benefits_not_replaced_amount) <> "" Then err_msg = err_msg & vbCr & "* You must enter the dollar amount of stolen benefits not replaced in a number format."
        If trim(stolen_benefits_not_replaced_amount) <> "" and trim(stolen_benefit_not_replaced_explanation) = "" Then err_msg = err_msg & vbCr & "* You must provide an explanation for why the full amount of stolen benefits was not replaced."
        IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
      Loop until err_msg = ""
      CALL check_for_password(are_we_passworded_out)
    LOOP UNTIL are_we_passworded_out = false

    CALL start_a_new_spec_memo(memo_opened, search_for_arep_and_swkr, forms_to_arep, forms_to_swkr, send_to_other, other_name, other_street, other_city, other_state, other_zip, False)

    Call write_variable_in_SPEC_MEMO("> Your request for the replacement of stolen EBT benefits has been approved.")
    Call write_variable_in_SPEC_MEMO("> The amount of stolen benefits that have been returned to your EBT account is: $" & stolen_benefits_replaced_amount)
    If stolen_benefits_not_replaced_amount <> "" Then Call write_variable_in_SPEC_MEMO("> The full amount of stolen benefit cannot be replaced because: " & stolen_benefit_not_replaced_explanation & "." & " The amount not replaced is $" & stolen_benefits_not_replaced_amount & ".")
    Call write_variable_in_SPEC_MEMO("> If you have not reported your compromised EBT card stolen and had a replacement card issued, this was done for you prior to benefit replacement.")
    Call write_variable_in_SPEC_MEMO("> If you do not agree with this decision, see the reverse side of this notice for your appeal rights. Contact your county worker to request an appeal.")

    'Save the SPEC/MEMO
    PF4

    'CASE/NOTE that SPEC/MEMO sent
    start_a_blank_case_note

    'Worker is awaiting additional information from forms, will CASE/NOTE and then end
    CALL write_variable_in_Case_Note("--EBT Stolen Benefits - SPEC/MEMO Sent--")
    CALL write_variable_in_Case_Note("SPEC/MEMO sent to client informing them of approval of benefits replacement request. Information provided in SPEC/MEMO provided below:")
    CALL write_variable_in_Case_Note("----")
    Call write_variable_in_Case_Note("Your request for the replacement of stolen EBT benefits has been approved.")
    Call write_variable_in_Case_Note("> The amount of stolen benefits that have been returned to your EBT account is: $" & stolen_benefits_replaced_amount)
    If stolen_benefits_not_replaced_amount <> "" Then Call write_variable_in_Case_Note("> The full amount of stolen benefit cannot be replaced because: " & stolen_benefit_not_replaced_explanation & "." & " The amount not replaced is $" & stolen_benefits_not_replaced_amount & ".")
    Call write_variable_in_Case_Note("> If you have not reported your compromised EBT card stolen and had a replacement card issued, this was done for you prior to benefit replacement.")
    Call write_variable_in_Case_Note("> If you do not agree with this decision, see the reverse side of this notice for your appeal rights. Contact your county worker to request an appeal.")
    CALL write_variable_in_Case_Note("----")
    CALL write_variable_in_Case_Note(worker_signature)    
    
    'Manual time calculation
    STATS_manualtime = STATS_manualtime + 210

    script_end_procedure_with_error_report("The SPEC/MEMO has been created. A CASE/NOTE regarding the SPEC/MEMO was also created.")
  End If
End If

'----------------------------------------------------------------------------------------------------Closing Project Documentation - Version date 01/12/2023
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------06/13/2024
'--Tab orders reviewed & confirmed----------------------------------------------06/13/2024
'--Mandatory fields all present & reviewed--------------------------------------06/17/2024
'--All variables in dialog match mandatory fields-------------------------------06/17/2024
'--Review dialog names for content and content fit in dialog--------------------06/13/2024
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------06/17/2024
'--CASE:NOTE Header doesn't look funky------------------------------------------06/17/2024
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------06/17/2024
'--write_variable_in_CASE_NOTE function: confirm that proper punctuation is used06/17/2024 
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------06/17/2024
'--MAXIS_background_check reviewed (if applicable)------------------------------06/17/2024
'--PRIV Case handling reviewed -------------------------------------------------06/17/2024
'--Out-of-County handling reviewed----------------------------------------------06/17/2024
'--script_end_procedures (w/ or w/o error messaging)----------------------------06/17/2024
'--BULK - review output of statistics and run time/count (if applicable)--------N/A
'--All strings for MAXIS entry are uppercase vs. lower case (Ex: "X")-----------06/17/2024
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------06/18/2024
'--Incrementors reviewed (if necessary)-----------------------------------------06/18/2024
'--Denomination reviewed -------------------------------------------------------06/18/2024
'--Script name reviewed---------------------------------------------------------06/18/2024
'--BULK - remove 1 incrementor at end of script reviewed------------------------06/18/2024

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------06/17/2024
'--comment Code-----------------------------------------------------------------06/17/2024
'--Update Changelog for release/update------------------------------------------06/17/2024
'--Remove testing message boxes-------------------------------------------------06/17/2024
'--Remove testing code/unnecessary code-----------------------------------------06/17/2024
'--Review/update SharePoint instructions----------------------------------------06/17/2024
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------06/17/2024
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------06/18/2024
'--COMPLETE LIST OF SCRIPTS update policy references----------------------------06/18/2024
'--Complete misc. documentation (if applicable)---------------------------------06/18/2024
'--Update project team/issue contact (if applicable)----------------------------06/17/2024
