'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - DISASTER FOOD REPLACEMENT.vbs"
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
CALL changelog_update("09/19/2022", "Update to ensure Worker Signature is in all scripts that CASE/NOTE.", "MiKayla Handley, Hennepin County") '#316
call changelog_update("11/27/2019", "Initial version.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================
'Connecting to BlueZone

EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number) 'Finds the case number
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)

'Initial dialog to determine food benefit replacement option and action
BeginDialog Dialog1, 0, 0, 201, 150, "Food Benefits Replacement"
  Text 5, 10, 50, 10, "Case number: "
  EditBox 55, 5, 45, 15, MAXIS_case_number
  Text 5, 30, 45, 10, "Footer Month:"
  EditBox 55, 25, 15, 15, MAXIS_footer_month
  Text 85, 30, 20, 10, "Year:"
  EditBox 105, 25, 15, 15, MAXIS_footer_year
  Text 5, 50, 175, 10, "Select the food benefits replacement process below:"
  DropListBox 5, 65, 190, 50, "Select Option:"+chr(9)+"Food Destroyed in Misfortune/Disaster"+chr(9)+"Replacing Stolen EBT Food - Client Notified"+chr(9)+"Replacing Stolen EBT Food - Client Requests", benefit_replacement_process
  Text 5, 90, 165, 10, "Select the action to take below:"
  DropListBox 5, 100, 190, 30, "Select Option:"+chr(9)+"Enter CASE/NOTE regarding request"+chr(9)+"Send SPEC/MEMO regarding decision on request", benefit_replacement_action
  ButtonGroup ButtonPressed
    OkButton 95, 130, 50, 15
    CancelButton 145, 130, 50, 15
EndDialog

'Dialog validation
Do
  Do
    err_msg = ""
    DIALOG dialog1
    Cancel_confirmation
    If benefit_replacement_process = "[Select Process]" or benefit_replacement_action = "[Select Process]" Then err_msg = err_msg & vbCr & "* You must select the food benefit replacement process and action." 
    Call validate_MAXIS_case_number(err_msg, "*")
    Call validate_footer_month_entry(MAXIS_footer_month, MAXIS_footer_year, err_msg, "*")
    IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
  LOOP UNTIL err_msg = ""									'loops until all errors are resolved
  CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in


'To do - verify client vs resident language
'Initial Dialog Box
'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 346, 350, "Replacing Food Destroyed in a Disaster"
  Text 10, 5, 265, 10, "When a client reports food destroyed in a disaster and all requirements are met"
  Text 10, 15, 125, 10, "(see CM0024.06.03.15 or TE02.11.18)"
  GroupBox 10, 25, 330, 75, "Loss Information"
  Text 15, 40, 50, 10, "Case number:"
  EditBox 75, 35, 40, 15, MAXIS_case_number
  Text 120, 40, 45, 10, "Date of Loss:"
  EditBox 170, 35, 50, 15, loss_date
  Text 225, 40, 70, 10, "Amount of Food Loss: "
  EditBox 300, 35, 35, 15, amount_loss
  Text 15, 60, 60, 10, "Type of Disaster:"
  EditBox 75, 55, 175, 15, disaster_type
  Text 255, 55, 80, 20, "(describe the disaster - power outage, fire, etc.)"
  Text 15, 80, 105, 10, "How the disaster was verified:"
  EditBox 120, 75, 130, 15, how_verif
  Text 255, 75, 80, 20, "(news reports, Social Worker, Red Cross etc.)"
  GroupBox 10, 105, 285, 70, "Nonreceipt/Replacement Affidavit (DHS-1609)"
  Text 15, 120, 225, 10, "Date DHS-1609 signed by the client (enter N/A if Pending DHS-1609):"
  EditBox 245, 115, 45, 15, dhs1609_sig_date
  Text 15, 140, 225, 10, "Date DHS-1609 received by county (enter N/A if Pending DHS-1609):"
  EditBox 245, 135, 45, 15, dhs1609_rcvd_date
  Text 15, 160, 230, 10, "Date DHS-1609 completed by county (enter N/A if Pending DHS-1609):"
  EditBox 245, 155, 45, 15, dhs1609_done_date
  GroupBox 10, 175, 330, 145, "County Review"
  Text 15, 190, 115, 10, "Date client reported loss to county:"
  EditBox 130, 185, 50, 15, report_date
  Text 185, 190, 100, 10, "Date loss verified by county:"
  EditBox 285, 185, 50, 15, loss_verified_date
  Text 15, 205, 65, 10, "Status of Request:"
  DropListBox 90, 205, 120, 15, "Select One:"+chr(9)+"Pending Complete DHS-1609"+chr(9)+"Pending Verification(s)"+chr(9)+"All Required Request Info Received", request_status
  Text 225, 205, 55, 10, "Replace as REI: "
  DropListBox 285, 205, 50, 15, "Select One:"+chr(9)+"YES"+chr(9)+"NO"+chr(9)+"N/A", rei_replacement
  Text 15, 225, 75, 10, "Verifications Needed: "
  EditBox 90, 225, 245, 15, verif_needed
  Text 15, 245, 75, 10, "Decision on Request: "
  DropListBox 90, 245, 120, 15, "Select One:"+chr(9)+"Pending Additional Information"+chr(9)+"Request Approved"+chr(9)+"Request Denied", request_decision
  Text 15, 265, 100, 15, "If denied, provide explanation:"
  EditBox 115, 265, 220, 15, denial_reason
  CheckBox 15, 285, 260, 10, "If approved, indicate here that request info has been submitted to TSS BENE", TSS_BENE_sent_checkbox
  ButtonGroup ButtonPressed
    PushButton 15, 300, 105, 15, "TSS BENE Submission Form", TSS_BENE_webform_link
  Text 15, 330, 60, 10, "Worker Signature: "
  EditBox 80, 325, 165, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 255, 325, 35, 15
    CancelButton 295, 325, 35, 15
EndDialog


'Info to the user of what this script currently covers
Do
  Do
    err_msg = ""
    DIALOG dialog1
    Cancel_confirmation
    'Validation for Loss Information group box
    Call validate_MAXIS_case_number(err_msg, "*")
    If IsDate(trim(loss_date)) <> True Then err_msg = err_msg & vbCr & "* You must enter the date the loss occurred."
    If trim(amount_loss) = "" Then err_msg = err_msg & vbCr & "* You must enter the dollar amount of food loss."
    If trim(disaster_type) = "" Then err_msg = err_msg & vbCr & "* You must enter the type of disaster - power outage, fire, etc."
    If trim(how_verif) = "" Then err_msg = err_msg & vbCr & "* You must explain how the disaster was verified - news reports, Social Worker, Red Cross, etc."
    
    'Validation for Nonreceipt/Replacement Affidavit (DHS-1609) group box
    If request_status = "Pending Complete DHS-1609" and (dhs1609_sig_date <> "N/A" or dhs1609_done_date <> "N/A" or dhs1609_rcvd_date <> "N/A") Then err_msg = err_msg & vbCr & "* You indicated that the request is Pending a Complete DHS-1609. Enter N/A for each of the DHS-1609 Dates."
    If (request_status = "Pending Verification(s)" or request_decision = "All Required Request Info Received") and (IsDate(trim(dhs1609_sig_date)) <> True or IsDate(trim(dhs1609_done_date)) <> True or IsDate(trim(dhs1609_rcvd_date)) <> True) Then err_msg = err_msg & vbCr & "* You must enter dates in the for each of the DHS-1609 fields."

    'Validation for County review group box
    If IsDate(trim(report_date)) <> TRUE Then err_msg = err_msg & vbCr & "* You must enter the date the resident reported the loss to the county."
    If IsDate(trim(loss_verified_date)) <> TRUE Then err_msg = err_msg & vbCr & "* You must enter the date the county verified the loss."
    If request_status = "Select One:" THEN err_msg = err_msg & vbCr & "* You must select the status of the request."
    If request_decision = "Select One:" THEN err_msg = err_msg & vbCr & "* You must select the status of the replacement."
    IF rei_replacement = "Select One:" THEN err_msg = err_msg & vbCr & "* You must make a selection for the Replace as REI field."

    If request_status = "Pending Verification(s)" and trim(verif_needed) = "" THEN err_msg = err_msg & vbCr & "* You must fill out the verifications needed field."
    If request_status = "All Required Request Info Received" and (request_decision = "Pending Additional Information" or request_decision = "Select One:") THEN err_msg = err_msg & vbCr & "* You indicated that all required request information has been received so you must indicate the decision made on the decision on request dropdown."
    If request_status <> "All Required Request Info Received" and request_decision <> "Pending Additional Information" THEN err_msg = err_msg & vbCr & "* You indicated that information regarding the request is pending so you cannot select the request approved option for the decision on request dropdown."
    If request_decision = "Request Denied" and trim(denial_reason) = "" THEN err_msg = err_msg & vbCr & "* You indicated the request is denied so you must provide a explanation for the denial."
    If request_decision = "Request Approved" and TSS_BENE_sent_checkbox = UNCHECKED THEN err_msg = err_msg & vbCr & "* You indicated the request is approved so you must submit the request to the TSS BENE Unit and check the box to confirm the request has been sent."
    If request_decision <> "Request Approved" and TSS_BENE_sent_checkbox = CHECKED THEN err_msg = err_msg & vbCr & "* The TSS BENE checkbox should only be checked if the request has been approved."

    IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
    IF worker_signature = "" THEN err_msg = err_msg & vbCr & "* Please sign your case note."
  LOOP UNTIL err_msg = ""									'loops until all errors are resolved
  CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in


start_a_blank_case_note
'writes case note for Baby Born
CALL write_variable_in_Case_Note("--Food Destroyed in a Disaster Reported - " & replacement_status & "--")
CALL write_bullet_and_variable_in_Case_Note("Type of Disaster", disaster_type)
CALL write_bullet_and_variable_in_Case_Note("How the disaster was verified", how_verif)
CALL write_bullet_and_variable_in_Case_Note("Date client reported the loss of food to county", report_date)
CALL write_bullet_and_variable_in_Case_Note("Date of Loss", loss_date)
CALL write_bullet_and_variable_in_Case_Note("Amount of Food Loss", amount_loss)
CALL write_bullet_and_variable_in_Case_Note("Replace as REI", rei_replacement)
IF TSS_BENE_sent_checkbox <> CHECKED THEN CALL write_bullet_and_variable_in_Case_Note("Stauts of Request", replacement_status)
CALL write_bullet_and_variable_in_Case_Note("Reason for Denial", denial_reason)
CALL write_bullet_and_variable_in_Case_Note("Verifcations Requested", verif_needed)
IF TSS_BENE_sent_checkbox = CHECKED THEN CALL write_variable_in_Case_Note("* Submitted a TSS BENE request (webform) through SIR")
CALL write_variable_in_Case_Note("Nonreceipt/Replacement Affidavit (DHS-1609)")
CALL write_variable_in_Case_Note(" Date DHS-1609 was signed by the client: " & dhs1609_sig_date)
CALL write_variable_in_Case_Note(" Date DHS-1609 was received by the county: " & dhs1609_rcvd_date)
CALL write_variable_in_Case_Note(" Date DHS-1609 was completed by the county: " & dhs1609_done_date)
CALL write_variable_in_Case_Note("----")
CALL write_variable_in_Case_Note(worker_signature)
PF3

script_end_procedure_with_error_report("The case note has been created please be sure to send verifications to ECF and submit a TSS BENE request (webform) through SIR.")
