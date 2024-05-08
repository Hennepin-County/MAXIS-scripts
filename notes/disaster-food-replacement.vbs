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
CALL changelog_update("05/07/2024", "Update to align with updated POLI/TEMP.", "Mark Riegel, Hennepin County") '#316
CALL changelog_update("09/19/2022", "Update to ensure Worker Signature is in all scripts that CASE/NOTE.", "MiKayla Handley, Hennepin County") '#316
call changelog_update("11/27/2019", "Initial version.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================
'Connecting to BlueZone
EMConnect ""

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

'Initial Dialog Box
'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 336, 410, "Replacing Food Destroyed in a Disaster"
  EditBox 55, 30, 35, 15, MAXIS_case_number
  DropListBox 150, 30, 120, 15, "Select One:"+chr(9)+"Initial report of loss of food"+chr(9)+"Update on replacement request"+chr(9)+"Decision on replacement request", process_step
  EditBox 110, 60, 45, 15, loss_date
  EditBox 110, 80, 45, 15, amount_loss
  EditBox 110, 100, 45, 15, report_date
  EditBox 110, 120, 210, 15, disaster_description
  EditBox 110, 140, 120, 15, how_verif
  EditBox 110, 160, 45, 15, loss_verification_date
  EditBox 90, 195, 235, 15, verif_needed
  DropListBox 90, 215, 50, 15, "Select One:"+chr(9)+"YES"+chr(9)+"NO"+chr(9)+"NA", rei_replacement
  DropListBox 90, 230, 180, 15, "Select One:"+chr(9)+"Pending Complete DHS-1609"+chr(9)+"Pending Verification(s)"+chr(9)+"Request Approved"+chr(9)+"Request Denied", replacement_status
  EditBox 90, 245, 235, 15, denial_reason
  CheckBox 10, 270, 135, 10, "Request was sent to TSS BENE Unit", TSS_BENE_sent_checkbox
  ButtonGroup ButtonPressed
    PushButton 145, 265, 180, 15, "TSS BENE Unit Webform", TSS_BENE_webform_btn
  EditBox 135, 305, 45, 15, dhs_1609_due_date
  CheckBox 190, 305, 95, 10, "Create TIKL for DHS-1609", dhs_1609_tikl
  EditBox 135, 325, 45, 15, dhs1609_sig_date
  EditBox 135, 345, 45, 15, dhs1609_rcvd_date
  EditBox 135, 365, 45, 15, dhs1609_done_date
  EditBox 75, 390, 165, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 250, 390, 35, 15
    CancelButton 290, 390, 35, 15
  Text 5, 5, 265, 10, "When a client reports food destroyed in a disaster and all requirements are met"
  Text 5, 15, 125, 10, "(see CM0024.06.03.15 or TE02.11.18)"
  Text 5, 35, 50, 10, "Case number:"
  Text 100, 35, 50, 10, "Process Step:"
  GroupBox 5, 50, 325, 130, "Food Loss Details"
  Text 10, 65, 45, 10, "Date of Loss:"
  Text 10, 85, 80, 10, "Amount of Food Loss ($): "
  Text 10, 105, 95, 10, "Date loss reported to county:"
  Text 10, 125, 75, 10, "Describe the disaster:"
  Text 10, 145, 90, 10, "How disaster was verified:"
  Text 240, 140, 80, 20, "(news report, social worker, Red Cross, etc.)"
  Text 10, 165, 65, 10, "Date Loss verified:"
  Text 10, 310, 95, 10, "Date DHS-1609 is due back:"
  GroupBox 5, 185, 325, 100, "Information for Replacement Request"
  Text 10, 200, 70, 10, "Verifications Needed: "
  Text 10, 215, 55, 10, "Replace as REI: "
  Text 10, 230, 70, 10, "Decision on Request: "
  Text 10, 250, 65, 10, "Reason for Denial: "
  GroupBox 5, 290, 325, 95, "Nonreceipt/Replacement Affidavit (DHS-1609)"
  Text 10, 330, 120, 10, "Date DHS-1609 signed by the client:"
  Text 10, 350, 115, 10, "Date DHS-1609 received by county: "
  Text 10, 370, 125, 10, "Date DHS-1609 completed by county: "
  Text 10, 395, 60, 10, "Worker Signature: "
EndDialog

'Info to the user of what this script currently covers
Do
  Do
    err_msg = ""
    DIALOG dialog1
    Cancel_confirmation
    Call validate_MAXIS_case_number(err_msg, "*")
    If ButtonPressed = TSS_BENE_webform_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://owa.dhssir.cty.dhs.state.mn.us/csedforms/MMR/TSSBENE_BENE_request.asp"
    IF process_step = "Select One:" THEN err_msg = err_msg & vbCr & "* Please select the step in the process you are providing an update on."
    IF rei_replacement = "Select One:" THEN err_msg = err_msg & vbCr & "* Please select if the replacement was REI, or select NA."
    IF replacement_status = "Select One:" THEN err_msg = err_msg & vbCr & "* Please select the status of the replacement."
    IF replacement_status = "Pending Verification(s)" and verif_needed = "" THEN err_msg = err_msg & vbCr & "* Please complete the pending verifications field."
    If IsDate(loss_date) <> TRUE or loss_date = "" Then err_msg = err_msg & vbCr & "* Please enter the date the client reports the loss occurred."
    If IsDate(report_date) <> TRUE or report_date = "" Then err_msg = err_msg & vbCr & "* Please enter the date the client reported the loss of food to county."
    If amount_loss = "" Then err_msg = err_msg & vbCr & "* Please enter the dollar amount the client reported."
    If disaster_description = "" Then err_msg = err_msg & vbCr & "* Please describe the type of disaster. If it was a power outage, please specify what caused the power outage."
    If how_verif = "" Then err_msg = err_msg & vbCr & "* Please indicate how the disaster was verified - news reports, social worker, Red Cross, utility confirmation, etc."
    IF replacement_status <> "Pending Complete DHS-1609" Then
      If IsDate(dhs_1609_due_date) <> TRUE or dhs_1609_due_date = "" Then err_msg = err_msg & vbCr & "* Please enter the date the DHS-1609 is due back."
      If IsDate(dhs1609_done_date) <> TRUE or dhs1609_done_date = "" Then err_msg = err_msg & vbCr & "* Please enter the date the county signed the form."
      IF IsDate(dhs1609_rcvd_date) <> TRUE or dhs1609_rcvd_date = "" Then err_msg = err_msg & vbCr & "* Please enter the date the county received the request."
      IF IsDate(dhs1609_sig_date) <> TRUE or dhs1609_sig_date = "" Then err_msg = err_msg & vbCr & "* Please enter the date the client signed the form."
    End If
    IF replacement_status = "Request Denied" and trim(denial_reason) = "" THEN err_msg = err_msg & vbCr & "* Please provide a reason for the denial."
    IF replacement_status <> "Request Denied" and denial_reason <> "" THEN err_msg = err_msg & vbCr & "* The reason for the denial field should be blank if the request is not being denied."
    IF TSS_BENE_sent_checkbox = UNCHECKED and replacement_status = "Request Approved" THEN err_msg = err_msg & vbCr & "* Please check that the TSS BENE Webform has been completed."
    IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
    IF worker_signature = "" THEN err_msg = err_msg & vbCr & "* Please sign your case note."
  LOOP UNTIL err_msg = ""									'loops until all errors are resolved
  CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'Create TIKL if selected
If dhs_1609_tikl = CHECKED Then Call create_TIKL("DHS-1609 was sent 10 days ago and is now due back.", 10, date, False, TIKL_note_text)

'Write CASE/NOTE with information
start_a_blank_case_note
CALL write_variable_in_Case_Note("--Food Destroyed in a Disaster Reported - " & replacement_status & "--")
CALL write_bullet_and_variable_in_Case_Note("Current Process Step", process_step)
CALL write_bullet_and_variable_in_Case_Note("Date of Loss", loss_date)
CALL write_bullet_and_variable_in_Case_Note("Amount of Food Loss", amount_loss)
CALL write_bullet_and_variable_in_Case_Note("Date client reported the loss of food to county", report_date)
CALL write_bullet_and_variable_in_Case_Note("Description of Disaster", disaster_description)
CALL write_bullet_and_variable_in_Case_Note("How the disaster was verified", how_verif)
CALL write_bullet_and_variable_in_Case_Note("Date Loss Verified", loss_verification_date)
CALL write_bullet_and_variable_in_Case_Note("Replace as REI", rei_replacement)
IF TSS_BENE_sent_checkbox <> CHECKED THEN CALL write_bullet_and_variable_in_Case_Note("Status of Request", replacement_status)
IF replacement_status = "Request Denied" Then CALL write_bullet_and_variable_in_Case_Note("Reason for Denial", denial_reason)
CALL write_bullet_and_variable_in_Case_Note("Verifications Requested", verif_needed)
IF TSS_BENE_sent_checkbox = CHECKED THEN CALL write_variable_in_Case_Note("* Submitted a TSS BENE request (webform) through SIR")
CALL write_variable_in_Case_Note("Nonreceipt/Replacement Affidavit (DHS-1609)")
CALL write_variable_in_Case_Note(" Date DHS-1609 is due back from the client: " & dhs_1609_due_date)
CALL write_variable_in_Case_Note(" Date DHS-1609 was signed by the client: " & dhs1609_sig_date)
CALL write_variable_in_Case_Note(" Date DHS-1609 was received by the county: " & dhs1609_rcvd_date)
CALL write_variable_in_Case_Note(" Date DHS-1609 was completed by the county: " & dhs1609_done_date)
CALL write_variable_in_Case_Note("----")
CALL write_variable_in_Case_Note(worker_signature)
PF3

script_end_procedure_with_error_report("The case note has been created. If request has been approved, please be sure to send verifications to ECF and submit a TSS BENE request (webform) through SIR.")
