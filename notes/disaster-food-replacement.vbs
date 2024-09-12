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
CALL changelog_update("09/10/2024", "Update to align with 06/2024 POLI/TEMP regarding denial of replacement requests.", "Mark Riegel, Hennepin County") '#1848
CALL changelog_update("05/07/2024", "Update to align with updated 02/2024 POLI/TEMP.", "Mark Riegel, Hennepin County") '#1796
CALL changelog_update("09/19/2022", "Update to ensure Worker Signature is in all scripts that CASE/NOTE.", "MiKayla Handley, Hennepin County") '#316
call changelog_update("11/27/2019", "Initial version.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

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

'Dialog to gather details on the request
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 336, 385, "Replacing Food Destroyed in a Disaster"
  EditBox 65, 25, 35, 15, MAXIS_case_number
  EditBox 110, 55, 45, 15, loss_date
  EditBox 110, 70, 45, 15, amount_loss
  EditBox 110, 85, 45, 15, report_date
  EditBox 110, 100, 210, 15, disaster_description
  EditBox 110, 115, 120, 15, how_verif
  EditBox 110, 130, 45, 15, loss_verification_date
  DropListBox 90, 170, 180, 15, "Select One:"+chr(9)+"Pending Complete DHS-1609"+chr(9)+"Pending Verification(s)"+chr(9)+"Request Approved"+chr(9)+"Request Denied", replacement_status
  DropListBox 90, 185, 55, 15, "Select One:"+chr(9)+"YES"+chr(9)+"NO"+chr(9)+"NA", rei_replacement
  DropListBox 90, 200, 180, 15, "Select One:"+chr(9)+"DHS-1609 never received"+chr(9)+"Late report of loss"+chr(9)+"Other denial reason", denial_reason
  EditBox 90, 220, 235, 15, denial_reason_other
  EditBox 90, 235, 235, 15, verif_needed
  CheckBox 10, 260, 135, 10, "Request was sent to TSS BENE Unit", TSS_BENE_sent_checkbox
  ButtonGroup ButtonPressed
    PushButton 145, 255, 180, 15, "TSS BENE Unit Webform", TSS_BENE_webform_btn
  EditBox 135, 295, 45, 15, dhs_1609_due_date
  CheckBox 190, 295, 95, 10, "Create TIKL for DHS-1609", dhs_1609_tikl
  EditBox 135, 310, 45, 15, dhs1609_sig_date
  EditBox 135, 325, 45, 15, dhs1609_rcvd_date
  EditBox 135, 340, 45, 15, dhs1609_done_date
  EditBox 75, 365, 165, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 250, 365, 35, 15
    CancelButton 290, 365, 35, 15
  Text 5, 5, 265, 10, "When a client reports food destroyed in a disaster and all requirements are met"
  Text 5, 15, 125, 10, "(see CM0024.06.03.15 or TE02.11.18)"
  Text 10, 30, 50, 10, "Case number:"
  GroupBox 5, 45, 325, 105, "Food Loss Details"
  Text 10, 60, 45, 10, "Date of Loss:"
  Text 10, 75, 80, 10, "Amount of Food Loss ($): "
  Text 10, 90, 95, 10, "Date loss reported to county:"
  Text 10, 105, 75, 10, "Describe the disaster:"
  Text 10, 120, 90, 10, "How disaster was verified:"
  Text 240, 120, 80, 20, "(news report, social worker, Red Cross, etc.)"
  Text 10, 135, 65, 10, "Date Loss verified:"
  GroupBox 5, 155, 325, 120, "Information for Replacement Request"
  Text 10, 170, 70, 10, "Status of Request: "
  Text 10, 185, 55, 10, "Replace as REI: "
  Text 10, 205, 65, 10, "Reason for Denial: "
  Text 10, 225, 75, 10, "Denial Reason (other):"
  Text 10, 240, 70, 10, "Verifications Needed: "
  GroupBox 5, 280, 325, 80, "Nonreceipt/Replacement Affidavit (DHS-1609)"
  Text 10, 300, 95, 10, "Date DHS-1609 is due back:"
  Text 10, 315, 120, 10, "Date DHS-1609 signed by the client:"
  Text 10, 330, 115, 10, "Date DHS-1609 received by county: "
  Text 10, 345, 125, 10, "Date DHS-1609 completed by county: "
  Text 10, 370, 60, 10, "Worker Signature: "
EndDialog

'Info to the user of what this script currently covers
Do
  Do
    err_msg = ""
    DIALOG dialog1
    Cancel_confirmation
    Call validate_MAXIS_case_number(err_msg, "*")
    If IsDate(loss_date) <> TRUE or loss_date = "" Then err_msg = err_msg & vbCr & "* Please enter the date the client reports the loss occurred."
    If amount_loss = "" Then err_msg = err_msg & vbCr & "* Please enter the dollar amount the client reported."
    If IsDate(report_date) <> TRUE or report_date = "" Then err_msg = err_msg & vbCr & "* Please enter the date the client reported the loss of food to county."
    If trim(disaster_description) = "" Then err_msg = err_msg & vbCr & "* Please describe the type of disaster. If it was a power outage, please specify what caused the power outage."
    If trim(how_verif) = "" Then err_msg = err_msg & vbCr & "* Please indicate how the disaster was verified - news reports, social worker, Red Cross, utility confirmation, etc."
    If IsDate(loss_verification_date) <> TRUE or loss_verification_date = "" Then err_msg = err_msg & vbCr & "* Please enter the date the county verified the loss of food."
    IF replacement_status = "Select One:" THEN err_msg = err_msg & vbCr & "* Please select the status of the replacement."
    IF replacement_status = "Pending Verification(s)" and trim(verif_needed) = "" THEN err_msg = err_msg & vbCr & "* Please complete the pending verifications field."
    IF replacement_status = "Request Denied" and denial_reason = "Select One:" THEN err_msg = err_msg & vbCr & "* Please select the reason for the denial."
    IF replacement_status <> "Request Denied" and trim(denial_reason_other) <> "" THEN err_msg = err_msg & vbCr & "* The other denial reason field should be blank if the request is not being denied."
    IF rei_replacement = "Select One:" THEN err_msg = err_msg & vbCr & "* Please select if the replacement was REI, or select NA."
    If (denial_reason = "DHS-1609 never received" or denial_reason = "Late report of loss") and trim(denial_reason_other) <> "" THEN err_msg = err_msg & vbCr & "* The Denial reason (other) field should be blank if the other denial reason is not selected."
    If denial_reason = "Other denial reason" and trim(denial_reason_other) = "" THEN err_msg = err_msg & vbCr & "* Please enter a reason for the denial in the Denial reason (other) field."
    IF TSS_BENE_sent_checkbox = UNCHECKED and replacement_status = "Request Approved" THEN err_msg = err_msg & vbCr & "* Please check that the TSS BENE Webform has been completed."
    IF TSS_BENE_sent_checkbox = CHECKED and replacement_status <> "Request Approved" THEN err_msg = err_msg & vbCr & "* Please only check the TSS BENE Webform checkbox if you are approving the request."
    If ButtonPressed = TSS_BENE_webform_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://owa.dhssir.cty.dhs.state.mn.us/csedforms/MMR/TSSBENE_BENE_request.asp"
    If IsDate(dhs_1609_due_date) <> TRUE and trim(dhs_1609_due_date) <> "" Then err_msg = err_msg & vbCr & "* Please enter the date the DHS-1609 is due back."
    IF IsDate(dhs1609_sig_date) <> TRUE and trim(dhs1609_sig_date) <> "" Then err_msg = err_msg & vbCr & "* Please enter the date the client signed the form."
    IF IsDate(dhs1609_rcvd_date) <> TRUE and trim(dhs1609_rcvd_date) <> "" Then err_msg = err_msg & vbCr & "* Please enter the date the county received the request."
    If IsDate(dhs1609_done_date) <> TRUE and trim(dhs1609_done_date) <> "" Then err_msg = err_msg & vbCr & "* Please enter the date the county signed the form."
    IF err_msg <> "" and ButtonPressed <> TSS_BENE_webform_btn THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
    IF trim(worker_signature) = "" THEN err_msg = err_msg & vbCr & "* Please sign your case note."
  LOOP UNTIL err_msg = "" and ButtonPressed <> TSS_BENE_webform_btn									'loops until all errors are resolved
  CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'Create TIKL if selected
If dhs_1609_tikl = CHECKED Then Call create_TIKL("DHS-1609 was sent 10 days ago and is now due back.", 10, date, False, TIKL_note_text)

'Write CASE/NOTE with information
start_a_blank_case_note
CALL write_variable_in_Case_Note("--Food Destroyed in Disaster Reported - " & replacement_status & "--")
CALL write_bullet_and_variable_in_Case_Note("Date of Loss", loss_date)
CALL write_bullet_and_variable_in_Case_Note("Amount of Food Loss", amount_loss)
CALL write_bullet_and_variable_in_Case_Note("Date client reported the loss of food to county", report_date)
CALL write_bullet_and_variable_in_Case_Note("Description of Disaster", disaster_description)
CALL write_bullet_and_variable_in_Case_Note("How the disaster was verified", how_verif)
CALL write_bullet_and_variable_in_Case_Note("Date Loss Verified", loss_verification_date)
CALL write_bullet_and_variable_in_Case_Note("Replace as REI", rei_replacement)
CALL write_bullet_and_variable_in_Case_Note("Status of Request", replacement_status)
IF denial_reason = "Other denial reason" Then 
  CALL write_bullet_and_variable_in_Case_Note("Reason for Denial", denial_reason_other)
ElseIf denial_reason = "DHS-1609 never received" or denial_reason = "Late report of loss" Then
  CALL write_bullet_and_variable_in_Case_Note("Reason for Denial", denial_reason)
End If
CALL write_bullet_and_variable_in_Case_Note("Verifications Requested", verif_needed)
IF TSS_BENE_sent_checkbox = CHECKED THEN CALL write_variable_in_Case_Note("* Submitted a TSS BENE request (webform) through SIR")
CALL write_variable_in_Case_Note("Nonreceipt/Replacement Affidavit (DHS-1609)")
CALL write_bullet_and_variable_in_Case_Note("Date DHS-1609 is due back from the client", dhs_1609_due_date)
If dhs_1609_tikl = CHECKED Then write_variable_in_Case_Note(" TIKL created for DHS-1609 due date.")
CALL write_bullet_and_variable_in_Case_Note("Date DHS-1609 was signed by the client", dhs1609_sig_date)
CALL write_bullet_and_variable_in_Case_Note("Date DHS-1609 was received by the county", dhs1609_rcvd_date)
CALL write_bullet_and_variable_in_Case_Note("Date DHS-1609 was completed by the county", dhs1609_done_date)
CALL write_variable_in_Case_Note("----")
CALL write_variable_in_Case_Note(worker_signature)

If denial_reason = "DHS-1609 never received" or denial_reason = "Late report of loss" Then

  'Script ends
  script_end_procedure_with_error_report("The case note has been created. You indicated that the request was denied due to the DHS-1609 never being received or due to a late report of the loss. MAXIS will automatically generate a denied replacement notice when the 03 and 04 replacement denial reasons are used. See TE02.11.07 (Replacing Benefits). These actions will need to be completed outside of the script.")

ElseIf denial_reason = "Other denial reason" Then

  'Dialog to allow for edits to CASE/NOTE before going to SPEC/MEMO
  Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 246, 75, "Food Destroyed in Disaster or Misfortune - Save CASE/NOTE"
      ButtonGroup ButtonPressed
        PushButton 125, 55, 65, 15, "Save CASE/NOTE", save_case_note_btn
        CancelButton 195, 55, 45, 15
      Text 10, 5, 230, 35, "Please review the generated CASE/NOTE and make any needed changes. Once the CASE/NOTE is finalized, the script will generate a SPEC/MEMO notifying the resident of the denial. Please click the 'Save CASE/NOTE' to navigate to SPEC/MEMO."
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

  'Start a SPEC/MEMO and add details about denial
  CALL start_a_new_spec_memo(memo_opened, search_for_arep_and_swkr, forms_to_arep, forms_to_swkr, send_to_other, other_name, other_street, other_city, other_state, other_zip, False)
  Call write_variable_in_SPEC_MEMO("Hello, on " & report_date & " you requested a replacement for food benefits due to " & disaster_description & ". Your replacement request has been denied because " & denial_reason_other & ".")

  'Script ends
  script_end_procedure_with_error_report("The SPEC/MEMO has been created.")
Else
  'Script ends
  script_end_procedure_with_error_report("The case note has been created. If request has been approved, please be sure to send verifications to ECF and submit a TSS BENE request (webform) through SIR.")
End If

'----------------------------------------------------------------------------------------------------Closing Project Documentation - Version date 01/12/2023
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------05/08/2024
'--Tab orders reviewed & confirmed----------------------------------------------09/10/2024
'--Mandatory fields all present & Reviewed--------------------------------------09/10/2024
'--All variables in dialog match mandatory fields-------------------------------05/08/2024
'Review dialog names for content and content fit in dialog----------------------05/08/2024
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------09/10/2024
'--CASE:NOTE Header doesn't look funky------------------------------------------09/10/2024
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------09/10/2024
'--write_variable_in_CASE_NOTE function: confirm that proper punctuation is used -----------------------------------
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------05/08/2024
'--MAXIS_background_check reviewed (if applicable)------------------------------05/08/2024
'--PRIV Case handling reviewed -------------------------------------------------05/08/2024
'--Out-of-County handling reviewed----------------------------------------------05/08/2024
'--script_end_procedures (w/ or w/o error messaging)----------------------------05/08/2024
'--BULK - review output of statistics and run time/count (if applicable)--------N/A
'--All strings for MAXIS entry are uppercase vs. lower case (Ex: "X")-----------05/08/2024
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------05/08/2024
'--Incrementors reviewed (if necessary)-----------------------------------------05/08/2024
'--Denomination reviewed -------------------------------------------------------05/08/2024
'--Script name reviewed---------------------------------------------------------05/08/2024
'--BULK - remove 1 incrementor at end of script reviewed------------------------N/A

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------09/10/2024
'--comment Code-----------------------------------------------------------------09/10/2024
'--Update Changelog for release/update------------------------------------------09/10/2024
'--Remove testing message boxes-------------------------------------------------09/10/2024
'--Remove testing code/unnecessary code-----------------------------------------09/10/2024
'--Review/update SharePoint instructions----------------------------------------09/10/2024
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------05/08/2024
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------05/08/2024
'--COMPLETE LIST OF SCRIPTS update policy references----------------------------05/08/2024
'--Complete misc. documentation (if applicable)---------------------------------05/08/2024
'--Update project team/issue contact (if applicable)----------------------------05/08/2024
