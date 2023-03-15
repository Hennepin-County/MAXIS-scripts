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

'Initial Dialog Box
'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 331, 265, "Replacing Food Destroyed in a Disaster"
  EditBox 65, 25, 35, 15, MAXIS_case_number
  EditBox 155, 25, 45, 15, loss_date
  EditBox 280, 25, 45, 15, amount_loss
  EditBox 65, 45, 135, 15, disaster_type
  EditBox 110, 65, 60, 15, how_verif
  EditBox 110, 85, 45, 15, report_date
  DropListBox 80, 105, 120, 15, "Select One:"+chr(9)+"Pending Complete DHS-1609"+chr(9)+"Pending Verification(s)"+chr(9)+"Submitted Request to TSS BENE", replacement_status
  DropListBox 275, 105, 50, 15, "Select One:"+chr(9)+"YES"+chr(9)+"NO"+chr(9)+"NA", rei_replacement
  EditBox 80, 125, 245, 15, verif_needed
  EditBox 80, 145, 245, 15, denial_reason
  EditBox 135, 175, 45, 15, dhs1609_sig_date
  EditBox 135, 195, 45, 15, dhs1609_rcvd_date
  EditBox 135, 215, 45, 15, dhs1609_done_date
  CheckBox 210, 175, 115, 10, "Request was sent to TSS BENE", TSS_BENE_sent_checkbox
  EditBox 75, 240, 165, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 250, 240, 35, 15
    CancelButton 290, 240, 35, 15
  Text 105, 30, 45, 10, "Date of Loss:"
  Text 205, 30, 75, 10, "Amount of Food Loss: "
  Text 5, 50, 60, 10, "Type of Disaster:"
  Text 205, 50, 115, 10, "(specify the cause of power outage)"
  Text 175, 70, 150, 10, "(news reports, Social Worker, Red Cross etc.)"
  Text 5, 110, 60, 10, "Approval Status: "
  Text 10, 180, 120, 10, "Date DHS-1609 signed by the client:"
  Text 10, 245, 60, 10, "Worker Signature: "
  GroupBox 5, 165, 190, 70, "Nonreceipt/Replacement Affidavit (DHS-1609)"
  Text 215, 110, 55, 10, "Replace as REI: "
  Text 10, 200, 115, 10, "Date DHS-1609 received by county: "
  Text 5, 150, 65, 10, "Reason for Denial: "
  Text 5, 90, 100, 10, "Date client reported to county:"
  Text 5, 130, 70, 10, "Verifications Needed: "
  Text 5, 70, 105, 10, "How the disaster was verified:"
  Text 5, 15, 125, 10, "(see CM0024.06.03.15 or TE02.11.18)"
  Text 5, 5, 265, 10, "When a client reports food destroyed in a disaster and all requirements are met"
  Text 5, 30, 50, 10, "Case number:"
  Text 10, 220, 125, 10, "Date DHS-1609 completed by county: "
EndDialog

'Info to the user of what this script currently covers
Do
    Do
	    err_msg = ""
	    DIALOG dialog1
	    Cancel_confirmation
	    IF MAXIS_case_number = "" OR (MAXIS_case_number <> "" AND len(MAXIS_case_number) > 8) OR (MAXIS_case_number <> "" AND IsNumeric(MAXIS_case_number) = False) THEN err_msg = err_msg & vbCr & "* Please enter a valid case number."
	    IF rei_replacement = "Select One:" THEN err_msg = err_msg & vbCr & "* Please select if the replacement was REI."
		IF replacement_status = "Select One:" THEN err_msg = err_msg & vbCr & "* Please select the status of the replacement."
		IF replacement_status = "Pending Verification(s)" and verif_needed = "" THEN err_msg = err_msg & vbCr & "* Please complete the pending verifications field"
		If IsDate(loss_date) <> TRUE or loss_date = "" Then err_msg = err_msg & vbCr & "* Please complete the date the client reports the loss occured."
		If IsDate(report_date) <> TRUE or report_date = "" Then err_msg = err_msg & vbCr & "* Please complete the date the client reported the loss of food to county."
		If amount_loss = "" Then err_msg = err_msg & vbCr & "* Please complete the amount the client reported."
		If disaster_type = "" Then err_msg = err_msg & vbCr & "* Please complete the type of disaster if power outage - specify what caused the power outage."
		If how_verif = "" Then err_msg = err_msg & vbCr & "* Please complete how the disaster was verified - news reports, Social Worker, RedCross, Excel confirmation etc."
		IF replacement_status <> "Pending Complete DHS-1609" and IsDate(dhs1609_done_date) <> TRUE or dhs1609_done_date = "" Then err_msg = err_msg & vbCr & "* Please complete the date the date the county signed the form or enter N/A."
		IF replacement_status <> "Pending Complete DHS-1609" and IsDate(dhs1609_rcvd_date) <> TRUE or dhs1609_rcvd_date = "" Then err_msg = err_msg & vbCr & "* Please complete the date county received the request or write N/A."
		IF replacement_status <> "Pending Complete DHS-1609" and IsDate(dhs1609_sig_date) <> TRUE or dhs1609_sig_date = "" Then err_msg = err_msg & vbCr & "* Please complete the date the client signed the form or write N/A."
    	IF TSS_BENE_sent_checkbox = UNCHECKED and replacement_status = "Submitted Request to TSS BENE" THEN err_msg = err_msg & vbCr & "* Please check that the TSS BENE Webform has been completed."
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
