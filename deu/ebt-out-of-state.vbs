'GATHERING STATS===========================================================================================
name_of_script = "NOTES - DEU-EBT OUT OF STATE.vbs"
start_time = timer
STATS_counter = 1
STATS_manualtime = 0
STATS_denominatinon = "C"
'END OF STATS BLOCK===========================================================================================

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
		FuncLib_URL = "C:\BZS-FuncLib\MASTER FUNCTIONS LIBRARY.vbs"
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
CALL changelog_update("10/06/2022", "Update to remove hard coded DEU signature all DEU scripts.", "MiKayla Handley, Hennepin County") '#316
CALL changelog_update("09/19/2022", "Update to ensure Worker Signature is in all scripts that CASE/NOTE.", "MiKayla Handley, Hennepin County") '#316
call changelog_update("02/22/2021", "Removed handling for other option.", "MiKayla Handley, Hennepin County")
call changelog_update("01/14/2021", "Updated handling for review case to update for overpayment at a later date.", "MiKayla Handley, Hennepin County")
call changelog_update("11/30/2018", "Initial version.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display

'END CHANGELOG BLOCK =======================================================================================================----------
'connecting to BlueZone and grabbing the case number
EMConnect ""
CALL MAXIS_case_number_finder (MAXIS_case_number)
date_due = dateadd("d", 10, date)

'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 291, 85, "EBT OUT OF STATE "
  EditBox 65, 5, 40, 15, maxis_case_number
  EditBox 200, 5, 50, 15, bene_date
  EditBox 65, 25, 40, 15, MEMB_number
  EditBox 200, 25, 50, 15, out_of_state
  DropListBox 65, 45, 65, 15, "Select One:"+chr(9)+"Active"+chr(9)+"Inactive", case_status
  DropListBox 200, 45, 85, 15, "Select One:"+chr(9)+"Initial review"+chr(9)+"Client responds to request"+chr(9)+"No response received", action_taken
  EditBox 65, 65, 100, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 180, 65, 50, 15
    CancelButton 235, 65, 50, 15
  Text 5, 10, 50, 10, "Case number:"
  Text 115, 10, 80, 10, "Date accessing benefits:"
  Text 5, 30, 50, 10, "MEMB number:"
  Text 170, 30, 30, 10, "State(s):"
  Text 5, 50, 45, 10, "Case status:"
  Text 155, 50, 45, 10, "Action taken:"
  Text 5, 70, 60, 10, "Worker signature:"
EndDialog

DO
	DO
		err_msg = ""
		Dialog Dialog1
		cancel_without_confirmation
		If MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbNewLine & "* Enter a valid case number."
		IF Isdate(bene_date) = false THEN err_msg = err_msg & vbNewLine & "* Please enter the date the client was accessing the benefits."
		If out_of_state = ""  then err_msg = err_msg & vbNewLine & "* Enter the state(s) that the client has used benefits in."
		IF worker_signature = "" THEN err_msg = err_msg & vbCr & "* Please sign your case note."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP UNTIL err_msg = ""
	CALL check_for_password(are_we_passworded_out)
LOOP UNTIL are_we_passworded_out = false

IF action_taken = "Initial review" THEN
    '-------------------------------------------------------------------------------------------------DIALOG
    Dialog1 = "" 'Blanking out previous dialog detail
    BeginDialog Dialog1, 0, 0, 196, 115, "EBT Out of State Initial Review"
      EditBox 40, 5, 40, 15, date_due
      DropListBox 140, 5, 50, 15, "Select One:"+chr(9)+"1"+chr(9)+"2"+chr(9)+"3"+chr(9)+"4"+chr(9)+"5"+chr(9)+"6"+chr(9)+"7"+chr(9)+"8"+chr(9)+"9"+chr(9)+"10"+chr(9)+"11"+chr(9)+"12"+chr(9)+"OVER 12", months_used
      CheckBox 10, 35, 75, 10, "Request for contact", request_contact_checkbox
      CheckBox 105, 35, 70, 10, "Shelter verification", shel_verf_checkbox
      CheckBox 10, 45, 90, 10, "Authorization to release", ATR_Verf_CheckBox
      CheckBox 105, 45, 80, 10, "Other (please specify)", other_checkbox
      EditBox 50, 65, 140, 15, other_notes
      CheckBox 5, 85, 90, 10, "Contacted other state(s)", other_state_contact_checkbox
      ButtonGroup ButtonPressed
    	OkButton 105, 95, 40, 15
    	CancelButton 150, 95, 40, 15
      Text 85, 10, 50, 10, "# Months used: "
      GroupBox 5, 25, 185, 35, "Verification Requested: "
      Text 5, 70, 45, 10, "Other notes: "
      Text 5, 10, 35, 10, "Date due:"
    EndDialog

 	Do
    	Do
            err_msg = ""
    		Dialog Dialog1
    		cancel_without_confirmation
    		IF Isdate(date_due) = false THEN err_msg = err_msg & vbNewLine & "* Please enter the due date."
    		IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
        Loop until err_msg = ""
     	Call check_for_password(are_we_passworded_out)
    LOOP UNTIL check_for_password(are_we_passworded_out) = False
END IF

IF action_taken = "Client responds to request" THEN
    '-------------------------------------------------------------------------------------------------DIALOG
    Dialog1 = "" 'Blanking out previous dialog detail
    BeginDialog Dialog1, 0, 0, 196, 105, "Client Response"
      EditBox 55, 5, 40, 15, date_received
      DropListBox 155, 5, 35, 15, "Select One:"+chr(9)+"0"+chr(9)+"1"+chr(9)+"2"+chr(9)+"3"+chr(9)+"4"+chr(9)+"5"+chr(9)+"6"+chr(9)+"7"+chr(9)+"8"+chr(9)+"9"+chr(9)+"10"+chr(9)+"11"+chr(9)+"12"+chr(9)+"OVER 12", months_used
      CheckBox 10, 35, 75, 10, "Request for contact", request_contact_checkbox
      CheckBox 10, 45, 90, 10, "Authorization to release", ATR_Verf_CheckBox
      CheckBox 105, 35, 70, 10, "Shelter verification", shel_verf_checkbox
      CheckBox 105, 45, 80, 10, "Other (please specify)", other_checkbox
      EditBox 50, 65, 140, 15, other_notes
      ButtonGroup ButtonPressed
    	OkButton 105, 85, 40, 15
    	CancelButton 150, 85, 40, 15
      Text 5, 10, 50, 10, "Date received:"
      Text 100, 10, 55, 10, "# Months used: "
      GroupBox 5, 25, 185, 35, "Verification received: "
      Text 5, 70, 45, 10, "Other notes: "
    EndDialog

	Do
    	Do
            err_msg = ""
    		Dialog Dialog1
    		cancel_without_confirmation
    		IF Isdate(date_received) = false THEN err_msg = err_msg & vbNewLine & "* Please enter the date received."
    		IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
        Loop until err_msg = ""
     	Call check_for_password(are_we_passworded_out)
    LOOP UNTIL check_for_password(are_we_passworded_out) = False
END IF

IF action_taken = "No response received" THEN
    '-------------------------------------------------------------------------------------------------DIALOG
    Dialog1 = "" 'Blanking out previous dialog detail
    BeginDialog Dialog1, 0, 0, 196, 125, "No response received"
      EditBox 50, 5, 40, 15, date_closed
      DropListBox 150, 5, 40, 15, "Select One:"+chr(9)+"1"+chr(9)+"2"+chr(9)+"3"+chr(9)+"4"+chr(9)+"5"+chr(9)+"6"+chr(9)+"7"+chr(9)+"8"+chr(9)+"9"+chr(9)+"10"+chr(9)+"11"+chr(9)+"12"+chr(9)+"YEAR plus", months_used
      CheckBox 10, 35, 75, 10, "Request for contact", request_contact_checkbox
      CheckBox 105, 35, 70, 10, "Shelter verification", shel_verf_checkbox
      CheckBox 10, 45, 90, 10, "Authorization to release", ATR_Verf_CheckBox
      CheckBox 105, 45, 80, 10, "Other (please specify)", other_checkbox
      EditBox 65, 65, 125, 15, reason_closed
      CheckBox 5, 85, 180, 10, "Overpayment possible to be reviewed at a later date", overpayment_checkbox
      CheckBox 5, 100, 90, 10, "Contacted other state(s)", other_state_contact_checkbox
      ButtonGroup ButtonPressed
        OkButton 105, 105, 40, 15
        CancelButton 150, 105, 40, 15
      Text 95, 10, 50, 10, "# Months used: "
      GroupBox 5, 25, 185, 35, "Verification Requested: "
      Text 5, 70, 55, 10, "Closure reason:"
      Text 5, 10, 45, 10, "Date closed:"
    EndDialog

	Do
    	Do
            err_msg = ""
    		Dialog Dialog1
    		cancel_without_confirmation
    		IF Isdate(date_closed) = false THEN err_msg = err_msg & vbNewLine & "* Please enter the closed date."
			IF reason_closed = "" THEN err_msg = err_msg & vbNewLine & "* Please enter the closure reason."
    		IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
        Loop until err_msg = ""
     	Call check_for_password(are_we_passworded_out)
    LOOP UNTIL check_for_password(are_we_passworded_out) = False
END IF

IF request_contact_checkbox = CHECKED THEN pending_verifs = pending_verifs & "Contact Request, "
IF shel_verf_checkbox = CHECKED THEN pending_verifs = pending_verifs & "Shelter Verification, "
IF ATR_Verf_CheckBox = CHECKED THEN pending_verifs = pending_verifs & "ATR, "
IF other_checkbox = CHECKED THEN pending_verifs = pending_verifs & "Other, "
out_of_state = ucase(out_of_state)
'-------------------------------------------------------------------trims excess spaces of pending_verifs
pending_verifs = trim(pending_verifs) 	'takes the last comma off of pending_verifs when autofilled into dialog if more than one app date is found and additional app is selected
IF right(pending_verifs, 1) = "," THEN pending_verifs = left(pending_verifs, len(pending_verifs) - 1)

start_a_blank_case_note      'navigates to case/note and puts case/note into edit mode
	IF action_taken = "Initial review" THEN Call write_variable_in_CASE_NOTE("-----EBT OUT OF STATE REVIEWED FOR M" & MEMB_number & "-----")
	IF action_taken = "Client responds to request" THEN Call write_variable_in_CASE_NOTE("-----EBT OUT OF STATE RESPONSE RECEIVED FOR M" & MEMB_number & "-----")
	IF action_taken = "No response received" THEN Call write_variable_in_CASE_NOTE("-----EBT OUT OF STATE REQUESTED NO REPONSE RECEIVED FOR M" & MEMB_number & "-----")
    Call write_bullet_and_variable_in_CASE_NOTE("Client has been accessing benefits out of state since", bene_date)
	Call write_bullet_and_variable_in_CASE_NOTE("State(s)", out_of_state)
	IF case_status = "Inactive" THEN
		Call write_variable_in_CASE_NOTE("* Client will need to verify MN residence when reapplying and a written statement explain why accessing benefits out of state.")
		Call write_variable_in_CASE_NOTE("* Agency will need to verify benefits received in the other state prior to reopening case.")
	END IF
	IF other_state_contact_checkbox = CHECKED THEN Call write_variable_in_CASE_NOTE("* Other state(s) have been contacted.")
	IF other_state_contact_checkbox = UNCHECKED THEN Call write_variable_in_CASE_NOTE("* Other state(s) have not been contacted.")
	Call write_variable_in_CASE_NOTE("* Request sent to client for explanation of benefits used in the other state and shelter request.")
	IF action_taken = "No response received" THEN Call write_variable_in_CASE_NOTE("* Client will need to verify residence when reapplying.")
	Call write_bullet_and_variable_in_CASE_NOTE("Date case was closed", date_closed)
	Call write_bullet_and_variable_in_CASE_NOTE("Explanation of action to close the case", reason_closed)
	IF overpayment_checkbox = CHECKED THEN Call write_variable_in_CASE_NOTE("* DEU will review for possible overpayment regarding out of state usage at a later date.")
	IF action_taken = "Client responds to request" THEN Call write_bullet_and_variable_in_CASE_NOTE("Verification received", pending_verifs)
	IF action_taken <> "Client responds to request" THEN CALL write_bullet_and_variable_in_CASE_NOTE("Verification requested", pending_verifs)
	IF action_taken = "Initial review" THEN
		Call write_bullet_and_variable_in_CASE_NOTE("Verification due", date_due)
		CALL write_variable_in_CASE_NOTE ("* Client must be provided 10 days to return requested verifications.")
	END IF
	CALL write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
	CALL write_variable_in_CASE_NOTE("----- ----- ----- ----- -----")
    CALL write_variable_in_CASE_NOTE(worker_signature)
script_end_procedure_with_error_report("EBT out of state case note complete.")
