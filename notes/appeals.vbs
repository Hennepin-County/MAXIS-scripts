'Required for statistical purposes========================================================================================== 
name_of_script = "NOTES - APPEALS.vbs"
start_time = timer 
STATS_counter = 1               'sets the stats counter at one 
STATS_manualtime = 95           'manual run time in seconds 
STATS_denomination = "C"        'C is for each case 
 'END OF stats block========================================================================================================= 

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN		'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else																		'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		End if
		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
		req.send													'Sends request
		IF req.Status = 200 THEN									'200 means great success
			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
			Execute req.responseText								'Executes the script code
		ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
			MsgBox 	"Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_
					vbCr & _
					"Before contacting Veronica Cary, please check to make sure you can load the main page at www.GitHub.com." & vbCr &_
					vbCr & _
					"If you can reach GitHub.com, but this script still does not work, ask an alpha user to contact Veronica Cary and provide the following information:" & vbCr &_
					vbTab & "- The name of the script you are running." & vbCr &_
					vbTab & "- Whether or not the script is ""erroring out"" for any other users." & vbCr &_
					vbTab & "- The name and email for an employee from your IT department," & vbCr & _
					vbTab & vbTab & "responsible for network issues." & vbCr &_
					vbTab & "- The URL indicated below (a screenshot should suffice)." & vbCr &_
					vbCr & _
					"Veronica will work with your IT department to try and solve this issue, if needed." & vbCr &_
					vbCr &_
					"URL: " & FuncLib_URL
					script_end_procedure("Script ended due to error connecting to GitHub.")
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
call changelog_update("12/15/2016", "New script that will document the information about an appeal, and the appeal process.", "Charles Clark, Hennepin County")
call changelog_update("12/12/2016", "Initial version.", "Charles Clark, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'The script----------------------------------------------------------------------------------------------------
'connecting to BlueZone and grabbing the case number
EMConnect ""
Call MAXIS_case_number_finder(maxis_case_number)

'Initial dialog and do...loop
BeginDialog, 0, 0, 201, 70, "Appeal initial dialog"
  EditBox 135, 5, 60, 15, maxis_case_number
  DropListBox 105, 25, 90, 15, "Select one..."+chr(9)+"Appeal Summary Completed"+chr(9)+"Appeal Hearing Info"+chr(9)+"Appeal Decision Received"+chr(9)+"Appeal Resolution", appeal_actions
  ButtonGroup ButtonPressed
    OkButton 90, 45, 50, 15
    CancelButton 145, 45, 50, 15
  Text 5, 30, 80, 10, "Select an appeal action:"
  Text 10, 10, 45, 10, "Case number:"
EndDialog
Do
	Do
		Dialog
		if ButtonPressed = 0 then StopScript
		if IsNumeric(maxis_case_number) = false or len(maxis_case_number) > 8 then MsgBox "* Please enter a valid case number."
		If appeal_actions = "Select one..." then MsgBox "Please select an appeal action."
	Loop until appeal_actions <> "Select one..." and IsNumeric(maxis_case_number) = true
 Call check_for_password(are_we_passworded_out)
LOOP UNTIL check_for_password(are_we_passworded_out) = False
	
If appeal_actions = "Appeal Summary Completed" then
    BeginDialog, 0, 0, 351, 195, "Appeal Summary Completed"
      EditBox 105, 10, 50, 15, date_appeal_received
      EditBox 295, 10, 50, 15, effective_date
      EditBox 95, 35, 250, 15, action_client_is_appealing
      CheckBox 100, 60, 30, 10, "CASH", cash_check
      CheckBox 135, 60, 30, 10, "SNAP", snap_check
      CheckBox 170, 60, 30, 10, "HC", hc_check
      DropListBox 160, 75, 55, 15, "Select one..."+chr(9)+"Yes"+chr(9)+"No", benefits_continuing
      EditBox 80, 95, 265, 15, proofs_attachments
      EditBox 80, 120, 265, 15, other_notes
      EditBox 80, 145, 265, 15, action_taken
      EditBox 145, 170, 85, 15, worker_signature
      ButtonGroup ButtonPressed
    	OkButton 240, 170, 50, 15
    	CancelButton 295, 170, 50, 15
      Text 5, 40, 85, 10, "Action client is appealing:"
      Text 75, 175, 65, 10, "Worker Signature:"
      Text 5, 15, 100, 10, "Date appeal request received:"
      Text 5, 125, 45, 10, "Other notes:"
      Text 5, 80, 150, 10, "Benefits continuing at pre-appeal level (Y/N):"
      Text 5, 150, 50, 10, " Actions taken:"
      Text 5, 60, 90, 10, "Programs client appealing:"
      Text 165, 15, 130, 10, "Effective date of action being appealed:"
      Text 5, 100, 70, 10, "Proofs/attachments:"
    EndDialog
		'Shows dialog and creates and displays an error message if worker completes things incorrectly.  
	DO		
		Do 
			err_msg = "" 
			Dialog
			cancel_confirmation 
			IF isdate(date_appeal_received) = false THEN err_msg = err_msg & vbNewLine & "* Please complete Date Appeal Request Received"
			IF isdate(effective_date) = false THEN err_msg = err_msg & vbNewLine & "* Please enter a valid numeric date for the action the client wishes to appeal" 
			IF action_client_is_appealing = "" THEN err_msg = err_msg & vbNewLine & "* Please enter the action client is appealing"
			IF (cash_check <> 1 AND snap_check <> 1 AND hc_check <> 1) THEN err_msg = err_msg & vbNewLine & "* Please select programs client is appealing"
			IF benefits_continuing = "Select one..." THEN err_msg = err_msg & vbNewLine & "* Please select if benefits will continue pending the outcome of the appeal."
			IF worker_signature = "" THEN err_msg = err_msg & vbNewLine & "* Please enter your worker signature."
			IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine		
		Loop until err_msg = ""	
		Call check_for_password(are_we_passworded_out)
	LOOP UNTIL check_for_password(are_we_passworded_out) = False
	 		
	'creating new variable for case note for programs appealing that is incremential	
	If cash_check = 1 then progs_appealing = progs_appealing & "CASH, "
	If snap_check = checked then progs_appealing = progs_appealing & "SNAP, " 
	IF hc_check = checked then progs_appealing = progs_appealing & "HC, "
	'trims excess spaces of progs_appealing
	progs_appealing = trim(progs_appealing)
	'takes the last comma off of progs_appealing when autofilled into dialog if more more than one app date is found and additional app is selected
	If right(progs_appealing, 1) = "," THEN progs_appealing = left(progs_appealing, len(progs_appealing) - 1) 
	
	 start_a_blank_case_note      'navigates to case/note and puts case/note into edit mode		 
	 Call write_variable_in_CASE_NOTE("---Appeal Summary Completed---")
	 call write_bullet_and_variable_in_CASE_NOTE("Date appeal request received", date_appeal_received)
	 Call write_bullet_and_variable_in_CASE_NOTE("Effective date of action being appealed", effective_date)
	 Call write_bullet_and_variable_in_CASE_NOTE("Action client is appealing", action_client_is_appealing)
	 Call write_bullet_and_variable_in_CASE_NOTE("Programs client appealing", progs_appealing)
	 Call write_bullet_and_variable_in_CASE_NOTE("Benefits continuing at pre-appeal level (Y/N)", benefits_continuing)
	 Call write_bullet_and_variable_in_CASE_NOTE("Proofs/attachments", proofs_attachments)
	 Call write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes) 
	 Call write_bullet_and_variable_in_CASE_NOTE("Actions taken", action_taken)
	 Call write_variable_in_CASE_NOTE ("---")
	 call write_variable_in_CASE_NOTE(worker_signature)	 
END If 	
	
If appeal_actions = "Appeal Hearing Info" then
    BeginDialog, 0, 0, 346, 140, "Appeal Hearing Info"
      EditBox 65, 5, 55, 15, hearing_date
      DropListBox 190, 5, 60, 15, "Select one..."+chr(9)+"Yes, in person"+chr(9)+"Yes, by phone"+chr(9)+"Did not attend", appeal_attendence	
      EditBox 65, 30, 265, 15, hearing_details
      EditBox 65, 50, 265, 15, other_notes
      EditBox 105, 80, 55, 15, anticipated_date_result
      EditBox 105, 105, 115, 15, worker_signature
      ButtonGroup ButtonPressed
    	OkButton 225, 105, 50, 15
    	CancelButton 280, 105, 50, 15
      Text 5, 10, 60, 10, "Date Of Hearing:"
      Text 15, 85, 85, 10, "Anticipated decision date:"
      Text 135, 10, 55, 10, "Client attended:"
      Text 5, 35, 55, 10, "Hearing Details:"
      Text 45, 110, 60, 10, "Worker Signature:"
      Text 20, 55, 40, 10, "Other notes:"
    EndDialog
	'Shows dialog and creates and displays an error message if worker completes things incorrectly.  
	DO
		Do 
			err_msg = "" 
			Dialog	
			cancel_confirmation 	
			IF isdate(hearing_date) = false THEN err_msg = err_msg & vbNewLine & "* Please complete date of hearing."
			If appeal_attendence = "Select one..." THEN err_msg = err_msg & vbNewLine & "* Please select if the client attended appeal, or if appeal was held by phone"
			IF hearing_details = "" THEN err_msg = err_msg & vbNewLine & "* Please enter hearing details"
			IF isdate(anticipated_date_result) = false THEN err_msg = err_msg & vbNewLine & "* Please enter a valid date for the  anticipated date of appeal decision" 
			IF worker_signature = "" THEN err_msg = err_msg & vbNewLine & "* Please enter your worker signature." 
			IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine		 
		Loop until err_msg = ""	
		Call check_for_password(are_we_passworded_out)
	LOOP UNTIL check_for_password(are_we_passworded_out) = False  
 
    start_a_blank_case_note      'navigates to case/note and puts case/note into edit mode 
 	Call write_variable_in_CASE_NOTE("---Appeal Hearing Info---")
	Call write_bullet_and_variable_in_CASE_NOTE("Date Of Hearing", hearing_date)
	Call write_bullet_and_variable_in_CASE_NOTE("Did Client Attend The Appeal", appeal_attendence) 
	Call write_bullet_and_variable_in_CASE_NOTE("Hearing Details", hearing_details)
	Call write_bullet_and_variable_in_CASE_NOTE("Aniticipated date of decision", anticipated_date_result)	
	Call write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
	Call write_variable_in_CASE_NOTE ("---")
	Call write_variable_in_CASE_NOTE(worker_signature)
End If

If appeal_actions = "Appeal Decision Received" then
    BeginDialog, 0, 0, 346, 120, "Appeal decision received"
      EditBox 85, 10, 245, 15, disposition_of_appeal
      EditBox 85, 35, 245, 15, actions_needed
      EditBox 85, 60, 60, 15, date_signed_by_judge
      DropListBox 275, 60, 55, 15, "Select one..."+chr(9)+"Yes"+chr(9)+"No"+chr(9)+"NA", compliance_form_needed
      EditBox 85, 90, 135, 15, worker_signature
      ButtonGroup ButtonPressed
    	OkButton 225, 90, 50, 15
    	CancelButton 280, 90, 50, 15
      Text 20, 95, 60, 10, "Worker Signature:"
      Text 5, 15, 75, 10, "Disposition of appeal:"
      Text 10, 65, 75, 10, "Date signed by judge:"
      Text 155, 65, 115, 10, "SNAP compliance form completed:"
      Text 25, 40, 55, 10, "Actions needed:"
    EndDialog
'Shows dialog and creates and displays an error message if worker completes things incorrectly.  
	Do
		Do 
			err_msg = "" 
			Dialog
			cancel_confirmation 
			IF disposition_of_appeal = "" THEN err_msg = err_msg & vbNewLine & "* Please enter the disposition of the appeal"
			IF actions_needed = "" THEN err_msg = err_msg & vbNewLine & "* Please enter actions needed"
			If compliance_form_needed = "Select one..." THEN err_msg = err_msg & vbNewLine & "* Please select whether a compliance form is needed"
			IF isdate(date_signed_by_judge) = false THEN err_msg = err_msg & vbNewLine & "* Please enter a valid date the appeal findings were signed by the Judge" 
			IF worker_signature = "" THEN err_msg = err_msg & vbNewLine & "* Please enter your worker signature." 
			IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
		Loop until err_msg = ""	
		Call check_for_password(are_we_passworded_out)
	LOOP UNTIL check_for_password(are_we_passworded_out) = False	 

	 start_a_blank_case_note      'navigates to case/note and puts case/note into edit mode
	 Call write_variable_in_CASE_NOTE("Appeal Decision Received")
	 Call write_bullet_and_variable_in_CASE_NOTE("Disposition of appeal", disposition_of_appeal)
	 Call write_bullet_and_variable_in_CASE_NOTE("Actions needed", actions_needed)
	 Call write_bullet_and_variable_in_CASE_NOTE("SNAP compliance form completed", compliance_form_needed)
	 Call write_bullet_and_variable_in_CASE_NOTE("Date signed by judge", date_signed_by_judge)
	 Call write_variable_in_CASE_NOTE ("---")
	 Call write_variable_in_CASE_NOTE(worker_signature)
End If

If appeal_actions = "Appeal Resolution" then
    BeginDialog, 0, 0, 241, 130, "Appeal resolution"
      DropListBox 125, 5, 55, 15, "Select one..."+chr(9)+"Yes"+chr(9)+"No", actions_needed
      DropListBox 125, 25, 55, 20, "Select one..."+chr(9)+"Yes"+chr(9)+"No", op_needed
      EditBox 125, 45, 85, 15, overpayment_amount
      EditBox 125, 70, 85, 15, worker_signature
      ButtonGroup ButtonPressed
    	OkButton 105, 95, 50, 15
    	CancelButton 160, 95, 50, 15
      Text 10, 10, 115, 10, "Is action needed?"
      Text 10, 30, 115, 10, "Overpayments required?"
      Text 10, 50, 105, 10, "Overpayment Amount, if any:"
      Text 10, 75, 65, 10, "Worker Signature:"
    EndDialog
	Do
		DO
			err_msg = "" 
			Dialog
			cancel_confirmation 
			If actions_needed = "Select one..." THEN err_msg = err_msg & vbNewLine & "* Please select whether action is needed by caseworker"
			If op_needed = "Select one..." THEN err_msg = err_msg & vbNewLine & "* Please select whether overpayments are required"
			IF overpayment_amount = "" THEN err_msg = err_msg & vbNewLine & "* Please enter the amount of the overpayment"
			IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine		
		Loop until err_msg = ""	
    	Call check_for_password(are_we_passworded_out)
	LOOP UNTIL check_for_password(are_we_passworded_out) = False		 

	start_a_blank_case_note      'navigates to case/note and puts case/note into edit mode
	Call write_variable_in_CASE_NOTE("Appeal Resolution")
	Call write_bullet_and_variable_in_CASE_NOTE("Is action needed?", actions_needed)
	Call write_bullet_and_variable_in_CASE_NOTE("Overpayments required?", op_needed)
	Call write_bullet_and_variable_in_CASE_NOTE("Overpayment Amount, if any", overpayment_amount)
	Call write_variable_in_CASE_NOTE ("---")
	call write_variable_in_CASE_NOTE(worker_signature)
END IF 

script_end_procedure("") 