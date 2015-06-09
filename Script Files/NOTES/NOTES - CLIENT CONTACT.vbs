'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - CLIENT CONTACT.vbs"
start_time = timer

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN		'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF default_directory = "C:\DHS-MAXIS-Scripts\Script Files\" THEN			'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		ELSEIF beta_agency = "" or beta_agency = True then							'If you're a beta agency, you should probably use the beta branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/BETA/MASTER%20FUNCTIONS%20LIBRARY.vbs"
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

'THE MAIN DIALOG--------------------------------------------------------------------------------------------------
BeginDialog client_contact_dialog, 0, 0, 386, 300, "Client contact"
  ComboBox 50, 5, 60, 15, "Phone call"+chr(9)+"Voicemail"+chr(9)+"Email"+chr(9)+"Office visit"+chr(9)+"Letter", contact_type
  DropListBox 115, 5, 45, 10, "from"+chr(9)+"to", contact_direction
  ComboBox 165, 5, 85, 15, "client"+chr(9)+"AREP"+chr(9)+"Non-AREP"+chr(9)+"SWKR", who_contacted
  EditBox 280, 5, 100, 15, regarding
  EditBox 80, 25, 65, 15, phone_number
  EditBox 225, 25, 85, 15, when_contact_was_made
  EditBox 55, 45, 55, 15, case_number
  EditBox 70, 65, 310, 15, contact_reason
  EditBox 55, 85, 325, 15, actions_taken
  EditBox 65, 120, 310, 15, verifs_needed
  EditBox 120, 140, 255, 15, cl_instructions
  EditBox 60, 160, 315, 15, case_status
  CheckBox 5, 185, 255, 10, "Check here if you want to TIKL out for this case after the case note is done.", TIKL_check
  CheckBox 5, 200, 255, 10, "Check here if you reminded client about the importance of the CAF 1.", caf_1_check
  CheckBox 5, 215, 135, 10, "Check here if you sent forms to AREP.", Sent_arep_checkbox
  CheckBox 5, 230, 120, 10, "Check here if follow-up is needed.", follow_up_needed_checkbox
  CheckBox 20, 265, 105, 10, "Answered caller's question", Call_center_answer_check
  CheckBox 20, 280, 105, 10, "Transferred caller to Worker", call_center_transfer_check
  EditBox 315, 255, 65, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 275, 280, 50, 15
    CancelButton 330, 280, 50, 15
  Text 5, 10, 45, 10, "Contact type:"
  Text 260, 10, 15, 10, "Re:"
  Text 25, 30, 50, 10, "Phone number: "
  Text 150, 30, 70, 10, "Date/Time of Contact"
  Text 5, 50, 50, 10, "Case number: "
  Text 5, 70, 65, 10, "Reason for contact:"
  Text 5, 90, 50, 10, "Actions taken: "
  GroupBox 0, 105, 380, 75, "Helpful info for call centers (or front desks) to pass on to clients"
  Text 10, 125, 50, 10, "Verifs needed: "
  Text 10, 145, 105, 10, "Instructions/message for client:"
  Text 10, 165, 45, 10, "Case status: "
  GroupBox 5, 250, 130, 45, "Call Center:"
  Text 240, 260, 70, 10, "Sign your case note: "
EndDialog


'THE SCRIPT--------------------------------------------------------------------------------------------------

EMConnect ""


'updating case number insert w/function name             
CALL MAXIS_case_number_finder(case_number)

'updates the "when contact was made" variable to show the current date & time
when_contact_was_made = date & ", " & time


DO
	Do
		Do
			Do
				DO
				Dialog client_contact_dialog
				If buttonpressed = 0 then stopscript
				IF contact_reason = "" or contact_type = "" Then MsgBox("You must enter a reason for contact, as well as a type (phone, etc.).")
			Loop until contact_reason <> "" and contact_type <> ""
			IF worker_signature = "" THEN MsgBox "Please sign your note"
			LOOP UNTIL worker_signature <>""
			If (isnumeric(case_number) = False and len(case_number) <> 8) then MsgBox "You must enter either a valid MAXIS or MCRE case number."
		Loop until (isnumeric(case_number) = True) or (isnumeric(case_number) = False and len(case_number) = 8)
		transmit
		If isnumeric(case_number) = True then
			EMReadScreen MAXIS_check, 5, 1, 39
			If MAXIS_check <> "MAXIS" then MsgBox "You are not in MAXIS. Navigate your screen to MAXIS and try again. You might be passworded out."
		Else
			MMIS_row = 1
			MMIS_col = 1
			EMSearch "MMIS", MMIS_row, MMIS_col
			If MMIS_row <> 1 then
				EMReadScreen OSLT_check, 4, 1, 52 'Because cases that are on the "OSLT" screen in MMIS don't contain the characters "MMIS" in the top line.
				If OSLT_check = "OSLT" then MMIS_row = 1
			End if
			If MMIS_row <> 1 then MsgBox "You are not in MMIS. Navigate your screen to MMIS and try again. You might be passworded out."
		End if
	Loop until MAXIS_check = "MAXIS" or MMIS_row = 1
	If isnumeric(case_number) = True then 
		call navigate_to_MAXIS_screen("case", "note")
		PF9
		EMReadScreen mode_check, 7, 20, 3
		If mode_check <> "Mode: A" and mode_check <> "Mode: E" then MsgBox "The script doesn't appear to be able to find your case note. Are you in inquiry? If so, navigate to production on the screen where you clicked the script button, and try again. Otherwise, you might have forgotten to type a valid case number."
	Else
		call MMIS_RKEY_finder
		EMWriteScreen "c", 2, 19
		EMWriteScreen case_number, 9, 19
		transmit
		EMReadScreen RKEY_check, 4, 1, 52 'CHECKING FOR RKEY, IF RENEWAL IS DUE A WARNING MESSAGE WILL NEED TO BE MOVED PAST.
		If RKEY_check = "RKEY" then transmit
		PF4
		PF11
		EMReadScreen MMIS_edit_mode_check, 5, 5, 2
		If MMIS_edit_mode_check <> "'''''" then script_end_procedure("MMIS edit mode not found. Are you in inquiry? Is MMIS not functioning? Shut down this script and try again. If it continues to not work, email your script administrator the case number, and a screenshot of MMIS.")
	End if
Loop until (mode_check = "Mode: A" or mode_check = "Mode: E") or (MMIS_edit_mode_check = "'''''") 



'Writing case note w/updated functions
CALL write_variable_in_CASE_NOTE(contact_type & " " & contact_direction & " " & who_contacted & " re: " & regarding)
CALL write_bullet_and_variable_in_CASE_NOTE("Contact was made", when_contact_was_made)
CALL write_bullet_and_variable_in_CASE_NOTE("Phone number", phone_number)
CALL write_bullet_and_variable_in_CASE_NOTE("Reason for contact", contact_reason)
CALL write_bullet_and_variable_in_CASE_NOTE("Actions Taken", actions_taken)
CALL write_bullet_and_variable_in_CASE_NOTE("Verifs Needed", verifs_needed)
CALL write_bullet_and_variable_in_CASE_NOTE("Instructions/Message for CL", cl_instructions)
CALL write_bullet_and_variable_in_CASE_NOTE("Case Status", case_status)

'checkbox results
IF caf_1_check = checked THEN CALL write_variable_in_CASE_NOTE("* Reminded client about the importance of submitting the CAF 1.")
IF Sent_arep_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* Sent form(s) to AREP.")
IF call_center_answer_check = checked THEN CALL write_variable_in_CASE_NOTE("* Call center answered caller's question.")
IF call_center_transfer_check = checked THEN CALL write_variable_in_CASE_NOTE("* Call center transferred call to a worker.")
IF follow_up_needed_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* Follow-up is needed.")

'Worker sig
CALL write_variable_in_CASE_NOTE("---")
CALL write_variable_in_CASE_NOTE(worker_signature)

'TIKLING
IF TIKL_check = checked THEN
	MsgBox "The script will now navigate to a TIKL."
	CALL navigate_to_MAXIS_screen("dail", "writ")
END IF

'If case requires followup, it will create a MsgBox (via script_end_procedure) explaining that followup is needed. This MsgBox gets inserted into the statistics database for counties using that function. This will allow counties to "pull statistics" on follow-up, including case numbers, which can be used to track outcomes.
If follow_up_needed_checkbox = checked then
	script_end_procedure("Success! Follow-up is needed for case number: " & case_number)
Else
	script_end_procedure("")
End if