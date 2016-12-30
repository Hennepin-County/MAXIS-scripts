'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - CLIENT CONTACT.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 195          'manual run time in seconds
STATS_denomination = "C"        'C is for each case
'END OF stats block=========================================================================================================

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else											'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
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
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE MAIN DIALOG--------------------------------------------------------------------------------------------------
BeginDialog client_contact_dialog, 0, 0, 386, 320, "Client contact"
  ComboBox 50, 5, 60, 15, "Phone call"+chr(9)+"Voicemail"+chr(9)+"Email"+chr(9)+"Office visit"+chr(9)+"Letter", contact_type
  DropListBox 115, 5, 45, 10, "from"+chr(9)+"to", contact_direction
  ComboBox 165, 5, 85, 15, "client"+chr(9)+"AREP"+chr(9)+"Non-AREP"+chr(9)+"SWKR", who_contacted
  EditBox 280, 5, 100, 15, regarding
  EditBox 70, 25, 65, 15, phone_number
  EditBox 225, 25, 85, 15, when_contact_was_made
  EditBox 70, 45, 65, 15, MAXIS_case_number
  EditBox 70, 65, 65, 15, Mnsure_IC_number
  EditBox 70, 85, 310, 15, contact_reason
  EditBox 70, 105, 310, 15, actions_taken
  EditBox 65, 140, 310, 15, verifs_needed
  EditBox 65, 160, 310, 15, case_status
  EditBox 80, 180, 295, 15, cl_instructions
  CheckBox 5, 205, 255, 10, "Check here if you want to TIKL out for this case after the case note is done.", TIKL_check
  CheckBox 5, 220, 255, 10, "Check here if you reminded client about the importance of the CAF 1.", caf_1_check
  CheckBox 5, 235, 135, 10, "Check here if you sent forms to AREP.", Sent_arep_checkbox
  CheckBox 5, 250, 120, 10, "Check here if follow-up is needed.", follow_up_needed_checkbox
  CheckBox 20, 285, 105, 10, "Answered caller's question", Call_center_answer_check
  CheckBox 20, 300, 105, 10, "Transferred caller to Worker", call_center_transfer_check
  EditBox 315, 275, 65, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 275, 300, 50, 15
    CancelButton 330, 300, 50, 15
  Text 5, 10, 45, 10, "Contact type:"
  Text 260, 10, 15, 10, "Re:"
  Text 5, 30, 50, 10, "Phone number: "
  Text 150, 30, 70, 10, "Date/Time of Contact"
  Text 5, 50, 50, 10, "Case number: "
  Text 5, 90, 65, 10, "Reason for contact:"
  Text 5, 110, 50, 10, "Actions taken: "
  GroupBox 0, 125, 380, 75, "Helpful info for call centers (or front desks) to pass on to clients"
  Text 5, 145, 50, 10, "Verifs needed: "
  Text 5, 165, 45, 10, "Case status: "
  Text 5, 185, 75, 10, "Instructions/message:"
  GroupBox 5, 270, 130, 45, "Call Center:"
  Text 240, 280, 70, 10, "Sign your case note: "
  CheckBox 150, 45, 65, 10, "Used Interpreter", used_interpreter_checkbox
  Text 5, 70, 60, 10, "Mnsure IC number:"
EndDialog


'THE SCRIPT--------------------------------------------------------------------------------------------------
'CONNECTING TO MAXIS & GRABBING THE CASE NUMBER
EMConnect ""
CALL MAXIS_case_number_finder(MAXIS_case_number)

'updates the "when contact was made" variable to show the current date & time
when_contact_was_made = date & ", " & time
DO
	Do
		Do
			Do
				Dialog client_contact_dialog
				cancel_confirmation
				IF contact_reason = "" or contact_type = "" Then MsgBox("You must enter a reason for contact, as well as a type (phone, etc.).")
			Loop until contact_reason <> "" and contact_type <> ""
			IF worker_signature = "" THEN MsgBox "Please sign your note"
		LOOP UNTIL worker_signature <>""
		If (isnumeric(MAXIS_case_number) = False and len(MAXIS_case_number) <> 8) then MsgBox "You must enter either a valid MAXIS or MCRE case number."
	Loop until (isnumeric(MAXIS_case_number) = True) or (isnumeric(MAXIS_case_number) = False and len(MAXIS_case_number) = 8)
	call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false

'checking for an active MAXIS session
Call check_for_MAXIS(False)

'THE CASE NOTE----------------------------------------------------------------------------------------------------
start_a_blank_case_note
CALL write_variable_in_CASE_NOTE(contact_type & " " & contact_direction & " " & who_contacted & " re: " & regarding)
If Used_interpreter_checkbox = checked THEN
	CALL write_variable_in_CASE_NOTE("* Contact was made: " & when_contact_was_made & " w/ interpreter")
Else
	CALL write_bullet_and_variable_in_CASE_NOTE("Contact was made", when_contact_was_made)
End if
CALL write_bullet_and_variable_in_CASE_NOTE("Phone number", phone_number)
CALL write_bullet_and_variable_in_CASE_NOTE("MNSURE/IC number", Mnsure_IC_number)
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
	script_end_procedure("Success! Follow-up is needed for case number: " & MAXIS_case_number)
Else
	script_end_procedure("")
End if
