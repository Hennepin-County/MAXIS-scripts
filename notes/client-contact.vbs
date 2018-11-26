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
call changelog_update("06/19/2018", "Added FAX to contact type, changed MNSURE IC # to METS IC #, updated look of dialog and back end mandatory field handling. Also removed message box prior to navigating to DAIL/WRIT.", "Ilse Ferris, Hennepin County")
call changelog_update("09/28/2017", "Removed call center information from bottom of dialog.", "Ilse Ferris, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE MAIN DIALOG--------------------------------------------------------------------------------------------------
BeginDialog client_contact_dialog, 0, 0, 386, 245, "Client contact"
  ComboBox 50, 5, 60, 15, "Phone call"+chr(9)+"Voicemail"+chr(9)+"Email"+chr(9)+"Fax"+chr(9)+"Office visit"+chr(9)+"Letter", contact_type
  DropListBox 115, 5, 45, 10, "from"+chr(9)+"to", contact_direction
  ComboBox 165, 5, 85, 15, "client"+chr(9)+"AREP"+chr(9)+"Non-AREP"+chr(9)+"SWKR", who_contacted
  EditBox 280, 5, 100, 15, regarding
  EditBox 70, 25, 65, 15, phone_number
  EditBox 280, 25, 100, 15, when_contact_was_made
  EditBox 70, 45, 65, 15, MAXIS_case_number
  CheckBox 145, 50, 65, 10, "Used Interpreter", used_interpreter_checkbox
  EditBox 280, 45, 65, 15, METS_IC_number
  EditBox 70, 70, 310, 15, contact_reason
  EditBox 70, 90, 310, 15, actions_taken
  EditBox 65, 125, 310, 15, verifs_needed
  EditBox 65, 145, 310, 15, case_status
  EditBox 85, 165, 290, 15, cl_instructions
  CheckBox 5, 190, 255, 10, "Check here if you want to TIKL out for this case after the case note is done.", TIKL_check
  CheckBox 5, 205, 255, 10, "Check here if you reminded client about the importance of the CAF 1.", caf_1_check
  CheckBox 260, 190, 95, 10, "Forms were sent to AREP.", Sent_arep_checkbox
  CheckBox 260, 205, 120, 10, "Follow up is needed on this case.", follow_up_needed_checkbox
  EditBox 70, 225, 195, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 275, 225, 50, 15
    CancelButton 330, 225, 50, 15
  Text 260, 10, 15, 10, "Re:"
  Text 5, 30, 50, 10, "Phone number: "
  Text 200, 30, 75, 10, "Date/Time of Contact:"
  Text 5, 50, 50, 10, "Case number: "
  Text 5, 75, 65, 10, "Reason for contact:"
  Text 5, 95, 50, 10, "Actions taken: "
  GroupBox 5, 110, 375, 75, "Helpful info for call centers (or front desks) to pass on to clients:"
  Text 10, 130, 50, 10, "Verifs needed: "
  Text 10, 150, 45, 10, "Case status: "
  Text 10, 170, 75, 10, "Instructions/message:"
  Text 5, 230, 60, 10, "Worker signature:"
  Text 5, 10, 45, 10, "Contact type:"
  Text 215, 50, 60, 10, "METS IC number:"
EndDialog

'THE SCRIPT--------------------------------------------------------------------------------------------------
'CONNECTING TO MAXIS & GRABBING THE CASE NUMBER
EMConnect ""
CALL MAXIS_case_number_finder(MAXIS_case_number)

'updates the "when contact was made" variable to show the current date & time
when_contact_was_made = date & ", " & time

Do 
    Do 
        err_msg = ""
		Dialog client_contact_dialog
		cancel_confirmation
        If (isnumeric(MAXIS_case_number) = False and len(MAXIS_case_number) <> 8) then err_msg = err_msg & vbcr & "* Enter a valid case number."
		If trim(contact_reason) = "" then err_msg = err_msg & vbcr & "* Enter a reason for contact."
        If trim(contact_type) = "" then err_msg = err_msg & vbcr & "* Enter the contact type (phone, etc.)."
		If trim(worker_signature) = "" then err_msg = err_msg & vbcr & "* Sign your case note."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
	LOOP UNTIL err_msg = ""									'loops until all errors are resolved
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

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
CALL write_bullet_and_variable_in_CASE_NOTE("METS/IC number", METS_IC_number)
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
CALL write_variable_in_CASE_NOTE("---")
CALL write_variable_in_CASE_NOTE(worker_signature)

'TIKLING
IF TIKL_check = checked THEN CALL navigate_to_MAXIS_screen("dail", "writ")

'If case requires followup, it will create a MsgBox (via script_end_procedure) explaining that followup is needed. This MsgBox gets inserted into the statistics database for counties using that function. This will allow counties to "pull statistics" on follow-up, including case numbers, which can be used to track outcomes.
If follow_up_needed_checkbox = checked then
	script_end_procedure("Success! Follow-up is needed for case number: " & MAXIS_case_number)
Else
	script_end_procedure("")
End if
