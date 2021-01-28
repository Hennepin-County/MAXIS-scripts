'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - MHC CLIENT CONTACT.vbs"
start_time = timer

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

'THE SCRIPT--------------------------------------------------------------------------------------------------
EMConnect ""

date_of_call = date & ""

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 311, 110, " MHC client contact"
  EditBox 60, 10, 50, 15, MAXIS_case_number
  EditBox 145, 10, 50, 15, person_pmi
  EditBox 255, 10, 50, 15, date_of_call
  EditBox 60, 30, 245, 15, Changes_reported
  EditBox 60, 50, 245, 15, actions_taken
  CheckBox 60, 70, 160, 10, "Check if worker sent manual enrollment form.", Check_send_enrollment
  EditBox 70, 90, 125, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 200, 90, 50, 15
    CancelButton 255, 90, 50, 15
  Text 5, 35, 55, 10, "Reason for call: "
  Text 10, 15, 50, 10, "Case Number:"
  Text 0, 95, 65, 10, "Worker's Signature:"
  Text 210, 15, 40, 10, "Date of call:"
  Text 125, 15, 15, 10, "PMI:"
  Text 5, 55, 50, 10, "Actions taken:"
EndDialog

Do
	Do
		'Do
			DO
				Dialog Dialog1
				If buttonpressed = 0 then stopscript
				IF worker_signature = "" THEN MsgBox "Please sign your note."
				IF actions_taken = "" then MsgBox "Please enter your actions taken."
			Loop until worker_signature <> "" AND actions_taken <> ""
		If (isnumeric(MAXIS_case_number) = False and isnumeric(person_pmi) = False) then MsgBox "You must enter either a valid MAXIS case number or PMI number"
	Loop until (isnumeric(MAXIS_case_number) = True) or (isnumeric(person_pmi) = True)
	transmit
	MMIS_row = 1
	MMIS_col = 1
	EMSearch "MMIS", MMIS_row, MMIS_col
	If MMIS_row <> 1 then
		EMReadScreen OSLT_check, 4, 1, 52 'Because cases that are on the "OSLT" screen in MMIS don't contain the characters "MMIS" in the top line.
		If OSLT_check = "OSLT" then MMIS_row = 1
	If MMIS_row <> 1 then MsgBox "You are not in MMIS. Navigate your screen to MMIS and try again. You might be passworded out."
	End if
Loop until MMIS_row = 1
If isnumeric(MAXIS_case_number) = True then
	If len(MAXIS_case_number) < 8 then 'This will generate an 8 digit Case Number.
		Do
			MAXIS_case_number = "0" & MAXIS_case_number
		Loop until len(MAXIS_case_number) = 8
	End if
	Call get_to_RKEY
	EMWriteScreen "C", 2, 19
	EMWriteScreen MAXIS_case_number, 9, 19
	transmit
	transmit
	transmit
	EMWriteScreen "x", 11, 2
	transmit
	PF4
	PF11
	EMReadScreen MMIS_edit_mode_check, 5, 5, 2
	If MMIS_edit_mode_check <> "'''''" then script_end_procedure("MMIS edit mode not found. Are you in inquiry? Is MMIS not functioning? Shut down this script and try again. If it continues to not work, email your script administrator the case number, and a screenshot of MMIS.")
Else
	If isnumeric(person_pmi) = true then
		If len(person_pmi) < 8 then 'This will generate an 8 digit PMI.
		Do
			person_pmi = "0" & person_pmi
		Loop until len(person_pmi) = 8
		End If
	End if
	Call get_to_RKEY
	EMWriteScreen "c", 2, 19
	EMWriteScreen person_pmi, 4, 19
	transmit
	PF4
	PF11
			EMReadScreen MMIS_edit_mode_check, 5, 5, 2
	If MMIS_edit_mode_check <> "'''''" then script_end_procedure("MMIS edit mode not found. Are you in inquiry? Is MMIS not functioning? Shut down this script and try again. If it continues to not work, email your script administrator the case number, and a screenshot of MMIS.")
	EMReadScreen MMIS_edit_mode_check, 5, 5, 2
	If MMIS_edit_mode_check <> "'''''" then script_end_procedure("MMIS edit mode not found. Are you in inquiry? Is MMIS not functioning? Shut down this script and try again. If it continues to not work, email your script administrator the case number, and a screenshot of MMIS.")
End if

CALL write_variable_in_MMIS_NOTE(variable)
CALL write_variable_in_MMIS_NOTE("Client Contact on " & date_of_call & " by phone")
CALL write_variable_in_MMIS_NOTE("Change reported: " & changes_reported)
CALL write_variable_in_MMIS_NOTE("Action Taken: " & actions_taken)
IF check_send_enrollment = checked THEN CALL write_variable_in_MMIS_NOTE("Sent enrollment form to client.")
CALL write_variable_in_MMIS_NOTE(worker_signature)
CALL write_variable_in_MMIS_NOTE ("*************************************************************************")
'logic to support the check box for worker checking off to case note that they are sending form to client.

IF check_send_enrollment THEN MsgBox "Remember to manually send an enrollment form to client as requested."
script_end_procedure("")
