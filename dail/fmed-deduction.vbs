'Required for statistical purposes===============================================================================
name_of_script = "DAIL - FMED DEDUCTION.vbs"
start_time = timer
STATS_counter = 1              'sets the stats counter at one
STATS_manualtime = 127         'manual run time in seconds
STATS_denomination = "C"       'C is for case
'END OF stats block==============================================================================================

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
CALL changelog_update("09/16/2022", "Update to ensure Worker Signature is in all scripts that CASE/NOTE.", "MiKayla Handley, Hennepin County") '#316
call changelog_update("07/13/2020", "Updated following verbiage: 'client' changed to 'resident', 'worker' changed to 'team'.", "Ilse Ferris, Hennepin County")
call changelog_update("05/28/2020", "Update to the notice wording, added virtual drop box information.", "MiKayla Handley, Hennepin County")
CALL changelog_update("12/29/2017", "Coordinates for sending MEMO's has changed in SPEC/MEMO. Updated script to support change.", "Ilse Ferris, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'<<<<<GO THROUGH THE SCRIPT AND REMOVE REDUNDANT FUNCTIONS, THANKS TO CUSTOM FUNCTIONS THEY ARE NOT REQUIRED.
EMConnect ""

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 281, 45, "Worker signature"
  EditBox 60, 5, 70, 15, MAXIS_case_number
  EditBox 80, 25, 135, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 225, 5, 50, 15
    CancelButton 225, 25, 50, 15
  Text 5, 10, 50, 10, "Case Number:"
  Text 5, 30, 70, 10, "Sign your case note:"
EndDialog

DO
	DO
		err_msg = ""
		Dialog Dialog1
		cancel_confirmation
		Call validate_MAXIS_case_number(err_msg, "*")
		IF worker_signature = "" THEN err_msg = err_msg & vbCr & "* Please sign your case note."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
	LOOP UNTIL err_msg = ""									'loops until all errors are resolved
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
LOOP UNTIL  are_we_passworded_out = false					'loops until user passwords back in
call start_a_new_spec_memo(memo_opened, True, forms_to_arep, forms_to_swkr, send_to_other, other_name, other_street, other_city, other_state, other_zip, True)	'navigates to spec/memo and opens into edit mode

'Writes the info into the MEMO.
Call write_variable_in_SPEC_MEMO("************************************************************")
Call write_variable_in_SPEC_MEMO("You are turning 60 next month, so you may be eligible for a new deduction for SNAP. Residents who are over 60 years old may receive increased SNAP benefits if they have recurring medical bills over $35 each month.")

Call write_variable_in_SPEC_MEMO("If you have medical bills over $35 each month, please contact your team to discuss adjusting your benefits. You will need to send in proof of the medical bills, such as pharmacy receipts, an explanation of benefits, or premium notices.")
CALL write_variable_in_SPEC_MEMO("You now have an option to use an email to return documents to Hennepin County. Write the case number and full name associated with the case in the body of the email. Only the following types are accepted PNG, JPG, TIFF, DOC, PDF, and HTML. You will not receive confirmation of receipt or failure. To obtain information about your case please contact your team. EMAIL: hhsews@hennepin.us ")
Call write_variable_in_SPEC_MEMO("************************************************************")
PF4

Call navigate_to_MAXIS_screen("CASE", "NOTE")
PF9 'edit mode
Call write_variable_in_case_note("MEMBER HAS TURNED 60 - NOTIFY ABOUT POSSIBLE FMED DEDUCTION")
Call write_variable_in_case_note("* Sent MEMO to client about FMED deductions.")
Call write_variable_in_case_note("---")
Call write_variable_in_case_note(worker_signature)

script_end_procedure("The script has sent a MEMO to the client about the possible FMED deduction, and has case noted the action.")
