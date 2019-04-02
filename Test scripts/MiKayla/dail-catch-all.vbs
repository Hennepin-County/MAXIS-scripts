'Required for statistical purposes==========================================================================================
name_of_script = "DAIL - CATCH ALL.vbs"
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
call changelog_update("04/01/2019", "Initial version.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE MAIN DIALOG--------------------------------------------------------------------------------------------------
BeginDialog catch_all_dialog, 0, 0, 376, 140, "DAIL CATCH ALL"
  EditBox 60, 5, 65, 15, MAXIS_case_number
  EditBox 305, 5, 65, 15, METS_IC_number
  EditBox 60, 25, 100, 15, when_contact_was_made
  EditBox 270, 25, 100, 15, DAIL_read
  EditBox 60, 45, 310, 15, actions_taken
  EditBox 60, 65, 310, 15, verifs_needed
  EditBox 60, 85, 310, 15, cl_instructions
  EditBox 65, 120, 195, 15, worker_signature
  CheckBox 5, 105, 110, 10, "Check here if you want to TIKL.", TIKL_check
  ButtonGroup ButtonPressed
    OkButton 275, 120, 45, 15
    CancelButton 325, 120, 45, 15
  Text 5, 10, 50, 10, "Case number: "
  Text 245, 10, 60, 10, "METS IC number:"
  Text 5, 30, 40, 10, "Date/Time:"
  Text 255, 30, 15, 10, "Re:"
  Text 5, 50, 50, 10, "Actions taken: "
  Text 5, 70, 50, 10, "Verifs needed: "
  Text 5, 90, 45, 10, "Other notes:"
  Text 5, 125, 60, 10, "Worker signature:"
EndDialog


'THE SCRIPT--------------------------------------------------------------------------------------------------
'CONNECTING TO MAXIS & GRABBING THE CASE NUMBER
EMConnect ""
EMReadscreen dail_check, 4, 2, 48 'changed from DAIL to view to ensure we are in DAIL/DAIL'
IF dail_check <> "DAIL" THEN script_end_procedure("Your cursor is not set on a message type. Please select an appropriate DAIL message and try again.")
IF dail_check = "DAIL" THEN
	EMSendKey "T"
	transmit
	'Making sure that the user is on an acceptable DAIL message
	EMReadScreen DAIL_type, 4, 6, 6
	EMReadScreen MAXIS_case_number, 8, 5, 73
	MAXIS_case_number= TRIM(MAXIS_case_number)
	EMSendKey "X"
	'THE ENTIRE MESSAGE TEXT IS DISPLAYED'
	EmReadScreen error_msg 37, 24, 02
	EmReadScreen DAIL_read 50, 6, 20
END IF

'updates the "when contact was made" variable to show the current date & time
when_contact_was_made = date & ", " & time

Do
    Do
        err_msg = ""
		Dialog catch_all_dialog
		cancel_confirmation
        If (isnumeric(MAXIS_case_number) = False and len(MAXIS_case_number) <> 8) then err_msg = err_msg & vbcr & "* Enter a valid case number."
		If trim(actions_taken) = "" then err_msg = err_msg & vbcr & "* Please enter the action taken."
        If trim(DAIL_read) = "" then err_msg = err_msg & vbcr & "* Please enter the DAIL."
		If trim(worker_signature) = "" then err_msg = err_msg & vbcr & "* Sign your case note."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
	LOOP UNTIL err_msg = ""									'loops until all errors are resolved
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'checking for an active MAXIS session
Call check_for_MAXIS(False)

'THE CASE NOTE----------------------------------------------------------------------------------------------------
start_a_blank_case_note
CALL write_variable_in_CASE_NOTE("Action Taken: " & " re: " & regarding)
CALL write_bullet_and_variable_in_CASE_NOTE("Action was taken: ", when_contact_was_made)
CALL write_bullet_and_variable_in_CASE_NOTE("Phone number", phone_number)
CALL write_bullet_and_variable_in_CASE_NOTE("METS/IC number", METS_IC_number)

CALL write_bullet_and_variable_in_CASE_NOTE("Actions Taken", actions_taken)
CALL write_bullet_and_variable_in_CASE_NOTE("Verifs Needed", verifs_needed)
CALL write_bullet_and_variable_in_CASE_NOTE("Instructions/Message for CL", cl_instructions)
CALL write_bullet_and_variable_in_CASE_NOTE("Case Status", case_status)
CALL write_variable_in_CASE_NOTE("---")
CALL write_variable_in_CASE_NOTE(worker_signature)

'TIKLING
IF TIKL_check = checked THEN CALL navigate_to_MAXIS_screen("dail", "writ")

script_end_procedure_with_error_reporting("DAIL has been case noted")
End if
