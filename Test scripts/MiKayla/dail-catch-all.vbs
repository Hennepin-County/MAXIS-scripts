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
BeginDialog catch_all_dialog, 0, 0, 361, 140, "DAIL_type   DAIL CATCH ALL   & MAXIS_case_number"
  EditBox 60, 45, 215, 15, actions_taken
  EditBox 60, 65, 215, 15, verifs_needed
  EditBox 60, 85, 215, 15, cl_instructions
  EditBox 65, 120, 95, 15, worker_signature
  CheckBox 5, 105, 110, 10, "Check here if you want to TIKL.", TIKL_check
  ButtonGroup ButtonPressed
    OkButton 180, 120, 45, 15
    CancelButton 230, 120, 45, 15
  Text 5, 50, 50, 10, "Actions taken: "
  Text 5, 70, 50, 10, "Verifs needed: "
  Text 5, 90, 45, 10, "Other notes:"
  Text 5, 125, 60, 10, "Worker signature:"
  CheckBox 135, 105, 140, 10, "Check here that ECF has been reviewed ", ECF_reviewed
  GroupBox 5, 5, 345, 45, ""DAIL for case # " MAXIS_case_number"
EndDialog
'do we need a date rcvd or save that for docs rcvd'

'THE SCRIPT--------------------------------------------------------------------------------------------------
'CONNECTING TO MAXIS & GRABBING THE CASE NUMBER
EMConnect ""
EMReadscreen dail_check, 4, 2, 48 'changed from DAIL to view to ensure we are in DAIL/DAIL'
IF dail_check <> "DAIL" THEN script_end_procedure("Your cursor is not set on a message type. Please select an appropriate DAIL message and try again.")
IF dail_check = "DAIL" THEN
	EMSendKey "T"
	TRANSMIT
	EMReadScreen DAIL_type, 4, 6, 6 'read the DAIL msg'
	DAIL_type = trim(DAIL_type)
	IF DAIL_type = "TIKL" or DAIL_type = "PEPR"  or DAIL_type = "SSN" or DAIL_type = "INFO" THEN
		match_found = TRUE
	ELSE
		match_found = FALSE
		script_end_procedure("This is not an supported DAIL currently.Please select a WAGE match DAIL, and run the script again.")
	END IF
	IF match_found = TRUE THEN
		'The following reads the message in full for the end part (which tells the worker which message was selected)
		EMReadScreen full_message, 58, 6, 20
		EmReadScreen MAXIS_case_number, 8, 5, 73
		MAXIS_case_number = trim(MAXIS_case_number)

		EMReadScreen extra_info, 1, 06, 80
		If extra_info = "+" or extra_info = "&" THEN
	    EMSendKey "X"
		TRANSMIT
	    'THE ENTIRE MESSAGE TEXT IS DISPLAYED'
	    EmReadScreen error_msg, 37, 24, 02
		row = 1
		col = 1
		EMSearch "Case Number", row, col 	'Has to search, because every once in a while the rows and columns can slide one or two positions.
		'If row = 0 then script_end_procedure("MAXIS may be busy: the script appears to have errored out. This should be temporary. Try again in a moment. If it happens repeatedly contact the alpha user for your agency.")
		EMReadScreen first_line, 61, row + 3, col - 40 'JOB DETAIL Reads each line for the case note. COL needs to be subtracted from because of NDNH message format differs from original new hire format.
			'first_line = replace(first_line, "FOR  ", "FOR ")	'need to replaces 2 blank spaces'
			first_line = trim(first_line)
		EMReadScreen second_line, 61, row + 4, col - 40
			second_line = trim(second_line)
		EMReadScreen third_line, 61, row + 5, col - 40 'maxis name'
			third_line = trim(third_line)
			'third_line = replace(third_line, ",", ", ")
		EMReadScreen fourth_line, 61, row + 6, col - 40'new hire name'
			fourth_line = trim(fourth_line)
			'fourth_line = replace(fourth_line, ",", ", ")
		EMReadScreen fifth_line, 61, row + 7, col - 40'new hire name'
			fifth_line = trim(fifth_line)
		EmReadScreen DAIL_read, 50, 6, 20
		DAIL_read = trim(DAIL_read)
	END IF
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
CALL write_variable_in_CASE_NOTE("Action taken re: " & DAIL_type)
CALL write_variable_in_case_note(first_line)
CALL write_variable_in_case_note(second_line)
CALL write_variable_in_case_note(third_line)
CALL write_variable_in_case_note(fourth_line)
CALL write_variable_in_case_note(fifth_line)
CALL write_variable_in_case_note("---")
CALL write_bullet_and_variable_in_case_note("Action was taken: ", when_contact_was_made)
CALL write_bullet_and_variable_in_case_note("Actions taken", actions_taken)
CALL write_bullet_and_variable_in_case_note("Action was taken ", when_contact_was_made)
CALL write_bullet_and_variable_in_case_note("Phone number", phone_number)
CALL write_bullet_and_variable_in_case_note("METS/IC number", METS_IC_number)
CALL write_bullet_and_variable_in_case_note("Verifs Needed", verifs_needed)
CALL write_bullet_and_variable_in_case_note("Instructions/Message for CL", cl_instructions)
CALL write_bullet_and_variable_in_case_note("Case Status", case_status)
CALL write_variable_in_CASE_NOTE("---")
CALL write_variable_in_CASE_NOTE(worker_signature)

'TIKLING
'IF TIKL_check = checked THEN CALL navigate_to_MAXIS_screen("dail", "writ")

script_end_procedure_with_error_report(DAIL_type & vbcr &  first_line & vbcr & " DAIL has been case noted")
End if
