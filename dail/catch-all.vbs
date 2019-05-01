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
'===============================================================================================END FUNCTIONS LIBRARY BLOCK

'===========================================================================================================CHANGELOG BLOCK
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("04/01/2019", "Initial version.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'=======================================================================================================END CHANGELOG BLOCK
'DIALOG
BeginDialog delete_message_dialog, 0, 0, 126, 45, "Double-Check the Computer's Work..."
  ButtonGroup ButtonPressed
    PushButton 10, 25, 50, 15, "YES", delete_button
    PushButton 60, 25, 50, 15, "NO", do_not_delete
  Text 30, 10, 65, 10, "Delete the DAIL??"
EndDialog

'-------------------------------------------------------------------------------------------------------THE SCRIPT
'CONNECTING TO MAXIS & GRABBING THE CASE NUMBER
EMConnect ""
'EMReadscreen dail_check, 4, 2, 48 'changed from DAIL to view to ensure we are in DAIL/DAIL'
'IF dail_check <> "DAIL" THEN script_end_procedure("Your cursor is not set on a message type. Please select an appropriate DAIL message and try again.")
'IF dail_check = "DAIL" THEN
'	EMSendKey "T"
'	TRANSMIT
'	EMReadScreen DAIL_type, 4, 6, 6 'read the DAIL msg'
'	DAIL_type = trim(DAIL_type)
'	IF DAIL_type = "TIKL" or DAIL_type = "PEPR"  or DAIL_type = "INFO" THEN
'		match_found = TRUE
'	ELSE
'		match_found = FALSE
'		script_end_procedure("This is not an supported DAIL currently. Please select TIKL, PEPR, SSN, or INFO DAIL, and run the script again.")
'	END IF
'	IF match_found = TRUE THEN
	    'do we need a date rcvd or save that for docs rcvd'
	    'The following reads the message in full for the end part (which tells the worker which message was selected)
	   ' EMReadScreen full_message, 59, 6, 20
		'full_message = trim(full_message)
	   ' EmReadScreen MAXIS_case_number, 8, 5, 73
	   ' MAXIS_case_number = trim(MAXIS_case_number)

	    'THE MAIN DIALOG--------------------------------------------------------------------------------------------------

        BeginDialog catch_all_dialog, 0, 0, 281, 150,  DAIL_type & "MESSAGE PROCESSED"
          EditBox 225, 35, 50, 15, docs_rcvd_date
          EditBox 65, 55, 210, 15, actions_taken
          EditBox 65, 75, 210, 15, verifs_needed
          EditBox 65, 95, 210, 15, other_notes
          EditBox 65, 130, 95, 15, worker_signature
          CheckBox 5, 40, 90, 10, "ECF has been reviewed ", ECF_reviewed
          CheckBox 5, 115, 110, 10, "Check here if you want to TIKL", TIKL_check
          ButtonGroup ButtonPressed
            OkButton 180, 130, 45, 15
            CancelButton 230, 130, 45, 15
          GroupBox 5, 5, 270, 25, "DAIL for case # " &  MAXIS_case_number
          Text 10, 15, 260, 10, full_message
          Text 5, 60, 50, 10, "Actions taken:"
          Text 5, 80, 50, 10, "Verifs needed:"
          Text 5, 100, 45, 10, "Other notes:"
          Text 5, 135, 60, 10, "Worker signature:"
          Text 120, 40, 100, 10, "If applicable - date doc(s) rcvd:"
        EndDialog

		EMWriteScreen "N", 6, 3         'Goes to Case Note - maintains tie with DAIL
		TRANSMIT
		'Starts a blank case note
		PF9
		EMReadScreen case_note_mode_check, 7, 20, 3
		If case_note_mode_check <> "Mode: A" then script_end_procedure("You are not in a case note on edit mode. You might be in inquiry. Try the script again in production.")
		'updates the "when contact was made" variable to show the current date & time
		when_contact_was_made = date & ", " & time

		Do
		    Do
		        err_msg = ""
				Dialog catch_all_dialog
				cancel_confirmation
		        If (isnumeric(MAXIS_case_number) = False and len(MAXIS_case_number) <> 8) then err_msg = err_msg & vbcr & "* Enter a valid case number."
				If trim(actions_taken) = "" then err_msg = err_msg & vbcr & "* Please enter the action taken."
		    	If trim(worker_signature) = "" then err_msg = err_msg & vbcr & "* Please ensure your case note is signed."
				IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
			LOOP UNTIL err_msg = ""									'loops until all errors are resolved
			CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
		Loop until are_we_passworded_out = false					'loops until user passwords back in

		'checking for an active MAXIS session
		Call check_for_MAXIS(False)
	'END IF
'END IF



'THE CASENOTE----------------------------------------------------------------------------------------------------
'start_a_blank_case_note
CALL write_variable_in_CASE_NOTE("=== " & DAIL_type & " - MESSAGE PROCESSED " & "===")
CALL write_variable_in_case_note("* " & full_message)
CALL write_variable_in_case_note(first_line)
CALL write_variable_in_case_note(second_line)
CALL write_variable_in_case_note(third_line)
CALL write_variable_in_case_note(fourth_line)
CALL write_variable_in_case_note(fifth_line)
CALL write_variable_in_case_note("---")
IF ECF_reviewed = CHECKED THEN CALL write_variable_in_case_note("* ECF reviewed and appropriate action taken")
CALL write_bullet_and_variable_in_case_note("Actions taken", actions_taken)
CALL write_bullet_and_variable_in_case_note("Action taken on", when_contact_was_made)
CALL write_bullet_and_variable_in_case_note("Verifications needed", verifs_needed)
CALL write_bullet_and_variable_in_case_note("Other notes", other_notes)
IF tikl_checkbox = checked THEN CALL write_variable_in_case_note("* TIKL'd to check for requested verifications")
CALL write_variable_in_CASE_NOTE("---")
CALL write_variable_in_CASE_NOTE(worker_signature)


DIALOG delete_message_dialog
    IF ButtonPressed = delete_button THEN
    	PF3 
    	PF3
    	DO
    		dail_read_row = 6
    		DO
    			EMReadScreen double_check, 59, dail_read_row, 20
				double_check = trim(double_check)
    			IF double_check = full_message THEN
                    EMWriteScreen "T", dail_read_row, 3
					TRANSMIT
                    EMReadScreen dail_case_number, 8, 5, 73
                    dail_case_number = trim(dail_case_number)
                    If dail_case_number = MAXIS_case_number Then EMWriteScreen "D", 6, 3
    				TRANSMIT
    				EXIT DO
    			ELSE
    				dail_read_row = dail_read_row + 1
    			END IF
    			IF dail_read_row = 19 THEN PF8
    		LOOP UNTIL dail_read_row = 19
    		EMReadScreen others_dail, 13, 24, 2
    		If others_dail = "** WARNING **" Then transmit
    	LOOP UNTIL double_check = full_message
    END IF

	'TIKLING
	IF TIKL_checkbox = checked THEN CALL navigate_to_MAXIS_screen("dail", "writ")
	'If worker checked to TIKL out, it goes to DAIL WRIT
	IF TIKL_checkbox = checked THEN
		CALL navigate_to_MAXIS_screen("DAIL","WRIT")
		CALL create_MAXIS_friendly_date(date, 10, 5, 18)
		EMSetCursor 9, 3
		EMSendKey "DAIL recieved " & DAIL_type & " " & verifs_needed & "."
	END IF

script_end_procedure_with_error_report(DAIL_type & vbcr &  first_line & vbcr & " DAIL has been case noted")
