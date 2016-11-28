'Required for statistical purposes==========================================================================================
name_of_script = "NOTICES - OVERDUE BABY.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 120                               'manual run time in seconds
STATS_denomination = "C"       'C is for each CASE
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

'DIALOG---------------------------------------------------------------------------------------------------------------------
BeginDialog NOTICES_overdue_baby_dialog, 0, 0, 141, 85, "NOTICES - OVERDUE BABY"
  EditBox 60, 5, 60, 15, MAXIS_case_number
  EditBox 70, 25, 60, 15, worker_signature
  CheckBox 5, 45, 100, 15, "TIKL for ten day follow up?", tikl_for_ten_day_follow_up_checkbox
  ButtonGroup ButtonPressed
    OkButton 30, 65, 50, 15
    CancelButton 80, 65, 50, 15
  Text 5, 5, 50, 15, "Case Number:"
  Text 5, 25, 60, 15, "Worker Signature:"
EndDialog

EndDialog

'THE SCRIPT------------------------------------------------------------------------------------------------------------------
'Connects to BlueZone default screen
EMConnect ""

'Searches for a case number
call MAXIS_case_number_finder(MAXIS_case_number)

Do
	Dialog NOTICES_overdue_baby_dialog
	If ButtonPressed = 0 then stopscript
	If MAXIS_case_number = ""  or isnumeric(MAXIS_case_number) = false then MsgBox "You did not enter a valid case number. Please try again."
	If worker_signature = "" then MsgBox "You did not sign your case note. Please try again."
Loop until MAXIS_case_number <> "" and isnumeric(MAXIS_case_number) = true and worker_signature <> ""
transmit
call check_for_MAXIS(True)

'Navigates into SPEC/MEMO
	call navigate_to_MAXIS_screen("SPEC", "MEMO")

'Checks to make sure we're past the SELF menu
	EMReadScreen still_self, 27, 2, 28
	If still_self = "Select Function Menu (SELF)" then script_end_procedure("Script was not able to get past SELF menu. Is case in background?")

'Creates a new MEMO. If it's unable the script will stop.
	PF5
	EMReadScreen memo_display_check, 12, 2, 33
	If memo_display_check = "Memo Display" then script_end_procedure("You are not able to go into update mode. Did you enter in inquiry by mistake? Please try again in production.")
	EMWriteScreen "x", 5, 10
	transmit

'Writes the info into the MEMO
EMSetCursor 3, 15
call write_variable_in_SPEC_MEMO("Our records indicate your due date has passed and you did not report the birth of your child or the pregnancy end date. Please contact us within 10 days of this notice with the following information or your case may close:")
call write_variable_in_SPEC_MEMO("")
call write_variable_in_SPEC_MEMO("* Date of the birth or pregnancy end date.")
call write_variable_in_SPEC_MEMO("* Baby's sex and full name.")
call write_variable_in_SPEC_MEMO("* Baby's social security number.")
call write_variable_in_SPEC_MEMO("* Full name of the baby's father.")
call write_variable_in_SPEC_MEMO("* Does the baby's father live in your home?")
call write_variable_in_SPEC_MEMO("* If so, does the father have a source of income?")
call write_variable_in_SPEC_MEMO("  (If so, what is the source of income?)")
call write_variable_in_SPEC_MEMO("* Is there other health insurance available through any       household member's employer, or privately?")
call write_variable_in_SPEC_MEMO("")
call write_variable_in_SPEC_MEMO("Thank you,")
PF4

'Navigates to blank case note
call navigate_to_MAXIS_screen("CASE", "NOTE")
PF9

'Writes the case note
call write_variable_in_CASE_NOTE("***Overdue Baby***")
call write_variable_in_CASE_NOTE("* SPEC/MEMO sent this date informing client that they need to report ")
call write_variable_in_CASE_NOTE("      information regarding the birth of their child, and/or pregnancy end ")
call write_variable_in_CASE_NOTE("      date, within 10 days or their case may close.")
call write_variable_in_CASE_NOTE("---")
call write_variable_in_CASE_NOTE(worker_signature)

'Navigates to TIKL (if selected)
If tikl_for_ten_day_follow_up_checkbox = checked then
	call navigate_to_MAXIS_screen("DAIL", "WRIT")
	call create_MAXIS_friendly_date(date, 10, 5, 18)
	call write_variable_in_TIKL("Has information on new baby/end of pregnancy been received? If not, consider case closure/take appropriate action.")
	transmit
	PF3
End If

script_end_procedure("Success! The script has case noted the overdue baby info, sent a SPEC/MEMO to the client, and TIKLed for 10-day return of information.")
