'Required for statistical purposes===============================================================================
name_of_script = "DAIL - MEDI MESSAGE.vbs"
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

'===========================================================================================================CHANGELOG BLOCK
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("05/01/2019", "Initial version.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'=======================================================================================================END CHANGELOG BLOCK
'THE MAIN DIALOG--------------------------------------------------------------------------------------------------
BeginDialog medi_dialog, 0, 0, 266, 150, "DAIL_type & MESSAGE PROCESSED"
  CheckBox 5, 40, 140, 10, "Client is eligible for the Medicare buy-in", medi_checkbox
  EditBox 210, 35, 50, 15, ELIG_date
  Text 5, 60, 195, 10, "If INELIG year that client will be eligible for Medicare Buy-In"
  EditBox 210, 55, 50, 15, ELIG_year
  CheckBox 5, 75, 110, 10, "Forms have been sent in ECF", ECF_sent
  EditBox 50, 90, 210, 15, other_notes
  EditBox 70, 110, 95, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 175, 130, 40, 15
    CancelButton 220, 130, 40, 15
  GroupBox 5, 5, 270, 25, "DAIL for case #  &  MAXIS_case_number"
  Text 10, 15, 260, 10, full_message
  Text 5, 95, 45, 10, "Other notes:"
  Text 5, 115, 60, 10, "Worker signature:"
  Text 175, 40, 35, 10, "ELIG date"
EndDialog

EMWriteScreen "N", 6, 3         'Goes to Case Note - maintains tie with DAIL
TRANSMIT
'Starts a blank case note
PF9
EMReadScreen case_note_mode_check, 7, 20, 3
If case_note_mode_check <> "Mode: A" then script_end_procedure("You are not in a case note on edit mode. You might be in inquiry. Try the script again in production.")

Do
    Do
        err_msg = ""
		Dialog medi_dialog
		cancel_confirmation
        If (isnumeric(MAXIS_case_number) = False and len(MAXIS_case_number) <> 8) then err_msg = err_msg & vbcr & "* Enter a valid case number."
		If trim(worker_signature) = "" then err_msg = err_msg & vbcr & "* Please ensure your case note is signed."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
	LOOP UNTIL err_msg = ""									'loops until all errors are resolved
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

due_date = dateadd("d", 30, ELIG_date)

'start_a_blank_case_note
CALL write_variable_in_CASE_NOTE("=== " & DAIL_type & " - MESSAGE PROCESSED " & "===")
CALL write_variable_in_case_note("* " & full_message)
CALL write_variable_in_case_note(first_line)
CALL write_variable_in_case_note(second_line)
CALL write_variable_in_case_note(third_line)
CALL write_variable_in_case_note(fourth_line)
CALL write_variable_in_case_note(fifth_line)
CALL write_variable_in_case_note("---")
IF medi_checkbox = CHEKED THEN
	Call write_variable_in_case_note("** Medicare Buy-in Referral mailed **")
	Call write_variable_in_case_note("Client is eligible for the Medicare buy-in as of " & ELIG_date & ". Proof due by " & due_date & "to apply.")
	Call write_variable_in_case_note("Mailed DHS-3439-ENG MHCP Medicare Buy-In Referral Letter - TIKL set to follow up.")
ELSE
	Call write_variable_in_case_note("** Medicare Referral **")
	Call write_variable_in_case_note("Client is not eligible for the Medicare buy-in. Enrollment is not until January " & ELIG_year & ", unable	to apply until the enrollment time.")
	Call write_variable_in_case_note("TIKL set to mail the Medicare Referral for November " & ELIG_year & ".")
END IF
IF ECF_sent = CHECKED THEN CALL write_variable_in_case_note("* ECF reviewed and appropriate action taken")
CALL write_bullet_and_variable_in_case_note("Other notes", other_notes)
CALL write_variable_in_CASE_NOTE("---")
CALL write_variable_in_CASE_NOTE(worker_signature)
PF3

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
