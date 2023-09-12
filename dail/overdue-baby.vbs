'Required for statistical purposes==========================================================================================
name_of_script = "DAIL - OVERDUE BABY.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 120                               'manual run time in seconds
STATS_denomination = "C"       'C is for each CASE
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
call changelog_update("05/28/2020", "Update to the notice wording, added virtual drop box information.", "MiKayla Handley, Hennepin County")
call changelog_update("03/01/2020", "Updated TIKL functionality and TIKL text in the case note.", "Ilse Ferris")
CALL changelog_update("04/26/2019", "Updated for DAIL function.", "MiKayla Handley, Hennepin County")
CALL changelog_update("12/29/2017", "Coordinates for sending MEMO's has changed in SPEC function. Updated script to support change.", "Ilse Ferris, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'PEPR UNBORN CHILD IS OVERDUE
'THE SCRIPT------------------------------------------------------------------------------------------------------------------
'Connects to BlueZone default screen
EMConnect ""

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 166, 155, "DAIL_type & MESSAGE PROCESSED"
  CheckBox 5, 35, 90, 10, "ECF has been reviewed ", ECF_reviewed
  CheckBox 5, 50, 90, 10, "Updated PREG panel", update_preg_CHECKBOX
  CheckBox 5, 65, 75, 10, "Send SPEC/MEMO", spec_memo_CHECKBOX
  CheckBox 5, 80, 90, 10, "TIKL for ten day review", tikl_CHECKBOX
  EditBox 50, 95, 110, 15, other_notes
  EditBox 50, 115, 110, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 75, 135, 40, 15
    CancelButton 120, 135, 40, 15
  Text 5, 100, 40, 10, "Other notes:"
  Text 10, 15, 160, 10, full_message
  GroupBox 5, 5, 155, 25, "DAIL for case # " & MAXIS_case_number
  Text 5, 120, 35, 10, "Signature:"
EndDialog

Do
	Do
		err_msg = ""
        Dialog Dialog1
        Cancel_without_confirmation
        If worker_signature = "" then err_msg = err_msg & "You did not sign your case note. Please try again."
	LOOP UNTIL err_msg = ""									'loops until all errors are resolved
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'Call create_TIKL(TIKL_text, num_of_days, date_to_start, ten_day_adjust, TIKL_note_text)
If tikl_checkbox = CHECKED then Call create_TIKL("Has information on new baby/end of pregnancy been received? If not, consider case closure/take appropriate action.", 10, date, False, TIKL_note_text)

IF spec_memo_CHECKBOX = CHECKED THEN
	CALL start_a_new_spec_memo(memo_opened, True, forms_to_arep, forms_to_swkr, send_to_other, other_name, other_street, other_city, other_state, other_zip, True)	'navigates to spec/memo and opens into edit mode

	CALL write_variable_in_SPEC_MEMO("Our records indicate your due date has passed and you did not report the birth of your child or the pregnancy end date. Please contact us within 10 days of this notice with the following information or your case may close:")
    CALL write_variable_in_SPEC_MEMO("")
    CALL write_variable_in_SPEC_MEMO("* Date of the birth or pregnancy end date.")
    CALL write_variable_in_SPEC_MEMO("* Baby's sex and full name.")
    CALL write_variable_in_SPEC_MEMO("* Baby's social security number.")
    CALL write_variable_in_SPEC_MEMO("* Full name of the baby's father.")
    CALL write_variable_in_SPEC_MEMO("* Does the baby's father live in your home?")
    CALL write_variable_in_SPEC_MEMO("* If so, does the father have a source of income?")
    CALL write_variable_in_SPEC_MEMO("  (If so, what is the source of income?)")
    CALL write_variable_in_SPEC_MEMO("* Is there other health insurance available through any       household member's employer, or privately?")
	CALL write_variable_in_SPEC_MEMO("")
	CALL write_variable_in_SPEC_MEMO("*** Submitting Documents:")
	CALL write_variable_in_SPEC_MEMO("- Online at infokeep.hennepin.us or MNBenefits.mn.gov")
	CALL write_variable_in_SPEC_MEMO("  Use InfoKeep to upload documents directly to your case.")
	CALL write_variable_in_SPEC_MEMO("- Mail, Fax, or Drop Boxes at Service Centers.")
	CALL write_variable_in_SPEC_MEMO("  More Info: https://www.hennepin.us/economic-supports")
    CALL write_variable_in_SPEC_MEMO("")
    CALL write_variable_in_SPEC_MEMO("Thank you.")
    PF4
END IF

'Navigates to blank case note
Call navigate_to_MAXIS_screen("CASE", "NOTE")
PF9 'edit mode
call write_variable_in_CASE_NOTE("=== OVERDUE BABY DAIL PROCESSED ===")
IF update_preg_CHECKBOX = CHECKED THEN CALL write_variable_in_CASE_NOTE("* Updated PREG panel, child has been added to case. No further action taken.")
IF spec_memo_CHECKBOX = CHECKED THEN CALL write_variable_in_CASE_NOTE("* SPEC/MEMO sent informing client that they need to report information regarding the birth of their child, and/or pregnancy end date, within 10 days or their case may close.")
CALL write_bullet_and_variable_in_case_note("Other Notes", other_notes)
CALL write_variable_in_CASE_NOTE("---")
CALL write_variable_in_CASE_NOTE(worker_signature)

script_end_procedure_with_error_report("Success! The script has case noted the overdue baby info.")
