'Required for statistical purposes===============================================================================
name_of_script = "ADMIN - AVS SUBMITTED.vbs"
start_time = timer
STATS_counter = 1              'sets the stats counter at one
STATS_manualtime = 90         'manual run time in seconds
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
call changelog_update("02/20/2020", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'=======================================================================================================END CHANGELOG BLOCK
'THE SCRIPT
'GRABBING THE CASE NUMBER, THE MEMB NUMBERS, AND THE FOOTER MONTH------------------------------------------------------------------------------------------------------------------------------------------------
EMConnect ""
Call maxis_case_number_finder(MAXIS_case_number)

'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 311, 110, "AVS Submitted"
  EditBox 65, 45, 45, 15, maxis_case_number
  EditBox 65, 65, 240, 15, other_notes
  EditBox 65, 85, 150, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 220, 85, 40, 15
    CancelButton 265, 85, 40, 15
  Text 20, 70, 45, 10, "Other Notes:"
  Text 5, 90, 60, 10, "Worker signature:"
  Text 15, 50, 50, 10, "Case Number:"
  GroupBox 5, 5, 300, 35, "Using the script:"
  Text 10, 15, 290, 20, "This script will case note and set a TIKL for 10-day return when an AVS request is sent. A DHS -7823 must be on file. For applicants OR change in basis residents only."
EndDialog

Do
    Do
        err_msg = ""
		Dialog Dialog1
		cancel_without_confirmation
		IF IsNumeric(maxis_case_number) = false or len(maxis_case_number) > 8 THEN err_msg = err_msg & vbNewLine & "* Please enter a valid case number."
		If trim(worker_signature) = "" then err_msg = err_msg & vbcr & "* Please ensure your case note is signed."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
	LOOP UNTIL err_msg = ""									'loops until all errors are resolved
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'create_TIKL(TIKL_text, num_of_days, date_to_start, ten_day_adjust, TIKL_note_text)
Call create_TIKL("AVS 10-day follow up is required. If you do not have access to the AVS system, QI can assist you with the results. Contact your supervisor about your AVS access.", 10, date, False, TIKL_note_text)

start_a_blank_case_note
Call write_variable_in_case_note("***AVS Request Submitted - 10 day follow-up needed***")
Call write_variable_in_case_note(TIKL_note_text)
CALL write_bullet_and_variable_in_case_note("Other Notes", other_notes)
CALL write_variable_in_CASE_NOTE("---")
CALL write_variable_in_CASE_NOTE(worker_signature)
PF3

script_end_procedure("Success! Your case note and TIKL have been created.")