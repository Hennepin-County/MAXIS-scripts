'GATHERING STATS===========================================================================================
name_of_script = "NOTES - DEU-APPEAL SUMMARY COMPLETED.vbs"
start_time = timer
STATS_counter = 1
STATS_manualtime = 60
STATS_denominatinon = "C"
'END OF STATS BLOCK===========================================================================================

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
call run_from_GitHub(script_repository & "notes/appeals.vbs")
'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
CALL changelog_update("09/16/2022", "Update to ensure Worker Signature is in all scripts that CASE/NOTE.", "MiKayla Handley, Hennepin County") '#316
call changelog_update("04/13/2017", "Initial version.", "MiKayla Handley, Hennepin County")
'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display

'END CHANGELOG BLOCK =======================================================================================================----------
'connecting to BlueZone and grabbing the case number
EMConnect ""
Call MAXIS_case_number_finder(maxis_case_number)
'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 276, 90, "Appeal Summary Completed"
  EditBox 55, 5, 65, 15, maxis_case_number
  EditBox 210, 5, 60, 15, date_appeal_rcvd
  EditBox 55, 25, 65, 15, claim_number
  EditBox 210, 25, 60, 15, effective_date
  EditBox 95, 45, 175, 15, action_client_is_appealing
  EditBox 70, 65, 90, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 165, 65, 50, 15
    CancelButton 220, 65, 50, 15
  Text 5, 10, 45, 10, "Case number:"
  Text 130, 10, 75, 10, "Date appeal received:"
  Text 5, 30, 50, 10, "Claim number:"
  Text 130, 30, 80, 10, "Effective date of action:"
  Text 5, 50, 85, 10, "Action client is appealing:"
  Text 5, 70, 60, 10, "Worker signature:"
EndDialog

Do
	Do
        err_msg = ""
		Dialog Dialog1
		cancel_confirmation
		IF IsNumeric(maxis_case_number) = false or len(maxis_case_number) > 8 THEN err_msg = err_msg & vbNewLine & "* Please enter a valid case number."
		IF Isdate(date_appeal_rcvd) = false THEN err_msg = err_msg & vbNewLine & "* Please enter a date for the appeal."
		IF IsNumeric(claim_number) = false THEN err_msg = err_msg & vbNewLine & "* Please enter a valid claim number."
		IF Isdate(effective_date) = false THEN err_msg = err_msg & vbNewLine & "* Please enter the effective date."
		IF action_client_is_appealing = "" THEN err_msg = err_msg & vbNewLine & "* Please enter action that client is appealing."
		IF worker_signature = "" THEN err_msg = err_msg & vbCr & "* Please sign your case note."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
    Loop until err_msg = ""
 	Call check_for_password(are_we_passworded_out)
LOOP UNTIL check_for_password(are_we_passworded_out) = False

start_a_blank_case_note      'navigates to case/note and puts case/note into edit mode
Call write_variable_in_CASE_NOTE("-----APPEAL SUMMARY COMPLETED-----")
Call write_bullet_and_variable_in_CASE_NOTE("Claim number:", claim_number)
Call write_bullet_and_variable_in_CASE_NOTE("Date appeal request received:", date_appeal_rcvd)
Call write_bullet_and_variable_in_CASE_NOTE("Effective date of action being appealed:", effective_date)
Call write_bullet_and_variable_in_CASE_NOTE("Action client is appealing:", action_client_is_appealing)
Call write_variable_in_CASE_NOTE("----- ----- ----- ----- ----- ----- -----")
CALL write_variable_in_CASE_NOTE(worker_signature)
Call write_variable_in_CASE_NOTE("DEBT ESTABLISHMENT UNIT 612-348-4290 PROMPTS 1-1-1")

script_end_procedure_with_error_report("Appeal Summary case note complete.")
