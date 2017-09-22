'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - REIN PROGS.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 90           'manual run time in seconds
STATS_denomination = "C"        'C is for each case
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

'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog Rein_dialog, 0, 0, 256, 260, "Rein"
  EditBox 80, 5, 60, 15, MAXIS_case_number
  EditBox 80, 25, 60, 15, rein_date
  CheckBox 30, 65, 50, 10, "SNAP", SNAP_checkbox
  CheckBox 90, 65, 50, 10, "CASH", CASH_checkbox
  CheckBox 155, 65, 50, 10, "HC", HC_checkbox
  CheckBox 30, 110, 50, 10, "SNAP", SNAP_rein_checkbox
  CheckBox 90, 110, 50, 10, "CASH", CASH_rein_checkbox
  CheckBox 155, 110, 50, 10, "HC", HC_rein_checkbox
  EditBox 100, 135, 50, 15, Progs_closed_date
  EditBox 100, 160, 115, 15, reason_rein
  EditBox 100, 185, 115, 15, Actions_taken
  EditBox 100, 210, 75, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 125, 235, 50, 15
    CancelButton 190, 235, 50, 15
  Text 30, 30, 45, 10, "Date of REIN:"
  Text 10, 140, 85, 10, "Programs last closed on:"
  Text 45, 185, 50, 10, "Actions Taken:"
  Text 40, 160, 65, 10, "Reason for REIN:"
  Text 10, 10, 75, 10, "Maxis case number:"
  GroupBox 5, 95, 220, 35, "Programs to REIN: "
  Text 35, 215, 65, 10, "Worker Signature:"
  GroupBox 5, 50, 215, 35, "Programs closed:"
EndDialog


'script code-----------------------------------------------------------------------------------------------

'Connect to Bluezone
EMConnect ""

'Grabs Maxis Case number
CALL MAXIS_case_number_finder(MAXIS_case_number)

'Shows dialog
DO
	DO

		Dialog REIN_dialog
		IF ButtonPressed = 0 THEN StopScript
		IF worker_signature = "" THEN MsgBox "You must sign your case note!"
		LOOP UNTIL worker_signature <> ""
	IF IsNumeric(MAXIS_case_number) = FALSE THEN MsgBox "You must type a valid numeric case number."
LOOP UNTIL IsNumeric(MAXIS_case_number) = TRUE


'Checks Maxis for password prompt
CALL check_for_MAXIS(True)


'Navigates to case note
CALL navigate_to_MAXIS_screen("CASE", "NOTE")

'Sends a PF9
PF9

'Writes the case note
CALL write_variable_in_case_note ("***REIN Programs***")
CALL write_bullet_and_variable_in_case_note("Date of REIN", rein_date)
CALL write_variable_in_case_note ("~~~Programs closed~~~")
IF SNAP_checkbox = 1 THEN call write_variable_in_case_note("* SNAP")
IF CASH_checkbox = 1 THEN call write_variable_in_case_note("* CASH")
IF HC_checkbox = 1 THEN call write_variable_in_case_note("* HC")
CALL write_variable_in_case_note ("~~~Programs to REIN~~~")
IF SNAP_REIN_checkbox = 1 THEN call write_variable_in_case_note("* SNAP")
IF CASH_REIN_checkbox = 1 THEN call write_variable_in_case_note("* CASH")
IF HC_REIN_checkbox = 1 THEN call write_variable_in_case_note("* HC")
CALL write_variable_in_case_note ("---")
CALL write_bullet_and_variable_in_case_note("Programs closed on", progs_closed_date)
CALL write_bullet_and_variable_in_case_note("Reason for Reinstatement", reason_rein)
CALL write_bullet_and_variable_in_case_note("Actions Taken", actions_taken)
CALL write_variable_in_case_note ("---")
CALL write_variable_in_case_note (worker_signature)


CALL script_end_procedure("")
