'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - MFIP SANCTION CURED.vbs"
start_time = timer

'DIM name_of_script, start_time, FuncLib_URL, run_locally, default_directory, beta_agency, req, fso

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN		'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else																		'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		End if
		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
		req.send													'Sends request
		IF req.Status = 200 THEN									'200 means great success
			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
			Execute req.responseText								'Executes the script code
		ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
			MsgBox 	"Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_
					vbCr & _
					"Before contacting Veronica Cary, please check to make sure you can load the main page at www.GitHub.com." & vbCr &_
					vbCr & _
					"If you can reach GitHub.com, but this script still does not work, ask an alpha user to contact Veronica Cary and provide the following information:" & vbCr &_
					vbTab & "- The name of the script you are running." & vbCr &_
					vbTab & "- Whether or not the script is ""erroring out"" for any other users." & vbCr &_
					vbTab & "- The name and email for an employee from your IT department," & vbCr & _
					vbTab & vbTab & "responsible for network issues." & vbCr &_
					vbTab & "- The URL indicated below (a screenshot should suffice)." & vbCr &_
					vbCr & _
					"Veronica will work with your IT department to try and solve this issue, if needed." & vbCr &_
					vbCr &_
					"URL: " & FuncLib_URL
					script_end_procedure("Script ended due to error connecting to GitHub.")
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

'Required for statistical purposes==========================================================================================
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 90           'manual run time in seconds
STATS_denomination = "C"        'C is for each case
'END OF stats block=========================================================================================================

'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog MFIP_sanction_cured_dialog, 0, 0, 396, 190, "MFIP Sanction Cured"
  EditBox 90, 5, 85, 15, case_number
  EditBox 90, 25, 85, 15, sanction_lifted_month
  EditBox 315, 25, 70, 15, compliance_date
  DropListBox 90, 50, 215, 15, "Select One..."+chr(9)+"Client complied with Employment Services"+chr(9)+"Client complied with Child Support"+chr(9)+"Client complied with Employment Services AND Child Support ", cured_reason
  EditBox 90, 75, 295, 15, action_taken
  DropListBox 90, 100, 85, 20, "Select One..."+chr(9)+"Letter"+chr(9)+"Phone Call"+chr(9)+"Email"+chr(9)+"Client Not Notified", notified_via
  EditBox 90, 120, 295, 15, other_notes
  EditBox 90, 145, 100, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 280, 165, 50, 15
    CancelButton 335, 165, 50, 15
  Text 195, 30, 115, 10, "Date Client Came into Compliance:"
  Text 40, 80, 50, 15, "Action Taken:"
  Text 25, 100, 70, 10, "Notified Client Via:"
  Text 10, 30, 75, 10, "Month Sanction Lifted:"
  Text 15, 150, 70, 10, "Sign Your Case Note:"
  Text 5, 50, 80, 10, "Sanction Cured Reason:"
  Text 5, 125, 80, 10, "Other Notes/Comments:"
  Text 15, 10, 70, 10, "Maxis Case Number:"
EndDialog


'THE SCRIPT--------------------------------------------------------------------------------------------------
'Connect to Bluezone
EMConnect ""

'Grabs Maxis Case number
CALL MAXIS_case_number_finder(case_number)

'Shows dialog
DO
	DO
		DO
			Dialog MFIP_sanction_cured_dialog
			IF ButtonPressed = 0 THEN StopScript
			IF worker_signature = "" THEN MsgBox "You must sign your case note!"
			LOOP UNTIL worker_signature <> ""
		IF IsNumeric(case_number) = FALSE THEN MsgBox "You must type a valid numeric case number."
	IF cured_reason = "Select One..." THEN MsgBox "You must select 'Reason for Sanction being Cured!'"
	IF notified_via = "Select One..." THEN MsgBox "You must select 'Notified Client Via!'"
	LOOP UNTIL cured_reason <> "Select One..."
LOOP UNTIL IsNumeric(case_number) = TRUE
	
'Checks Maxis for password prompt
CALL check_for_MAXIS(True)

'Navigates to case note
CALL navigate_to_screen("CASE", "NOTE")

'Sends a PF9
PF9

'Writes the case note

CALL write_variable_in_case_note ("~~$~~MFIP SANCTION CURED~~$~~")                                         'Writes title in Case note
CALL write_bullet_and_variable_in_case_note("Month Sanction Cured", sanction_lifted_month)                 'Writes Month the Sanction was lifted
CALL write_bullet_and_variable_in_case_note("Client Came Into Compliance On", compliance_date)             'Writes the Date the Client came into Compliance
CALL write_bullet_and_variable_in_case_note("Sanction Cured Reason", cured_reason)                         'Writes the reason why the sanction was cured
CALL write_bullet_and_variable_in_case_note("Actions Taken", action_taken)                                 'Writes any actions taken
CALL write_bullet_and_variable_in_case_note("Client was notified Via", notified_via)                       'Writes the way the client was notified that their sanction was lifted
CALL write_bullet_and_variable_in_case_note("Other Notes/Comments", other_notes)                           'Writes any other notes/comments
CALL write_variable_in_case_note ("---")   
CALL write_variable_in_CASE_NOTE(worker_signature)                                                         'Writes worker signature in note

CALL script_end_procedure("")
