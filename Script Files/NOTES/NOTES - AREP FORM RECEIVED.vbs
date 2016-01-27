'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - AREP FORM RECEIVED.vbs"
start_time = timer

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

'Required for stats purposes================================================================================================
STATS_counter = 1         'sets the stats counter to 1
STATS_manualtime = 45    'sets the manual run time
STATS_denomination = "C"  'C is for case
'END OF stats block=========================================================================================================

'DIALOGS--------------------------------------------------------------------------------------------------
BeginDialog AREP_dlg, 0, 0, 226, 135, "AREP Case Note"
  EditBox 60, 5, 100, 15, case_number
  CheckBox 65, 30, 35, 10, "SNAP", SNAP_AREP_check
  CheckBox 105, 30, 50, 10, "Health Care", HC_AREP_check
  CheckBox 160, 30, 30, 10, "Cash", CASH_AREP_check
  EditBox 125, 45, 65, 15, arep_signature_date
  CheckBox 5, 65, 75, 10, "ID on file for AREP?", AREP_ID_check
  CheckBox 5, 80, 215, 10, "TIKL to get new HC form 12 months after date form was signed?", TIKL_check
  EditBox 75, 95, 100, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 70, 115, 50, 15
    CancelButton 125, 115, 50, 15
  Text 5, 25, 55, 20, "Programs Authorized for:"
  Text 5, 100, 65, 10, "Worker Signature:"
  Text 5, 50, 115, 10, "Date form was signed (MM/DD/YY):"
  Text 5, 10, 50, 10, "Case Number:"
EndDialog



'THE SCRIPT--------------------------------------------------------------------------------------------------

'Connects to BlueZone
EMConnect ""
'Calls a MAXIS case number
call MAXIS_case_number_finder(case_number)

'Shows dialog and creates and displays an error message if worker completes things incorrectly. 
Do
	err_msg = ""
	dialog AREP_dlg
	cancel_confirmation
	IF SNAP_AREP_check <> checked AND HC_AREP_check <> checked AND CASH_AREP_check <> checked THEN err_msg = err_msg & "Please select a program" & vbNewLine
	IF isdate(arep_signature_date) = false THEN err_msg = err_msg & "Please enter a valid date for the date the form was signed/valid from." & vbNewLine
	IF case_number = "" THEN err_msg = err_msg & "Please enter a case number." & vbNewLine
	IF worker_signature = "" THEN err_msg = err_msg & "Please enter your worker signature." & vbNewLine
	IF (TIKL_check = checked AND arep_signature_date = "") THEN err_msg = err_msg & "You have requested the script to TIKL based on the signature date but you did not enter the signature date." & vbNewLine
	IF err_msg <> "" THEN msgbox err_msg
Loop until err_msg = ""													


'checking for an active MAXIS session
Call check_for_MAXIS(False)

'formatting programs into one variable to write in case note
IF SNAP_AREP_check = checked THEN AREP_programs = "SNAP "
IF HC_AREP_check = checked THEN AREP_programs = AREP_programs & "HC "
IF CASH_AREP_check = checked THEN AREP_programs = AREP_programs & "CASH "


'THE CASE NOTE----------------------------------------------------------------------------------------------------
'Writes a new line, then writes each additional line if there's data in the dialog's edit box (uses if/then statement to decide).
start_a_blank_CASE_NOTE
call write_variable_in_case_note("---AREP FORM RECEIVED---")
call write_Bullet_and_variable_in_case_note("Programs Authorized for", AREP_programs)
call write_Bullet_and_variable_in_case_note("AREP valid start date", arep_signature_date)
call write_variable_in_case_note("* Client and AREP signed AREP form.")
IF AREP_ID_check = checked THEN write_variable_in_CASE_NOTE("* AREP ID on file.")
IF TIKL_check = checked THEN write_variable_in_CASE_NOTE("* TIKL'd for 12 months to get new HC AREP form.")
Call write_variable_in_case_note("---")
call write_variable_in_case_note("* Please see AREP panel to check if AREP is still current and active. This case note does not take the place of an AREP panel.")
Call write_variable_in_case_note("---")
call write_variable_in_CASE_NOTE(worker_signature)

'THE TIKL----------------------------------------------------------------------------------------------------
'If TIKL_check isn't checked this is the end
If TIKL_check = checked then
	call navigate_to_MAXIS_screen("dail", "writ")
	call create_MAXIS_friendly_date(dateadd("m", 12, arep_signature_date), 0, 5, 18)
	call write_variable_in_TIKL("Client's AREP release for HC is now 12 months old and no longer valid. Take appropriate action.")
End If

'Script ends
script_end_procedure("Success! Case note has been added. If you selected to add a TIKL you can edit the TIKL now if needed.")
