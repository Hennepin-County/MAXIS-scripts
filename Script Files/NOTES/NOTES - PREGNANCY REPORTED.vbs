'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - PREGNANCY REPORTED.vbs"
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

'Required for statistical purposes==========================================================================================
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 180          'manual run time in seconds
STATS_denomination = "C"        'C is for each case
'END OF stats block=========================================================================================================

'THE DIALOG--------------------------------------------------------------------------------------------------
BeginDialog Dialog1, 0, 0, 351, 185, "Pregnancy Reported"
  EditBox 95, 5, 80, 15, maxis_case_number
  EditBox 95, 25, 80, 15, member_preg
  EditBox 260, 25, 70, 15, due_date
  DropListBox 95, 60, 95, 15, "Select One..."+chr(9)+"Self Attestation"+chr(9)+"Change Report Form"+chr(9)+"Pregnancy Verification Form"+chr(9)+"Renewal Form"+chr(9)+"Other", report_method
  EditBox 95, 80, 235, 15, other_notes
  CheckBox 35, 120, 25, 15, "MA", ma_checkbox
  CheckBox 85, 120, 35, 15, "CASH", cash_checkbox
  CheckBox 190, 110, 70, 10, "Updated in MMIS", mmis_checkbox
  CheckBox 190, 130, 125, 10, "Verification Request sent for CASH", verification_checkbox
  EditBox 90, 155, 120, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 240, 155, 50, 15
    CancelButton 295, 155, 50, 15
  Text 15, 85, 80, 10, "Other Comments/Notes:"
  Text 15, 30, 75, 10, "HH Member Pregnant:"
  Text 20, 10, 70, 10, "Maxis Case Number:"
  Text 10, 60, 85, 15, "Pregnancy Reported Via:"
  Text 265, 40, 75, 10, "Example:  MM/DD/YY"
  GroupBox 10, 105, 130, 40, "Program Pregnancy Reported For:"
  Text 20, 160, 70, 10, "Sign your Case Note:"
  Text 185, 30, 70, 10, "Pregnancy Due Date:"
  Text 100, 40, 60, 10, "Example: 01, 03"
EndDialog

'THE SCRIPT------------------------------------------------------------------------------------------------------
'Connects to BLUEZONE
EMConnect ""

'Grabs the MAXIS case number
CALL MAXIS_case_number_finder(case_number)

'Shows dialog
DO
	err_msg = ""
	Dialog Dialog1 
		IF ButtonPressed = 0 THEN StopScript
		IF report_method = "Select One..." THEN err_msg = err_msg & vbCr & "* You must select how the pregnancy was reported!"
		IF IsNumeric(case_number) = FALSE THEN err_msg = err_msg & vbCr & "* You must type a valid numeric case number."
		IF due_date = "" OR (due_date <> "" AND IsDate(due_date) = False) THEN err_msg = err_msg & vbCr & "* You must enter a due date in a MM/DD/YY format."
		IF worker_signature = "" THEN err_msg = err_msg & vbCr & "* You must sign your case note!"
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
LOOP UNTIL err_msg = ""

'Checks Maxis for password prompt
CALL check_for_MAXIS(True)

'Script calculates the Conception date based off the due date entered in the dialog box
conception_date = DateAdd("d", -280, due_date)

'The script reads what member number was manually entered, and navigates to that member's stat/preg panel
CALL navigate_to_MAXIS_screen("STAT", "PREG")
EMWriteScreen member_preg, 20, 76
EMWriteScreen "nn", 20, 79
transmit

'Writes the auto-calucated conception date in the Conception Date field and the Due date in that field
CALL create_MAXIS_friendly_date(conception_date, 0, 6, 53)
CALL create_MAXIS_friendly_date(due_date, 0, 10, 53)

EMWriteScreen "n", 8, 75

'If under Program Pregnancy applied for, FW has check MA or MA/CASH then script will write Y in the Verified field on stat/preg
IF ma_checkbox = checked and cash_checkbox = checked THEN EMWritescreen "Y", 6, 75

'If under Program Pregnancy applied for, FW has checked CASH then script will write N in the Verified field on stat/preg
IF cash_checkbox = checked THEN EMWritescreen "N", 6, 75
transmit

'Opens new case note
start_a_blank_case_note

'Writes the Case Note
CALL write_variable_in_case_note ("---Pregnancy Reported---")
CALL write_bullet_and_variable_in_case_note("Household Member Pregnant", member_preg)
CALL write_bullet_and_variable_in_case_note("Conception Date", conception_date)
CALL write_bullet_and_variable_in_case_note("Pregnancy Due Date", due_date)
CALL write_bullet_and_variable_in_case_note("Pregnancy Reported Via", report_method)
IF ma_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* Program Pregnancy Reported for: MA")         'HAVING TROUBLES STARTING HERE....
IF cash_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* Program Pregnancy Reported for: CASH")
IF ma_checkbox and cash_checkbox = checked THEN CALL write_variable_in_case_note("* Programs Pregnancy Reported for: MA & CASH")
IF mmis_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* Updated in MMIS")
IF verification_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* Sent verification request for CASH")
CALL write_bullet_and_variable_in_CASE_NOTE("Other Comments/Notes", other_notes)
CALL write_variable_in_case_note("---")
CALL write_variable_in_case_note(worker_signature)

script_end_procedure("")
