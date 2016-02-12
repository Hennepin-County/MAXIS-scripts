'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - MSQ.vbs"
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

'Required for statistical purposes===========================================================================================
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 300          'manual run time in seconds
STATS_denomination = "C"        'C is for each case
'END OF stats block==========================================================================================================
 
'THE DIALOG----------------------------------------------------------------------------------------------------------
BeginDialog msq_dialog, 0, 0, 321, 125, "MSQ"
  EditBox 80, 5, 70, 15, case_number
  EditBox 75, 30, 70, 15, member_injured
  EditBox 205, 30, 70, 15, injury_date
  EditBox 75, 65, 175, 15, other_notes
  EditBox 80, 95, 80, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 200, 95, 50, 15
    CancelButton 255, 95, 50, 15
  Text 5, 70, 70, 10, "Action Taken/Notes:"
  Text 165, 35, 40, 10, "Injury Date:"
  Text 5, 35, 70, 10, "HH Member Injured:"
  Text 5, 100, 70, 10, "Sign your Case Note:"
  Text 5, 10, 70, 10, "Maxis Case Number:"
  Text 75, 45, 40, 10, "(Ex: 01, 02)"
  Text 205, 45, 70, 10, "(Ex: MM/DD/YY)"
EndDialog


'THE SCRIPT--------------------------------------------------------------------------------------------------------------

'Connects to BLUEZONE
EMConnect ""

'Grabs the MAXIS case number            
CALL MAXIS_case_number_finder(case_number)

'Shows dialog
DO
	err_msg = ""		
	Dialog msq_dialog
		IF ButtonPressed = 0 THEN StopScript
		IF IsNumeric(case_number) = FALSE THEN err_msg = err_msg & vbCr & "* You must type a valid numeric case number."
		IF worker_signature = "" THEN err_msg = err_msg & vbCr & "* You must sign your case note!"
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
LOOP UNTIL err_msg = ""

'Checks Maxis for password prompt
CALL check_for_MAXIS(True)

'The script reads what member number was manually entered, and navigates to that member's STAT/ACCI panel
CALL navigate_to_MAXIS_screen("STAT", "ACCI")
EMWriteScreen member_injured, 20, 76
EMWriteScreen "nn", 20, 79
transmit

EMWriteScreen "n", 8, 75

'Writes 13 in Accident Type field
EMWriteScreen "13", 6, 47

'Writes the Injury Date in the Injury date field
CALL create_MAXIS_friendly_date(injury_date, 0, 6, 73)

'Writes N in the Med Cooperation field
EMWriteScreen "N", 7, 47

'Writes N in the Good cause field
EMWriteScreen "N", 7, 73

'Writes a N in Pend Litigation
EMWritescreen "N", 9, 47

'Opens new case note
start_a_blank_case_note


'Writes the Case Note
CALL write_variable_in_case_note("*** MSQ Form ***")
CALL write_bullet_and_variable_in_case_note("Household Member Injured", member_injured)
CALL write_bullet_and_variable_in_case_note("Injury Date", injury_date)
CALL write_bullet_and_variable_in_CASE_NOTE("Actions Taken/Notes", other_notes)
CALL write_variable_in_case_note("---")
CALL write_variable_in_case_note(worker_signature)

script_end_procedure("Success! Remember to update MMIS.")
