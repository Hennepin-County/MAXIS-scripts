'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "NAV - PRISM SCREEN FINDER.vbs"
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

'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog PRISM_screen_finder_dialog, 0, 0, 261, 135, "PRISM screen finder"
  ButtonGroup ButtonPressed
    CancelButton 210, 120, 50, 15
    PushButton 140, 70, 45, 10, "DDPL", DDPL_button
    PushButton 140, 40, 45, 10, "CAAD", CAAD_button
    PushButton 140, 55, 45, 10, "CAFS", CAFS_button
    PushButton 140, 85, 45, 10, "GCSC", GCSC_button
    PushButton 140, 115, 45, 10, "PESE", PESE_button
 
  Text 35, 70, 90, 10, "Direct deposit listing:"
  Text 35, 40, 65, 10, "Case notes:"
  Text 35, 55, 100, 10, "Case financial summary:"
  Text 35, 85, 100, 10, "Good cause/safety concerns:"
  Text 35, 115, 65, 10, "Person search:"
  Text 10, 0, 250, 25, "Press a button below to navigate to PRISM screens.  Then press F1 in the case number or MCI number field to select the participant or case information you are looking for."
EndDialog


'THE SCRIPT----------------------------------------------------------------------------------------------------

'Connect to BlueZone
EMConnect ""

CALL check_for_PRISM(FALSE)

DO

	Dialog PRISM_screen_finder_dialog

	'Now it'll navigate to any of the screens chosen
	If buttonpressed = DDPL_button then call navigate_to_PRISM_screen("DDPL")
	If buttonpressed = CAAD_button then call navigate_to_PRISM_screen("CAAD")
	If buttonpressed = CAFS_button then call navigate_to_PRISM_screen("CAFS")
	If buttonpressed = GCSC_button then call navigate_to_PRISM_screen("GCSC")
	If buttonpressed = PESE_button then call navigate_to_PRISM_screen("PESE")
LOOP until buttonpressed = 0

script_end_procedure("")


