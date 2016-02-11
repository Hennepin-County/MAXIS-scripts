'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "UTILITIES - MAIN MENU.vbs"
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
BeginDialog UTILITIES_scripts_main_menu_dialog, 0, 0, 461, 85, "Utilities scripts main menu dialog"
  ButtonGroup ButtonPressed
    CancelButton 405, 65, 50, 15
    PushButton 5, 20, 95, 10, "Banked Month DB Updater", banked_month_database_updater_button
    PushButton 60, 35, 40, 10, "INFO", INFO_button
    PushButton 5, 50, 95, 10, "Update Worker Signature", UPDATE_WORKER_SIGNATURE_button
    PushButton 385, 5, 70, 10, "SIR instructions", SIR_instructions_button
  Text 5, 5, 250, 10, "Utilities scripts main menu: select the script to run from the choices below."
  Text 105, 20, 305, 10, "-- NEW 02/2016!!! Updates cases in the banked month database with actual MAXIS status."
  Text 105, 35, 265, 10, "-- NEW 01/2016!!! Displays information about your BlueZone Scripts installation."
  Text 105, 50, 195, 10, "-- Sets or updates the default worker signature for this user."
EndDialog

'Variables to declare
IF script_repository = "" THEN script_repository = "https://raw.githubusercontent.com/MN-Script-Team/DHS-MAXIS-Scripts/master/Script Files"		'If it's blank, we're assuming the user is a scriptwriter, ergo, master branch.

'THE SCRIPT----------------------------------------------------------------------------------------------------
'Shows dialog, which asks user which script to run.
Do
	dialog UTILITIES_scripts_main_menu_dialog
	If buttonpressed = cancel then stopscript
	If buttonpressed = SIR_instructions_button then CreateObject("WScript.Shell").Run("https://www.dhssir.cty.dhs.state.mn.us/MAXIS/blzn/Script%20Instructions%20Wiki/Utilities%20scripts.aspx")
Loop until buttonpressed <> SIR_instructions_button

'Connecting to BlueZone
EMConnect ""

IF buttonpressed = banked_month_database_updater_button 		then call run_from_GitHub(script_repository & "/UTILITIES/UTILITIES - BANKED MONTH DATABASE UPDATER.vbs")
IF buttonpressed = INFO_button 									then call run_from_GitHub(script_repository & "/UTILITIES/UTILITIES - INFO.vbs")
IF buttonpressed = UPDATE_WORKER_SIGNATURE_button				then call run_from_GitHub(script_repository & "/UTILITIES/UTILITIES - UPDATE WORKER SIGNATURE.vbs")

'Logging usage stats
script_end_procedure("If you see this, it's because you clicked a button that, for some reason, does not have an outcome in the script. Contact your alpha user to report this bug. Thank you!")
