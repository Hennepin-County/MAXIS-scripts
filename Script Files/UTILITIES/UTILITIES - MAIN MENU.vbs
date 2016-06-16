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
BeginDialog UTILITIES_scripts_main_menu_dialog, 0, 0, 461, 175, "Utilities scripts main menu dialog"
  ButtonGroup ButtonPressed
    CancelButton 405, 155, 50, 15
    PushButton 30, 20, 95, 10, "Banked Month DB Updater", banked_month_database_updater_button
    PushButton 10, 35, 105, 10, "Copy CASE/NOTE Elsewhere", COPY_TO_CLAIM_button
    PushButton 85, 50, 40, 10, "INFO", INFO_button
    PushButton 5, 65, 120, 10, "Move Production Screen to Inquiry", MOVE_PRODUCTION_SCREEN_TO_INQUIRY_button
    PushButton 5, 80, 120, 10, "Phone Number or Name Look Up", PHONE_NUMBER_OR_NAME_LOOK_UP_button
    PushButton 60, 100, 65, 10, "POLI TEMP List", POLI_TEMP_LIST_button
    PushButton 45, 115, 80, 10, "PRISM Screen Finder", PRISM_SCREENFINDER_button
    PushButton 45, 130, 80, 10, "Training Case Creator", TRAINING_CASE_CREATOR_button
    PushButton 30, 145, 95, 10, "Update Worker Signature", UPDATE_WORKER_SIGNATURE_button
    PushButton 385, 5, 70, 10, "SIR instructions", SIR_instructions_button
  Text 130, 100, 260, 10, "-- Creates a list of current POLI/TEMP topics, TEMP reference and revised date."
  Text 130, 130, 300, 10, "-- Creates training case scenarios en masse and XFERs them to workers."
  Text 130, 145, 195, 10, "-- Sets or updates the default worker signature for this user."
  Text 130, 20, 245, 10, "-- Updates cases in the banked month database with actual MAXIS status."
  Text 5, 5, 250, 10, "Utilities scripts main menu: select the script to run from the choices below."
  Text 130, 50, 215, 10, "-- Displays information about your BlueZone Scripts installation."
  Text 130, 65, 220, 10, "-- Moves a screen from MAXIS prouduction mode to MAXIS inquiry."
  Text 130, 80, 320, 20, "-- Checks every case on PND1, PND2, ACTV, REVW, or INAC, to find a case number when you have a phone number. *OR* Searches for a specific case on multiple REPT screens by last name."
  Text 130, 115, 310, 10, "-- Navigates to popular PRISM screens. The navigation window stays open until user closes it."
  Text 130, 35, 245, 10, "-- Copies a case note to a claim note or a MEMO"
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
IF buttonpressed = COPY_TO_CLAIM_button					 		then call run_from_GitHub(script_repository & "/UTILITIES/UTILITIES - COPY CASE NOTE ELSEWHERE.vbs")
IF buttonpressed = INFO_button 									then call run_from_GitHub(script_repository & "/UTILITIES/UTILITIES - INFO.vbs")
If buttonpressed = MOVE_PRODUCTION_SCREEN_TO_INQUIRY_button		then call run_from_GitHub(script_repository & "/UTILITIES/UTILITIES - MOVE PRODUCTION SCREEN TO INQUIRY.vbs")
IF ButtonPressed = PHONE_NUMBER_OR_NAME_LOOK_UP_button			then call run_from_GitHub(script_repository & "/UTILITIES/UTILITIES - PHONE NUMBER OR NAME LOOK UP.vbs")
IF buttonpressed = POLI_TEMP_LIST_button						then call run_from_GitHub(script_repository & "/UTILITIES/UTILITIES - POLI TEMP LIST.vbs")
IF buttonpressed = PRISM_SCREENFINDER_button                    then call run_from_GitHub(script_repository & "/UTILITIES/UTILITIES - PRISM SCREEN FINDER.vbs")         
IF buttonpressed = TRAINING_CASE_CREATOR_button 				then call run_from_GitHub(script_repository & "/UTILITIES/UTILITIES - TRAINING CASE CREATOR.vbs")
IF buttonpressed = UPDATE_WORKER_SIGNATURE_button				then call run_from_GitHub(script_repository & "/UTILITIES/UTILITIES - UPDATE WORKER SIGNATURE.vbs")

'Logging usage stats
script_end_procedure("If you see this, it's because you clicked a button that, for some reason, does not have an outcome in the script. Contact your alpha user to report this bug. Thank you!")
