'Required for statistical purposes===============================================================================
name_of_script = "NAV - OTHER NAV MAIN MENU.vbs"
start_time = timer

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

'DIALOG----------------------------------------------------------------------------------------------------
BeginDialog OTHER_NAV_scripts_main_menu_dialog, 0, 0, 456, 65, "Other NAV scripts main menu dialog"
  ButtonGroup ButtonPressed
    CancelButton 400, 120, 50, 20
    PushButton 370, 5, 70, 10, "SIR instructions", SIR_instructions_button
    PushButton 5, 15, 100, 10, "Look up MAXIS case in MMIS", 		LOOK_UP_MAXIS_CASE_IN_MMIS_button
    PushButton 5, 30, 100, 10, "Look up MMIS PMI in MAXIS", 		LOOK_UP_MMIS_PMI_IN_MAXIS_button
    PushButton 5, 45, 40, 10, "View INFC", 					VIEW_INFC_button
  Text 0, 0, 250, 10, "Other nav scripts main menu: select the script to run from the choices below."
  Text 110, 15, 250, 10, "--- Navigates to RELG in MMIS for a selected case. Navigates to person 01."
  Text 110, 30, 140, 10, "--- Jumps from MMIS to MAXIS for a case."
  Text 50, 45, 400, 10, "--- Views an INFC panel for a case."
EndDialog

'Variables to declare
IF script_repository = "" THEN script_repository = "https://raw.githubusercontent.com/MN-Script-Team/DHS-MAXIS-Scripts/master/Script Files"		'If it's blank, we're assuming the user is a scriptwriter, ergo, master branch.

'THE SCRIPT----------------------------------------------------------------------------------------------------

'Shows main menu dialog, which asks user which script to run. Loops until a button other than the SIR instructions button is clicked.
Do
	dialog OTHER_NAV_scripts_main_menu_dialog
	If buttonpressed = cancel then stopscript
	If buttonpressed = SIR_instructions_button then CreateObject("WScript.Shell").Run("https://www.dhssir.cty.dhs.state.mn.us/MAXIS/blzn/scriptwiki/Wiki%20Pages/Navigation%20scripts.aspx")
Loop until buttonpressed <> SIR_instructions_button

'Connecting to BlueZone
EMConnect ""

If buttonpressed = LOOK_UP_MAXIS_CASE_IN_MMIS_button						then call run_from_GitHub(script_repository & "/NAV/NAV - FIND MAXIS CASE IN MMIS.vbs")
If buttonpressed = LOOK_UP_MMIS_PMI_IN_MAXIS_button						then call run_from_GitHub(script_repository & "/NAV/NAV - FIND MMIS PMI IN MAXIS.vbs")
If buttonpressed = VIEW_INFC_button									then call run_from_GitHub(script_repository & "/NAV/NAV - VIEW INFC.vbs")

'Logging usage stats
script_end_procedure("If you see this, it's because you clicked a button that, for some reason, does not have an outcome in the script. Contact your alpha user to report this bug. Thank you!")
