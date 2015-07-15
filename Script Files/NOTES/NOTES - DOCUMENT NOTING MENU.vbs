'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - DOCUMENT NOTING MENU.vbs"
start_time = timer

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN		'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF default_directory = "C:\DHS-MAXIS-Scripts\Script Files\" THEN			'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		ELSEIF beta_agency = "" or beta_agency = True then							'If you're a beta agency, you should probably use the beta branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/BETA/MASTER%20FUNCTIONS%20LIBRARY.vbs"
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
BeginDialog NOTES_document_noting_menu_dialog, 0, 0, 411, 150, "Document noting scripts main menu dialog"
  ButtonGroup ButtonPressed
    CancelButton 355, 130, 50, 15
    PushButton 340, 10, 65, 10, "SIR instructions", SIR_instructions_button
    PushButton 5, 20, 100, 10, "Change Report Form received", CHANGE_REPORT_FORM_RECEIVED_button
    PushButton 5, 35, 45, 10, "LTC - 1503", LTC_1503_button
    PushButton 5, 50, 45, 10, "LTC - 5181", LTC_5181_button
    PushButton 5, 65, 110, 10, "Medical Opinion Form received", MEDICAL_OPINION_FORM_RECEIVED_button
    PushButton 5, 80, 80, 10, "Shelter Form received", SHELTER_FORM_RECEIVED_button
    PushButton 5, 105, 75, 10, "Documents received", DOCUMENTS_RECEIVED_button
  Text 5, 5, 275, 10, "Document noting scripts main menu: select the script to run from the choices below."
  Text 110, 20, 180, 10, "--- Template for a Change Report Form (CRF) received."
  Text 55, 35, 130, 10, "--- Template for processing DHS-1503."
  Text 55, 50, 180, 10, "--- NEW 06/2015!!! Template for processing DHS-5181."
  Text 120, 65, 195, 10, "--- Template for information about a Medical Opinion Form."
  Text 90, 80, 160, 10, "--- Template for noting Shelter Form information."
  Text 85, 105, 320, 20, "--- Template to indicate what documents you've received for a case (utility proof, AREP verification, etc.). FOR USE WHEN OTHER TEMPLATES DON'T APPLY."
EndDialog


'Variables to declare
IF script_repository = "" THEN script_repository = "https://raw.githubusercontent.com/MN-Script-Team/DHS-MAXIS-Scripts/master/Script Files"		'If it's blank, we're assuming the user is a scriptwriter, ergo, master branch.

'THE SCRIPT----------------------------------------------------------------------------------------------------

'Shows main menu dialog, which asks user which script to run. Loops until a button other than the SIR instructions button is clicked.
Do
	Dialog NOTES_document_noting_menu_dialog
	IF ButtonPressed = cancel THEN StopScript
	If buttonpressed = SIR_instructions_button then CreateObject("WScript.Shell").Run("https://www.dhssir.cty.dhs.state.mn.us/MAXIS/blzn/scriptwiki/Wiki%20Pages/Notes%20scripts.aspx")
Loop until buttonpressed <> SIR_instructions_button

'Connecting to BlueZone
EMConnect ""

		
IF ButtonPressed = CHANGE_REPORT_FORM_RECEIVED_button				THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - CHANGE REPORT FORM RECEIVED.vbs")
IF ButtonPressed = DOCUMENTS_RECEIVED_button						THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - DOCUMENTS RECEIVED.vbs")				
IF ButtonPressed = LTC_1503_button									THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - LTC - 1503.vbs")
IF ButtonPressed = LTC_5181_button									THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - LTC - 5181.vbs")
IF ButtonPressed = MEDICAL_OPINION_FORM_RECEIVED_button				THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - MEDICAL OPINION FORM RECEIVED.vbs")
IF ButtonPressed = SHELTER_FORM_RECEIVED_button						THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - SHELTER FORM RECEIVED.vbs")

'Logging usage stats
script_end_procedure("If you see this, it's because you clicked a button that, for some reason, does not have an outcome in the script. Contact your alpha user to report this bug. Thank you!")