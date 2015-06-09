'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ACTIONS - MAIN MENU.vbs"
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
BeginDialog ACTIONS_scripts_main_menu_dialog, 0, 0, 456, 215, "Actions scripts main menu dialog"
  ButtonGroup ButtonPressed
    CancelButton 400, 195, 50, 15
	PushButton 375, 10, 65, 10, "SIR instructions", SIR_instructions_button
    PushButton 5, 20, 50, 10, "BILS updater", BILS_UPDATER_button
    PushButton 5, 35, 50, 10, "Check EDRS", CHECK_EDRS_button
    PushButton 5, 50, 75, 10, "Copy panels to Word", COPY_PANELS_TO_WORD_button
    PushButton 5, 65, 105, 10, "LTC-Spousal Allocation FIATer", LTC_SPOUSAL_ALLOCATION_FIATER_button
    PushButton 5, 80, 105, 10, "MA-EPD earned income FIATer", MA_EPD_EI_FIAT_button
    PushButton 5, 95, 60, 10, "New job reported", NEW_JOB_REPORTED_button
    PushButton 5, 120, 60, 10, "PA verif request", PA_VERIF_REQUEST_button
    PushButton 5, 135, 70, 10, "Paystubs Received", PAYSTUBS_RECEIVED_button
    PushButton 5, 160, 45, 10, "Send SVES", SEND_SVES_button
    PushButton 5, 175, 55, 10, "Transfer case", TRANSFER_CASE_button
    PushButton 5, 200, 85, 10, "Update worker signature", UPDATE_WORKER_SIGNATURE_button
  Text 5, 5, 245, 10, "Action scripts main menu: select the script to run from the choices below."
  Text 60, 20, 215, 10, "--- Updates a BILS panel with reoccurring or actual BILS received."
  Text 60, 35, 185, 10, "--- sends an EDRS request for a HH member on a case."
  Text 85, 50, 180, 10, "--- Copies MAXIS panels to Word en masse for a case."
  Text 115, 65, 175, 10, "--- FIATs a spousal allocation across a budget period."
  Text 115, 80, 295, 10, "--- FIATs MA-EPD earned income (JOBS income) to be even across an entire budget period."
  Text 70, 95, 380, 20, "--- Creates a JOBS panel, a CASE/NOTE, and a TIKL, when client reports a new job. For new HIRE messages on the DAIL, use the DAIL scrubber instead."
  Text 70, 120, 320, 10, "--- Creates a Word document with PA benefit totals for other agencies to determine client benefits."
  Text 80, 135, 370, 20, "--- Enter in paystubs on one dialog, and it puts that information on JOBS (both retrospective and prospective if applicable), as well as the PIC and HC pop-up, and it'll case note the income as well."
  Text 55, 160, 90, 10, "--- Sends a SVES/QURY."
  Text 65, 175, 380, 20, "--- SPEC/XFERs a case, and can send a memo to the new client. For in-agency as well as between agencies (out-of-county XFERs)."
  Text 95, 200, 185, 10, "--- Updates the default worker signature on your scripts."
EndDialog

'Variables to declare
IF script_repository = "" THEN script_repository = "https://raw.githubusercontent.com/MN-Script-Team/DHS-MAXIS-Scripts/master/Script Files"		'If it's blank, we're assuming the user is a scriptwriter, ergo, master branch.

'THE SCRIPT----------------------------------------------------------------------------------------------------

'Shows dialog, which asks user which script to run.
Do
	dialog ACTIONS_scripts_main_menu_dialog
	If buttonpressed = cancel then stopscript
	If buttonpressed = SIR_instructions_button then CreateObject("WScript.Shell").Run("https://www.dhssir.cty.dhs.state.mn.us/MAXIS/blzn/scriptwiki/Wiki%20Pages/Actions%20scripts.aspx")
Loop until buttonpressed <> SIR_instructions_button



'Connecting to BlueZone
EMConnect ""

If buttonpressed = BILS_UPDATER_button then call run_from_GitHub(script_repository & "/ACTIONS/ACTIONS - BILS UPDATER.vbs")
If buttonpressed = CHECK_EDRS_button then call run_from_GitHub(script_repository & "/ACTIONS/ACTIONS - CHECK EDRS.vbs")
If buttonpressed = COPY_PANELS_TO_WORD_button then call run_from_GitHub(script_repository & "/ACTIONS/ACTIONS - COPY PANELS TO WORD.vbs")
If buttonpressed = LTC_SPOUSAL_ALLOCATION_FIATER_button then call run_from_GitHub(script_repository & "/ACTIONS/ACTIONS - LTC - SPOUSAL ALLOCATION FIATER.vbs")
If buttonpressed = MA_EPD_EI_FIAT_button then call run_from_GitHub(script_repository & "/ACTIONS/ACTIONS - MA-EPD EI FIAT.vbs")
If buttonpressed = NEW_JOB_REPORTED_button then call run_from_GitHub(script_repository & "/ACTIONS/ACTIONS - NEW JOB REPORTED.vbs")
If buttonpressed = PA_VERIF_REQUEST_button then call run_from_GitHub(script_repository & "/ACTIONS/ACTIONS - PA VERIF REQUEST.vbs")
If buttonpressed = PAYSTUBS_RECEIVED_button then call run_from_GitHub(script_repository & "/ACTIONS/ACTIONS - PAYSTUBS RECEIVED.vbs")
If buttonpressed = SEND_SVES_button then call run_from_GitHub(script_repository & "/ACTIONS/ACTIONS - SEND SVES.vbs")
If buttonpressed = TRANSFER_CASE_button then call run_from_GitHub(script_repository & "/ACTIONS/ACTIONS - TRANSFER CASE.vbs")
If ButtonPressed = UPDATE_WORKER_SIGNATURE_button then call run_from_GitHub(script_repository & "/ACTIONS/ACTIONS - UPDATE WORKER SIGNATURE.vbs")

'Logging usage stats
script_end_procedure("If you see this, it's because you clicked a button that, for some reason, does not have an outcome in the script. Contact your alpha user to report this bug. Thank you!")
