'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ACTIONS - MAIN MENU.vbs"
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
BeginDialog ACTIONS_scripts_main_menu_dialog, 0, 0, 461, 350, "Actions scripts main menu dialog"
  ButtonGroup ButtonPressed
    CancelButton 395, 330, 50, 15
    PushButton 5, 20, 110, 10, "ABAWD Banked Months FIATer", ABAWD_BANKED_MONTHS_button
    PushButton 5, 35, 110, 10, "ABAWD FSET Exemption Check", ABAWD_FSET_EXEMPTION_button
    PushButton 5, 50, 90, 10, "ABAWD Screening Tool", ABAWD_tool
    PushButton 5, 65, 50, 10, "BILS updater", BILS_UPDATER_button
    PushButton 5, 80, 50, 10, "Check EDRS", CHECK_EDRS_button
    PushButton 5, 95, 110, 10, "Child Support Disregard FIATer", CS_FIAT_button
    PushButton 5, 110, 80, 10, "Copy panels to Word", COPY_PANELS_TO_WORD_button
    PushButton 5, 125, 60, 10, " FSET sanction", FSET_sanction_button
    PushButton 5, 140, 80, 10, "Housing Grant FIATer", HOUSING_GRANT_FIATER_button
    PushButton 5, 155, 110, 10, "LTC-Spousal Allocation FIATer", LTC_SPOUSAL_ALLOCATION_FIATER_button
    PushButton 5, 175, 110, 10, "MA-EPD earned income FIATer", MA_EPD_EI_FIAT_button
    PushButton 5, 195, 60, 10, "New job reported", NEW_JOB_REPORTED_button
    PushButton 5, 210, 60, 10, "PA verif request", PA_VERIF_REQUEST_button
    PushButton 5, 225, 70, 10, "Pay stubs Received", PAYSTUBS_RECEIVED_button
    PushButton 5, 250, 100, 10, "Shelter Expense Verif Recv'd", SHELTER_EXPENSE_button
    PushButton 5, 275, 50, 10, "Send SVES", SEND_SVES_button
    PushButton 5, 295, 60, 10, "Transfer case", TRANSFER_CASE_button
    PushButton 5, 310, 65, 10, "TYMA TIKLer", TYMA_TIKLER_button
    PushButton 310, 5, 65, 10, "UTILITIES scripts", UTILITIES_SCRIPTS_button
    PushButton 385, 5, 70, 10, "SIR instructions", SIR_instructions_button
  Text 5, 5, 250, 10, "Action scripts main menu: select the script to run from the choices below."
  Text 120, 20, 335, 10, "-- NEW 01/2016!! FIATS SNAP eligibility, income and deductions for HH members using banked months."
  Text 125, 35, 320, 10, "--- Double checks a case to see if any possible ABAWD/FSET exemptions exist."
  Text 105, 50, 270, 10, "--- A tool to walk through a screening to determine if client is ABAWD."
  Text 65, 65, 220, 10, "--- Updates a BILS panel with reoccurring or actual BILS received."
  Text 65, 80, 190, 10, "--- sends an EDRS request for a HH member on a case."
  Text 125, 95, 330, 10, "--- FIATS in the CS disregard for MFIP and DWP as described in CM 17.15.03"
  Text 95, 110, 180, 10, "--- Copies MAXIS panels to Word en masse for a case."
  Text 75, 125, 370, 10, "--- Updates the WREG panel, and case notes when imposing or resolving a FSET sanction."
  Text 95, 140, 350, 15, "--- FIATs out the SHEL Housing Subsidy making the MFIP case eligible for the Housing Grant."
  Text 125, 155, 180, 10, "--- FIATs a spousal allocation across a budget period."
  Text 125, 175, 300, 10, "--- FIATs MA-EPD earned income (JOBS income) to be even across an entire budget period."
  Text 75, 195, 380, 10, "--- Creates a JOBS panel, CASE/NOTE and TIKL when a new job is reported. Use the DAIL scrubber for new hire DAILs."
  Text 75, 210, 320, 10, "--- Creates a Word document with PA benefit totals for other agencies to determine client benefits."
  Text 85, 225, 370, 20, "--- Enter in pay stubs on one dialog, and it puts that information on JOBS (both retrospective and prospective if applicable), as well as the PIC and HC pop-up, and it'll case note the income as well."
  Text 110, 250, 340, 20, "-- NEW 01/2016!! Enter shelter expense and address information in a single dialog and the script updates SHEL, HEST, and ADDR and case notes."
  Text 65, 275, 90, 10, "--- Sends a SVES/QURY."
  Text 75, 295, 330, 10, "--- SPEC/XFERs a case, and can send a client memo. For in-agency as well as out-of-county XFERs."
  Text 80, 310, 355, 10, "--- NEW 02/2016!!! TIKLS for TYMA report forms to be sent. "
EndDialog


'Variables to declare
IF script_repository = "" THEN script_repository = "https://raw.githubusercontent.com/MN-Script-Team/DHS-MAXIS-Scripts/master/Script Files"		'If it's blank, we're assuming the user is a scriptwriter, ergo, master branch.

'THE SCRIPT----------------------------------------------------------------------------------------------------
'Shows dialog, which asks user which script to run.
Do
	dialog ACTIONS_scripts_main_menu_dialog
	If buttonpressed = cancel then stopscript
	If buttonpressed = SIR_instructions_button then CreateObject("WScript.Shell").Run("https://www.dhssir.cty.dhs.state.mn.us/MAXIS/blzn/Script%20Instructions%20Wiki/Actions%20scripts.aspx")
Loop until buttonpressed <> SIR_instructions_button

'Connecting to BlueZone
EMConnect ""

If buttonpressed = ABAWD_BANKED_MONTHS_button then call run_from_GitHub(script_repository & "/ACTIONS/ACTIONS - ABAWD BANKED MONTHS FIATER.vbs")
IF buttonpressed = ABAWD_FSET_EXEMPTION_button then call run_from_GitHub(script_repository & "/ACTIONS/ACTIONS - ABAWD FSET EXEMPTION CHECK.vbs")
IF buttonpressed = ABAWD_tool then call run_from_GitHub(script_repository & "/ACTIONS/ACTIONS - ABAWD SCREENING TOOL.vbs")
If buttonpressed = BILS_UPDATER_button then call run_from_GitHub(script_repository & "/ACTIONS/ACTIONS - BILS UPDATER.vbs")
If buttonpressed = CHECK_EDRS_button then call run_from_GitHub(script_repository & "/ACTIONS/ACTIONS - CHECK EDRS.vbs")
If buttonpressed = CS_FIAT_button then call run_from_GitHub(script_repository & "/ACTIONS/ACTIONS - CS DISREGARD FIAT.vbs")
If buttonpressed = COPY_PANELS_TO_WORD_button then call run_from_GitHub(script_repository & "/ACTIONS/ACTIONS - COPY PANELS TO WORD.vbs")
IF ButtonPressed = FSET_sanction_button	THEN CALL run_from_GitHub(script_repository & "/ACTIONS/ACTIONS - FSET SANCTION.vbs")
IF ButtonPressed = HOUSING_GRANT_FIATER_button THEN call run_from_GitHub(script_repository & "/ACTIONS/ACTIONS - HOUSING GRANT FIATER.vbs")
If buttonpressed = LTC_SPOUSAL_ALLOCATION_FIATER_button then call run_from_GitHub(script_repository & "/ACTIONS/ACTIONS - LTC - SPOUSAL ALLOCATION FIATER.vbs")
If buttonpressed = MA_EPD_EI_FIAT_button then call run_from_GitHub(script_repository & "/ACTIONS/ACTIONS - MA-EPD EI FIAT.vbs")
If buttonpressed = NEW_JOB_REPORTED_button then call run_from_GitHub(script_repository & "/ACTIONS/ACTIONS - NEW JOB REPORTED.vbs")
If buttonpressed = PA_VERIF_REQUEST_button then call run_from_GitHub(script_repository & "/ACTIONS/ACTIONS - PA VERIF REQUEST.vbs")
If buttonpressed = PAYSTUBS_RECEIVED_button then call run_from_GitHub(script_repository & "/ACTIONS/ACTIONS - PAYSTUBS RECEIVED.vbs")
If buttonpressed = SEND_SVES_button then call run_from_GitHub(script_repository & "/ACTIONS/ACTIONS - SEND SVES.vbs")
IF ButtonPressed = SHELTER_EXPENSE_button THEN CALL run_from_GitHub(script_repository & "/ACTIONS/ACTIONS - SHELTER EXPENSE VERIF RECEIVED.vbs")
If buttonpressed = TRANSFER_CASE_button then call run_from_GitHub(script_repository & "/ACTIONS/ACTIONS - TRANSFER CASE.vbs")
If buttonpressed = TYMA_TIKLER_button then call run_from_GitHub(script_repository & "/ACTIONS/ACTIONS - TYMA TIKLER.vbs")

If ButtonPressed = UTILITIES_SCRIPTS_button then call run_from_GitHub(script_repository & "/UTILITIES/UTILITIES - MAIN MENU.vbs")

'Logging usage stats
script_end_procedure("If you see this, it's because you clicked a button that, for some reason, does not have an outcome in the script. Contact your alpha user to report this bug. Thank you!")
