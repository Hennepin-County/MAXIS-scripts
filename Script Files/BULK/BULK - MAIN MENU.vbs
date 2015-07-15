'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - MAIN MENU.vbs"
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

DIM ButtonPressed 
DIM SIR_instructions_button, dialog_name
DIM BULK_list_scripts_button, other_BULK_scripts_button
DIM BULK_TIKLER_button, CASE_NOTE_FROM_EXCEL_LIST_button, CEI_PREMIUM_NOTER_button, COLA_AUTO_APPROVED_DAIL_NOTER_button, INAC_SCRUBBER_button, RETURNED_MAIL_button, REVW_MONT_CLOSURES_button
DIM ACTV_LIST_button, DAIL_REPORT_button, EOMC_LIST_button, PND1_LIST_button, PND2_LIST_button, REVS_LIST_button, REVW_LIST_button, MFCM_LIST_button, ADDRESS_LIST_button, ARST_LIST_button, CHECK_SNAP_FOR_GA_RCA_button, LTC_GRH_LIST_GENERATOR_button, MAEPD_MEDICARE_LIST_button, MISC_NON_MAGI_HC_DEDUCTIONS_button, SWKR_LIST_GENERATOR_button


FUNCTION create_BULK_main_menu(dialog_name)
	IF dialog_name = "OTHER BULK" THEN 
		BeginDialog dialog_name, 0, 0, 456, 285, "Other Bulk Scripts Main Menu"
		ButtonGroup ButtonPressed
			PushButton 5, 25, 55, 15, "BULK LISTS", BULK_list_scripts_button
			PushButton 375, 5, 65, 10, "SIR instructions", SIR_instructions_button
			PushButton 10, 80, 45, 10, "Bulk TIKLer", BULK_TIKLER_button
			PushButton 10, 95, 95, 10, "CASE/NOTE from Excel list", CASE_NOTE_FROM_EXCEL_LIST_button
			PushButton 10, 120, 70, 10, "CEI premium noter", CEI_PREMIUM_NOTER_button
			PushButton 10, 135, 110, 10, "COLA auto approved DAIL noter", COLA_AUTO_APPROVED_DAIL_NOTER_button
			PushButton 10, 160, 55, 10, "INAC scrubber", INAC_SCRUBBER_button
			PushButton 10, 200, 55, 10, "Returned mail", RETURNED_MAIL_button
			PushButton 10, 225, 80, 10, "REVW/MONT closures", REVW_MONT_CLOSURES_button
			CancelButton 400, 265, 50, 15
		Text 5, 5, 235, 10, "Bulk scripts main menu: select the script to run from the choices below."
		GroupBox 5, 60, 440, 190, "Other BULK Scripts"
		Text 60, 80, 175, 10, "--- Creates the same TIKL on up to 60 cases at once."
		Text 110, 95, 335, 20, "--- Creates the same CASE/NOTE on potentially hundreds of cases listed on an Excel spreadsheet of your choice."
		Text 85, 120, 240, 10, "--- Case notes recurring CEI premiums on multiple cases simultaneously."
		Text 125, 135, 320, 20, "--- NEW 06/2015!!! Case notes all cases on DAIL/DAIL that have a message indicating that COLA was auto-approved, copies the messages to an Excel spreadsheet, and deletes the DAIL."
		Text 70, 160, 375, 35, "--- Checks all cases on REPT/INAC (in the month before the current footer month, or prior) for MMIS discrepancies, active claims, DAIL messages, and ABPS panels in need of update (for Good Cause status), and adds them to a Word document. After that, it case notes all of the cases without DAIL messages or MMIS discrepancies. If your agency uses a closed-file worker number, it will SPEC/XFER the cases from your number into that number."
		Text 70, 200, 375, 20, "--- Case notes that returned mail (without a forwarding address) was received for up to 60 cases simultaneously, and TIKLs for 10-day return of proofs."
		Text 95, 225, 350, 20, "--- Case notes all cases on REPT/REVW or REPT/MONT that are closing for missing or incomplete CAF/HRF/CSR/HC ER. Case notes ''last day of REIN'' as well as ''date case becomes an intake.''"
		Text 65, 30, 375, 10, "-- This will navigate you to the menu for the BULK List Generators."
		EndDialog
	ELSEIF dialog_name = "BULK LISTS" THEN
		BeginDialog dialog_name, 0, 0, 456, 285, "BULK List Generators"
		ButtonGroup ButtonPressed
			PushButton 5, 25, 55, 15, "OTHER BULK", other_BULK_scripts_button
			PushButton 375, 5, 65, 10, "SIR instructions", SIR_instructions_button
			PushButton 10, 105, 25, 10, "ACTV", ACTV_LIST_button
			PushButton 35, 105, 25, 10, "DAIL", DAIL_REPORT_button
			PushButton 60, 105, 25, 10, "EOMC", EOMC_LIST_button
			PushButton 85, 105, 25, 10, "PND1", PND1_LIST_button
			PushButton 10, 115, 25, 10, "PND2", PND2_LIST_button
			PushButton 35, 115, 25, 10, "REVS", REVS_LIST_button
			PushButton 60, 115, 30, 10, "REVW", REVW_LIST_button
			PushButton 85, 115, 25, 10, "MFCM", MFCM_LIST_button
			PushButton 125, 75, 25, 10, "ADDR", ADDRESS_LIST_button
			PushButton 125, 90, 25, 10, "ARST", ARST_LIST_button
			PushButton 125, 110, 85, 10, "Check SNAP for GA/RCA", CHECK_SNAP_FOR_GA_RCA_button
			PushButton 125, 135, 65, 10, "LTC-GRH list gen", LTC_GRH_LIST_GENERATOR_button
			PushButton 125, 160, 80, 10, "MA-EPD/Medi Pt B CEI", MAEPD_MEDICARE_LIST_button
			PushButton 125, 185, 105, 10, "Misc. non-MAGI HC deductions", MISC_NON_MAGI_HC_DEDUCTIONS_button
			PushButton 125, 210, 55, 10, "SWKR list gen", SWKR_LIST_GENERATOR_button
			CancelButton 400, 265, 50, 15
		Text 5, 5, 235, 10, "Bulk scripts main menu: select the script to run from the choices below."
		GroupBox 5, 60, 110, 70, "Case lists"
		Text 10, 75, 100, 25, "Case list scripts pull a list of cases into an Excel spreadsheet."
		GroupBox 120, 60, 330, 175, "Other bulk lists"
		Text 155, 90, 215, 10, "--- Caseload stats by worker. Includes most MAXIS programs."
		Text 195, 135, 250, 20, "--- Creates a list of FACIs, AREPs, and waiver types assigned to the various cases in a caseload (or group of caseloads)."
		Text 210, 160, 230, 20, "--- NEW 07/2015!!! Creates a list of cases and clients active on MA-EPD and Medicare Part B that are eligible for Part B reimbursement."
		Text 235, 185, 210, 20, "--- NEW 06/2015!!! Creates a list of cases with non-MAGI HC deductions."
		Text 185, 210, 260, 20, "--- Creates a list of SWKRs assigned to the various cases in a caseload (or group of caseloads)."
		Text 155, 75, 285, 10, "-- NEW 07/2015!!! Creates a list of all addresses for an entire caseload."
		Text 65, 30, 375, 10, "-- This will navigate you to the menu for the Other BULK scripts."
		Text 215, 110, 225, 20, "--- NEW 07/2015!!! Compares the amount of GA and RCA FIAT'd into SNAP and creates a list of the results."
		EndDialog
	END IF

	DIALOG dialog_name

END FUNCTION

'=====THE SCRIPT=====
EMConnect ""

'Setting the default menu to the BULK LISTS
dialog_name = "BULK LISTS"
DO
	'Calling the function that loads the dialogs
	CALL create_BULK_main_menu(dialog_name)
		IF ButtonPressed = 0 THEN stopscript
		'Opening the SIR Instructions
		IF buttonpressed = SIR_instructions_button then CreateObject("WScript.Shell").Run("https://www.dhssir.cty.dhs.state.mn.us/MAXIS/blzn/scriptwiki/Wiki%20Pages/Bulk%20scripts.aspx")

		'If the user selects the other sub-menu, the script do-loops with the new dialog_name
		IF ButtonPressed = BULK_list_scripts_button THEN 
			dialog_name = "BULK LISTS"
		ELSEIF ButtonPressed = other_BULK_scripts_button THEN 
			dialog_name = "OTHER BULK"
		END IF

		'If the user selects a script button, the script will exit the do-loop
LOOP UNTIL ButtonPressed <> SIR_instructions_button AND ButtonPressed <> BULK_list_scripts_button AND ButtonPressed <> other_BULK_scripts_button

'Available scripts
If ButtonPressed = ACTV_LIST_button then call run_from_GitHub(script_repository & "/BULK/BULK - REPT-ACTV LIST.vbs")
IF ButtonPressed = DAIL_REPORT_button THEN CALL run_from_GitHub(script_repository & "/BULK/BULK - DAIL REPORT.vbs")
If ButtonPressed = EOMC_LIST_button then call run_from_GitHub(script_repository & "/BULK/BULK - REPT-EOMC LIST.vbs")
If ButtonPressed = PND1_LIST_button then call run_from_GitHub(script_repository & "/BULK/BULK - REPT-PND1 LIST.vbs")
If ButtonPressed = PND2_LIST_button then call run_from_GitHub(script_repository & "/BULK/BULK - REPT-PND2 LIST.vbs")
If ButtonPressed = REVS_LIST_button then call run_from_GitHub(script_repository & "/BULK/BULK - REPT-REVS LIST.vbs")
If ButtonPressed = REVW_LIST_button then call run_from_GitHub(script_repository & "/BULK/BULK - REPT-REVW LIST.vbs")
If ButtonPressed = MFCM_LIST_button then call run_from_GitHub(script_repository & "/BULK/BULK - REPT-MFCM LIST.vbs")
IF ButtonPressed = ADDRESS_LIST_button THEN CALL run_from_GitHub(script_repository & "/BULK/BULK - ADDRESS REPORT.vbs")
If ButtonPressed = ARST_LIST_button then call run_from_GitHub(script_repository & "/BULK/BULK - REPT-ARST LIST.vbs")
IF ButtonPressed = CHECK_SNAP_FOR_GA_RCA_button THEN CALL run_from_GitHub(script_repository & "/BULK/BULK - CHECK SNAP FOR GA RCA.vbs")
If ButtonPressed = LTC_GRH_LIST_GENERATOR_button then call run_from_GitHub(script_repository & "/BULK/BULK - LTC-GRH LIST GENERATOR.vbs")
IF ButtonPressed = MAEPD_MEDICARE_LIST_button THEN CALL run_from_GitHub(script_repository & "/BULK/BULK - FIND MAEPD MEDI CEI.vbs")
If ButtonPressed = MISC_NON_MAGI_HC_DEDUCTIONS_button then call run_from_GitHub(script_repository & "/BULK/BULK - MISC NON-MAGI HC DEDUCTIONS.vbs")
If ButtonPressed = SWKR_LIST_GENERATOR_button then call run_from_GitHub(script_repository & "/BULK/BULK - SWKR LIST GENERATOR.vbs")
If ButtonPressed = BULK_TIKLER_button then call run_from_GitHub(script_repository & "/BULK/BULK - BULK TIKLER.vbs")
If ButtonPressed = CASE_NOTE_FROM_EXCEL_LIST_button then call run_from_GitHub(script_repository & "/BULK/BULK - CASE NOTE FROM EXCEL LIST.vbs")
If ButtonPressed = CEI_PREMIUM_NOTER_button then call run_from_GitHub(script_repository & "/BULK/BULK - CEI PREMIUM NOTER.vbs")
If ButtonPressed = COLA_AUTO_APPROVED_DAIL_NOTER_button then call run_from_GitHub(script_repository & "/BULK/BULK - COLA AUTO APPROVED DAIL NOTER.vbs")
If ButtonPressed = INAC_SCRUBBER_button then call run_from_GitHub(script_repository & "/BULK/BULK - INAC SCRUBBER.vbs")
If ButtonPressed = RETURNED_MAIL_button then call run_from_GitHub(script_repository & "/BULK/BULK - RETURNED MAIL.vbs")
If ButtonPressed = REVW_MONT_CLOSURES_button then call run_from_GitHub(script_repository & "/BULK/BULK - REVW-MONT CLOSURES.vbs")

'Logging usage stats
script_end_procedure("If you see this, it's because you clicked a button that, for some reason, does not have an outcome in the script. Contact your alpha user to report this bug. Thank you!")
