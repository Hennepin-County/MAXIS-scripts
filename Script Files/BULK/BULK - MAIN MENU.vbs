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

'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog BULK_scripts_main_menu_dialog, 0, 0, 456, 330, "Bulk scripts main menu dialog"
  ButtonGroup ButtonPressed
    CancelButton 400, 310, 50, 15
    PushButton 375, 5, 65, 10, "SIR instructions", SIR_instructions_button
    PushButton 10, 65, 25, 10, "ACTV", ACTV_LIST_button
    PushButton 35, 65, 25, 10, "EOMC", EOMC_LIST_button
    PushButton 60, 65, 25, 10, "PND1", PND1_LIST_button
    PushButton 85, 65, 25, 10, "PND2", PND2_LIST_button
    PushButton 20, 75, 25, 10, "REVS", REVS_LIST_button
    PushButton 45, 75, 30, 10, "REVW", REVW_LIST_button
    PushButton 75, 75, 25, 10, "MFCM", MFCM_LIST_button
    PushButton 125, 30, 25, 10, "ARST", ARST_LIST_button
    PushButton 125, 45, 65, 10, "LTC-GRH list gen", LTC_GRH_LIST_GENERATOR_button
    PushButton 125, 70, 105, 10, "Misc. non-MAGI HC deductions", MISC_NON_MAGI_HC_DEDUCTIONS_button
    PushButton 125, 95, 55, 10, "SWKR list gen", SWKR_LIST_GENERATOR_button
    PushButton 10, 135, 45, 10, "Bulk TIKLer", BULK_TIKLER_button
    PushButton 10, 150, 95, 10, "CASE/NOTE from Excel list", CASE_NOTE_FROM_EXCEL_LIST_button
    PushButton 10, 175, 70, 10, "CEI premium noter", CEI_PREMIUM_NOTER_button
    PushButton 10, 190, 110, 10, "COLA auto approved DAIL noter", COLA_AUTO_APPROVED_DAIL_NOTER_button
    PushButton 10, 215, 55, 10, "INAC scrubber", INAC_SCRUBBER_button
    PushButton 10, 255, 55, 10, "Returned mail", RETURNED_MAIL_button
    PushButton 10, 280, 80, 10, "REVW/MONT closures", REVW_MONT_CLOSURES_button
  Text 5, 5, 235, 10, "Bulk scripts main menu: select the script to run from the choices below."
  GroupBox 5, 20, 110, 70, "Case lists"
  Text 10, 35, 100, 25, "Case list scripts pull a list of cases into an Excel spreadsheet."
  GroupBox 120, 20, 330, 100, "Other bulk lists"
  Text 155, 30, 215, 10, "--- Caseload stats by worker. Includes most MAXIS programs."
  Text 195, 45, 250, 20, "--- Creates a list of FACIs, AREPs, and waiver types assigned to the various cases in a caseload (or group of caseloads)."
  Text 235, 70, 210, 20, "--- NEW 06/2015!!! Creates a list of cases with non-MAGI HC deductions."
  Text 185, 95, 260, 20, "--- Creates a list of SWKRs assigned to the various cases in a caseload (or group of caseloads)."
  GroupBox 5, 120, 445, 185, "Other bulk scripts"
  Text 60, 135, 175, 10, "--- Creates the same TIKL on up to 60 cases at once."
  Text 110, 150, 335, 20, "--- Creates the same CASE/NOTE on potentially hundreds of cases listed on an Excel spreadsheet of your choice."
  Text 85, 175, 240, 10, "--- Case notes recurring CEI premiums on multiple cases simultaneously."
  Text 125, 190, 320, 20, "--- NEW 06/2015!!! Case notes all cases on DAIL/DAIL that have a message indicating that COLA was auto-approved, copies the messages to an Excel spreadsheet, and deletes the DAIL."
  Text 70, 215, 375, 35, "--- Checks all cases on REPT/INAC (in the month before the current footer month, or prior) for MMIS discrepancies, active claims, DAIL messages, and ABPS panels in need of update (for Good Cause status), and adds them to a Word document. After that, it case notes all of the cases without DAIL messages or MMIS discrepancies. If your agency uses a closed-file worker number, it will SPEC/XFER the cases from your number into that number."
  Text 70, 255, 375, 20, "--- Case notes that returned mail (without a forwarding address) was received for up to 60 cases simultaneously, and TIKLs for 10-day return of proofs."
  Text 95, 280, 350, 20, "--- Case notes all cases on REPT/REVW or REPT/MONT that are closing for missing or incomplete CAF/HRF/CSR/HC ER. Case notes ''last day of REIN'' as well as ''date case becomes an intake.''"
EndDialog





'VARIABLES TO DECLARE (there's more here than usual because this was originally one big BULK script. This should be added to the other bulk scripts as needed, and removed from here.)
all_case_numbers_array = " "					'Creating blank variable for the future array
call worker_county_code_determination(worker_county_code, two_digit_county_code)	'Determines worker county code
is_not_blank_excel_string = Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34)	'This is the string required to tell excel to ignore blank cells in a COUNTIFS function
IF script_repository = "" THEN script_repository = "https://raw.githubusercontent.com/MN-Script-Team/DHS-MAXIS-Scripts/master/Script Files"		'If it's blank, we're assuming the user is a scriptwriter, ergo, master branch.

'THE SCRIPT----------------------------------------------------------------------------------------------------

'Shows main menu dialog, which asks user which script to run. Loops until a button other than the SIR instructions button is clicked.
Do
	dialog BULK_scripts_main_menu_dialog
	If buttonpressed = cancel then stopscript
	If buttonpressed = SIR_instructions_button then CreateObject("WScript.Shell").Run("https://www.dhssir.cty.dhs.state.mn.us/MAXIS/blzn/scriptwiki/Wiki%20Pages/Bulk%20scripts.aspx")
Loop until buttonpressed <> SIR_instructions_button

'Connecting to BlueZone
EMConnect ""

If ButtonPressed = ACTV_LIST_button then call run_from_GitHub(script_repository & "/BULK/BULK - REPT-ACTV LIST.vbs")
If ButtonPressed = EOMC_LIST_button then call run_from_GitHub(script_repository & "/BULK/BULK - REPT-EOMC LIST.vbs")
If ButtonPressed = PND1_LIST_button then call run_from_GitHub(script_repository & "/BULK/BULK - REPT-PND1 LIST.vbs")
If ButtonPressed = PND2_LIST_button then call run_from_GitHub(script_repository & "/BULK/BULK - REPT-PND2 LIST.vbs")
If ButtonPressed = REVS_LIST_button then call run_from_GitHub(script_repository & "/BULK/BULK - REPT-REVS LIST.vbs")
If ButtonPressed = REVW_LIST_button then call run_from_GitHub(script_repository & "/BULK/BULK - REPT-REVW LIST.vbs")
If ButtonPressed = MFCM_LIST_button then call run_from_GitHub(script_repository & "/BULK/BULK - REPT-MFCM LIST.vbs")
If ButtonPressed = ARST_LIST_button then call run_from_GitHub(script_repository & "/BULK/BULK - REPT-ARST LIST.vbs")
If ButtonPressed = LTC_GRH_LIST_GENERATOR_button then call run_from_GitHub(script_repository & "/BULK/BULK - LTC-GRH LIST GENERATOR.vbs")
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
