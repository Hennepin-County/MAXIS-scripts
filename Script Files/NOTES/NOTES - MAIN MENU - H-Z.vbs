'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - MAIN MENU (H-Z).vbs"
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
BeginDialog NOTES_H_Z_scripts_main_menu_dialog, 0, 0, 516, 335, "Notes (H-Z) scripts main menu dialog"
  ButtonGroup ButtonPressed
    CancelButton 460, 315, 50, 15
    PushButton 445, 10, 65, 10, "SIR instructions", SIR_instructions_button
    PushButton 5, 20, 60, 10, "Document noting", DOCUMENT_NOTING_button
    PushButton 5, 45, 50, 10, "HC Renewal", HC_RENEWAL_button
    PushButton 5, 60, 30, 10, "HCAPP", HCAPP_button
    PushButton 5, 75, 65, 10, "HH comp change", HH_COMP_CHANGE_button
    PushButton 5, 90, 25, 10, "HRF", HRF_button
    PushButton 5, 105, 45, 10, "LEP - SAVE", LEP_SAVE_button
    PushButton 5, 120, 80, 10, "LEP - Sponsor income", LEP_SPONSOR_INCOME_button
    PushButton 5, 135, 90, 10, "LTC - Application received", LTC_APPLICATION_RECEIVED_button
    PushButton 5, 150, 85, 10, "LTC - Asset assessment", LTC_ASSET_ASSESSMENT_button
    PushButton 5, 165, 95, 10, "LTC - COLA summary 2015", LTC_COLA_SUMMARY_2015_button
    PushButton 5, 180, 75, 10, "LTC - Intake approval", LTC_INTAKE_APPROVAL_button
    PushButton 5, 195, 65, 10, "LTC - MA approval", LTC_MA_APPROVAL_button
    PushButton 5, 210, 55, 10, "LTC - Renewal", LTC_RENEWAL_button
    PushButton 5, 225, 80, 10, "LTC - Transfer penalty", LTC_TRANSFER_PENALTY_button
    PushButton 5, 240, 125, 10, "MFIP sanction/DWP disqualification", MFIP_SANCTION_AND_DWP_DISQUALIFICATION_button
    PushButton 5, 255, 110, 10, "Mileage reimbursement request", MILEAGE_REIMBURSEMENT_REQUEST_button
    PushButton 5, 270, 110, 10, "MNsure - Documents requested", MNSURE_DOCUMENTS_REQUESTED_button
    PushButton 5, 285, 50, 10, "Overpayment", OVERPAYMENT_button
    PushButton 5, 300, 75, 10, "Verifications needed", VERIFICATIONS_NEEDED_button
  Text 5, 5, 435, 10, "Notes scripts main menu: select the script to run from the choices below. Notes with autofill functionality marked with an asterisk (*)."
  Text 75, 20, 330, 10, "--- NEW 06/2015!!! Templates that are for use when various documents are received."
  Text 60, 45, 140, 10, "--- Template for HC renewals.*"
  Text 40, 60, 120, 10, "--- Template for HCAPPs.*"
  Text 75, 75, 240, 10, "--- NEW 06/2015!!! Template for when you update the HH comp of a case."
  Text 35, 90, 240, 10, "--- Template for HRFs (for GRH, use the ''GRH - HRF'' script).*"
  Text 55, 105, 255, 10, "--- Template for the SAVE system for verifying immigration status."
  Text 90, 120, 345, 10, "--- Template for the sponsor income deeming calculation (it will also help calculate it for you)."
  Text 100, 135, 205, 10, "--- Template for initial details of a LTC application.*"
  Text 95, 150, 340, 10, "--- Template for the LTC asset assessment. Will enter both person and case notes if desired."
  Text 105, 165, 250, 10, "--- Template to summarize actions for the 2015 COLA.*"
  Text 85, 180, 205, 10, "--- Template for use when approving a LTC intake.*"
  Text 75, 195, 355, 10, "--- Template for approving LTC MA (can be used for changes, initial application, or recertification).*"
  Text 65, 210, 140, 10, "--- Template for LTC renewals.*"
  Text 90, 225, 225, 10, "--- Template for noting a transfer penalty."
  Text 135, 240, 290, 10, "--- Template for MFIP sanctions and DWP disqualifications, both CS and ES."
  Text 120, 255, 260, 10, "--- Template for actions taken on medical mileage reimbursements."
  Text 120, 270, 250, 10, "--- Template for when MNsure documents have been requested."
  Text 60, 285, 240, 10, "--- Template for noting basic information about overpayments."
  Text 85, 300, 425, 10, "--- Template for when verifications are needed (enters each verification clearly)."
EndDialog

'Variables to declare
IF script_repository = "" THEN script_repository = "https://raw.githubusercontent.com/MN-Script-Team/DHS-MAXIS-Scripts/master/Script Files"		'If it's blank, we're assuming the user is a scriptwriter, ergo, master branch.

'THE SCRIPT----------------------------------------------------------------------------------------------------

'Shows main menu dialog, which asks user which script to run. Loops until a button other than the SIR instructions button is clicked.
Do
	Dialog NOTES_H_Z_scripts_main_menu_dialog
	IF ButtonPressed = cancel THEN StopScript
	If buttonpressed = SIR_instructions_button then CreateObject("WScript.Shell").Run("https://www.dhssir.cty.dhs.state.mn.us/MAXIS/blzn/scriptwiki/Wiki%20Pages/Notes%20scripts.aspx")
Loop until buttonpressed <> SIR_instructions_button

'Connecting to BlueZone
EMConnect ""

IF ButtonPressed = DOCUMENT_NOTING_button							THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - DOCUMENT NOTING MENU.vbs")
IF ButtonPressed = HC_RENEWAL_button								THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - HC RENEWAL.vbs")
IF ButtonPressed = HCAPP_button										THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - HCAPP.vbs")
IF ButtonPressed = HH_COMP_CHANGE_button							THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - HH COMP CHANGE.vbs")
IF ButtonPressed = HRF_button										THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - HRF.vbs")
IF ButtonPressed = LEP_SAVE_button									THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - LEP - SAVE.vbs")
IF ButtonPressed = LEP_SPONSOR_INCOME_button						THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - LEP - SPONSOR INCOME.vbs")
IF ButtonPressed = LTC_APPLICATION_RECEIVED_button					THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - LTC - APPLICATION RECEIVED.vbs")
IF ButtonPressed = LTC_ASSET_ASSESSMENT_button						THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - LTC - ASSET ASSESSMENT.vbs")
IF ButtonPressed = LTC_COLA_SUMMARY_2015_button						THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - LTC - COLA SUMMARY 2015.vbs")
IF ButtonPressed = LTC_INTAKE_APPROVAL_button						THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - LTC - INTAKE APPROVAL.vbs")
IF ButtonPressed = LTC_MA_APPROVAL_button							THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - LTC - MA APPROVAL.vbs")
IF ButtonPressed = LTC_RENEWAL_button								THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - LTC - RENEWAL.vbs")
IF ButtonPressed = LTC_TRANSFER_PENALTY_button						THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - LTC - TRANSFER PENALTY.vbs")
IF ButtonPressed = MFIP_SANCTION_AND_DWP_DISQUALIFICATION_button	THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - MFIP SANCTION AND DWP DISQUALIFICATION.vbs")
IF ButtonPressed = MILEAGE_REIMBURSEMENT_REQUEST_button				THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - MILEAGE REIMBURSEMENT REQUEST.vbs")
IF ButtonPressed = MNSURE_DOCUMENTS_REQUESTED_button				THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - MNSURE - DOCUMENTS REQUESTED.vbs")
IF ButtonPressed = OVERPAYMENT_button								THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - OVERPAYMENT.vbs")
IF ButtonPressed = VERIFICATIONS_NEEDED_button						THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - VERIFICATIONS NEEDED.vbs")

'Logging usage stats
script_end_procedure("If you see this, it's because you clicked a button that, for some reason, does not have an outcome in the script. Contact your alpha user to report this bug. Thank you!")