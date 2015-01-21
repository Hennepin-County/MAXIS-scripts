'LOADING ROUTINE FUNCTIONS-------------------------------------------------------------------------------------------
url = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER FUNCTIONS LIBRARY.vbs"
Set req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a URL
req.open "GET", url, False									'Attempts to open the URL
req.send													'Sends request
IF req.Status = 200 Then									'200 means great success
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
			"URL: " & url
			script_end_procedure("Script ended due to error connecting to GitHub.")
END IF


'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog NOTES_H_Z_scripts_main_menu_dialog, 0, 0, 456, 280, "Notes (H-Z) scripts main menu dialog"
  ButtonGroup ButtonPressed
    CancelButton 400, 260, 50, 15
    PushButton 10, 25, 50, 10, "HC Renewal", HC_RENEWAL_button
    PushButton 10, 40, 30, 10, "HCAPP", HCAPP_button
    PushButton 10, 55, 25, 10, "HRF", HRF_button
    PushButton 10, 70, 45, 10, "LEP - SAVE", LEP_SAVE_button
    PushButton 10, 85, 80, 10, "LEP - Sponsor income", LEP_SPONSOR_INCOME_button
    PushButton 10, 100, 45, 10, "LTC - 1503", LTC_1503_button
    PushButton 10, 115, 90, 10, "LTC - Application received", LTC_APPLICATION_RECEIVED_button
    PushButton 10, 130, 85, 10, "LTC - Asset assessment", LTC_ASSET_ASSESSMENT_button
    PushButton 10, 145, 95, 10, "LTC - COLA summary 2015", LTC_COLA_SUMMARY_2015_button
    PushButton 10, 160, 75, 10, "LTC - Intake approval", LTC_INTAKE_APPROVAL_button
    PushButton 10, 175, 65, 10, "LTC - MA approval", LTC_MA_APPROVAL_button
    PushButton 10, 190, 55, 10, "LTC - Renewal", LTC_RENEWAL_button
    PushButton 10, 205, 125, 10, "MFIP sanction/DWP disqualification", MFIP_SANCTION_AND_DWP_DISQUALIFICATION_button
    PushButton 10, 220, 110, 10, "Mileage reimbursement request", MILEAGE_REIMBURSEMENT_REQUEST_button
    PushButton 10, 235, 110, 10, "MNsure - Documents requested", MNSURE_DOCUMENTS_REQUESTED_button
    PushButton 10, 250, 50, 10, "Overpayment", OVERPAYMENT_button
    PushButton 10, 265, 75, 10, "Verifications needed", VERIFICATIONS_NEEDED_button
  Text 5, 5, 245, 10, "Notes scripts main menu: select the script to run from the choices below."
  Text 65, 25, 140, 10, "--- A case note template for HC renewals."
  Text 45, 40, 120, 10, "--- A case note template for HCAPPs."
  Text 40, 55, 240, 10, "--- A case note template for HRFs (for GRH, use the ''GRH - HRF'' script)."
  Text 60, 70, 255, 10, "--- A case note template for the SAVE system for verifying immigration status."
  Text 95, 85, 345, 10, "--- A case note template for the sponsor income deeming calculation (it will also help calculate it for you)."
  Text 60, 100, 165, 10, "--- A case note template for processing DHS-1503."
  Text 105, 115, 205, 10, "--- A case note template for initial details of a LTC application."
  Text 100, 130, 340, 10, "--- A case note template for the LTC asset assessment. Will enter both person and case notes if desired."
  Text 110, 145, 250, 10, "--- A case note template to summarize actions for the 2015 COLA."
  Text 90, 160, 205, 10, "--- A case note template for use when approving a LTC intake."
  Text 80, 175, 355, 10, "--- A case note template for approving LTC MA (can be used for changes, initial application, or recertification)."
  Text 70, 190, 140, 10, "--- A case note template for LTC renewals."
  Text 140, 205, 290, 10, "--- A case note template for MFIP sanctions and DWP disqualifications, both CS and ES."
  Text 125, 220, 260, 10, "--- A case note template for actions taken on medical mileage reimbursements."
  Text 125, 235, 250, 10, "--- A case note template for when MNsure documents have been requested."
  Text 65, 250, 240, 10, "--- A case note template for noting basic information about overpayments."
  Text 90, 265, 300, 10, "--- A case note template for when verifications are needed (enters each verification clearly)."
EndDialog



'THE SCRIPT----------------------------------------------------------------------------------------------------

'Shows report scanning dialog, which asks user which report to generate.
Dialog NOTES_H_Z_scripts_main_menu_dialog
IF ButtonPressed = cancel THEN StopScript

'Connecting to BlueZone
EMConnect ""

IF ButtonPressed = 	HC_RENEWAL_button								THEN CALL run_from_GitHub(script_repository & "NOTES/NOTES - HC RENEWAL.vbs")
IF ButtonPressed = 	HCAPP_button									THEN CALL run_from_GitHub(script_repository & "NOTES/NOTES - HCAPP.vbs")
IF ButtonPressed = 	HRF_button										THEN CALL run_from_GitHub(script_repository & "NOTES/NOTES - HRF.vbs")
IF ButtonPressed = 	LEP_SAVE_button									THEN CALL run_from_GitHub(script_repository & "NOTES/NOTES - LEP - SAVE.vbs")
IF ButtonPressed = 	LEP_SPONSOR_INCOME_button						THEN CALL run_from_GitHub(script_repository & "NOTES/NOTES - LEP - SPONSOR INCOME.vbs")
IF ButtonPressed = 	LTC_1503_button									THEN CALL run_from_GitHub(script_repository & "NOTES/NOTES - LTC - 1503.vbs")
IF ButtonPressed = 	LTC_APPLICATION_RECEIVED_button					THEN CALL run_from_GitHub(script_repository & "NOTES/NOTES - LTC - APPLICATION RECEIVED.vbs")
IF ButtonPressed = 	LTC_ASSET_ASSESSMENT_button						THEN CALL run_from_GitHub(script_repository & "NOTES/NOTES - LTC - ASSET ASSESSMENT.vbs")
IF ButtonPressed = 	LTC_COLA_SUMMARY_2015_button					THEN CALL run_from_GitHub(script_repository & "NOTES/NOTES - LTC - COLA SUMMARY 2015.vbs")
IF ButtonPressed = 	LTC_INTAKE_APPROVAL_button						THEN CALL run_from_GitHub(script_repository & "NOTES/NOTES - LTC - INTAKE APPROVAL.vbs")
IF ButtonPressed = 	LTC_MA_APPROVAL_button							THEN CALL run_from_GitHub(script_repository & "NOTES/NOTES - LTC - MA APPROVAL.vbs")
IF ButtonPressed = 	LTC_RENEWAL_button								THEN CALL run_from_GitHub(script_repository & "NOTES/NOTES - LTC - RENEWAL.vbs")
IF ButtonPressed = 	MFIP_SANCTION_AND_DWP_DISQUALIFICATION_button	THEN CALL run_from_GitHub(script_repository & "NOTES/NOTES - MFIP SANCTION AND DWP DISQUALIFICATION.vbs")
IF ButtonPressed = 	MILEAGE_REIMBURSEMENT_REQUEST_button			THEN CALL run_from_GitHub(script_repository & "NOTES/NOTES - MILEAGE REIMBURSEMENT REQUEST.vbs")
IF ButtonPressed = 	MNSURE_DOCUMENTS_REQUESTED_button				THEN CALL run_from_GitHub(script_repository & "NOTES/NOTES - MNSURE - DOCUMENTS REQUESTED.vbs")
IF ButtonPressed = 	OVERPAYMENT_button								THEN CALL run_from_GitHub(script_repository & "NOTES/NOTES - OVERPAYMENT.vbs")
IF ButtonPressed = 	VERIFICATIONS_NEEDED_button						THEN CALL run_from_GitHub(script_repository & "NOTES/NOTES - VERIFICATIONS NEEDED.vbs")


'Logging usage stats
script_end_procedure("If you see this, it's because you clicked a button that, for some reason, does not have an outcome in the script. Contact your alpha user to report this bug. Thank you!")