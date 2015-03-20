'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - MAIN MENU (0-G).vbs"
start_time = timer

'LOADING ROUTINE FUNCTIONS-------------------------------------------------------------------------------------------
If beta_agency = "" or beta_agency = True then
	url = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/BETA/MASTER%20FUNCTIONS%20LIBRARY.vbs"
Else
	url = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
End if
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
BeginDialog NOTES_0_G_scripts_main_menu_dialog, 0, 0, 546, 350, "Notes (0-G) scripts main menu dialog"
  ButtonGroup ButtonPressed
    CancelButton 490, 330, 50, 15
    PushButton 475, 10, 65, 10, "SIR instructions", SIR_instructions_button
    PushButton 10, 25, 105, 10, "ApplyMN application received", APPLYMN_APPLICATION_RECEIVED_button
    PushButton 10, 40, 70, 10, "Approved programs", APPROVED_PROGRAMS_button
    PushButton 10, 55, 45, 10, "Baby born", BABY_BORN_button
    PushButton 10, 70, 50, 10, "Burial assets", BURIAL_ASSETS_button
    PushButton 10, 85, 20, 10, "CAF", CAF_button
    PushButton 10, 100, 100, 10, "Change report form received", CHANGE_REPORT_FORM_RECEIVED_button
    PushButton 10, 115, 95, 10, "Citizenship/identity verified", CITIZENSHIP_IDENTITY_VERIFIED_button
    PushButton 10, 130, 115, 10, "Client contact (call center version)", CLIENT_CONTACT_CALL_CENTER_VERSION_button
    PushButton 10, 145, 50, 10, "Client contact", CLIENT_CONTACT_button
    PushButton 10, 160, 60, 10, "Closed programs", CLOSED_PROGRAMS_button
    PushButton 10, 180, 50, 10, "Combined AR", COMBINED_AR_button
    PushButton 10, 195, 20, 10, "CSR", CSR_button
    PushButton 10, 210, 60, 10, "Denied programs", DENIED_PROGRAMS_button
    PushButton 10, 230, 75, 10, "Documents received", DOCUMENTS_RECEIVED_button
    PushButton 10, 245, 40, 10, "Drug felon", DRUG_FELON_button
    PushButton 10, 260, 50, 10, "DWP budget", DWP_BUDGET_button
    PushButton 10, 275, 45, 10, "Emergency", EMERGENCY_button
    PushButton 10, 290, 75, 10, "Expedited screening", EXPEDITED_SCREENING_button
    PushButton 10, 305, 40, 10, "Fraud info", FRAUD_INFO_button
    PushButton 10, 320, 65, 10, "Gas cards issued", GAS_CARDS_ISSUED_button
    PushButton 10, 335, 45, 10, "GRH - HRF", GRH_HRF_button
  Text 5, 5, 245, 10, "Notes scripts main menu: select the script to run from the choices below."
  Text 120, 25, 330, 10, "--- A case note template for documenting details about an ApplyMN application recevied."
  Text 85, 40, 455, 10, "--- A case note template for when you approve a clients programs. Can autofill some data about the approval (like benefit totals) from MAXIS."
  Text 60, 55, 270, 10, "--- NEW 04/2015!!! A case note template for a baby born and added to household."
  Text 65, 70, 135, 10, "--- A case note template for burial assets."
  Text 35, 85, 390, 10, "--- A case note template for when youre processing a CAF. Works for intake as well as recertification and reapplication."
  Text 115, 100, 265, 10, "--- NEW 04/2015!!! A case note template for a change report form (CRF) received."
  Text 110, 115, 295, 10, "--- A case note template for documenting citizenship/identity status for a client (or clients)."
  Text 130, 130, 285, 10, "--- A case note template for documenting client contact (for call center/phone bank staff)."
  Text 65, 145, 260, 10, "--- A case note template for documenting client contact, either from or to a client."
  Text 75, 160, 455, 20, "--- A case note template for indicating which programs are closing, and when. Also case notes intake/REIN dates based on programs closing and conditions of closure."
  Text 65, 180, 250, 10, "--- A case note template for the Combined Annual Renewal (or Combined AR)."
  Text 35, 195, 120, 10, "--- A case note template for the CSR."
  Text 75, 210, 465, 20, "--- A case note template for indicating which programs youve denied, and when. Also case notes intake/REIN dates based on programs denied and conditions of denial."
  Text 90, 230, 405, 10, "--- A case note template to clearly indicate what documents youve received for a case (shelter form, AREP verification, etc.)."
  Text 55, 245, 215, 10, "--- NEW 04/2015!!! A case note template for noting drug felon info."
  Text 65, 260, 215, 10, "--- NEW 04/2015!!! A case note template for noting DWP budgets."
  Text 60, 275, 240, 10, "--- A case note template for emergency assistance applications (EA/EGA)."
  Text 90, 290, 220, 10, "--- A case note template for screening a client for expedited status."
  Text 55, 305, 200, 10, "--- NEW 04/2015!!! A case note template for noting fraud info."
  Text 80, 320, 450, 10, "--- NEW 04/2015!!! A case note template for gas card issuance. Consult with a supervisor to make sure this is appropriate for your agency."
  Text 60, 335, 210, 10, "--- A case note template for GRH HRFs. Case must be post-pay."
EndDialog


'THE SCRIPT----------------------------------------------------------------------------------------------------

'Shows main menu dialog, which asks user which script to run. Loops until a button other than the SIR instructions button is clicked.
Do
	Dialog NOTES_0_G_scripts_main_menu_dialog
	IF ButtonPressed = cancel THEN StopScript
	If buttonpressed = SIR_instructions_button then CreateObject("WScript.Shell").Run("https://www.dhssir.cty.dhs.state.mn.us/MAXIS/blzn/scriptwiki/Wiki%20Pages/Notes%20scripts.aspx")
Loop until buttonpressed <> SIR_instructions_button

'Connecting to BlueZone
EMConnect ""

IF ButtonPressed = APPLYMN_APPLICATION_RECEIVED_button			THEN CALL run_from_GitHub(script_repository & "NOTES/NOTES - APPLYMN APPLICATION RECEIVED.vbs")		
IF ButtonPressed = APPROVED_PROGRAMS_button						THEN CALL run_from_GitHub(script_repository & "NOTES/NOTES - APPROVED PROGRAMS.vbs")					
IF ButtonPressed = BABY_BORN_button								THEN CALL run_from_GitHub(script_repository & "NOTES/NOTES - BABY BORN.vbs")
IF ButtonPressed = BURIAL_ASSETS_button							THEN CALL run_from_GitHub(script_repository & "NOTES/NOTES - BURIAL ASSETS.vbs")						
IF ButtonPressed = CAF_button									THEN CALL run_from_GitHub(script_repository & "NOTES/NOTES - CAF.vbs")								
IF ButtonPressed = CHANGE_REPORT_FORM_RECEIVED_button			THEN CALL run_from_GitHub(script_repository & "NOTES/NOTES - CHANGE REPORT FORM RECEIVED.vbs")
IF ButtonPressed = CITIZENSHIP_IDENTITY_VERIFIED_button			THEN CALL run_from_GitHub(script_repository & "NOTES/NOTES - CITIZENSHIP-IDENTITY VERIFIED.vbs")		
IF ButtonPressed = CLIENT_CONTACT_CALL_CENTER_VERSION_button	THEN CALL run_from_GitHub(script_repository & "NOTES/NOTES - CLIENT CONTACT (CALL CENTER VERSION).vbs")
IF ButtonPressed = CLIENT_CONTACT_button						THEN CALL run_from_GitHub(script_repository & "NOTES/NOTES - CLIENT CONTACT.vbs")					
IF ButtonPressed = CLOSED_PROGRAMS_button						THEN CALL run_from_GitHub(script_repository & "NOTES/NOTES - CLOSED PROGRAMS.vbs")					
IF ButtonPressed = COMBINED_AR_button							THEN CALL run_from_GitHub(script_repository & "NOTES/NOTES - COMBINED AR.vbs")						
IF ButtonPressed = CSR_button									THEN CALL run_from_GitHub(script_repository & "NOTES/NOTES - CSR.vbs")								
IF ButtonPressed = DENIED_PROGRAMS_button						THEN CALL run_from_GitHub(script_repository & "NOTES/NOTES - DENIED PROGRAMS.vbs")					
IF ButtonPressed = DOCUMENTS_RECEIVED_button					THEN CALL run_from_GitHub(script_repository & "NOTES/NOTES - DOCUMENTS RECEIVED.vbs")				
IF ButtonPressed = DRUG_FELON_button							THEN CALL run_from_GitHub(script_repository & "NOTES/NOTES - DRUG FELON.vbs")
IF ButtonPressed = DWP_BUDGET_button							THEN CALL run_from_GitHub(script_repository & "NOTES/NOTES - DWP BUDGET.vbs")
IF ButtonPressed = EMERGENCY_button								THEN CALL run_from_GitHub(script_repository & "NOTES/NOTES - EMERGENCY.vbs")							
IF ButtonPressed = EXPEDITED_SCREENING_button					THEN CALL run_from_GitHub(script_repository & "NOTES/NOTES - EXPEDITED SCREENING.vbs")				
IF ButtonPressed = FRAUD_INFO_button							THEN CALL run_from_GitHub(script_repository & "NOTES/NOTES - FRAUD INFO.vbs")
IF ButtonPressed = GAS_CARDS_ISSUED_button						THEN CALL run_from_GitHub(script_repository & "NOTES/NOTES - GAS CARDS ISSUED.vbs")
IF ButtonPressed = GRH_HRF_button								THEN CALL run_from_GitHub(script_repository & "NOTES/NOTES - GRH - HRF.vbs")							

'Logging usage stats
script_end_procedure("If you see this, it's because you clicked a button that, for some reason, does not have an outcome in the script. Contact your alpha user to report this bug. Thank you!")