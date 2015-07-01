'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - MAIN MENU (0-G).vbs"
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
BeginDialog NOTES_0_G_scripts_main_menu_dialog, 0, 0, 516, 350, "Notes (0-G) scripts main menu dialog"
  ButtonGroup ButtonPressed
    CancelButton 460, 330, 50, 15
    PushButton 445, 10, 65, 10, "SIR instructions", SIR_instructions_button
    PushButton 5, 20, 60, 10, "Document noting", DOCUMENT_NOTING_button
    PushButton 5, 45, 105, 10, "ApplyMN application received", APPLYMN_APPLICATION_RECEIVED_button
    PushButton 5, 60, 70, 10, "Approved programs", APPROVED_PROGRAMS_button
    PushButton 5, 75, 45, 10, "Baby born", BABY_BORN_button
    PushButton 5, 90, 50, 10, "Burial assets", BURIAL_ASSETS_button
    PushButton 5, 105, 20, 10, "CAF", CAF_button
    PushButton 5, 120, 95, 10, "Citizenship/identity verified", CITIZENSHIP_IDENTITY_VERIFIED_button
    PushButton 5, 135, 50, 10, "Client contact", CLIENT_CONTACT_button
    PushButton 5, 150, 60, 10, "Closed programs", CLOSED_PROGRAMS_button
    PushButton 5, 165, 50, 10, "Combined AR", COMBINED_AR_button
    PushButton 5, 180, 20, 10, "CSR", CSR_button
    PushButton 5, 195, 60, 10, "Denied programs", DENIED_PROGRAMS_button
    PushButton 5, 210, 40, 10, "Drug felon", DRUG_FELON_button
    PushButton 5, 225, 50, 10, "DWP budget", DWP_BUDGET_button
    PushButton 5, 240, 45, 10, "Emergency", EMERGENCY_button
    PushButton 5, 255, 120, 10, "Employment plan or status update", EMPLOYMENT_PLAN_OR_STATUS_UPDATE_button
    PushButton 5, 270, 75, 10, "Expedited screening", EXPEDITED_SCREENING_button
    PushButton 5, 285, 40, 10, "Fraud info", FRAUD_INFO_button
    PushButton 5, 300, 65, 10, "Gas cards issued", GAS_CARDS_ISSUED_button
    PushButton 5, 315, 45, 10, "GRH - HRF", GRH_HRF_button
  Text 5, 5, 435, 10, "Notes scripts main menu: select the script to run from the choices below. Notes with autofill functionality marked with an asterisk (*)."
  Text 75, 20, 330, 10, "--- NEW 06/2015!!! Templates that are for use when various documents are received."
  Text 115, 45, 330, 10, "--- Template for documenting details about an ApplyMN application recevied."
  Text 80, 60, 325, 10, "--- Template for when you approve a client's programs."
  Text 55, 75, 270, 10, "--- Template for a baby born and added to household."
  Text 60, 90, 135, 10, "--- Template for burial assets."
  Text 30, 105, 390, 10, "--- Template for when you're processing a CAF. Works for intake as well as recertification and reapplication.*"
  Text 105, 120, 295, 10, "--- Template for documenting citizenship/identity status for a case."
  Text 60, 135, 430, 10, "--- Template for documenting client contact, either from or to a client. MERGED WITH CALL CENTER VERSION 06/2015!!!"
  Text 70, 150, 430, 10, "--- Template for indicating which programs are closing, and when. Also case notes intake/REIN dates based on various selections."
  Text 60, 165, 250, 10, "--- Template for the Combined Annual Renewal.*"
  Text 30, 180, 120, 10, "--- Template for the CSR.*"
  Text 70, 195, 435, 10, "--- Template for indicating which programs you've denied, and when. Also case notes intake/REIN dates based on various selections."
  Text 50, 210, 215, 10, "--- Template for noting drug felon info."
  Text 60, 225, 215, 10, "--- Template for noting DWP budgets."
  Text 55, 240, 240, 10, "--- Template for EA/EGA applications.*"
  Text 130, 255, 345, 10, "--- NEW 06/2015!!! Template for case noting an employment plan or status update for family cash cases."
  Text 85, 270, 220, 10, "--- Template for screening a client for expedited status."
  Text 50, 285, 200, 10, "--- Template for noting fraud info."
  Text 75, 300, 375, 10, "--- Template for gas card issuance. Consult with a supervisor to make sure this is appropriate for your agency."
  Text 55, 315, 190, 10, "--- Template for GRH HRFs. Case must be post-pay.*"
EndDialog

'Variables to declare
IF script_repository = "" THEN script_repository = "https://raw.githubusercontent.com/MN-Script-Team/DHS-MAXIS-Scripts/master/Script Files"		'If it's blank, we're assuming the user is a scriptwriter, ergo, master branch.

'THE SCRIPT----------------------------------------------------------------------------------------------------

'Shows main menu dialog, which asks user which script to run. Loops until a button other than the SIR instructions button is clicked.
Do
	Dialog NOTES_0_G_scripts_main_menu_dialog
	IF ButtonPressed = cancel THEN StopScript
	If buttonpressed = SIR_instructions_button then CreateObject("WScript.Shell").Run("https://www.dhssir.cty.dhs.state.mn.us/MAXIS/blzn/scriptwiki/Wiki%20Pages/Notes%20scripts.aspx")
Loop until buttonpressed <> SIR_instructions_button

'Connecting to BlueZone
EMConnect ""

IF ButtonPressed = DOCUMENT_NOTING_button						THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - DOCUMENT NOTING MENU.vbs")				
IF ButtonPressed = APPLYMN_APPLICATION_RECEIVED_button			THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - APPLYMN APPLICATION RECEIVED.vbs")		
IF ButtonPressed = APPROVED_PROGRAMS_button						THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - APPROVED PROGRAMS.vbs")					
IF ButtonPressed = BABY_BORN_button								THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - BABY BORN.vbs")
IF ButtonPressed = BURIAL_ASSETS_button							THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - BURIAL ASSETS.vbs")						
IF ButtonPressed = CAF_button									THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - CAF.vbs")								
IF ButtonPressed = CHANGE_REPORT_FORM_RECEIVED_button			THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - CHANGE REPORT FORM RECEIVED.vbs")
IF ButtonPressed = CITIZENSHIP_IDENTITY_VERIFIED_button			THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - CITIZENSHIP-IDENTITY VERIFIED.vbs")		
IF ButtonPressed = CLIENT_CONTACT_CALL_CENTER_VERSION_button	THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - CLIENT CONTACT (CALL CENTER VERSION).vbs")
IF ButtonPressed = CLIENT_CONTACT_button						THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - CLIENT CONTACT.vbs")					
IF ButtonPressed = CLOSED_PROGRAMS_button						THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - CLOSED PROGRAMS.vbs")					
IF ButtonPressed = COMBINED_AR_button							THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - COMBINED AR.vbs")						
IF ButtonPressed = CSR_button									THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - CSR.vbs")								
IF ButtonPressed = DENIED_PROGRAMS_button						THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - DENIED PROGRAMS.vbs")					
IF ButtonPressed = DRUG_FELON_button							THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - DRUG FELON.vbs")
IF ButtonPressed = DWP_BUDGET_button							THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - DWP BUDGET.vbs")
IF ButtonPressed = EMERGENCY_button								THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - EMERGENCY.vbs")						
IF ButtonPressed = EMPLOYMENT_PLAN_OR_STATUS_UPDATE_button		THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - EMPLOYMENT PLAN OR STATUS UPDATE.vbs")
IF ButtonPressed = EXPEDITED_SCREENING_button					THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - EXPEDITED SCREENING.vbs")				
IF ButtonPressed = FRAUD_INFO_button							THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - FRAUD INFO.vbs")
IF ButtonPressed = GAS_CARDS_ISSUED_button						THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - GAS CARDS ISSUED.vbs")
IF ButtonPressed = GRH_HRF_button								THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - GRH - HRF.vbs")							

'Logging usage stats
script_end_procedure("If you see this, it's because you clicked a button that, for some reason, does not have an outcome in the script. Contact your alpha user to report this bug. Thank you!")