'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - MAIN MENU.vbs"
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

'Declaring the variable names to cut down on the number of arguments that need to be passed through the function.
DIM ButtonPressed 
DIM SIR_instructions_button, dialog_name
DIM number_through_c_notes_button, d_through_g_notes_button, h_through_z_notes_button, ltc_notes_button
DIM DOCUMENT_NOTING_button, APPLYMN_APPLICATION_RECEIVED_button, APPROVED_PROGRAMS_button, BABY_BORN_button, BURIAL_ASSETS_button								
DIM CAF_button, CITIZENSHIP_IDENTITY_VERIFIED_button, CLIENT_CONTACT_CALL_CENTER_VERSION_button, CLIENT_CONTACT_button, CLOSED_PROGRAMS_button							
DIM COMBINED_AR_button, CSR_button, DENIED_PROGRAMS_button, DRUG_FELON_button, DWP_BUDGET_button								
DIM EMERGENCY_button, EMPLOYMENT_PLAN_OR_STATUS_UPDATE_button, EXPEDITED_SCREENING_button, FRAUD_INFO_button, GAS_CARDS_ISSUED_button							
DIM GRH_HRF_button, HC_RENEWAL_button, HCAPP_button, HH_COMP_CHANGE_button, HRF_button										
DIM LEP_SAVE_button, LEP_SPONSOR_INCOME_button						
DIM LTC_APPLICATION_RECEIVED_button, LTC_ASSET_ASSESSMENT_button, LTC_COLA_SUMMARY_2015_button, LTC_INTAKE_APPROVAL_button, LTC_MA_APPROVAL_button							
DIM LTC_RENEWAL_button, LTC_TRANSFER_PENALTY_button, MFIP_SANCTION_AND_DWP_DISQUALIFICATION_button, LTC_1503_button, LTC_5181_button	
DIM MILEAGE_REIMBURSEMENT_REQUEST_button, MNSURE_DOCUMENTS_REQUESTED_button, OVERPAYMENT_button, VERIFICATIONS_NEEDED_button, CHANGE_REPORT_FORM_RECEIVED_button				
DIM DOCUMENTS_RECEIVED_button, MEDICAL_OPINION_FORM_RECEIVED_button, SHELTER_FORM_RECEIVED_button						

'The function that creates the 4 dialogs depending on the dialog_name being sent through.
FUNCTION create_NOTES_main_menu(dialog_name)
	IF dialog_name = "#-C" THEN 
        BeginDialog dialog_name, 0, 0, 516, 270, "# - C NOTES Scripts"
          ButtonGroup ButtonPressed
            PushButton 15, 35, 30, 15, "# - C", number_through_c_notes_button
            PushButton 45, 35, 30, 15, "D - G", d_through_g_notes_button
            PushButton 75, 35, 30, 15, "H - Z", h_through_z_notes_button
            PushButton 105, 35, 30, 15, "LTC", ltc_notes_button
            PushButton 445, 10, 65, 10, "SIR instructions", SIR_instructions_button
            PushButton 5, 70, 105, 10, "ApplyMN application received", APPLYMN_APPLICATION_RECEIVED_button
            PushButton 5, 85, 70, 10, "Approved programs", APPROVED_PROGRAMS_button
            PushButton 5, 100, 45, 10, "Baby born", BABY_BORN_button
            PushButton 5, 115, 50, 10, "Burial assets", BURIAL_ASSETS_button
            PushButton 5, 130, 20, 10, "CAF", CAF_button
            PushButton 5, 145, 95, 10, "Citizenship/identity verified", CITIZENSHIP_IDENTITY_VERIFIED_button
            PushButton 5, 160, 50, 10, "Client contact", CLIENT_CONTACT_button
            PushButton 5, 175, 60, 10, "Closed programs", CLOSED_PROGRAMS_button
            PushButton 5, 190, 50, 10, "Combined AR", COMBINED_AR_button
            PushButton 5, 205, 20, 10, "CSR", CSR_button
            CancelButton 460, 250, 50, 15
          Text 5, 5, 435, 10, "Notes scripts main menu: select the script to run from the choices below. Notes with autofill functionality marked with an asterisk (*)."
          Text 115, 70, 330, 10, "--- Template for documenting details about an ApplyMN application recevied."
          Text 80, 85, 325, 10, "--- Template for when you approve a client's programs."
          Text 55, 100, 270, 10, "--- Template for a baby born and added to household."
          Text 60, 115, 135, 10, "--- Template for burial assets."
          Text 30, 130, 390, 10, "--- Template for when you're processing a CAF. Works for intake as well as recertification and reapplication.*"
          Text 105, 145, 295, 10, "--- Template for documenting citizenship/identity status for a case."
          Text 60, 160, 430, 10, "--- Template for documenting client contact, either from or to a client."
          Text 70, 175, 430, 10, "--- Template for indicating which programs are closing, and when. Also case notes intake/REIN dates based on various selections."
          Text 60, 190, 250, 10, "--- Template for the Combined Annual Renewal.*"
          Text 30, 205, 120, 10, "--- Template for the CSR.*"
          GroupBox 5, 20, 140, 35, "NOTES Sub-Menus"
        EndDialog
	ELSEIF dialog_name = "D-G" THEN
        BeginDialog dialog_name, 0, 0, 516, 270, "D - G NOTES Scripts"
          ButtonGroup ButtonPressed
            PushButton 15, 35, 30, 15, "# - C", number_through_c_notes_button
            PushButton 45, 35, 30, 15, "D - G", d_through_g_notes_button
            PushButton 75, 35, 30, 15, "H - Z", h_through_z_notes_button
            PushButton 105, 35, 30, 15, "LTC", ltc_notes_button
            PushButton 445, 10, 65, 10, "SIR instructions", SIR_instructions_button
            PushButton 5, 70, 60, 10, "Denied programs", DENIED_PROGRAMS_button
            PushButton 5, 85, 40, 10, "Drug felon", DRUG_FELON_button
            PushButton 5, 100, 50, 10, "DWP budget", DWP_BUDGET_button
            PushButton 5, 115, 45, 10, "Emergency", EMERGENCY_button
            PushButton 5, 130, 120, 10, "Employment plan or status update", EMPLOYMENT_PLAN_OR_STATUS_UPDATE_button
            PushButton 5, 145, 75, 10, "Expedited screening", EXPEDITED_SCREENING_button
            PushButton 5, 160, 40, 10, "Fraud info", FRAUD_INFO_button
            PushButton 5, 175, 55, 10, "FSET sanction ", FSET_sanction_button
            PushButton 5, 190, 65, 10, "Gas cards issued", GAS_CARDS_ISSUED_button
            PushButton 5, 205, 45, 10, "GRH - HRF", GRH_HRF_button
            CancelButton 460, 250, 50, 15
          Text 70, 70, 435, 10, "--- Template for indicating which programs you've denied, and when. Also case notes intake/REIN dates based on various selections."
          Text 50, 85, 215, 10, "--- Template for noting drug felon info."
          Text 60, 100, 215, 10, "--- Template for noting DWP budgets."
          Text 55, 115, 240, 10, "--- Template for EA/EGA applications.*"
          Text 130, 130, 345, 10, "--- Template for case noting an employment plan or status update for family cash cases."
          Text 85, 145, 220, 10, "--- Template for screening a client for expedited status."
          Text 50, 160, 200, 10, "--- Template for noting fraud info."
          Text 75, 190, 375, 10, "--- Template for gas card issuance. Consult with a supervisor to make sure this is appropriate for your agency."
          Text 55, 205, 190, 10, "--- Template for GRH HRFs. Case must be post-pay.*"
          Text 5, 5, 435, 10, "Notes scripts main menu: select the script to run from the choices below. Notes with autofill functionality marked with an asterisk (*)."
          Text 65, 175, 395, 10, "--- NEW 07/2015 Template for for imposing or resolving an FSET sanction which will also update the MAXIS WREG panel."
          GroupBox 5, 20, 140, 35, "NOTES Sub-Menus"
        EndDialog
	ELSEIF dialog_name = "H-Z" THEN 
        BeginDialog dialog_name, 0, 0, 516, 270, "Notes (H-Z) scripts main menu dialog"
          ButtonGroup ButtonPressed
            PushButton 15, 35, 30, 15, "# - C", number_through_c_notes_button
            PushButton 45, 35, 30, 15, "D - G", d_through_g_notes_button
            PushButton 75, 35, 30, 15, "H - Z", h_through_z_notes_button
            PushButton 105, 35, 30, 15, "LTC", ltc_notes_button
            PushButton 445, 10, 65, 10, "SIR instructions", SIR_instructions_button
            PushButton 5, 70, 50, 10, "HC Renewal", HC_RENEWAL_button
            PushButton 5, 85, 30, 10, "HCAPP", HCAPP_button
            PushButton 5, 100, 65, 10, "HH comp change", HH_COMP_CHANGE_button
            PushButton 5, 115, 25, 10, "HRF", HRF_button
            PushButton 5, 130, 45, 10, "LEP - SAVE", LEP_SAVE_button
            PushButton 5, 145, 80, 10, "LEP - Sponsor income", LEP_SPONSOR_INCOME_button
            PushButton 5, 160, 125, 10, "MFIP sanction/DWP disqualification", MFIP_SANCTION_AND_DWP_DISQUALIFICATION_button
            PushButton 5, 175, 110, 10, "Mileage reimbursement request", MILEAGE_REIMBURSEMENT_REQUEST_button
            PushButton 5, 190, 110, 10, "MNsure - Documents requested", MNSURE_DOCUMENTS_REQUESTED_button
            PushButton 5, 205, 50, 10, "Overpayment", OVERPAYMENT_button
            PushButton 5, 220, 75, 10, "SNAP case review", SNAP_case_review
            PushButton 5, 235, 75, 10, "Verifications needed", VERIFICATIONS_NEEDED_button
            CancelButton 460, 250, 50, 15
          Text 60, 70, 140, 10, "--- Template for HC renewals.*"
          Text 40, 85, 120, 10, "--- Template for HCAPPs.*"
          Text 75, 100, 240, 10, "--- Template for when you update the HH comp of a case."
          Text 35, 115, 240, 10, "--- Template for HRFs (for GRH, use the ''GRH - HRF'' script).*"
          Text 55, 130, 255, 10, "--- Template for the SAVE system for verifying immigration status."
          Text 90, 145, 345, 10, "--- Template for the sponsor income deeming calculation (it will also help calculate it for you)."
          Text 135, 160, 290, 10, "--- Template for MFIP sanctions and DWP disqualifications, both CS and ES."
          Text 120, 175, 260, 10, "--- Template for actions taken on medical mileage reimbursements."
          Text 120, 190, 250, 10, "--- Template for when MNsure documents have been requested."
          Text 60, 205, 240, 10, "--- Template for noting basic information about overpayments."
          Text 85, 235, 270, 10, "--- Template for when verifications are needed (enters each verification clearly)."
          Text 5, 5, 435, 10, "Notes scripts main menu: select the script to run from the choices below. Notes with autofill functionality marked with an asterisk (*)."
          Text 85, 220, 345, 10, "---NEW 07/2015  Template for SNAP reviewers to use that will case note the status  SNAP quality review."
          GroupBox 5, 20, 140, 35, "NOTES Sub-Menus"
        EndDialog	
	ELSEIF dialog_name = "LTC" THEN 
        BeginDialog dialog_name, 0, 0, 516, 270, "Notes (LTC) scripts main menu dialog"
          ButtonGroup ButtonPressed
            PushButton 15, 35, 30, 15, "# - C", number_through_c_notes_button
            PushButton 45, 35, 30, 15, "D - G", d_through_g_notes_button
            PushButton 75, 35, 30, 15, "H - Z", h_through_z_notes_button
            PushButton 105, 35, 30, 15, "LTC", ltc_notes_button
            PushButton 5, 70, 45, 10, "LTC - 1503", LTC_1503_button
            PushButton 5, 85, 45, 10, "LTC - 5181", LTC_5181_button
            PushButton 445, 10, 65, 10, "SIR instructions", SIR_instructions_button
            PushButton 5, 100, 90, 10, "LTC - Application received", LTC_APPLICATION_RECEIVED_button
            PushButton 5, 115, 85, 10, "LTC - Asset assessment", LTC_ASSET_ASSESSMENT_button
            PushButton 5, 130, 95, 10, "LTC - COLA summary 2015", LTC_COLA_SUMMARY_2015_button
            PushButton 5, 145, 75, 10, "LTC - Intake approval", LTC_INTAKE_APPROVAL_button
            PushButton 5, 160, 65, 10, "LTC - MA approval", LTC_MA_APPROVAL_button
            PushButton 5, 175, 55, 10, "LTC - Renewal", LTC_RENEWAL_button
            PushButton 5, 190, 80, 10, "LTC - Transfer penalty", LTC_TRANSFER_PENALTY_button
            CancelButton 460, 245, 50, 15
          Text 55, 70, 130, 10, "--- Template for processing DHS-1503."
          Text 55, 85, 180, 10, "--- NEW 06/2015!!! Template for processing DHS-5181."
          Text 100, 100, 205, 10, "--- Template for initial details of a LTC application.*"
          Text 95, 115, 340, 10, "--- Template for the LTC asset assessment. Will enter both person and case notes if desired."
          Text 105, 130, 250, 10, "--- Template to summarize actions for the 2015 COLA.*"
          Text 85, 145, 205, 10, "--- Template for use when approving a LTC intake.*"
          Text 75, 160, 355, 10, "--- Template for approving LTC MA (can be used for changes, initial application, or recertification).*"
          Text 65, 175, 140, 10, "--- Template for LTC renewals.*"
          Text 90, 190, 225, 10, "--- Template for noting a transfer penalty."
          Text 5, 5, 435, 10, "Notes scripts main menu: select the script to run from the choices below. Notes with autofill functionality marked with an asterisk (*)."
          GroupBox 5, 20, 140, 35, "NOTES Sub-Menus"
        EndDialog
	END IF

	DIALOG dialog_name

END FUNCTION

'=====THE SCRIPT=====
EMConnect ""

'Setting the default menu to the #-C notes
dialog_name = "#-C"
DO
	'Calling the function that loads the dialogs
	CALL create_NOTES_main_menu(dialog_name)
		IF ButtonPressed = 0 THEN stopscript
		'Opening the SIR Instructions
		IF buttonpressed = SIR_instructions_button then CreateObject("WScript.Shell").Run("https://www.dhssir.cty.dhs.state.mn.us/MAXIS/blzn/scriptwiki/Wiki%20Pages/Bulk%20scripts.aspx")

		'If the user selects the other sub-menu, the script do-loops with the new dialog_name
		IF ButtonPressed = number_through_c_notes_button THEN 
			dialog_name = "#-C"
		ELSEIF ButtonPressed = d_through_g_notes_button THEN 
			dialog_name = "D-G"
		ELSEIF ButtonPressed = h_through_z_notes_button THEN
			dialog_name = "H-Z"
		ELSEIF ButtonPressed = LTC_notes_button THEN 
			dialog_name = "LTC"
		END IF

		'If the user selects a script button, the script will exit the do-loop
LOOP UNTIL ButtonPressed <> SIR_instructions_button AND ButtonPressed <> number_through_c_notes_button AND ButtonPressed <> d_through_g_notes_button AND ButtonPressed <> h_through_z_notes_button AND ButtonPressed <> LTC_notes_button

'Available scripts
IF ButtonPressed = APPLYMN_APPLICATION_RECEIVED_button				THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - APPLYMN APPLICATION RECEIVED.vbs")		
IF ButtonPressed = APPROVED_PROGRAMS_button							THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - APPROVED PROGRAMS.vbs")					
IF ButtonPressed = BABY_BORN_button									THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - BABY BORN.vbs")
IF ButtonPressed = BURIAL_ASSETS_button								THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - BURIAL ASSETS.vbs")						
IF ButtonPressed = CAF_button										THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - CAF.vbs")								
IF ButtonPressed = CHANGE_REPORT_FORM_RECEIVED_button				THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - CHANGE REPORT FORM RECEIVED.vbs")
IF ButtonPressed = CITIZENSHIP_IDENTITY_VERIFIED_button				THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - CITIZENSHIP-IDENTITY VERIFIED.vbs")		
IF ButtonPressed = CLIENT_CONTACT_CALL_CENTER_VERSION_button		THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - CLIENT CONTACT (CALL CENTER VERSION).vbs")
IF ButtonPressed = CLIENT_CONTACT_button							THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - CLIENT CONTACT.vbs")					
IF ButtonPressed = CLOSED_PROGRAMS_button							THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - CLOSED PROGRAMS.vbs")					
IF ButtonPressed = COMBINED_AR_button								THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - COMBINED AR.vbs")						
IF ButtonPressed = CSR_button										THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - CSR.vbs")								
IF ButtonPressed = DENIED_PROGRAMS_button							THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - DENIED PROGRAMS.vbs")					
IF ButtonPressed = DRUG_FELON_button								THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - DRUG FELON.vbs")
IF ButtonPressed = DWP_BUDGET_button								THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - DWP BUDGET.vbs")
IF ButtonPressed = EMERGENCY_button									THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - EMERGENCY.vbs")						
IF ButtonPressed = EMPLOYMENT_PLAN_OR_STATUS_UPDATE_button			THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - EMPLOYMENT PLAN OR STATUS UPDATE.vbs")
IF ButtonPressed = EXPEDITED_SCREENING_button						THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - EXPEDITED SCREENING.vbs")				
IF ButtonPressed = FRAUD_INFO_button								THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - FRAUD INFO.vbs")
IF ButtonPressed = GAS_CARDS_ISSUED_button							THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - GAS CARDS ISSUED.vbs")
IF ButtonPressed = GRH_HRF_button									THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - GRH - HRF.vbs")							
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
IF ButtonPressed = DOCUMENTS_RECEIVED_button						THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - DOCUMENTS RECEIVED.vbs")				
IF ButtonPressed = LTC_1503_button									THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - LTC - 1503.vbs")
IF ButtonPressed = LTC_5181_button									THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - LTC - 5181.vbs")
IF ButtonPressed = MEDICAL_OPINION_FORM_RECEIVED_button				THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - MEDICAL OPINION FORM RECEIVED.vbs")
IF ButtonPressed = SHELTER_FORM_RECEIVED_button						THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - SHELTER FORM RECEIVED.vbs")

'Logging usage stats
script_end_procedure("If you see this, it's because you clicked a button that, for some reason, does not have an outcome in the script. Contact your alpha user to report this bug. Thank you!")
