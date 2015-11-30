'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - MAIN MENU.vbs"
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

'Declaring the variable names to cut down on the number of arguments that need to be passed through the function.
DIM ButtonPressed
DIM SIR_instructions_button, dialog_name
DIM number_through_c_notes_button, d_through_f_notes_button, g_through_l_notes_button, m_through_q_notes_button, r_through_z_notes_button, ltc_notes_button

DIM APPLICATION_RECEIVED_button, APPROVED_PROGRAMS_button
DIM BABY_BORN_button, BURIAL_ASSETS_button
DIM CAF_button, CHANGE_REPORT_FORM_RECEIVED_button, CITIZENSHIP_IDENTITY_VERIFIED_button, CLIENT_CONTACT_button, CLOSED_PROGRAMS_button, COMBINED_AR_button, CSR_button, COUNTY_BURIAL_button

DIM DENIED_PROGRAMS_button, DOCUMENTS_RECEIVED_button, DRUG_FELON_button, DWP_BUDGET_button
DIM EMERGENCY_button, EMPLOYMENT_PLAN_OR_STATUS_UPDATE_button, EMPLOYMENT_VERIFICATION_button, ES_REFERRAL_button, EXPEDITED_SCREENING_button, FRAUD_INFO_button

DIM GAS_CARDS_ISSUED_button, GOOD_CAUSE_CLAIMED_button, GRH_HRF_button
DIM HC_RENEWAL_button, HCAPP_button, HH_COMP_CHANGE_button, HRF_button, IEVS_NOTICE_RECEIVED_button, INCARCERATION_button, INTERVIEW_COMPLETED_button
DIM LEP_EMA_button, LEP_SAVE_button, LEP_SPONSOR_INCOME_button, LOBBY_NO_SHOW_button

DIM MEDICAL_OPINION_FORM_RECEIVED_button, MFIP_SANCTION_AND_DWP_DISQUALIFICATION_button, MFIP_SANCTION_CURED_button, MFIP_TO_SNAP_TRANSITION_button, MILEAGE_REIMBURSEMENT_REQUEST_button, MNSURE_DOCUMENTS_REQUESTED_button, MNSURE_RETRO_HC_APPLICATION_button
DIM OVERPAYMENT_button, MTAF_button
DIM PREGNANCY_REPORTED_button, PROOF_OF_RELATIONSHIP_button
DIM REIN_PROGS_button
DIM SHELTER_FORM_RECEIVED_button, SNAP_CASE_REVIEW_button, SUBMIT_CASE_FOR_SNAP_REVIEW_button
DIM VERIFICATIONS_NEEDED_button

DIM LTC_APPLICATION_RECEIVED_button, LTC_ASSET_ASSESSMENT_button, LTC_COLA_SUMMARY_2016_button, LTC_INTAKE_APPROVAL_button, LTC_MA_APPROVAL_button
DIM LTC_RENEWAL_button, LTC_TRANSFER_PENALTY_button, LTC_1503_button, LTC_5181_button

'The function that creates the 4 dialogs depending on the dialog_name being sent through.
FUNCTION create_NOTES_main_menu(dialog_name)
	IF dialog_name = "#-C" THEN
        BeginDialog dialog_name, 0, 0, 516, 270, "# - C NOTES Scripts"
          Text 5, 5, 435, 10, "Notes scripts main menu: select the script to run from the choices below. Notes with autofill functionality marked with an asterisk (*)."
		  GroupBox 5, 20, 205, 35, "NOTES Sub-Menus"
		  ButtonGroup ButtonPressed
			PushButton 15, 35, 30, 15, "# - C", number_through_c_notes_button
			PushButton 45, 35, 30, 15, "D - F", d_through_f_notes_button
			PushButton 75, 35, 30, 15, "G - L", g_through_l_notes_button
			PushButton 105, 35, 30, 15, "M - Q", m_through_q_notes_button
			PushButton 135, 35, 30, 15, "R - Z", r_through_z_notes_button
            PushButton 165, 35, 30, 15, "LTC", ltc_notes_button
          	PushButton 445, 10, 65, 10, "SIR instructions", SIR_instructions_button
          	PushButton 5, 70, 80, 10, "Application received", APPLICATION_RECEIVED_button
          	PushButton 5, 85, 70, 10, "Approved programs", APPROVED_PROGRAMS_button
          	PushButton 5, 100, 45, 10, "Baby born", BABY_BORN_button
          	PushButton 5, 115, 50, 10, "Burial assets", BURIAL_ASSETS_button
          	PushButton 5, 130, 20, 10, "CAF", CAF_button
          	PushButton 5, 145, 105, 10, "Change Report Form Received", CHANGE_REPORT_FORM_RECEIVED_button
          	PushButton 5, 160, 95, 10, "Citizenship/identity verified", CITIZENSHIP_IDENTITY_VERIFIED_button
          	PushButton 5, 175, 50, 10, "Client contact", CLIENT_CONTACT_button
          	PushButton 5, 190, 60, 10, "Closed programs", CLOSED_PROGRAMS_button
          	PushButton 5, 205, 50, 10, "Combined AR", COMBINED_AR_button
          	PushButton 5, 220, 20, 10, "CSR", CSR_button
			PushButton 5, 235, 85, 10, "County Burial Application", COUNTY_BURIAL_button
          	CancelButton 460, 245, 50, 15
          Text 90, 70, 330, 10, "--- Template for documenting details about an application recevied."
          Text 80, 85, 325, 10, "--- Template for when you approve a client's programs."
          Text 55, 100, 270, 10, "--- Template for a baby born and added to household."
          Text 60, 115, 135, 10, "--- Template for burial assets."
          Text 30, 130, 390, 10, "--- Template for when you're processing a CAF. Works for intake as well as recertification and reapplication.*"
          Text 120, 145, 255, 10, "--- Template for case noting information reported from a Change Report Form."
          Text 105, 160, 295, 10, "--- Template for documenting citizenship/identity status for a case."
          Text 60, 175, 430, 10, "--- Template for documenting client contact, either from or to a client."
          Text 70, 190, 430, 10, "--- Template for indicating which programs are closing, and when. Also case notes intake/REIN dates based on various selections."
          Text 60, 205, 250, 10, "--- Template for the Combined Annual Renewal.*"
          Text 30, 220, 120, 10, "--- Template for the CSR.*"
		  Text 95, 235, 230, 10, "---- Template for a application for paying funeral costs. "
        EndDialog
	ELSEIF dialog_name = "D-F" THEN
        BeginDialog dialog_name, 0, 0, 516, 270, "D - F NOTES Scripts"
          Text 5, 5, 435, 10, "Notes scripts main menu: select the script to run from the choices below. Notes with autofill functionality marked with an asterisk (*)."
          	GroupBox 5, 20, 205, 35, "NOTES Sub-Menus"
			ButtonGroup ButtonPressed
			PushButton 15, 35, 30, 15, "# - C", number_through_c_notes_button
			PushButton 45, 35, 30, 15, "D - F", d_through_f_notes_button
			PushButton 75, 35, 30, 15, "G - L", g_through_l_notes_button
			PushButton 105, 35, 30, 15, "M - Q", m_through_Q_notes_button
			PushButton 135, 35, 30, 15, "R - Z", r_through_z_notes_button
            PushButton 165, 35, 30, 15, "LTC", ltc_notes_button
            PushButton 445, 10, 65, 10, "SIR instructions", SIR_instructions_button
            PushButton 5, 70, 60, 10, "Denied programs", DENIED_PROGRAMS_button
            PushButton 5, 85, 55, 10, "Docs Received", DOCUMENTS_RECEIVED_button
            PushButton 5, 100, 40, 10, "Drug felon", DRUG_FELON_button
            PushButton 5, 115, 50, 10, "DWP budget", DWP_BUDGET_button
            PushButton 5, 130, 45, 10, "Emergency", EMERGENCY_button
            PushButton 5, 145, 120, 10, "Employment plan or status update", EMPLOYMENT_PLAN_OR_STATUS_UPDATE_button
            PushButton 5, 160, 90, 10, "Employment Verif Recv'd", EMPLOYMENT_VERIFICATION_button
            PushButton 5, 175, 45, 10, "ES Referral", ES_REFERRAL_button
	    PushButton 5, 190, 75, 10, "Expedited screening", EXPEDITED_SCREENING_button
       	    PushButton 5, 205, 40, 10, "Fraud info", FRAUD_INFO_button
	    CancelButton 460, 245, 50, 15
   	Text 70, 70, 435, 10, "--- Template for indicating which programs you've denied, and when. Also case notes intake/REIN dates based on various selections."
        Text 65, 85, 440, 10, "--- Template for case noting information about documents received."
	Text 50, 100, 215, 10, "--- Template for noting drug felon info."
	Text 60, 115, 215, 10, "--- Template for noting DWP budgets."
  	Text 55, 130, 240, 10, "--- Template for EA/EGA applications.*"
	Text 130, 145, 345, 10, "--- Template for case noting an employment plan or status update for family cash cases."
	Text 100, 160, 370, 10, "--- Template for noting information about an employment verification received by the agency."
	Text 55, 175, 245, 10, "--- New 11/2015!!! Template for sending an MFIP or DWP referral to employment services."
  	Text 85, 190, 220, 10, "--- Template for screening a client for expedited status."
	Text 50, 205, 200, 10, "--- Template for noting fraud info."
        EndDialog
	ELSEIF dialog_name = "G-L" THEN
		BeginDialog dialog_name, 0, 0, 516, 270, "Notes (G-L) scripts main menu dialog"
		ButtonGroup ButtonPressed
			CancelButton 455, 245, 50, 15
			PushButton 445, 10, 65, 10, "SIR instructions", SIR_instructions_button
			PushButton 15, 35, 30, 15, "# - C", number_through_c_notes_button
			PushButton 45, 35, 30, 15, "D - F", d_through_f_notes_button
			PushButton 75, 35, 30, 15, "G - L", g_through_l_notes_button
			PushButton 105, 35, 30, 15, "M - Q", m_through_Q_notes_button
			PushButton 135, 35, 30, 15, "R - Z", r_through_z_notes_button
			PushButton 165, 35, 30, 15, "LTC", ltc_notes_button
			PushButton 5, 60, 65, 10, "Gas cards issued", GAS_CARDS_ISSUED_button
			PushButton 5, 75, 75, 10, "Good Cause Claimed", GOOD_CAUSE_CLAIMED_button
			PushButton 5, 90, 45, 10, "GRH - HRF", GRH_HRF_button
			PushButton 5, 105, 50, 10, "HC Renewal", HC_RENEWAL_button
			PushButton 5, 120, 30, 10, "HCAPP", HCAPP_button
			PushButton 5, 135, 65, 10, "HH comp change", HH_COMP_CHANGE_button
			PushButton 5, 150, 25, 10, "HRF", HRF_button
			PushButton 5, 165, 80, 10, "IEVS Notice Received", IEVS_NOTICE_RECEIVED_button
			PushButton 5, 180, 50, 10, " Incarceration ", Incarceration_button
			PushButton 5, 195, 75, 10, "Interview Completed", INTERVIEW_COMPLETED_button
			PushButton 5, 210, 45, 10, "LEP - EMA", LEP_EMA_button
			PushButton 5, 225, 45, 10, "LEP - SAVE", LEP_SAVE_button
			PushButton 5, 240, 80, 10, "LEP - Sponsor income", LEP_SPONSOR_INCOME_button
			PushButton 5, 255, 60, 10, "Lobby No Show", LOBBY_NO_SHOW_button
		Text 5, 5, 435, 10, "Notes scripts main menu: select the script to run from the choices below. Notes with autofill functionality marked with an asterisk (*)."
		Text 75, 60, 375, 10, "--- Template for gas card issuance. Consult with a supervisor to make sure this is appropriate for your agency."
		Text 85, 75, 235, 10, "--- Template for requests of good cause to not receive child support."
		Text 55, 90, 190, 10, "--- Template for GRH HRFs. Case must be post-pay.*"
		Text 60, 105, 140, 10, "--- Template for HC renewals.*"
		Text 40, 120, 120, 10, "--- Template for HCAPPs.*"
		Text 75, 135, 240, 10, "--- Template for when you update the HH comp of a case."
		Text 35, 150, 240, 10, "--- Template for HRFs (for GRH, use the ''GRH - HRF'' script).*"
		Text 95, 165, 215, 10, "--- Template to case note when a IEVS notice is returned."
		Text 60, 180, 290, 10, "--- Template to note details of an incarceration, and also updates STAT/FACI if necessary."
		Text 90, 195, 325, 10, "--- Template to case note an interview being completed but no stat panels updated."  
		Text 90, 195, 325, 10, "--- Template to case note an interview being completed but no stat panels updated."
		Text 55, 210, 255, 10, "--- Template for EMA applications."
		Text 55, 225, 255, 10, "--- Template for the SAVE system for verifying immigration status."
		Text 90, 240, 310, 10, "--- Template for the sponsor income deeming calculation (it will also help calculate it for you)."
		Text 70, 255, 320, 10, "--- Template for case noting a client's no-showing their in-office appointment."
		GroupBox 5, 20, 205, 35, "NOTES Sub-Menus"
		EndDialog
	ELSEIF dialog_name = "M-Q" THEN
		BeginDialog dialog_name, 0, 0, 516, 270, "Notes (M-Q) scripts main menu dialog"
		ButtonGroup ButtonPressed
			PushButton 445, 10, 65, 10, "SIR instructions", SIR_instructions_button
			PushButton 15, 35, 30, 15, "# - C", number_through_c_notes_button
			PushButton 45, 35, 30, 15, "D - F", d_through_f_notes_button
			PushButton 75, 35, 30, 15, "G - L", g_through_l_notes_button
			PushButton 105, 35, 30, 15, "M - Q", m_through_Q_notes_button
			PushButton 135, 35, 30, 15, "R - Z", r_through_z_notes_button
			PushButton 165, 35, 30, 15, "LTC", ltc_notes_button
			PushButton 5, 70, 110, 10, "Medical Opinion Form Received", MEDICAL_OPINION_FORM_RECEIVED_button
			PushButton 5, 85, 80, 10, "MFIP Sanction Cured", MFIP_SANCTION_CURED_button
			PushButton 5, 100, 125, 10, "MFIP sanction/DWP disqualification", MFIP_SANCTION_AND_DWP_DISQUALIFICATION_button
			PushButton 5, 115, 90, 10, "MFIP to SNAP Transition", MFIP_TO_SNAP_TRANSITION_button
			PushButton 5, 130, 110, 10, "Mileage reimbursement request", MILEAGE_REIMBURSEMENT_REQUEST_button
			PushButton 5, 145, 110, 10, "MNsure - Documents requested", MNSURE_DOCUMENTS_REQUESTED_button
			PushButton 5, 160, 105, 10, "MNsure - Retro HC Application", MNSURE_RETRO_HC_APPLICATION_button
			PushButton 5, 175, 30, 10, "MTAF", MTAF_button
			PushButton 5, 190, 50, 10, "Overpayment", OVERPAYMENT_button
			PushButton 5, 205, 75, 10, "Pregnancy Reported", PREGNANCY_REPORTED_button
			PushButton 5, 220, 70, 10, "Proof of relationship", PROOF_OF_RELATIONSHIP_button
		Text 5, 5, 435, 10, "Notes scripts main menu: select the script to run from the choices below. Notes with autofill functionality marked with an asterisk (*)."
		GroupBox 5, 20, 205, 35, "NOTES Sub-Menus"
		Text 120, 70, 335, 10, "--- Template for case noting information about a Medical Opinion Form."
		Text 90, 85, 350, 10, "--- Template for noting the curing of an MFIP sanction."
		Text 135, 100, 290, 10, "--- Template for MFIP sanctions and DWP disqualifications, both CS and ES."
		Text 100, 115, 265, 10, "--- Template for noting when closing MFIP and opening SNAP."
		Text 120, 130, 260, 10, "--- Template for actions taken on medical mileage reimbursements."
		Text 120, 145, 250, 10, "--- Template for when MNsure documents have been requested."
		Text 115, 160, 340, 10, "--- Template for when MNsure retro HC has been requested."
		Text 60, 190, 240, 10, "--- Template for noting basic information about overpayments."
		Text 85, 205, 405, 10, "--- Template for case noting a pregnancy. This script can update STAT/PREG."
		Text 85, 220, 415, 10, "--- Template for documenting proof of relationship between a member 01 and someone else in the household."
		ButtonGroup ButtonPressed
			CancelButton 460, 245, 50, 15
		Text 45, 175, 290, 10, "--- NEW 10/2015!!!! Template for case noting an MTAF (MFIP Transition Application Form)."
		EndDialog
	ELSEIF dialog_name = "R-Z" THEN
		BeginDialog dialog_name, 0, 0, 516, 270, "Notes (R-Z) scripts main menu dialog"
			Text 5, 5, 435, 10, "Notes scripts main menu: select the script to run from the choices below. Notes with autofill functionality marked with an asterisk (*)."
			GroupBox 5, 20, 205, 35, "NOTES Sub-Menus"
			ButtonGroup ButtonPressed
			  PushButton 15, 35, 30, 15, "# - C", number_through_c_notes_button
		      PushButton 45, 35, 30, 15, "D - F", d_through_f_notes_button
		      PushButton 75, 35, 30, 15, "G - L", g_through_l_notes_button
		      PushButton 105, 35, 30, 15, "M - Q", m_through_Q_notes_button
		      PushButton 135, 35, 30, 15, "R - Z", r_through_z_notes_button
		      PushButton 165, 35, 30, 15, "LTC", ltc_notes_button
		      PushButton 445, 10, 65, 10, "SIR instructions", SIR_instructions_button
		      PushButton 5, 70, 30, 10, "REIN", REIN_PROGS_button
		      PushButton 5, 85, 80, 10, "Shelter Form Received", SHELTER_FORM_RECEIVED_button
		      PushButton 5, 100, 70, 10, "SNAP case review", SNAP_CASE_REVIEW_button
		      PushButton 5, 115, 100, 10, "Submit case for SNAP review", SUBMIT_CASE_FOR_SNAP_REVIEW_button
		      PushButton 5, 130, 75, 10, "Verifications needed", VERIFICATIONS_NEEDED_button
		    CancelButton 460, 245, 50, 15
			Text 40, 70, 415, 10, "--- Template for noting program reinstatement information."
			Text 90, 85, 350, 10, "--- Template for case noting information about a Shelter Form received."
			Text 80, 100, 365, 10, "--- Template for SNAP reviewers to use that will case note the status  SNAP quality review."
			Text 85, 130, 270, 10, "--- Template for when verifications are needed (enters each verification clearly)."
			Text 110, 115, 345, 10, "--- Template for when a worker sends a case to be reviewed by a supervisor or coworker."
		EndDialog
	ELSEIF dialog_name = "LTC" THEN
       BeginDialog dialog_name, 0, 0, 516, 270, "Notes (LTC) scripts main menu dialog"
         Text 5, 5, 435, 10, "Notes scripts main menu: select the script to run from the choices below. Notes with autofill functionality marked with an asterisk (*).		"
		  GroupBox 5, 20, 205, 35, "NOTES Sub-Menus"
		  ButtonGroup ButtonPressed
			PushButton 15, 35, 30, 15, "# - C", number_through_c_notes_button
			PushButton 45, 35, 30, 15, "D - F", d_through_f_notes_button
			PushButton 75, 35, 30, 15, "G - L", g_through_l_notes_button
			PushButton 105, 35, 30, 15, "M - Q", m_through_Q_notes_button
			PushButton 135, 35, 30, 15, "R - Z", r_through_z_notes_button
           PushButton 165, 35, 30, 15, "LTC", ltc_notes_button
           PushButton 445, 10, 65, 10, "SIR instructions", SIR_instructions_button
           PushButton 5, 70, 45, 10, "LTC - 1503", LTC_1503_button
           PushButton 5, 85, 45, 10, "LTC - 5181", LTC_5181_button
           PushButton 5, 100, 90, 10, "LTC - Application received", LTC_APPLICATION_RECEIVED_button
           PushButton 5, 115, 85, 10, "LTC - Asset assessment", LTC_ASSET_ASSESSMENT_button
           PushButton 5, 130, 95, 10, "LTC - COLA summary 2016", LTC_COLA_SUMMARY_2016_button
           PushButton 5, 145, 75, 10, "LTC - Intake approval", LTC_INTAKE_APPROVAL_button
           PushButton 5, 160, 65, 10, "LTC - MA approval", LTC_MA_APPROVAL_button
           PushButton 5, 175, 55, 10, "LTC - Renewal", LTC_RENEWAL_button
           PushButton 5, 190, 80, 10, "LTC - Transfer penalty", LTC_TRANSFER_PENALTY_button
           CancelButton 460, 245, 50, 15
         Text 55, 70, 130, 10, "--- Template for processing DHS-1503."
         Text 55, 85, 180, 10, "--- Template for processing DHS-5181."
         Text 100, 100, 205, 10, "--- Template for initial details of a LTC application.*"
         Text 95, 115, 340, 10, "--- Template for the LTC asset assessment. Will enter both person and case notes if desired."
         Text 105, 130, 250, 10, "--- Template to summarize actions for the 2016 COLA.*"
         Text 85, 145, 205, 10, "--- Template for use when approving a LTC intake.*"
         Text 75, 160, 355, 10, "--- Template for approving LTC MA (can be used for changes, initial application, or recertification).*"
         Text 65, 175, 140, 10, "--- Template for LTC renewals.*"
         Text 90, 190, 225, 10, "--- Template for noting a transfer penalty."
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
		IF buttonpressed = SIR_instructions_button then CreateObject("WScript.Shell").Run("https://www.dhssir.cty.dhs.state.mn.us/MAXIS/blzn/Script%20Instructions%20Wiki/Notes%20scripts.aspx")

		'If the user selects the other sub-menu, the script do-loops with the new dialog_name
		IF ButtonPressed = number_through_c_notes_button THEN
			dialog_name = "#-C"
		ELSEIF ButtonPressed = d_through_f_notes_button THEN
			dialog_name = "D-F"
		ELSEIF ButtonPressed = g_through_l_notes_button THEN
			dialog_name = "G-L"
		ELSEIF ButtonPressed = m_through_q_notes_button THEN
			dialog_name = "M-Q"
		ELSEIF ButtonPressed = r_through_z_notes_button THEN
			dialog_name = "R-Z"
		ELSEIF ButtonPressed = LTC_notes_button THEN
			dialog_name = "LTC"
		END IF

		'If the user selects a script button, the script will exit the do-loop
LOOP UNTIL ButtonPressed <> SIR_instructions_button AND _
			ButtonPressed <> number_through_c_notes_button AND _
			ButtonPressed <> d_through_f_notes_button AND _
			ButtonPressed <> g_through_l_notes_button AND _
			ButtonPressed <> m_through_q_notes_button AND _
			ButtonPressed <> r_through_z_notes_button AND _
			ButtonPressed <> LTC_notes_button

'Available scripts
IF ButtonPressed = APPLICATION_RECEIVED_button						THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - APPLICATION RECEIVED.vbs")
IF ButtonPressed = APPROVED_PROGRAMS_button							THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - APPROVED PROGRAMS.vbs")
IF ButtonPressed = BABY_BORN_button									THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - BABY BORN.vbs")
IF ButtonPressed = BURIAL_ASSETS_button								THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - BURIAL ASSETS.vbs")
IF ButtonPressed = CAF_button										THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - CAF.vbs")
IF ButtonPressed = CHANGE_REPORT_FORM_RECEIVED_button				THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - CHANGE REPORT FORM RECEIVED.vbs")
IF ButtonPressed = CITIZENSHIP_IDENTITY_VERIFIED_button				THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - CITIZENSHIP-IDENTITY VERIFIED.vbs")
IF ButtonPressed = CLIENT_CONTACT_button							THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - CLIENT CONTACT.vbs")
IF ButtonPressed = CLOSED_PROGRAMS_button							THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - CLOSED PROGRAMS.vbs")
IF ButtonPressed = COMBINED_AR_button								THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - COMBINED AR.vbs")
IF ButtonPressed = CSR_button										THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - CSR.vbs")
IF ButtonPressed = COUNTY_BURIAL_button								THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - COUNTY BURIAL APPLICATION.vbs")

IF ButtonPressed = DENIED_PROGRAMS_button							THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - DENIED PROGRAMS.vbs")
IF ButtonPressed = DOCUMENTS_RECEIVED_button						THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - DOCUMENTS RECEIVED.vbs")
IF ButtonPressed = DRUG_FELON_button								THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - DRUG FELON.vbs")
IF ButtonPressed = DWP_BUDGET_button								THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - DWP BUDGET.vbs")
IF ButtonPressed = EMERGENCY_button									THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - EMERGENCY.vbs")
IF ButtonPressed = EMPLOYMENT_PLAN_OR_STATUS_UPDATE_button			THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - EMPLOYMENT PLAN OR STATUS UPDATE.vbs")
IF ButtonPressed = EMPLOYMENT_VERIFICATION_button					THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - EVF RECEIVED.vbs")
IF ButtonPressed = ES_REFERRAL_button								THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - ES REFERRAL.vbs")
IF ButtonPressed = EXPEDITED_SCREENING_button						THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - EXPEDITED SCREENING.vbs")
IF ButtonPressed = FRAUD_INFO_button								THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - FRAUD INFO.vbs")

IF ButtonPressed = GAS_CARDS_ISSUED_button							THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - GAS CARDS ISSUED.vbs")
IF ButtonPressed = GOOD_CAUSE_CLAIMED_button						THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - GOOD CAUSE CLAIMED.vbs")
IF ButtonPressed = GRH_HRF_button									THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - GRH - HRF.vbs")
IF ButtonPressed = HC_RENEWAL_button								THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - HC RENEWAL.vbs")
IF ButtonPressed = HCAPP_button										THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - HCAPP.vbs")
IF ButtonPressed = HH_COMP_CHANGE_button							THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - HH COMP CHANGE.vbs")
IF ButtonPressed = HRF_button										THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - HRF.vbs")
IF ButtonPressed = IEVS_NOTICE_RECEIVED_button						THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - IEVS NOTICE RECEIVED.vbs")
If ButtonPressed = Incarceration_button								THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES – INCARCERATION.vbs")
IF ButtonPressed = INTERVIEW_COMPLETED_button						THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - INTERVIEW COMPLETED.vbs")
IF ButtonPressed = LEP_EMA_button									THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - LEP - EMA.vbs")
IF ButtonPressed = LEP_SAVE_button									THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - LEP - SAVE.vbs")
IF ButtonPressed = LEP_SPONSOR_INCOME_button						THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - LEP - SPONSOR INCOME.vbs")
IF ButtonPressed = LOBBY_NO_SHOW_button 							THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - LOBBY NO SHOW.vbs")

IF ButtonPressed = MEDICAL_OPINION_FORM_RECEIVED_button				THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - MEDICAL OPINION FORM RECEIVED.vbs")
IF ButtonPressed = MFIP_SANCTION_AND_DWP_DISQUALIFICATION_button	THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - MFIP SANCTION AND DWP DISQUALIFICATION.vbs")
IF ButtonPressed = MFIP_SANCTION_CURED_button						THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - MFIP SANCTION CURED.vbs")
IF ButtonPressed = MFIP_TO_SNAP_TRANSITION_button					THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - MFIP TO SNAP TRANSITION.vbs")
IF ButtonPressed = MILEAGE_REIMBURSEMENT_REQUEST_button				THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - MILEAGE REIMBURSEMENT REQUEST.vbs")
IF ButtonPressed = MNSURE_DOCUMENTS_REQUESTED_button				THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - MNSURE - DOCUMENTS REQUESTED.vbs")
IF ButtonPressed = MNSURE_RETRO_HC_APPLICATION_button				THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - MNSURE RETRO HC APPLICATION.vbs")
If ButtonPressed = MTAF_button										THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - MTAF.vbs")
IF ButtonPressed = OVERPAYMENT_button								THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - OVERPAYMENT.vbs")
IF ButtonPressed = PREGNANCY_REPORTED_button						THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - PREGNANCY REPORTED.vbs")
IF ButtonPressed = PROOF_OF_RELATIONSHIP_button						THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - PROOF OF RELATIONSHIP.vbs")
IF ButtonPressed = REIN_PROGS_button								THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - REIN PROGS.vbs")
IF ButtonPressed = SHELTER_FORM_RECEIVED_button						THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - SHELTER FORM RECEIVED.vbs")
IF ButtonPressed = SUBMIT_CASE_FOR_SNAP_REVIEW_button				THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - SUBMIT CASE FOR SNAP REVIEW.vbs")
IF ButtonPressed = SNAP_CASE_REVIEW_button							THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - SNAP CASE REVIEW.vbs")
IF ButtonPressed = VERIFICATIONS_NEEDED_button						THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - VERIFICATIONS NEEDED.vbs")

IF ButtonPressed = LTC_1503_button									THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - LTC - 1503.vbs")
IF ButtonPressed = LTC_5181_button									THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - LTC - 5181.vbs")
IF ButtonPressed = LTC_APPLICATION_RECEIVED_button					THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - LTC - APPLICATION RECEIVED.vbs")
IF ButtonPressed = LTC_ASSET_ASSESSMENT_button						THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - LTC - ASSET ASSESSMENT.vbs")
IF ButtonPressed = LTC_COLA_SUMMARY_2016_button						THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - LTC - COLA SUMMARY 2016.vbs")
IF ButtonPressed = LTC_INTAKE_APPROVAL_button						THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - LTC - INTAKE APPROVAL.vbs")
IF ButtonPressed = LTC_MA_APPROVAL_button							THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - LTC - MA APPROVAL.vbs")
IF ButtonPressed = LTC_RENEWAL_button								THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - LTC - RENEWAL.vbs")
IF ButtonPressed = LTC_TRANSFER_PENALTY_button						THEN CALL run_from_GitHub(script_repository & "/NOTES/NOTES - LTC - TRANSFER PENALTY.vbs")

'Logging usage stats
script_end_procedure("If you see this, it's because you clicked a button that, for some reason, does not have an outcome in the script. Contact your alpha user to report this bug. Thank you!")
