'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - CAF.vbs"
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

'DATE CALCULATIONS----------------------------------------------------------------------------------------------------
footer_month = datepart("m", date)
If len(footer_month) = 1 then footer_month = "0" & footer_month
footer_year = "" & datepart("yyyy", date) - 2000

'DIALOGS-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
BeginDialog case_number_dialog, 0, 0, 181, 120, "Case number dialog"
  EditBox 80, 5, 70, 15, case_number
  EditBox 65, 25, 30, 15, footer_month
  EditBox 140, 25, 30, 15, footer_year
  CheckBox 10, 60, 30, 10, "cash", cash_checkbox
  CheckBox 50, 60, 30, 10, "HC", HC_checkbox
  CheckBox 90, 60, 35, 10, "SNAP", SNAP_checkbox
  CheckBox 135, 60, 35, 10, "EMER", EMER_checkbox
  DropListBox 70, 80, 75, 15, "Intake"+chr(9)+"Reapplication"+chr(9)+"Recertification"+chr(9)+"Add program"+chr(9)+"Addendum", CAF_type
  ButtonGroup ButtonPressed
	OkButton 35, 100, 50, 15
	CancelButton 95, 100, 50, 15
  Text 25, 10, 50, 10, "Case number:"
  Text 10, 30, 50, 10, "Footer month:"
  Text 110, 30, 25, 10, "Year:"
  GroupBox 5, 45, 170, 30, "Programs applied for"
  Text 30, 85, 35, 10, "CAF type:"
EndDialog


BeginDialog CAF_dialog_01, 0, 0, 451, 260, "CAF dialog part 1"
  EditBox 60, 5, 50, 15, CAF_datestamp
  ComboBox 175, 5, 70, 15, " "+chr(9)+"phone"+chr(9)+"office", interview_type
  EditBox 60, 25, 50, 15, interview_date
  ComboBox 230, 25, 95, 15, " "+chr(9)+"in-person"+chr(9)+"dropped off"+chr(9)+"mailed in"+chr(9)+"ApplyMN"+chr(9)+"faxed", how_app_was_received
  ComboBox 220, 45, 105, 15, " "+chr(9)+"DHS-2128 (LTC Renewal)"+chr(9)+"DHS-3417B (Req. to Apply...)"+chr(9)+"DHS-3418 (HC Renewal)"+chr(9)+"DHS-3531 (LTC Application)", HC_document_received
  EditBox 390, 45, 50, 15, HC_datestamp
  EditBox 75, 70, 370, 15, HH_comp
  EditBox 35, 90, 200, 15, cit_id
  EditBox 265, 90, 180, 15, IMIG
  EditBox 60, 110, 120, 15, AREP
  EditBox 270, 110, 175, 15, SCHL
  EditBox 60, 130, 210, 15, DISA
  EditBox 310, 130, 135, 15, FACI
  EditBox 35, 160, 410, 15, PREG
  EditBox 35, 180, 410, 15, ABPS
  EditBox 55, 210, 390, 15, verifs_needed
  ButtonGroup ButtonPressed
	PushButton 340, 240, 50, 15, "NEXT", next_to_page_02_button
	CancelButton 395, 240, 50, 15
	PushButton 335, 15, 45, 10, "prev. panel", prev_panel_button
	PushButton 335, 25, 45, 10, "next panel", next_panel_button
	PushButton 395, 15, 45, 10, "prev. memb", prev_memb_button
	PushButton 395, 25, 45, 10, "next memb", next_memb_button
	PushButton 5, 75, 60, 10, "HH comp/EATS:", EATS_button
	PushButton 240, 95, 20, 10, "IMIG:", IMIG_button
	PushButton 5, 115, 25, 10, "AREP/", AREP_button
	PushButton 30, 115, 25, 10, "ALTP:", ALTP_button
	PushButton 190, 115, 25, 10, "SCHL/", SCHL_button
	PushButton 215, 115, 25, 10, "STIN/", STIN_button
	PushButton 240, 115, 25, 10, "STEC:", STEC_button
	PushButton 5, 135, 25, 10, "DISA/", DISA_button
	PushButton 30, 135, 25, 10, "PDED:", PDED_button
	PushButton 280, 135, 25, 10, "FACI:", FACI_button
	PushButton 5, 165, 25, 10, "PREG:", PREG_button
	PushButton 5, 185, 25, 10, "ABPS:", ABPS_button
	PushButton 10, 240, 20, 10, "DWP", ELIG_DWP_button
	PushButton 30, 240, 15, 10, "FS", ELIG_FS_button
	PushButton 45, 240, 15, 10, "GA", ELIG_GA_button
	PushButton 60, 240, 15, 10, "HC", ELIG_HC_button
	PushButton 75, 240, 20, 10, "MFIP", ELIG_MFIP_button
	PushButton 95, 240, 20, 10, "MSA", ELIG_MSA_button
	PushButton 115, 240, 15, 10, "WB", ELIG_WB_button
	PushButton 150, 240, 25, 10, "ADDR", ADDR_button
	PushButton 175, 240, 25, 10, "MEMB", MEMB_button
	PushButton 200, 240, 25, 10, "MEMI", MEMI_button
	PushButton 225, 240, 25, 10, "PROG", PROG_button
	PushButton 250, 240, 25, 10, "REVW", REVW_button
	PushButton 275, 240, 25, 10, "TYPE", TYPE_button
  Text 5, 10, 55, 10, "CAF datestamp:"
  Text 120, 10, 50, 10, "Interview type:"
  GroupBox 330, 5, 115, 35, "STAT-based navigation"
  Text 5, 30, 55, 10, "Interview date:"
  Text 120, 30, 110, 10, "How was application received?:"
  Text 5, 50, 210, 10, "If HC applied for (or recertifying): what document was received?:"
  Text 335, 50, 55, 10, "HC datestamp:"
  Text 5, 95, 25, 10, "CIT/ID:"
  Text 5, 215, 50, 10, "Verifs needed:"
  GroupBox 5, 230, 130, 25, "ELIG panels:"
  GroupBox 145, 230, 160, 25, "other STAT panels:"
EndDialog


BeginDialog CAF_dialog_02, 0, 0, 451, 315, "CAF dialog part 2"
  EditBox 60, 45, 385, 15, earned_income
  EditBox 70, 65, 375, 15, unearned_income
  EditBox 85, 85, 360, 15, income_changes
  EditBox 65, 105, 380, 15, notes_on_abawd
  EditBox 105, 125, 340, 15, notes_on_income
  EditBox 155, 145, 290, 15, is_any_work_temporary
  EditBox 60, 175, 385, 15, SHEL_HEST
  EditBox 60, 195, 250, 15, COEX_DCEX
  EditBox 65, 225, 380, 15, CASH_ACCTs
  EditBox 155, 245, 290, 15, other_assets
  EditBox 55, 275, 390, 15, verifs_needed
  ButtonGroup ButtonPressed
	PushButton 340, 295, 50, 15, "NEXT", next_to_page_03_button
	CancelButton 395, 295, 50, 15
	PushButton 275, 300, 60, 10, "previous page", previous_to_page_01_button
	PushButton 10, 15, 20, 10, "DWP", ELIG_DWP_button
	PushButton 30, 15, 15, 10, "FS", ELIG_FS_button
	PushButton 45, 15, 15, 10, "GA", ELIG_GA_button
	PushButton 60, 15, 15, 10, "HC", ELIG_HC_button
	PushButton 75, 15, 20, 10, "MFIP", ELIG_MFIP_button
	PushButton 95, 15, 20, 10, "MSA", ELIG_MSA_button
	PushButton 115, 15, 15, 10, "WB", ELIG_WB_button
	PushButton 150, 15, 25, 10, "BUSI", BUSI_button
	PushButton 175, 15, 25, 10, "JOBS", JOBS_button
	PushButton 200, 15, 25, 10, "PBEN", PBEN_button
	PushButton 225, 15, 25, 10, "RBIC", RBIC_button
	PushButton 250, 15, 25, 10, "UNEA", UNEA_button
	PushButton 335, 15, 45, 10, "prev. panel", prev_panel_button
	PushButton 335, 25, 45, 10, "next panel", next_panel_button
	PushButton 395, 15, 45, 10, "prev. memb", prev_memb_button
	PushButton 395, 25, 45, 10, "next memb", next_memb_button
	PushButton 5, 90, 75, 10, "STWK/inc. changes:", STWK_button
	PushButton 5, 180, 25, 10, "SHEL/", SHEL_button
	PushButton 30, 180, 25, 10, "HEST:", HEST_button
	PushButton 5, 200, 25, 10, "COEX/", COEX_button
	PushButton 30, 200, 25, 10, "DCEX:", DCEX_button
	PushButton 5, 230, 25, 10, "CASH/", CASH_button
	PushButton 30, 230, 30, 10, "ACCTs:", ACCT_button
	PushButton 5, 250, 25, 10, "CARS/", CARS_button
	PushButton 30, 250, 25, 10, "REST/", REST_button
	PushButton 55, 250, 25, 10, "SECU/", SECU_button
	PushButton 80, 250, 25, 10, "TRAN/", TRAN_button
	PushButton 105, 250, 45, 10, "other assets:", OTHR_button
  GroupBox 5, 5, 130, 25, "ELIG panels:"
  GroupBox 145, 5, 135, 25, "Income panels"
  GroupBox 330, 5, 115, 35, "STAT-based navigation"
  Text 5, 50, 55, 10, "Earned income:"
  Text 5, 70, 65, 10, "Unearned income:"
  Text 5, 110, 50, 10, "ABAWD notes:"
  Text 5, 130, 100, 10, "Notes on income and budget:"
  Text 5, 150, 150, 10, "Is any work temporary? If so, explain details:"
  Text 5, 280, 50, 10, "Verifs needed:"
EndDialog

'CAF_status needs to have the " "+chr(9)+ manually added each time.
BeginDialog CAF_dialog_03, 0, 0, 451, 365, "CAF dialog part 3"
  EditBox 60, 45, 385, 15, INSA
  EditBox 35, 65, 410, 15, ACCI
  EditBox 35, 85, 175, 15, DIET
  EditBox 245, 85, 200, 15, BILS
  EditBox 35, 105, 285, 15, FMED
  EditBox 390, 105, 55, 15, retro_request
  EditBox 180, 130, 265, 15, reason_expedited_wasnt_processed
  EditBox 100, 150, 345, 15, FIAT_reasons
  CheckBox 15, 190, 80, 10, "Application signed?", application_signed_checkbox
  CheckBox 100, 190, 65, 10, "Appt letter sent?", appt_letter_sent_checkbox
  CheckBox 175, 190, 70, 10, "EBT referral sent?", EBT_referral_checkbox
  CheckBox 255, 190, 50, 10, "eDRS sent?", eDRS_sent_checkbox
  CheckBox 315, 190, 50, 10, "Expedited?", expedited_checkbox
  CheckBox 375, 190, 70, 10, "IAAs/OMB given?", IAA_checkbox
  CheckBox 15, 205, 80, 10, "Intake packet given?", intake_packet_checkbox
  CheckBox 100, 205, 105, 10, "Managed care packet sent?", managed_care_packet_checkbox
  CheckBox 210, 205, 110, 10, "Managed care referral made?", managed_care_referral_checkbox
  CheckBox 330, 205, 65, 10, "R/R explained?", R_R_checkbox
  CheckBox 15, 220, 65, 10, "Updated MMIS?", updated_MMIS_checkbox
  CheckBox 90, 220, 95, 10, "Workforce referral made?", WF1_checkbox
  CheckBox 190, 220, 85, 10, "Sent forms to AREP?", Sent_arep_checkbox
  EditBox 55, 240, 230, 15, other_notes
  ComboBox 330, 240, 115, 15, " "+chr(9)+"incomplete"+chr(9)+"approved", CAF_status
  EditBox 55, 260, 390, 15, verifs_needed
  EditBox 55, 280, 390, 15, actions_taken
  CheckBox 15, 315, 240, 10, "Check here to update PND2 to show client delay (pending cases only).", client_delay_checkbox
  CheckBox 15, 330, 200, 10, "Check here to create a TIKL to deny at the 30/45 day mark.", TIKL_checkbox
  CheckBox 15, 345, 265, 10, "Check here to send a TIKL (10 days from now) to update PND2 for Client Delay.", client_delay_TIKL_checkbox
  EditBox 395, 325, 50, 15, worker_signature
  ButtonGroup ButtonPressed
	OkButton 340, 345, 50, 15
	CancelButton 395, 345, 50, 15
	PushButton 10, 15, 20, 10, "DWP", ELIG_DWP_button
	PushButton 30, 15, 15, 10, "FS", ELIG_FS_button
	PushButton 45, 15, 15, 10, "GA", ELIG_GA_button
	PushButton 60, 15, 15, 10, "HC", ELIG_HC_button
	PushButton 75, 15, 20, 10, "MFIP", ELIG_MFIP_button
	PushButton 95, 15, 20, 10, "MSA", ELIG_MSA_button
	PushButton 115, 15, 15, 10, "WB", ELIG_WB_button
	PushButton 335, 15, 45, 10, "prev. panel", prev_panel_button
	PushButton 335, 25, 45, 10, "next panel", next_panel_button
	PushButton 395, 15, 45, 10, "prev. memb", prev_memb_button
	PushButton 395, 25, 45, 10, "next memb", next_memb_button
	PushButton 5, 50, 25, 10, "INSA/", INSA_button
	PushButton 30, 50, 25, 10, "MEDI:", MEDI_button
	PushButton 5, 70, 25, 10, "ACCI:", ACCI_button
	PushButton 5, 90, 25, 10, "DIET:", DIET_button
	PushButton 215, 90, 25, 10, "BILS:", BILS_button
	PushButton 5, 110, 25, 10, "FMED:", FMED_button
	PushButton 325, 110, 60, 10, "Retro Req. date:", HCRE_button
	PushButton 290, 350, 45, 10, "prev. page", previous_to_page_02_button
  GroupBox 5, 5, 130, 25, "ELIG panels:"
  GroupBox 330, 5, 115, 35, "STAT-based navigation"
  Text 5, 135, 170, 10, "Reason expedited wasn't processed (if applicable):"
  Text 5, 155, 95, 10, "FIAT reasons (if applicable):"
  GroupBox 5, 175, 440, 60, "Common elements workers should case note:"
  Text 5, 245, 50, 10, "Other notes:"
  Text 290, 245, 40, 10, "CAF status:"
  Text 5, 265, 50, 10, "Verifs needed:"
  Text 5, 285, 50, 10, "Actions taken:"
  GroupBox 5, 300, 280, 60, "Actions the script can do:"
  Text 330, 330, 60, 10, "Worker signature:"
EndDialog




'VARIABLES WHICH NEED DECLARING------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
HH_memb_row = 5 'This helps the navigation buttons work!
Dim row
Dim col
application_signed_checkbox = checked 'The script should default to having the application signed.


'GRABBING THE CASE NUMBER, THE MEMB NUMBERS, AND THE FOOTER MONTH------------------------------------------------------------------------------------------------------------------------------------------------
EMConnect ""

call find_variable("Case Nbr: ", case_number, 8)
case_number = trim(case_number)
case_number = replace(case_number, "_", "")
If IsNumeric(case_number) = False then case_number = ""

call find_variable("Month: ", MAXIS_footer_month, 2)
If row <> 0 then 
  footer_month = MAXIS_footer_month
  call find_variable("Month: " & footer_month & " ", MAXIS_footer_year, 2)
  If row <> 0 then footer_year = MAXIS_footer_year
End if

case_number = trim(case_number)
case_number = replace(case_number, "_", "")
If IsNumeric(case_number) = False then case_number = ""

Do
  Dialog case_number_dialog
  If ButtonPressed = 0 then stopscript
  If case_number = "" or IsNumeric(case_number) = False or len(case_number) > 8 then MsgBox "You need to type a valid case number."
Loop until case_number <> "" and IsNumeric(case_number) = True and len(case_number) <= 8
transmit
call check_for_MAXIS(True)


'GRABBING THE DATE RECEIVED AND THE HH MEMBERS---------------------------------------------------------------------------------------------------------------------------------------------------------------------
call navigate_to_screen("stat", "hcre")
EMReadScreen STAT_check, 4, 20, 21
If STAT_check <> "STAT" then script_end_procedure("Can't get in to STAT. This case may be in background. Wait a few seconds and try again. If the case is not in background contact an alpha user for your agency.")


'Creating a custom dialog for determining who the HH members are
call HH_member_custom_dialog(HH_member_array)

'GRABBING THE INFO FOR THE CASE NOTE-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

If CAF_type = "Recertification" then                                                          'For recerts it goes to one area for the CAF datestamp. For other app types it goes to STAT/PROG.
	call autofill_editbox_from_MAXIS(HH_member_array, "REVW", CAF_datestamp)
Else
	call autofill_editbox_from_MAXIS(HH_member_array, "PROG", CAF_datestamp)
End if
If HC_checkbox = checked and CAF_type <> "Recertification" then call autofill_editbox_from_MAXIS(HH_member_array, "HCRE-retro", retro_request)     'Grabbing retro info for HC cases that aren't recertifying
call autofill_editbox_from_MAXIS(HH_member_array, "MEMB", HH_comp)                                                                        'Grabbing HH comp info from MEMB.
If SNAP_checkbox = checked then call autofill_editbox_from_MAXIS(HH_member_array, "EATS", HH_comp)                                                 'Grabbing EATS info for SNAP cases, puts on HH_comp variable
'Removing semicolons from HH_comp variable, it is not needed.
HH_comp = replace(HH_comp, "; ", "")


'I put these sections in here, just because SHEL should come before HEST, it just looks cleaner.
call autofill_editbox_from_MAXIS(HH_member_array, "SHEL", SHEL_HEST) 
call autofill_editbox_from_MAXIS(HH_member_array, "HEST", SHEL_HEST) 

'Now it grabs the rest of the info, not dependent on which programs are selected.
call autofill_editbox_from_MAXIS(HH_member_array, "ABPS", ABPS)
call autofill_editbox_from_MAXIS(HH_member_array, "ACCI", ACCI)
call autofill_editbox_from_MAXIS(HH_member_array, "ACCT", CASH_ACCTs)
call autofill_editbox_from_MAXIS(HH_member_array, "AREP", AREP)
call autofill_editbox_from_MAXIS(HH_member_array, "BILS", BILS)
call autofill_editbox_from_MAXIS(HH_member_array, "BUSI", earned_income)
call autofill_editbox_from_MAXIS(HH_member_array, "CASH", CASH_ACCTs)
call autofill_editbox_from_MAXIS(HH_member_array, "CARS", other_assets)
call autofill_editbox_from_MAXIS(HH_member_array, "COEX", COEX_DCEX)
call autofill_editbox_from_MAXIS(HH_member_array, "DCEX", COEX_DCEX)
call autofill_editbox_from_MAXIS(HH_member_array, "DIET", DIET)
call autofill_editbox_from_MAXIS(HH_member_array, "DISA", DISA)
call autofill_editbox_from_MAXIS(HH_member_array, "FACI", FACI)
call autofill_editbox_from_MAXIS(HH_member_array, "FMED", FMED)
call autofill_editbox_from_MAXIS(HH_member_array, "IMIG", IMIG)
call autofill_editbox_from_MAXIS(HH_member_array, "INSA", INSA)
call autofill_editbox_from_MAXIS(HH_member_array, "JOBS", earned_income)
call autofill_editbox_from_MAXIS(HH_member_array, "MEDI", INSA)
call autofill_editbox_from_MAXIS(HH_member_array, "MEMI", cit_id)
call autofill_editbox_from_MAXIS(HH_member_array, "OTHR", other_assets)
call autofill_editbox_from_MAXIS(HH_member_array, "PBEN", income_changes)
call autofill_editbox_from_MAXIS(HH_member_array, "PREG", PREG)
call autofill_editbox_from_MAXIS(HH_member_array, "RBIC", earned_income)
call autofill_editbox_from_MAXIS(HH_member_array, "REST", other_assets)
call autofill_editbox_from_MAXIS(HH_member_array, "SCHL", SCHL)
call autofill_editbox_from_MAXIS(HH_member_array, "SECU", other_assets)
call autofill_editbox_from_MAXIS(HH_member_array, "STWK", income_changes)
call autofill_editbox_from_MAXIS(HH_member_array, "UNEA", unearned_income)
call autofill_editbox_from_MAXIS(HH_member_array, "WREG", notes_on_abawd)

'MAKING THE GATHERED INFORMATION LOOK BETTER FOR THE CASE NOTE
If cash_checkbox = checked then programs_applied_for = programs_applied_for & "cash, "
If HC_checkbox = checked then programs_applied_for = programs_applied_for & "HC, "
If SNAP_checkbox = checked then programs_applied_for = programs_applied_for & "SNAP, "
If EMER_checkbox = checked then programs_applied_for = programs_applied_for & "emergency, "
programs_applied_for = trim(programs_applied_for)
if right(programs_applied_for, 1) = "," then programs_applied_for = left(programs_applied_for, len(programs_applied_for) - 1)


'SHOULD DEFAULT TO TIKLING FOR APPLICATIONS THAT AREN'T RECERTS.
If CAF_type <> "Recertification" then TIKL_checkbox = checked


'CASE NOTE DIALOG--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Do
	Do
		Do
			Dialog CAF_dialog_01			'Displays the first dialog
			cancel_confirmation				'Asks if you're sure you want to cancel, and cancels if you select that.	
			MAXIS_dialog_navigation			'Navigates around MAXIS using a custom function (works with the prev/next buttons and all the navigation buttons)
		Loop until ButtonPressed = next_to_page_02_button
		Do
			Do
				Dialog CAF_dialog_02			'Displays the second dialog
				cancel_confirmation				'Asks if you're sure you want to cancel, and cancels if you select that.
				MAXIS_dialog_navigation			'Navigates around MAXIS using a custom function (works with the prev/next buttons and all the navigation buttons)
			Loop until ButtonPressed = next_to_page_03_button or ButtonPressed = previous_to_page_01_button		'If you press either the next or previous button, this loop ends
			If ButtonPressed = previous_to_page_01_button then exit do		'If the button was previous, it exits this do loop and is caught in the next one, which sends you back to Dialog 1 because of the "If ButtonPressed = previous_to_page_01_button then exit do" later on
			Do
				Dialog CAF_dialog_03			'Displays the third dialog
				cancel_confirmation				'Asks if you're sure you want to cancel, and cancels if you select that.
				MAXIS_dialog_navigation			'Navigates around MAXIS using a custom function (works with the prev/next buttons and all the navigation buttons)
				If ButtonPressed = previous_to_page_02_button then exit do		'Exits this do...loop here if you press previous. The second ""loop until ButtonPressed = -1" gets caught, and it loops back to the "Do" after "Loop until ButtonPressed = next_to_page_02_button"
			Loop until ButtonPressed = -1 or ButtonPressed = previous_to_page_02_button		'If OK or PREV, it exits the loop here, which is weird because the above also causes it to exit
		Loop until ButtonPressed = -1	'Because this is in here a second time, it triggers a return to the "Dialog CAF_dialog_02" line, where all those "DOs" start again!!!!!
		If ButtonPressed = previous_to_page_01_button then exit do 	'This exits this particular loop again for prev button on page 2, which sends you back to page 1!!
		If actions_taken = "" or CAF_datestamp = "" or worker_signature = "" or CAF_status = "" THEN 'Tells the worker what's required in a MsgBox.
			MsgBox "You need to:" & chr(13) & chr(13) & _
			  "-Fill in the datestamp, and/or" & chr(13) & _
			  "-Actions taken sections, and/or" & chr(13) & _
			  "-HCAPP Status, and/or" & chr(13) & _
			  "-Sign your case note." & chr(13) & chr(13) & _
			  "Check these items after pressing ''OK''."	
		End if
	Loop until actions_taken <> "" and CAF_datestamp <> "" and worker_signature <> "" and CAF_status <> ""		'Loops all of that until those four sections are finished. Let's move that over to those particular pages. Folks would be less angry that way I bet.
	CALL proceed_confirmation(case_note_confirm)			'Checks to make sure that we're ready to case note.
Loop until case_note_confirm = TRUE							'Loops until we affirm that we're ready to case note.

check_for_maxis(FALSE)  'allows for looping to check for maxis after worker has complete dialog box so as not to lose a giant CAF case note if they get timed out while writing. 

'Now, the client_delay_checkbox business. It'll update client delay if the box is checked and it isn't a recert.
If client_delay_checkbox = checked and CAF_type <> "Recertification" then 
	call navigate_to_screen("rept", "pnd2")
	EMGetCursor PND2_row, PND2_col
	for i = 0 to 1 'This is put in a for...next statement so that it will check for "additional app" situations, where the case could be on multiple lines in REPT/PND2. It exits after one if it can't find an additional app.
		EMReadScreen PND2_SNAP_status_check, 1, PND2_row, 62
		If PND2_SNAP_status_check = "P" then EMWriteScreen "C", PND2_row, 62
		EMReadScreen PND2_HC_status_check, 1, PND2_row, 65
		If PND2_HC_status_check = "P" then
			EMWriteScreen "x", PND2_row, 3
			transmit
			person_delay_row = 7
			Do
				EMReadScreen person_delay_check, 1, person_delay_row, 39
				If person_delay_check <> " " then EMWriteScreen "c", person_delay_row, 39
				person_delay_row = person_delay_row + 2
			Loop until person_delay_check = " " or person_delay_row > 20
			PF3
		End if
		EMReadScreen additional_app_check, 14, PND2_row + 1, 17
		If additional_app_check <> "ADDITIONAL APP" then exit for
		PND2_row = PND2_row + 1
	next
	PF3
	EMReadScreen PND2_check, 4, 2, 52
	If PND2_check = "PND2" then
		MsgBox "PND2 might not have been updated for client delay. There may have been a MAXIS error. Check this manually after case noting."
		PF10
		client_delay_checkbox = unchecked		'Probably unnecessary except that it changes the case note parameters
	End if
End if

'Going to TIKL, there's a custom function for this. Evaluate using it.
If TIKL_checkbox = checked and CAF_type <> "Recertification" then
	If cash_checkbox = checked or EMER_checkbox = checked or SNAP_checkbox = checked then
		call navigate_to_screen("dail", "writ")
		call create_MAXIS_friendly_date(CAF_datestamp, 30, 5, 18) 
		EMSetCursor 9, 3
		If cash_checkbox = checked then EMSendKey "cash/"
		If SNAP_checkbox = checked then EMSendKey "SNAP/"
		If EMER_checkbox = checked then EMSendKey "EMER/"
		EMSendKey "<backspace>" & " pending 30 days. Evaluate for possible denial."
		transmit
		PF3
	End if
	If HC_checkbox = checked then
		call navigate_to_screen("dail", "writ")
		call create_MAXIS_friendly_date(CAF_datestamp, 45, 5, 18) 
		EMSetCursor 9, 3
		EMSendKey "HC pending 45 days. Evaluate for possible denial. If any members are elderly/disabled, allow an additional 15 days and reTIKL out."
		transmit
		PF3
	End if
End if
If client_delay_TIKL_checkbox = checked then
	call navigate_to_screen("dail", "writ")
	call create_MAXIS_friendly_date(date, 10, 5, 18) 
	EMSetCursor 9, 3
	EMSendKey ">>>UPDATE PND2 FOR CLIENT DELAY IF APPROPRIATE<<<"
	transmit
	PF3
End if
'----Here's the new bit to TIKL to APPL the CAF for CAF_datestamp if the CL fails to complete the CASH/SNAP reinstate and then TIKL again for DateAdd("D", 30, CAF_datestamp) to evaluate for possible denial.
'----IF the DatePart("M", CAF_datestamp) = footer_month (DatePart("M", CAF_datestamp) is converted to footer_comparo_month for the sake of comparison) and the CAF_status <> "Approved" and CAF_type is a recertification AND cash or snap is checked, then 
'---------the script generates a TIKL.
footer_comparison_month = DatePart("M", CAF_datestamp)
IF len(footer_comparison_month) <> 2 THEN footer_comparison_month = "0" & footer_comparison_month
IF CAF_type = "Recertification" AND footer_month = footer_comparison_month AND CAF_status <> "approved" AND (cash_checkbox = checked OR SNAP_checkbox = checked) THEN 
	CALL navigate_to_MAXIS_screen("DAIL", "WRIT")
	start_of_next_month = DatePart("M", DateAdd("M", 1, CAF_datestamp)) & "/01/" & DatePart("YYYY", DateAdd("M", 1, CAF_datestamp))
	denial_consider_date = DateAdd("D", 30, CAF_datestamp)
	CALL create_MAXIS_friendly_date(start_of_next_month, 0, 5, 18)
	EMWriteScreen ("IF CLIENT HAS NOT COMPLETED RECERT, APPL CAF FOR " & CAF_datestamp), 9, 3
	EMWriteScreen ("AND TIKL FOR " & denial_consider_date & " TO EVALUATE FOR POSSIBLE DENIAL."), 10, 3
	transmit
	PF3
END IF	
'--------------------END OF TIKL BUSINESS

'Navigates to case note, and checks to make sure we aren't in inquiry.
start_a_blank_CASE_NOTE


'Adding a colon to the beginning of the CAF status variable if it isn't blank (simplifies writing the header of the case note)
If CAF_status <> "" then CAF_status = ": " & CAF_status

'Adding footer month to the recertification case notes
If CAF_type = "Recertification" then CAF_type = footer_month & "/" & footer_year & " recert"

'THE CASE NOTE-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
CALL write_variable_in_CASE_NOTE("***" & CAF_type & CAF_status & "***")
IF move_verifs_needed = TRUE THEN CALL write_bullet_and_variable_in_CASE_NOTE("Verifs needed", verifs_needed)			'IF global variable move_verifs_needed = True (on FUNCTIONS FILE), it'll case note at the top.
CALL write_bullet_and_variable_in_CASE_NOTE("CAF datestamp", CAF_datestamp)
CALL write_bullet_and_variable_in_CASE_NOTE("Interview type", interview_type)											
CALL write_bullet_and_variable_in_CASE_NOTE("Interview date", interview_date)
CALL write_bullet_and_variable_in_CASE_NOTE("HC document received", HC_document_received)								
CALL write_bullet_and_variable_in_CASE_NOTE("HC datestamp", HC_datestamp)
CALL write_bullet_and_variable_in_CASE_NOTE("Programs applied for", programs_applied_for)
CALL write_bullet_and_variable_in_CASE_NOTE("How CAF was received", how_app_was_received)								
CALL write_bullet_and_variable_in_CASE_NOTE("HH comp/EATS", HH_comp)
CALL write_bullet_and_variable_in_CASE_NOTE("Cit/ID", cit_id)
CALL write_bullet_and_variable_in_CASE_NOTE("IMIG", IMIG)
CALL write_bullet_and_variable_in_CASE_NOTE("AREP", AREP)
CALL write_bullet_and_variable_in_CASE_NOTE("FACI", FACI)
CALL write_bullet_and_variable_in_CASE_NOTE("SCHL/STIN/STEC", SCHL)
CALL write_bullet_and_variable_in_CASE_NOTE("DISA", DISA)
CALL write_bullet_and_variable_in_CASE_NOTE("PREG", PREG)
CALL write_bullet_and_variable_in_CASE_NOTE("ABPS", ABPS)
CALL write_bullet_and_variable_in_CASE_NOTE("Earned inc.", earned_income)
CALL write_bullet_and_variable_in_CASE_NOTE("UNEA", unearned_income)
CALL write_bullet_and_variable_in_CASE_NOTE("STWK/inc. changes", income_changes)
CALL write_bullet_and_variable_in_CASE_NOTE("ABAWD Notes", notes_on_abawd)
CALL write_bullet_and_variable_in_CASE_NOTE("Notes on income and budget", notes_on_income)
CALL write_bullet_and_variable_in_CASE_NOTE("Is any work temporary", is_any_work_temporary)
CALL write_bullet_and_variable_in_CASE_NOTE("SHEL/HEST", SHEL_HEST)
CALL write_bullet_and_variable_in_CASE_NOTE("COEX/DCEX", COEX_DCEX)
CALL write_bullet_and_variable_in_CASE_NOTE("CASH/ACCTs", CASH_ACCTs)
CALL write_bullet_and_variable_in_CASE_NOTE("Other assets", other_assets)
CALL write_bullet_and_variable_in_CASE_NOTE("INSA", INSA)
CALL write_bullet_and_variable_in_CASE_NOTE("ACCI", ACCI)
CALL write_bullet_and_variable_in_CASE_NOTE("DIET", DIET)
CALL write_bullet_and_variable_in_CASE_NOTE("BILS", BILS)
CALL write_bullet_and_variable_in_CASE_NOTE("FMED", FMED)
CALL write_bullet_and_variable_in_CASE_NOTE("Retro Request (IF applicable)", retro_request)
IF application_signed_checkbox = checked THEN 
	CALL write_variable_in_CASE_NOTE("* Application was signed.")
Else
	CALL write_variable_in_CASE_NOTE("* Application was not signed.")
END IF
IF appt_letter_sent_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* Appointment letter was sent before interview.")
IF EBT_referral_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* EBT referral made for client.")
IF eDRS_sent_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* eDRS sent.")
IF expedited_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* Expedited SNAP.")
CALL write_bullet_and_variable_in_CASE_NOTE("Reason expedited wasn't processed", reason_expedited_wasnt_processed)		'This is strategically placed next to expedited checkbox entry.
IF IAA_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* IAAs/OMB given to client.")
IF intake_packet_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* Client received intake packet.")
IF managed_care_packet_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* Client received managed care packet.")
IF managed_care_referral_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* Managed care referral made.")
IF R_R_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* R/R explained to client.")
IF updated_MMIS_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* Updated MMIS.")
IF WF1_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* Workforce referral made.")
IF Sent_arep_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* Sent form(s) to AREP.")
IF client_delay_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* PND2 updated to show client delay.")
CALL write_bullet_and_variable_in_CASE_NOTE("FIAT reasons", FIAT_reasons)
CALL write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
IF move_verifs_needed = False THEN CALL write_bullet_and_variable_in_CASE_NOTE("Verifs needed", verifs_needed)			'IF global variable move_verifs_needed = False (on FUNCTIONS FILE), it'll case note at the bottom.
CALL write_bullet_and_variable_in_CASE_NOTE("Actions taken", actions_taken)
CALL write_variable_in_CASE_NOTE("---")
CALL write_variable_in_CASE_NOTE(worker_signature)

script_end_procedure("Success! CAF has been successfully noted. Please remember to run the Approved Programs, Closed Programs, or Denied Programs scripts if  results have been APP'd.")
