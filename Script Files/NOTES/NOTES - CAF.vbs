'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - CAF.vbs"
start_time = timer

'LOADING ROUTINE FUNCTIONS FROM GITHUB REPOSITORY---------------------------------------------------------------------------
url = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a URL
req.open "GET", url, FALSE									'Attempts to open the URL
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
			"URL: " & url
			script_end_procedure("Script ended due to error connecting to GitHub.")
END IF

'DATE CALCULATIONS----------------------------------------------------------------------------------------------------
next_month = dateadd("m", + 1, date)
footer_month = datepart("m", next_month)
If len(footer_month) = 1 then footer_month = "0" & footer_month
footer_year = datepart("yyyy", next_month)
footer_year = "" & footer_year - 2000

'DIALOGS-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
BeginDialog case_number_dialog, 0, 0, 181, 120, "Case number dialog"
  EditBox 80, 5, 70, 15, case_number
  EditBox 65, 25, 30, 15, footer_month
  EditBox 140, 25, 30, 15, footer_year
  CheckBox 10, 60, 30, 10, "cash", cash_checkbox
  CheckBox 50, 60, 30, 10, "HC", HC_checkbox
  CheckBox 90, 60, 35, 10, "SNAP", SNAP_checkbox
  CheckBox 135, 60, 35, 10, "EMER", EMER_checkbox
  DropListBox 70, 80, 75, 15, "Intake"+chr(9)+"Reapplication"+chr(9)+"Recertification"+chr(9)+"Add program", CAF_type
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
      Do
        Do
          Dialog CAF_dialog_01
          If ButtonPressed = 0 then 
            cancel_confirm = MsgBox("Are you sure you want to cancel the script? Press YES to cancel. Press NO to return to the script.", vbYesNo)
            If cancel_confirm = vbYes then stopscript
          End if
        Loop until ButtonPressed <> no_cancel_button
        EMReadScreen STAT_check, 4, 20, 21
        If STAT_check = "STAT" then call stat_navigation
        transmit 'Forces a screen refresh, to keep MAXIS from erroring out in the event of a password prompt.
        EMReadScreen MAXIS_check, 5, 1, 39
        If MAXIS_check <> "MAXIS" and MAXIS_check <> "AXIS " then MsgBox "You do not appear to be in MAXIS. Are you passworded out? Or in MMIS? Check these and try again."
      Loop until MAXIS_check = "MAXIS" or MAXIS_check = "AXIS " 
      If ButtonPressed <> next_to_page_02_button then call navigation_buttons
    Loop until ButtonPressed = next_to_page_02_button
    Do
      Do
        Do
          Do
            Dialog CAF_dialog_02
	            If ButtonPressed = 0 then 
      	      cancel_confirm = MsgBox("Are you sure you want to cancel the script? Press YES to cancel. Press NO to return to the script.", vbYesNo)
            If cancel_confirm = vbYes then stopscript
            End if
          Loop until ButtonPressed <> no_cancel_button
          EMReadScreen STAT_check, 4, 20, 21
          If STAT_check = "STAT" then call stat_navigation
          transmit 'Forces a screen refresh, to keep MAXIS from erroring out in the event of a password prompt.
          EMReadScreen MAXIS_check, 5, 1, 39
          If MAXIS_check <> "MAXIS" and MAXIS_check <> "AXIS " then MsgBox "You do not appear to be in MAXIS. Are you passworded out? Or in MMIS? Check these and try again."
        Loop until MAXIS_check = "MAXIS" or MAXIS_check = "AXIS " 
        If ButtonPressed <> next_to_page_03_button then call navigation_buttons
      Loop until ButtonPressed = next_to_page_03_button or ButtonPressed = previous_to_page_01_button
      If ButtonPressed = previous_to_page_01_button then exit do
      Do
        Do
          Do
            Dialog CAF_dialog_03
            If ButtonPressed = 0 then 
	            cancel_confirm = MsgBox("Are you sure you want to cancel the script? Press YES to cancel. Press NO to return to the script.", vbYesNo)
      	      If cancel_confirm = vbYes then stopscript
            End if
          Loop until ButtonPressed <> no_cancel_button
          EMReadScreen STAT_check, 4, 20, 21
          If STAT_check = "STAT" then call stat_navigation
          transmit 'Forces a screen refresh, to keep MAXIS from erroring out in the event of a password prompt.
          EMReadScreen MAXIS_check, 5, 1, 39
          If MAXIS_check <> "MAXIS" and MAXIS_check <> "AXIS " then MsgBox "You do not appear to be in MAXIS. Are you passworded out? Or in MMIS? Check these and try again."
        Loop until MAXIS_check = "MAXIS" or MAXIS_check = "AXIS " 
        If ButtonPressed <> -1 then call navigation_buttons
        If ButtonPressed = previous_to_page_02_button then exit do
      Loop until ButtonPressed = -1 or ButtonPressed = previous_to_page_02_button
    Loop until ButtonPressed = -1
    If ButtonPressed = previous_to_page_01_button then exit do 'In case the script skipped the third page as a result of hitting "previous page" on part 2
    If actions_taken = "" or CAF_datestamp = "" or worker_signature = "" or CAF_status = "" THEN MsgBox "You need to:" & chr(13) & chr(13) & "-Fill in the datestamp, and/or" & chr(13) & "-Actions taken sections, and/or" & chr(13) & "-HCAPP Status, and/or" & chr(13) & "-Sign your case note." & chr(13) & chr(13) & "Check these items after pressing ''OK''."
  Loop until actions_taken <> "" and CAF_datestamp <> "" and worker_signature <> "" and CAF_status <> ""
  If ButtonPressed = -1 then case_note_confirm = MsgBox("Do you want to case note? Press YES to confirm. Press NO to return to the script.", vbYesNo)
  If case_note_confirm = vbYes then
    If client_delay_checkbox = checked and CAF_type <> "Recertification" then 'UPDATES PND2 FOR CLIENT DELAY IF CHECKED
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
        client_delay_checkbox = unchecked
      End if
    End if
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
    call navigate_to_screen("case", "note")
    PF9
    EMReadScreen case_note_check, 17, 2, 33
    EMReadScreen mode_check, 1, 20, 09
    If case_note_check <> "Case Notes (NOTE)" or mode_check <> "A" then MsgBox "The script can't open a case note. Are you in inquiry? Check MAXIS and try again."
  End if
Loop until case_note_check = "Case Notes (NOTE)" and mode_check = "A"


'Adding a colon to the beginning of the CAF status variable if it isn't blank (simplifies writing the header of the case note)
If CAF_status <> "" then CAF_status = ": " & CAF_status

'Adding footer month to the recertification case notes
If CAF_type = "Recertification" then CAF_type = footer_month & "/" & footer_year & " recert"


'THE CASE NOTE-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

EMSendKey "<home>" & "***" & CAF_type & CAF_status & "***" & "<newline>"
If move_verifs_needed = True and verifs_needed <> "" then call write_bullet_and_variable_in_case_note("Verifs needed", verifs_needed)		'If global variable move_verifs_needed = True (on FUNCTIONS FILE), it'll case note at the top.
call write_bullet_and_variable_in_case_note("CAF datestamp", CAF_datestamp)
If interview_type <> "" and interview_type <> " " then call write_bullet_and_variable_in_case_note("Interview type", interview_type)
If interview_date <> "" then call write_bullet_and_variable_in_case_note("Interview date", interview_date)
If HC_document_received <> "" and HC_document_received <> " " then call write_bullet_and_variable_in_case_note("HC document received", HC_document_received)
If HC_datestamp <> "" then call write_bullet_and_variable_in_case_note("HC datestamp", HC_datestamp)
call write_bullet_and_variable_in_case_note("Programs applied for", programs_applied_for)
if how_app_was_received <> "" or how_app_was_received <> " " then call write_bullet_and_variable_in_case_note("How CAF was received", how_app_was_received)	'This one also uses " " as option, because that is the default
If HH_comp <> "" then call write_bullet_and_variable_in_case_note("HH comp/EATS", HH_comp)
If cit_id <> "" then call write_bullet_and_variable_in_case_note("Cit/ID", cit_id)
If IMIG <> "" then call write_bullet_and_variable_in_case_note("IMIG", IMIG)
If AREP <> "" then call write_bullet_and_variable_in_case_note("AREP", AREP)
If FACI <> "" then call write_bullet_and_variable_in_case_note("FACI", FACI)
If SCHL <> "" then call write_bullet_and_variable_in_case_note("SCHL/STIN/STEC", SCHL)
If DISA <> "" then call write_bullet_and_variable_in_case_note("DISA", DISA)
If PREG <> "" then call write_bullet_and_variable_in_case_note("PREG", PREG)
If ABPS <> "" then call write_bullet_and_variable_in_case_note("ABPS", ABPS)
If earned_income <> "" then call write_bullet_and_variable_in_case_note("Earned income", earned_income)
If unearned_income <> "" then call write_bullet_and_variable_in_case_note("Unearned income", unearned_income)
If income_changes <> "" then call write_bullet_and_variable_in_case_note("STWK/inc. changes", income_changes)
IF notes_on_abawd <> "" then call write_bullet_and_variable_in_case_note("ABAWD Notes", notes_on_abawd)
If notes_on_income <> "" then call write_bullet_and_variable_in_case_note("Notes on income and budget", notes_on_income)
If is_any_work_temporary <> "" then call write_bullet_and_variable_in_case_note("Is any work temporary", is_any_work_temporary)
If SHEL_HEST <> "" then call write_bullet_and_variable_in_case_note("SHEL/HEST", SHEL_HEST)
If COEX_DCEX <> "" then call write_bullet_and_variable_in_case_note("COEX/DCEX", COEX_DCEX)
If CASH_ACCTs <> "" then call write_bullet_and_variable_in_case_note("CASH/ACCTs", CASH_ACCTs)
If other_assets <> "" then call write_bullet_and_variable_in_case_note("Other assets", other_assets)
If INSA <> "" then call write_bullet_and_variable_in_case_note("INSA", INSA)
If ACCI <> "" then call write_bullet_and_variable_in_case_note("ACCI", ACCI)
If DIET <> "" then call write_bullet_and_variable_in_case_note("DIET", DIET)
If BILS <> "" then call write_bullet_and_variable_in_case_note("BILS", BILS)
If FMED <> "" then call write_bullet_and_variable_in_case_note("FMED", FMED)
If retro_request <> "" then call write_bullet_and_variable_in_case_note("Retro Request (if applicable)", retro_request)
If application_signed_checkbox = checked then call write_variable_in_case_note("* Application was signed.")
If application_signed_checkbox = unchecked then call write_variable_in_case_note("* Application was not signed.")
If appt_letter_sent_checkbox = checked then call write_variable_in_case_note("* Appointment letter was sent before interview.")
If EBT_referral_checkbox = checked then call write_variable_in_case_note("* EBT referral made for client.")
If eDRS_sent_checkbox = checked then call write_variable_in_case_note("* eDRS sent.")
If expedited_checkbox = checked then call write_variable_in_case_note("* Expedited SNAP.")
If reason_expedited_wasnt_processed <> "" then call write_bullet_and_variable_in_case_note("Reason expedited wasn't processed", reason_expedited_wasnt_processed)	'This is strategically placed next to expedited checkbox entry.
If IAA_checkbox = checked then call write_variable_in_case_note("* IAAs/OMB given to client.")
If intake_packet_checkbox = checked then call write_variable_in_case_note("* Client received intake packet.")
If managed_care_packet_checkbox = checked then call write_variable_in_case_note("* Client received managed care packet.")
If managed_care_referral_checkbox = checked then call write_variable_in_case_note("* Managed care referral made.")
If R_R_checkbox = checked then call write_variable_in_case_note("* R/R explained to client.")
If updated_MMIS_checkbox = checked then call write_variable_in_case_note("* Updated MMIS.")
If WF1_checkbox = checked then call write_variable_in_case_note("* Workforce referral made.")
If client_delay_checkbox = checked then call write_variable_in_case_note("* PND2 updated to show client delay.")
if FIAT_reasons <> "" then call write_bullet_and_variable_in_case_note("FIAT reasons", FIAT_reasons)
if other_notes <> "" then call write_bullet_and_variable_in_case_note("Other notes", other_notes)
If move_verifs_needed = False and verifs_needed <> "" then call write_bullet_and_variable_in_case_note("Verifs needed", verifs_needed)		'If global variable move_verifs_needed = False (on FUNCTIONS FILE), it'll case note at the bottom.
call write_bullet_and_variable_in_case_note("Actions taken", actions_taken)
call write_variable_in_case_note("---")
call write_variable_in_case_note(worker_signature)

script_end_procedure("")
