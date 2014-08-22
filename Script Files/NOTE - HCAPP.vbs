'Grabbing stats----------------------------------------------------------------------------------------------------
name_of_script = "NOTE - HCAPP"
start_time = timer

'LOADING ROUTINE FUNCTIONS----------------------------------------------------------------------------------------------------
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\MAXIS-BZ-Scripts-County-Beta\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'DATE CALCULATIONS----------------------------------------------------------------------------------------------------

next_month = dateadd("m", + 1, date)
footer_month = datepart("m", next_month)
If len(footer_month) = 1 then footer_month = "0" & footer_month
footer_year = datepart("yyyy", next_month)
footer_year = "" & footer_year - 2000


'DIALOGS-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

BeginDialog case_number_and_footer_month_dialog, 0, 0, 161, 65, "Case number and footer month"
  Text 5, 10, 85, 10, "Enter your case number:"
  EditBox 95, 5, 60, 15, case_number
  Text 15, 30, 50, 10, "Footer month:"
  EditBox 65, 25, 25, 15, footer_month
  Text 95, 30, 20, 10, "Year:"
  EditBox 120, 25, 25, 15, footer_year
  ButtonGroup ButtonPressed
    OkButton 25, 45, 50, 15
    CancelButton 85, 45, 50, 15
EndDialog


BeginDialog HCAPP_dialog_01, 0, 0, 446, 300, "HCAPP dialog part 1"
  EditBox 75, 5, 50, 15, HCAPP_datestamp
  EditBox 45, 25, 250, 15, HH_comp
  EditBox 45, 45, 225, 15, cit_id
  EditBox 330, 45, 115, 15, AREP
  EditBox 85, 65, 90, 15, SCHL
  EditBox 235, 65, 210, 15, DISA
  EditBox 35, 85, 160, 15, PREG
  EditBox 235, 85, 210, 15, ABPS
  EditBox 60, 105, 180, 15, retro_request
  EditBox 60, 140, 385, 15, earned_income
  EditBox 70, 160, 375, 15, unearned_income
  EditBox 35, 180, 245, 15, STWK
  EditBox 350, 180, 95, 15, COEX_DCEX
  EditBox 65, 200, 380, 15, notes_on_income
  EditBox 155, 220, 290, 15, is_any_work_temporary
  EditBox 55, 255, 390, 15, verifs_needed
  ButtonGroup ButtonPressed
    PushButton 335, 280, 50, 15, "NEXT", next_page_button
    CancelButton 390, 280, 50, 15
    PushButton 335, 15, 45, 10, "prev. panel", prev_panel_button
    PushButton 335, 25, 45, 10, "next panel", next_panel_button
    PushButton 395, 15, 45, 10, "prev. memb", prev_memb_button
    PushButton 395, 25, 45, 10, "next memb", next_memb_button
    PushButton 275, 50, 25, 10, "AREP/", AREP_button
    PushButton 300, 50, 25, 10, "ALTP:", ALTP_button
    PushButton 5, 70, 25, 10, "SCHL/", SCHL_button
    PushButton 30, 70, 25, 10, "STIN/", STIN_button
    PushButton 55, 70, 25, 10, "STEC:", STEC_button
    PushButton 180, 70, 25, 10, "DISA/", DISA_button
    PushButton 205, 70, 25, 10, "PDED:", PDED_button
    PushButton 5, 90, 25, 10, "PREG:", PREG_button
    PushButton 205, 90, 25, 10, "ABPS:", ABPS_button
    PushButton 5, 185, 25, 10, "STWK:", STWK_button
    PushButton 295, 185, 25, 10, "COEX/", COEX_button
    PushButton 320, 185, 25, 10, "DCEX:", DCEX_button
    PushButton 10, 285, 25, 10, "BUSI", BUSI_button
    PushButton 35, 285, 25, 10, "JOBS", JOBS_button
    PushButton 60, 285, 25, 10, "RBIC", RBIC_button
    PushButton 85, 285, 25, 10, "UNEA", UNEA_button
    PushButton 125, 285, 25, 10, "MEMB", MEMB_button
    PushButton 150, 285, 25, 10, "MEMI", MEMI_button
    PushButton 175, 285, 25, 10, "REVW", REVW_button
    PushButton 215, 285, 35, 10, "ELIG/HC", ELIG_HC_button
  GroupBox 330, 5, 115, 35, "STAT-based navigation"
  Text 5, 10, 65, 10, "HCAPP datestamp:"
  Text 5, 30, 35, 10, "HH comp:"
  Text 5, 50, 40, 10, "Cit/ID/imig:"
  Text 5, 110, 50, 10, "Retro request:"
  Text 5, 145, 55, 10, "Earned income:"
  Text 5, 165, 65, 10, "Unearned income:"
  Text 5, 205, 60, 10, "Notes on income:"
  Text 5, 225, 150, 10, "Is any work temporary? If so, explain details:"
  Text 5, 260, 50, 10, "Verifs needed:"
  GroupBox 5, 275, 110, 25, "Income panels"
  GroupBox 120, 275, 85, 25, "other STAT panels:"
EndDialog

BeginDialog HCAPP_dialog_02, 0, 0, 451, 310, "HCAPP dialog part 2"
  EditBox 35, 50, 410, 15, assets
  EditBox 60, 80, 385, 15, INSA
  EditBox 35, 100, 410, 15, ACCI
  EditBox 35, 120, 410, 15, BILS
  EditBox 125, 140, 125, 15, FACI
  CheckBox 255, 145, 80, 10, "Application signed?", application_signed_check
  CheckBox 350, 145, 65, 10, "MMIS updated?", MMIS_updated_check
  CheckBox 20, 160, 290, 10, "Check here to have the script update PND2 to show client delay (pending cases only).", client_delay_check
  CheckBox 20, 175, 245, 10, "Check here to have the script create a TIKL to deny at the 45 day mark.", TIKL_check
  EditBox 100, 190, 345, 15, FIAT_reasons
  EditBox 55, 210, 215, 15, other_notes
  ComboBox 330, 210, 115, 15, ""+chr(9)+"incomplete"+chr(9)+"approved", HCAPP_status
  EditBox 55, 230, 390, 15, verifs_needed
  EditBox 55, 250, 390, 15, actions_taken
  EditBox 395, 270, 50, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 340, 290, 50, 15
    CancelButton 395, 290, 50, 15
    PushButton 10, 15, 25, 10, "ACCT", ACCT_button
    PushButton 35, 15, 25, 10, "CARS", CARS_button
    PushButton 60, 15, 25, 10, "CASH", CASH_button
    PushButton 85, 15, 25, 10, "OTHR", OTHR_button
    PushButton 10, 25, 25, 10, "REST", REST_button
    PushButton 35, 25, 25, 10, "SECU", SECU_button
    PushButton 60, 25, 25, 10, "TRAN", TRAN_button
    PushButton 335, 15, 45, 10, "prev. panel", prev_panel_button
    PushButton 335, 25, 45, 10, "next panel", next_panel_button
    PushButton 395, 15, 45, 10, "prev. memb", prev_memb_button
    PushButton 395, 25, 45, 10, "next memb", next_memb_button
    PushButton 5, 85, 25, 10, "INSA/", INSA_button
    PushButton 30, 85, 25, 10, "MEDI:", MEDI_button
    PushButton 5, 105, 25, 10, "ACCI:", ACCI_button
    PushButton 5, 125, 25, 10, "BILS:", BILS_button
    PushButton 5, 145, 25, 10, "FACI/", FACI_button
    PushButton 10, 280, 25, 10, "MEMB", MEMB_button
    PushButton 35, 280, 25, 10, "MEMI", MEMI_button
    PushButton 60, 280, 25, 10, "REVW", REVW_button
    PushButton 95, 280, 35, 10, "ELIG/HC", ELIG_HC_button
    PushButton 225, 295, 75, 10, "previous page", previous_page_button
  GroupBox 5, 5, 110, 35, "Asset panels"
  GroupBox 330, 5, 115, 35, "STAT-based navigation"
  Text 5, 55, 30, 10, "Assets:"
  Text 35, 145, 90, 10, "residency/miscellaneous:"
  Text 5, 195, 95, 10, "FIAT reasons (if applicable):"
  Text 5, 215, 45, 10, "Other notes:"
  Text 280, 215, 50, 10, "HCAPP status:"
  Text 5, 235, 50, 10, "Verifs needed:"
  Text 5, 255, 50, 10, "Actions taken:"
  GroupBox 5, 270, 85, 25, "other STAT panels:"
  Text 330, 275, 65, 10, "Worker signature:"
EndDialog

BeginDialog case_note_dialog, 0, 0, 136, 51, "Case note dialog"
  ButtonGroup ButtonPressed
    PushButton 15, 20, 105, 10, "Yes, take me to case note.", yes_case_note_button
    PushButton 5, 35, 125, 10, "No, take me back to the script dialog.", no_case_note_button
  Text 10, 5, 125, 10, "Are you sure you want to case note?"
EndDialog

BeginDialog cancel_dialog, 0, 0, 141, 51, "Cancel dialog"
  Text 5, 5, 135, 10, "Are you sure you want to end this script?"
  ButtonGroup ButtonPressed
    PushButton 10, 20, 125, 10, "No, take me back to the script dialog.", no_cancel_button
    PushButton 20, 35, 105, 10, "Yes, close this script.", yes_cancel_button
EndDialog

'VARIABLES WHICH NEED DECLARING------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
HH_memb_row = 5
Dim row
Dim col
HC_check = 1 'This is so the functions will work without having to select a program. It uses the same dialogs as the CSR, which can look in multiple places. This is HC only, so it doesn't need those.
application_signed_check = 1 'The script should default to having the application signed.


'THE SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------

'Connecting to MAXIS
EMConnect ""

'Grabbing the case number
call find_variable("Case Nbr: ", case_number, 8)
case_number = trim(case_number)
case_number = replace(case_number, "_", "")
If IsNumeric(case_number) = False then case_number = ""

'Grabbing the footer month/year
call find_variable("Month: ", MAXIS_footer_month, 2)
If row <> 0 then 
  footer_month = MAXIS_footer_month
  call find_variable("Month: " & footer_month & " ", MAXIS_footer_year, 2)
  If row <> 0 then footer_year = MAXIS_footer_year
End if

'Showing the case number
Do
  Dialog case_number_and_footer_month_dialog
  If ButtonPressed = 0 then stopscript
  If case_number = "" or IsNumeric(case_number) = False or len(case_number) > 8 then MsgBox "You need to type a valid case number."
Loop until case_number <> "" and IsNumeric(case_number) = True and len(case_number) <= 8
transmit

'Checking to see that we're in MAXIS
EMReadScreen MAXIS_check, 5, 1, 39
If MAXIS_check <> "MAXIS" and MAXIS_check <> "AXIS " then call script_end_procedure("You are not in MAXIS or you are locked out of your case.")

'Navigating to STAT, grabbing the HH members
call navigate_to_screen("stat", "hcre")
EMReadScreen STAT_check, 4, 20, 21
If STAT_check <> "STAT" then call script_end_procedure("Can't get in to STAT. This case may be in background. Wait a few seconds and try again. If the case is not in background contact a Support Team member.")
EMReadScreen ERRR_check, 4, 2, 52
If ERRR_check = "ERRR" then transmit 'For error prone cases.

'Creating a custom dialog for determining who the HH members are
call HH_member_custom_dialog(HH_member_array)

'Autofilling case info
call autofill_editbox_from_MAXIS(HH_member_array, "HCRE-retro", retro_request)
call autofill_editbox_from_MAXIS(HH_member_array, "HCRE", HCAPP_datestamp)
call autofill_editbox_from_MAXIS(HH_member_array, "ABPS", ABPS)
call autofill_editbox_from_MAXIS(HH_member_array, "ACCI", ACCI)
call autofill_editbox_from_MAXIS(HH_member_array, "ACCT", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "AREP", AREP)
call autofill_editbox_from_MAXIS(HH_member_array, "BUSI", earned_income)
call autofill_editbox_from_MAXIS(HH_member_array, "CASH", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "CARS", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "BILS", BILS)
call autofill_editbox_from_MAXIS(HH_member_array, "COEX", COEX_DCEX)
call autofill_editbox_from_MAXIS(HH_member_array, "DCEX", COEX_DCEX)
call autofill_editbox_from_MAXIS(HH_member_array, "DISA", DISA)
call autofill_editbox_from_MAXIS(HH_member_array, "FACI", FACI)
call autofill_editbox_from_MAXIS(HH_member_array, "INSA", INSA)
call autofill_editbox_from_MAXIS(HH_member_array, "JOBS", earned_income)
call autofill_editbox_from_MAXIS(HH_member_array, "MEDI", INSA)
call autofill_editbox_from_MAXIS(HH_member_array, "MEMB", HH_comp)
call autofill_editbox_from_MAXIS(HH_member_array, "MEMI", cit_id)
call autofill_editbox_from_MAXIS(HH_member_array, "OTHR", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "PREG", PREG)
call autofill_editbox_from_MAXIS(HH_member_array, "RBIC", earned_income)
call autofill_editbox_from_MAXIS(HH_member_array, "REST", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "SCHL", SCHL)
call autofill_editbox_from_MAXIS(HH_member_array, "SECU", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "STWK", STWK)
call autofill_editbox_from_MAXIS(HH_member_array, "UNEA", unearned_income)

'Cleaning up data fields. This could probably get removed if this can be successfully put into the autofill function.
earned_income = trim(earned_income)
if right(earned_income, 1) = ";" then earned_income = left(earned_income, len(earned_income) - 1)
earned_income = replace(earned_income, "$________/non-monthly", "amt unknown")
earned_income = replace(earned_income, "$________/monthly", "amt unknown")
earned_income = replace(earned_income, "$________/weekly", "amt unknown")
earned_income = replace(earned_income, "$________/biweekly", "amt unknown")
earned_income = replace(earned_income, "$________/semimonthly", "amt unknown")
unearned_income = trim(unearned_income)
if right(unearned_income, 1) = ";" then unearned_income = left(unearned_income, len(unearned_income) - 1)
unearned_income = replace(unearned_income, "$________/non-monthly", "amt unknown")
unearned_income = replace(unearned_income, "$________/monthly", "amt unknown")
unearned_income = replace(unearned_income, "$________/weekly", "amt unknown")
unearned_income = replace(unearned_income, "$________/biweekly", "amt unknown")
unearned_income = replace(unearned_income, "$________/semimonthly", "amt unknown")
assets = trim(assets)
if right(assets, 1) = ";" then assets = left(assets, len(assets) - 1)
COEX_DCEX = trim(COEX_DCEX)
if right(COEX_DCEX, 1) = ";" then COEX_DCEX = left(COEX_DCEX, len(COEX_DCEX) - 1)
SCHL = trim(SCHL)
if right(SCHL, 1) = ";" then SCHL = left(SCHL, len(SCHL) - 1)
DISA = trim(DISA)
if right(DISA, 1) = ";" then DISA = left(DISA, len(DISA) - 1)
FACI = trim(FACI)
if right(FACI, 1) = ";" then FACI = left(FACI, len(FACI) - 1)
INSA = trim(INSA)
if right(INSA, 1) = ";" then INSA = left(INSA, len(INSA) - 1)
ACCI = trim(ACCI)
if right(ACCI, 1) = ";" then ACCI = left(ACCI, len(ACCI) - 1)
cit_ID = trim(cit_ID)
if right(cit_ID, 1) = ";" then cit_ID = left(cit_ID, len(cit_ID) - 1)
PREG = trim(PREG)
if right(PREG, 1) = ";" then PREG = left(PREG, len(PREG) - 1)
STWK = trim(STWK)
if right(STWK, 1) = ";" then STWK = left(STWK, len(STWK) - 1)


'SECTION 07: CASE NOTE DIALOG--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Do
  Do
    Do
      Do
        Do
          Dialog HCAPP_dialog_01
          If ButtonPressed = 0 then 
            dialog cancel_dialog
            If ButtonPressed = yes_cancel_button then stopscript
          End if
        Loop until ButtonPressed <> no_cancel_button
        EMReadScreen STAT_check, 4, 20, 21
        If STAT_check = "STAT" then
          If ButtonPressed = prev_panel_button then call panel_navigation_prev
          If ButtonPressed = next_panel_button then call panel_navigation_next
          If ButtonPressed = prev_memb_button then call memb_navigation_prev
          If ButtonPressed = next_memb_button then call memb_navigation_next
        End if
        transmit 'Forces a screen refresh, to keep MAXIS from erroring out in the event of a password prompt.
        EMReadScreen MAXIS_check, 5, 1, 39
        If MAXIS_check <> "MAXIS" and MAXIS_check <> "AXIS " then MsgBox "You do not appear to be in MAXIS. Are you passworded out? Or in MMIS? Check these and try again."
      Loop until MAXIS_check = "MAXIS" or MAXIS_check = "AXIS " 
      If ButtonPressed = ELIG_HC_button then call navigate_to_screen("elig", "HC__")
      If ButtonPressed = ABPS_button then call navigate_to_screen("stat", "ABPS")
      If ButtonPressed = AREP_button then call navigate_to_screen("stat", "AREP")
      If ButtonPressed = ALTP_button then call navigate_to_screen("stat", "ALTP")
      If ButtonPressed = SCHL_button then call navigate_to_screen("stat", "SCHL")
      If ButtonPressed = STIN_button then call navigate_to_screen("stat", "STIN")
      If ButtonPressed = STEC_button then call navigate_to_screen("stat", "STEC")
      If ButtonPressed = DISA_button then call navigate_to_screen("stat", "DISA")
      If ButtonPressed = PDED_button then call navigate_to_screen("stat", "PDED")
      If ButtonPressed = BUSI_button then call navigate_to_screen("stat", "BUSI")
      If ButtonPressed = JOBS_button then call navigate_to_screen("stat", "JOBS")
      If ButtonPressed = RBIC_button then call navigate_to_screen("stat", "RBIC")
      If ButtonPressed = UNEA_button then call navigate_to_screen("stat", "UNEA")
      If ButtonPressed = COEX_button then call navigate_to_screen("stat", "COEX")
      If ButtonPressed = DCEX_button then call navigate_to_screen("stat", "DCEX")
      If ButtonPressed = MEMB_button then call navigate_to_screen("stat", "MEMB")
      If ButtonPressed = MEMI_button then call navigate_to_screen("stat", "MEMI")
      If ButtonPressed = REVW_button then call navigate_to_screen("stat", "REVW")
    Loop until ButtonPressed = next_page_button
    Do
      Do
        Do
          Dialog HCAPP_dialog_02
          If ButtonPressed = 0 then 
            dialog cancel_dialog
            If ButtonPressed = yes_cancel_button then stopscript
          End if
        Loop until ButtonPressed <> no_cancel_button
        EMReadScreen STAT_check, 4, 20, 21
        If STAT_check = "STAT" then
          If ButtonPressed = prev_panel_button then call panel_navigation_prev
          If ButtonPressed = next_panel_button then call panel_navigation_next
          If ButtonPressed = prev_memb_button then call memb_navigation_prev
          If ButtonPressed = next_memb_button then call memb_navigation_next
        End if
        transmit 'Forces a screen refresh, to keep MAXIS from erroring out in the event of a password prompt.
        EMReadScreen MAXIS_check, 5, 1, 39
        If MAXIS_check <> "MAXIS" and MAXIS_check <> "AXIS " then MsgBox "You do not appear to be in MAXIS. Are you passworded out? Or in MMIS? Check these and try again."
      Loop until MAXIS_check = "MAXIS" or MAXIS_check = "AXIS " 
      If ButtonPressed = ELIG_HC_button then call navigate_to_screen("elig", "HC__")
      If ButtonPressed = ACCT_button then call navigate_to_screen("stat", "ACCT")
      If ButtonPressed = CARS_button then call navigate_to_screen("stat", "CARS")
      If ButtonPressed = CASH_button then call navigate_to_screen("stat", "CASH")
      If ButtonPressed = OTHR_button then call navigate_to_screen("stat", "OTHR")
      If ButtonPressed = REST_button then call navigate_to_screen("stat", "REST")
      If ButtonPressed = SECU_button then call navigate_to_screen("stat", "SECU")
      If ButtonPressed = TRAN_button then call navigate_to_screen("stat", "TRAN")
      If ButtonPressed = INSA_button then call navigate_to_screen("stat", "INSA")
      If ButtonPressed = MEDI_button then call navigate_to_screen("stat", "MEDI")
      If ButtonPressed = ACCI_button then call navigate_to_screen("stat", "ACCI")
      If ButtonPressed = BILS_button then call navigate_to_screen("stat", "BILS")
      If ButtonPressed = FACI_button then call navigate_to_screen("stat", "FACI")
      If ButtonPressed = MEMB_button then call navigate_to_screen("stat", "MEMB")
      If ButtonPressed = MEMI_button then call navigate_to_screen("stat", "MEMI")
      If ButtonPressed = REVW_button then call navigate_to_screen("stat", "REVW")
    Loop until ButtonPressed = -1 or ButtonPressed = previous_page_button
    If ButtonPressed = previous_page_button then exit do
    If actions_taken = "" or HCAPP_datestamp = "" or worker_signature = "" then MsgBox "You need to fill in the datestamp and actions taken sections, as well as sign your case note. Check these items after pressing ''OK''."
  Loop until actions_taken <> "" and HCAPP_datestamp <> "" and worker_signature <> "" 
  If ButtonPressed = -1 then dialog case_note_dialog
  If buttonpressed = yes_case_note_button then
    If client_delay_check = 1 then 'UPDATES PND2 FOR CLIENT DELAY IF CHECKED
      call navigate_to_screen("rept", "pnd2")
      EMGetCursor PND2_row, PND2_col
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
      PF3
      EMReadScreen PND2_check, 4, 2, 52
      If PND2_check = "PND2" then
        MsgBox "PND2 might not have been updated for client delay. There may have been a MAXIS error. Check this manually after case noting."
        PF10
        client_delay_check = 0
      End if
    End if
    If TIKL_check = 1 then
      call navigate_to_screen("dail", "writ")
      call create_MAXIS_friendly_date(HCAPP_datestamp, 45, 5, 18) 
      EMSetCursor 9, 3
      EMSendKey "HC pending 45 days. Evaluate for possible denial. If any members are elderly/disabled, allow an additional 15 days and reTIKL out."
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

'Adding a colon to the beginning of the HCAPP status variable if it isn't blank (simplifies writing the header of the case note)
If HCAPP_status <> "" then HCAPP_status = ": " & HCAPP_status

'SECTION 08: THE CASE NOTE-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

EMSendKey "<home>" & "***HCAPP received " & HCAPP_datestamp & HCAPP_status & "***" & "<newline>"
If HH_comp <> "" then call write_editbox_in_case_note("HH comp", HH_comp, 6)
If cit_id <> "" then call write_editbox_in_case_note("Cit/ID", cit_id, 6)
If AREP <> "" then call write_editbox_in_case_note("AREP", AREP, 6)
If SCHL <> "" then call write_editbox_in_case_note("SCHL/STIN/STEC", SCHL, 6)
If DISA <> "" then call write_editbox_in_case_note("DISA", DISA, 6)
If PREG <> "" then call write_editbox_in_case_note("PREG", PREG, 6)
If retro_request <> "" then call write_editbox_in_case_note("Retro request", retro_request, 6)
If ABPS <> "" then call write_editbox_in_case_note("ABPS", ABPS, 6)
If earned_income <> "" then call write_editbox_in_case_note("Earned income", earned_income, 6)
If unearned_income <> "" then call write_editbox_in_case_note("Unearned income", unearned_income, 6)
If STWK <> "" then call write_editbox_in_case_note("STWK", STWK, 6)
If COEX_DCEX <> "" then call write_editbox_in_case_note("COEX/DCEX", COEX_DCEX, 6)
If notes_on_income <> "" then call write_editbox_in_case_note("Notes on income", notes_on_income, 6)
If is_any_work_temporary <> "" then call write_editbox_in_case_note("Is any work temporary", is_any_work_temporary, 6)
If assets <> "" then call write_editbox_in_case_note("Assets", assets, 6)
If INSA <> "" then call write_editbox_in_case_note("INSA", INSA, 6)
If ACCI <> "" then call write_editbox_in_case_note("ACCI", ACCI, 6)
If BILS <> "" then call write_editbox_in_case_note("BILS", BILS, 6)
If FACI <> "" then call write_editbox_in_case_note("FACI", FACI, 6)
If application_signed_check = 1 then call write_new_line_in_case_note("* Application was signed.")
If application_signed_check = 0 then call write_new_line_in_case_note("* Application was not signed.")
If client_delay_check = 1 then call write_new_line_in_case_note("* PND2 updated to show client delay.")
if FIAT_reasons <> "" then call write_editbox_in_case_note("FIAT reasons", FIAT_reasons, 6)
if other_notes <> "" then call write_editbox_in_case_note("Other notes", other_notes, 6)
if verifs_needed <> "" then call write_editbox_in_case_note("Verifs needed", verifs_needed, 6)
call write_editbox_in_case_note("Actions taken", actions_taken, 6)
If MMIS_update_check = 1 then call write_new_line_in_case_note("* MMIS updated.")
call write_new_line_in_case_note("---")
call write_new_line_in_case_note(worker_signature)

call script_end_procedure("")






