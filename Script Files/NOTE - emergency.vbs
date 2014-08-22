'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTE - Emergency"
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
BeginDialog case_number_dialog, 0, 0, 181, 97, "Case number dialog"
  EditBox 80, 5, 70, 15, case_number
  EditBox 65, 25, 30, 15, footer_month
  EditBox 140, 25, 30, 15, footer_year
  CheckBox 10, 60, 30, 10, "cash", cash_check
  CheckBox 50, 60, 30, 10, "HC", HC_check
  CheckBox 90, 60, 35, 10, "SNAP", SNAP_check
  ButtonGroup ButtonPressed
    OkButton 35, 80, 50, 15
    CancelButton 95, 80, 50, 15
  Text 25, 10, 50, 10, "Case number:"
  Text 10, 30, 50, 10, "Footer month:"
  Text 110, 30, 25, 10, "Year:"
  GroupBox 5, 45, 170, 30, "Other programs open or applied for:"
EndDialog

BeginDialog emergency_dialog, 0, 0, 321, 380, "Emergency Dialog"
  EditBox 60, 45, 65, 15, interview_date
  EditBox 170, 45, 150, 15, HH_comp
  CheckBox 25, 75, 40, 10, "Eviction", eviction_check
  CheckBox 75, 75, 70, 10, "Utility disconnect", utility_disconnect_check
  CheckBox 155, 75, 60, 10, "Homelessness", homelessness_check
  CheckBox 230, 75, 65, 10, "Security deposit", security_deposit_check
  EditBox 65, 100, 255, 15, cause_of_crisis
  EditBox 85, 160, 235, 15, income
  EditBox 110, 180, 210, 15, income_under_200_FPG
  EditBox 60, 200, 260, 15, thirty_percent_rule_notes
  EditBox 75, 220, 245, 15, monthly_expense
  EditBox 40, 240, 280, 15, assets
  EditBox 60, 260, 260, 15, verifs_needed
  EditBox 75, 280, 245, 15, crisis_resolvable
  EditBox 80, 300, 240, 15, discussion_of_crisis
  EditBox 60, 320, 260, 15, actions_taken
  EditBox 50, 340, 270, 15, referrals
  EditBox 75, 360, 90, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 195, 360, 50, 15
    CancelButton 255, 360, 50, 15
    PushButton 10, 15, 25, 10, "ADDR", ADDR_button
    PushButton 35, 15, 25, 10, "MEMB", MEMB_button
    PushButton 60, 15, 25, 10, "MEMI", MEMI_button
    PushButton 10, 25, 25, 10, "PROG", PROG_button
    PushButton 35, 25, 25, 10, "TYPE", TYPE_button
    PushButton 125, 20, 50, 10, "ELIG/EMER", ELIG_EMER_button
    PushButton 210, 15, 45, 10, "prev. panel", prev_panel_button
    PushButton 210, 25, 45, 10, "next panel", next_panel_button
    PushButton 270, 15, 45, 10, "prev. memb", prev_memb_button
    PushButton 270, 25, 45, 10, "next memb", next_memb_button
    PushButton 75, 130, 25, 10, "BUSI", BUSI_button
    PushButton 100, 130, 25, 10, "JOBS", JOBS_button
    PushButton 75, 140, 25, 10, "RBIC", RBIC_button
    PushButton 100, 140, 25, 10, "UNEA", UNEA_button
    PushButton 150, 130, 25, 10, "ACCT", ACCT_button
    PushButton 175, 130, 25, 10, "CARS", CARS_button
    PushButton 200, 130, 25, 10, "CASH", CASH_button
    PushButton 225, 130, 25, 10, "OTHR", OTHR_button
    PushButton 150, 140, 25, 10, "REST", REST_button
    PushButton 175, 140, 25, 10, "SECU", SECU_button
    PushButton 200, 140, 25, 10, "TRAN", TRAN_button
  GroupBox 5, 5, 85, 35, "other STAT panels:"
  GroupBox 205, 5, 115, 35, "STAT-based navigation"
  Text 5, 50, 50, 10, "Interview date:"
  Text 130, 50, 35, 10, "HH Comp:"
  GroupBox 20, 65, 280, 25, "Crisis (check all that apply):"
  Text 5, 105, 55, 10, "Cause of crisis:"
  GroupBox 70, 120, 60, 35, "Income panels"
  GroupBox 145, 120, 110, 35, "Asset panels"
  Text 5, 165, 75, 10, "Income (past 30 days):"
  Text 5, 185, 100, 10, "Is income under 200% FPG?:"
  Text 5, 205, 55, 10, "30% rule notes:"
  Text 5, 225, 60, 10, "Monthly expense:"
  Text 5, 245, 30, 10, "Assets:"
  Text 5, 265, 50, 10, "Verifs needed:"
  Text 5, 285, 65, 10, "Crisis resolvable?:"
  Text 5, 305, 75, 10, "Discussion of Crisis:"
  Text 5, 325, 50, 10, "Actions taken:"
  Text 5, 345, 40, 10, "Referrals:"
  Text 5, 365, 65, 10, "Worker signature:"
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
application_signed_check = 1 'The script should default to having the application signed.

'THE SCRIPT--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'Connecting to BlueZone
EMConnect ""

'Grabbing the case number
call find_variable("Case Nbr: ", case_number, 8)
case_number = trim(case_number)
case_number = replace(case_number, "_", "")
If IsNumeric(case_number) = False then case_number = ""

'Grabbing the footer month
call find_variable("Month: ", MAXIS_footer_month, 2)
If row <> 0 then 
  footer_month = MAXIS_footer_month
  call find_variable("Month: " & footer_month & " ", MAXIS_footer_year, 2)
  If row <> 0 then footer_year = MAXIS_footer_year
End if

'Showing the case number dialog
Do
  Dialog case_number_dialog
  If ButtonPressed = 0 then stopscript
  If case_number = "" or IsNumeric(case_number) = False or len(case_number) > 8 then MsgBox "You need to type a valid case number."
Loop until case_number <> "" and IsNumeric(case_number) = True and len(case_number) <= 8

'Checking for MAXIS
transmit
EMReadScreen MAXIS_check, 5, 1, 39
If MAXIS_check <> "MAXIS" and MAXIS_check <> "AXIS " then script_end_procedure("You are not in MAXIS, or you are locked out of your case.")

'Jumping into STAT
call navigate_to_screen("stat", "hcre")
EMReadScreen STAT_check, 4, 20, 21
If STAT_check <> "STAT" then script_end_procedure("Can't get in to STAT. This case may be in background. Wait a few seconds and try again. If the case is not in background contact a Support Team member.")
EMReadScreen ERRR_check, 4, 2, 52
If ERRR_check = "ERRR" then transmit 'For error prone cases.

'Creating a custom dialog for determining who the HH members are
call HH_member_custom_dialog(HH_member_array)

'Autofilling
call autofill_editbox_from_MAXIS(HH_member_array, "ACCT", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "BUSI", income)
call autofill_editbox_from_MAXIS(HH_member_array, "CARS", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "CASH", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "COEX", monthly_expense)
call autofill_editbox_from_MAXIS(HH_member_array, "DCEX", monthly_expense)
call autofill_editbox_from_MAXIS(HH_member_array, "JOBS", income)
call autofill_editbox_from_MAXIS(HH_member_array, "MEMB", HH_comp)
call autofill_editbox_from_MAXIS(HH_member_array, "OTHR", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "RBIC", income)
call autofill_editbox_from_MAXIS(HH_member_array, "REST", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "SECU", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "SHEL", monthly_expense)
call autofill_editbox_from_MAXIS(HH_member_array, "UNEA", income)
call autofill_editbox_from_MAXIS(HH_member_array, "HEST", monthly_expense) 'Does this last because people like it tacked on to the end, not before. The rest are alphabetical.


'Showing the case note
Do
  Do
    Do
      Do
        Do
          Dialog emergency_dialog
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
      If ButtonPressed = ADDR_button then call navigate_to_screen("stat", "ADDR")
      If ButtonPressed = MEMB_button then call navigate_to_screen("stat", "MEMB")
      If ButtonPressed = MEMI_button then call navigate_to_screen("stat", "MEMI")
      If ButtonPressed = PROG_button then call navigate_to_screen("stat", "PROG")
      If ButtonPressed = TYPE_button then call navigate_to_screen("stat", "TYPE")
      If ButtonPressed = ELIG_EMER_button then call navigate_to_screen("elig", "emer")
      If ButtonPressed = BUSI_button then call navigate_to_screen("stat", "BUSI")
      If ButtonPressed = JOBS_button then call navigate_to_screen("stat", "JOBS")
      If ButtonPressed = RBIC_button then call navigate_to_screen("stat", "RBIC")
      If ButtonPressed = UNEA_button then call navigate_to_screen("stat", "UNEA")
      If ButtonPressed = ACCT_button then call navigate_to_screen("stat", "ACCT")
      If ButtonPressed = CARS_button then call navigate_to_screen("stat", "CARS")
      If ButtonPressed = CASH_button then call navigate_to_screen("stat", "CASH")
      If ButtonPressed = OTHR_button then call navigate_to_screen("stat", "OTHR")
      If ButtonPressed = REST_button then call navigate_to_screen("stat", "REST")
      If ButtonPressed = SECU_button then call navigate_to_screen("stat", "SECU")
      If ButtonPressed = TRAN_button then call navigate_to_screen("stat", "TRAN")
    Loop until ButtonPressed = -1
    If income = "" or actions_taken = "" or worker_signature = "" then MsgBox "You need to fill in the income and actions taken sections, as well as sign your case note. Check these items after pressing ''OK''."
  Loop until income <> "" and actions_taken <> "" and worker_signature <> ""
  If ButtonPressed = -1 then dialog case_note_dialog
  If buttonpressed = yes_case_note_button then
    call navigate_to_screen("case", "note")
    PF9
    EMReadScreen case_note_check, 17, 2, 33
    EMReadScreen mode_check, 1, 20, 09
    If case_note_check <> "Case Notes (NOTE)" or mode_check <> "A" then MsgBox "The script can't open a case note. Are you in inquiry? Check MAXIS and try again."
  End if
Loop until case_note_check = "Case Notes (NOTE)" and mode_check = "A"

'Logic to enter what the "crisis" variable is from the checkboxes indicated
If eviction_check = 1 then crisis = crisis & "eviction, "
If utility_disconnect_check = 1 then crisis = crisis & "utility disconnect, "
If homelessness_check = 1 then crisis = crisis & "homelessness, "
If security_deposit_check = 1 then crisis = crisis & "security deposit, "
If eviction_check = 0 and utility_disconnect_check = 0 and homelessness_check = 0 and security_deposit_check = 0 then
  crisis = "no crisis given."
Else
  crisis = trim(crisis)
  crisis = left(crisis, len(crisis) - 1) & "."
End if

'Writing the case note
EMSendKey "<home>" & "***Emergency application received***" & "<newline>"
If interview_date <> "" then call write_editbox_in_case_note("Interview date", interview_date, 6)
If HH_comp <> "" then call write_editbox_in_case_note("HH comp", HH_comp, 6)
If crisis <> "" then call write_editbox_in_case_note("Crisis", crisis, 6)
If cause_of_crisis <> "" then call write_editbox_in_case_note("Cause of crisis", cause_of_crisis, 6)
If income <> "" then call write_editbox_in_case_note("Income, past 30 days", income, 6)
If income_under_200_FPG <> "" then call write_editbox_in_case_note("Income under 200% FPG", income_under_200_FPG, 6)
If thirty_percent_rule_notes <> "" then call write_editbox_in_case_note("30% rule notes", thirty_percent_rule_notes, 6)
If monthly_expense <> "" then call write_editbox_in_case_note("Monthly expense", monthly_expense, 6)
If assets <> "" then call write_editbox_in_case_note("Assets", assets, 6)
if verifs_needed <> "" then call write_editbox_in_case_note("Verifs needed", verifs_needed, 6)
If crisis_resolvable <> "" then call write_editbox_in_case_note("Crisis resolvable?", crisis_resolvable, 6)
If discussion_of_crisis <> "" then call write_editbox_in_case_note("Discussion of crisis", discussion_of_crisis, 6)
If actions_taken <> "" then call write_editbox_in_case_note("Actions taken", actions_taken, 6)
If referrals <> "" then call write_editbox_in_case_note("Referrals", referrals, 6)
call write_new_line_in_case_note("---")
If FSPR_check = "FSPR" then
  call write_new_line_in_case_note("   " & FSSM_line_01)
  call write_new_line_in_case_note("   " & FSSM_line_02)
  call write_new_line_in_case_note("   " & FSSM_line_03)
  call write_new_line_in_case_note("---")
End if
call write_new_line_in_case_note(worker_signature)

script_end_procedure("")


