'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTE - MA approval"
start_time = timer


'LOADING ROUTINE FUNCTIONS----------------------------------------------------------------------------------------------------
'>>>>NOTE: these were added as a batch process. Check below for any 'StopScript' functions and convert manually to the script_end_procedure("") function
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\MAXIS-BZ-Scripts-County-Beta\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'DATE CALCULATIONS----------------------------------------------------------------------------------------------------

next_month = dateadd("m", + 1, date)

footer_month = datepart("m", next_month) & ""
If len(footer_month) = 1 then footer_month = "0" & footer_month
footer_year = datepart("yyyy", next_month)
footer_year = "" & footer_year - 2000

'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog case_number_dialog, 0, 0, 161, 61, "Case number"
  Text 5, 5, 85, 10, "Enter your case number:"
  EditBox 95, 0, 60, 15, case_number
  Text 15, 25, 50, 10, "Footer month:"
  EditBox 65, 20, 25, 15, footer_month
  Text 95, 25, 20, 10, "Year:"
  EditBox 120, 20, 25, 15, footer_year
  ButtonGroup ButtonPressed
    OkButton 25, 40, 50, 15
    CancelButton 85, 40, 50, 15
EndDialog

BeginDialog BBUD_Dialog, 0, 0, 191, 76, "BBUD"
  Text 5, 10, 180, 10, "This is a method B budget. What would you like to do?"
  ButtonGroup ButtonPressed
    PushButton 20, 25, 70, 15, "Jump to STAT/BILS", BILS_button
    PushButton 100, 25, 70, 15, "Stay in ELIG/HC", ELIG_button
    CancelButton 135, 55, 50, 15
EndDialog

BeginDialog approval_dialog, 0, 0, 376, 147, "Approval dialog"
  DropListBox 45, 5, 30, 15, "EX"+chr(9)+"DX"+chr(9)+"DP", elig_type
  DropListBox 135, 5, 30, 15, "L"+chr(9)+"S"+chr(9)+"B", budget_type
  EditBox 285, 5, 85, 15, recipient_amt
  EditBox 90, 25, 280, 15, income
  EditBox 50, 45, 320, 15, deductions
  CheckBox 5, 65, 70, 10, "Updated RSPD?", updated_RSPD_check
  CheckBox 85, 65, 110, 10, "Approved new MAXIS results?", approved_check
  CheckBox 210, 65, 70, 10, "Sent DHS-3050?", DHS_3050_check
  EditBox 85, 85, 120, 15, designated_provider
  Text 5, 110, 70, 10, "Other (if applicable):"
  EditBox 75, 105, 295, 15, other
  Text 5, 130, 70, 10, "Sign your case note:"
  EditBox 80, 125, 70, 15, worker_sig
  ButtonGroup ButtonPressed
    OkButton 165, 125, 50, 15
    CancelButton 225, 125, 50, 15
    PushButton 215, 85, 35, 10, "ELIG/HC", ELIG_HC_button
    PushButton 265, 85, 25, 10, "BILS", BILS_button
    PushButton 290, 85, 25, 10, "FACI", FACI_button
    PushButton 315, 85, 25, 10, "HCMI", HCMI_button
    PushButton 340, 85, 25, 10, "UNEA", UNEA_button
  Text 5, 10, 35, 10, "Elig type:"
  Text 85, 10, 45, 10, "Budget type:"
  Text 175, 10, 110, 10, "Waiver obilgation/recipient amt:"
  Text 5, 30, 80, 10, "Total countable income:"
  Text 5, 50, 45, 10, "Deductions:"
  Text 5, 90, 75, 10, "Designated provider:"
  GroupBox 260, 75, 110, 25, "STAT based navigation"
EndDialog

'THE SCRIPT----------------------------------------------------------------------------------------------------

'Connects to BlueZone
EMConnect ""

'Finds case number after setting row/col variables
row = 1
col = 1
EMSearch "Case Nbr:", row, col
If row <> 0 then EMReadScreen case_number, 8, row, col + 10

'Shows case number dialog
Dialog case_number_dialog
If ButtonPressed = 0 then stopscript

'Sends transmit to check for MAXIS
transmit
EMReadScreen MAXIS_check, 5, 1, 39
If MAXIS_check <> "MAXIS" then script_end_procedure("You are not in MAXIS or you are locked out of your case.")

'Navigates back to self
back_to_self

'Navigates to HCMI check to see if it's codedin order to autofill the designated provider info
EMWriteScreen "stat", 16, 43
EMWriteScreen "________", 18, 43
EMWriteScreen case_number, 18, 43
EMWriteScreen footer_month, 20, 43
EMWriteScreen footer_year, 20, 46
EMWriteScreen "hcmi", 21, 70
transmit

'Checks to make sure it's in HCMI, due to error prone cases
EMReadScreen HCMI_check, 4, 2, 55 
If HCMI_check <> "HCMI" then transmit

'Checks the spenddown option. If one is indicated it will navigate to FACI and pull the current FACI into the designated provider box. If no FACI is given it will generate a warning message to the worker to check MMIS.
EMReadScreen spenddown_option, 2, 10, 57
If spenddown_option <> "__" then
  call autofill_editbox_from_MAXIS(array("01"), "FACI", FACI)
  If FACI = "" then
    MsgBox "A current facility could not be found. Check MMIS for a designated provider."
  Else
    designated_provider = FACI
  End if
End if

'Gets back to SELF
back_to_self

'Jumps into ELIG/HC for the footer month listed earlier
EMWriteScreen "elig", 16, 43
EMWriteScreen "________", 18, 43
EMWriteScreen case_number, 18, 43
EMWriteScreen footer_month, 20, 43
EMWriteScreen footer_year, 20, 46
EMWriteScreen "hc", 21, 70
transmit

'Checks to see if MEMB 01 has a case. If not it'll try MEMB 02. If that doesn't work the script will error out on its own due to MAXIS intervention.
EMReadScreen person_check, 2, 8, 31
If person_check = "NO" then
  MsgBox "Person 01 does not have HC on this case. The script will attempt to execute this on person 02. Please check this for errors before approving any results."
  EMWriteScreen "x", 9, 26
End if
If person_check <> "NO" then EMWriteScreen "x", 8, 26
transmit

'Searching for the footer month span after defining row/col variables. If a span can't be found the script will shut down.
row = 1
col = 1
EMSearch " " & footer_month & "/" & footer_year & " ", row, col
If row = 0 then script_end_procedure("A " & footer_month & "/" & footer_year & " span could not be found. Try this again. You may need to run the case through background.")

'Grabbing the elig type and budget type
EMReadScreen elig_type, 2, 12, col - 1
EMReadScreen budget_type, 1, 13, col + 3

'Transmitting into the budget breakdown screen
EMWriteScreen "x", 9, col + 3
transmit

'Checking to see if this is an LBUD. If so, it'll grab the info from the appropriate places.
EMReadScreen LBUD_check, 4, 3, 45
If LBUD_check = "LBUD" then
  EMReadScreen recipient_amt, 10, 15, 70
  recipient_amt = "$" & trim(recipient_amt)
  EMReadScreen income, 10, 12, 32
  income = "$" & trim(income)
  EMReadScreen LTC_exclusions, 10, 14, 32
  If LTC_exclusions <> "__________" then deductions = deductions & "LTC exclusions ($" & replace(LTC_exclusions, "_", "") & "). "
  EMReadScreen medicare_premium, 10, 15, 32
  If medicare_premium <> "__________" then deductions = deductions & "Medicare ($" & replace(medicare_premium, "_", "") & "). "
  EMReadScreen pers_cloth_needs, 10, 16, 32
  If pers_cloth_needs <> "__________" then deductions = deductions & "Personal needs ($" & replace(pers_cloth_needs, "_", "") & "). "
  EMReadScreen home_maintenance_allowance, 10, 17, 32
  If home_maintenance_allowance <> "__________" then deductions = deductions & "Home maintenance allowance ($" & replace(home_maintenance_allowance, "_", "") & "). "
  EMReadScreen guard_rep_payee_fee, 10, 18, 32
  If guard_rep_payee_fee <> "__________" then deductions = deductions & "Payee fee ($" & replace(guard_rep_payee_fee, "_", "") & "). "
  EMReadScreen spousal_allocation, 10, 8, 70
  If spousal_allocation <> "          " then deductions = deductions & "Spousal allocation ($" & replace(spousal_allocation, " ", "") & "). "
  EMReadScreen family_allocation, 10, 9, 70
  If family_allocation <> "__________" then deductions = deductions & "Family allocation ($" & replace(family_allocation, "_", "") & "). "
  EMReadScreen health_ins_premium, 10, 10, 70
  If health_ins_premium <> "__________" then deductions = deductions & "Health insurance premium ($" & replace(health_ins_premium, "_", "") & "). "
  EMReadScreen other_med_expense, 10, 11, 70
  If other_med_expense <> "__________" then deductions = deductions & "Other medical expense ($" & replace(other_med_expense, "_", "") & "). "
  EMReadScreen SSI_1611_benefits, 10, 12, 70
  If SSI_1611_benefits <> "__________" then deductions = deductions & "SSI 1611 benefits ($" & replace(SSI_1611_benefits, "_", "") & "). "
  EMReadScreen other_deductions, 10, 13, 70
  If other_deductions <> "__________" then deductions = deductions & "Other deductions ($" & replace(other_deductions, "_", "") & "). "
End if

'Now it checks to see if this is an SBUD. If so, it'll grab the info from the appropriate places.
EMReadScreen SBUD_check, 4, 3, 44
If SBUD_check = "SBUD" then
  EMReadScreen recipient_amt, 10, 16, 71
  recipient_amt = "$" & trim(recipient_amt)
  EMReadScreen income, 10, 13, 32
  income = "$" & trim(income)
  EMReadScreen LTC_exclusions, 10, 15, 32
  If LTC_exclusions <> "__________" then deductions = deductions & "LTC exclusions ($" & replace(LTC_exclusions, "_", "") & "). "
  EMReadScreen medicare_premium, 10, 16, 32
  If medicare_premium <> "__________" then deductions = deductions & "Medicare ($" & replace(medicare_premium, "_", "") & "). "
  EMReadScreen pers_cloth_needs, 10, 17, 32
  If pers_cloth_needs <> "__________" then deductions = deductions & "Maintenance needs allowance ($" & replace(pers_cloth_needs, "_", "") & "). "
  EMReadScreen guard_rep_payee_fee, 10, 18, 32
  If guard_rep_payee_fee <> "__________" then deductions = deductions & "Payee fee ($" & replace(guard_rep_payee_fee, "_", "") & "). "
  EMReadScreen spousal_allocation, 10, 9, 71
  If spousal_allocation <> "          " then deductions = deductions & "Spousal allocation ($" & replace(spousal_allocation, " ", "") & "). "
  EMReadScreen family_allocation, 10, 10, 71
  If family_allocation <> "__________" then deductions = deductions & "Family allocation ($" & replace(family_allocation, "_", "") & "). "
  EMReadScreen health_ins_premium, 10, 11, 71
  If health_ins_premium <> "__________" then deductions = deductions & "Health insurance premium ($" & replace(health_ins_premium, "_", "") & "). "
  EMReadScreen other_med_expense, 10, 12, 71
  If other_med_expense <> "__________" then deductions = deductions & "Other medical expense ($" & replace(other_med_expense, "_", "") & "). "
  EMReadScreen SSI_1611_benefits, 10, 13, 71
  If SSI_1611_benefits <> "__________" then deductions = deductions & "SSI 1611 benefits ($" & replace(SSI_1611_benefits, "_", "") & "). "
  EMReadScreen other_deductions, 10, 14, 71
  If other_deductions <> "__________" then deductions = deductions & "Other deductions ($" & replace(other_deductions, "_", "") & "). "
End if

'Now it checks to see if this is an EBUD. If so, it'll grab the info from the appropriate places.
EMReadScreen EBUD_check, 4, 3, 60
If EBUD_check = "EBUD" then
  EMReadScreen income, 10, 9, 69
  income = "$" & trim(income)
  EMReadScreen MA_EPD_premium, 10, 13, 69
  other = "MA-EPD premium is $" & trim(MA_EPD_premium) & "/mo."
End if

'Now it checks to see if this is a BBUD. If so, it'll read the info, then offer the worker the chance to navigate to BILS
EMReadScreen BBUD_check, 4, 3, 47
If BBUD_check = "BBUD" then
  EMReadScreen income, 10, 12, 32
  income = "$" & trim(income)
  Dialog BBUD_dialog
  If ButtonPressed = 0 then stopscript
  If ButtonPressed = 4 then
    PF3
    EMReadScreen MAXIS_check, 5, 1, 39
    If MAXIS_check <> "MAXIS" then
      Do
        Dialog BBUD_Dialog
        If buttonpressed = 0 then stopscript
      Loop until MAXIS_check = "MAXIS"
    End if
    back_to_SELF
    EMWriteScreen "stat", 16, 43
    EMWriteScreen "bils", 21, 70
    transmit
    EMReadScreen BILS_check, 4, 2, 54
    If BILS_check <> "BILS" then transmit
  End if
End if


'Cleans up the recipient_amt variable in case it's blank
If recipient_amt = "$" then recipient_amt = "$0"

'Shows the MA approval dialog, checks for MAXIS and allows navigation buttons
Do
  Do
    Dialog approval_dialog
    If buttonpressed = 0 then stopscript
    transmit 'checking for password prompt
    EMReadScreen MAXIS_check, 5, 1, 39
    If MAXIS_check <> "MAXIS" then MsgBox "You are not in MAXIS, or are passworded out. Please navigate back to MAXIS production and try again."
  Loop until MAXIS_check = "MAXIS"
  If buttonpressed = ELIG_HC_button then call navigate_to_screen("elig", "hc__")
  If buttonpressed = BILS_button then call navigate_to_screen("stat", "bils")
  If buttonpressed = FACI_button then call navigate_to_screen("stat", "faci")
  If buttonpressed = HCMI_button then call navigate_to_screen("stat", "hcmi")
  If buttonpressed = UNEA_button then call navigate_to_screen("stat", "unea")
Loop until buttonpressed = -1

'Navigates to a blank case note
call navigate_to_screen("case", "note")
PF9

'Enters the case note info
EMSendKey "<home>"
EMSendKey "**Approved " & elig_type & "-" & budget_type & " for " & footer_month & "/" & footer_year
If elig_type <> "DP" then 
  EMSendKey ", " & recipient_amt
  If budget_type = "L" then EMSendKey " LTC sd**"
  If budget_type = "S" then EMSendKey " SISEW waiver obl**"
  If budget_type = "B" then EMSendKey " recip amt**"
Else
  EMSendKey "**"
End if
EMSendKey "<newline>"
call write_editbox_in_case_note ("Income", income, 6)
call write_editbox_in_case_note ("Deductions", deductions, 6)
call write_new_line_in_case_note ("---")
If updated_RSPD_check = 1 then call write_new_line_in_case_note ("* Updated RSPD in MMIS.")
If designated_provider <> "" then call write_editbox_in_case_note ("Designated provider", designated_provider, 6)
If approved_check = 1 then call write_new_line_in_case_note ("* Approved new MAXIS results.")
If DHS_3050_check = 1 then call write_new_line_in_case_note ("* Sent DHS-3050 LTC communication form to facility.")
If other <> "" then call write_editbox_in_case_note ("Other", other, 6)
call write_new_line_in_case_note ("---")
call write_new_line_in_case_note (worker_sig)

script_end_procedure("")






