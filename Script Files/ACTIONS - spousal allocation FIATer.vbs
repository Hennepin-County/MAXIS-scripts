'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ACTIONS - spousal allocation FIATer"
start_time = timer

'LOADING ROUTINE FUNCTIONS----------------------------------------------------------------------------------------------------
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\MAXIS-BZ-Scripts-County-Beta\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog spousal_maintenance_dialog, 0, 0, 256, 185, "Spousal Maintenance Dialog"
  EditBox 5, 15, 35, 15, gross_spousal_unearned_income_type_01
  EditBox 60, 15, 50, 15, gross_spousal_unearned_income_01
  EditBox 5, 35, 35, 15, gross_spousal_unearned_income_type_02
  EditBox 60, 35, 50, 15, gross_spousal_unearned_income_02
  EditBox 5, 55, 35, 15, gross_spousal_unearned_income_type_03
  EditBox 60, 55, 50, 15, gross_spousal_unearned_income_03
  EditBox 5, 75, 35, 15, gross_spousal_unearned_income_type_04
  EditBox 60, 75, 50, 15, gross_spousal_unearned_income_04
  EditBox 5, 105, 35, 15, gross_spousal_earned_income_type_01
  EditBox 60, 105, 50, 15, gross_spousal_earned_income_01
  EditBox 5, 125, 35, 15, gross_spousal_earned_income_type_02
  EditBox 60, 125, 50, 15, gross_spousal_earned_income_02
  EditBox 5, 145, 35, 15, gross_spousal_earned_income_type_03
  EditBox 60, 145, 50, 15, gross_spousal_earned_income_03
  EditBox 5, 165, 35, 15, gross_spousal_earned_income_type_04
  EditBox 60, 165, 50, 15, gross_spousal_earned_income_04
  EditBox 205, 85, 35, 15, mort_rent_payment
  EditBox 205, 105, 35, 15, taxes_and_insuance
  EditBox 210, 125, 35, 15, coop_condo_maint_fees
  EditBox 190, 145, 45, 15, utility_allowance
  ButtonGroup ButtonPressed
    OkButton 135, 165, 50, 15
    CancelButton 190, 165, 50, 15
    PushButton 125, 15, 45, 10, "prev. panel", prev_panel_button
    PushButton 125, 25, 45, 10, "next panel", next_panel_button
    PushButton 185, 15, 45, 10, "prev. memb", prev_memb_button
    PushButton 185, 25, 45, 10, "next memb", next_memb_button
    PushButton 125, 60, 25, 10, "HEST", HEST_button
    PushButton 150, 60, 25, 10, "JOBS", JOBS_button
    PushButton 175, 60, 25, 10, "SHEL", SHEL_button
    PushButton 200, 60, 25, 10, "UNEA", UNEA_button
  Text 10, 5, 25, 10, "UI type"
  Text 75, 5, 30, 10, "Amount"
  Text 10, 95, 25, 10, "EI type"
  Text 75, 95, 30, 10, "Amount"
  GroupBox 120, 5, 115, 35, "STAT-based navigation"
  GroupBox 120, 50, 110, 25, "MAXIS panels"
  Text 130, 90, 65, 10, "Mort-rent payment:"
  Text 130, 110, 75, 10, "Taxes and insurance:"
  Text 130, 130, 80, 10, "Coop-condo maint fees:"
  Text 130, 150, 60, 10, "Utility allowance:"
EndDialog

BeginDialog case_number_dialog, 0, 0, 216, 80, "Case number"
  EditBox 120, 0, 60, 15, case_number
  EditBox 100, 20, 25, 15, footer_month
  EditBox 155, 20, 25, 15, footer_year
  EditBox 115, 40, 25, 15, spousal_allocation_footer_month
  EditBox 175, 40, 25, 15, spousal_allocation_footer_year
  ButtonGroup ButtonPressed
    OkButton 55, 60, 50, 15
    CancelButton 115, 60, 50, 15
  Text 30, 5, 85, 10, "Enter your case number:"
  Text 30, 25, 70, 10, "MAXIS footer month:"
  Text 130, 25, 20, 10, "Year:"
  Text 5, 45, 105, 10, "Spousal allocation footer month:"
  Text 150, 45, 20, 10, "Year:"
EndDialog

'THE SCRIPT----------------------------------------------------------------------------------------------------

'Connects to MAXIS
EMConnect ""

'Grabs case number from MAXIS
call find_variable("Case Nbr: ", case_number, 8)
case_number = trim(case_number)
case_number = replace(case_number, "_", "")
If IsNumeric(case_number) = False then case_number = ""

'Grabs footer month/year from MAXIS
call find_variable("Month: ", footer_month, 2)
call find_variable("Month: " & footer_month & " ", footer_year, 2)
spousal_allocation_footer_month = footer_month
spousal_allocation_footer_year = footer_year

'Dupli

'Shows case number dialog
dialog case_number_dialog
if ButtonPressed = 0 then stopscript

'Navigates back to the SELF menu
back_to_self

'Enters into STAT for the client
EMWriteScreen "stat", 16, 43
EMWriteScreen "________", 18, 43
EMWriteScreen case_number, 18, 43
EMWriteScreen "memb", 21, 70
EMWriteScreen footer_month, 20, 43
EMWriteScreen footer_year, 20, 46
transmit

'Checks to see we're past SELF. If not past SELF (due to error) the script will stop
EMReadScreen SELF_check, 4, 2, 50
If SELF_check = "SELF" then script_end_procedure("You don't appear to have gone past SELF. This case might be in background. Wait for it to come out of background and try again.")

'Checks for which HH member is the spouse. The spouse is coded as "02" on STAT/MEMB.
Do
  EMReadScreen spouse_check, 2, 10, 42
  If spouse_check = "02" then EMReadScreen spousal_reference_number, 2, 4, 33
  EMReadScreen current_memb, 1, 2, 73
  EMReadScreen total_membs, 1, 2, 78
  transmit
Loop until cint(current_memb) = cint(total_membs)

'Jumps to STAT/SHEL.
EMWriteScreen "shel", 20, 71
transmit

'Reads the info off of STAT/SHEL into variables for each type of shelter expense. This is used to autofill the allocation dialog.
EMReadScreen rent_verif, 2, 11, 67
If rent_verif <> "__" and rent_verif <> "NO" and rent_verif <> "?_" then EMReadScreen rent, 8, 11, 56
If rent_verif = "__" or rent_verif = "NO" or rent_verif = "?_" then rent = "0"
EMReadScreen lot_rent_verif, 2, 12, 67
If lot_rent_verif <> "__" and lot_rent_verif <> "NO" and lot_rent_verif <> "?_" then EMReadScreen lot_rent, 8, 12, 56
If lot_rent_verif = "__" or lot_rent_verif = "NO" or lot_rent_verif = "?_" then lot_rent = "0"
EMReadScreen mortgage_verif, 2, 13, 67
If mortgage_verif <> "__" and mortgage_verif <> "NO" and mortgage_verif <> "?_" then EMReadScreen mortgage, 8, 13, 56
If mortgage_verif = "__" or mortgage_verif = "NO" or mortgage_verif = "?_" then mortgage = "0"
EMReadScreen insurance_verif, 2, 14, 67
If insurance_verif <> "__" and insurance_verif <> "NO" and insurance_verif <> "?_" then EMReadScreen insurance, 8, 14, 56
If insurance_verif = "__" or insurance_verif = "NO" or insurance_verif = "?_" then insurance = "0"
EMReadScreen taxes_verif, 2, 15, 67
If taxes_verif <> "__" and taxes_verif <> "NO" and taxes_verif <> "?_" then EMReadScreen taxes, 8, 15, 56
If taxes_verif = "__" or taxes_verif = "NO" or taxes_verif = "?_" then taxes = "0"
EMReadScreen room_verif, 2, 16, 67
If room_verif <> "__" and room_verif <> "NO" and room_verif <> "?_" then EMReadScreen room, 8, 16, 56
If room_verif = "__" or room_verif = "NO" or room_verif = "?_" then room = "0"
EMReadScreen garage_verif, 2, 17, 67
If garage_verif <> "__" and garage_verif <> "NO" and garage_verif <> "?_" then EMReadScreen garage, 8, 17, 56
If garage_verif = "__" or garage_verif = "NO" or garage_verif = "?_" then garage = "0"
EMReadScreen subsidy_verif, 2, 18, 67
If subsidy_verif <> "__" and subsidy_verif <> "NO" and subsidy_verif <> "?_" then EMReadScreen subsidy, 8, 18, 56
If subsidy_verif = "__" or subsidy_verif = "NO" or subsidy_verif = "?_" then subsidy = "0"
mort_rent_payment = cint(rent) + cint(mortgage)
mort_rent_payment = "" & mort_rent_payment
If mort_rent_payment = "0" then mort_rent_payment = ""
taxes_and_insuance = cint(taxes) + cint(insurance)
taxes_and_insuance = "" & taxes_and_insuance
If taxes_and_insuance = "0" then taxes_and_insuance = ""
coop_condo_maint_fees = cint(lot_rent)
coop_condo_maint_fees = "" & coop_condo_maint_fees
If coop_condo_maint_fees = "0" then coop_condo_maint_fees = ""

'Jumps to STAT/HEST
EMWriteScreen "hest", 20, 71
transmit

'Reads the info off of STAT/HEST into variables for the utility expenses. This is used to autofill the allocation dialog.
EMReadScreen utility_allowance, 6, 13, 75
If utility_allowance = "      " then EMReadScreen utility_allowance, 6, 14, 75
If utility_allowance = "      " then EMReadScreen utility_allowance, 6, 15, 75
If utility_allowance = "      " then utility_allowance = ""

'Navigates to UNEA and grabs info on the spouse's income (if a spouse is found in MEMB).
If spousal_reference_number <> "" then
  EMWriteScreen "unea", 20, 71
  EMWriteScreen spousal_reference_number, 20, 76
  transmit
  EMReadScreen current_panel_number, 1, 2, 73
  If current_panel_number = "1" then
    EMReadScreen gross_spousal_unearned_income_type_01, 2, 5, 37
    EMReadScreen gross_spousal_unearned_income_01, 8, 18, 68
    transmit
  End if
  EMReadScreen current_panel_number, 1, 2, 73
  If current_panel_number = "2" then
    EMReadScreen gross_spousal_unearned_income_type_02, 2, 5, 37
    EMReadScreen gross_spousal_unearned_income_02, 8, 18, 68
    transmit
  End if
  EMReadScreen current_panel_number, 1, 2, 73
  If current_panel_number = "3" then
    EMReadScreen gross_spousal_unearned_income_type_03, 2, 5, 37
    EMReadScreen gross_spousal_unearned_income_03, 8, 18, 68
    transmit
  End if
  EMReadScreen current_panel_number, 1, 2, 73
  If current_panel_number = "4" then
    EMReadScreen gross_spousal_unearned_income_type_04, 2, 5, 37
    EMReadScreen gross_spousal_unearned_income_04, 8, 18, 68
    transmit
  End if
  EMWriteScreen "jobs", 20, 71
  EMWriteScreen spousal_reference_number, 20, 76
  transmit
  EMReadScreen current_panel_number, 1, 2, 73
  If current_panel_number = "1" then
    earned_income_number = 1
    EMReadScreen gross_spousal_earned_income_type_01, 1, 5, 38
    If gross_spousal_earned_income_type_01 = "W" then gross_spousal_earned_income_type_01 = "02"
    EMReadScreen gross_spousal_earned_income_01, 8, 17, 67
    transmit
  End if
  EMReadScreen current_panel_number, 1, 2, 73
  If current_panel_number = "2" then
    earned_income_number = earned_income_number + 1
    EMReadScreen gross_spousal_earned_income_type_02, 2, 5, 37
    If gross_spousal_earned_income_type_02 = "W" then gross_spousal_earned_income_type_02 = "02"
    EMReadScreen gross_spousal_earned_income_02, 8, 17, 67
    transmit
  End if
  EMReadScreen current_panel_number, 1, 2, 73
  If current_panel_number = "3" then
    earned_income_number = earned_income_number + 1
    EMReadScreen gross_spousal_earned_income_type_03, 2, 5, 37
    If gross_spousal_earned_income_type_03 = "W" then gross_spousal_earned_income_type_03 = "02"
    EMReadScreen gross_spousal_earned_income_03, 8, 17, 67
    transmit
  End if
  EMReadScreen current_panel_number, 1, 2, 73
  If current_panel_number = "4" then
    earned_income_number = earned_income_number + 1
    EMReadScreen gross_spousal_earned_income_type_04, 2, 5, 37
    If gross_spousal_earned_income_type_04 = "W" then gross_spousal_earned_income_type_04 = "02"
    EMReadScreen gross_spousal_earned_income_04, 8, 17, 67
    transmit
  End if
  EMWriteScreen "busi", 20, 71
  EMWriteScreen spousal_reference_number, 20, 76
  transmit
  EMReadScreen current_panel_number, 1, 2, 73
  If current_panel_number <> "0" then 
    MsgBox "There is self-employment for the spouse. Add this manually to the calculation. The dialog will pop-up after it checks for RBIC."
    has_self_employment = True
  End if
  EMWriteScreen "rbic", 20, 71
  EMWriteScreen spousal_reference_number, 20, 76
  transmit
  EMReadScreen current_panel_number, 1, 2, 73
  If current_panel_number <> "0" then MsgBox "There is room/board income for the spouse. Add this manually to the calculation. The dialog will pop-up next."
  If current_panel_number = "0" and has_self_employment = True then
    EMWriteScreen "busi", 20, 71
    EMWriteScreen spousal_reference_number, 20, 76
    transmit
  End if
End if

'It should trim up the spaces from the screens.
gross_spousal_unearned_income_01 = trim(gross_spousal_unearned_income_01)
gross_spousal_unearned_income_02 = trim(gross_spousal_unearned_income_02)
gross_spousal_unearned_income_03 = trim(gross_spousal_unearned_income_03)
gross_spousal_unearned_income_04 = trim(gross_spousal_unearned_income_04)
gross_spousal_earned_income_01 = trim(gross_spousal_earned_income_01)
gross_spousal_earned_income_02 = trim(gross_spousal_earned_income_02)
gross_spousal_earned_income_03 = trim(gross_spousal_earned_income_03)
gross_spousal_earned_income_04 = trim(gross_spousal_earned_income_04)

'Defining the HH_memb_row variable for the navigation buttons
HH_memb_row = 6

'Shows the spousal maintenance dialog
Do
  Do
    dialog spousal_maintenance_dialog
    If ButtonPressed = 0 then stopscript
    EMReadScreen STAT_check, 4, 20, 21
    If STAT_check = "STAT" then call stat_navigation
    transmit 'Forces a screen refresh, to keep MAXIS from erroring out in the event of a password prompt.
    EMReadScreen MAXIS_check, 5, 1, 39
    If MAXIS_check <> "MAXIS" and MAXIS_check <> "AXIS " then MsgBox "You do not appear to be in MAXIS. Are you passworded out? Or in MMIS? Check these and try again."
  Loop until MAXIS_check = "MAXIS" or MAXIS_check = "AXIS " 
  If ButtonPressed <> -1 then call navigation_buttons
Loop until ButtonPressed = -1

'Jumps back to the SELF menu
back_to_self

'Navigates to ELIG/HC.
EMWriteScreen "elig", 16, 43
EMWriteScreen "hc", 21, 70
transmit

'Checks to see if MEMB 01 has HC, and puts an "x" there. If not it'll try MEMB 02. 
EMReadScreen person_check, 1, 8, 26
If person_check <> "_" then
  MsgBox "Person 01 does not have HC on this case. The script will attempt to execute this on person 02. Please check this for errors before approving any results."
  EMWriteScreen "x", 9, 26
End if
If person_check = "_" then EMWriteScreen "x", 8, 26

'Gets into ELIG/HC for that particular member.
transmit

'Turns on FIAT mode, checks to make sure it worked and sets the reason as "06".
PF9
EMReadScreen FIAT_check, 4, 24, 45
If FIAT_check <> "FIAT" then
  EMSendKey "06"
  transmit
End if

'Defining the variables for the following search
ELIG_HC_row = 6
ELIG_HC_col = 1

'Determining the col variable based on the indicated footer month/year
EMSearch spousal_allocation_footer_month & "/" & spousal_allocation_footer_year, ELIG_HC_row, ELIG_HC_col
If ELIG_HC_col = 0 then script_end_procedure("Requested footer month not found. You may have entered ELIG/HC in an invalid footer month, or results haven't been generated for that month. Check these out and try again.")
col = ELIG_HC_col + 1

'setting the variable for the next do...loop
budget_months = 0 

'Fills all budget months with "x's", so that the script will go into each one in succession.
Do
  EMReadScreen budget_check, 1, 12, col
  If budget_check = "/" then 
    budget_months = budget_months + 1
    EMWriteScreen "x", 9, col + 1
  End if
  col = col + 11
Loop until col > 75

'Jumps into the budget screen (LBUD or SBUD)
transmit

For amt_of_months_to_do = 1 to budget_months

  'For an unknown (as of 06/24/2013) reason, some cases seem to stay in the first budget month and not move on. This gathers the current month to see if we've moved past it later.
  EMReadScreen starting_bdgt_month, 5, 6, 14

  'Checks to see if this is an LBUD or SBUD. It'll stop if it's neither.
  EMReadScreen LBUD_check, 4, 3, 45
  If LBUD_check <> "LBUD" then
    EMReadScreen SBUD_check, 4, 3, 44
    If SBUD_check <> "SBUD" then script_end_procedure("This is not a method L or method S case.")
  End if

  'Transmits to the next screen after putting an "x" on the "spousal allocation" line. This is different on LBUD and SBUD, so the script includes logic.
  If LBUD_check = "LBUD" then EMWriteScreen "x", 8, 44
  If SBUD_check = "SBUD" then EMWriteScreen "x", 9, 44
  transmit

  'Puts an "x" on the Gross Unearned Income and transmits
  EMWriteScreen "x", 5, 5
  transmit

  'Blanks out what's already there, then writes the unearned income type, and income amount, and whether or not it's excluded. Repeats for up to four income types.
  EMWriteScreen "__", 8, 8
  EMWriteScreen gross_spousal_unearned_income_type_01, 8, 8
  EMWriteScreen "__________", 8, 43
  EMWriteScreen gross_spousal_unearned_income_01, 8, 43
  EMWriteScreen "_", 8, 58
  If gross_spousal_unearned_income_type_01 <> "" then
    If gross_spousal_unearned_income_excluded_check_01 = 1 then EMWriteScreen "Y", 8, 58
    If gross_spousal_unearned_income_excluded_check_01 <> 1 then EMWriteScreen "N", 8, 58
  End if
  EMWriteScreen "__", 9, 8
  EMWriteScreen gross_spousal_unearned_income_type_02, 9, 8
  EMWriteScreen "__________", 9, 43
  EMWriteScreen gross_spousal_unearned_income_02, 9, 43
  EMWriteScreen "_", 9, 58
  If gross_spousal_unearned_income_type_02 <> "" then
    If gross_spousal_unearned_income_excluded_check_02 = 1 then EMWriteScreen "Y", 9, 58
    If gross_spousal_unearned_income_excluded_check_02 <> 1 then EMWriteScreen "N", 9, 58
  End if
  EMWriteScreen "__", 10, 8
  EMWriteScreen gross_spousal_unearned_income_type_03, 10, 8
  EMWriteScreen "__________", 10, 43
  EMWriteScreen gross_spousal_unearned_income_03, 10, 43
  EMWriteScreen "_", 10, 58
  If gross_spousal_unearned_income_type_03 <> "" then
    If gross_spousal_unearned_income_excluded_check_03 = 1 then EMWriteScreen "Y", 10, 58
    If gross_spousal_unearned_income_excluded_check_03 <> 1 then EMWriteScreen "N", 10, 58
  End if
  EMWriteScreen "__", 11, 8
  EMWriteScreen gross_spousal_unearned_income_type_04, 11, 8
  EMWriteScreen "__________", 11, 43
  EMWriteScreen gross_spousal_unearned_income_04, 11, 43
  EMWriteScreen "_", 11, 58
  If gross_spousal_unearned_income_type_04 <> "" then
    If gross_spousal_unearned_income_excluded_check_04 = 1 then EMWriteScreen "Y", 11, 58
    If gross_spousal_unearned_income_excluded_check_04 <> 1 then EMWriteScreen "N", 11, 58
  End if

  'Gets out of unearned and heads into earned income
  PF3
  EMWriteScreen "x", 6, 5
  transmit

  'Blanks out what's already there, then writes the earned income type, and income amount, and whether or not it's excluded. Repeats for up to four income types.
  EMWriteScreen "__", 8, 8
  EMWriteScreen gross_spousal_earned_income_type_01, 8, 8
  EMWriteScreen "___________", 8, 43
  EMWriteScreen gross_spousal_earned_income_01, 8, 43
  EMWriteScreen "_", 8, 59
  If gross_spousal_earned_income_type_01 <> "" then
    If gross_spousal_earned_income_excluded_check_01 = 1 then EMWriteScreen "Y", 8, 59
    If gross_spousal_earned_income_excluded_check_01 <> 1 then EMWriteScreen "N", 8, 59
  End if
  EMWriteScreen "__", 9, 8
  EMWriteScreen gross_spousal_earned_income_type_02, 9, 8
  EMWriteScreen "___________", 9, 43
  EMWriteScreen gross_spousal_earned_income_02, 9, 43
  EMWriteScreen "_", 9, 59
  If gross_spousal_earned_income_type_02 <> "" then
    If gross_spousal_earned_income_excluded_check_02 = 1 then EMWriteScreen "Y", 9, 59
    If gross_spousal_earned_income_excluded_check_02 <> 1 then EMWriteScreen "N", 9, 59
  End if
  EMWriteScreen "__", 10, 8
  EMWriteScreen gross_spousal_earned_income_type_03, 10, 8
  EMWriteScreen "___________", 10, 43
  EMWriteScreen gross_spousal_earned_income_03, 10, 43
  EMWriteScreen "_", 10, 59
  If gross_spousal_earned_income_type_03 <> "" then
    If gross_spousal_earned_income_excluded_check_03 = 1 then EMWriteScreen "Y", 10, 59
    If gross_spousal_earned_income_excluded_check_03 <> 1 then EMWriteScreen "N", 10, 59
  End if
  EMWriteScreen "__", 11, 8
  EMWriteScreen gross_spousal_earned_income_type_04, 11, 8
  EMWriteScreen "___________", 11, 43
  EMWriteScreen gross_spousal_earned_income_04, 11, 43
  EMWriteScreen "_", 11, 59
  If gross_spousal_earned_income_type_04 <> "" then
    If gross_spousal_earned_income_excluded_check_04 = 1 then EMWriteScreen "Y", 11, 59
    If gross_spousal_earned_income_excluded_check_04 <> 1 then EMWriteScreen "N", 11, 59
  End if

  'Gets out of earned income
  PF3

  'Blanks out Mort/Rent Payment, Taxes & Insurance, Coop/Condo Maint Fees and Utility Allowance, writes in the amounts from the dialog
  EMWriteScreen "__________", 9, 33
  EMWriteScreen mort_rent_payment, 9, 33
  EMWriteScreen "__________", 10, 33
  EMWriteScreen taxes_and_insuance, 10, 33
  EMWriteScreen "__________", 11, 33
  EMWriteScreen coop_condo_maint_fees, 11, 33
  EMWriteScreen "__________", 12, 33
  EMWriteScreen utility_allowance, 12, 33
  
  'Transmits to get past the spousal allocation screen
  transmit

  'Checks to see if the ACT maint needs check box popped up. If it did, it'll transmit to get past that
  EMReadScreen ACT_maint_needs_check, 3, 17, 4
  If ACT_maint_needs_check = "ACT" then transmit

  'Transmits to jump to the next month.
  transmit

  'For an unknown (as of 06/24/2013) reason, some cases seem to stay in the first budget month and not move on. This is a fix for that.
  EMReadScreen ending_bdgt_month, 5, 6, 14
  If starting_bdgt_month = ending_bdgt_month then transmit
 
  'Resets the variables to check on the next month.
  LBUD_check = ""
  SBUD_check = ""
next

script_end_procedure("")






