'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - LTC - RENEWAL.vbs"
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

BeginDialog LTC_recert_dialog, 5, 5, 431, 247, "LTC recert dialog"
  EditBox 75, 45, 40, 15, recert_month
  EditBox 170, 45, 215, 15, US_citizen
  EditBox 65, 65, 50, 15, MA_type
  EditBox 220, 65, 35, 15, MEDI_reimbursement_prog
  EditBox 360, 65, 65, 15, net_income_amt
  EditBox 45, 85, 115, 15, HH_comp
  EditBox 200, 85, 165, 15, AREP
  EditBox 35, 105, 245, 15, FACI
  EditBox 35, 125, 390, 15, income
  EditBox 35, 145, 390, 15, assets
  EditBox 60, 165, 365, 15, recipient_amt
  EditBox 50, 185, 375, 15, deductions
  EditBox 50, 205, 375, 15, other_notes
  DropListBox 60, 225, 75, 15, "complete"+chr(9)+"incomplete", review_status
  EditBox 215, 225, 65, 15, worker_sig
  ButtonGroup ButtonPressed
    OkButton 320, 225, 50, 15
    CancelButton 375, 225, 50, 15
  GroupBox 20, 5, 60, 35, "Income panels"
  ButtonGroup ButtonPressed
    PushButton 25, 15, 25, 10, "BUSI", BUSI_button
    PushButton 50, 15, 25, 10, "JOBS", JOBS_button
    PushButton 25, 25, 25, 10, "RBIC", RBIC_button
    PushButton 50, 25, 25, 10, "UNEA", UNEA_button
  GroupBox 85, 5, 110, 35, "Asset panels"
  ButtonGroup ButtonPressed
    PushButton 90, 15, 25, 10, "ACCT", ACCT_button
    PushButton 115, 15, 25, 10, "CARS", CARS_button
    PushButton 140, 15, 25, 10, "CASH", CASH_button
    PushButton 165, 15, 25, 10, "OTHR", OTHR_button
    PushButton 90, 25, 25, 10, "REST", REST_button
    PushButton 115, 25, 25, 10, "SECU", SECU_button
    PushButton 140, 25, 25, 10, "TRAN", TRAN_button
  GroupBox 200, 5, 100, 35, "Other important panels:"
  ButtonGroup ButtonPressed
    PushButton 205, 15, 25, 10, "HCRE", HCRE_button
    PushButton 230, 15, 25, 10, "REVW", REVW_button
    PushButton 205, 25, 25, 10, "MEMB", MEMB_button
    PushButton 230, 25, 25, 10, "MEMI", MEMI_button
    PushButton 260, 15, 35, 10, "ELIG/HC", ELIG_HC_button
  GroupBox 305, 5, 105, 35, "STAT-based navigation"
  ButtonGroup ButtonPressed
    PushButton 310, 15, 45, 10, "prev. panel", prev_panel_button
    PushButton 310, 25, 45, 10, "next panel", next_panel_button
    PushButton 360, 15, 45, 10, "prev. memb", prev_memb_button
    PushButton 360, 25, 45, 10, "next memb", next_memb_button
    PushButton 170, 90, 25, 10, "AREP:", AREP_button
    PushButton 5, 110, 25, 10, "FACI:", FACI_button
  Text 5, 50, 70, 10, "Recert footer month:"
  Text 125, 50, 40, 10, "US citizen?:"
  Text 5, 70, 60, 10, "MA type (ie EX-S):"
  Text 130, 70, 90, 10, "MEDI reimbursement prog:"
  Text 275, 70, 80, 10, "Total countable income:"
  Text 5, 90, 40, 10, "HH Comp:"
  Text 5, 130, 30, 10, "Income:"
  Text 5, 150, 25, 10, "Assets:"
  Text 5, 170, 50, 10, "Recipient amt:"
  Text 5, 190, 40, 10, "Deductions:"
  Text 5, 210, 40, 10, "Other notes:"
  Text 5, 230, 55, 10, "Review status:"
  Text 145, 230, 65, 10, "Sign the case note:"
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

'VARIABLES WHICH NEED DECLARING----------------------------------------------------------------------------------------------------
HH_memb_row = 5
HC_check = 1 'It does this so that the shared function actually checks the HC income amounts.
Dim row
Dim col
HH_member_array = array("01") 'Because this script will only be used on member 01 on a case (MA-LTC)

'THE SCRIPT----------------------------------------------------------------------------------------------------

'Connecting to BlueZone
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

'Jumping to STAT
call navigate_to_screen("stat", "memb")
EMReadScreen SELF_check, 4, 2, 50
If SELF_check = "SELF" then script_end_procedure("Couldn't get past the SELF menu. Is your case in background?") 'This is error-proofing.

'This grabs the HH comp in greater detail than the shared function as of 07/11/2013. This includes age/gender/marital status of client.
EMReadScreen client_age, 3, 8, 76
client_age = trim(client_age)
EMReadScreen client_gender, 1, 9, 42
If client_gender = "F" then client_gender = "female"
If client_gender = "M" then client_gender = "male"
HH_comp = client_age & " y/o " & client_gender
call navigate_to_screen("stat", "memi") 
EMReadScreen marital_status, 1, 7, 49
If marital_status = "N" then marital_status = "never married"
If marital_status = "M" then marital_status = "married"
If marital_status = "S" then marital_status = "married living apart"
If marital_status = "L" then marital_status = "legally sep"
If marital_status = "D" then marital_status = "divorced"
If marital_status = "W" then marital_status = "widowed"
HH_comp = HH_comp & ", " & marital_status

'Autofill for the rest of the STAT panels
call autofill_editbox_from_MAXIS(HH_member_array, "ACCT", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "AREP", AREP)
call autofill_editbox_from_MAXIS(HH_member_array, "BUSI", income)
call autofill_editbox_from_MAXIS(HH_member_array, "CASH", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "CARS", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "FACI", FACI)
call autofill_editbox_from_MAXIS(HH_member_array, "JOBS", income)
call autofill_editbox_from_MAXIS(HH_member_array, "MEMI", US_citizen)
call autofill_editbox_from_MAXIS(HH_member_array, "OTHR", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "RBIC", income)
call autofill_editbox_from_MAXIS(HH_member_array, "REST", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "SECU", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "UNEA", income)

'Determining the resert month by combining footer month and year, elig/HC searching will need this
recert_month = footer_month & "/" & footer_year

'Jumping to elig/HC
call navigate_to_screen("elig", "hc")

'Checking to see if person 01 has HC. If not it tries person 02
EMReadScreen person_check, 2, 8, 31
If person_check = "NO" then
  MsgBox "Person 01 does not have HC on this case. The script will attempt to execute this on person 02. Please check this for errors before approving any results."
  EMWriteScreen "x", 9, 26
Else
  EMWriteScreen "x", 8, 26
End if


'Scans for possible secondary programs (QMB/SLMB/QI1)
EMReadScreen second_program_elig_result, 4, 9, 41 
If second_program_elig_result = "ELIG" then 
  EMReadScreen QMB_SLMB_check, 4, 9, 28
  If trim(QMB_SLMB_check) = "QMB" or trim(QMB_SLMB_check) = "SLMB" or trim(QMB_SLMB_check) = "QI1" then MEDI_reimbursement_prog = trim(QMB_SLMB_check)
End if

'Jumps into the ELIG/HC screen for MA
transmit

'Searching for the footer month/year
row = 4 'because the first several rows can contain other "MM/YY" data which can interfere with the search
col = 1
EMSearch footer_month & "/" & footer_year & " ", row, col
If row = 0 then script_end_procedure("A " & footer_month & "/" & footer_year & " span could not be found. Try this again. You may need to run the case through background.")

'Reading info from ELIG/HC, going into the budget breakdown to grab even more info
EMReadScreen elig_type, 2, 12, col - 2
EMReadScreen budget_type, 1, 13, col + 2
EMWriteScreen "x", 9, col + 2
transmit

'Combining the elig type and budget type to create the MA_type variable
MA_type = elig_type & "-" & budget_type

'Checks for LBUD, then if it's an LBUD it'll read info about the budget
EMReadScreen LBUD_check, 4, 3, 45
If LBUD_check = "LBUD" then
  EMReadScreen recipient_amt, 10, 15, 70
  recipient_amt = "$" & trim(recipient_amt)
  EMReadScreen net_income_amt, 10, 12, 32
  net_income_amt = "$" & trim(net_income_amt)
  EMReadScreen LTC_exclusions, 10, 14, 32
  If LTC_exclusions <> "__________" then deductions = deductions & "LTC exclusions ($" & replace(LTC_exclusions, "_", "") & "); "
  EMReadScreen medicare_premium, 10, 15, 32
  If medicare_premium <> "__________" then deductions = deductions & "Medicare ($" & replace(medicare_premium, "_", "") & "); "
  EMReadScreen pers_cloth_needs, 10, 16, 32
  If pers_cloth_needs <> "__________" then deductions = deductions & "Personal needs ($" & replace(pers_cloth_needs, "_", "") & "); "
  EMReadScreen home_maintenance_allowance, 10, 17, 32
  If home_maintenance_allowance <> "__________" then deductions = deductions & "Home maintenance allowance ($" & replace(home_maintenance_allowance, "_", "") & "); "
  EMReadScreen guard_rep_payee_fee, 10, 18, 32
  If guard_rep_payee_fee <> "__________" then deductions = deductions & "Payee fee ($" & replace(guard_rep_payee_fee, "_", "") & "); "
  EMReadScreen spousal_allocation, 10, 8, 70
  If spousal_allocation <> "          " then deductions = deductions & "Spousal allocation ($" & replace(spousal_allocation, " ", "") & "); "
  EMReadScreen family_allocation, 10, 9, 70
  If family_allocation <> "__________" then deductions = deductions & "Family allocation ($" & replace(family_allocation, "_", "") & "); "
  EMReadScreen health_ins_premium, 10, 10, 70
  If health_ins_premium <> "__________" then deductions = deductions & "Health insurance premium ($" & replace(health_ins_premium, "_", "") & "); "
  EMReadScreen other_med_expense, 10, 11, 70
  If other_med_expense <> "__________" then deductions = deductions & "Other medical expense ($" & replace(other_med_expense, "_", "") & "); "
  EMReadScreen SSI_1611_benefits, 10, 12, 70
  If SSI_1611_benefits <> "__________" then deductions = deductions & "SSI 1611 benefits ($" & replace(SSI_1611_benefits, "_", "") & "); "
  EMReadScreen other_deductions, 10, 13, 70
  If other_deductions <> "__________" then deductions = deductions & "Other deductions ($" & replace(other_deductions, "_", "") & "); "
End if

'Checks for SBUD, then if it's an SBUD it'll read info about the budget
EMReadScreen SBUD_check, 4, 3, 44
If SBUD_check = "SBUD" then
  EMReadScreen recipient_amt, 10, 16, 71
  recipient_amt = "$" & trim(recipient_amt)
  EMReadScreen net_income_amt, 10, 13, 32
  net_income_amt = "$" & trim(net_income_amt)
  EMReadScreen LTC_exclusions, 10, 15, 32
  If LTC_exclusions <> "__________" then deductions = deductions & "LTC exclusions ($" & replace(LTC_exclusions, "_", "") & "); "
  EMReadScreen medicare_premium, 10, 16, 32
  If medicare_premium <> "__________" then deductions = deductions & "Medicare ($" & replace(medicare_premium, "_", "") & "); "
  EMReadScreen pers_cloth_needs, 10, 17, 32
  If pers_cloth_needs <> "__________" then deductions = deductions & "Maintenance needs allowance ($" & replace(pers_cloth_needs, "_", "") & "); "
  EMReadScreen guard_rep_payee_fee, 10, 18, 32
  If guard_rep_payee_fee <> "__________" then deductions = deductions & "Payee fee ($" & replace(guard_rep_payee_fee, "_", "") & "); "
  EMReadScreen spousal_allocation, 10, 9, 71
  If spousal_allocation <> "          " then deductions = deductions & "Spousal allocation ($" & replace(spousal_allocation, " ", "") & "); "
  EMReadScreen family_allocation, 10, 10, 71
  If family_allocation <> "__________" then deductions = deductions & "Family allocation ($" & replace(family_allocation, "_", "") & "); "
  EMReadScreen health_ins_premium, 10, 11, 71
  If health_ins_premium <> "__________" then deductions = deductions & "Health insurance premium ($" & replace(health_ins_premium, "_", "") & "); "
  EMReadScreen other_med_expense, 10, 12, 71
  If other_med_expense <> "__________" then deductions = deductions & "Other medical expense ($" & replace(other_med_expense, "_", "") & "); "
  EMReadScreen SSI_1611_benefits, 10, 13, 71
  If SSI_1611_benefits <> "__________" then deductions = deductions & "SSI 1611 benefits ($" & replace(SSI_1611_benefits, "_", "") & "); "
  EMReadScreen other_deductions, 10, 14, 71
  If other_deductions <> "__________" then deductions = deductions & "Other deductions ($" & replace(other_deductions, "_", "") & "); "
End if

'Checks for EBUD, then if it's an EBUD it'll read info about the budget
EMReadScreen EBUD_check, 4, 3, 60
If EBUD_check = "EBUD" then
  EMReadScreen net_income_amt, 10, 9, 69
  net_income_amt = "$" & trim(net_income_amt)
  EMReadScreen MA_EPD_premium, 10, 13, 69
  other = "MA-EPD premium is $" & trim(MA_EPD_premium) & "/mo."
End if

'Checks for BBUD, then if it's a BBUD it'll read info about the budget, and offer a chance to auto-navigate to STAT/BILS to manually fill in budget info
EMReadScreen BBUD_check, 4, 3, 47
If BBUD_check = "BBUD" then
  EMReadScreen net_income_amt, 10, 12, 32
  net_income_amt = "$" & trim(net_income_amt)
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
    call navigate_to_screen("stat", "bils")
    EMReadScreen BILS_check, 4, 2, 54
    If BILS_check <> "BILS" then transmit
  End if
End if

'Cleans up the recipient_amt variable and deductions variable
If recipient_amt = "$" then recipient_amt = "$0"
If right(deductions, 2) = "; " then deductions = left(deductions, len(deductions) - 2)

'Shows the dialog
Do
  Do
    Do
      Do
        Do
          Dialog LTC_recert_dialog
          If ButtonPressed = 0 then 
            dialog cancel_dialog
            If ButtonPressed = yes_cancel_button then stopscript
          End if
        Loop until ButtonPressed <> no_cancel_button
        EMReadScreen STAT_check, 4, 20, 21
        If STAT_check = "STAT" then
          If ButtonPressed = prev_panel_button then call prev_panel_navigation
          If ButtonPressed = next_panel_button then call next_panel_navigation
          If ButtonPressed = prev_memb_button then call prev_memb_navigation
          If ButtonPressed = next_memb_button then call next_memb_navigation
        End if
        transmit 'Forces a screen refresh, to keep MAXIS from erroring out in the event of a password prompt.
        EMReadScreen MAXIS_check, 5, 1, 39
        If MAXIS_check <> "MAXIS" and MAXIS_check <> "AXIS " then MsgBox "You do not appear to be in MAXIS. Are you passworded out? Or in MMIS? Check these and try again."
      Loop until MAXIS_check = "MAXIS" or MAXIS_check = "AXIS " 
      If ButtonPressed = AREP_button then call navigate_to_screen("stat", "arep")
      If ButtonPressed = FACI_button then call navigate_to_screen("stat", "FACI")
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
      If ButtonPressed = HCRE_button then call navigate_to_screen("stat", "HCRE")
      If ButtonPressed = REVW_button then call navigate_to_screen("stat", "REVW")
      If ButtonPressed = MEMB_button then call navigate_to_screen("stat", "MEMB")
      If ButtonPressed = MEMI_button then call navigate_to_screen("stat", "MEMI")
      If ButtonPressed = ELIG_HC_button then call navigate_to_screen("elig", "HC__")
    Loop until ButtonPressed = -1
    If worker_sig = "" then MsgBox "You must sign your case note."
  Loop until worker_sig <> ""
  If ButtonPressed = -1 then dialog case_note_dialog
  If buttonpressed = yes_case_note_button then
    call navigate_to_screen("case", "note")
    PF9
    EMReadScreen case_note_check, 17, 2, 33
    EMReadScreen mode_check, 1, 20, 09
    If case_note_check <> "Case Notes (NOTE)" or mode_check <> "A" then MsgBox "The script can't open a case note. Are you in inquiry? Check MAXIS and try again."
  End if
Loop until case_note_check = "Case Notes (NOTE)" and mode_check = "A"

'Logic to fix the naming in the "recipient amt" variable (not everyone likes calling it "recipient amt"
If budget_type = "L" then recipient_amt_name = "spenddown: "
If budget_type = "S" then recipient_amt_name = "waiver obl: "
If budget_type = "B" then recipient_amt_name = "recipient amt: "

'Logic to add a slash to the MEDI reimbursement variable if it isn't blank
If MEDI_reimbursement_prog <> "" then MEDI_reimbursement_prog = "/" & MEDI_reimbursement_prog

'Writing the case note
EMSendKey "<home>" & "***" & recert_month & " ER " & review_status & ": " & MA_type & MEDI_reimbursement_prog & ", " & recipient_amt_name & recipient_amt & "***" & "<newline>"
call write_editbox_in_case_note("HH comp", HH_comp, 6)
call write_editbox_in_case_note("Citizenship", US_citizen, 6)
call write_editbox_in_case_note("AREP", AREP, 6)
call write_editbox_in_case_note("FACI", FACI, 6)
call write_editbox_in_case_note("Income", income, 6)
call write_editbox_in_case_note("Total countable income", net_income_amt, 6)
call write_editbox_in_case_note("Assets", assets, 6)
call write_editbox_in_case_note("Recipient amt", recipient_amt, 6)
call write_editbox_in_case_note("Deducts", deductions, 6)
If other_notes <> "" then call write_editbox_in_case_note("Notes", other_notes, 6)
call write_new_line_in_case_note("---")
call write_new_line_in_case_note(worker_sig)

script_end_procedure("")


