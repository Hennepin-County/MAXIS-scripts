'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - LTC - RENEWAL.vbs"
start_time = timer

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN		'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF default_directory = "C:\DHS-MAXIS-Scripts\Script Files\" OR default_directory = "" THEN			'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
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
next_month = dateadd("m", + 1, date)
MAXIS_footer_month = datepart("m", next_month)
If len(MAXIS_footer_month) = 1 then MAXIS_footer_month = "0" & MAXIS_footer_month
MAXIS_footer_year = datepart("yyyy", next_month)
MAXIS_footer_year = "" & MAXIS_footer_year - 2000

'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog case_number_dialog, 0, 0, 146, 70, "Case number dialog"
  EditBox 80, 5, 60, 15, case_number					
  EditBox 80, 25, 25, 15, MAXIS_footer_month					
  EditBox 115, 25, 25, 15, MAXIS_footer_year
  ButtonGroup ButtonPressed
    OkButton 35, 45, 50, 15
    CancelButton 90, 45, 50, 15
  Text 10, 30, 65, 10, "Footer month/year:"
  Text 10, 10, 45, 10, "Case number: "
EndDialog


BeginDialog BBUD_Dialog, 0, 0, 191, 76, "BBUD"
  Text 5, 10, 180, 10, "This is a method B budget. What would you like to do?"
  ButtonGroup ButtonPressed
    PushButton 20, 25, 70, 15, "Jump to STAT/BILS", BILS_button
    PushButton 100, 25, 70, 15, "Stay in ELIG/HC", ELIG_button
    CancelButton 135, 55, 50, 15
EndDialog

BeginDialog LTC_recert_dialog, 0, 0, 431, 260, "LTC recert dialog"
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
  CheckBox 5, 225, 100, 10, "Sent forms to AREP?", sent_arep_checkbox
  DropListBox 60, 240, 75, 15, "complete"+chr(9)+"incomplete", review_status
  EditBox 215, 240, 65, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 320, 240, 50, 15
    CancelButton 375, 240, 50, 15
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
  Text 5, 245, 55, 10, "Review status:"
  Text 145, 245, 65, 10, "Sign the case note:"
EndDialog


'VARIABLES WHICH NEED DECLARING----------------------------------------------------------------------------------------------------
HH_memb_row = 5
HC_check = 1 'It does this so that the shared function actually checks the HC income amounts.
Dim row
Dim col
HH_member_array = array("01") 'Because this script will only be used on member 01 on a case (MA-LTC)

'THE SCRIPT----------------------------------------------------------------------------------------------------

'Connecting to BlueZone & grabbing case number & footer month
EMConnect ""
call MAXIS_case_number_finder(case_number)
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)


'Showing the case number dialog
Do
  Dialog case_number_dialog
  If ButtonPressed = 0 then stopscript
  If case_number = "" or IsNumeric(case_number) = False or len(case_number) > 8 then MsgBox "You need to type a valid case number."
Loop until case_number <> "" and IsNumeric(case_number) = True and len(case_number) <= 8

'checking for an active MAXIS session
Call check_for_MAXIS (FALSE)

'This grabs the HH comp in greater detail than the shared function as of 07/11/2013. This includes age/gender/marital status of client.
call navigate_to_MAXIS_screen("stat", "memb")
EMReadScreen client_age, 3, 8, 76
client_age = trim(client_age)
EMReadScreen client_gender, 1, 9, 42
If client_gender = "F" then client_gender = "female"
If client_gender = "M" then client_gender = "male"
HH_comp = client_age & " y/o " & client_gender
call navigate_to_MAXIS_screen("stat", "memi") 
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
recert_month = MAXIS_footer_month & "/" & MAXIS_footer_year


'Checking to see if person 01 has HC. If not it tries person 02
call navigate_to_MAXIS_screen("elig", "hc")
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
EMSearch MAXIS_footer_month & "/" & MAXIS_footer_year & " ", row, col
If row = 0 then script_end_procedure("A " & MAXIS_footer_month & "/" & MAXIS_footer_year & " span could not be found. Try this again. You may need to run the case through background.")

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
    EMReadScreen check_for_MAXIS(True), 5, 1, 39
    If check_for_MAXIS(True) <> "MAXIS" then
      Do
        Dialog BBUD_Dialog
        If buttonpressed = 0 then stopscript
      Loop until check_for_MAXIS(True) = "MAXIS"
    End if
    call navigate_to_MAXIS_screen("stat", "bils")
    EMReadScreen BILS_check, 4, 2, 54
    If BILS_check <> "BILS" then transmit
  End if
End if

'Cleans up the recipient_amt variable and deductions variable
If recipient_amt = "$" then recipient_amt = "$0"
If right(deductions, 2) = "; " then deductions = left(deductions, len(deductions) - 2)

'Shows the recert dialog
Do
    Dialog LTC_recert_dialog
    cancel_confirmation
	MAXIS_dialog_navigation
LOOP until ButtonPressed = -1

'Functions to confirm proceeding to case note & for an active MAXIS session
proceed_confirmation(TRUE)   
Call check_for_MAXIS(False)


'Logic to fix the naming in the "recipient amt" variable (not everyone likes calling it "recipient amt"
If budget_type = "L" then recipient_amt_name = "spenddown: "
If budget_type = "S" then recipient_amt_name = "waiver obl: "
If budget_type = "B" then recipient_amt_name = "recipient amt: "

'Logic to add a slash to the MEDI reimbursement variable if it isn't blank
If MEDI_reimbursement_prog <> "" then MEDI_reimbursement_prog = "/" & MEDI_reimbursement_prog


'THE CASE NOTE----------------------------------------------------------------------------------------------------
start_a_blank_CASE_NOTE
'Writing the case note
Call write_variable_in_case_note("***" & recert_month & " ER " & review_status & ": " & MA_type & MEDI_reimbursement_prog & ", " & recipient_amt_name & recipient_amt & "***")
call write_bullet_and_variable_in_case_note("HH comp", HH_comp)
call write_bullet_and_variable_in_case_note("Citizenship", US_citizen)
call write_bullet_and_variable_in_case_note("AREP", AREP)
call write_bullet_and_variable_in_case_note("FACI", FACI)
call write_bullet_and_variable_in_case_note("Income", income)
call write_bullet_and_variable_in_case_note("Total countable income", net_income_amt)
call write_bullet_and_variable_in_case_note("Assets", assets)
call write_bullet_and_variable_in_case_note("Recipient amt", recipient_amt)
call write_bullet_and_variable_in_case_note("Deducts", deductions)
If other_notes <> "" then call write_bullet_and_variable_in_case_note("Notes", other_notes)
IF Sent_arep_checkbox = 1 THEN CALL write_variable_in_case_note("* Sent form(s) to AREP.")
call write_variable_in_CASE_NOTE("---")
call write_variable_in_CASE_NOTE(worker_signature)

script_end_procedure("")