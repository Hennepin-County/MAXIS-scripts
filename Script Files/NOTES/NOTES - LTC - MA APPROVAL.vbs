'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - LTC - MA APPROVAL.vbs"
start_time = timer

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN		'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF default_directory = "C:\DHS-MAXIS-Scripts\Script Files\" OR default_directory = "" THEN						
			'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
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

'>>>>NOTE: these were added as a batch process. Check below for any 'StopScript' functions and convert manually to the script_end_procedure("") function

'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog case_number_dialog, 0, 0, 161, 61, "Case number"
  Text 5, 5, 85, 10, "Enter your case number:"
  EditBox 95, 0, 60, 15, case_number
  Text 15, 25, 50, 10, "Footer month:"
  EditBox 65, 20, 25, 15, MAXIS_footer_month
  Text 95, 25, 20, 10, "Year:"
  EditBox 120, 20, 25, 15, MAXIS_footer_year
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


BeginDialog approval_dialog, 0, 0, 376, 140, "Approval dialog"
  DropListBox 45, 5, 30, 15, "EX"+chr(9)+"DX"+chr(9)+"DP", elig_type
  DropListBox 135, 5, 30, 15, "L"+chr(9)+"S"+chr(9)+"B", budget_type
  EditBox 285, 5, 85, 15, recipient_amt
  EditBox 90, 25, 280, 15, income
  EditBox 50, 45, 320, 15, deductions
  CheckBox 5, 65, 70, 10, "Updated RSPD?", updated_RSPD_check
  CheckBox 75, 65, 110, 10, "Approved new MAXIS results?", approved_check
  CheckBox 190, 65, 70, 10, "Sent DHS-3050?", DHS_3050_check
  EditBox 75, 80, 140, 15, designated_provider
  EditBox 75, 100, 295, 15, other
  DropListBox 60, 120, 60, 15, "None"+chr(9)+"Paperless IR"+chr(9)+"HRF", special_header_droplist
  EditBox 190, 120, 70, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 265, 120, 50, 15
    CancelButton 320, 120, 50, 15
    PushButton 265, 80, 25, 10, "BILS", BILS_button
    PushButton 290, 80, 25, 10, "FACI", FACI_button
    PushButton 315, 80, 25, 10, "HCMI", HCMI_button
    PushButton 340, 80, 25, 10, "UNEA", UNEA_button
    PushButton 220, 80, 35, 10, "ELIG/HC", ELIG_HC_button
  Text 85, 10, 45, 10, "Budget type:"
  Text 5, 10, 35, 10, "Elig type:"
  Text 175, 10, 110, 10, "Waiver obilgation/recipient amt:"
  Text 5, 30, 80, 10, "Total countable income:"
  Text 5, 50, 45, 10, "Deductions:"
  Text 5, 85, 70, 10, "Designated provider:"
  GroupBox 260, 70, 110, 25, "STAT based navigation"
  Text 5, 105, 65, 10, "Other (if applicable):"
  Text 130, 125, 60, 10, "Worker signature:"
  Text 5, 125, 53, 10, "Special header:"
EndDialog


'THE SCRIPT----------------------------------------------------------------------------------------------------
'Connects to BlueZone
EMConnect ""

'Grabbing case number & footer month/year
Call MAXIS_case_number_finder(case_number)
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)
 
'Shows case number dialog
Dialog case_number_dialog
cancel_confirmation

'Sends transmit to check for MAXIS
Call check_for_MAXIS(FALSE)

'Going to ELIG/HC for the correct footer month
back_to_self
EMWriteScreen "elig", 16, 43
EMWriteScreen "________", 18, 43
EMWriteScreen case_number, 18, 43
EMWriteScreen MAXIS_footer_month, 20, 43
EMWriteScreen MAXIS_footer_year, 20, 46
EMWriteScreen "hc", 21, 70
transmit


'Checks to make sure it's in HCMI, due to error prone cases
call navigate_to_MAXIS_screen("STAT", "HCMI")
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

'Jumps into ELIG/HC for the footer month listed earlier
Call navigate_to_MAXIS_screen("ELIG", "HC__")

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
EMSearch " " & MAXIS_footer_month & "/" & MAXIS_footer_year & " ", row, col
If row = 0 then script_end_procedure("A " & MAXIS_footer_month & "/" & MAXIS_footer_year & " span could not be found. Try this again. You may need to run the case through background.")

'Grabbing the elig type and budget type
EMReadScreen elig_type, 2, 12, col - 1
EMReadScreen budget_type, 1, 13, col + 3
If (budget_type <> "L" AND budget_type <> "S" AND budget_type <> "B") THEN
	script_end_procedure ("This case is not a L, S or B budget case.  Use the ""Approved Programs"" script instead.")
END if

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
  cancel_confirmation
  If ButtonPressed = 4 then
    PF3
    Call check_for_MAXIS(False)
      Do
        Dialog BBUD_Dialog
        cancel_confirmation
      Loop until MAXIS_check = "MAXIS"
    End if
    Call navigate_to_MAXIS_screen("STAT", "BILS")
    EMReadScreen BILS_check, 4, 2, 54
    If BILS_check <> "BILS" then transmit
End if

'auto-fills and cleans up information to entered into the approval_dialog
If recipient_amt = "$" or recipient_amt = "" then recipient_amt = "$0"
If income = "$" or income = "" then income = "$0"
If deductions = "$" or deductions = "" then deductions = "$0"


'Shows the MA approval dialog, checks for MAXIS and allows navigation buttons
Do
	err_msg = ""
	Dialog approval_dialog
	cancel_confirmation
	MAXIS_dialog_navigation	
	If worker_signature = "" then err_msg = err_msg & vbNewLine & "* Sign your case note."
	IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
LOOP until err_msg = ""


'checking for an active MAXIS session
Call check_for_MAXIS (FALSE)

'THE CASE NOTE----------------------------------------------------------------------------------------------------
start_a_blank_CASE_NOTE
'if case is L budget
If (special_header_droplist = "HRF" AND budget_type = "L") then Call write_variable_in_CASE_NOTE("**Approved " & elig_type & "-" & budget_type & " for HRF " & MAXIS_footer_month & "/" & MAXIS_footer_year & ", " & recipient_amt & " LTC SD**")
If (special_header_droplist = "Paperless IR" AND budget_type = "L") then Call write_variable_in_CASE_NOTE("**Approved " & elig_type & "-" & budget_type & " for paperless IR for " & MAXIS_footer_month & "/" & MAXIS_footer_year & ", " & recipient_amt & " LTC SD**")
If (special_header_droplist = "None" AND budget_type = "L") then Call write_variable_in_case_note("**Approved " & elig_type & "-" & budget_type & " for " & MAXIS_footer_month & "/" & MAXIS_footer_year & ", " & recipient_amt & " LTC SD**")
'if case is S budget
If (special_header_droplist = "HRF" AND budget_type = "S") then Call write_variable_in_CASE_NOTE("**Approved HRF " & elig_type & "-" & budget_type & " for HRF " & MAXIS_footer_month & "/" & MAXIS_footer_year & ", " & recipient_amt & " SISEW waiver obl**")
If (special_header_droplist = "Paperless IR" AND budget_type = "S") then Call write_variable_in_CASE_NOTE("**Approved " & elig_type & "-" & budget_type & " for paperless IR for " & MAXIS_footer_month & "/" & MAXIS_footer_year & ", " & recipient_amt & " SISEW waiver obl**")
If (special_header_droplist = "None" AND budget_type = "S") then Call write_variable_in_case_note("**Approved " & elig_type & "-" & budget_type & " for " & MAXIS_footer_month & "/" & MAXIS_footer_year & ", " & recipient_amt & " SISEW waiver obl**")
'if case is B budget
If (special_header_droplist = "HRF" AND budget_type = "B") then Call write_variable_in_CASE_NOTE("**Approved HRF " & elig_type & "-" & budget_type & " for HRF " & MAXIS_footer_month & "/" & MAXIS_footer_year & ", " & recipient_amt & " recip amt**")
If (special_header_droplist = "Paperless IR" AND budget_type = "B") then Call write_variable_in_CASE_NOTE("**Approved " & elig_type & "-" & budget_type & " for paperless IR for " & MAXIS_footer_month & "/" & MAXIS_footer_year & ", " & recipient_amt & " recip amt**")
If (special_header_droplist = "None" AND budget_type = "B") then Call write_variable_in_case_note("**Approved " & elig_type & "-" & budget_type & " for " & MAXIS_footer_month & "/" & MAXIS_footer_year & ", " & recipient_amt & " recip amt**")
call write_bullet_and_variable_in_case_note ("Income", income)
call write_bullet_and_variable_in_case_note ("Deductions", deductions)
call write_variable_in_case_note ("")
If updated_RSPD_check = 1 then call write_variable_in_case_note("* Updated RSPD in MMIS.")
call write_bullet_and_variable_in_case_note ("Designated provider", designated_provider)
If approved_check = 1 then call write_variable_in_case_note ("* Approved new MAXIS results.")
If DHS_3050_check = 1 then call write_variable_in_case_note ("* Sent DHS-3050 LTC communication form to facility.")
call write_bullet_and_variable_in_case_note ("Other", other)
call write_variable_in_case_note ("---")
call write_variable_in_case_note (worker_signature)

script_end_procedure("")
