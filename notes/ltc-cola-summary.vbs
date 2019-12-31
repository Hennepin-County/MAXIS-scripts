'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - LTC - COLA SUMMARY.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 480          'manual run time in seconds
STATS_denomination = "C"        'C is for each case
'END OF stats block=========================================================================================================

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/Hennepin-County/MAXIS-scripts/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else											'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/Hennepin-County/MAXIS-scripts/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		End if
		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
		req.send													'Sends request
		IF req.Status = 200 THEN									'200 means great success
			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
			Execute req.responseText								'Executes the script code
		ELSE														'Error message
			critical_error_msgbox = MsgBox ("Something has gone wrong. The Functions Library code stored on GitHub was not able to be reached." & vbNewLine & vbNewLine &_
                                            "FuncLib URL: " & FuncLib_URL & vbNewLine & vbNewLine &_
                                            "The script has stopped. Please check your Internet connection. Consult a scripts administrator with any questions.", _
                                            vbOKonly + vbCritical, "BlueZone Scripts Critical Error")
            StopScript
		END IF
	ELSE
		FuncLib_URL = "C:\MAXIS-scripts\MASTER FUNCTIONS LIBRARY.vbs"
		Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
		Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
		text_from_the_other_script = fso_command.ReadAll
		fso_command.Close
		Execute text_from_the_other_script
	END IF
END IF
'END FUNCTIONS LIBRARY BLOCK================================================================================================

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("09/25/2017", "Added header to be specific to the MAXIS footer month/year approved.", "Ilse Ferris, Hennepin County")
call changelog_update("11/29/2016", "Added header update for 2017 in case notes, and made this a variable year vs. hard coding this information into the script, and needing yearly updates.", "Ilse Ferris, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 391, 200, "COLA income dialog"
  EditBox 30, 15, 350, 15, unearned_income
  EditBox 30, 35, 350, 15, earned_income
  EditBox 80, 55, 300, 15, medicare_part_B
  EditBox 30, 100, 350, 15, unearned_income_spouse
  EditBox 30, 120, 350, 15, earned_income_spouse
  EditBox 50, 155, 335, 15, other_notes
  EditBox 200, 175, 65, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 275, 175, 50, 15
    CancelButton 330, 175, 50, 15
  GroupBox 5, 5, 380, 75, "Member 01"
  Text 15, 20, 15, 10, "UI:"
  Text 15, 40, 15, 10, "EI:"
  Text 15, 60, 60, 10, "Medicare Part B:"
  GroupBox 5, 90, 380, 55, "Spouse"
  Text 15, 105, 15, 10, "UI:"
  Text 15, 125, 15, 10, "EI:"
  Text 5, 160, 45, 10, "Other notes:"
  Text 135, 175, 60, 15, "Worker signature:"
EndDialog

'VARIABLES WHICH NEED DECLARING------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
HH_memb_row = 5
HC_check = 1
Dim row
Dim col
HC_check = 1 'This is so the functions will work without having to select a program. It uses the same dialogs as the CSR, which can look in multiple places. This is HC only, so it doesn't need those.

'THE SUB routine----------------------------------------------------------------------------------------------------
Function approval_summary
  PF3
  Call check_for_MAXIS(False)

'First it goes to HCMI check to see if it's coded. If so it'll autofill the designated provider info
  back_to_self
  EMWriteScreen "stat", 16, 43
  EMWriteScreen "________", 18, 43
  EMWriteScreen MAXIS_case_number, 18, 43
  EMWriteScreen MAXIS_footer_month, 20, 43
  EMWriteScreen MAXIS_footer_year, 20, 46
  EMWriteScreen "hcmi", 21, 70
  transmit

  EMReadScreen HCMI_check, 4, 2, 55 'Error prone checking
  If HCMI_check <> "HCMI" then transmit

  EMReadScreen spenddown_option, 2, 10, 57
  If spenddown_option <> "__" then
    EMWriteScreen "faci", 20, 71
    transmit

  'Now it checks to see if there are multiple FACI panels. It will automatically go to the most recent one.
    Do
      EMReadScreen in_year_check_01, 4, 14, 53
      EMReadScreen in_year_check_02, 4, 15, 53
      EMReadScreen in_year_check_03, 4, 16, 53
      EMReadScreen in_year_check_04, 4, 17, 53
      EMReadScreen in_year_check_05, 4, 18, 53
      EMReadScreen out_year_check_01, 4, 14, 77
      EMReadScreen out_year_check_02, 4, 15, 77
      EMReadScreen out_year_check_03, 4, 16, 77
      EMReadScreen out_year_check_04, 4, 17, 77
      EMReadScreen out_year_check_05, 4, 18, 77
      If in_year_check_01 <> "____" and out_year_check_01 = "____" then exit do
      If in_year_check_02 <> "____" and out_year_check_02 = "____" then exit do
      If in_year_check_03 <> "____" and out_year_check_03 = "____" then exit do
      If in_year_check_04 <> "____" and out_year_check_04 = "____" then exit do
      If in_year_check_05 <> "____" and out_year_check_05 = "____" then exit do
      EMReadScreen FACI_current_panel, 1, 2, 73
      EMReadScreen FACI_total_check, 1, 2, 78
      transmit
    Loop until FACI_current_panel = FACI_total_check

  'Now it sends the most recent FACI to the designated_provider variable.
    EMReadScreen FACI_name, 30, 6, 43
    FACI_name = replace(FACI_name, "_", "")
    FACI_name = split(FACI_name)
    For each word in FACI_name
      If word <> "" then
        first_letter_of_word = ucase(left(word, 1))
        rest_of_word = LCase(right(word, len(word) -1))
        If len(word) > 3 then
          designated_provider = designated_provider & first_letter_of_word & rest_of_word & " "
        Else
          designated_provider = designated_provider & word & " "
        End if
      End if
    Next
  End if
  back_to_self

'navigates to elig HC for first month of the year.
  EMWriteScreen "elig", 16, 43
  EMWriteScreen "________", 18, 43
  EMWriteScreen MAXIS_case_number, 18, 43
  EMWriteScreen MAXIS_footer_month, 20, 43
  EMWriteScreen MAXIS_footer_year, 20, 46
  EMWriteScreen "hc", 21, 70
  transmit
 'checks if the first person has HC if not it selects person 02.
  EMReadScreen person_check, 2, 8, 31
  If person_check = "NO" then
    MsgBox "Person 01 does not have HC on this case. The script will attempt to execute this on person 02. Please check this for errors before approving any results."
    EMWriteScreen "x", 9, 26
  End if
  If person_check <> "NO" then EMWriteScreen "x", 8, 26
  transmit

  row = 3
  col = 1
  EMSearch MAXIS_footer_month & "/" & MAXIS_footer_year, row, col
  If row = 0 then
    MsgBox "A " & MAXIS_footer_month & "/" & MAXIS_footer_year & " span could not be found. Try this again. You may need to run the case through background."
    stopscript
  End if

  EMReadScreen elig_type, 2, 12, col - 2
  EMReadScreen budget_type, 1, 12, col + 3
  EMWriteScreen "x", 9, col + 2
  transmit
  'Reads which BUD screen it ended up on and then acts accordingly
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

  EMReadScreen BBUD_check, 4, 3, 47
  If BBUD_check = "BBUD" then
    EMReadScreen income, 10, 12, 32
    income = "$" & trim(income)
    Dialog BBUD_dialog
    cancel_confirmation
    If ButtonPressed = BILS_button then
      PF3
      EMReadScreen check_for_MAXIS(True), 5, 1, 39
      If check_for_MAXIS(True) <> "MAXIS" then
	  	'-------------------------------------------------------------------------------------------------DIALOG
		Dialog1 = "" 'Blanking out previous dialog detail
	  	BeginDialog Dialog1, 0, 0, 191, 76, "BBUD"
	      Text 5, 10, 180, 10, "This is a method B budget. What would you like to do?"
	      ButtonGroup ButtonPressed
	      	PushButton 20, 25, 70, 15, "Jump to STAT/BILS", BILS_button
	      	PushButton 100, 25, 70, 15, "Stay in ELIG/HC", ELIG_button
	      	CancelButton 135, 55, 50, 15
	  	EndDialog
        Do
          Dialog Dialog1
          cancel_confirmation
        Loop until check_for_MAXIS(True) = "MAXIS"
      End if
      back_to_SELF
      Call navigate_to_MAXIS_screen("STAT", "BILS")
      transmit
      EMReadScreen BILS_check, 4, 2, 54
      If BILS_check <> "BILS" then transmit  'ERR checking
    End if
  End if

  EMReadScreen EBUD_check, 4, 3, 60
  If EBUD_check = "EBUD" then
      EMReadScreen income, 10, 16, 32
      income = "$" & trim(income)
      EMReadScreen ma_epd_premium_amt, 7, 13, 72
      ma_epd_premium_amt = trim(ma_epd_premium_amt)
  End If

  	'-------------------------------------------------------------------------------------------------DIALOG
	Dialog1 = "" 'Blanking out previous dialog detail
    '---Dialog is dynamically created and needs results from BBUD to be created therefore needs to be seperate from other dialogs.
    BeginDialog Dialog1, 0, 0, 376, 165, "COLA"
      DropListBox 45, 5, 30, 15, "EX"+chr(9)+"DX"+chr(9)+"DP", elig_type
      DropListBox 135, 5, 30, 15, "L"+chr(9)+"S"+chr(9)+"B"+chr(9)+"E", budget_type
      EditBox 285, 5, 85, 15, recipient_amt
      EditBox 285, 25, 85, 15, ma_epd_premium_amt
      EditBox 90, 45, 280, 15, income
      EditBox 50, 65, 320, 15, deductions
      EditBox 85, 85, 185, 15, designated_provider
      If BBUD_check = "BBUD" then
        Text 280, 85, 25, 10, "NAV to:"
        ButtonGroup ButtonPressed
          PushButton 310, 85, 55, 10, "STAT/BILS", BILS_button_COLADLG
          PushButton 310, 95, 55, 10, "ELIG/BBUD", BBUD_button
      End If
      CheckBox 5, 110, 65, 10, "Updated RSPL?", updated_RSPL_check
      CheckBox 85, 110, 110, 10, "Approved new MAXIS results?", approved_check
      CheckBox 210, 110, 70, 10, "Sent DHS-3050?", DHS_3050_check
      CheckBox 295, 110, 70, 10, "MADE Email Sent", made_email_checkbox
      EditBox 75, 125, 295, 15, other
      EditBox 80, 145, 120, 15, worker_signature
      ButtonGroup ButtonPressed
        OkButton 265, 145, 50, 15
        CancelButton 320, 145, 50, 15
      Text 5, 10, 35, 10, "Elig type:"
      Text 85, 10, 45, 10, "Budget type:"
      Text 175, 10, 110, 10, "Waiver obilgation/recipient amt:"
      Text 220, 30, 60, 10, "MA-EPD Premium:"
      Text 5, 50, 80, 10, "Total countable income:"
      Text 5, 70, 45, 10, "Deductions:"
      Text 5, 90, 70, 10, "Designated provider:"
      Text 5, 130, 65, 10, "Other (if applicable):"
      Text 5, 150, 70, 10, "Sign your case note:"
    EndDialog


   Do
	      Dialog Dialog1
	      If buttonpressed = 0 then stopscript
		If buttonpressed = BBUD_button then
			Call check_for_MAXIS(False)
			back_to_self
			Call navigate_to_MAXIS_screen("ELIG", "HC")
 			 EMReadScreen person_check, 2, 8, 31
 			 If person_check = "NO" then
 			   MsgBox "Person 01 does not have HC on this case. The script will attempt to execute this on person 02. Please check this for errors before approving any results."
  			   EMWriteScreen "x", 9, 26
		       End if
			 If person_check <> "NO" then EMWriteScreen "x", 8, 26
  			 transmit
			 row = 3
  			 col = 1
  			 EMSearch MAXIS_footer_month & "/" & MAXIS_footer_year, row, col
 			 If row = 0 then
 			   MsgBox "A " & MAXIS_footer_month & "/" & MAXIS_footer_year & " span could not be found. Try this again. You may need to run the case through background."
 			   stopscript
			  End if
			  EMReadScreen elig_type, 2, 12, col - 2
			  EMReadScreen budget_type, 1, 13, col + 2
			  EMWriteScreen "x", 9, col + 2
			  transmit
		End if
		If buttonpressed = BILS_button_COLADLG then
			call check_for_MAXIS(False)
			back_to_self
			Call navigate_to_MAXIS_screen("STAT", "BILS")
			EMReadScreen BILS_check, 4, 2, 54
   			If BILS_check <> "BILS" then transmit  'ERR checking
		End if
  Loop until buttonpressed = OK
  back_to_self
  EMWriteScreen MAXIS_case_number, 18, 43

  Call start_a_blank_CASE_NOTE

  If ma_epd_premium_amt <> "" Then
    Call write_variable_in_CASE_NOTE("Approved COLA updates " & MAXIS_footer_month & "/" & MAXIS_footer_year & ": " & elig_type & "-" & budget_type & " - EPD Premium $" & ma_epd_premium_amt)
  else
    Call write_variable_in_CASE_NOTE("Approved COLA updates " & MAXIS_footer_month & "/" & MAXIS_footer_year & ": " & elig_type & "-" & budget_type & " " & recipient_amt)
  end if
  If budget_type = "L" then EMSendKey " LTC SD**"
  If budget_type = "S" then EMSendKey " SISEW waiver obl**"
  If budget_type = "B" then EMSendKey " Recip amt.**"
  Call write_bullet_and_variable_in_CASE_NOTE("Income", income)
  Call write_bullet_and_variable_in_CASE_NOTE("Deductions", deductions)
  call write_bullet_and_variable_in_CASE_NOTE("Designated Provider", designated_provider)
  Call write_variable_in_CASE_NOTE("---")
  If updated_RSPL_check = 1 then call write_variable_in_CASE_NOTE ("* Updated RSPL in MMIS.")
  If designated_provider_check = 1 then write_variable_in_CASE_NOTE("* Client has designated provider.")
  If made_email_checkbox = 1 then write_variable_in_CASE_NOTE("* MADE emailed")
  If approved_check = 1 then call write_variable_in_CASE_NOTE("* Approved new MAXIS results.")
  If DHS_3050_check = 1 then call write_variable_in_CASE_NOTE ("* Sent DHS-3050 LTC communication form to facility.")
  call write_bullet_and_variable_in_CASE_NOTE("Other", other)
  Call write_variable_in_CASE_NOTE("---")
  Call write_variable_in_CASE_NOTE(worker_signature)
end function

' Function for script--------------------------------------------------------------------------------------------------------
function income_summary
	EMConnect ""
	back_to_self
	EMWriteScreen MAXIS_footer_month, 20, 43
	EMWriteScreen MAXIS_footer_year, 20, 46

	'GRABBING THE HH MEMBERS---------------------------------------------------------------------------------------------------------------------------------------------------------------------
	call navigate_to_MAXIS_screen("stat", "unea")
	'checking for an active MAXIS session
	Call check_for_MAXIS(False)

	'THE FOLLOWING DIALOG WILL DYNAMICALLY CHANGE DEPENDING ON THE HH COMP. IT WILL ALLOW A WORKER TO SELECT WHICH HH MEMBERS NEED TO BE INCLUDED IN THE SCRIPT.
	EMReadScreen HH_member_01, 18, 5, 3                                       'THIS GATHERS THE HH MEMBERS DIRECTLY FROM A MAXIS SCREEN.
	EMReadScreen HH_member_02, 18, 6, 3
	EMReadScreen HH_member_03, 18, 7, 3
	EMReadScreen HH_member_04, 18, 8, 3
	EMReadScreen HH_member_05, 18, 9, 3
	EMReadScreen HH_member_06, 18, 10, 3
	EMReadScreen HH_member_07, 18, 11, 3
	EMReadScreen HH_member_08, 18, 12, 3
	EMReadScreen HH_member_09, 18, 13, 3
	EMReadScreen HH_member_10, 18, 14, 3
	EMReadScreen HH_member_11, 18, 15, 3
	EMReadScreen HH_member_12, 18, 16, 3
	EMReadScreen HH_member_13, 18, 17, 3
	EMReadScreen HH_member_14, 18, 18, 3
	EMReadScreen HH_member_15, 18, 19, 3
	dialog_size_variable = 50             'DEFAULT IS 50, BUT IT CHANGES DEPENDING ON THE AMOUNT OF HH MEMBERS.
	If HH_member_03 <> "                  " then dialog_size_variable = 65
	If HH_member_04 <> "                  " then dialog_size_variable = 80
	If HH_member_05 <> "                  " then dialog_size_variable = 95
	If HH_member_06 <> "                  " then dialog_size_variable = 110
	If HH_member_07 <> "                  " then dialog_size_variable = 125
	If HH_member_08 <> "                  " then dialog_size_variable = 140
	If HH_member_09 <> "                  " then dialog_size_variable = 155
	If HH_member_10 <> "                  " then dialog_size_variable = 170
	If HH_member_11 <> "                  " then dialog_size_variable = 185
	If HH_member_12 <> "                  " then dialog_size_variable = 200
	If HH_member_13 <> "                  " then dialog_size_variable = 215
	If HH_member_14 <> "                  " then dialog_size_variable = 230
	If HH_member_15 <> "                  " then dialog_size_variable = 245
	If HH_member_01 <> "                  " then client_01_check = 1          'ALL CHECKBOXES DEFAULT TO CHECKED, AS USUALLY WE NEED ALL HH MEMBER INFO.
	If HH_member_02 <> "                  " then client_02_check = 1
	If HH_member_03 <> "                  " then client_03_check = 1
	If HH_member_04 <> "                  " then client_04_check = 1
	If HH_member_05 <> "                  " then client_05_check = 1
	If HH_member_06 <> "                  " then client_06_check = 1
	If HH_member_07 <> "                  " then client_07_check = 1
	If HH_member_08 <> "                  " then client_08_check = 1
	If HH_member_09 <> "                  " then client_09_check = 1
	If HH_member_10 <> "                  " then client_10_check = 1
	If HH_member_11 <> "                  " then client_11_check = 1
	If HH_member_12 <> "                  " then client_12_check = 1
	If HH_member_13 <> "                  " then client_13_check = 1
	If HH_member_14 <> "                  " then client_14_check = 1
	If HH_member_15 <> "                  " then client_15_check = 1
	'-------------------------------------------------------------------------------------------------DIALOG
	Dialog1 = "" 'Blanking out previous dialog detail
	BeginDialog Dialog1, 0, 0, 191, dialog_size_variable, "HH member dialog"
	ButtonGroup ButtonPressed
		OkButton 135, 10, 50, 15
		CancelButton 135, 30, 50, 15
	Text 10, 5, 105, 10, "Household members to look at:"
	If HH_member_01 <> "                  " then CheckBox 10, 20, 120, 10, HH_member_01, client_01_check
	If HH_member_02 <> "                  " then CheckBox 10, 35, 120, 10, HH_member_02, client_02_check
	If HH_member_03 <> "                  " then CheckBox 10, 50, 120, 10, HH_member_03, client_03_check
	If HH_member_04 <> "                  " then CheckBox 10, 65, 120, 10, HH_member_04, client_04_check
	If HH_member_05 <> "                  " then CheckBox 10, 80, 120, 10, HH_member_05, client_05_check
	If HH_member_06 <> "                  " then CheckBox 10, 95, 120, 10, HH_member_06, client_06_check
	If HH_member_07 <> "                  " then CheckBox 10, 110, 120, 10, HH_member_07, client_07_check
	If HH_member_08 <> "                  " then CheckBox 10, 125, 120, 10, HH_member_08, client_08_check
	If HH_member_09 <> "                  " then CheckBox 10, 140, 120, 10, HH_member_09, client_09_check
	If HH_member_10 <> "                  " then CheckBox 10, 155, 120, 10, HH_member_10, client_10_check
	If HH_member_11 <> "                  " then CheckBox 10, 170, 120, 10, HH_member_11, client_11_check
	If HH_member_12 <> "                  " then CheckBox 10, 185, 120, 10, HH_member_12, client_12_check
	If HH_member_13 <> "                  " then CheckBox 10, 200, 120, 10, HH_member_13, client_13_check
	If HH_member_14 <> "                  " then CheckBox 10, 215, 120, 10, HH_member_14, client_14_check
	If HH_member_15 <> "                  " then CheckBox 10, 230, 120, 10, HH_member_15, client_15_check
	EndDialog

	'NOW IT SHOWS THE DIALOG FROM THE LAST SCREEN
	Dialog Dialog1
	cancel_confirmation
	Call check_for_MAXIS(False)

	'DETERMINING WHICH HH MEMBERS TO LOOK AT
	If client_01_check = 1 then HH_member_array = HH_member_array & left(HH_member_01, 2) & " "
	If client_02_check = 1 then HH_member_array = HH_member_array & left(HH_member_02, 2) & " "
	If client_03_check = 1 then HH_member_array = HH_member_array & left(HH_member_03, 2) & " "
	If client_04_check = 1 then HH_member_array = HH_member_array & left(HH_member_04, 2) & " "
	If client_05_check = 1 then HH_member_array = HH_member_array & left(HH_member_05, 2) & " "
	If client_06_check = 1 then HH_member_array = HH_member_array & left(HH_member_06, 2) & " "
	If client_07_check = 1 then HH_member_array = HH_member_array & left(HH_member_07, 2) & " "
	If client_08_check = 1 then HH_member_array = HH_member_array & left(HH_member_08, 2) & " "
	If client_09_check = 1 then HH_member_array = HH_member_array & left(HH_member_09, 2) & " "
	If client_10_check = 1 then HH_member_array = HH_member_array & left(HH_member_10, 2) & " "
	If client_11_check = 1 then HH_member_array = HH_member_array & left(HH_member_11, 2) & " "
	If client_12_check = 1 then HH_member_array = HH_member_array & left(HH_member_12, 2) & " "
	If client_13_check = 1 then HH_member_array = HH_member_array & left(HH_member_13, 2) & " "
	If client_14_check = 1 then HH_member_array = HH_member_array & left(HH_member_14, 2) & " "
	If client_15_check = 1 then HH_member_array = HH_member_array & left(HH_member_15, 2) & " "
	HH_member_array = trim(HH_member_array)
	HH_member_array = split(HH_member_array, " ")

	'GRABBING THE INFO FOR THE CASE NOTE-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	'DETERMINES THE UNEARNED INCOME RECEIVED BY THE CLIENT
	For each HH_member in HH_member_array
	call navigate_to_MAXIS_screen("stat", "unea")
	EMWriteScreen HH_member, 20, 76
	EMWriteScreen "01", 20, 79
	transmit
	EMReadScreen UNEA_total, 1, 2, 78
	If UNEA_total <> 0 then
		Do
		If HH_member = "01" then
			call add_UNEA_to_variable(unearned_income)
		Else
			call add_unea_to_variable(unearned_income_spouse)
		End if
		EMReadScreen UNEA_panel_current, 1, 2, 73
		If cint(UNEA_panel_current) < cint(UNEA_total) then transmit
		Loop until cint(UNEA_panel_current) = cint(UNEA_total)
	End if
	Next

	'DETERMINES THE JOBS INCOME RECEIVED BY THE CLIENT
	For each HH_member in HH_member_array
	call navigate_to_MAXIS_screen("stat", "jobs")
	EMWriteScreen HH_member, 20, 76
	EMWriteScreen "01", 20, 79
	transmit
	EMReadScreen JOBS_total, 1, 2, 78
	If JOBS_total <> 0 then
		Do
		If HH_member = "01" then
			call add_JOBS_to_variable(earned_income)
		Else
			call add_JOBS_to_variable(earned_income_spouse)
		End if
		EMReadScreen JOBS_panel_current, 1, 2, 73
		If cint(JOBS_panel_current) < cint(JOBS_total) then transmit
		Loop until cint(JOBS_panel_current) = cint(JOBS_total)
	End if
	Next

	'DETERMINES THE BUSI INCOME RECEIVED BY THE CLIENT
	For each HH_member in HH_member_array
	call navigate_to_MAXIS_screen("stat", "busi")
	EMWriteScreen HH_member, 20, 76
	EMWriteScreen "01", 20, 79
	transmit
	EMReadScreen BUSI_total, 1, 2, 78
	If BUSI_total <> 0 then
		Do
		If HH_member = "01" then
			call add_BUSI_to_variable(earned_income)
		Else
			call add_BUSI_to_variable(earned_income_spouse)
		End if
		EMReadScreen BUSI_panel_current, 1, 2, 73
		If cint(BUSI_panel_current) < cint(BUSI_total) then transmit
		Loop until cint(BUSI_panel_current) = cint(BUSI_total)
	End if
	Next

	'DETERMINES THE RBIC INCOME RECEIVED BY THE CLIENT
	For each HH_member in HH_member_array
	call navigate_to_MAXIS_screen("stat", "rbic")
	EMWriteScreen HH_member, 20, 76
	EMWriteScreen "01", 20, 79
	transmit
	EMReadScreen RBIC_total, 1, 2, 78
	If RBIC_total <> 0 then
		Do
		If HH_member = "01" then
			call add_RBIC_to_variable(earned_income)
		Else
			call add_RBIC_to_variable(earned_income_spouse)
		End if
		EMReadScreen RBIC_panel_current, 1, 2, 73
		If cint(RBIC_panel_current) < cint(RBIC_total) then transmit
		Loop until cint(RBIC_panel_current) = cint(RBIC_total)
	End if
	Next

	'DETERMINES THE MEDICARE PART B PAID BY THE CLIENT
	call navigate_to_MAXIS_screen("stat", "medi")
	EMWriteScreen "01", 20, 76
	transmit
	EMReadScreen MEDI_total, 1, 2, 78
	If MEDI_total <> 0 then
	EMReadScreen medicare_part_B, 8, 7, 73
	medicare_part_B = "$" & trim(medicare_part_B)
	End if

	'IT HAS TO CLEAN UP EDIT BOXES--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

	'CLEANS UP THE unearned_income EDITBOX
	unearned_income = trim(unearned_income)
	if right(unearned_income, 1) = ";" then unearned_income = left(unearned_income, len(unearned_income) - 1)
	unearned_income = replace(unearned_income, "/)", ")")
	unearned_income = replace(unearned_income, "$________/non-monthly", "amt unknown")
	unearned_income = replace(unearned_income, "$________/monthly", "amt unknown")
	unearned_income = replace(unearned_income, "$________/weekly", "amt unknown")
	unearned_income = replace(unearned_income, "$________/biweekly", "amt unknown")
	unearned_income = replace(unearned_income, "$________/semimonthly", "amt unknown")

	'CLEANS UP THE earned_income EDITBOX
	earned_income = trim(earned_income)
	if right(earned_income, 1) = ";" then earned_income = left(earned_income, len(earned_income) - 1)
	earned_income = replace(earned_income, "/)", ")")
	earned_income = replace(earned_income, "$________/non-monthly", "amt unknown")
	earned_income = replace(earned_income, "$________/monthly", "amt unknown")
	earned_income = replace(earned_income, "$________/weekly", "amt unknown")
	earned_income = replace(earned_income, "$________/biweekly", "amt unknown")
	earned_income = replace(earned_income, "$________/semimonthly", "amt unknown")

	'CLEANS UP THE unearned_income_spouse EDITBOX
	unearned_income_spouse = trim(unearned_income_spouse)
	if right(unearned_income_spouse, 1) = ";" then unearned_income_spouse = left(unearned_income_spouse, len(unearned_income_spouse) - 1)
	unearned_income_spouse = replace(unearned_income_spouse, "/)", ")")
	unearned_income_spouse = replace(unearned_income_spouse, "$________/non-monthly", "amt unknown")
	unearned_income_spouse = replace(unearned_income_spouse, "$________/monthly", "amt unknown")
	unearned_income_spouse = replace(unearned_income_spouse, "$________/weekly", "amt unknown")
	unearned_income_spouse = replace(unearned_income_spouse, "$________/biweekly", "amt unknown")
	unearned_income_spouse = replace(unearned_income_spouse, "$________/semimonthly", "amt unknown")

	'CLEANS UP THE earned_income_spouse EDITBOX
	earned_income_spouse = trim(earned_income_spouse)
	if right(earned_income_spouse, 1) = ";" then earned_income_spouse = left(earned_income_spouse, len(earned_income_spouse) - 1)
	earned_income_spouse = replace(earned_income_spouse, "/)", ")")
	earned_income_spouse = replace(earned_income_spouse, "$________/non-monthly", "amt unknown")
	earned_income_spouse = replace(earned_income_spouse, "$________/monthly", "amt unknown")
	earned_income_spouse = replace(earned_income_spouse, "$________/weekly", "amt unknown")
	earned_income_spouse = replace(earned_income_spouse, "$________/biweekly", "amt unknown")
	earned_income_spouse = replace(earned_income_spouse, "$________/semimonthly", "amt unknown")

	'CASE NOTE DIALOG--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	Dialog COLA_income_dialog
	cancel_confirmation
	Call check_for_MAXIS(False)

	call start_a_blank_CASE_NOTE
	Call write_variable_in_CASE_NOTE ("===COLA INCOME SUMMARY for " & MAXIS_footer_month & "/" & MAXIS_footer_year & "===")
	call write_bullet_and_variable_in_case_note("Unearned income", unearned_income)
	call write_bullet_and_variable_in_case_note("Earned income", earned_income)
	call write_bullet_and_variable_in_case_note("Medicare Part B premium", medicare_part_B)
	call write_bullet_and_variable_in_case_note("Spousal unearned income", unearned_income_spouse)
	call write_bullet_and_variable_in_case_note("Spousal earned income", earned_income_spouse)
	call write_bullet_and_variable_in_case_note("Other notes", other_notes)
	call write_variable_in_CASE_NOTE("---")
	call write_variable_in_CASE_NOTE(worker_signature)
End function

'CONNECTS TO MAXIS--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
EMConnect ""

Call MAXIS_case_number_finder(MAXIS_case_number)
MAXIS_footer_month = CM_plus_1_mo
MAXIS_footer_year = CM_plus_1_yr

'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 171, 125, "COLA case number dialog"
  EditBox 100, 10, 60, 15, MAXIS_case_number
  EditBox 115, 30, 20, 15, MAXIS_footer_month
  EditBox 140, 30, 20, 15, MAXIS_footer_year
  CheckBox 50, 65, 75, 10, "Approval Summary", approval_summary_check
  CheckBox 50, 80, 70, 10, "Income Summary", income_summary_check
  ButtonGroup ButtonPressed
    OkButton 30, 100, 50, 15
    CancelButton 95, 100, 50, 15
  Text 5, 35, 105, 10, "Approval Month/Year (MM/YY):"
  Text 10, 15, 85, 10, "Enter your case number:"
  GroupBox 30, 55, 110, 40, "COLA case note types"
EndDialog

	DO
		err_msg = ""
		Dialog Dialog1
		cancel_confirmation
		If MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & "* You need to type a valid case number." & vBCr
		If len(MAXIS_footer_month) <> 2 THEN err_msg = err_msg & "* Enter the Approval month in MM format (include leading zero if needed)" & vBCr
		If len(MAXIS_footer_year) <> 2 THEN err_msg = err_msg & "* Enter the Approval year in YY format (include leading zero if needed)" & vBCr
		IF err_msg <> "" THEN msgbox err_msg
	LOOP until err_msg = ""

	If approval_summary_check = 1 then call approval_summary
	If income_summary_check = 1 then call income_summary

script_end_procedure("")
